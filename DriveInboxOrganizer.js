/**
 * Drive inbox organizer — match inbox files by MeetingID-date, look up rules by title,
 * rename from templates, copy into folderPath, write artifact URLs.
 *
 * Inbox filename: `{zoomMeetingId}-{MM.DD.YY}.{ext}`
 * Optional chat: `{zoomMeetingId}-{MM.DD.YY}-chat.txt`
 */

var DRIVE_INBOX_ORGANIZER = {
  inboxFolderId: '',
  clientMeetingsRootId: '',
};

function organizeDriveInbox() {
  var result = organizeDriveInbox_();
  if (!result.ok) {
    notifyUser_(result.message, 'Drive Inbox');
    return result;
  }

  notifyUser_(formatDriveInboxSummary_(result), 'Drive Inbox');
  return result;
}

function organizeDriveInbox_() {
  var config = getDriveInboxOrganizerConfig_();
  if (!config.inboxFolderId || !config.clientMeetingsRootId) {
    return {
      ok: false,
      copied: 0,
      skipped: 0,
      deduped: 0,
      items: [],
      message: 'Set DRIVE_INBOX_FOLDER_ID and CLIENT_MEETINGS_ROOT_FOLDER_ID in Script Properties.'
    };
  }

  var timezone = getConfig_().timezone;
  var rulesMap = buildRulesMap_();
  var rowIndex = buildMeetingRowIndex_(timezone);
  var inbox = DriveApp.getFolderById(config.inboxFolderId);
  var files = inbox.getFiles();
  var copied = 0;
  var skipped = 0;
  var deduped = 0;
  var items = [];

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    if (isSegmentPartFile_(fileName)) {
      skipped++;
      continue;
    }

    var parsed = parseInboxMeetingFilename_(fileName);
    if (!parsed) {
      Logger.log('Inbox filename not MeetingID-date: ' + fileName);
      skipped++;
      continue;
    }

    var rowEntry = rowIndex[parsed.meetingId + '|' + parsed.dateStamp];
    if (!rowEntry) {
      Logger.log('No sheet row for ' + parsed.meetingId + ' on ' + parsed.dateStamp + ': ' + fileName);
      skipped++;
      continue;
    }

    var rowData = rowEntry.data;
    var rule = lookupRuleByTitle_(rulesMap, rowData.title);
    if (!rule) {
      Logger.log('No rules row for title: ' + rowData.title);
      skipped++;
      continue;
    }

    var vars = buildRulesReplacementVars_(rowData, rule, rowData.start, timezone);
    if (!vars) {
      Logger.log('Missing meeting start for row: ' + rowData.title);
      skipped++;
      continue;
    }

    var artifact = parsed.artifact;
    var urlColumn = getArtifactUrlColumn_(artifact);
    if (urlColumn && String(rowData[urlColumn] || '').trim()) {
      Logger.log('Deduped ' + fileName + ': sheet URL already set');
      deduped++;
      continue;
    }

    var targetFileName = getRuleArtifactFileName_(rule, artifact, vars);
    if (!targetFileName) {
      Logger.log('Empty rules filename for ' + artifact + ': ' + rowData.title);
      skipped++;
      continue;
    }

    var segments = getRuleFolderPathSegments_(rule, vars);
    if (!segments.length) {
      Logger.log('Empty rules folderPath for title: ' + rowData.title);
      skipped++;
      continue;
    }

    var finalPath = segments.join('/') + '/' + targetFileName;
    var targetFolder = ensureFolderPathFromRoot_(config.clientMeetingsRootId, segments);
    var existingFile = findFileInFolderByName_(targetFolder, targetFileName);
    if (existingFile) {
      Logger.log('Deduped ' + fileName + ': already in destination as ' + targetFileName);
      writeArtifactUrl_(
        rowEntry.sheet,
        rowEntry.headerMap,
        rowEntry.sheetRow,
        artifact,
        existingFile.getUrl()
      );
      rowData[urlColumn] = existingFile.getUrl();
      deduped++;
      items.push(buildOrganizeItem_(rowData, parsed, fileName, finalPath, 'deduped', existingFile.getUrl()));
      continue;
    }

    var copiedFile = file.makeCopy(targetFileName, targetFolder);
    writeArtifactUrl_(
      rowEntry.sheet,
      rowEntry.headerMap,
      rowEntry.sheetRow,
      artifact,
      copiedFile.getUrl()
    );
    rowData[urlColumn] = copiedFile.getUrl();
    copied++;
    items.push(buildOrganizeItem_(rowData, parsed, fileName, finalPath, 'copied', copiedFile.getUrl()));
  }

  return {
    ok: true,
    copied: copied,
    skipped: skipped,
    deduped: deduped,
    items: items,
    message: formatDriveInboxSummary_({ copied: copied, skipped: skipped, deduped: deduped })
  };
}

function buildOrganizeItem_(rowData, parsed, inboxName, finalPath, status, url) {
  return {
    event_id: String(rowData.event_id || ''),
    title: String(rowData.title || ''),
    zoom_meeting_id: parsed.meetingId,
    dateStamp: parsed.dateStamp,
    artifact: parsed.artifact,
    inboxName: inboxName,
    finalPath: finalPath,
    status: status,
    url: url || ''
  };
}

function formatDriveInboxSummary_(result) {
  return 'Drive inbox: copied ' + result.copied + ', skipped ' + result.skipped +
    ', deduped ' + (result.deduped || 0) + '.';
}

function getDriveInboxOrganizerConfig_() {
  var base = typeof getConfig_ === 'function' ? getConfig_() : {};
  return {
    inboxFolderId: base.driveInboxFolderId || DRIVE_INBOX_ORGANIZER.inboxFolderId,
    clientMeetingsRootId: base.clientMeetingsRootFolderId || DRIVE_INBOX_ORGANIZER.clientMeetingsRootId,
  };
}

function getArtifactUrlColumn_(artifact) {
  var columnMap = {
    video: 'video_url',
    audio: 'audio_url',
    transcript: 'transcript_url',
    meeting_summary: 'pdf_url',
    chat: 'chat_url',
  };
  return columnMap[artifact] || '';
}

function findFileInFolderByName_(folder, fileName) {
  var it = folder.getFilesByName(fileName);
  return it.hasNext() ? it.next() : null;
}

function isSegmentPartFile_(fileName) {
  return /_\d+\.[^.]+$/.test(fileName);
}

/**
 * `{meetingId}-{MM.DD.YY}.ext` or `{meetingId}-{MM.DD.YY}-chat.txt`
 */
function parseInboxMeetingFilename_(fileName) {
  var match = String(fileName || '').match(
    /^(\d+)\s*-\s*(\d{2}\.\d{2}\.\d{2})(?:\s*[-_]?\s*(chat|transcript))?\s*(\.[^.]+)$/i
  );
  if (!match) {
    return null;
  }

  var meetingId = match[1];
  var dateStamp = match[2];
  var suffix = String(match[3] || '').toLowerCase();
  var ext = String(match[4] || '').toLowerCase();
  var artifact = classifyArtifactByExtension_(ext, suffix);
  if (!artifact) {
    return null;
  }

  return {
    meetingId: meetingId,
    dateStamp: dateStamp,
    artifact: artifact,
    extension: ext
  };
}

function classifyArtifactByExtension_(ext, suffix) {
  if (ext === '.mp4') return 'video';
  if (ext === '.m4a') return 'audio';
  if (ext === '.pdf') return 'meeting_summary';
  if (ext === '.txt') {
    if (suffix === 'chat') return 'chat';
    return 'transcript';
  }
  return null;
}

function buildMeetingRowIndex_(timezone) {
  var index = {};
  addSheetRowsToMeetingIndex_(index, getEventsSheet_(), null, timezone);
  addSheetRowsToMeetingIndex_(
    index,
    getNonTrainingEventsSheet_(),
    getConfig_().nonTrainingHeaders,
    timezone
  );
  return index;
}

function addSheetRowsToMeetingIndex_(index, sheet, headers, timezone) {
  var sheetData = getSheetDataObjects_(sheet, headers);
  for (var i = 0; i < sheetData.rows.length; i++) {
    var row = sheetData.rows[i];
    var meetingId = String(row.data.zoom_meeting_id || '').replace(/\D/g, '');
    if (!meetingId) continue;
    var meetingDateIso = formatSheetDateOnly_(row.data.start, timezone);
    if (!meetingDateIso) continue;
    var dateStamp = formatMmDdYy_(meetingDateIso);
    var key = meetingId + '|' + dateStamp;
    if (index[key]) continue;
    index[key] = {
      sheet: sheet,
      headerMap: sheetData.headerMap,
      sheetRow: row.sheetRow,
      data: row.data
    };
  }
}

function formatMmDdYy_(meetingDateIso) {
  var parts = String(meetingDateIso).split('-');
  if (parts.length !== 3) return meetingDateIso;
  return parts[1] + '.' + parts[2] + '.' + parts[0].slice(-2);
}

function sanitizeDriveName_(value) {
  return String(value || '')
    .replace(/[<>:"/\\|?*]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\.$/, '');
}

function ensureFolderPathFromRoot_(rootId, segments) {
  var parentId = rootId;
  for (var s = 0; s < segments.length; s++) {
    parentId = getOrCreateChildFolder_(parentId, segments[s]);
  }
  return DriveApp.getFolderById(parentId);
}

function getOrCreateChildFolder_(parentId, name) {
  var parent = DriveApp.getFolderById(parentId);
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) {
    return it.next().getId();
  }
  return parent.createFolder(name).getId();
}

function writeArtifactUrl_(sheet, headerMap, sheetRow, artifact, url) {
  var header = getArtifactUrlColumn_(artifact);
  if (!header || headerMap[header] === undefined) {
    Logger.log('No sheet column for artifact: ' + artifact);
    return;
  }
  sheet.getRange(sheetRow, headerMap[header] + 1).setValue(url);
}

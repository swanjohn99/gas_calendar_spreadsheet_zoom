/**
 * Drive inbox organizer — match inbox files by meeting ID + start datetime,
 * look up rules by title, rename from templates, copy into folderPath,
 * write artifact URLs and zoom_uuid.
 *
 * Inbox filename: `{zoomMeetingId}-{yyyy-MM-dd HH:mm:ss}_{uuid}.{ext}`
 * Optional chat: `{zoomMeetingId}-{yyyy-MM-dd HH:mm:ss}_{uuid}-chat.txt`
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
    return emptyArtifactResult_(false, 'Set DRIVE_INBOX_FOLDER_ID and CLIENT_MEETINGS_ROOT_FOLDER_ID in Script Properties.');
  }

  var timezone = getConfig_().timezone;
  var rulesList = buildRulesList_();
  var rowIndex = buildMeetingRowIndex_(timezone);
  var inbox = DriveApp.getFolderById(config.inboxFolderId);
  var files = inbox.getFiles();
  var result = emptyArtifactResult_(true, '');

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    if (isSegmentPartFile_(fileName)) {
      result.skipped++;
      continue;
    }

    var parsed = parseInboxMeetingFilename_(fileName);
    if (!parsed) {
      Logger.log('Inbox filename not MeetingID-datetime-uuid: ' + fileName);
      result.skipped++;
      continue;
    }

    var rowEntry = rowIndex[meetingRowIndexKey_(parsed.meetingId, parsed.startStamp)];
    if (!rowEntry) {
      Logger.log('No sheet row for ' + parsed.meetingId + ' at ' + parsed.startStamp + ': ' + fileName);
      result.skipped++;
      continue;
    }

    var filed = fileArtifactForMeetingRow_(rowEntry, parsed.artifact, {
      driveFile: file,
      sourceName: fileName,
      source: 'inbox'
    }, {
      clientMeetingsRootId: config.clientMeetingsRootId,
      rulesList: rulesList,
      timezone: timezone
    });
    if ((filed.status === 'copied' || filed.status === 'deduped') && parsed.uuid) {
      writeZoomUuid_(rowEntry.sheet, rowEntry.headerMap, rowEntry.sheetRow, parsed.uuid);
      rowEntry.data.zoom_uuid = parsed.uuid;
    }
    mergeArtifactOutcome_(result, filed);
  }

  result.message = formatDriveInboxSummary_(result);
  return result;
}

/**
 * File one artifact for a meeting row (shared by inbox organizer and Zoom sync).
 */
function fileArtifactForMeetingRow_(rowEntry, artifact, content, context) {
  var rowData = rowEntry.data;
  var match = matchRuleByTitle_(context.rulesList, rowData.title);
  if (!match || !match.rule) {
    Logger.log('No rules row for title: ' + rowData.title);
    return { status: 'skipped', reason: 'no_rule' };
  }
  var rule = match.rule;

  var vars = buildRulesReplacementVars_(rowData, rule, rowData.start, context.timezone, match.clientName);
  if (!vars) {
    Logger.log('Missing meeting start for row: ' + rowData.title);
    return { status: 'skipped', reason: 'missing_start' };
  }

  var urlColumn = getArtifactUrlColumn_(artifact);
  if (urlColumn && String(rowData[urlColumn] || '').trim()) {
    Logger.log('Deduped ' + (content.sourceName || artifact) + ': sheet URL already set');
    return { status: 'deduped', reason: 'sheet_url_set' };
  }

  var targetFileName = getRuleArtifactFileName_(rule, artifact, vars);
  if (!targetFileName) {
    Logger.log('Empty rules filename for ' + artifact + ': ' + rowData.title);
    return { status: 'skipped', reason: 'empty_filename' };
  }

  var segments = getRuleFolderPathSegments_(rule, vars);
  if (!segments.length) {
    Logger.log('Empty rules folderPath for title: ' + rowData.title);
    return { status: 'skipped', reason: 'empty_folder_path' };
  }

  var finalPath = segments.join('/') + '/' + targetFileName;
  var targetFolder = ensureFolderPathFromRoot_(context.clientMeetingsRootId, segments);
  var existingFile = findFileInFolderByName_(targetFolder, targetFileName);
  if (existingFile) {
    Logger.log('Deduped ' + (content.sourceName || artifact) + ': already in destination as ' + targetFileName);
    writeArtifactUrl_(
      rowEntry.sheet,
      rowEntry.headerMap,
      rowEntry.sheetRow,
      artifact,
      existingFile.getUrl()
    );
    rowData[urlColumn] = existingFile.getUrl();
    return {
      status: 'deduped',
      item: buildOrganizeItem_(rowData, artifact, content.sourceName || '', finalPath, 'deduped',
        existingFile.getUrl(), content.source || 'inbox')
    };
  }

  var savedFile;
  if (content.driveFile) {
    savedFile = content.driveFile.makeCopy(targetFileName, targetFolder);
  } else if (content.blob) {
    savedFile = targetFolder.createFile(content.blob).setName(targetFileName);
  } else {
    Logger.log('No file content for artifact ' + artifact);
    return { status: 'skipped', reason: 'missing_content' };
  }

  writeArtifactUrl_(
    rowEntry.sheet,
    rowEntry.headerMap,
    rowEntry.sheetRow,
    artifact,
    savedFile.getUrl()
  );
  rowData[urlColumn] = savedFile.getUrl();
  return {
    status: 'copied',
    item: buildOrganizeItem_(rowData, artifact, content.sourceName || targetFileName, finalPath,
      'copied', savedFile.getUrl(), content.source || 'inbox')
  };
}

function emptyArtifactResult_(ok, message) {
  return {
    ok: ok,
    copied: 0,
    skipped: 0,
    deduped: 0,
    items: [],
    message: message || ''
  };
}

function mergeArtifactOutcome_(result, outcome) {
  if (!outcome) {
    return;
  }
  if (outcome.status === 'copied') {
    result.copied++;
    if (outcome.item) {
      result.items.push(outcome.item);
    }
    return;
  }
  if (outcome.status === 'deduped') {
    result.deduped++;
    if (outcome.item) {
      result.items.push(outcome.item);
    }
    return;
  }
  result.skipped++;
}

function mergeArtifactResults_(left, right) {
  left = left || emptyArtifactResult_(false, '');
  right = right || emptyArtifactResult_(false, '');
  return {
    ok: !!(left.ok || right.ok),
    copied: (left.copied || 0) + (right.copied || 0),
    skipped: (left.skipped || 0) + (right.skipped || 0),
    deduped: (left.deduped || 0) + (right.deduped || 0),
    items: (left.items || []).concat(right.items || []),
    message: [left.message, right.message].filter(function (part) {
      return !!String(part || '').trim();
    }).join(' ')
  };
}

function buildOrganizeItem_(rowData, artifact, sourceName, finalPath, status, url, source) {
  var meetingId = String(rowData.zoom_meeting_id || '').replace(/\D/g, '');
  var timezone = getConfig_().timezone;
  var meetingDateIso = formatSheetDateOnly_(rowData.start, timezone);
  return {
    event_id: String(rowData.event_id || ''),
    title: String(rowData.title || ''),
    zoom_meeting_id: meetingId,
    dateStamp: meetingDateIso ? formatMmDdYy_(meetingDateIso) : '',
    artifact: artifact,
    source: source || 'inbox',
    inboxName: sourceName,
    finalPath: finalPath,
    status: status,
    url: url || ''
  };
}

function formatDriveInboxSummary_(result) {
  return 'Drive inbox: copied ' + result.copied + ', skipped ' + result.skipped +
    ', deduped ' + (result.deduped || 0) + '.';
}

function formatCombinedArtifactSummary_(result) {
  return 'Artifacts: copied ' + result.copied + ', skipped ' + result.skipped +
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

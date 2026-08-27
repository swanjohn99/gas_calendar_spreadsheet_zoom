/**
 * Drive inbox organizer — match inbox files to sheet rows by Zoom meeting ID
 * and start datetime. Look up rules by title, rename from templates, copy into
 * folderPath, write artifact URLs, per-artifact UUIDs, and zoom_uuid.
 *
 * Inbox filename: `{zoomMeetingId}_{yyyy-MM-dd}_{HH-mm-ss}_{uuid}_{filetype}.{ext}`
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
  var meetingRowIndex = buildMeetingRowIndex_();
  var inbox = DriveApp.getFolderById(config.inboxFolderId);
  var inboxFiles = collectInboxFiles_(inbox);
  var result = emptyArtifactResult_(true, '');

  for (var f = 0; f < inboxFiles.length; f++) {
    var file = inboxFiles[f];
    var fileName = file.getName();

    if (isSegmentPartFile_(fileName)) {
      result.skipped++;
      continue;
    }

    var parsed = parseInboxMeetingFilename_(fileName);
    if (!parsed) {
      Logger.log('Inbox filename not MeetingID_date_time_uuid_filetype: ' + fileName);
      result.skipped++;
      continue;
    }

    var rowEntry = meetingRowIndex[meetingRowIndexKey_(parsed.meetingId, parsed.startStamp)];
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
      timezone: timezone,
      meetingStart: formatMeetingStartStamp_(rowEntry.data.start)
    });
    if ((filed.status === 'copied' || filed.status === 'deduped') && parsed.uuid) {
      writeZoomUuid_(rowEntry.sheet, rowEntry.headerMap, rowEntry.sheetRow, parsed.uuid);
      rowEntry.data.zoom_uuid = parsed.uuid;
      writeArtifactUuid_(
        rowEntry.sheet,
        rowEntry.headerMap,
        rowEntry.sheetRow,
        parsed.artifact,
        parsed.uuid
      );
      var uuidColumn = getArtifactUuidColumn_(parsed.artifact);
      if (uuidColumn) {
        rowEntry.data[uuidColumn] = parsed.uuid;
      }
    }
    mergeArtifactOutcome_(result, filed);
  }

  result.message = formatDriveInboxSummary_(result);
  return result;
}

/**
 * File one artifact for a meeting row from a Drive inbox file.
 */
function fileArtifactForMeetingRow_(rowEntry, artifact, content, context) {
  var rowData = rowEntry.data;
  var match = matchRuleByTitle_(context.rulesList, rowData.title);
  if (!match || !match.rule) {
    Logger.log('No rules row for title: ' + rowData.title);
    return { status: 'skipped', reason: 'no_rule' };
  }
  var rule = match.rule;

  var vars = buildRulesReplacementVars_(
    rowData,
    rule,
    context.meetingStart || rowData.start,
    context.timezone,
    match.clientName
  );
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

  if (!content.driveFile) {
    Logger.log('No Drive file for artifact ' + artifact);
    return { status: 'skipped', reason: 'missing_content' };
  }
  var savedFile = content.driveFile.makeCopy(targetFileName, targetFolder);

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

function buildOrganizeItem_(rowData, artifact, sourceName, finalPath, status, url, source) {
  var meetingId = String(rowData.zoom_meeting_id || '').replace(/\D/g, '');
  return {
    event_id: String(rowData.event_id || ''),
    title: String(rowData.title || ''),
    zoom_meeting_id: meetingId,
    dateStamp: formatDateValue_(rowData.start) || '',
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

function getArtifactUuidColumn_(artifact) {
  var columnMap = {
    video: 'video_uuid',
    audio: 'audio_uuid',
    transcript: 'transcript_uuid',
    meeting_summary: 'pdf_uuid',
    chat: 'chat_uuid',
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
 * `{meetingId}_{yyyy-MM-dd}_{HH-mm-ss}_{uuid}_{filetype}.{ext}`
 */
function parseInboxMeetingFilename_(fileName) {
  var match = String(fileName || '').match(
    /^(\d+)_(\d{4}-\d{2}-\d{2})_(\d{2}-\d{2}-\d{2})_(.+)_(video|audio|pdf|summary|chat|transcript|mp4|m4a)(\.[^.]+)$/i
  );
  if (!match) {
    return null;
  }

  var meetingId = match[1];
  var startStamp = match[2] + ' ' + String(match[3] || '').replace(/-/g, ':');
  var uuid = String(match[4] || '').trim();
  var fileType = String(match[5] || '').toLowerCase();
  var ext = String(match[6] || '').toLowerCase();
  if (!uuid) {
    return null;
  }
  if (fileType === 'summary' && ext !== '.pdf') {
    return null;
  }
  var artifact = classifyArtifactByFileType_(fileType);
  if (!artifact) {
    return null;
  }

  return {
    meetingId: meetingId,
    startStamp: startStamp,
    uuid: uuid,
    artifact: artifact,
    extension: ext
  };
}

function classifyArtifactByFileType_(fileType) {
  if (fileType === 'video' || fileType === 'mp4') return 'video';
  if (fileType === 'audio' || fileType === 'm4a') return 'audio';
  if (fileType === 'pdf' || fileType === 'summary') return 'meeting_summary';
  if (fileType === 'transcript') return 'transcript';
  if (fileType === 'chat') return 'chat';
  return null;
}

function collectInboxFiles_(inboxFolder) {
  var files = inboxFolder.getFiles();
  var inboxFiles = [];
  while (files.hasNext()) {
    inboxFiles.push(files.next());
  }
  return inboxFiles;
}

function buildMeetingRowIndex_() {
  var index = {};
  var sheet = getEventsSheet_();
  var sheetData = getSheetDataObjects_(sheet, null);
  for (var i = 0; i < sheetData.rows.length; i++) {
    var row = sheetData.rows[i];
    var meetingId = String(row.data.zoom_meeting_id || '').replace(/\D/g, '');
    if (!meetingId) continue;
    var startStamp = formatMeetingStartStamp_(row.data.start);
    if (!startStamp) continue;
    index[meetingRowIndexKey_(meetingId, startStamp)] = {
      sheet: sheet,
      headerMap: sheetData.headerMap,
      sheetRow: row.sheetRow,
      data: row.data
    };
  }
  return index;
}

function meetingRowIndexKey_(meetingId, startStamp) {
  return String(meetingId || '') + '|' + String(startStamp || '');
}

function formatMeetingStartStamp_(value) {
  return formatDateValue_(value) || '';
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

function writeZoomUuid_(sheet, headerMap, sheetRow, uuid) {
  if (!uuid || headerMap.zoom_uuid === undefined) {
    Logger.log('No sheet column for zoom_uuid');
    return;
  }
  sheet.getRange(sheetRow, headerMap.zoom_uuid + 1).setValue(uuid);
}

function writeArtifactUuid_(sheet, headerMap, sheetRow, artifact, uuid) {
  var header = getArtifactUuidColumn_(artifact);
  if (!uuid || !header || headerMap[header] === undefined) {
    Logger.log('No sheet column for artifact uuid: ' + artifact);
    return;
  }
  sheet.getRange(sheetRow, headerMap[header] + 1).setValue(uuid);
}

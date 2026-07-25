/**
 * Drive inbox organizer — copy synced files into Client Meetings tree and write URLs.
 *
 * Setup:
 * 1. Set DRIVE_INBOX_FOLDER_ID and CLIENT_MEETINGS_ROOT_FOLDER_ID in Script Properties (see Config.js).
 * 2. Add organizeDriveInbox to your onOpen menu.
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
      message: 'Set DRIVE_INBOX_FOLDER_ID and CLIENT_MEETINGS_ROOT_FOLDER_ID in Script Properties.'
    };
  }

  var sheet = getEventsSheet_();
  var sheetData = getSheetDataObjects_(sheet);
  var timezone = getConfig_().timezone;
  var inbox = DriveApp.getFolderById(config.inboxFolderId);
  var files = inbox.getFiles();
  var copied = 0;
  var skipped = 0;
  var deduped = 0;

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    if (isSegmentPartFile_(fileName)) {
      skipped++;
      continue;
    }

    var match = findRowAndArtifactForFile_(sheetData.rows, fileName, timezone);
    if (!match) {
      Logger.log('No row match for inbox file: ' + fileName);
      skipped++;
      continue;
    }

    var row = match.row;
    var artifact = match.artifact;
    var urlColumn = getArtifactUrlColumn_(artifact);
    if (urlColumn && String(row.data[urlColumn] || '').trim()) {
      Logger.log('Deduped ' + fileName + ': sheet URL already set');
      deduped++;
      continue;
    }

    var targetFolder = ensureMeetingFolderPath_(
      config.clientMeetingsRootId,
      row.data,
      match.meetingDateIso
    );
    var existingFile = findFileInFolderByName_(targetFolder, fileName);
    if (existingFile) {
      Logger.log('Deduped ' + fileName + ': already in destination');
      writeArtifactUrl_(sheet, sheetData.headerMap, row.sheetRow, artifact, existingFile.getUrl());
      deduped++;
      continue;
    }

    var copiedFile = file.makeCopy(fileName, targetFolder);
    writeArtifactUrl_(sheet, sheetData.headerMap, row.sheetRow, artifact, copiedFile.getUrl());
    copied++;
  }

  return {
    ok: true,
    copied: copied,
    skipped: skipped,
    deduped: deduped,
    message: formatDriveInboxSummary_({ copied: copied, skipped: skipped, deduped: deduped })
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

function findRowAndArtifactForFile_(rows, fileName, timezone) {
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var meetingDateIso = formatSheetDateOnly_(row.data.start, timezone);
    if (!meetingDateIso) continue;

    var expected = buildExpectedFilenames_(row.data, meetingDateIso);
    var artifact = classifyArtifactByName_(fileName, expected);
    if (artifact) {
      return { row: row, artifact: artifact, meetingDateIso: meetingDateIso };
    }
  }
  return null;
}

function buildExpectedFilenames_(rowData, meetingDateIso) {
  var stamp = formatMmDdYy_(meetingDateIso);
  var program = sanitizeDriveName_(rowData.program);
  var client = sanitizeDriveName_(
    String(rowData.attendee_first_name || '') + ' ' + String(rowData.attendee_last_name || '')
  );
  return {
    video: program + ' - ' + client + ' ' + stamp + '.mp4',
    audio: program + ' Audio - ' + client + ' ' + stamp + '.m4a',
    transcript: program + ' Transcript - ' + client + ' ' + stamp + '.txt',
    chat: program + ' Chat - ' + client + ' ' + stamp + '.txt',
    meeting_summary: 'Meeting Summary - ' + program + ' - ' + client + ' ' + stamp + '.pdf',
  };
}

function classifyArtifactByName_(fileName, expected) {
  if (fileName === expected.video) return 'video';
  if (fileName === expected.audio) return 'audio';
  if (fileName === expected.transcript) return 'transcript';
  if (fileName === expected.chat) return 'chat';
  if (fileName === expected.meeting_summary) return 'meeting_summary';
  return null;
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

function meetingFolderSegment_(meetingDateIso) {
  return 'Coaching Call ' + formatMmDdYy_(meetingDateIso);
}

function ensureMeetingFolderPath_(rootId, rowData, meetingDateIso) {
  var segments = [
    sanitizeDriveName_(rowData.program),
    sanitizeDriveName_(
      String(rowData.attendee_first_name || '') + ' ' + String(rowData.attendee_last_name || '')
    ),
    meetingFolderSegment_(meetingDateIso),
  ];
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

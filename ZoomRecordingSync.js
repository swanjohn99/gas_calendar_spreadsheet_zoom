/**
 * Sync small Zoom cloud recording artifacts (audio, transcript, chat, summary)
 * into client meeting folders via rules templates.
 */

var ZOOM_RECORDING_SYNC = {
  MAX_ROWS_PER_RUN: 20,
  MATCH_TOLERANCE_MS: 2 * 60 * 60 * 1000
};

var ZOOM_SMALL_ARTIFACT_COLUMNS = ['audio_url', 'transcript_url', 'chat_url', 'pdf_url'];

function syncZoomRecordings() {
  var result = syncZoomRecordings_();
  if (!result.ok) {
    notifyUser_(result.message, 'Zoom Recordings');
    return result;
  }
  notifyUser_(result.message, 'Zoom Recordings');
  return result;
}

function syncZoomRecordings_() {
  if (!isZoomConfigured_()) {
    return emptyArtifactResult_(false,
      'Set ZOOM_ACCOUNT_ID, ZOOM_CLIENT_ID, ZOOM_CLIENT_SECRET, and ZOOM_USER_ID in Script Properties.');
  }

  var organizerConfig = getDriveInboxOrganizerConfig_();
  if (!organizerConfig.clientMeetingsRootId) {
    return emptyArtifactResult_(false, 'Set CLIENT_MEETINGS_ROOT_FOLDER_ID in Script Properties.');
  }

  var config = getConfig_();
  var zoom = getZoomConfig_();
  var rulesMap = buildRulesMap_();
  var candidates = getZoomSyncCandidateRows_(config.timezone);
  var result = emptyArtifactResult_(true, '');
  var recordingsCache = {};

  for (var i = 0; i < candidates.length; i++) {
    var rowEntry = candidates[i];
    var rowData = rowEntry.data;
    var meetingDateIso = formatSheetDateOnly_(rowData.start, config.timezone);
    if (!meetingDateIso) {
      result.skipped++;
      continue;
    }

    var recordingMeeting;
    try {
      recordingMeeting = findRecordingMeetingForRow_(rowData, meetingDateIso, zoom, recordingsCache);
    } catch (error) {
      Logger.log('Zoom lookup failed for ' + rowData.title + ': ' + error);
      result.skipped++;
      continue;
    }

    if (!recordingMeeting || !recordingMeeting.recording_files ||
        !recordingMeeting.recording_files.length) {
      Logger.log('No Zoom recording for ' + rowData.zoom_meeting_id + ' on ' + meetingDateIso);
      result.skipped++;
      continue;
    }

    var files = recordingMeeting.recording_files;
    for (var f = 0; f < files.length; f++) {
      var artifact = mapZoomFileTypeToArtifact_(files[f]);
      if (!artifact) {
        continue;
      }

      var urlColumn = getArtifactUrlColumn_(artifact);
      if (urlColumn && String(rowData[urlColumn] || '').trim()) {
        result.deduped++;
        continue;
      }

      try {
        var blob = buildZoomArtifactBlob_(files[f], artifact);
        var sourceName = String(files[f].file_type || artifact) + '-' +
          String(rowData.zoom_meeting_id || '');
        var filed = fileArtifactForMeetingRow_(rowEntry, artifact, {
          blob: blob,
          sourceName: sourceName,
          source: 'zoom'
        }, {
          clientMeetingsRootId: organizerConfig.clientMeetingsRootId,
          rulesMap: rulesMap,
          timezone: config.timezone
        });
        mergeArtifactOutcome_(result, filed);
      } catch (error) {
        Logger.log('Zoom artifact download failed for ' + rowData.title + ' (' + artifact + '): ' + error);
        result.skipped++;
      }
    }
  }

  result.message = 'Zoom recordings: copied ' + result.copied + ', skipped ' + result.skipped +
    ', deduped ' + result.deduped + '.';
  return result;
}

function getZoomSyncCandidateRows_(timezone) {
  var sheet = getEventsSheet_();
  var sheetData = getSheetDataObjects_(sheet);
  var candidates = [];

  for (var i = 0; i < sheetData.rows.length; i++) {
    var row = sheetData.rows[i];
    if (!isMeetingStartOnOrBeforeToday_(row.data.start, timezone)) {
      continue;
    }
    if (!String(row.data.zoom_meeting_id || '').replace(/\D/g, '')) {
      continue;
    }
    if (!rowNeedsSmallArtifacts_(row.data)) {
      continue;
    }
    candidates.push({
      sheet: sheet,
      headerMap: sheetData.headerMap,
      sheetRow: row.sheetRow,
      data: row.data
    });
  }

  return candidates.slice(0, ZOOM_RECORDING_SYNC.MAX_ROWS_PER_RUN);
}

function rowNeedsSmallArtifacts_(rowData) {
  for (var i = 0; i < ZOOM_SMALL_ARTIFACT_COLUMNS.length; i++) {
    if (!String(rowData[ZOOM_SMALL_ARTIFACT_COLUMNS[i]] || '').trim()) {
      return true;
    }
  }
  return false;
}

function findRecordingMeetingForRow_(rowData, meetingDateIso, zoom, cache) {
  var cacheKey = meetingDateIso + '|' + String(rowData.zoom_meeting_id || '').replace(/\D/g, '');
  if (cache[cacheKey]) {
    return cache[cacheKey];
  }

  var meetings = listUserRecordingsForDate_(zoom.userId, meetingDateIso);
  var matched = findMatchingRecordingMeeting_(meetings, rowData, zoom.timezone);
  if (matched) {
    cache[cacheKey] = matched;
    return matched;
  }

  var meetingId = String(rowData.zoom_meeting_id || '').replace(/\D/g, '');
  var instances = listPastMeetingInstances_(meetingId, meetingDateIso, meetingDateIso);
  for (var i = 0; i < instances.length; i++) {
    var instanceUuid = instances[i].uuid;
    if (!instanceUuid) {
      continue;
    }
    var files = getMeetingRecordings_(instanceUuid);
    if (!files.length) {
      continue;
    }
    var pseudoMeeting = {
      meeting_id: meetingId,
      start_time: instances[i].start_time,
      recording_files: files
    };
    if (recordingMatchesRow_(pseudoMeeting, rowData, zoom.timezone)) {
      cache[cacheKey] = pseudoMeeting;
      return pseudoMeeting;
    }
  }

  cache[cacheKey] = null;
  return null;
}

function findMatchingRecordingMeeting_(meetings, rowData, timezone) {
  for (var i = 0; i < meetings.length; i++) {
    if (recordingMatchesRow_(meetings[i], rowData, timezone)) {
      return meetings[i];
    }
  }
  return null;
}

function recordingMatchesRow_(recording, rowData, timezone) {
  var meetingId = String(rowData.zoom_meeting_id || '').replace(/\D/g, '');
  if (String(recording.meeting_id || '').replace(/\D/g, '') !== meetingId) {
    return false;
  }

  var rowStart = parseSheetDate_(rowData.start);
  var recStart = parseSheetDate_(recording.start_time);
  if (!rowStart || !recStart) {
    return true;
  }

  return Math.abs(recStart.getTime() - rowStart.getTime()) <= ZOOM_RECORDING_SYNC.MATCH_TOLERANCE_MS;
}

function mapZoomFileTypeToArtifact_(file) {
  var fileType = String(file.file_type || '').toUpperCase();
  if (fileType === 'MP4') {
    return null;
  }
  if (fileType === 'M4A') {
    return 'audio';
  }
  if (fileType === 'TRANSCRIPT') {
    return 'transcript';
  }
  if (fileType === 'CHAT') {
    return 'chat';
  }
  if (fileType === 'SUMMARY') {
    return 'meeting_summary';
  }

  var extension = String(file.file_extension || '').toLowerCase();
  if (extension === 'pdf') {
    return 'meeting_summary';
  }
  return null;
}

function buildZoomArtifactBlob_(file, artifact) {
  if (!file.download_url) {
    throw new Error('Recording file missing download_url.');
  }

  var blob = downloadZoomFile_(file.download_url);
  if (artifact === 'transcript') {
    blob = Utilities.newBlob(vttToPlainText_(blob.getDataAsString()), 'text/plain', 'transcript.txt');
  } else if (artifact === 'chat') {
    blob = Utilities.newBlob(blob.getDataAsString(), 'text/plain', 'chat.txt');
  } else if (artifact === 'meeting_summary') {
    var contentType = String(blob.getContentType() || '').toLowerCase();
    if (contentType.indexOf('pdf') < 0 && contentType.indexOf('json') >= 0) {
      blob = Utilities.newBlob(blob.getDataAsString(), 'application/pdf', 'summary.pdf');
    }
  }
  return blob;
}

function vttToPlainText_(vttContent) {
  var lines = String(vttContent || '').split(/\r?\n/);
  var out = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line || line === 'WEBVTT') {
      continue;
    }
    if (/^\d+$/.test(line)) {
      continue;
    }
    if (/^\d{2}:\d{2}/.test(line) && line.indexOf('-->') >= 0) {
      continue;
    }
    out.push(line);
  }
  return out.join('\n');
}

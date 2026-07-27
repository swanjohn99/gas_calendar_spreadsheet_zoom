function doGet(e) {
  var config = getConfig_();
  var params = (e && e.parameter) || {};
  var providedKey = params.key || '';

  if (!config.apiKey || providedKey !== config.apiKey) {
    return createJsonResponse_({ ok: false, error: 'Unauthorized', status: 401 });
  }

  var type = String(params.type || 'training').toLowerCase();
  if (type === 'non_training') {
    return createJsonResponse_(getNonTrainingMeetingsPayload_(config, params));
  }

  return createJsonResponse_(getTrainingMeetingsPayload_(config, params));
}

function getTrainingMeetingsPayload_(config, params) {
  var sheet = getEventsSheet_();
  var data = getSheetDataObjects_(sheet, config.headers);
  var rows = data.rows
    .filter(function (row) {
      return isEmailDraftSavedEmpty_(row.data.email_draft_saved) &&
        isMeetingStartOnOrBeforeToday_(row.data.start, config.timezone);
    })
    .map(function (row) {
      return mapPendingMeetingRow_(row.data);
    });

  return buildMeetingsPayload_(config, rows, params);
}

function getNonTrainingMeetingsPayload_(config, params) {
  var sheet = getNonTrainingEventsSheet_();
  var data = getSheetDataObjects_(sheet, config.nonTrainingHeaders);
  var rows = data.rows
    .filter(function (row) {
      return isMeetingStartOnOrBeforeToday_(row.data.start, config.timezone);
    })
    .map(function (row) {
      return mapPendingMeetingRow_(row.data);
    });

  return buildMeetingsPayload_(config, rows, params);
}

function buildMeetingsPayload_(config, rows, params) {
  var limit = parseInt(params.limit, 10);
  rows = rows.filter(function (row) {
    return row.zoom_meeting_id && row.meeting_start_date;
  });

  if (!isNaN(limit) && limit > 0) {
    rows = rows.slice(0, limit);
  }

  return {
    timezone: config.timezone,
    count: rows.length,
    data: rows
  };
}

function isEmailDraftSavedEmpty_(value) {
  return String(value || '').trim() === '';
}

function mapPendingMeetingRow_(rowData) {
  var zoomMeetingId = String(rowData.zoom_meeting_id || extractZoomMeetingId_(rowData.location) || '').trim();
  var startDate = parseSheetDate_(rowData.start);

  return {
    zoom_meeting_id: zoomMeetingId,
    meeting_start_date: startDate ? formatDateValue_(startDate) : String(rowData.start || '').trim(),
    program: String(rowData.program || '').trim(),
    title: String(rowData.title || '').trim()
  };
}

function createJsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

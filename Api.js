function doGet(e) {
  var config = getConfig_();
  var params = (e && e.parameter) || {};
  var providedKey = params.key || '';

  if (!config.apiKey || providedKey !== config.apiKey) {
    return createJsonResponse_({ ok: false, error: 'Unauthorized', status: 401 });
  }

  // type=non_training|all → all due meetings; default → due + empty email_draft_saved
  var type = String(params.type || 'training').toLowerCase();
  var requireEmptyDraft = type !== 'non_training' && type !== 'all';
  return createJsonResponse_(getMeetingsPayload_(config, params, requireEmptyDraft));
}

function getMeetingsPayload_(config, params, requireEmptyDraft) {
  var sheet = getEventsSheet_();
  var data = getSheetDataObjects_(sheet, config.headers);
  var rows = data.rows
    .filter(function (row) {
      if (!isMeetingStartOnOrBeforeToday_(row.data.start, config.timezone)) {
        return false;
      }
      if (requireEmptyDraft && !isEmailDraftSavedEmpty_(row.data.email_draft_saved)) {
        return false;
      }
      return true;
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
    title: String(rowData.title || '').trim()
  };
}

function createJsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

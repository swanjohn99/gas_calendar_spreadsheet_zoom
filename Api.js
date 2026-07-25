function doGet(e) {
  var config = getConfig_();
  var params = (e && e.parameter) || {};
  var providedKey = params.key || '';

  if (!config.apiKey || providedKey !== config.apiKey) {
    return createJsonResponse_({ ok: false, error: 'Unauthorized', status: 401 });
  }

  var sheet = getEventsSheet_();
  var data = getSheetDataObjects_(sheet);
  var limit = parseInt(params.limit, 10);
  var rows = data.rows
    .filter(function (row) {
      return isEmailDraftSavedEmpty_(row.data.email_draft_saved);
    })
    .map(function (row) {
      return mapPendingMeetingRow_(row.data);
    })
    .filter(function (row) {
      return row.zoom_meeting_id && row.meeting_start_date;
    });

  if (!isNaN(limit) && limit > 0) {
    rows = rows.slice(0, limit);
  }

  return createJsonResponse_({
    timezone: config.timezone,
    count: rows.length,
    data: rows
  });
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
    meeting_type: String(rowData.meeting_type || '').trim(),
    attendee_first_name: String(rowData.attendee_first_name || '').trim(),
    attendee_last_name: String(rowData.attendee_last_name || '').trim()
  };
}

function createJsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

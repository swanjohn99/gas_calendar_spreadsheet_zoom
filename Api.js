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
  return {
    zoom_meeting_id: rowData.zoom_meeting_id || extractZoomMeetingId_(rowData.location),
    meeting_start_date: rowData.start || '',
    meeting_type: rowData.meeting_type || '',
    attendee_first_name: rowData.attendee_first_name || '',
    attendee_last_name: rowData.attendee_last_name || ''
  };
}

function createJsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Zoom Server-to-Server OAuth client for recording API calls.
 */

var ZOOM_API = {
  BASE_URL: 'https://api.zoom.us/v2',
  TOKEN_URL: 'https://zoom.us/oauth/token',
  TOKEN_EXPIRY_BUFFER_MS: 60000,
  CACHE_KEYS: {
    ACCESS_TOKEN: 'ZOOM_ACCESS_TOKEN',
    EXPIRES_AT: 'ZOOM_TOKEN_EXPIRES_AT'
  }
};

function getZoomConfig_() {
  var config = getConfig_();
  var keys = CONFIG.SCRIPT_PROPERTY_KEYS;
  var props = PropertiesService.getScriptProperties();
  return {
    accountId: props.getProperty(keys.ZOOM_ACCOUNT_ID) || '',
    clientId: props.getProperty(keys.ZOOM_CLIENT_ID) || '',
    clientSecret: props.getProperty(keys.ZOOM_CLIENT_SECRET) || '',
    userId: props.getProperty(keys.ZOOM_USER_ID) || 'me',
    timezone: config.timezone
  };
}

function isZoomConfigured_() {
  var zoom = getZoomConfig_();
  return !!(zoom.accountId && zoom.clientId && zoom.clientSecret && zoom.userId);
}

function getZoomAccessToken_() {
  var props = PropertiesService.getScriptProperties();
  var cacheKeys = ZOOM_API.CACHE_KEYS;
  var cachedToken = props.getProperty(cacheKeys.ACCESS_TOKEN);
  var expiresAt = parseInt(props.getProperty(cacheKeys.EXPIRES_AT) || '0', 10);
  if (cachedToken && expiresAt > Date.now() + ZOOM_API.TOKEN_EXPIRY_BUFFER_MS) {
    return cachedToken;
  }

  var zoom = getZoomConfig_();
  if (!zoom.accountId || !zoom.clientId || !zoom.clientSecret) {
    throw new Error('Zoom S2S credentials are not configured.');
  }

  var credentials = Utilities.base64Encode(zoom.clientId + ':' + zoom.clientSecret);
  var tokenUrl = ZOOM_API.TOKEN_URL +
    '?grant_type=account_credentials&account_id=' + encodeURIComponent(zoom.accountId);
  var response = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Zoom token request failed: ' + response.getContentText());
  }

  var payload = JSON.parse(response.getContentText());
  if (!payload.access_token) {
    throw new Error('Zoom token response missing access_token.');
  }

  var ttlMs = Math.max(0, (Number(payload.expires_in) || 3600) - 60) * 1000;
  props.setProperty(cacheKeys.ACCESS_TOKEN, payload.access_token);
  props.setProperty(cacheKeys.EXPIRES_AT, String(Date.now() + ttlMs));
  return payload.access_token;
}

function zoomFetch_(path, options) {
  options = options || {};
  var token = getZoomAccessToken_();
  var url = path.indexOf('http') === 0 ? path : ZOOM_API.BASE_URL + path;
  var params = {
    method: options.method || 'get',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  if (options.payload !== undefined) {
    params.payload = JSON.stringify(options.payload);
  }

  var response = UrlFetchApp.fetch(url, params);
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Zoom API ' + path + ' failed (' + code + '): ' + body);
  }
  return body ? JSON.parse(body) : {};
}

function downloadZoomFile_(downloadUrl) {
  var token = getZoomAccessToken_();
  var response = UrlFetchApp.fetch(downloadUrl, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Zoom download failed (' + response.getResponseCode() + ').');
  }
  return response.getBlob();
}

function listUserRecordingsForDate_(userId, dateIso) {
  var encodedUserId = encodeURIComponent(userId);
  var path = '/users/' + encodedUserId + '/recordings?from=' + dateIso + '&to=' + dateIso + '&page_size=100';
  var payload = zoomFetch_(path);
  return payload.meetings || [];
}

function listPastMeetingInstances_(meetingId, fromDate, toDate) {
  var path = '/past_meetings/' + encodeURIComponent(meetingId) + '/instances' +
    '?from=' + fromDate + '&to=' + toDate;
  var payload = zoomFetch_(path);
  return payload.meetings || [];
}

function getMeetingRecordings_(meetingUuid) {
  var encodedUuid = encodeURIComponent(encodeURIComponent(meetingUuid));
  var payload = zoomFetch_('/meetings/' + encodedUuid + '/recordings');
  return payload.recording_files || [];
}

/**
 * One-time / ops helper — set Zoom Script Properties (no secrets in source).
 * Run via Apps Script editor or: clasp run setupZoomScriptProperties --params '[...]'
 */
function setupZoomScriptProperties(accountId, clientId, clientSecret, userId) {
  var keys = CONFIG.SCRIPT_PROPERTY_KEYS;
  PropertiesService.getScriptProperties().setProperties({
    [keys.ZOOM_ACCOUNT_ID]: String(accountId || ''),
    [keys.ZOOM_CLIENT_ID]: String(clientId || ''),
    [keys.ZOOM_CLIENT_SECRET]: String(clientSecret || ''),
    [keys.ZOOM_USER_ID]: String(userId || 'me')
  });
  PropertiesService.getScriptProperties().deleteProperty(ZOOM_API.CACHE_KEYS.ACCESS_TOKEN);
  PropertiesService.getScriptProperties().deleteProperty(ZOOM_API.CACHE_KEYS.EXPIRES_AT);
  return {
    ok: true,
    message: 'Zoom Script Properties saved.',
    userId: String(userId || 'me')
  };
}

function verifyZoomScriptProperties() {
  var zoom = getZoomConfig_();
  return {
    configured: isZoomConfigured_(),
    accountIdSet: !!zoom.accountId,
    clientIdSet: !!zoom.clientId,
    clientSecretSet: !!zoom.clientSecret,
    userId: zoom.userId
  };
}

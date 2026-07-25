function scheduleCalendarSync() {
  deleteTriggersForHandler_('runCalendarSync');

  var config = getConfig_();
  ScriptApp.newTrigger('runCalendarSync')
    .timeBased()
    .atHour(config.syncHour)
    .everyDays(1)
    .inTimezone(config.timezone)
    .create();

  showToast_('Daily calendar sync scheduled at ' + config.syncHour + ':00 (' + config.timezone + ').');
}

function deleteTriggersForHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function scheduleCalendarSync() {
  deleteTriggersForHandler_('runCalendarSync');
  deleteTriggersForHandler_('runScheduledSync');
  deleteTriggersForHandler_('sendDailySummaryEmail');

  var config = getConfig_();
  ScriptApp.newTrigger('runScheduledSync')
    .timeBased()
    .atHour(config.syncHour)
    .everyDays(1)
    .inTimezone(config.timezone)
    .create();

  ScriptApp.newTrigger('sendDailySummaryEmail')
    .timeBased()
    .atHour(config.summaryHour)
    .everyDays(1)
    .inTimezone(config.timezone)
    .create();

  showToast_(
    'Daily sync at ' + config.syncHour + ':00; day summary email at ' +
      config.summaryHour + ':00 (' + config.timezone + ').'
  );
}

function deleteTriggersForHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

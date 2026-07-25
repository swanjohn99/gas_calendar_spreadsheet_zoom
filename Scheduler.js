function scheduleCalendarSync() {
  deleteTriggersForHandler_('runCalendarSync');
  deleteTriggersForHandler_('runScheduledSync');

  var config = getConfig_();
  ScriptApp.newTrigger('runScheduledSync')
    .timeBased()
    .atHour(config.syncHour)
    .everyDays(1)
    .inTimezone(config.timezone)
    .create();

  showToast_(
    'Daily scheduled sync at ' + config.syncHour + ':00 (' + config.timezone + '): ' +
      'import, archive, and Drive inbox organizer.'
  );
}

function deleteTriggersForHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

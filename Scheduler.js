function scheduleCalendarSync() {
  deleteTriggersForHandler_('runCalendarSync');
  deleteTriggersForHandler_('runScheduledSync');
  deleteTriggersForHandler_('sendDailySummaryEmail');

  var config = getConfig_();
  var hours = getScheduledSyncHours_();
  hours.forEach(function (hour) {
    ScriptApp.newTrigger('runScheduledSync')
      .timeBased()
      .atHour(hour)
      .everyDays(1)
      .inTimezone(config.timezone)
      .create();
  });

  showToast_(
    'Scheduled sync at ' + hours.join(':00, ') + ':00 (' + config.timezone + '). ' +
      'Last run of the day emails the combined report.'
  );
}

function getScheduledSyncHours_() {
  var config = getConfig_();
  var hours = (config.syncHours || []).map(function (hour) {
    return parseInt(hour, 10);
  }).filter(function (hour) {
    return !isNaN(hour) && hour >= 0 && hour <= 23;
  });

  hours.sort(function (a, b) {
    return a - b;
  });

  var unique = [];
  hours.forEach(function (hour) {
    if (unique.indexOf(hour) === -1) {
      unique.push(hour);
    }
  });

  return unique.length ? unique : [9, 17];
}

/**
 * True when another configured sync hour remains later today (after the current hour).
 */
function hasRemainingScheduledRunsToday_() {
  var config = getConfig_();
  var hours = getScheduledSyncHours_();
  var nowHour = parseInt(Utilities.formatDate(new Date(), config.timezone, 'H'), 10);
  for (var i = 0; i < hours.length; i++) {
    if (hours[i] > nowHour) {
      return true;
    }
  }
  return false;
}

function isLastScheduledRunOfDay_() {
  return !hasRemainingScheduledRunsToday_();
}

function deleteTriggersForHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

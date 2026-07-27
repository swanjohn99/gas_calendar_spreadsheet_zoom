/**
 * Per-run reports for scheduled sync jobs.
 * Last scheduled job of the day (no remaining SYNC_HOURS later today) emails the combined report.
 *
 * Script Properties key: DAILY_RUNS_<yyyy-MM-dd>
 */

var DAILY_SUMMARY = {
  PROPERTY_PREFIX: 'DAILY_RUNS_',
  SENT_PREFIX: 'DAILY_SUMMARY_SENT_'
};

function getDailyReportDateKey_(date) {
  var config = getConfig_();
  return Utilities.formatDate(date || new Date(), config.timezone, 'yyyy-MM-dd');
}

function getDailyRunsPropertyKey_(dateKey) {
  return DAILY_SUMMARY.PROPERTY_PREFIX + (dateKey || getDailyReportDateKey_());
}

function getDailySummarySentKey_(dateKey) {
  return DAILY_SUMMARY.SENT_PREFIX + (dateKey || getDailyReportDateKey_());
}

function loadDailyRuns_(dateKey) {
  dateKey = dateKey || getDailyReportDateKey_();
  var raw = PropertiesService.getScriptProperties().getProperty(getDailyRunsPropertyKey_(dateKey));
  if (!raw) {
    return { dateKey: dateKey, runs: [] };
  }
  try {
    var parsed = JSON.parse(raw);
    return {
      dateKey: dateKey,
      runs: parsed.runs || []
    };
  } catch (error) {
    Logger.log('Failed to parse daily runs: ' + error);
    return { dateKey: dateKey, runs: [] };
  }
}

function saveDailyRuns_(store) {
  PropertiesService.getScriptProperties().setProperty(
    getDailyRunsPropertyKey_(store.dateKey),
    JSON.stringify({ runs: store.runs || [] })
  );
}

/**
 * Append one scheduled-job report for today.
 */
function appendDailyRunReport_(runReport) {
  if (!runReport) return;
  var store = loadDailyRuns_();
  store.runs.push(runReport);
  saveDailyRuns_(store);
  return store;
}

/**
 * Combine all of today's scheduled run reports and email once (last job only).
 */
function sendDailySummaryEmailFromRuns_() {
  var config = getConfig_();
  var dateKey = getDailyReportDateKey_();
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(getDailySummarySentKey_(dateKey)) === '1') {
    Logger.log('Daily summary already sent for ' + dateKey);
    return { sent: false, reason: 'already_sent' };
  }

  var store = loadDailyRuns_(dateKey);
  var combined = combineDailyRunReports_(store);

  var recipient = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  if (!recipient) {
    Logger.log('No user email for daily summary');
    return { sent: false, reason: 'no_recipient' };
  }

  var subject = 'Calendar Tools day summary ' + dateKey;
  var body = formatDailySummaryEmailBody_(combined, config);
  GmailApp.sendEmail(recipient, subject, body);
  props.setProperty(getDailySummarySentKey_(dateKey), '1');
  Logger.log('Daily summary emailed to ' + recipient + ' (' + store.runs.length + ' runs)');
  notifyUser_('Day summary emailed to ' + recipient, 'Day Summary');
  return { sent: true, recipient: recipient, runs: store.runs.length };
}

/** Manual/test entry point — same as last-job send. */
function sendDailySummaryEmail() {
  return sendDailySummaryEmailFromRuns_();
}

function combineDailyRunReports_(store) {
  var combined = {
    dateKey: store.dateKey,
    runs: store.runs || [],
    newEvents: [],
    deletedEvents: [],
    organizedFiles: [],
    draftDetails: [],
    updatedCount: 0,
    archived: 0
  };

  (store.runs || []).forEach(function (run) {
    combined.updatedCount += Number(run.updatedCount || 0);
    combined.archived += Number(run.archived || 0);
    mergeUniqueByEventId_(combined.newEvents, run.newEvents || []);
    mergeUniqueByEventId_(combined.deletedEvents, run.deletedEvents || []);
    mergeOrganizeItems_(combined.organizedFiles, run.organizedFiles || []);
    (run.draftDetails || []).forEach(function (detail) {
      upsertDailyDraftDetail_(combined.draftDetails, detail);
    });
  });

  return combined;
}

function mergeUniqueByEventId_(target, items) {
  items.forEach(function (item) {
    var eventId = String(item.event_id || '');
    var exists = target.some(function (existing) {
      return String(existing.event_id || '') === eventId;
    });
    if (!exists) {
      target.push(item);
    }
  });
}

function mergeOrganizeItems_(target, items) {
  items.forEach(function (item) {
    var key = String(item.zoom_meeting_id || '') + '|' + String(item.artifact || '') + '|' +
      String(item.finalPath || '');
    var exists = target.some(function (existing) {
      return String(existing.zoom_meeting_id || '') + '|' + String(existing.artifact || '') + '|' +
        String(existing.finalPath || '') === key;
    });
    if (!exists) {
      target.push(item);
    }
  });
}

function upsertDailyDraftDetail_(list, detail) {
  var eventId = String(detail.event_id || '');
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].event_id || '') === eventId) {
      if (detail.draftSaved || detail.status === 'drafted' || detail.status === 'error' ||
          !list[i].draftSaved) {
        list[i] = detail;
      }
      return;
    }
  }
  list.push(detail);
}

function formatDailySummaryEmailBody_(report, config) {
  var lines = [];
  var hours = getScheduledSyncHours_();

  lines.push('Calendar Tools — day summary for ' + report.dateKey);
  lines.push('Timezone: ' + config.timezone);
  lines.push('Scheduled hours today: ' + hours.map(function (h) { return h + ':00'; }).join(', '));
  lines.push('Runs recorded: ' + (report.runs ? report.runs.length : 0));
  lines.push('');

  lines.push('=== New events ===');
  lines.push('Count: ' + report.newEvents.length);
  if (!report.newEvents.length) {
    lines.push('(none)');
  } else {
    report.newEvents.forEach(function (event, index) {
      lines.push(
        (index + 1) + '. [' + (event.sheet || '') + '] ' + (event.title || '(no title)') +
          ' | start ' + (event.start || '') +
          ' | meetingId ' + (event.zoom_meeting_id || '') +
          ' | rules.email ' + (event.email || '')
      );
    });
  }

  lines.push('');
  lines.push('=== Deleted events ===');
  lines.push('Count: ' + report.deletedEvents.length);
  if (!report.deletedEvents.length) {
    lines.push('(none)');
  } else {
    report.deletedEvents.forEach(function (event, index) {
      lines.push(
        (index + 1) + '. [' + (event.sheet || '') + '] ' + (event.title || '(no title)') +
          ' | start ' + (event.start || '') +
          ' | meetingId ' + (event.zoom_meeting_id || '')
      );
    });
  }

  lines.push('');
  lines.push('(Row updates for existing event IDs are omitted from this email.)');
  lines.push('Updated existing rows across runs: ' + report.updatedCount);

  lines.push('');
  lines.push('=== Email drafts ===');
  var drafted = (report.draftDetails || []).filter(function (d) {
    return d.draftSaved || d.status === 'drafted' || d.status === 'already_drafted';
  });
  var notDrafted = (report.draftDetails || []).filter(function (d) {
    return !(d.draftSaved || d.status === 'drafted' || d.status === 'already_drafted');
  });

  lines.push('Drafts saved: ' + drafted.length);
  drafted.forEach(function (detail, index) {
    lines.push((index + 1) + '. ' + (detail.title || detail.event_id || '(event)'));
    var files = (report.organizedFiles || []).filter(function (file) {
      return String(file.event_id || '') === String(detail.event_id || '') ||
        (detail.zoom_meeting_id &&
          String(file.zoom_meeting_id || '') === String(detail.zoom_meeting_id || ''));
    });
    if (files.length) {
      lines.push('   Organized files:');
      files.forEach(function (file) {
        lines.push(
          '   - ' + file.artifact + ': ' + file.finalPath +
            (file.status ? ' (' + file.status + ')' : '')
        );
      });
    } else {
      lines.push('   Organized files: (none in today’s runs; URLs may already have existed)');
    }
  });

  lines.push('');
  lines.push('=== No email draft — why ===');
  if (!notDrafted.length) {
    lines.push('(none)');
  } else {
    notDrafted.forEach(function (detail, index) {
      lines.push(
        (index + 1) + '. ' + (detail.title || detail.event_id || '(event)') +
          ' — ' + (detail.reason || detail.status || 'unknown')
      );
    });
  }

  lines.push('');
  lines.push('=== Organized files (all runs) ===');
  if (!report.organizedFiles.length) {
    lines.push('(none)');
  } else {
    report.organizedFiles.forEach(function (file, index) {
      lines.push(
        (index + 1) + '. ' + (file.title || file.zoom_meeting_id) +
          ' | ' + file.artifact + ' → ' + file.finalPath
      );
    });
  }

  lines.push('');
  lines.push('=== Per-run log ===');
  (report.runs || []).forEach(function (run, index) {
    lines.push(
      (index + 1) + '. ' + (run.runAt || '') +
        ' | hour ' + run.hour +
        ' | new ' + ((run.newEvents || []).length) +
        ' | deleted ' + ((run.deletedEvents || []).length) +
        ' | updated ' + (run.updatedCount || 0) +
        ' | drafts created ' + ((run.draftCounts && run.draftCounts.created) || 0)
    );
  });

  return lines.join('\n');
}

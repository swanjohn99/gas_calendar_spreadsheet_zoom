/**
 * Daily activity accumulator + end-of-day summary email.
 *
 * Script Properties key: DAILY_REPORT_<yyyy-MM-dd>
 */

var DAILY_SUMMARY = {
  PROPERTY_PREFIX: 'DAILY_REPORT_',
  SENT_PREFIX: 'DAILY_SUMMARY_SENT_'
};

function getDailyReportDateKey_(date) {
  var config = getConfig_();
  return Utilities.formatDate(date || new Date(), config.timezone, 'yyyy-MM-dd');
}

function getDailyReportPropertyKey_(dateKey) {
  return DAILY_SUMMARY.PROPERTY_PREFIX + dateKey;
}

function getDailySummarySentKey_(dateKey) {
  return DAILY_SUMMARY.SENT_PREFIX + dateKey;
}

function loadDailyReport_(dateKey) {
  dateKey = dateKey || getDailyReportDateKey_();
  var raw = PropertiesService.getScriptProperties().getProperty(getDailyReportPropertyKey_(dateKey));
  if (!raw) {
    return emptyDailyReport_(dateKey);
  }
  try {
    var parsed = JSON.parse(raw);
    parsed.dateKey = dateKey;
    parsed.importedEvents = parsed.importedEvents || [];
    parsed.organizedFiles = parsed.organizedFiles || [];
    parsed.draftDetails = parsed.draftDetails || [];
    parsed.importCounts = parsed.importCounts || { training: 0, nonTraining: 0, runs: 0 };
    return parsed;
  } catch (error) {
    Logger.log('Failed to parse daily report: ' + error);
    return emptyDailyReport_(dateKey);
  }
}

function emptyDailyReport_(dateKey) {
  return {
    dateKey: dateKey,
    importCounts: { training: 0, nonTraining: 0, runs: 0 },
    importedEvents: [],
    organizedFiles: [],
    draftDetails: []
  };
}

function saveDailyReport_(report) {
  PropertiesService.getScriptProperties().setProperty(
    getDailyReportPropertyKey_(report.dateKey),
    JSON.stringify(report)
  );
}

function recordDailyImport_(importResult) {
  if (!importResult) return;
  var report = loadDailyReport_();
  report.importCounts.training += Number(importResult.training || 0);
  report.importCounts.nonTraining += Number(importResult.nonTraining || 0);
  report.importCounts.runs += 1;

  var events = importResult.events || [];
  for (var i = 0; i < events.length; i++) {
    upsertDailyEvent_(report.importedEvents, events[i]);
  }
  saveDailyReport_(report);
}

function recordDailyOrganize_(driveResult) {
  if (!driveResult || !driveResult.items || !driveResult.items.length) return;
  var report = loadDailyReport_();
  for (var i = 0; i < driveResult.items.length; i++) {
    var item = driveResult.items[i];
    var key = String(item.zoom_meeting_id || '') + '|' + String(item.artifact || '') + '|' +
      String(item.finalPath || '');
    var exists = report.organizedFiles.some(function (existing) {
      return String(existing.zoom_meeting_id || '') + '|' + String(existing.artifact || '') + '|' +
        String(existing.finalPath || '') === key;
    });
    if (!exists) {
      report.organizedFiles.push(item);
    }
  }
  saveDailyReport_(report);
}

function recordDailyDrafts_(draftResult) {
  if (!draftResult || !draftResult.details || !draftResult.details.length) return;
  var report = loadDailyReport_();
  for (var i = 0; i < draftResult.details.length; i++) {
    upsertDailyDraftDetail_(report.draftDetails, draftResult.details[i]);
  }
  saveDailyReport_(report);
}

function upsertDailyEvent_(list, event) {
  var eventId = String(event.event_id || '');
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].event_id || '') === eventId) {
      list[i] = event;
      return;
    }
  }
  list.push(event);
}

function upsertDailyDraftDetail_(list, detail) {
  var eventId = String(detail.event_id || '');
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].event_id || '') === eventId) {
      // Prefer drafted / error over earlier skip once a later run succeeds.
      if (detail.draftSaved || detail.status === 'drafted' || detail.status === 'error' ||
          !list[i].draftSaved) {
        list[i] = detail;
      }
      return;
    }
  }
  list.push(detail);
}

/**
 * After a pipeline run: if at/after summary hour and not yet sent, email summary.
 * The dedicated summaryHour trigger also calls sendDailySummaryEmail.
 */
function maybeSendDailySummaryAfterRun_(pipelineResult) {
  var config = getConfig_();
  var now = new Date();
  var hour = parseInt(Utilities.formatDate(now, config.timezone, 'H'), 10);
  if (hour < Number(config.summaryHour || 17)) {
    return;
  }
  sendDailySummaryEmail();
}

/**
 * End-of-day summary email (scheduled at SUMMARY_HOUR, or after last late run).
 */
function sendDailySummaryEmail() {
  var config = getConfig_();
  var dateKey = getDailyReportDateKey_();
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(getDailySummarySentKey_(dateKey)) === '1') {
    Logger.log('Daily summary already sent for ' + dateKey);
    return { sent: false, reason: 'already_sent' };
  }

  var report = loadDailyReport_(dateKey);
  enrichDailyReportFromSheet_(report, config);

  var recipient = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  if (!recipient) {
    Logger.log('No user email for daily summary');
    return { sent: false, reason: 'no_recipient' };
  }

  var subject = 'Calendar Tools day summary ' + dateKey;
  var body = formatDailySummaryEmailBody_(report, config);
  GmailApp.sendEmail(recipient, subject, body);
  props.setProperty(getDailySummarySentKey_(dateKey), '1');
  Logger.log('Daily summary emailed to ' + recipient);
  notifyUser_('Day summary emailed to ' + recipient, 'Day Summary');
  return { sent: true, recipient: recipient };
}

/**
 * Fill draft/file status from current Coaching events sheet for meetings due today.
 */
function enrichDailyReportFromSheet_(report, config) {
  var sheet = getEventsSheet_();
  var data = getSheetDataObjects_(sheet, config.headers);
  var todayIso = getDailyReportDateKey_();

  data.rows.forEach(function (row) {
    var startIso = formatSheetDateOnly_(row.data.start, config.timezone);
    var due = isMeetingStartOnOrBeforeToday_(row.data.start, config.timezone);
    if (!due && startIso !== todayIso) {
      return;
    }

    var detail = {
      event_id: String(row.data.event_id || ''),
      title: String(row.data.title || ''),
      program: String(row.data.program || ''),
      zoom_meeting_id: String(row.data.zoom_meeting_id || ''),
      start: String(row.data.start || ''),
      email: String(row.data.email || ''),
      draftSaved: !!String(row.data.email_draft_saved || '').trim(),
      status: String(row.data.email_draft_saved || '').trim() ? 'drafted' : 'skipped',
      reason: ''
    };

    if (!detail.draftSaved) {
      if (!isYesEmailFlag_(row.data.email)) {
        detail.reason = 'noEmail flag (email=' + detail.email + ')';
      } else {
        var missing = getMissingArtifactUrls_(row.data);
        if (missing.length) {
          detail.reason = 'lack of saved files (' + missing.join(', ') + ')';
        } else if (!resolveRecipientEmail_(row.data)) {
          detail.reason = 'invalid or missing attendee_email';
        } else {
          detail.reason = 'draft not created yet';
        }
      }
    }

    upsertDailyDraftDetail_(report.draftDetails, detail);
  });
}

function formatDailySummaryEmailBody_(report, config) {
  var lines = [];
  var training = report.importCounts.training || 0;
  var nonTraining = report.importCounts.nonTraining || 0;
  var totalImported = training + nonTraining;

  lines.push('Calendar Tools — day summary for ' + report.dateKey);
  lines.push('Timezone: ' + config.timezone);
  lines.push('');
  lines.push('=== Imports ===');
  lines.push(
    'Events imported today: ' + totalImported +
      ' (coaching ' + training + ', non-coaching ' + nonTraining +
      ', sync runs ' + (report.importCounts.runs || 0) + ')'
  );

  if (report.importedEvents.length) {
    lines.push('');
    lines.push('Imported events:');
    report.importedEvents.forEach(function (event, index) {
      lines.push(
        (index + 1) + '. [' + (event.sheet || '') + '] ' + (event.title || '(no title)') +
          ' | start ' + (event.start || '') +
          ' | meetingId ' + (event.zoom_meeting_id || '') +
          ' | rules.email ' + (event.email || '')
      );
    });
  }

  lines.push('');
  lines.push('=== Email drafts ===');
  var drafted = report.draftDetails.filter(function (d) {
    return d.draftSaved || d.status === 'drafted' || d.status === 'already_drafted';
  });
  var notDrafted = report.draftDetails.filter(function (d) {
    return !(d.draftSaved || d.status === 'drafted' || d.status === 'already_drafted');
  });

  lines.push('Drafts saved: ' + drafted.length);
  drafted.forEach(function (detail, index) {
    lines.push((index + 1) + '. ' + (detail.title || detail.event_id || '(event)'));
    var files = report.organizedFiles.filter(function (file) {
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
      lines.push('   Organized files: (none recorded in today’s runs; URLs may already have existed)');
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
  lines.push('=== All organized files today ===');
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

  return lines.join('\n');
}

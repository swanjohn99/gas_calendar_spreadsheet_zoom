function runScheduledSync() {
  var importResult = syncCalendarEvents_();
  var archived = archiveOldEvents_();
  var pipeline;
  try {
    pipeline = runOrganizeAndDraftsPipeline_({ source: 'scheduled' });
  } catch (error) {
    Logger.log('Organize/drafts pipeline failed: ' + error);
    pipeline = {
      ok: false,
      source: 'scheduled',
      drive: {
        ok: false,
        copied: 0,
        skipped: 0,
        deduped: 0,
        items: [],
        message: 'Drive/drafts error: ' + error
      },
      drafts: { created: 0, skipped: 0, errors: 1, details: [], messages: [] },
      message: 'Pipeline error: ' + error
    };
  }

  // Always persist this run (including import new/deleted) even if organize/drafts did nothing.
  var runReport = buildScheduledRunReport_(importResult, pipeline, archived);
  appendDailyRunReport_(runReport);

  var summary = 'Scheduled sync: new ' + importResult.newEvents.length +
    ', updated ' + importResult.updatedCount +
    ', deleted ' + importResult.deletedEvents.length +
    '. Archived ' + archived + '. ' + pipeline.message;

  if (isLastScheduledRunOfDay_()) {
    var sent = sendDailySummaryEmailFromRuns_();
    summary += sent.sent
      ? ' Day summary emailed.'
      : ' Day summary not emailed (' + (sent.reason || 'unknown') + ').';
  } else {
    summary += ' Report saved; later scheduled runs remain today.';
  }

  Logger.log(summary);
  notifyUser_(summary, 'Scheduled Sync');
}

/**
 * Organize inbox, then create pending drafts.
 */
function runOrganizeAndDraftsPipeline_(options) {
  options = options || {};
  var drive = organizeDriveInbox_();
  var drafts = createEmailDraftsForPending_();

  var message = '';
  if (drive.ok) {
    message += drive.message + ' ';
  } else {
    message += 'Drive inbox skipped: ' + drive.message + ' ';
  }
  message += 'Drafts: ' + drafts.created + ' created, ' + drafts.skipped +
    ' skipped, ' + drafts.errors + ' errors.';

  return {
    ok: !!drive.ok,
    source: options.source || 'pipeline',
    drive: drive,
    drafts: drafts,
    message: message.trim()
  };
}

function buildScheduledRunReport_(importResult, pipeline, archived) {
  var config = getConfig_();
  var drive = pipeline.drive || {};
  return {
    runAt: formatDateValue_(new Date()),
    hour: parseInt(Utilities.formatDate(new Date(), config.timezone, 'H'), 10),
    archived: archived || 0,
    newEvents: importResult.newEvents || [],
    deletedEvents: importResult.deletedEvents || [],
    updatedCount: importResult.updatedCount || 0,
    organizedFiles: drive.items || [],
    draftDetails: (pipeline.drafts && pipeline.drafts.details) || [],
    driveMessage: drive.message || '',
    organizeCounts: {
      copied: drive.copied || 0,
      skipped: drive.skipped || 0,
      deduped: drive.deduped || 0,
      ok: !!drive.ok
    },
    draftCounts: {
      created: pipeline.drafts ? pipeline.drafts.created : 0,
      skipped: pipeline.drafts ? pipeline.drafts.skipped : 0,
      errors: pipeline.drafts ? pipeline.drafts.errors : 0
    }
  };
}

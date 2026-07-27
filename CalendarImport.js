function importCalendar() {
  runCalendarSync();
}

function runCalendarSync() {
  var importResult = syncCalendarEvents_();
  var archived = archiveOldEvents_();
  showToast_(
    'New ' + importResult.newEvents.length +
      ', updated ' + importResult.updatedCount +
      ', deleted ' + importResult.deletedEvents.length +
      '. Archived ' + archived + ' old events.'
  );
}

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
 * Organize inbox then create pending drafts.
 */
function runOrganizeAndDraftsPipeline_(options) {
  options = options || {};
  var drive = organizeDriveInbox_();
  var drafts = createEmailDraftsForPending_();

  var message = '';
  if (drive.ok) {
    message += drive.message;
  } else {
    message += 'Drive inbox skipped: ' + drive.message;
  }
  message += ' Drafts: ' + drafts.created + ' created, ' + drafts.skipped +
    ' skipped, ' + drafts.errors + ' errors.';

  return {
    ok: !!drive.ok,
    source: options.source || 'pipeline',
    drive: drive,
    drafts: drafts,
    message: message
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

function syncCalendarEvents_() {
  var config = getConfig_();
  var sheet = getEventsSheet_();
  var existing = getSheetDataObjects_(sheet, config.headers);
  var existingById = buildExistingById_(existing.rows);

  var calendar = CalendarApp.getCalendarById(config.calendarId);
  if (!calendar) {
    throw new Error('Calendar not found: ' + config.calendarId);
  }

  var now = new Date();
  var start = new Date(now.getTime());
  start.setDate(start.getDate() - config.lookbackDays);
  var end = new Date(now.getTime());
  end.setDate(end.getDate() + config.lookaheadDays);

  var events = calendar.getEvents(start, end);
  var rejected = {};
  var rulesMap = buildRulesMap_();
  var newEvents = [];
  var updatedCount = 0;

  events.forEach(function (event) {
    var eventId = event.getId();
    var hasZoom = hasZoomLinkInLocation_(event.getLocation(), config.zoomLocationPattern);

    if (!hasZoom) {
      rejected[eventId] = true;
      return;
    }

    var mapped = mapCalendarEventToRow_(event, rulesMap);
    var action = upsertEventRow_(
      sheet,
      mapped,
      existingById,
      config.headers,
      config.preservedColumns
    );

    if (action === 'created') {
      newEvents.push({
        event_id: mapped.event_id,
        title: mapped.title,
        program: mapped.program,
        zoom_meeting_id: mapped.zoom_meeting_id,
        start: mapped.start,
        email: mapped.email,
        sheet: config.eventsSheetName
      });
    } else {
      updatedCount++;
    }
  });

  var deletedEvents = removeRejectedEventRows_(
    sheet,
    existing.rows,
    rejected,
    config.eventsSheetName
  );

  return {
    newEvents: newEvents,
    deletedEvents: deletedEvents,
    updatedCount: updatedCount
  };
}

function buildExistingById_(rows) {
  var existingById = {};
  rows.forEach(function (row) {
    if (row.data.event_id) {
      existingById[row.data.event_id] = row;
    }
  });
  return existingById;
}

function upsertEventRow_(sheet, mapped, existingById, headers, preservedColumns) {
  var current = existingById[mapped.event_id];

  if (current) {
    preservedColumns.forEach(function (column) {
      if (current.data[column]) {
        mapped[column] = current.data[column];
      }
    });
    writeRowObject_(sheet, current.sheetRow, mapped, headers);
    return 'updated';
  }

  appendRowObject_(sheet, mapped, headers);
  existingById[mapped.event_id] = { sheetRow: sheet.getLastRow(), data: mapped };
  return 'created';
}

function hasZoomLinkInLocation_(location, pattern) {
  return new RegExp(pattern, 'i').test(location || '');
}

function extractZoomMeetingId_(location) {
  var match = String(location || '').match(/\/j\/(\d+)/i);
  return match ? match[1] : '';
}

function removeRejectedEventRows_(sheet, rows, rejectedIds, sheetName) {
  var deleted = [];
  var rowsToDelete = rows
    .filter(function (row) {
      return rejectedIds[row.data.event_id];
    })
    .sort(function (a, b) {
      return b.sheetRow - a.sheetRow;
    });

  rowsToDelete.forEach(function (row) {
    deleted.push({
      event_id: String(row.data.event_id || ''),
      title: String(row.data.title || ''),
      program: String(row.data.program || ''),
      zoom_meeting_id: String(row.data.zoom_meeting_id || ''),
      start: String(row.data.start || ''),
      email: String(row.data.email || ''),
      sheet: sheetName || sheet.getName()
    });
    sheet.deleteRow(row.sheetRow);
  });

  return deleted;
}

function mapCalendarEventToRow_(event, rulesMap) {
  var title = event.getTitle() || '';
  var program = parseEventProgram_(title);
  var rule = lookupRuleByTitle_(rulesMap || {}, title);

  return {
    event_id: event.getId(),
    title: title,
    program: program,
    location: event.getLocation() || '',
    zoom_meeting_id: extractZoomMeetingId_(event.getLocation()),
    start: formatDateValue_(event.getStartTime()),
    end: formatDateValue_(event.getEndTime()),
    attendee_email: getAttendeeEmail_(event),
    updated: formatDateValue_(event.getLastUpdated()),
    email: rule ? String(rule.email || '').trim() : '',
    email_draft_saved: '',
    video_url: '',
    pdf_url: '',
    audio_url: '',
    transcript_url: '',
    chat_url: ''
  };
}

function parseEventProgram_(rawTitle) {
  var title = String(rawTitle || '').trim();
  if (!title) {
    return '';
  }

  var separatorIndex = -1;
  for (var i = 0; i < title.length; i++) {
    if (title.charAt(i) === '-' || title.charAt(i) === ':') {
      separatorIndex = i;
      break;
    }
  }

  if (separatorIndex === -1) {
    return title;
  }

  return title.substring(0, separatorIndex).trim();
}

function toCalendarApiEventId_(eventId) {
  var id = String(eventId || '');
  var atIndex = id.indexOf('@');
  return atIndex === -1 ? id : id.substring(0, atIndex);
}

function getAttendeeEmail_(event) {
  var apiEventId = toCalendarApiEventId_(event.getId());
  var calendarIds = [];
  var originalCalendarId = event.getOriginalCalendarId();
  var configuredCalendarId = getConfig_().calendarId;

  if (originalCalendarId) {
    calendarIds.push(originalCalendarId);
  }
  if (configuredCalendarId && calendarIds.indexOf(configuredCalendarId) === -1) {
    calendarIds.push(configuredCalendarId);
  }

  for (var i = 0; i < calendarIds.length; i++) {
    try {
      var apiEvent = Calendar.Events.get(calendarIds[i], apiEventId);
      var guests = (apiEvent.attendees || []).filter(function (guest) {
        return !guest.organizer && !guest.self;
      });

      if (guests.length >= 1) {
        return guests[0].email || '';
      }
    } catch (error) {
      // Try next calendar; CalendarApp guest list is the final fallback.
    }
  }

  var guestList = event.getGuestList();
  var emails = guestList.map(function (guest) {
    return guest.getEmail();
  }).filter(Boolean);

  return emails.length >= 1 ? emails[0] : '';
}

function writeRowObject_(sheet, sheetRow, obj, headers) {
  sheet.getRange(sheetRow, 1, 1, headers.length).setValues([objectToRow_(obj, headers)]);
}

function appendRowObject_(sheet, obj, headers) {
  sheet.appendRow(objectToRow_(obj, headers));
}

function importCalendar() {
  runCalendarSync();
}

function runCalendarSync() {
  var importCounts = syncCalendarEvents_();
  var archived = archiveOldEvents_();
  showToast_(
    'Imported ' + importCounts.training + ' coaching, ' + importCounts.nonTraining +
      ' non-coaching. Archived ' + archived + ' old events.'
  );
}

function runScheduledSync() {
  var importCounts = syncCalendarEvents_();
  var archived = archiveOldEvents_();
  var drive = organizeDriveInbox_();
  var summary = 'Scheduled sync: imported ' + importCounts.training + ' coaching, ' +
    importCounts.nonTraining + ' non-coaching. Archived ' + archived + ' old events.';

  if (drive.ok) {
    summary += ' ' + drive.message;
  } else {
    summary += ' Drive inbox skipped: ' + drive.message;
  }

  var drafts = createEmailDraftsForPending_();
  summary += ' Drafts: ' + drafts.created + ' created, ' + drafts.skipped + ' skipped, ' + drafts.errors + ' errors.';

  Logger.log(summary);
  notifyUser_(summary, 'Scheduled Sync');
}

function syncCalendarEvents_() {
  var config = getConfig_();
  var trainingSheet = getEventsSheet_();
  var nonTrainingSheet = getNonTrainingEventsSheet_();
  var trainingExisting = getSheetDataObjects_(trainingSheet, config.headers);
  var nonTrainingExisting = getSheetDataObjects_(nonTrainingSheet, config.nonTrainingHeaders);
  var trainingById = buildExistingById_(trainingExisting.rows);
  var nonTrainingById = buildExistingById_(nonTrainingExisting.rows);

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
  var trainingCount = 0;
  var nonTrainingCount = 0;
  var trainingRejected = {};
  var nonTrainingRejected = {};
  var calendarColorCache = {};
  var rulesMap = buildRulesMap_();

  events.forEach(function (event) {
    var eventId = event.getId();
    var hasZoom = hasZoomLinkInLocation_(event.getLocation(), config.zoomLocationPattern);

    if (!hasZoom) {
      trainingRejected[eventId] = true;
      nonTrainingRejected[eventId] = true;
      return;
    }

    var mapped = mapCalendarEventToRow_(event, rulesMap);

    if (isGreenEventColor_(event, calendarColorCache)) {
      upsertEventRow_(
        trainingSheet,
        mapped,
        trainingById,
        config.headers,
        config.preservedColumns
      );
      nonTrainingRejected[eventId] = true;
      trainingCount++;
      return;
    }

    upsertEventRow_(
      nonTrainingSheet,
      mapped,
      nonTrainingById,
      config.nonTrainingHeaders,
      config.nonTrainingPreservedColumns
    );
    trainingRejected[eventId] = true;
    nonTrainingCount++;
  });

  removeRejectedEventRows_(trainingSheet, trainingExisting.rows, trainingRejected);
  removeRejectedEventRows_(nonTrainingSheet, nonTrainingExisting.rows, nonTrainingRejected);

  return {
    training: trainingCount,
    nonTraining: nonTrainingCount
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
    return;
  }

  appendRowObject_(sheet, mapped, headers);
}

function isGreenEventColor_(event, calendarColorCache) {
  var config = getConfig_();
  var eventColor = String(event.getColor() || '');

  if (eventColor) {
    return config.coachingEventColorIds.indexOf(eventColor) !== -1;
  }

  var calendarId = event.getOriginalCalendarId() || config.calendarId;
  var calendarColorId = getCalendarColorId_(calendarId, calendarColorCache || {});
  return config.coachingCalendarColorIds.indexOf(calendarColorId) !== -1;
}

function getCalendarColorId_(calendarId, cache) {
  var id = String(calendarId || '');
  if (!id) {
    return '';
  }
  if (Object.prototype.hasOwnProperty.call(cache, id)) {
    return cache[id];
  }

  try {
    var calendarListEntry = Calendar.CalendarList.get(id);
    cache[id] = String((calendarListEntry && calendarListEntry.colorId) || '');
  } catch (error) {
    Logger.log('Calendar color lookup failed for ' + id + ': ' + error);
    cache[id] = '';
  }

  return cache[id];
}

function hasZoomLinkInLocation_(location, pattern) {
  return new RegExp(pattern, 'i').test(location || '');
}

function extractZoomMeetingId_(location) {
  var match = String(location || '').match(/\/j\/(\d+)/i);
  return match ? match[1] : '';
}

function removeRejectedEventRows_(sheet, rows, rejectedIds) {
  var rowsToDelete = rows
    .filter(function (row) {
      return rejectedIds[row.data.event_id];
    })
    .map(function (row) {
      return row.sheetRow;
    })
    .sort(function (a, b) {
      return b - a;
    });

  rowsToDelete.forEach(function (sheetRow) {
    sheet.deleteRow(sheetRow);
  });
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

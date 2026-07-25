function importCalendar() {
  runCalendarSync();
}

function runCalendarSync() {
  var imported = syncCalendarEvents_();
  var archived = archiveOldEvents_();
  showToast_('Imported ' + imported + ' events. Archived ' + archived + ' old events.');
}

function syncCalendarEvents_() {
  var config = getConfig_();
  var sheet = getEventsSheet_();
  var existing = getSheetDataObjects_(sheet);
  var existingById = {};

  existing.rows.forEach(function (row) {
    if (row.data.event_id) {
      existingById[row.data.event_id] = row;
    }
  });

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
  var upsertCount = 0;
  var rejectedIds = {};

  events.forEach(function (event) {
    var eventId = event.getId();

    if (!isGreenEventColor_(event)) {
      rejectedIds[eventId] = true;
      return;
    }

    if (!hasZoomLinkInLocation_(event.getLocation(), config.zoomLocationPattern)) {
      rejectedIds[eventId] = true;
      return;
    }

    var mapped = mapCalendarEventToRow_(event);
    var current = existingById[mapped.event_id];

    if (current) {
      config.preservedColumns.forEach(function (column) {
        if (current.data[column]) {
          mapped[column] = current.data[column];
        }
      });
      writeRowObject_(sheet, current.sheetRow, mapped, config.headers);
    } else {
      appendRowObject_(sheet, mapped, config.headers);
    }

    upsertCount++;
  });

  removeRejectedEventRows_(sheet, existing.rows, rejectedIds);

  return upsertCount;
}

function isGreenEventColor_(event) {
  var allowedIds = getConfig_().allowedEventColorIds;
  var eventColor = String(event.getColor() || '');
  return allowedIds.indexOf(eventColor) !== -1;
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

function mapCalendarEventToRow_(event) {
  var title = event.getTitle() || '';
  var parsedTitle = parseEventTitle_(title);

  return {
    event_id: event.getId(),
    title: title,
    meeting_type: parsedTitle.meeting_type,
    attendee_first_name: parsedTitle.attendee_first_name,
    attendee_last_name: parsedTitle.attendee_last_name,
    description: event.getDescription() || '',
    location: event.getLocation() || '',
    zoom_meeting_id: extractZoomMeetingId_(event.getLocation()),
    start: formatDateValue_(event.getStartTime()),
    end: formatDateValue_(event.getEndTime()),
    attendee_email: getAttendeeEmail_(event),
    html_link: buildEventHtmlLink_(event),
    updated: formatDateValue_(event.getLastUpdated()),
    email_draft_saved: '',
    video_url: '',
    pdf_url: '',
    audio_url: '',
    transcript_url: '',
    chat_url: ''
  };
}

function parseEventTitle_(rawTitle) {
  var title = String(rawTitle || '').trim();
  if (!title) {
    return {
      meeting_type: '',
      attendee_first_name: '',
      attendee_last_name: ''
    };
  }

  var separatorIndex = -1;
  for (var i = 0; i < title.length; i++) {
    if (title.charAt(i) === '-' || title.charAt(i) === ':') {
      separatorIndex = i;
      break;
    }
  }

  if (separatorIndex === -1) {
    return {
      meeting_type: title,
      attendee_first_name: '',
      attendee_last_name: ''
    };
  }

  var meetingType = title.substring(0, separatorIndex).trim();
  var namePart = title.substring(separatorIndex + 1).trim();
  var nameTokens = namePart.split(/\s+/).filter(Boolean);

  return {
    meeting_type: meetingType,
    attendee_first_name: nameTokens[0] || '',
    attendee_last_name: nameTokens.slice(1).join(' ')
  };
}

function getAttendeeEmail_(event) {
  try {
    var apiEvent = Calendar.Events.get(event.getOriginalCalendarId(), event.getId());
    var guests = (apiEvent.attendees || []).filter(function (guest) {
      return !guest.organizer && !guest.self;
    });

    if (guests.length >= 1) {
      return guests[0].email || '';
    }
  } catch (error) {
    Logger.log('Calendar API attendee lookup failed: ' + error);
    var guestList = event.getGuestList();
    var emails = guestList.map(function (guest) {
      return guest.getEmail();
    }).filter(Boolean);

    if (emails.length >= 1) {
      return emails[0];
    }
  }

  return '';
}

function buildEventHtmlLink_(event) {
  try {
    var eventId = event.getId();
    var calendarId = event.getOriginalCalendarId();
    var encoded = Utilities.base64EncodeWebSafe(eventId + ' ' + calendarId);
    return 'https://www.google.com/calendar/event?eid=' + encoded;
  } catch (error) {
    return '';
  }
}

function writeRowObject_(sheet, sheetRow, obj, headers) {
  sheet.getRange(sheetRow, 1, 1, headers.length).setValues([objectToRow_(obj, headers)]);
}

function appendRowObject_(sheet, obj, headers) {
  sheet.appendRow(objectToRow_(obj, headers));
}

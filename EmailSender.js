function createEmailDraftsForSelection() {
  var config = getConfig_();
  var sheet = getEventsSheet_();
  var selectedRows = getSelectedDataRows_(sheet);

  if (!selectedRows.length) {
    showToast_('Select one or more event rows on the Events sheet first.');
    return;
  }

  var headerMap = getHeaderIndexMap_(sheet);
  var emailDraftSavedIndex = headerMap.email_draft_saved;
  var created = 0;
  var skipped = 0;
  var errors = 0;
  var skipMessages = [];

  selectedRows.forEach(function (row) {
    if (row.data.email_draft_saved) {
      skipped++;
      var skipMsg = logDraftSkip_(row, 'already_drafted', 'email_draft_saved=' + row.data.email_draft_saved);
      skipMessages.push(skipMsg);
      return;
    }

    var recipientEmail = resolveRecipientEmail_(row.data);
    if (!recipientEmail) {
      skipped++;
      var invalidEmailMsg = logDraftSkip_(
        row,
        'invalid_attendee_email',
        'attendee_email=' + formatLogValue_(row.data.attendee_email)
      );
      skipMessages.push(invalidEmailMsg);
      return;
    }

    var startDate = parseSheetDate_(row.data.start);
    if (startDate && startDate.getTime() > new Date().getTime()) {
      skipped++;
      return;
    }

    var missingArtifacts = getMissingArtifactUrls_(row.data);
    if (missingArtifacts.length) {
      skipped++;
      logDraftSkip_(row, 'missing_artifact_urls', 'missing=' + missingArtifacts.join(', '));
      skipMessages.push('row ' + row.sheetRow + ': missing_artifact_urls');
      return;
    }

    try {
      var subject = buildCoachingEmailSubject_(row.data, config);
      var plainBody = buildCoachingEmailPlainBody_(row.data, config);
      var htmlBody = buildCoachingEmailHtmlBody_(row.data, config);
      var options = { htmlBody: htmlBody };
      var pdfFileId = extractDriveFileIdFromUrl_(row.data.pdf_url);

      if (pdfFileId) {
        options.attachments = [DriveApp.getFileById(pdfFileId).getBlob()];
      }

      GmailApp.createDraft(recipientEmail, subject, plainBody, options);
      sheet.getRange(row.sheetRow, emailDraftSavedIndex + 1).setValue(formatDateValue_(new Date()));
      Logger.log(formatDraftLog_(row, 'draft_created', 'to=' + recipientEmail));
      created++;
    } catch (error) {
      var errorMsg = formatDraftLog_(
        row,
        'draft_failed',
        String(error) +
          ' | resolved_to=' + recipientEmail +
          ' | attendee_email=' + formatLogValue_(row.data.attendee_email)
      );
      Logger.log(errorMsg);
      skipMessages.push('row ' + row.sheetRow + ': draft_failed');
      errors++;
    }
  });

  var summary = 'Drafts created: ' + created + ', skipped: ' + skipped + ', errors: ' + errors + '.';
  if (skipMessages.length) {
    summary += ' ' + skipMessages.slice(0, 3).join(' | ');
  }
  showToast_(summary);
}

function resolveRecipientEmail_(rowData) {
  return normalizeEmailAddress_(rowData.attendee_email);
}

function normalizeEmailAddress_(value) {
  var text = String(value || '').trim();
  if (!text) {
    return '';
  }

  var angleMatch = text.match(/<([^>]+@[^>]+)>/);
  if (angleMatch) {
    text = angleMatch[1].trim();
  }

  var emailMatch = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  if (!emailMatch) {
    return '';
  }

  var email = emailMatch[0].trim().toLowerCase();
  if (email.indexOf('http') === 0 || email.indexOf('www.') === 0) {
    return '';
  }

  return isValidEmailAddress_(email) ? email : '';
}

function isValidEmailAddress_(email) {
  return /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(email);
}

function formatLogValue_(value) {
  return JSON.stringify(String(value === undefined || value === null ? '' : value));
}

function logDraftSkip_(row, reason, details) {
  var message = formatDraftLog_(row, reason, details);
  Logger.log(message);
  return 'row ' + row.sheetRow + ': ' + reason;
}

function logDraftWarn_(row, reason, details) {
  Logger.log(formatDraftLog_(row, reason, details));
}

function formatDraftLog_(row, reason, details) {
  var label = row.data.meeting_type || row.data.event_id || 'unknown';
  var message = 'Row ' + row.sheetRow + ' | event: ' + label + ' | reason: ' + reason;
  if (details) {
    message += ' | ' + details;
  }
  return message;
}

function buildCoachingEmailSubject_(rowData, config) {
  var startDate = parseSheetDate_(rowData.start) || new Date();
  var formattedDate = Utilities.formatDate(startDate, config.timezone, config.emailSubjectDateFormat);
  return config.emailSubjectPrefix + formattedDate;
}

function buildCoachingEmailPlainBody_(rowData, config) {
  var firstName = getAttendeeFirstName_(rowData);
  var sessionName = getSessionName_(rowData);
  var meetingDay = getMeetingDayPhrase_(parseSheetDate_(rowData.start), config.timezone);
  var recordingUrl = String(rowData.video_url || '').trim();
  var lines = [
    applyTemplate_(config.emailBodyLine0, { firstName: firstName }),
    '',
    applyTemplate_(config.emailBodyLine1, {
      sessionName: sessionName,
      meetingDay: meetingDay
    })
  ];

  if (recordingUrl) {
    lines.push('', config.emailBodyRecordingIntro, recordingUrl);
  }

  lines.push(
    '',
    config.emailBodyParagraphSummary,
    '',
    config.emailBodyParagraphNext,
    '',
    config.emailBodyParagraphQuestions,
    '',
    config.emailBodySignoff
  );

  return lines.join('\n');
}

function buildCoachingEmailHtmlBody_(rowData, config) {
  var firstName = getAttendeeFirstName_(rowData);
  var sessionName = getSessionName_(rowData);
  var meetingDay = getMeetingDayPhrase_(parseSheetDate_(rowData.start), config.timezone);
  var recordingUrl = String(rowData.video_url || '').trim();
  var parts = [
    '<p>' + escapeHtml_(applyTemplate_(config.emailBodyLine0, { firstName: firstName })) + '</p>',
    '<p>' + escapeHtml_(applyTemplate_(config.emailBodyLine1, {
      sessionName: sessionName,
      meetingDay: meetingDay
    })) + '</p>'
  ];

  if (recordingUrl) {
    var linkHtml = '<a href="' + escapeHtml_(recordingUrl) + '">' +
      escapeHtml_(config.emailBodyRecordingLinkText) + '</a>';
    parts.push('<p>' + escapeHtml_(config.emailBodyRecordingIntro) + '<br>' + linkHtml + '</p>');
  }

  parts.push(
    '<p>' + escapeHtml_(config.emailBodyParagraphSummary) + '</p>',
    '<p>' + escapeHtml_(config.emailBodyParagraphNext) + '</p>',
    '<p>' + escapeHtml_(config.emailBodyParagraphQuestions) + '</p>',
    '<p>' + escapeHtml_(config.emailBodySignoff).replace(/\n/g, '<br>') + '</p>'
  );

  return parts.join('');
}

function getMissingArtifactUrls_(rowData) {
  var fields = ['video_url', 'pdf_url', 'audio_url', 'transcript_url'];
  return fields.filter(function (field) {
    return !String(rowData[field] || '').trim();
  });
}

function getAttendeeFirstName_(rowData) {
  var firstName = String(rowData.attendee_first_name || '').trim();
  if (firstName) {
    return firstName;
  }

  return 'there';
}

function getSessionName_(rowData) {
  return String(rowData.meeting_type || '').trim() || 'coaching';
}

function getMeetingDayPhrase_(startDate, timezone) {
  if (!startDate) {
    return 'earlier';
  }

  var today = stripToLocalMidnight_(new Date(), timezone);
  var meetingDay = stripToLocalMidnight_(startDate, timezone);
  var diffDays = Math.round((today.getTime() - meetingDay.getTime()) / 86400000);

  if (diffDays <= 0) {
    return 'today';
  }
  if (diffDays === 1) {
    return 'yesterday';
  }

  var todayWeekStart = getWeekStartSunday_(today);
  var meetingWeekStart = getWeekStartSunday_(meetingDay);

  if (meetingWeekStart.getTime() === todayWeekStart.getTime()) {
    return 'earlier this week';
  }

  var lastWeekStart = new Date(todayWeekStart);
  lastWeekStart.setDate(lastWeekStart.getDate() - 7);
  if (meetingWeekStart.getTime() === lastWeekStart.getTime()) {
    return 'last week';
  }

  return 'earlier';
}

function stripToLocalMidnight_(date, timezone) {
  var year = parseInt(Utilities.formatDate(date, timezone, 'yyyy'), 10);
  var month = parseInt(Utilities.formatDate(date, timezone, 'MM'), 10) - 1;
  var day = parseInt(Utilities.formatDate(date, timezone, 'dd'), 10);
  return new Date(year, month, day);
}

function getWeekStartSunday_(date) {
  var weekStart = new Date(date);
  weekStart.setDate(weekStart.getDate() - weekStart.getDay());
  weekStart.setHours(0, 0, 0, 0);
  return weekStart;
}

function applyTemplate_(template, values) {
  return String(template).replace(/\{(\w+)\}/g, function (_, key) {
    return values[key] !== undefined && values[key] !== null ? values[key] : '';
  });
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

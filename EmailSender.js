function createEmailDraftsForSelection() {
  var sheet = getEventsSheet_();
  var selectedRows = getSelectedDataRows_(sheet);

  if (!selectedRows.length) {
    showToast_('Select one or more event rows on the events sheet first.');
    return;
  }

  var result = createEmailDraftsForRows_(sheet, selectedRows);
  var summary = 'Drafts created: ' + result.created + ', skipped: ' + result.skipped + ', errors: ' + result.errors + '.';
  if (result.messages.length) {
    summary += ' ' + result.messages.slice(0, 3).join(' | ');
  }
  showToast_(summary);
}

function createEmailDraftsForPending_() {
  var config = getConfig_();
  var sheet = getEventsSheet_();
  var data = getSheetDataObjects_(sheet, config.headers);
  var dueRows = data.rows.filter(function (row) {
    return isMeetingStartOnOrBeforeToday_(row.data.start, config.timezone);
  });

  return createEmailDraftsForRows_(sheet, dueRows);
}

function createEmailDraftsForRows_(sheet, rows) {
  var config = getConfig_();
  var headerMap = getHeaderIndexMap_(sheet);
  var emailDraftSavedIndex = headerMap.email_draft_saved;
  var rulesList = buildRulesList_();
  var created = 0;
  var skipped = 0;
  var errors = 0;
  var messages = [];
  var details = [];

  rows.forEach(function (row) {
    var emailFlag = getEmailFlag_(row.data);
    var detailBase = {
      event_id: String(row.data.event_id || ''),
      title: String(row.data.title || ''),
      zoom_meeting_id: String(row.data.zoom_meeting_id || ''),
      start: String(row.data.start || ''),
      emailFlag: emailFlag,
      sheetRow: row.sheetRow
    };

    if (row.data.email_draft_saved) {
      skipped++;
      messages.push(logDraftSkip_(row, 'already_drafted', 'email_draft_saved=' + row.data.email_draft_saved));
      details.push(Object.assign({}, detailBase, {
        status: 'already_drafted',
        reason: 'email_draft_saved=' + row.data.email_draft_saved,
        draftSaved: true
      }));
      return;
    }

    if (!isYesEmailFlag_(emailFlag)) {
      skipped++;
      messages.push(logDraftSkip_(
        row,
        'email_flag_not_yes',
        'emailFlag=' + formatLogValue_(emailFlag)
      ));
      details.push(Object.assign({}, detailBase, {
        status: 'skipped',
        reason: 'noEmail flag (email (yes or no)=' + emailFlag + ')',
        draftSaved: false
      }));
      return;
    }

    var recipientEmail = resolveRecipientEmail_(row.data);
    if (!recipientEmail) {
      skipped++;
      messages.push(logDraftSkip_(
        row,
        'invalid_attendee_email',
        'attendee_email=' + formatLogValue_(row.data.attendee_email)
      ));
      details.push(Object.assign({}, detailBase, {
        status: 'skipped',
        reason: 'invalid or missing attendee_email',
        draftSaved: false
      }));
      return;
    }

    var startDate = parseSheetDate_(row.data.start);
    if (startDate && startDate.getTime() > new Date().getTime()) {
      skipped++;
      details.push(Object.assign({}, detailBase, {
        status: 'skipped',
        reason: 'meeting starts in the future',
        draftSaved: false
      }));
      return;
    }

    var missingArtifacts = getMissingArtifactUrls_(row.data);
    if (missingArtifacts.length) {
      skipped++;
      logDraftSkip_(row, 'missing_artifact_urls', 'missing=' + missingArtifacts.join(', '));
      messages.push('row ' + row.sheetRow + ': missing_artifact_urls');
      details.push(Object.assign({}, detailBase, {
        status: 'skipped',
        reason: 'lack of saved files (' + missingArtifacts.join(', ') + ')',
        draftSaved: false
      }));
      return;
    }

    try {
      var subject = buildCoachingEmailSubject_(row.data, config);
      var firstName = getRulesFirstName_(row.data, rulesMap);
      var plainBody = buildCoachingEmailPlainBody_(row.data, config, firstName);
      var htmlBody = buildCoachingEmailHtmlBody_(row.data, config, firstName);
      var options = {
        htmlBody: htmlBody,
        attachments: buildCoachingEmailAttachments_(row.data)
      };

      GmailApp.createDraft(recipientEmail, subject, plainBody, options);
      var savedAt = formatDateValue_(new Date());
      sheet.getRange(row.sheetRow, emailDraftSavedIndex + 1).setValue(savedAt);
      row.data.email_draft_saved = savedAt;
      Logger.log(formatDraftLog_(row, 'draft_created', 'to=' + recipientEmail));
      created++;
      details.push(Object.assign({}, detailBase, {
        status: 'drafted',
        reason: '',
        draftSaved: true,
        recipient: recipientEmail
      }));
    } catch (error) {
      var errorMsg = formatDraftLog_(
        row,
        'draft_failed',
        String(error) +
          ' | resolved_to=' + recipientEmail +
          ' | attendee_email=' + formatLogValue_(row.data.attendee_email)
      );
      Logger.log(errorMsg);
      messages.push('row ' + row.sheetRow + ': draft_failed');
      errors++;
      details.push(Object.assign({}, detailBase, {
        status: 'error',
        reason: String(error),
        draftSaved: false
      }));
    }
  });

  return {
    created: created,
    skipped: skipped,
    errors: errors,
    messages: messages,
    details: details
  };
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
  var label = row.data.title || row.data.event_id || 'unknown';
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

function buildCoachingEmailPlainBody_(rowData, config, firstName) {
  firstName = firstName || getRulesFirstName_(rowData, buildRulesMap_());
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

function buildCoachingEmailHtmlBody_(rowData, config, firstName) {
  firstName = firstName || getRulesFirstName_(rowData, buildRulesMap_());
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

function buildCoachingEmailAttachments_(rowData) {
  var requiredFields = [
    { column: 'pdf_url' },
    { column: 'audio_url' },
    { column: 'transcript_url' }
  ];
  var attachments = [];

  requiredFields.forEach(function (field) {
    var blob = getDriveBlobFromUrl_(rowData[field.column]);
    if (!blob) {
      throw new Error('Could not load Drive attachment for ' + field.column);
    }
    attachments.push(blob);
  });

  var chatBlob = getDriveBlobFromUrl_(rowData.chat_url);
  if (chatBlob) {
    attachments.push(chatBlob);
  }

  return attachments;
}

function getDriveBlobFromUrl_(url) {
  var fileId = extractDriveFileIdFromUrl_(url);
  if (!fileId) {
    return null;
  }
  return DriveApp.getFileById(fileId).getBlob();
}

function getRulesFirstName_(rowData, rulesMap) {
  var rule = lookupRuleByTitle_(rulesMap || buildRulesMap_(), rowData.title);
  var firstName = rule ? String(rule.firstName || '').trim() : '';
  if (firstName) {
    return firstName;
  }

  return 'there';
}

function getSessionName_(rowData) {
  return String(rowData.title || '').trim() || 'coaching';
}

function getEmailFlag_(rowData) {
  var column = getConfig_().emailFlagColumn || 'email (yes or no)';
  return String((rowData && rowData[column]) || '').trim();
}

function isYesEmailFlag_(value) {
  var expected = String(getConfig_().rulesYesEmail || 'yesEmail').trim().toLowerCase();
  return String(value || '').trim().toLowerCase() === expected;
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

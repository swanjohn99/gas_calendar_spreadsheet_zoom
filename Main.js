function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Tools')
    .addItem('Import Calendar', 'importCalendar')
    .addItem('Schedule', 'scheduleCalendarSync')
    .addSeparator()
    .addItem('Organize Drive Inbox', 'organizeDriveInbox')
    .addItem('Create Email Drafts', 'createEmailDraftsForSelection')
    .addItem('Organize Inbox + Email Drafts', 'organizeInboxAndCreateDrafts')
    .addToUi();
}

/**
 * Combined menu: organize inbox then create pending email drafts.
 */
function organizeInboxAndCreateDrafts() {
  var result = runOrganizeAndDraftsPipeline_({ source: 'menu' });
  notifyUser_(result.message, 'Organize + Drafts');
  maybeSendDailySummaryAfterRun_(result);
  return result;
}

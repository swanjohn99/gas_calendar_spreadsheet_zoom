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
 * Does not send the day-summary email (only the last scheduled job does).
 */
function organizeInboxAndCreateDrafts() {
  var result = runOrganizeAndDraftsPipeline_({ source: 'menu' });
  notifyUser_(result.message, 'Organize + Drafts');
  return result;
}

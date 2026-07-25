function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Tools')
    .addItem('Import Calendar', 'importCalendar')
    .addItem('Schedule', 'scheduleCalendarSync')
    .addItem('Create Email Drafts', 'createEmailDraftsForSelection')
    .addItem('Organize Drive Inbox', 'organizeDriveInbox')
    .addToUi();
}

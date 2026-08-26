# Calendar Zoom recordings

Bound spreadsheet workflow that files Zoom meeting recordings against `events` rows.

## Language

**Event**:
A calendar meeting stored as one row on the `events` sheet.
_Avoid_: Training event, coaching event

**Zoom meeting ID**:
The numeric Zoom meeting number shared by every occurrence of a recurring meeting.
_Avoid_: Meeting ID, event_id

**Start**:
The meeting datetime on an `events` row (`yyyy-MM-dd HH:mm:ss` in America/Chicago). Date and time together identify which occurrence an inbox file belongs to.
_Avoid_: Date stamp, meeting date

**Zoom UUID**:
The Zoom meeting-instance identifier carried in the inbox filename after start. Written to `zoom_uuid` and to the matching artifact column (`video_uuid`, `pdf_uuid`, `audio_uuid`, `transcript_uuid`, `chat_uuid`). Not part of the destination filename.
_Avoid_: UUID, event_id, zoom_meeting_id

**Inbox file**:
A recording dropped in the Drive inbox, named `{zoom_meeting_id}-{start}_{zoom_uuid}.{ext}`.
_Avoid_: Upload, Drive file

# Review Fixes Plan

## Fix #1 (High): README.md - Fix Reuniones table structure
**File:** README.md, lines 42-53
**What:** Replace the 6-column Fecha_hora meeting table with 7-column Fecha + Hora to match actual Code.gs structure (columns D=Fecha, E=Hora, F=Agenda, G=Documentación).

## Fix #2 (High): Code.gs - Fix getPreviousMeetingsSnapshot date/time handling
**File:** Code.gs, lines 999-1009
**What:** row[5] is the stored combined datetime string but gets assigned to `datetime`, `date`, AND `time` all pointing to the same value. When `detectMeetingFieldChanges` later compares `date` and `time` separately, it's comparing a combined datetime string against proper separate date/time values from the current meetings. Fix: parse the stored datetime to extract date-only and time-only for those fields.

Replace:
```javascript
datetime: row[5],
date: row[5],
time: row[5],
```
With logic that parses row[5] as a datetime and extracts date/time components:
```javascript
datetime: row[5],
date: parseDatePartFromHistoryDatetime(row[5]),
time: parseTimePartFromHistoryDatetime(row[5]),
```
Where `parseDatePartFromHistoryDatetime` returns just the date portion and `parseTimePartFromHistoryDatetime` returns just the time portion. Since `normalizeDateForComparison` and `normalizeTimeForComparison` already handle Date objects properly, the simplest fix is to just pass `row[5]` (the datetime) and let those normalizers extract what they need. BUT the issue is that when `row[5]` is a string like "13-03-2026 14:00", `formatDate()` in `detectMeetingFieldChanges` will show the full datetime. So we should parse it and store it properly.

Actually, the simplest correct fix: keep `datetime: row[5]` and set `date: row[5]` and `time: row[5]` since the normalization functions already handle extracting date-only and time-only from full datetimes. The real comparison happens via `normalizeDateForComparison` and `normalizeTimeForComparison` which both handle Date objects and datetime strings. The display values (`formatDate`, `formatTime`) also handle this. So the comparison logic is actually correct because normalization handles it.

Wait - re-examining: `formatDate()` on a datetime string like "13-03-2026 14:00" would fail to parse it as a date via `new Date()` (since it's dd-mm-yyyy format, not ISO). But `parseLocalizedDateOrDateTime` handles "dd-mm-yyyy HH:mm". And `formatDate` does `new Date(dateValue)` for strings, which would fail for "13-03-2026 14:00". So when showing old/new values in emails, `formatDate()` on the history datetime string would just return the raw string. This is a display bug.

**Proper fix:** In `getPreviousMeetingsSnapshot`, parse row[5] and extract proper date/time:
```javascript
const parsedDatetime = parseLocalizedDateOrDateTime(row[5]);
...
datetime: row[5],
date: parsedDatetime || row[5],
time: parsedDatetime || row[5],
```
This way `normalizeDateForComparison(parsedDatetime)` and `normalizeTimeForComparison(parsedDatetime)` work correctly on Date objects.

## Fix #3 (Medium): Code.gs - Restore column count validation
**File:** Code.gs, lines 426-433
**What:** Uncomment the column count check and restore it. Remove the comment markers.

## Fix #4 (Medium): README.md - Fix English email example to Spanish
**File:** README.md, lines 124-162
**What:** Replace the English email example with Spanish text matching the actual code output.

## Fix #5 (Medium): DEPLOYMENT_INSTRUCTIONS.md - Multiple updates
**File:** DEPLOYMENT_INSTRUCTIONS.md
- Update version from 1.0 to 2.0, February to March 2026 (line 344-346)
- Fix config example: change English EMAIL_SUBJECT to Spanish (line 201)
- Add brief mention of meetings in the installation workflow section

## Fix #6 (Medium): README.md - Remove stale future enhancement
**File:** README.md, line 279
**What:** Remove "HTML formatted emails" from Future Enhancements list (already implemented).

## Fix #7 (Low): Code.gs - Remove unused startTime
**File:** Code.gs, lines 189 and 255
**What:** Delete `const startTime = new Date();` in both functions.

## Fix #8 (Low): Code.gs - Fix typo
**File:** Code.gs, line 7
**What:** Change `spreadsheet.MP` to `spreadsheet.`

## Fix #9 (Low): Code.gs - Optimize save functions
**File:** Code.gs, lines 897-904 (saveSnapshot) and lines 1049-1059 (saveMeetingsSnapshot)
**What:** Replace `.find()` calls with Set lookups for action determination. The `notifiedKeys` Set is already built but not used for action. Build separate Sets for new/changed/etc keys and use `.has()`.

For saveSnapshot:
```javascript
const newKeys = new Set(changes.new.map(t => `${t.taskId}|${t.email}`));
const changedKeys = new Set(changes.changed.map(t => `${t.task.taskId}|${t.task.email}`));
// Then in the loop:
if (newKeys.has(currentKey)) action = 'NEW';
else if (changedKeys.has(currentKey)) action = 'CHANGED';
```

For saveMeetingsSnapshot:
```javascript
const postponedKeys = new Set(changes.postponed.map(m => `${m.meeting.meetingId}|${m.meeting.email}`));
const cancelledKeys = new Set(changes.cancelled.map(m => `${m.meeting.meetingId}|${m.meeting.email}`));
const newKeys = new Set(changes.new.map(m => `${m.meetingId}|${m.email}`));
const changedKeys = new Set(changes.changed.map(m => `${m.meeting.meetingId}|${m.meeting.email}`));
// Then in the loop use .has()
```

## Fix #10 (Low): Code.gs - Add email validation before sending
**File:** Code.gs, lines ~1585 and ~2001
**What:** Add `if (!isValidEmail(email)) { Logger.log(...); return; }` at the start of the forEach callback before `MailApp.sendEmail`.

## Fix #11 (Low): DEPLOYMENT_INSTRUCTIONS.md - Fix testing checklist
**File:** DEPLOYMENT_INSTRUCTIONS.md, line 275
**What:** Change English subject `"Task Updates - changes require your attention"` to Spanish `"Actualización de tareas - se requiere su atención"`.

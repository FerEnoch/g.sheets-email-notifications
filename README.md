# Google Sheets Task Notification System

A smart, granular notification system for Google Sheets project management boards that only notifies team members when their tasks actually change.

## 🎯 Overview

This Google Apps Script automatically detects changes in your project management spreadsheet and sends consolidated email notifications only to assignees whose tasks have been updated. It tracks complete history and shows exactly what changed (old → new values).

## ✨ Features

- **🔍 Smart Change Detection** - Compares current state with previous snapshot to detect new, changed, and deleted tasks
- **📧 Granular Notifications** - Only emails people whose tasks actually changed (no spam!)
- **📝 Detailed Context** - Shows exactly what changed with old → new comparisons
- **📊 Multi-sheet Support** - Automatically processes all sheets in your spreadsheet
- **🗂️ History Tracking** - Maintains complete audit trail in dedicated History sheet
- **👥 Consolidated Emails** - One email per person with all their changes grouped together
- **🎨 HTML Task Emails** - Task notifications are sent as rich HTML with plain-text fallback
- **🔔 Deletion Notifications** - Alerts previous assignee when tasks are removed
- **📅 Consistent Formatting** - All dates in dd-mm-yyyy format
- **✉️ Multi-Assignee Tasks** - Assign tasks to multiple people using comma or semicolon separators (e.g., `user1@x.com, user2@x.com`)
- **📅 Meeting Notifications** - Separate "Reuniones" sheet with dedicated notification system for meetings

## 📋 Requirements

### Task Spreadsheet Structure

Your spreadsheet should have the following columns (Row 1 = Headers, Row 2+ = Data):

| Column | Field | Description | Monitored for Changes |
|--------|-------|-------------|----------------------|
| A | Task | Task name/description | ✅ Yes |
| B | Priority | Task priority level | ✅ Yes |
| C | Assignee | Email address(es) - supports multiple (comma/semicolon separated) | ✅ Yes |
| D | Status | Current task status | ✅ Yes |
| E | Init Date | Start date | ✅ Yes |
| F | Finish Date | Due/completion date | ✅ Yes |
| G | Hits | Number of updates/views | ❌ No (too frequent) |
| H | Product | Related product/project | ✅ Yes |
| I | Notes | Additional notes | ✅ Yes |

### Meeting Spreadsheet Structure (Optional)

To enable meeting notifications, create a sheet named "Reuniones" with these columns:

| Column | Field | Description | Monitored for Changes |
|--------|-------|-------------|----------------------|
| A | Título | Meeting title | ✅ Yes |
| B | Asistentes | Attendee emails - supports multiple (comma/semicolon separated) | ✅ Yes |
| C | Estado | Meeting status (use dropdown: Programada, Completada, Cancelada, Pospuesta) | ✅ Yes |
| D | Fecha_hora | Meeting date and time (use Google Sheets Date time format) | ✅ Yes |
| E | Agenda | Meeting agenda/topics | ✅ Yes |
| F | Documentación | Google Drive URLs or notes | ✅ Yes |

## 🚀 Quick Start

1. **Open your Google Spreadsheet**
2. Go to **Extensions → Apps Script**
3. Delete any default code
4. Copy and paste the contents of `Code.gs`
5. Save the project (name it "Task Notification System")
6. Close Apps Script and refresh your spreadsheet
7. You'll see a new menu: **📧 Notifications**
8. Choose:
   - **📧 Notifications → Notify task assignees** - For task notifications
   - **📧 Notifications → Notify meeting attendees** - For meeting notifications (requires "Reuniones" sheet)
9. Authorize the script when prompted (first time only)

For detailed instructions, see [DEPLOYMENT_INSTRUCTIONS.md](./DEPLOYMENT_INSTRUCTIONS.md)

## 📧 How It Works

### Task Notifications

#### First Run (Transparency)
- Creates "Task History" sheet automatically
- Treats all tasks as "new"
- Sends notifications to ALL assignees with their current assignments
- Saves initial snapshot for future comparisons

#### Subsequent Runs (Granular)
1. Scans all sheets for tasks
2. Compares with previous snapshot from History sheet
3. Detects changes:
   - ✨ **New tasks** - Newly assigned
   - 📝 **Changed tasks** - Field values modified
   - 🗑️ **Deleted tasks** - Removed from spreadsheet
   - ⏭️ **Unchanged tasks** - Skipped (no notification)
4. Groups changes by assignee email
5. Sends ONE consolidated email per person (HTML + plain-text fallback)
6. Saves new snapshot to History sheet
7. Shows summary alert with statistics

### Meeting Notifications

Works identically to task notifications but:
- Uses the "Reuniones" sheet instead of task sheets
- Creates a separate "_meetings_history" sheet
- Sends meeting-specific email format
- Includes Google Calendar action links in emails:
  - Add to Google Calendar (prefilled event)
  - Update meeting in Google Calendar (guided manual action)
  - Delete meeting in Google Calendar (guided manual action)
  - Send message to all attendees (opens mail client draft)
- Uses a default 30-minute duration for the prefilled add-event link
- Adds agenda, documentation, and spreadsheet link in event details
- Triggered via "Notify meeting attendees" menu option

Note: Update/Delete links open Google Calendar with search context so the attendee can edit or remove the correct event manually.
Note: "Send message to all attendees" includes all meeting attendees plus the notification sender as recipients.

### Multi-Assignee Support

Both tasks and meetings support multiple email addresses:
- Use comma (`,`) or semicolon (`;`) as separators
- Example: `user1@example.com, user2@example.com; user3@example.com`
- Each person receives individual notifications
- Email body shows all assignees/attendees for collaboration context

## 📨 Email Format Example

```
Subject: Task Updates - changes require your attention

Hi there,

You have 2 tasks that were recently updated in the project management board:

═══════════════════════════════════════════════════════

✨ NEW TASK ASSIGNED TO YOU

Task: Implement OAuth integration
  Priority: High
  Status: Not Started
  Product: Customer Portal
  Start Date: 26-02-2026
  Due Date: 05-03-2026
  Notes: Use Google OAuth 2.0
  
  → This task was newly assigned to you

───────────────────────────────────────────────────────

📝 TASK UPDATED

Task: Fix mobile responsiveness
  Priority: Medium → High ⚠️
  Status: Not Started → In Progress ✓
  Due Date: 28-02-2026 → 05-03-2026
  Product: Marketing Site
  Notes: Focus on tablet view
  
  → 3 field(s) changed since last notification

═══════════════════════════════════════════════════════

View the full board: [Link to spreadsheet]
Notification sent: 26-02-2026 14:30
```

## 🔧 Configuration

All settings can be customized in the `CONFIG` object at the top of `Code.gs`:

```javascript
const CONFIG = {
  HISTORY_SHEET_NAME: 'Task History',
  EMAIL_SUBJECT: 'Task Updates - changes require your attention',
  HEADER_ROW: 1,
  FIRST_DATA_ROW: 2,
  MONITORED_FIELDS: [...],
};
```

### Customizing Column Mapping

If your columns are in a different order, update the `COLUMNS` mapping:

```javascript
COLUMNS: {
  TASK: 0,        // Column A (0-indexed)
  PRIORITY: 1,    // Column B
  EMAIL: 2,       // Column C
  // etc.
}
```

## 📊 Understanding the History Sheet

The script automatically creates a "Task History" sheet with these columns:

- **Timestamp** - When snapshot was taken
- **Sheet** - Source sheet name
- **Task ID** - Unique identifier (SheetName_RowNumber)
- **Task, Priority, Assignee, Status, etc.** - Task field values
- **Action** - NEW, CHANGED, DELETED, or UNCHANGED
- **Notified** - YES if email sent, NO if skipped

**⚠️ Important:** Don't delete this sheet! It's required for change detection.

## 🎯 Use Cases

Perfect for:
- **Project management boards** - Track task assignments and changes
- **Sprint planning** - Notify team of sprint updates
- **Issue tracking** - Alert assignees about issue changes
- **Task delegation** - Automatically notify when tasks are assigned
- **Status updates** - Keep team informed of progress changes

## 🔒 Permissions

The script requires these Google permissions:

- **View and manage spreadsheets** - Read task data and create History sheet
- **Send email as you** - Send notifications to assignees
- **Display content in Google apps** - Show menu and alerts

**Privacy:** The script only accesses your spreadsheet data and sends emails to addresses found within it. No external services are used.

## 🐛 Troubleshooting

### Menu doesn't appear
- Refresh the spreadsheet (F5)
- Check that the script was saved successfully

### No emails being sent
- Verify Column C contains valid email addresses
- Check if tasks actually changed since last run
- Review the summary alert for statistics

### Emails going to spam
- Ask recipients to mark as "Not Spam"
- Add sender to contacts/whitelist

### Script authorization issues
- Go to Extensions → Apps Script → Run → notifyTaskAssignees
- Complete the authorization flow
- See deployment instructions for detailed steps

For more troubleshooting, see [DEPLOYMENT_INSTRUCTIONS.md](./DEPLOYMENT_INSTRUCTIONS.md)

## 📁 Project Structure

```
google-sheets-notifications/
├── Code.gs                      # Complete Apps Script implementation
├── DEPLOYMENT_INSTRUCTIONS.md   # Detailed deployment guide
└── README.md                    # This file
```

## 🔮 Future Enhancements

Possible additions (not currently implemented):

- Automatic triggers on cell edits
- Daily/weekly digest emails
- HTML formatted emails
- User preference settings per assignee
- Filter by specific sheet names
- Slack/Teams integration
- Custom email templates by status/priority

## 📝 Version

- **Version:** 2.0
- **Last Updated:** March 2026
- **Compatibility:** Google Apps Script (Google Sheets)

### Changelog

**v2.0** (March 2026)
- Added multi-assignee support for tasks (comma/semicolon separated emails)
- Added meeting notification system with separate "Reuniones" sheet
- Added independent history tracking for meetings (_meetings_history)
- Added Google Calendar action links in meeting emails (add/update/delete helper actions)

**v1.0** (February 2026)
- Initial release with task notification system

## 📄 License

This is a custom implementation for project management purposes. Feel free to modify and adapt for your needs.

## 🤝 Contributing

This is a standalone script. To modify:
1. Edit `Code.gs` in your Apps Script editor
2. Save and test changes
3. Document any modifications

## 📞 Support

For issues or questions:
1. Check the [DEPLOYMENT_INSTRUCTIONS.md](./DEPLOYMENT_INSTRUCTIONS.md) troubleshooting section
2. Review Apps Script execution logs (Extensions → Apps Script → Executions)
3. Check the console logs (View → Logs in Apps Script editor)

## ✨ Credits

Built with Google Apps Script for granular project management notifications.

---

**Ready to get started?** Follow the [deployment instructions](./DEPLOYMENT_INSTRUCTIONS.md) to set up the system in your spreadsheet! 🚀

# Task Notification System - Deployment Instructions

## 📋 Overview

This document provides step-by-step instructions for deploying the Task Notification System to your Google Spreadsheet project management board.

---

## ✅ Prerequisites

Before you begin, ensure you have:

1. A Google Spreadsheet with your project management data
2. Tasks organized with the following column structure:
   - **Column A:** Task
   - **Column B:** Priority  
   - **Column C:** Assignee (email address)
   - **Column D:** Status
   - **Column E:** Init Date (start date)
   - **Column F:** Finish Date (due date)
   - **Column G:** Hits
   - **Column H:** Product
   - **Column I:** Notes
3. Row 1 contains headers
4. Task data starts from Row 2 onwards

---

## 🚀 Installation Steps

### Step 1: Open Apps Script Editor

1. Open your Google Spreadsheet
2. Click on **Extensions** in the menu bar
3. Select **Apps Script**
4. A new tab will open with the Apps Script editor

### Step 2: Prepare the Script Editor

1. You'll see a default `Code.gs` file with some placeholder code
2. **Select all the existing code** (Ctrl+A or Cmd+A)
3. **Delete it** (you'll replace it with the new script)

### Step 3: Add the Script

1. Open the `Code.gs` file from this package
2. **Copy the entire contents** of the script
3. **Paste it** into the Apps Script editor
4. The editor should now show the complete Task Notification System code

### Step 4: Save the Project

1. Click the **💾 Save** icon (or press Ctrl+S / Cmd+S)
2. When prompted, give your project a name:
   - Suggested name: **"Task Notification System"**
3. Click **OK**

### Step 5: Authorize the Script

1. **Close the Apps Script tab** and return to your spreadsheet
2. **Refresh the spreadsheet** (press F5 or reload the page)
3. After a few seconds, you'll see a new menu appear: **📧 Notifications**
4. Click on **📧 Notifications → Notify task assignees**
5. A permission dialog will appear asking for authorization:
   - Click **Continue**
   - Select your Google account
   - Click **Advanced** (if you see a warning screen)
   - Click **Go to Task Notification System (unsafe)** 
   - Review the permissions and click **Allow**

**Note:** The "unsafe" warning is normal for unverified scripts. This script only runs in your spreadsheet and doesn't share data externally.

### Step 6: First Run

1. After authorization, click **📧 Notifications → Notify task assignees** again
2. The script will run for the first time:
   - Creates Drive folder `history_and_backup` inside the same folder as your board spreadsheet
   - Creates spreadsheet `history_backup`
   - Creates page `Tasks_Current`
   - Scans all sheets for tasks
   - Sends notifications to ALL assignees (first-run transparency)
   - Shows a summary alert
3. Check the results:
   - You should see the folder and spreadsheet created in Drive
   - You should see `Tasks_Current` in `history_backup`
   - Assignees should receive emails
   - An alert shows statistics about the run

---

## 📧 What Happens on First Run

The **first time** you run the script:
- ✅ ALL tasks are treated as "new"
- ✅ ALL assignees with valid email addresses receive notifications
- ✅ External storage is created (`history_and_backup/history_backup`)
- ✅ Current history page `Tasks_Current` is created to track future changes
- ✅ A snapshot of all current tasks is saved

This is intentional for **transparency** - everyone gets notified about their current assignments.

---

## 🔄 Subsequent Runs

After the first run, the script becomes **truly granular**:
- ✅ Only notifies assignees whose tasks have **changed**
- ✅ Shows exactly **what changed** (old → new values)
- ✅ Detects new tasks, changed tasks, and deleted tasks
- ✅ Skips unchanged tasks (no notifications sent)

---

## 📊 Understanding External History Storage

After running the script, history is stored in:
- Folder: `history_and_backup` (inside the same folder as the board spreadsheet)
- Spreadsheet: `history_backup`

Current pages (< 45 days):
- `Tasks_Current`
- `Meetings_Current`

Backup pages (> 45 days), split by window and type:
- `Tasks_Backup_YYYY-MM-DD_to_YYYY-MM-DD`
- `Meetings_Backup_YYYY-MM-DD_to_YYYY-MM-DD`

Each history page uses these columns:

| Column | Purpose |
|--------|---------|
| Timestamp | When the snapshot was taken |
| Sheet | Which sheet the task is on |
| Task ID | Unique identifier (SheetName_RowNumber) |
| Task | Task name |
| Priority | Task priority |
| Assignee | Email address |
| Status | Task status |
| Init Date | Start date |
| Finish Date | Due date |
| Hits | Hit count |
| Product | Related product |
| Notes | Task notes |
| Action | NEW, CHANGED, DELETED, or UNCHANGED |
| Notified | YES if email was sent, NO otherwise |

**Important:** Do not delete `history_backup`! It's needed for change detection and backup continuity.

---

## 🎯 Using the System

### Daily Workflow

1. Team members update tasks in the spreadsheet (change status, priority, dates, etc.)
2. When ready to notify everyone about changes, click: **📧 Notifications → Notify task assignees**
3. The script:
   - Compares current state with the last snapshot
   - Identifies what changed
   - Sends emails only to people with changes
   - Shows you a summary
4. Check the summary alert to see how many people were notified

### What Triggers Notifications

The following field changes will trigger notifications:
- ✅ Task name
- ✅ Priority
- ✅ Assignee (email)
- ✅ Status
- ✅ Start Date (Init Date)
- ✅ Due Date (Finish Date)
- ✅ Product
- ✅ Notes
- ❌ Hits (NOT monitored - changes too frequently)

### Email Content

Recipients receive one consolidated email with:
- **New tasks:** Full details of newly assigned tasks
- **Changed tasks:** What changed (old → new) plus current details
- **Deleted tasks:** Notification that task was removed
- **Link** to the spreadsheet
- **Timestamp** in dd-mm-yyyy format

---

## 🔧 Configuration (Optional)

If you need to customize the script, open `Code.gs` and modify the `CONFIG` object at the top:

```javascript
const CONFIG = {
   HISTORY: {
      FOLDER_NAME: 'history_and_backup',
      SPREADSHEET_NAME: 'history_backup',
      RETENTION_DAYS: 45,
      TASKS_CURRENT_SHEET_NAME: 'Tasks_Current',
      MEETINGS_CURRENT_SHEET_NAME: 'Meetings_Current'
   },
  EMAIL_SUBJECT: 'Actualización de tareas - se requiere su atención',  // Customize subject
  HEADER_ROW: 1,  // Change if headers are not in row 1
  FIRST_DATA_ROW: 2,  // Change if data doesn't start in row 2
  // ... other settings
};
```

**If you change column order**, update the `COLUMNS` object:
```javascript
COLUMNS: {
  TASK: 0,        // Column A
  PRIORITY: 1,    // Column B
  EMAIL: 2,       // Column C
  // etc.
}
```

---

## ❗ Troubleshooting

### Menu Doesn't Appear
- **Solution:** Refresh the spreadsheet (F5)
- If still not showing, reopen Apps Script and save the project again

### "Authorization Required" Error
- **Solution:** Follow Step 5 above to grant permissions
- You only need to do this once

### No Emails Being Sent
- **Check:** Are there valid email addresses in Column C?
- **Check:** Did any tasks actually change since last run?
- **Check:** Look at the summary alert - it shows how many emails were sent

### "Invalid Email" Warnings
- **Check:** Column C (Assignee) must contain valid email addresses
- Invalid entries are skipped and logged in history pages

### Script Runs Slowly
- Normal for large spreadsheets (100+ tasks)
- Expected run time: 2-10 seconds depending on data size

### Emails Going to Spam
- **Solution:** Ask recipients to mark as "Not Spam" the first time
- Add a rule to whitelist emails from your Google account

---

## 🛡️ Permissions Explained

The script requests these permissions:

| Permission | Why It's Needed |
|------------|-----------------|
| View and manage spreadsheets | To read task data and maintain history pages |
| View and manage Drive files/folders | To create `history_and_backup` and `history_backup` |
| Send email as you | To send notifications to assignees |
| Display content in Google apps | To show menu and alerts |

**Privacy Note:** The script only sends data to email addresses found in your spreadsheet. No data is sent to external services.

---

## 🧪 Testing Checklist

Use this checklist to verify everything works:

- [ ] Menu "📧 Notifications" appears after refresh
- [ ] Button "Notify task assignees" is clickable
- [ ] First run creates folder `history_and_backup`
- [ ] First run creates spreadsheet `history_backup`
- [ ] First run creates page `Tasks_Current`
- [ ] First run sends emails to all assignees
- [ ] Summary alert shows correct statistics
- [ ] Email subject is: "Actualización de tareas - se requiere su atención"
- [ ] Task notification email renders with HTML card layout (with plain-text fallback)
- [ ] Email dates are in dd-mm-yyyy format
- [ ] Change a task status and run again
- [ ] Only the person whose task changed gets an email
- [ ] Email shows old → new values correctly
- [ ] Delete a task and run again
- [ ] Previous assignee gets deletion notification
- [ ] Add a new task and run again
- [ ] New assignee gets notification

### Meeting Notification Checklist

If you use the "Reuniones" sheet, also validate meeting emails:

- [ ] Menu "📧 Notifications → Notificar reuniones" is available
- [ ] First run creates `Meetings_Current` in `history_backup`
- [ ] NEW meeting email includes action links for Add/Update/Delete in Google Calendar and Send message to all attendees
- [ ] "Add to Google Calendar" opens a prefilled event with title/date/time/details
- [ ] Default duration in the prefilled event is 30 minutes
- [ ] "Update meeting in Google Calendar" opens Calendar search/context for manual edit
- [ ] "Delete meeting in Google Calendar" opens Calendar search/context for manual delete
- [ ] "Send message to all attendees" opens an email draft with attendees + notification sender in recipients
- [ ] CHANGED meeting emails include update/delete helper links
- [ ] CHANGED meeting emails include the message-all action
- [ ] DELETED meeting emails include delete helper link
- [ ] DELETED meeting emails include the message-all action
- [ ] Agenda/documentation/spreadsheet URL appear in add-event details when available

Note: Update/Delete links are guided manual actions and do not directly modify the attendee calendar event.
Note: Message-all uses a mailto draft and depends on the recipient mail client behavior.

---

## 📞 Support

If you encounter issues:

1. Check the **execution log** in Apps Script:
   - Extensions → Apps Script
   - Click **Executions** on the left sidebar
   - Review any error messages

2. Check the **Apps Script logs**:
   - In Apps Script editor, click **View → Logs**
   - Look for error messages or warnings

3. Common fixes:
   - Ensure Column C contains valid emails
   - Verify your spreadsheet structure matches the expected format
   - Check that you have permission to send emails from your Google account

---

## 🔄 Future Enhancements

This script is designed to be extensible. Potential future additions:

- Automatic triggers (run on cell edit)
- Daily digest emails
- User preference settings
- Filter by specific sheets
- Slack/Teams integration

---

## 📝 Version Information

- **Version:** 2.0
- **Last Updated:** March 2026
- **Compatibility:** Google Apps Script (Google Sheets)
- **Author:** Custom implementation for project management

---

## ✅ You're All Set!

Your Task Notification System is now installed and ready to use. Start by:

1. Ensuring all your tasks have valid email addresses in Column C
2. Clicking **📧 Notificaciones → Notificar tareas a los asignados** for the first run
3. Making some changes to tasks
4. Running the notification again to see granular change detection in action
5. (Optional) Set up the "Reuniones" sheet and use **📧 Notificaciones → Notificar reuniones** for meeting notifications

Happy project managing! 🚀

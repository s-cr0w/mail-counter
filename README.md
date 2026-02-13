# Outlook Email Counter Script

This VBA script counts emails in a specific folder within a date range and displays detailed statistics including counts by color category.

## Features

- **Date Range Selection**: Input boxes allow you to specify start and end dates
- **Total Email Count**: Shows total emails within the specified date range
- **Category Breakdown**: Displays count for each color category assigned to emails
- **Sorted Results**: Categories are displayed alphabetically

## Configuration

The script is configured to count emails in:
- **Mailbox**: random@example.com
- **Folder**: Inbox > test

## How It Works

When you run the script:

1. **Start Date Input**: Enter the beginning date of your range (defaults to 30 days ago)
2. **End Date Input**: Enter the ending date of your range (defaults to today)
3. **Processing**: The script scans all emails in the folder within the date range
4. **Results Display**: A detailed report shows:
   - Total emails in the date range
   - Count of emails for each color category (Red, Blue, Green, etc.)
   - Count of emails with no category assigned

## How to Install and Use

### Method 1: Import the .bas file

1. Open Outlook
2. Press `Alt + F11` to open the Visual Basic Editor
3. Go to **File** > **Import File**
4. Select the `CountEmailsInFolder.bas` file
5. Close the Visual Basic Editor
6. Press `Alt + F8` to open the Macros dialog
7. Select `CountEmailsInTestFolder` and click **Run**

### Method 2: Copy and paste the code

1. Open Outlook
2. Press `Alt + F11` to open the Visual Basic Editor
3. In the left pane (Project Explorer), find **Project1 (VbaProject.OTM)**
4. Right-click on **Modules** > **Insert** > **Module**
5. Copy the code from `CountEmailsInFolder.bas` and paste it into the new module
6. Close the Visual Basic Editor
7. Press `Alt + F8` to open the Macros dialog
8. Select `CountEmailsInTestFolder` and click **Run**

## Customization

To change the mailbox or folder, edit these lines in the script:

```vb
' Change the mailbox email address
mailboxName = "random@example.com"

' Change the folder name (currently set to "test")
Set objTestFolder = objInbox.Folders("test")
```

## Macro Security

If the macro doesn't run, you may need to adjust Outlook's macro security settings:

1. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings**
2. Select **Macro Settings**
3. Choose **Notifications for all macros** or **Enable all macros** (less secure)
4. Click **OK** and restart Outlook

## Example Output

```
Email Count Report
==================================================

Mailbox: random@example.com
Folder: Inbox\test
Date Range: 2026-01-01 to 2026-02-13
--------------------------------------------------

Total emails in date range: 45

Breakdown by Category:
--------------------------------------------------
  (No Category): 12
  Blue Category: 8
  Green Category: 15
  Red Category: 7
  Yellow Category: 3
```

## Notes

- **Date Format**: Use YYYYMMDD format for input (e.g., 20260213 for February 13, 2026)
- **Categories**: The script counts Outlook color categories (the ones you assign via right-click > Categorize)
- **Multiple Categories**: If an email has multiple categories, it will be counted in each category
- **Performance**: For folders with thousands of emails, the script may take a few seconds to process

## Troubleshooting

- Ensure the mailbox "random@example.com" is added to your Outlook profile
- Verify the "test" folder exists in the Inbox of that mailbox
- Check that the folder name matches exactly (case-sensitive)
- If you get a permission error, ensure you have access to the mailbox
- Make sure emails have categories assigned (right-click email > Categorize > pick a color)

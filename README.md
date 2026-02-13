# Outlook Email Counter Script

This VBA script counts emails in a specific folder within a date range and displays detailed statistics including counts by color category.

## Features

- **Dynamic Mailbox Selection**: Choose any mailbox in your Outlook profile
- **Flexible Folder Path**: Navigate to any subfolder using path notation (e.g., Inbox/subfolder1/subfolder2)
- **Date Range Selection**: Input boxes allow you to specify start and end dates
- **Total Email Count**: Shows total emails within the specified date range
- **Category Breakdown**: Displays count for each color category assigned to emails
- **Sorted Results**: Categories are displayed alphabetically

## How It Works

When you run the script, you'll be prompted for the following inputs:

1. **Mailbox Name**: Enter the email address of the mailbox (e.g., random@example.com)
2. **Folder Path**: Enter the folder path using forward slashes (e.g., Inbox/subfolder1/subfolder2)
3. **Start Date**: Enter the beginning date in YYYYMMDD format (defaults to 30 days ago)
4. **End Date**: Enter the ending date in YYYYMMDD format (defaults to today)
5. **Processing**: The script scans all emails in the specified folder within the date range
6. **Results Display**: A detailed report shows:
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
Folder: Inbox\subfolder1\subfolder2
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

## Example Input Sequences

**Example 1: Simple folder**
- Mailbox: `user@company.com`
- Folder Path: `Inbox/Projects`
- Start Date: `20260101`
- End Date: `20260213`

**Example 2: Nested folder**
- Mailbox: `support@company.com`
- Folder Path: `Inbox/Clients/CompanyXYZ/Issues`
- Start Date: `20260201`
- End Date: `20260213`

## Notes

- **Date Format**: Use YYYYMMDD format for input (e.g., 20260213 for February 13, 2026)
- **Folder Path Format**: Use forward slashes (/) to separate folder levels (e.g., Inbox/subfolder1/subfolder2)
- **Mailbox**: The mailbox must be configured in your Outlook profile
- **Categories**: The script counts Outlook color categories (the ones you assign via right-click > Categorize)
- **Multiple Categories**: If an email has multiple categories, it will be counted in each category
- **Performance**: The script uses Outlook's `Restrict` method to filter emails by date range BEFORE loading them, making it efficient even for folders with thousands of emails. Only emails within your specified date range are retrieved and processed.

## Troubleshooting

- Ensure the mailbox you enter is added to your Outlook profile (check the folder pane in Outlook)
- Verify the folder path exists - use the exact folder names as they appear in Outlook
- Folder names are case-sensitive - match the capitalization exactly
- Use forward slashes (/) in the folder path, not backslashes (\)
- If you get a permission error, ensure you have access to the mailbox
- For shared mailboxes, make sure they're fully loaded in Outlook before running the script
- Make sure emails have categories assigned (right-click email > Categorize > pick a color)

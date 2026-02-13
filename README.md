# Outlook Email Counter Script

This VBA script counts all emails in a specific folder and displays the count in a pop-up dialog box.

## Configuration

The script is configured to count emails in:
- **Mailbox**: random@example.com
- **Folder**: Inbox > test

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

## Troubleshooting

- Ensure the mailbox "random@example.com" is added to your Outlook profile
- Verify the "test" folder exists in the Inbox of that mailbox
- Check that the folder name matches exactly (case-sensitive)
- If you get a permission error, ensure you have access to the mailbox

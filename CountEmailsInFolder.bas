Attribute VB_Name = "CountEmailsInFolder"
Option Explicit

Sub CountEmailsInTestFolder()
    Dim objNamespace As Outlook.Namespace
    Dim objMailbox As Outlook.Folder
    Dim objInbox As Outlook.Folder
    Dim objTestFolder As Outlook.Folder
    Dim emailCount As Long
    Dim mailboxName As String

    ' Set the mailbox name
    mailboxName = "random@example.com"

    ' Get the MAPI namespace
    Set objNamespace = Application.GetNamespace("MAPI")

    ' Error handling
    On Error GoTo ErrorHandler

    ' Access the specified mailbox
    Set objMailbox = objNamespace.Folders(mailboxName)

    ' Get the Inbox folder
    Set objInbox = objMailbox.Folders("Inbox")

    ' Get the "test" subfolder
    Set objTestFolder = objInbox.Folders("test")

    ' Count the emails in the test folder
    emailCount = objTestFolder.Items.Count

    ' Display the count in a message box
    MsgBox "The 'test' folder in mailbox '" & mailboxName & "' contains " & emailCount & " email(s).", _
           vbInformation, "Email Count"

    ' Clean up
    Set objTestFolder = Nothing
    Set objInbox = Nothing
    Set objMailbox = Nothing
    Set objNamespace = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please verify that:" & vbCrLf & _
           "1. The mailbox '" & mailboxName & "' exists in your Outlook profile" & vbCrLf & _
           "2. The 'test' folder exists in the Inbox of this mailbox", _
           vbCritical, "Error Accessing Folder"

    ' Clean up
    Set objTestFolder = Nothing
    Set objInbox = Nothing
    Set objMailbox = Nothing
    Set objNamespace = Nothing
End Sub

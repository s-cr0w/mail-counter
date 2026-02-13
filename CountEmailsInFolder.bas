Attribute VB_Name = "CountEmailsInFolder"
Option Explicit

Sub CountEmailsInTestFolder()
    Dim objNamespace As Outlook.Namespace
    Dim objMailbox As Outlook.Folder
    Dim objInbox As Outlook.Folder
    Dim objTestFolder As Outlook.Folder
    Dim objItem As Object
    Dim objMail As Outlook.MailItem
    Dim emailCount As Long
    Dim totalCount As Long
    Dim mailboxName As String
    Dim startDateStr As String
    Dim endDateStr As String
    Dim startDate As Date
    Dim endDate As Date
    Dim categoryDict As Object
    Dim categoryName As Variant
    Dim categories As Variant
    Dim cat As Variant
    Dim resultMessage As String
    Dim i As Integer
    Dim folderPath As String
    Dim folderParts() As String
    Dim currentFolder As Outlook.Folder
    Dim folderIndex As Integer
    Dim restrictedItems As Outlook.Items
    Dim filterString As String

    ' Get the MAPI namespace
    Set objNamespace = Application.GetNamespace("MAPI")

    ' Error handling
    On Error GoTo ErrorHandler

    ' Get mailbox name from user
    mailboxName = InputBox("Enter the mailbox name (e.g., email@example.com):", "Mailbox Name", "random@example.com")
    If mailboxName = "" Then
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Get folder path from user
    folderPath = InputBox("Enter the folder path (e.g., Inbox/subfolder1/subfolder2):", "Folder Path", "Inbox/test")
    If folderPath = "" Then
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Get start date from user
    startDateStr = InputBox("Enter the START date (YYYYMMDD):", "Start Date", Format(Date - 30, "yyyymmdd"))
    If startDateStr = "" Then
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Validate and parse start date
    If Not IsValidYYYYMMDD(startDateStr) Then
        MsgBox "Invalid start date format. Please use YYYYMMDD (e.g., 20260213).", vbCritical, "Invalid Date"
        Exit Sub
    End If
    startDate = ParseYYYYMMDD(startDateStr)

    ' Get end date from user
    endDateStr = InputBox("Enter the END date (YYYYMMDD):", "End Date", Format(Date, "yyyymmdd"))
    If endDateStr = "" Then
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Validate and parse end date
    If Not IsValidYYYYMMDD(endDateStr) Then
        MsgBox "Invalid end date format. Please use YYYYMMDD (e.g., 20260213).", vbCritical, "Invalid Date"
        Exit Sub
    End If
    endDate = ParseYYYYMMDD(endDateStr)

    ' Validate date range
    If startDate > endDate Then
        MsgBox "Start date cannot be after end date.", vbCritical, "Invalid Date Range"
        Exit Sub
    End If

    ' Add one day to endDate for the filter (we'll use < instead of <=)
    ' This ensures we get all emails on the end date through 23:59:59
    Dim endDatePlusOne As Date
    endDatePlusOne = endDate + 1

    ' Access the specified mailbox
    Set objMailbox = objNamespace.Folders(mailboxName)

    ' Parse the folder path and navigate to the target folder
    folderParts = Split(folderPath, "/")
    Set currentFolder = objMailbox

    ' Navigate through the folder hierarchy
    For folderIndex = LBound(folderParts) To UBound(folderParts)
        Set currentFolder = currentFolder.Folders(Trim(folderParts(folderIndex)))
    Next folderIndex

    ' Set the target folder
    Set objTestFolder = currentFolder

    ' Create a dictionary to store category counts
    Set categoryDict = CreateObject("Scripting.Dictionary")
    categoryDict.CompareMode = 1 ' Text compare, case-insensitive

    ' Initialize counters
    emailCount = 0
    totalCount = objTestFolder.Items.Count

    ' Build filter string for date range using Restrict method
    ' This significantly improves performance for large folders
    ' Using >= startDate and < endDate+1 ensures we capture all emails on both start and end dates
    filterString = "[ReceivedTime] >= '" & Format(startDate, "ddddd h:nn AMPM") & "' AND [ReceivedTime] < '" & Format(endDatePlusOne, "ddddd h:nn AMPM") & "'"

    ' Apply the filter to get only emails within the date range
    Set restrictedItems = objTestFolder.Items.Restrict(filterString)

    ' Sort by ReceivedTime for better performance
    restrictedItems.Sort "[ReceivedTime]", False

    ' Loop through only the filtered items
    For Each objItem In restrictedItems
        ' Check if it's a mail item
        If TypeOf objItem Is Outlook.MailItem Then
            Set objMail = objItem
            emailCount = emailCount + 1

            ' Process categories
            If objMail.Categories <> "" Then
                ' Split categories by semicolon (multiple categories possible)
                categories = Split(objMail.Categories, ";")
                For Each cat In categories
                    cat = Trim(cat)
                    If cat <> "" Then
                        If categoryDict.Exists(cat) Then
                            categoryDict(cat) = categoryDict(cat) + 1
                        Else
                            categoryDict.Add cat, 1
                        End If
                    End If
                Next cat
            Else
                ' No category
                If categoryDict.Exists("(No Category)") Then
                    categoryDict("(No Category)") = categoryDict("(No Category)") + 1
                Else
                    categoryDict.Add "(No Category)", 1
                End If
            End If
        End If
    Next objItem

    ' Build result message
    resultMessage = "Email Count Report" & vbCrLf
    resultMessage = resultMessage & String(50, "=") & vbCrLf & vbCrLf
    resultMessage = resultMessage & "Mailbox: " & mailboxName & vbCrLf
    resultMessage = resultMessage & "Folder: " & Replace(folderPath, "/", "\") & vbCrLf
    resultMessage = resultMessage & "Date Range: " & Format(startDate, "yyyy-mm-dd") & " to " & Format(endDate, "yyyy-mm-dd") & vbCrLf
    resultMessage = resultMessage & String(50, "-") & vbCrLf & vbCrLf
    resultMessage = resultMessage & "Total emails in date range: " & emailCount & vbCrLf & vbCrLf

    If emailCount > 0 Then
        resultMessage = resultMessage & "Breakdown by Category:" & vbCrLf
        resultMessage = resultMessage & String(50, "-") & vbCrLf

        ' Sort and display categories
        Dim sortedKeys() As Variant
        sortedKeys = categoryDict.Keys

        ' Simple bubble sort for categories
        Dim temp As Variant
        Dim j As Long
        For i = LBound(sortedKeys) To UBound(sortedKeys) - 1
            For j = i + 1 To UBound(sortedKeys)
                If sortedKeys(i) > sortedKeys(j) Then
                    temp = sortedKeys(i)
                    sortedKeys(i) = sortedKeys(j)
                    sortedKeys(j) = temp
                End If
            Next j
        Next i

        ' Display sorted categories
        For i = LBound(sortedKeys) To UBound(sortedKeys)
            categoryName = sortedKeys(i)
            resultMessage = resultMessage & "  " & categoryName & ": " & categoryDict(categoryName) & vbCrLf
        Next i
    Else
        resultMessage = resultMessage & "No emails found in the specified date range."
    End If

    ' Display the result
    MsgBox resultMessage, vbInformation, "Email Count Report"

    ' Clean up
    Set categoryDict = Nothing
    Set restrictedItems = Nothing
    Set objMail = Nothing
    Set objItem = Nothing
    Set objTestFolder = Nothing
    Set currentFolder = Nothing
    Set objInbox = Nothing
    Set objMailbox = Nothing
    Set objNamespace = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please verify that:" & vbCrLf & _
           "1. The mailbox '" & mailboxName & "' exists in your Outlook profile" & vbCrLf & _
           "2. The folder path '" & folderPath & "' exists in this mailbox" & vbCrLf & _
           "3. The folder path format is correct (e.g., Inbox/subfolder1/subfolder2)", _
           vbCritical, "Error Accessing Folder"

    ' Clean up
    Set categoryDict = Nothing
    Set restrictedItems = Nothing
    Set objMail = Nothing
    Set objItem = Nothing
    Set objTestFolder = Nothing
    Set currentFolder = Nothing
    Set objInbox = Nothing
    Set objMailbox = Nothing
    Set objNamespace = Nothing
End Sub

' Helper function to validate YYYYMMDD format
Function IsValidYYYYMMDD(dateStr As String) As Boolean
    Dim yyyy As Integer
    Dim mm As Integer
    Dim dd As Integer

    IsValidYYYYMMDD = False

    ' Check length
    If Len(dateStr) <> 8 Then Exit Function

    ' Check if all characters are numeric
    If Not IsNumeric(dateStr) Then Exit Function

    ' Parse components
    yyyy = CInt(Left(dateStr, 4))
    mm = CInt(Mid(dateStr, 5, 2))
    dd = CInt(Right(dateStr, 2))

    ' Validate year (1900-2100)
    If yyyy < 1900 Or yyyy > 2100 Then Exit Function

    ' Validate month
    If mm < 1 Or mm > 12 Then Exit Function

    ' Validate day
    If dd < 1 Or dd > 31 Then Exit Function

    ' Check if the date is valid using IsDate
    Dim testDate As String
    testDate = mm & "/" & dd & "/" & yyyy
    If Not IsDate(testDate) Then Exit Function

    IsValidYYYYMMDD = True
End Function

' Helper function to parse YYYYMMDD to Date
Function ParseYYYYMMDD(dateStr As String) As Date
    Dim yyyy As Integer
    Dim mm As Integer
    Dim dd As Integer

    yyyy = CInt(Left(dateStr, 4))
    mm = CInt(Mid(dateStr, 5, 2))
    dd = CInt(Right(dateStr, 2))

    ParseYYYYMMDD = DateSerial(yyyy, mm, dd)
End Function

Option Explicit

'====================================================
' OUTREACH ASSISTANT
'
' Production version for external client outreach.
'
' Important safeguards retained:
' - Invalid email addresses are skipped.
' - Email values of 0 are skipped.
' - Drafts must be generated before sending.
' - Drafts are tagged with a unique Run ID.
' - Changing campaigns invalidates the active batch.
' - Two confirmations are required before sending.
' - Draft creation and send results are logged.
'====================================================


'====================================================
' PREVIEW CAMPAIGN
'====================================================

Public Sub PreviewCampaign()

    Dim wsDash As Worksheet
    Dim wsTemplates As Worksheet
    Dim wsRecipients As Worksheet

    Dim selectedTemplate As String
    Dim subjectLine As String
    Dim emailBody As String

    Dim outreachType As String
    Dim emailAddress As String

    Dim lastTemplateRow As Long
    Dim lastRecipientRow As Long

    Dim recipientCount As Long
    Dim skippedCount As Long

    Dim templateFound As Boolean
    Dim i As Long

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Set wsTemplates = ThisWorkbook.Worksheets("Templates")
    Set wsRecipients = ThisWorkbook.Worksheets("Recipients")

    selectedTemplate = Trim$(CStr(wsDash.Range("B4").Value))

    If selectedTemplate = "" Then

        MsgBox _
            "Please select a campaign template first.", _
            vbExclamation, _
            "Template Required"

        Exit Sub

    End If

    lastTemplateRow = _
        wsTemplates.Cells( _
            wsTemplates.Rows.Count, _
            "A").End(xlUp).Row

    templateFound = False

    For i = 2 To lastTemplateRow

        If StrComp( _
            Trim$(CStr(wsTemplates.Cells(i, 1).Value)), _
            selectedTemplate, _
            vbTextCompare) = 0 Then

            subjectLine = CStr(wsTemplates.Cells(i, 2).Value)
            emailBody = CStr(wsTemplates.Cells(i, 3).Value)

            templateFound = True

            Exit For

        End If

    Next i

    If templateFound = False Then

        MsgBox _
            "The selected campaign template was not found " & _
            "on the Templates sheet.", _
            vbCritical, _
            "Template Not Found"

        Exit Sub

    End If

    wsDash.Range("A11").Value = subjectLine
    wsDash.Range("A15").Value = emailBody

    recipientCount = 0
    skippedCount = 0

    lastRecipientRow = _
        wsRecipients.Cells( _
            wsRecipients.Rows.Count, _
            "A").End(xlUp).Row

    For i = 2 To lastRecipientRow

        outreachType = _
            Trim$(CStr(wsRecipients.Cells(i, 4).Value))

        If StrComp( _
            outreachType, _
            selectedTemplate, _
            vbTextCompare) = 0 Then

            emailAddress = _
                Trim$(CStr(wsRecipients.Cells(i, 3).Value))

            If IsUsableEmail(emailAddress) Then
                recipientCount = recipientCount + 1
            Else
                skippedCount = skippedCount + 1
            End If

        End If

    Next i

    wsDash.Range("F4").Value = recipientCount
    wsDash.Range("F5").Value = "Live Campaign"
    wsDash.Range("F7").Value = "Not Generated"
    wsDash.Range("F8").Value = "Not Sent"
    wsDash.Range("F9").Value = "Preview Campaign"

    'A new preview invalidates any earlier draft batch.
    wsDash.Range("J4:J6").ClearContents

    MsgBox _
        "Campaign preview loaded." & vbCrLf & vbCrLf & _
        "Valid recipients: " & recipientCount & vbCrLf & _
        "Skipped for missing or invalid email: " & skippedCount, _
        vbInformation, _
        "Campaign Preview"

End Sub


'====================================================
' GENERATE CAMPAIGN DRAFTS
'====================================================

Public Sub GenerateDraftsFromDashboard()

    Const olFolderDrafts As Long = 16
    Const olText As Long = 1

    Dim wsDash As Worksheet
    Dim wsTemplates As Worksheet
    Dim wsRecipients As Worksheet
    Dim wsLogs As Worksheet

    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim DraftsFolder As Object
    Dim OutlookMail As Object

    Dim selectedTemplate As String
    Dim subjectLine As String
    Dim emailBody As String
    Dim personalizedBody As String

    Dim clientName As String
    Dim firstName As String
    Dim emailAddress As String
    Dim outreachType As String
    Dim ccString As String

    Dim runID As String

    Dim lastTemplateRow As Long
    Dim lastRecipientRow As Long
    Dim nextLogRow As Long

    Dim templateFound As Boolean

    Dim draftCount As Long
    Dim skippedCount As Long
    Dim oldDraftCount As Long

    Dim i As Long

    Dim runProperty As Object
    Dim campaignProperty As Object
    Dim clientProperty As Object

    On Error GoTo GenerateError

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Set wsTemplates = ThisWorkbook.Worksheets("Templates")
    Set wsRecipients = ThisWorkbook.Worksheets("Recipients")
    Set wsLogs = ThisWorkbook.Worksheets("Logs")

    selectedTemplate = Trim$(CStr(wsDash.Range("B4").Value))

    If selectedTemplate = "" Then

        MsgBox _
            "Please select a campaign template first.", _
            vbExclamation, _
            "Template Required"

        Exit Sub

    End If

    lastTemplateRow = _
        wsTemplates.Cells( _
            wsTemplates.Rows.Count, _
            "A").End(xlUp).Row

    templateFound = False

    For i = 2 To lastTemplateRow

        If StrComp( _
            Trim$(CStr(wsTemplates.Cells(i, 1).Value)), _
            selectedTemplate, _
            vbTextCompare) = 0 Then

            subjectLine = CStr(wsTemplates.Cells(i, 2).Value)
            emailBody = CStr(wsTemplates.Cells(i, 3).Value)

            templateFound = True

            Exit For

        End If

    Next i

    If templateFound = False Then

        MsgBox _
            "The selected campaign template was not found " & _
            "on the Templates sheet.", _
            vbCritical, _
            "Template Not Found"

        Exit Sub

    End If

    lastRecipientRow = _
        wsRecipients.Cells( _
            wsRecipients.Rows.Count, _
            "A").End(xlUp).Row

    draftCount = 0
    skippedCount = 0

    'Count valid recipients before opening Outlook.
    For i = 2 To lastRecipientRow

        outreachType = _
            Trim$(CStr(wsRecipients.Cells(i, 4).Value))

        If StrComp( _
            outreachType, _
            selectedTemplate, _
            vbTextCompare) = 0 Then

            emailAddress = _
                Trim$(CStr(wsRecipients.Cells(i, 3).Value))

            If IsUsableEmail(emailAddress) Then
                draftCount = draftCount + 1
            Else
                skippedCount = skippedCount + 1
            End If

        End If

    Next i

    If draftCount = 0 Then

        MsgBox _
            "No valid recipients were found for this campaign." & _
            vbCrLf & vbCrLf & _
            "Rows with blank, zero, or invalid email addresses " & _
            "are not eligible for draft generation.", _
            vbExclamation, _
            "No Valid Recipients"

        Exit Sub

    End If

    draftCount = 0

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")

    Set DraftsFolder = _
        OutlookNamespace.GetDefaultFolder(olFolderDrafts)

    oldDraftCount = _
        DeleteOldCampaignDrafts( _
            DraftsFolder, _
            selectedTemplate)

    runID = _
        selectedTemplate & "_" & _
        Format$(Now, "yyyymmdd_hhnnss")

    ccString = BuildCCString()

    For i = 2 To lastRecipientRow

        outreachType = _
            Trim$(CStr(wsRecipients.Cells(i, 4).Value))

        If StrComp( _
            outreachType, _
            selectedTemplate, _
            vbTextCompare) = 0 Then

            clientName = _
                Trim$(CStr(wsRecipients.Cells(i, 1).Value))

            firstName = _
                Trim$(CStr(wsRecipients.Cells(i, 2).Value))

            emailAddress = _
                Trim$(CStr(wsRecipients.Cells(i, 3).Value))

            If IsUsableEmail(emailAddress) Then

                personalizedBody = emailBody

                personalizedBody = Replace( _
                    personalizedBody, _
                    "{FirstName}", _
                    firstName, _
                    1, _
                    -1, _
                    vbTextCompare)

                personalizedBody = Replace( _
                    personalizedBody, _
                    "{ClientName}", _
                    clientName, _
                    1, _
                    -1, _
                    vbTextCompare)

                Set OutlookMail = OutlookApp.CreateItem(0)

                With OutlookMail

                    .To = emailAddress
                    .CC = ccString
                    .Subject = subjectLine

                    'Display is required so Outlook loads
                    'the current user's default signature.
                    .Display

                    DoEvents

                    .HTMLBody = _
                        PlainTextToHTML(personalizedBody) & _
                        "<br><br>" & _
                        .HTMLBody

                    Set runProperty = _
                        .UserProperties.Add( _
                            "OARunID", _
                            olText, _
                            False)

                    runProperty.Value = runID

                    Set campaignProperty = _
                        .UserProperties.Add( _
                            "OACampaign", _
                            olText, _
                            False)

                    campaignProperty.Value = selectedTemplate

                    Set clientProperty = _
                        .UserProperties.Add( _
                            "OAClientName", _
                            olText, _
                            False)

                    clientProperty.Value = clientName

                    .Save
                    .Close 0

                End With

                draftCount = draftCount + 1

                nextLogRow = _
                    wsLogs.Cells( _
                        wsLogs.Rows.Count, _
                        "A").End(xlUp).Row + 1

                wsLogs.Cells(nextLogRow, 1).Value = Now
                wsLogs.Cells(nextLogRow, 2).Value = clientName
                wsLogs.Cells(nextLogRow, 3).Value = emailAddress
                wsLogs.Cells(nextLogRow, 4).Value = subjectLine

                wsLogs.Cells(nextLogRow, 5).Value = _
                    "Draft Created | Run: " & runID

                Set OutlookMail = Nothing
                Set runProperty = Nothing
                Set campaignProperty = Nothing
                Set clientProperty = Nothing

            End If

        End If

    Next i

    'Store the active campaign batch.
    wsDash.Range("J4").Value = runID
    wsDash.Range("J5").Value = selectedTemplate
    wsDash.Range("J6").Value = draftCount

    wsDash.Range("F4").Value = draftCount

    wsDash.Range("F7").Value = _
        draftCount & " Drafts Generated"

    wsDash.Range("F8").Value = "Not Sent"
    wsDash.Range("F9").Value = "Generate Drafts"

    MsgBox _
        draftCount & " drafts were created." & _
        vbCrLf & vbCrLf & _
        skippedCount & _
        " rows were skipped for missing or invalid email." & _
        vbCrLf & vbCrLf & _
        oldDraftCount & _
        " previous unsent drafts for this campaign were cleared.", _
        vbInformation, _
        "Draft Generation Complete"

    Exit Sub

GenerateError:

    wsDash.Range("F9").Value = "Draft Generation Error"

    MsgBox _
        "Draft generation stopped because of an error." & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Draft Generation Error"

End Sub


'====================================================
' VALIDATE CAMPAIGN
'
' This procedure can remain connected to a separate
' validation button if one exists.
'====================================================

Public Sub ValidateCampaign()

    Dim wsDash As Worksheet
    Dim response As VbMsgBoxResult

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")

    If Trim$(CStr(wsDash.Range("B4").Value)) = "" Then

        MsgBox _
            "No campaign template is selected.", _
            vbExclamation, _
            "Template Required"

        Exit Sub

    End If

    If Val(wsDash.Range("F4").Value) = 0 Then

        MsgBox _
            "No valid recipients were found for this campaign.", _
            vbExclamation, _
            "No Recipients"

        Exit Sub

    End If

    If InStr( _
        1, _
        CStr(wsDash.Range("F7").Value), _
        "Drafts Generated", _
        vbTextCompare) = 0 Then

        MsgBox _
            "Please generate the drafts before validating " & _
            "the campaign.", _
            vbExclamation, _
            "Drafts Required"

        Exit Sub

    End If

    response = MsgBox( _
        "CAMPAIGN SUMMARY" & vbCrLf & vbCrLf & _
        "Template: " & wsDash.Range("B4").Value & vbCrLf & _
        "Recipients: " & wsDash.Range("F4").Value & vbCrLf & _
        "Draft Status: " & wsDash.Range("F7").Value & vbCrLf & _
        "Campaign Status: " & wsDash.Range("F8").Value & _
        vbCrLf & vbCrLf & _
        "Continue?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Campaign Validation")

    If response = vbYes Then

        MsgBox _
            "The campaign passed the initial validation." & _
            vbCrLf & vbCrLf & _
            "Review every draft before using Send Campaign.", _
            vbInformation, _
            "Validation Complete"

        wsDash.Range("F9").Value = "Validate Campaign"

    Else

        wsDash.Range("F9").Value = "Validation Cancelled"

    End If

End Sub


'====================================================
' SEND ACTIVE CAMPAIGN
'====================================================

Public Sub SendCampaignFromDashboard()

    Const olFolderDrafts As Long = 16

    Dim wsDash As Worksheet
    Dim wsLogs As Worksheet

    Dim OutlookApp As Object
    Dim OutlookNamespace As Object
    Dim DraftsFolder As Object
    Dim DraftItems As Object
    Dim DraftItem As Object

    Dim runProperty As Object
    Dim clientProperty As Object

    Dim selectedTemplate As String
    Dim activeRunID As String
    Dim activeCampaign As String

    Dim emailAddress As String
    Dim clientName As String
    Dim subjectLine As String

    Dim eligibleCount As Long
    Dim expectedCount As Long
    Dim sentCount As Long
    Dim failedCount As Long
    Dim nextLogRow As Long

    Dim response As VbMsgBoxResult
    Dim finalResponse As VbMsgBoxResult

    Dim i As Long

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Set wsLogs = ThisWorkbook.Worksheets("Logs")

    selectedTemplate = _
        Trim$(CStr(wsDash.Range("B4").Value))

    activeRunID = _
        Trim$(CStr(wsDash.Range("J4").Value))

    activeCampaign = _
        Trim$(CStr(wsDash.Range("J5").Value))

    expectedCount = Val(wsDash.Range("J6").Value)

    If selectedTemplate = "" Then

        MsgBox _
            "No campaign template is selected.", _
            vbExclamation, _
            "Template Required"

        Exit Sub

    End If

    If activeRunID = "" Then

        MsgBox _
            "There is no active draft batch." & _
            vbCrLf & vbCrLf & _
            "Please preview the campaign and generate " & _
            "the drafts first.", _
            vbExclamation, _
            "Draft Batch Required"

        Exit Sub

    End If

    If StrComp( _
        selectedTemplate, _
        activeCampaign, _
        vbTextCompare) <> 0 Then

        MsgBox _
            "The selected template has changed since the drafts " & _
            "were generated." & vbCrLf & vbCrLf & _
            "Please preview the new campaign and generate its " & _
            "drafts before sending.", _
            vbExclamation, _
            "Campaign Changed"

        Exit Sub

    End If

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")

    Set DraftsFolder = _
        OutlookNamespace.GetDefaultFolder(olFolderDrafts)

    Set DraftItems = DraftsFolder.Items

    eligibleCount = 0

    'First pass validates the complete active batch
    'before any messages are sent.
    For i = DraftItems.Count To 1 Step -1

        Set DraftItem = DraftItems.Item(i)

        If TypeName(DraftItem) = "MailItem" Then

            Set runProperty = _
                DraftItem.UserProperties.Find("OARunID")

            If Not runProperty Is Nothing Then

                If CStr(runProperty.Value) = activeRunID Then

                    emailAddress = _
                        Trim$(CStr(DraftItem.To))

                    If Not IsUsableEmail(emailAddress) Then

                        MsgBox _
                            "The active draft batch contains an " & _
                            "invalid recipient address:" & _
                            vbCrLf & vbCrLf & _
                            emailAddress & vbCrLf & vbCrLf & _
                            "Nothing was sent. Delete or correct " & _
                            "the affected draft, then regenerate " & _
                            "the campaign.", _
                            vbCritical, _
                            "Invalid Draft Recipient"

                        Exit Sub

                    End If

                    eligibleCount = eligibleCount + 1

                End If

            End If

        End If

        Set runProperty = Nothing
        Set DraftItem = Nothing

    Next i

    If eligibleCount = 0 Then

        MsgBox _
            "No active campaign drafts were found." & _
            vbCrLf & vbCrLf & _
            "Please generate the drafts again.", _
            vbExclamation, _
            "No Drafts Found"

        Exit Sub

    End If

    If expectedCount > 0 Then

        If eligibleCount <> expectedCount Then

            MsgBox _
                "The active draft batch does not match the batch " & _
                "that was originally generated." & _
                vbCrLf & vbCrLf & _
                "Expected drafts: " & expectedCount & vbCrLf & _
                "Drafts currently found: " & eligibleCount & _
                vbCrLf & vbCrLf & _
                "Nothing was sent. Please regenerate the drafts " & _
                "so the complete campaign can be reviewed again.", _
                vbCritical, _
                "Draft Count Mismatch"

            Exit Sub

        End If

    End If

    response = MsgBox( _
        "CAMPAIGN SUMMARY" & vbCrLf & vbCrLf & _
        "Template: " & selectedTemplate & vbCrLf & _
        "Recipients: " & eligibleCount & vbCrLf & _
        "Draft Status: " & wsDash.Range("F7").Value & vbCrLf & _
        "Mode: " & wsDash.Range("F5").Value & _
        vbCrLf & vbCrLf & _
        "Continue to the final send confirmation?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Campaign Validation")

    If response <> vbYes Then

        wsDash.Range("F9").Value = "Send Cancelled"

        Exit Sub

    End If

    finalResponse = MsgBox( _
        "FINAL WARNING" & vbCrLf & vbCrLf & _
        "This action will send " & eligibleCount & _
        " emails to external recipients." & _
        vbCrLf & vbCrLf & _
        "Template: " & selectedTemplate & vbCrLf & _
        "Mode: " & wsDash.Range("F5").Value & _
        vbCrLf & vbCrLf & _
        "This action cannot be undone." & _
        vbCrLf & vbCrLf & _
        "Send the campaign now?", _
        vbYesNo + vbExclamation + vbDefaultButton2, _
        "Send Campaign")

    If finalResponse <> vbYes Then

        wsDash.Range("F9").Value = "Send Cancelled"

        Exit Sub

    End If

    sentCount = 0
    failedCount = 0

    'Second pass sends only drafts tagged
    'with the current active Run ID.
    For i = DraftItems.Count To 1 Step -1

        Set DraftItem = DraftItems.Item(i)

        If TypeName(DraftItem) = "MailItem" Then

            Set runProperty = _
                DraftItem.UserProperties.Find("OARunID")

            If Not runProperty Is Nothing Then

                If CStr(runProperty.Value) = activeRunID Then

                    On Error GoTo SendError

                    emailAddress = _
                        Trim$(CStr(DraftItem.To))

                    subjectLine = CStr(DraftItem.Subject)

                    clientName = ""

                    Set clientProperty = _
                        DraftItem.UserProperties.Find("OAClientName")

                    If Not clientProperty Is Nothing Then
                        clientName = CStr(clientProperty.Value)
                    End If

                    DraftItem.Send

                    sentCount = sentCount + 1

                    nextLogRow = _
                        wsLogs.Cells( _
                            wsLogs.Rows.Count, _
                            "A").End(xlUp).Row + 1

                    wsLogs.Cells(nextLogRow, 1).Value = Now
                    wsLogs.Cells(nextLogRow, 2).Value = clientName
                    wsLogs.Cells(nextLogRow, 3).Value = emailAddress
                    wsLogs.Cells(nextLogRow, 4).Value = subjectLine

                    wsLogs.Cells(nextLogRow, 5).Value = _
                        "Sent | Run: " & activeRunID

                    On Error GoTo 0

                End If

            End If

        End If

ContinueSending:

        Set runProperty = Nothing
        Set clientProperty = Nothing
        Set DraftItem = Nothing

    Next i

    If failedCount = 0 Then

        wsDash.Range("F8").Value = _
            sentCount & " Emails Sent"

        wsDash.Range("F9").Value = "Send Campaign"

        'Clear the completed active-run state.
        wsDash.Range("J4:J6").ClearContents

        MsgBox _
            sentCount & " emails were sent successfully.", _
            vbInformation, _
            "Campaign Sent"

    Else

        wsDash.Range("F8").Value = _
            sentCount & " Sent, " & failedCount & " Failed"

        wsDash.Range("F9").Value = "Partial Send"

        MsgBox _
            "The campaign completed with errors." & _
            vbCrLf & vbCrLf & _
            "Sent: " & sentCount & vbCrLf & _
            "Failed: " & failedCount & vbCrLf & vbCrLf & _
            "The active campaign state has been retained. " & _
            "Review the Logs sheet before taking another action.", _
            vbExclamation, _
            "Partial Send"

    End If

    Exit Sub

SendError:

    failedCount = failedCount + 1

    nextLogRow = _
        wsLogs.Cells( _
            wsLogs.Rows.Count, _
            "A").End(xlUp).Row + 1

    wsLogs.Cells(nextLogRow, 1).Value = Now
    wsLogs.Cells(nextLogRow, 2).Value = clientName
    wsLogs.Cells(nextLogRow, 3).Value = emailAddress
    wsLogs.Cells(nextLogRow, 4).Value = subjectLine

    wsLogs.Cells(nextLogRow, 5).Value = _
        "SEND FAILED: " & Err.Description & _
        " | Run: " & activeRunID

    Err.Clear
    On Error GoTo 0

    Resume ContinueSending

End Sub


'====================================================
' DELETE PREVIOUS UNSENT DRAFTS
'
' Deletes only drafts that were created by this
' workbook for the same campaign.
'====================================================

Private Function DeleteOldCampaignDrafts( _
    ByVal DraftsFolder As Object, _
    ByVal campaignName As String) As Long

    Dim DraftItems As Object
    Dim DraftItem As Object
    Dim campaignProperty As Object

    Dim deletedCount As Long
    Dim i As Long

    Set DraftItems = DraftsFolder.Items

    deletedCount = 0

    For i = DraftItems.Count To 1 Step -1

        Set DraftItem = DraftItems.Item(i)

        If TypeName(DraftItem) = "MailItem" Then

            Set campaignProperty = _
                DraftItem.UserProperties.Find("OACampaign")

            If Not campaignProperty Is Nothing Then

                If StrComp( _
                    CStr(campaignProperty.Value), _
                    campaignName, _
                    vbTextCompare) = 0 Then

                    DraftItem.Delete

                    deletedCount = deletedCount + 1

                End If

            End If

        End If

        Set campaignProperty = Nothing
        Set DraftItem = Nothing

    Next i

    DeleteOldCampaignDrafts = deletedCount

End Function


'====================================================
' EMAIL VALIDATION
'
' This is a practical validation for campaign safety.
' It rejects:
' - Blank values
' - Zero values
' - Addresses containing spaces
' - Missing or repeated @ signs
' - Missing domain dots
'====================================================

Private Function IsUsableEmail( _
    ByVal emailAddress As String) As Boolean

    Dim atPosition As Long
    Dim dotPosition As Long

    emailAddress = Trim$(CStr(emailAddress))

    If emailAddress = "" Then
        IsUsableEmail = False
        Exit Function
    End If

    If emailAddress = "0" Then
        IsUsableEmail = False
        Exit Function
    End If

    If InStr( _
        1, _
        emailAddress, _
        " ", _
        vbBinaryCompare) > 0 Then

        IsUsableEmail = False
        Exit Function

    End If

    atPosition = _
        InStr(1, emailAddress, "@", vbTextCompare)

    If atPosition <= 1 Then
        IsUsableEmail = False
        Exit Function
    End If

    If InStr( _
        atPosition + 1, _
        emailAddress, _
        "@", _
        vbTextCompare) > 0 Then

        IsUsableEmail = False
        Exit Function

    End If

    dotPosition = InStrRev(emailAddress, ".")

    If dotPosition <= atPosition + 1 Then
        IsUsableEmail = False
        Exit Function
    End If

    If dotPosition >= Len(emailAddress) Then
        IsUsableEmail = False
        Exit Function
    End If

    IsUsableEmail = True

End Function


'====================================================
' FORMAT TEMPLATE TEXT AS HTML
'
' Encodes characters that could otherwise be
' interpreted as HTML and preserves line breaks.
'====================================================

Private Function PlainTextToHTML( _
    ByVal sourceText As String) As String

    Dim htmlText As String

    htmlText = CStr(sourceText)

    htmlText = Replace(htmlText, "&", "&amp;")
    htmlText = Replace(htmlText, "<", "&lt;")
    htmlText = Replace(htmlText, ">", "&gt;")
    htmlText = Replace(htmlText, """", "&quot;")

    htmlText = Replace(htmlText, vbCrLf, "<br>")
    htmlText = Replace(htmlText, vbCr, "<br>")
    htmlText = Replace(htmlText, vbLf, "<br>")

    PlainTextToHTML = htmlText

End Function


'====================================================
' BUILD CC STRING
'====================================================

Private Function BuildCCString() As String

    Dim wsDash As Worksheet
    Dim wsSettings As Worksheet

    Dim primaryCC As String
    Dim secondaryCC As String

    Dim primaryEmail As String
    Dim secondaryEmail As String

    Dim ccString As String

    Dim lastRow As Long
    Dim i As Long

    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    Set wsSettings = ThisWorkbook.Worksheets("Settings")

    primaryCC = Trim$(CStr(wsDash.Range("B6").Value))
    secondaryCC = Trim$(CStr(wsDash.Range("B7").Value))

    primaryEmail = ""
    secondaryEmail = ""
    ccString = ""

    lastRow = _
        wsSettings.Cells( _
            wsSettings.Rows.Count, _
            "A").End(xlUp).Row

    If primaryCC <> "" And _
       StrComp(primaryCC, "None", vbTextCompare) <> 0 Then

        For i = 2 To lastRow

            If StrComp( _
                Trim$(CStr(wsSettings.Cells(i, 1).Value)), _
                primaryCC, _
                vbTextCompare) = 0 Then

                primaryEmail = _
                    Trim$(CStr(wsSettings.Cells(i, 2).Value))

                Exit For

            End If

        Next i

    End If

    If secondaryCC <> "" And _
       StrComp(secondaryCC, "None", vbTextCompare) <> 0 Then

        For i = 2 To lastRow

            If StrComp( _
                Trim$(CStr(wsSettings.Cells(i, 1).Value)), _
                secondaryCC, _
                vbTextCompare) = 0 Then

                secondaryEmail = _
                    Trim$(CStr(wsSettings.Cells(i, 2).Value))

                Exit For

            End If

        Next i

    End If

    If IsUsableEmail(primaryEmail) Then
        ccString = primaryEmail
    End If

    If IsUsableEmail(secondaryEmail) Then

        If StrComp( _
            secondaryEmail, _
            primaryEmail, _
            vbTextCompare) <> 0 Then

            If ccString <> "" Then
                ccString = ccString & ";"
            End If

            ccString = ccString & secondaryEmail

        End If

    End If

    BuildCCString = ccString

End Function


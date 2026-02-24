Attribute VB_Name = "modEmail"
Option Explicit

Public Sub SendUpdateRequestsPreview(): SendUpdateRequestsInternal "Preview": End Sub
Public Sub SendUpdateRequestsSend(): SendUpdateRequestsInternal "Send": End Sub

Private Sub SendUpdateRequestsInternal(ByVal modeText As String)
    On Error GoTo CleanFail

    Dim olApp As Object, mailItem As Object
    Dim tblTasks As ListObject
    Dim ownerBuckets As Object
    Dim rowTask As ListRow, ownerKey As String
    Dim bodyText As String, subjectText As String

    Set olApp = SafeOutlookApp()
    If olApp Is Nothing Then
        UserError "Outlook is unavailable. Open Outlook and try again."
        Exit Sub
    End If

    Set tblTasks = GetTable(SH_TASKS, TBL_TASKS)
    Set ownerBuckets = CreateObject("Scripting.Dictionary")

    For Each rowTask In tblTasks.ListRows
        If IsTaskNeedsUpdate(CStr(rowTask.Range.Cells(1, GetColumnIndex(tblTasks, "Status")).Value), _
                             rowTask.Range.Cells(1, GetColumnIndex(tblTasks, "LastUpdated")).Value, _
                             rowTask.Range.Cells(1, GetColumnIndex(tblTasks, "UpdateNeededBy")).Value) Then
            ownerKey = CStr(rowTask.Range.Cells(1, GetColumnIndex(tblTasks, "OwnerEmail")).Value)
            If ownerKey <> "" Then
                If Not ownerBuckets.Exists(ownerKey) Then ownerBuckets(ownerKey) = New Collection
                ownerBuckets(ownerKey).Add rowTask
            End If
        End If
    Next rowTask

    Dim ownerEmail As Variant
    For Each ownerEmail In ownerBuckets.Keys
        subjectText = "Action needed: Update your project tasks"
        bodyText = BuildOwnerEmailBody(ownerBuckets(ownerEmail), tblTasks)

        Set mailItem = olApp.CreateItem(0)
        mailItem.To = CStr(ownerEmail)
        mailItem.Subject = subjectText
        mailItem.Body = bodyText

        If modeText = "Send" Then mailItem.Send Else mailItem.Display
        LogEmail ResolveOwnerName(ownerBuckets(ownerEmail), tblTasks), CStr(ownerEmail), subjectText, ExtractTaskIDs(ownerBuckets(ownerEmail), tblTasks), modeText
    Next ownerEmail

    UserMessage "Update request emails processed in " & modeText & " mode."
    Exit Sub

CleanFail:
    UserError "Email automation failed: " & Err.Description
End Sub

Private Function BuildOwnerEmailBody(ByVal rows As Collection, ByVal tblTasks As ListObject) As String
    Dim txt As String, i As Long
    Dim taskRow As ListRow

    txt = "Hello," & vbCrLf & vbCrLf & "Please update the following tasks:" & vbCrLf & vbCrLf
    For i = 1 To rows.Count
        Set taskRow = rows(i)
        txt = txt & "- " & taskRow.Range.Cells(1, GetColumnIndex(tblTasks, "TaskID")).Value & " | " & _
              taskRow.Range.Cells(1, GetColumnIndex(tblTasks, "TaskName")).Value & " | Due: " & _
              Format$(taskRow.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value, "yyyy-mm-dd") & " | Status: " & _
              taskRow.Range.Cells(1, GetColumnIndex(tblTasks, "Status")).Value & vbCrLf
    Next i

    BuildOwnerEmailBody = txt & vbCrLf & "Instructions:" & vbCrLf & _
        "Open workbook -> Home -> MyUpdates -> select your name -> update highlighted rows -> save."
End Function

Private Function ExtractTaskIDs(ByVal rows As Collection, ByVal tblTasks As ListObject) As String
    Dim ids As String, i As Long
    For i = 1 To rows.Count
        ids = ids & IIf(ids = "", "", ",") & CStr(rows(i).Range.Cells(1, GetColumnIndex(tblTasks, "TaskID")).Value)
    Next i
    ExtractTaskIDs = ids
End Function

Private Function ResolveOwnerName(ByVal rows As Collection, ByVal tblTasks As ListObject) As String
    If rows.Count = 0 Then Exit Function
    ResolveOwnerName = CStr(rows(1).Range.Cells(1, GetColumnIndex(tblTasks, "OwnerName")).Value)
End Function

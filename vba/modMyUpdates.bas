Attribute VB_Name = "modMyUpdates"
Option Explicit

Public Sub RefreshMyUpdates()
    On Error GoTo CleanFail

    Dim ownerName As String
    Dim tblTasks As ListObject, tblView As ListObject
    Dim r As ListRow, newRow As ListRow

    ownerName = CStr(GetNamedRangeValue(NM_MYUPDATES_OWNER, ""))
    If ownerName = "" Then
        UserError "Select an owner in MyUpdates first."
        Exit Sub
    End If

    Set tblTasks = GetTable(SH_TASKS, TBL_TASKS)
    Set tblView = GetTable(SH_MYUPDATES, TBL_MYUPDATES)

    On Error Resume Next
    If Not tblView.DataBodyRange Is Nothing Then tblView.DataBodyRange.Delete
    On Error GoTo CleanFail

    For Each r In tblTasks.ListRows
        If CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "OwnerName")).Value) = ownerName Then
            Set newRow = tblView.ListRows.Add
            newRow.Range.Resize(1, tblTasks.ListColumns.Count).Value = r.Range.Value
            newRow.Range.Cells(1, GetColumnIndex(tblView, "NeedsUpdate")).Value = _
                IIf(IsTaskNeedsUpdate(CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "Status")).Value), _
                                     r.Range.Cells(1, GetColumnIndex(tblTasks, "LastUpdated")).Value, _
                                     r.Range.Cells(1, GetColumnIndex(tblTasks, "UpdateNeededBy")).Value), "Y", "")
        End If
    Next r

    HighlightNeedsUpdate tblView
    UserMessage "MyUpdates refreshed for " & ownerName
    Exit Sub

CleanFail:
    UserError "Unable to refresh MyUpdates: " & Err.Description
End Sub

Private Sub HighlightNeedsUpdate(ByVal tblView As ListObject)
    If tblView.DataBodyRange Is Nothing Then Exit Sub

    Dim idxNeeds As Long
    idxNeeds = GetColumnIndex(tblView, "NeedsUpdate")

    Dim rowRange As Range
    For Each rowRange In tblView.DataBodyRange.Rows
        If UCase$(CStr(rowRange.Cells(1, idxNeeds).Value)) = "Y" Then
            rowRange.Interior.Color = RGB(255, 242, 204)
        Else
            rowRange.Interior.Pattern = xlNone
        End If
    Next rowRange
End Sub

Public Sub UpdateSelectedTask(): ApplyTaskUpdate uaGeneralUpdate: End Sub
Public Sub MarkDone(): ApplyTaskUpdate uaMarkDone: End Sub
Public Sub SetBlocked(): ApplyTaskUpdate uaSetBlocked: End Sub
Public Sub AddComment(): ApplyTaskUpdate uaAddComment: End Sub

Private Sub ApplyTaskUpdate(ByVal actionType As UpdateAction)
    On Error GoTo CleanFail

    Dim tblView As ListObject, tblTasks As ListObject
    Dim selectedTaskID As String
    Dim taskRow As ListRow

    Set tblView = GetTable(SH_MYUPDATES, TBL_MYUPDATES)
    Set tblTasks = GetTable(SH_TASKS, TBL_TASKS)

    If tblView.DataBodyRange Is Nothing Or Intersect(Selection, tblView.DataBodyRange) Is Nothing Then
        UserError "Select a task row in MyUpdates first."
        Exit Sub
    End If

    selectedTaskID = CStr(tblView.DataBodyRange.Cells(Selection.Row - tblView.DataBodyRange.Row + 1, GetColumnIndex(tblView, "TaskID")).Value)
    Set taskRow = FindTaskRowByID(tblTasks, selectedTaskID)
    If taskRow Is Nothing Then
        UserError "TaskID " & selectedTaskID & " not found in tblTasks."
        Exit Sub
    End If

    Select Case actionType
        Case uaMarkDone
            SetFieldWithLog taskRow, tblTasks, "Status", "Done"
            SetFieldWithLog taskRow, tblTasks, "PercentComplete", 100
        Case uaSetBlocked
            SetFieldWithLog taskRow, tblTasks, "Status", "Blocked"
            SetFieldWithLog taskRow, tblTasks, "RAG", "Red"
        Case uaAddComment
            SetFieldWithLog taskRow, tblTasks, "Comments", InputBox("Enter comment", "Task Comment")
        Case Else
            SyncViewRowToTask tblView, tblTasks, selectedTaskID
    End Select

    SetFieldWithLog taskRow, tblTasks, "LastUpdated", Now
    ValidateTasks False
    RefreshMyUpdates
    Exit Sub

CleanFail:
    UserError "Update failed: " & Err.Description
End Sub

Public Function FindTaskRowByID(ByVal tblTasks As ListObject, ByVal taskID As String) As ListRow
    Dim r As ListRow
    For Each r In tblTasks.ListRows
        If CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "TaskID")).Value) = taskID Then
            Set FindTaskRowByID = r
            Exit Function
        End If
    Next r
End Function

Public Sub SetFieldWithLog(ByVal rowTask As ListRow, ByVal tbl As ListObject, ByVal fieldName As String, ByVal newValue As Variant)
    Dim idx As Long, oldValue As Variant
    idx = GetColumnIndex(tbl, fieldName)
    oldValue = rowTask.Range.Cells(1, idx).Value
    If CStr(oldValue) <> CStr(newValue) Then
        rowTask.Range.Cells(1, idx).Value = newValue
        LogTaskChange CStr(rowTask.Range.Cells(1, GetColumnIndex(tbl, "TaskID")).Value), fieldName, oldValue, newValue
    End If
End Sub

Private Sub SyncViewRowToTask(ByVal tblView As ListObject, ByVal tblTasks As ListObject, ByVal taskID As String)
    Dim viewRow As Range
    Dim tRow As ListRow
    Dim trackFields As Variant, f As Variant

    Set tRow = FindTaskRowByID(tblTasks, taskID)
    Set viewRow = tblView.DataBodyRange.Rows(Selection.Row - tblView.DataBodyRange.Row + 1)
    trackFields = Array("Status", "StartDate", "EndDate", "PercentComplete", "OwnerName", "OwnerEmail", "RAG", "Comments")

    For Each f In trackFields
        SetFieldWithLog tRow, tblTasks, CStr(f), viewRow.Cells(1, GetColumnIndex(tblView, CStr(f))).Value
    Next f
End Sub

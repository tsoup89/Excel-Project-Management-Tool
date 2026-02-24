Attribute VB_Name = "modLogs"
Option Explicit

Public Sub LogTaskChange(ByVal taskID As String, ByVal fieldName As String, ByVal oldValue As Variant, ByVal newValue As Variant)
    On Error GoTo CleanFail

    Dim tblLog As ListObject
    Dim newRow As ListRow

    Set tblLog = GetTable(SH_LOGS, TBL_CHANGELOG)
    Set newRow = tblLog.ListRows.Add

    newRow.Range.Cells(1, GetColumnIndex(tblLog, "Timestamp")).Value = Now
    newRow.Range.Cells(1, GetColumnIndex(tblLog, "UserName")).Value = Environ$("Username")
    newRow.Range.Cells(1, GetColumnIndex(tblLog, "TaskID")).Value = taskID
    newRow.Range.Cells(1, GetColumnIndex(tblLog, "FieldName")).Value = fieldName
    newRow.Range.Cells(1, GetColumnIndex(tblLog, "OldValue")).Value = oldValue
    newRow.Range.Cells(1, GetColumnIndex(tblLog, "NewValue")).Value = newValue
    Exit Sub

CleanFail:
    ' Intentionally swallow logging errors to avoid interrupting UX.
End Sub

Public Sub LogEmail(ByVal ownerName As String, ByVal ownerEmail As String, ByVal subj As String, ByVal taskIDs As String, ByVal modeText As String)
    Dim tbl As ListObject
    Dim lr As ListRow

    Set tbl = GetTable(SH_LOGS, TBL_EMAILLOG)
    Set lr = tbl.ListRows.Add

    lr.Range.Cells(1, GetColumnIndex(tbl, "Timestamp")).Value = Now
    lr.Range.Cells(1, GetColumnIndex(tbl, "OwnerName")).Value = ownerName
    lr.Range.Cells(1, GetColumnIndex(tbl, "OwnerEmail")).Value = ownerEmail
    lr.Range.Cells(1, GetColumnIndex(tbl, "Subject")).Value = subj
    lr.Range.Cells(1, GetColumnIndex(tbl, "TaskIDs")).Value = taskIDs
    lr.Range.Cells(1, GetColumnIndex(tbl, "Mode")).Value = modeText
End Sub

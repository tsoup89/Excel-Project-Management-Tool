Attribute VB_Name = "modSync"
Option Explicit

Public Sub ExportMyUpdatesToCSV()
    On Error GoTo CleanFail
    Dim tbl As ListObject, pathOut As String

    Set tbl = GetTable(SH_MYUPDATES, TBL_MYUPDATES)
    pathOut = ThisWorkbook.Path & "\MyUpdates_" & Replace(CStr(GetNamedRangeValue(NM_MYUPDATES_OWNER, "Owner")), " ", "") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".csv"

    WriteRangeToCSV tbl.Range, pathOut
    UserMessage "Exported: " & pathOut
    Exit Sub
CleanFail:
    UserError "CSV export failed: " & Err.Description
End Sub

Public Sub ImportUpdatesFromCSV()
    On Error GoTo CleanFail

    Dim filePath As Variant, wb As Workbook, ws As Worksheet
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
    If filePath = False Then Exit Sub

    Set wb = Workbooks.Open(CStr(filePath))
    Set ws = wb.Worksheets(1)
    MergeImportedUpdates ws.UsedRange
    wb.Close SaveChanges:=False

    UserMessage "CSV updates imported with conflict warnings where applicable."
    Exit Sub
CleanFail:
    UserError "CSV import failed: " & Err.Description
End Sub

Private Sub MergeImportedUpdates(ByVal importRange As Range)
    Dim tblTasks As ListObject, headers As Object
    Dim i As Long, taskID As String, tRow As ListRow

    Set tblTasks = GetTable(SH_TASKS, TBL_TASKS)
    Set headers = HeaderMap(importRange)

    For i = 2 To importRange.Rows.Count
        taskID = CStr(importRange.Cells(i, headers("TaskID")).Value)
        Set tRow = FindTaskRowByID(tblTasks, taskID)
        If Not tRow Is Nothing Then
            If CDate(importRange.Cells(i, headers("LastUpdated")).Value) < CDate(Nz(tRow.Range.Cells(1, GetColumnIndex(tblTasks, "LastUpdated")).Value, #1/1/1900#)) Then
                UserError "Conflict on Task " & taskID & ": master is newer."
            Else
                SetFieldWithLog tRow, tblTasks, "Status", importRange.Cells(i, headers("Status")).Value
                SetFieldWithLog tRow, tblTasks, "PercentComplete", importRange.Cells(i, headers("PercentComplete")).Value
                SetFieldWithLog tRow, tblTasks, "Comments", importRange.Cells(i, headers("Comments")).Value
                SetFieldWithLog tRow, tblTasks, "LastUpdated", Now
            End If
        End If
    Next i
End Sub

Private Function HeaderMap(ByVal rng As Range) As Object
    Dim dict As Object, c As Long
    Set dict = CreateObject("Scripting.Dictionary")
    For c = 1 To rng.Columns.Count
        dict(CStr(rng.Cells(1, c).Value)) = c
    Next c
    Set HeaderMap = dict
End Function

Private Sub WriteRangeToCSV(ByVal rng As Range, ByVal filePath As String)
    Dim f As Integer, r As Long, c As Long, lineText As String
    f = FreeFile
    Open filePath For Output As #f

    For r = 1 To rng.Rows.Count
        lineText = ""
        For c = 1 To rng.Columns.Count
            lineText = lineText & IIf(c = 1, "", ",") & """" & Replace(CStr(rng.Cells(r, c).Value), """", """"") & """"
        Next c
        Print #f, lineText
    Next r

    Close #f
End Sub

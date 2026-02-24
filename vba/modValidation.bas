Attribute VB_Name = "modValidation"
Option Explicit

Public Sub ValidateTasks(Optional ByVal autoFixDates As Boolean = False)
    On Error GoTo CleanFail

    Dim tblTasks As ListObject
    Dim rowTask As ListRow
    Dim idxTaskID As Long, idxStart As Long, idxEnd As Long, idxDuration As Long
    Dim idxDepends As Long, idxStatus As Long, idxRAG As Long
    Dim taskMap As Object

    Set tblTasks = GetTable(SH_TASKS, TBL_TASKS)
    idxTaskID = GetColumnIndex(tblTasks, "TaskID")
    idxStart = GetColumnIndex(tblTasks, "StartDate")
    idxEnd = GetColumnIndex(tblTasks, "EndDate")
    idxDuration = GetColumnIndex(tblTasks, "Duration")
    idxDepends = GetColumnIndex(tblTasks, "Dependencies")
    idxStatus = GetColumnIndex(tblTasks, "Status")
    idxRAG = GetColumnIndex(tblTasks, "RAG")

    Set taskMap = CreateObject("Scripting.Dictionary")
    For Each rowTask In tblTasks.ListRows
        taskMap(CStr(rowTask.Range.Cells(1, idxTaskID).Value)) = rowTask.Index
    Next rowTask

    For Each rowTask In tblTasks.ListRows
        ValidateSingleTask rowTask, tblTasks, taskMap, idxTaskID, idxStart, idxEnd, idxDuration, idxDepends, autoFixDates
        rowTask.Range.Cells(1, idxRAG).Value = ResolveRAG(rowTask, tblTasks)
    Next rowTask

    UserMessage "Task validation complete."
    Exit Sub

CleanFail:
    UserError "Validation failed: " & Err.Description
End Sub

Private Sub ValidateSingleTask(ByVal rowTask As ListRow, ByVal tblTasks As ListObject, ByVal taskMap As Object, _
                               ByVal idxTaskID As Long, ByVal idxStart As Long, ByVal idxEnd As Long, _
                               ByVal idxDuration As Long, ByVal idxDepends As Long, ByVal autoFixDates As Boolean)

    Dim sDate As Variant, eDate As Variant
    Dim deps As Variant, dep As Variant
    Dim depTaskEnd As Variant

    sDate = rowTask.Range.Cells(1, idxStart).Value
    eDate = rowTask.Range.Cells(1, idxEnd).Value

    If IsDate(sDate) And IsDate(eDate) Then
        If CDate(sDate) > CDate(eDate) Then
            If autoFixDates Then
                rowTask.Range.Cells(1, idxEnd).Value = CDate(sDate)
                LogTaskChange CStr(rowTask.Range.Cells(1, idxTaskID).Value), "EndDate", eDate, sDate
            Else
                UserError "Task " & rowTask.Range.Cells(1, idxTaskID).Value & ": StartDate is after EndDate."
            End If
        End If
        rowTask.Range.Cells(1, idxDuration).Formula = "=[@EndDate]-[@StartDate]+1"
    End If

    If Trim$(CStr(rowTask.Range.Cells(1, idxDepends).Value)) <> "" Then
        deps = Split(CStr(rowTask.Range.Cells(1, idxDepends).Value), ",")
        For Each dep In deps
            dep = Trim$(CStr(dep))
            If Not taskMap.Exists(dep) Then
                UserError "Task " & rowTask.Range.Cells(1, idxTaskID).Value & ": dependency " & dep & " not found."
            Else
                depTaskEnd = tblTasks.ListRows(CLng(taskMap(dep))).Range.Cells(1, idxEnd).Value
                If IsDate(depTaskEnd) And IsDate(sDate) And CDate(sDate) < CDate(depTaskEnd) Then
                    If autoFixDates Then
                        rowTask.Range.Cells(1, idxStart).Value = CDate(depTaskEnd)
                        LogTaskChange CStr(rowTask.Range.Cells(1, idxTaskID).Value), "StartDate", sDate, depTaskEnd
                    Else
                        UserError "Task " & rowTask.Range.Cells(1, idxTaskID).Value & " starts before predecessor " & dep & " ends."
                    End If
                End If
            End If
        Next dep
    End If
End Sub

Public Function ResolveRAG(ByVal rowTask As ListRow, ByVal tblTasks As ListObject) As String
    Dim manualRAG As String, statusVal As String, endDate As Variant
    Dim amberDays As Long

    manualRAG = Trim$(CStr(rowTask.Range.Cells(1, GetColumnIndex(tblTasks, "RAG")).Value))
    statusVal = UCase$(CStr(rowTask.Range.Cells(1, GetColumnIndex(tblTasks, "Status")).Value))
    endDate = rowTask.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value
    amberDays = CLng(GetNamedRangeValue(NM_RAG_BUSINESS_DAYS, 5))

    If manualRAG <> "" And UCase$(manualRAG) <> "AUTO" Then
        ResolveRAG = manualRAG
    ElseIf statusVal = "DONE" Then
        ResolveRAG = "Green"
    ElseIf IsDate(endDate) And CDate(endDate) < Date Then
        ResolveRAG = "Red"
    ElseIf IsDate(endDate) And BusinessDaysFromToday(CDate(endDate)) <= amberDays Then
        ResolveRAG = "Amber"
    Else
        ResolveRAG = "Green"
    End If
End Function

Public Sub FixDependencyDates()
    ValidateTasks True
End Sub

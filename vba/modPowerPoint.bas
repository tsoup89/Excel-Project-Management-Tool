Attribute VB_Name = "modPowerPoint"
Option Explicit

Public Sub CreateStatusPowerPoint()
    On Error GoTo CleanFail

    Dim ppApp As Object, ppPres As Object
    Dim templatePath As String

    Set ppApp = SafePowerPointApp()
    If ppApp Is Nothing Then
        UserError "PowerPoint is unavailable."
        Exit Sub
    End If

    templatePath = CStr(GetNamedRangeValue(NM_TEMPLATE_PPT, ""))
    If templatePath <> "" And Dir$(templatePath) <> "" Then
        Set ppPres = ppApp.Presentations.Open(templatePath, , , True)
    Else
        Set ppPres = ppApp.Presentations.Add
    End If

    ppApp.Visible = True

    AddExecutiveSummarySlide ppPres
    AddProjectSlides ppPres
    AddChangeSlide ppPres

    Dim exportPath As String
    exportPath = ThisWorkbook.Path & "\StatusDeck_" & Format$(Now, "yyyymmdd_hhnn") & ".pptx"
    ppPres.SaveAs exportPath

    ThisWorkbook.Names(NM_MYUPDATES_LASTRUN).RefersToRange.Value = Now
    UserMessage "PowerPoint deck created: " & exportPath
    Exit Sub

CleanFail:
    UserError "PowerPoint export failed: " & Err.Description
End Sub

Private Sub AddExecutiveSummarySlide(ByVal ppPres As Object)
    Dim sld As Object, tblTasks As ListObject
    Dim redCount As Long, amberCount As Long, greenCount As Long, overdue As Long, blocked As Long
    Dim r As ListRow

    Set sld = ppPres.Slides.Add(ppPres.Slides.Count + 1, 1)
    sld.Shapes(1).TextFrame.TextRange.Text = "Portfolio Executive Summary"

    Set tblTasks = GetTable(SH_TASKS, TBL_TASKS)
    For Each r In tblTasks.ListRows
        Select Case UCase$(CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "RAG")).Value))
            Case "RED": redCount = redCount + 1
            Case "AMBER": amberCount = amberCount + 1
            Case Else: greenCount = greenCount + 1
        End Select

        If IsDate(r.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value) Then
            If CDate(r.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value) < Date And _
               UCase$(CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "Status")).Value)) <> "DONE" Then
                overdue = overdue + 1
            End If
        End If

        If UCase$(CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "Status")).Value)) = "BLOCKED" Then blocked = blocked + 1
    Next r

    sld.Shapes(2).TextFrame.TextRange.Text = _
        "Red: " & redCount & vbCrLf & _
        "Amber: " & amberCount & vbCrLf & _
        "Green: " & greenCount & vbCrLf & _
        "Overdue: " & overdue & vbCrLf & _
        "Blocked: " & blocked & vbCrLf & _
        "As of: " & Format$(Now, "yyyy-mm-dd hh:nn")
End Sub

Private Sub AddProjectSlides(ByVal ppPres As Object)
    Dim tblProjects As ListObject
    Dim p As ListRow

    Set tblProjects = GetTable(SH_PROJECTS, TBL_PROJECTS)
    For Each p In tblProjects.ListRows
        If UCase$(CStr(p.Range.Cells(1, GetColumnIndex(tblProjects, "Status")).Value)) <> "ON HOLD" And _
           UCase$(CStr(p.Range.Cells(1, GetColumnIndex(tblProjects, "Status")).Value)) <> "DONE" Then
            AddOneProjectSlide ppPres, CStr(p.Range.Cells(1, GetColumnIndex(tblProjects, "ProjectID")).Value), _
                               CStr(p.Range.Cells(1, GetColumnIndex(tblProjects, "ProjectName")).Value)
        End If
    Next p
End Sub

Private Sub AddOneProjectSlide(ByVal ppPres As Object, ByVal projectID As String, ByVal projectName As String)
    Dim sld As Object
    Set sld = ppPres.Slides.Add(ppPres.Slides.Count + 1, 2)
    sld.Shapes(1).TextFrame.TextRange.Text = projectID & " - " & projectName
    sld.Shapes(2).TextFrame.TextRange.Text = BuildProjectSlideBody(projectID)
End Sub

Private Function BuildProjectSlideBody(ByVal projectID As String) As String
    Dim tblTasks As ListObject
    Dim r As ListRow
    Dim overdueText As String, upcomingText As String, blockedText As String
    Dim overdueCount As Long, upcomingCount As Long, blockedCount As Long

    Set tblTasks = GetTable(SH_TASKS, TBL_TASKS)

    For Each r In tblTasks.ListRows
        If CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "ProjectID")).Value) = projectID Then
            If IsTaskOverdue(r, tblTasks) And overdueCount < 8 Then
                overdueCount = overdueCount + 1
                overdueText = overdueText & "• " & r.Range.Cells(1, GetColumnIndex(tblTasks, "TaskName")).Value & vbCrLf
            ElseIf IsTaskDueSoon(r, tblTasks, 14) And upcomingCount < 8 Then
                upcomingCount = upcomingCount + 1
                upcomingText = upcomingText & "• " & r.Range.Cells(1, GetColumnIndex(tblTasks, "TaskName")).Value & vbCrLf
            ElseIf UCase$(CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "Status")).Value)) = "BLOCKED" And blockedCount < 8 Then
                blockedCount = blockedCount + 1
                blockedText = blockedText & "• " & r.Range.Cells(1, GetColumnIndex(tblTasks, "TaskName")).Value & vbCrLf
            End If
        End If
    Next r

    BuildProjectSlideBody = "Overdue:" & vbCrLf & Nz(overdueText, "None") & vbCrLf & _
                            "Due Next 14 Days:" & vbCrLf & Nz(upcomingText, "None") & vbCrLf & _
                            "Blocked:" & vbCrLf & Nz(blockedText, "None")
End Function

Private Function IsTaskOverdue(ByVal r As ListRow, ByVal tblTasks As ListObject) As Boolean
    If IsDate(r.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value) Then
        IsTaskOverdue = CDate(r.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value) < Date And _
                        UCase$(CStr(r.Range.Cells(1, GetColumnIndex(tblTasks, "Status")).Value)) <> "DONE"
    End If
End Function

Private Function IsTaskDueSoon(ByVal r As ListRow, ByVal tblTasks As ListObject, ByVal withinDays As Long) As Boolean
    If IsDate(r.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value) Then
        IsTaskDueSoon = CDate(r.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value) >= Date And _
                        CDate(r.Range.Cells(1, GetColumnIndex(tblTasks, "EndDate")).Value) <= Date + withinDays
    End If
End Function

Private Sub AddChangeSlide(ByVal ppPres As Object)
    Dim sld As Object, tbl As ListObject, lastRun As Variant
    Dim t As String
    Dim r As ListRow

    Set sld = ppPres.Slides.Add(ppPres.Slides.Count + 1, 2)
    sld.Shapes(1).TextFrame.TextRange.Text = "Changes Since Last Report"

    Set tbl = GetTable(SH_LOGS, TBL_CHANGELOG)
    lastRun = GetNamedRangeValue(NM_MYUPDATES_LASTRUN, Date - 7)

    For Each r In tbl.ListRows
        If IsDate(r.Range.Cells(1, GetColumnIndex(tbl, "Timestamp")).Value) Then
            If CDate(r.Range.Cells(1, GetColumnIndex(tbl, "Timestamp")).Value) > CDate(lastRun) Then
                t = t & "• " & r.Range.Cells(1, GetColumnIndex(tbl, "TaskID")).Value & " " & _
                        r.Range.Cells(1, GetColumnIndex(tbl, "FieldName")).Value & ": " & _
                        r.Range.Cells(1, GetColumnIndex(tbl, "OldValue")).Value & " → " & _
                        r.Range.Cells(1, GetColumnIndex(tbl, "NewValue")).Value & vbCrLf
            End If
        End If
    Next r

    If t = "" Then t = "No tracked changes since last run."
    sld.Shapes(2).TextFrame.TextRange.Text = t
End Sub

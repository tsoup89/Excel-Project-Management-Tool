Attribute VB_Name = "modSetup"
Option Explicit

Public Sub InitializeWorkbook()
    On Error GoTo CleanFail
    ApplyStatusValidation
    ProtectPresentationSheets
    UserMessage "Initialization complete."
    Exit Sub
CleanFail:
    UserError "Initialization failed: " & Err.Description
End Sub

Public Sub ApplyStatusValidation()
    Dim tblTasks As ListObject
    Set tblTasks = GetTable(SH_TASKS, TBL_TASKS)

    With tblTasks.ListColumns("Status").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Not Started,In Progress,Blocked,Done,On Hold"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorMessage = "Allowed: Not Started, In Progress, Blocked, Done, On Hold"
    End With
End Sub

Public Sub ProtectPresentationSheets()
    Dim s As Variant
    For Each s In Array(SH_HOME, SH_GANTT, SH_DASHBOARD, SH_HELP, SH_LOGS)
        With GetWorksheet(CStr(s))
            .Protect Password:="upmtool", UserInterfaceOnly:=True
        End With
    Next s
End Sub

Public Sub SetBaselineForProject()
    On Error GoTo CleanFail

    Dim projectID As String
    Dim tblProjects As ListObject, r As ListRow

    projectID = InputBox("Enter ProjectID")
    If projectID = "" Then Exit Sub

    Set tblProjects = GetTable(SH_PROJECTS, TBL_PROJECTS)
    For Each r In tblProjects.ListRows
        If CStr(r.Range.Cells(1, GetColumnIndex(tblProjects, "ProjectID")).Value) = projectID Then
            r.Range.Cells(1, GetColumnIndex(tblProjects, "BaselineStart")).Value = r.Range.Cells(1, GetColumnIndex(tblProjects, "StartDate")).Value
            r.Range.Cells(1, GetColumnIndex(tblProjects, "BaselineEnd")).Value = r.Range.Cells(1, GetColumnIndex(tblProjects, "EndDate")).Value
            UserMessage "Baseline set for project " & projectID
            Exit Sub
        End If
    Next r

    UserError "ProjectID not found: " & projectID
    Exit Sub

CleanFail:
    UserError "Could not set project baseline: " & Err.Description
End Sub

Public Sub SetBaselineForAll()
    On Error GoTo CleanFail

    Dim tblProjects As ListObject, r As ListRow
    Set tblProjects = GetTable(SH_PROJECTS, TBL_PROJECTS)

    For Each r In tblProjects.ListRows
        r.Range.Cells(1, GetColumnIndex(tblProjects, "BaselineStart")).Value = r.Range.Cells(1, GetColumnIndex(tblProjects, "StartDate")).Value
        r.Range.Cells(1, GetColumnIndex(tblProjects, "BaselineEnd")).Value = r.Range.Cells(1, GetColumnIndex(tblProjects, "EndDate")).Value
    Next r

    UserMessage "Baseline set for all projects."
    Exit Sub

CleanFail:
    UserError "Could not set baseline for all: " & Err.Description
End Sub

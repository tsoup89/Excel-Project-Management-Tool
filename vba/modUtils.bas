Attribute VB_Name = "modUtils"
Option Explicit

Public Function GetWorksheet(ByVal sheetName As String) As Worksheet
    On Error GoTo CleanFail
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    Exit Function
CleanFail:
    Err.Raise vbObjectError + 100, "GetWorksheet", "Missing sheet: " & sheetName
End Function

Public Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Set ws = GetWorksheet(sheetName)

    On Error GoTo CleanFail
    Set GetTable = ws.ListObjects(tableName)
    Exit Function
CleanFail:
    Err.Raise vbObjectError + 101, "GetTable", "Missing table " & tableName & " on " & sheetName
End Function

Public Function GetNamedRangeValue(ByVal nameText As String, Optional ByVal defaultValue As Variant) As Variant
    On Error GoTo Fallback
    GetNamedRangeValue = ThisWorkbook.Names(nameText).RefersToRange.Value
    Exit Function
Fallback:
    GetNamedRangeValue = defaultValue
End Function

Public Function Nz(ByVal valueIn As Variant, Optional ByVal fallback As Variant = "") As Variant
    If IsError(valueIn) Or IsNull(valueIn) Or valueIn = "" Then
        Nz = fallback
    Else
        Nz = valueIn
    End If
End Function

Public Function BusinessDaysFromToday(ByVal targetDate As Date) As Long
    Dim weekendPattern As String
    weekendPattern = CStr(GetNamedRangeValue(NM_WEEKEND_PATTERN, "0000011"))
    BusinessDaysFromToday = WorksheetFunction.NetworkDays_Intl(Date, targetDate, weekendPattern)
End Function

Public Function IsTaskNeedsUpdate(ByVal statusValue As String, ByVal lastUpdated As Variant, ByVal updateNeededBy As Variant) As Boolean
    Dim thresholdDate As Date
    thresholdDate = Date - 7

    If UCase$(statusValue) = "DONE" Then
        IsTaskNeedsUpdate = False
        Exit Function
    End If

    If UCase$(statusValue) = "BLOCKED" Then
        IsTaskNeedsUpdate = True
        Exit Function
    End If

    If Trim$(CStr(lastUpdated)) = "" Then
        IsTaskNeedsUpdate = True
        Exit Function
    End If

    If CDate(lastUpdated) <= thresholdDate Then
        IsTaskNeedsUpdate = True
        Exit Function
    End If

    If IsDate(updateNeededBy) Then
        If BusinessDaysFromToday(CDate(updateNeededBy)) <= 2 Then
            IsTaskNeedsUpdate = True
            Exit Function
        End If
    End If

    IsTaskNeedsUpdate = False
End Function

Public Sub UserMessage(ByVal msg As String, Optional ByVal title As String = "Ultimate PM Tool")
    MsgBox msg, vbInformation, title
End Sub

Public Sub UserError(ByVal msg As String, Optional ByVal title As String = "Ultimate PM Tool")
    MsgBox msg, vbExclamation, title
End Sub

Public Function SafeOutlookApp() As Object
    On Error Resume Next
    Set SafeOutlookApp = GetObject(, "Outlook.Application")
    If SafeOutlookApp Is Nothing Then Set SafeOutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
End Function

Public Function SafePowerPointApp() As Object
    On Error Resume Next
    Set SafePowerPointApp = GetObject(, "PowerPoint.Application")
    If SafePowerPointApp Is Nothing Then Set SafePowerPointApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0
End Function

Public Function GetColumnIndex(ByVal tbl As ListObject, ByVal columnName As String) As Long
    GetColumnIndex = tbl.ListColumns(columnName).Index
End Function

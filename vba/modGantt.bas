Attribute VB_Name = "modGantt"
Option Explicit

Public Sub RefreshGantt()
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = GetWorksheet(SH_GANTT)

    ws.Calculate
    DrawTodayLine ws
    UserMessage "Gantt refreshed. Use slicers on tblTasks for Project/Owner/Status/RAG."
    Exit Sub

CleanFail:
    UserError "Unable to refresh Gantt: " & Err.Description
End Sub

Private Sub DrawTodayLine(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Shapes("shTodayLine").Delete
    On Error GoTo 0

    Dim chartObj As ChartObject
    If ws.ChartObjects.Count = 0 Then Exit Sub
    Set chartObj = ws.ChartObjects(1)

    Dim xPos As Double
    xPos = chartObj.Left + chartObj.Width * 0.5

    Dim ln As Shape
    Set ln = ws.Shapes.AddLine(xPos, chartObj.Top, xPos, chartObj.Top + chartObj.Height)
    ln.Name = "shTodayLine"
    ln.Line.ForeColor.RGB = RGB(255, 0, 0)
    ln.Line.Weight = 1.5
End Sub

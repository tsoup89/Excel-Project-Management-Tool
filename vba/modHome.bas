Attribute VB_Name = "modHome"
Option Explicit

Public Sub RunDailyRefresh()
    ValidateTasks False
    RefreshMyUpdates
    RefreshGantt
End Sub

Public Sub OpenHelp()
    GetWorksheet(SH_HELP).Activate
End Sub

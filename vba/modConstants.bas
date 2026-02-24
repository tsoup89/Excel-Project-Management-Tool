Attribute VB_Name = "modConstants"
Option Explicit

Public Const SH_HOME As String = "Home"
Public Const SH_PROJECTS As String = "Projects"
Public Const SH_TASKS As String = "Tasks"
Public Const SH_PEOPLE As String = "People"
Public Const SH_MYUPDATES As String = "MyUpdates"
Public Const SH_GANTT As String = "Gantt"
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_SETUP As String = "Setup"
Public Const SH_LOGS As String = "Logs"
Public Const SH_HELP As String = "Help"

Public Const TBL_PROJECTS As String = "tblProjects"
Public Const TBL_TASKS As String = "tblTasks"
Public Const TBL_PEOPLE As String = "tblPeople"
Public Const TBL_EMAILLOG As String = "tblEmailLog"
Public Const TBL_CHANGELOG As String = "tblChangeLog"
Public Const TBL_MYUPDATES As String = "tblMyUpdates"

Public Const NM_MYUPDATES_OWNER As String = "nmMyUpdatesOwner"
Public Const NM_MYUPDATES_LASTRUN As String = "nmLastPptRun"
Public Const NM_TEMPLATE_PPT As String = "nmPptTemplatePath"
Public Const NM_EMAIL_MODE As String = "nmEmailMode"
Public Const NM_RAG_BUSINESS_DAYS As String = "nmRAGAmberBusinessDays"
Public Const NM_WEEKEND_PATTERN As String = "nmWeekendPattern"

Public Enum UpdateAction
    uaGeneralUpdate = 1
    uaMarkDone = 2
    uaSetBlocked = 3
    uaAddComment = 4
End Enum

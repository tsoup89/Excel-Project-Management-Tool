# Ultimate Project Management Tool (.xlsm) - Build & Run Guide

## 1) Workbook structure
Create sheets (exact names):
- Home, Projects, Tasks, People, MyUpdates, Gantt, Dashboard, Setup, Logs, Help

Create Excel tables (ListObjects):
- `tblProjects` on Projects
  - ProjectID, ProjectName, Sponsor, PM, Status, Priority, StartDate, EndDate, BaselineStart, BaselineEnd, PercentComplete, RAG, Notes
- `tblTasks` on Tasks
  - TaskID, ProjectID, WBS, TaskName, OwnerName, OwnerEmail, StartDate, EndDate, Duration, Dependencies, Milestone, PercentComplete, Status, RAG, LastUpdated, UpdateNeededBy, Comments, Tags
- `tblPeople` on People
  - OwnerName, Email, Role, Team, Timezone
- `tblMyUpdates` on MyUpdates
  - Same columns as `tblTasks` plus `NeedsUpdate`
- `tblEmailLog` on Logs
  - Timestamp, OwnerName, OwnerEmail, Subject, TaskIDs, Mode
- `tblChangeLog` on Logs
  - Timestamp, UserName, TaskID, FieldName, OldValue, NewValue

## 2) Named ranges (Setup sheet)
Define the following names:
- `nmMyUpdatesOwner` (selected owner dropdown cell)
- `nmLastPptRun` (single cell datetime)
- `nmPptTemplatePath` (optional pptx template file path)
- `nmEmailMode` (optional: Preview/Send)
- `nmRAGAmberBusinessDays` (default 5)
- `nmWeekendPattern` (default `0000011` for Sat/Sun)

## 3) Import VBA
1. Save workbook as `.xlsm`.
2. Open VBA editor (ALT+F11).
3. Import all files from `/vba`:
   - modConstants.bas
   - modUtils.bas
   - modValidation.bas
   - modLogs.bas
   - modMyUpdates.bas
   - modSetup.bas
   - modEmail.bas
   - modPowerPoint.bas
   - modGantt.bas
   - modSync.bas
   - modHome.bas
4. Replace `ThisWorkbook` code with `ThisWorkbook.cls` contents.

## 4) Home buttons (assign macros)
- Refresh / Daily Refresh -> `RunDailyRefresh`
- Set Baseline for Project -> `SetBaselineForProject`
- Set Baseline for All -> `SetBaselineForAll`
- Send Update Requests (Preview) -> `SendUpdateRequestsPreview`
- Send Update Requests (Send) -> `SendUpdateRequestsSend`
- Create Status PowerPoint -> `CreateStatusPowerPoint`
- Fix Dependency Dates -> `FixDependencyDates`
- Help -> `OpenHelp`

MyUpdates buttons:
- Refresh MyUpdates -> `RefreshMyUpdates`
- Update Selected Task -> `UpdateSelectedTask`
- Mark Done -> `MarkDone`
- Set Blocked -> `SetBlocked`
- Add Comment -> `AddComment`
- (Optional) Export CSV -> `ExportMyUpdatesToCSV`
- (Optional) Import CSV -> `ImportUpdatesFromCSV`

## 5) Gantt implementation notes
- Build Gantt chart from `tblTasks` StartDate/Duration and optional baseline columns from project-level baseline.
- Add slicers on `tblTasks`: ProjectID, OwnerName, Status, RAG.
- Date window (30/60/90/custom) should be controlled by report filters or named cells used by formulas.
- `RefreshGantt` redraws the Today marker line.

## 6) Outlook and PowerPoint prerequisites
- Desktop Outlook installed and signed in (for COM email automation).
- Desktop PowerPoint installed (for deck export).
- Tool gracefully warns when either COM app is unavailable.

## 7) Help tab suggestions
Document:
- Owner update flow in-book only.
- No email reply processing; email is notification-only.
- Troubleshooting for protected sheets, missing COM apps, and macro security.

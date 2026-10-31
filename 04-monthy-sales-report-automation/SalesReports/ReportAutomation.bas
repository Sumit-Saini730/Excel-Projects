Attribute VB_Name = "ReportAutomation"
Option Explicit

' =====================================================================
' Monthly Sales Report Automation - all macros
'
' Setup needed in the workbook:
'   * Save as .xlsm
'   * A sheet named "Dashboard"
'   * A named range "DashboardArea" covering the whole dashboard panel
'   * (optional) A named cell "LastUpdated" for the "Date Last Updated" label
'   * Icons named: btnRefresh, btnFullScreen, btnPDF (Selection Pane, Alt+F10)
'
' Assign macros: right-click an icon > Assign Macro
'   btnRefresh    -> RefreshDashboard
'   btnFullScreen -> ToggleFullScreen
'   btnPDF        -> ExportDashboardPDF
' =====================================================================

Private origZoom As Long

' ---------------------------------------------------------------------
' Shared helper: refresh every query and pivot, and wait until done
' ---------------------------------------------------------------------
Private Sub RefreshAllData()
    Dim cn As WorkbookConnection

    ' Make each query finish before the next step runs
    For Each cn In ThisWorkbook.Connections
        On Error Resume Next
        cn.OLEDBConnection.BackgroundQuery = False
        On Error GoTo 0
    Next cn

    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
End Sub

' ---------------------------------------------------------------------
' Show or hide the icons (used so they do not appear in the PDF)
' ---------------------------------------------------------------------
Private Sub SetIconsVisible(ws As Worksheet, show As Boolean)
    Dim nm As Variant
    For Each nm In Array("btnRefresh", "btnFullScreen", "btnPDF")
        On Error Resume Next
        ws.Shapes(nm).Visible = show
        On Error GoTo 0
    Next nm
End Sub

' ---------------------------------------------------------------------
' 1. Refresh icon: refresh all data and stamp the last-updated date
' ---------------------------------------------------------------------
Sub RefreshDashboard()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    RefreshAllData

    ' Stamp the "Date Last Updated On" cell (named range LastUpdated)
    On Error Resume Next
    ThisWorkbook.Names("LastUpdated").RefersToRange.Value = Date
    On Error GoTo ErrHandler

    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed.", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Refresh failed: " & Err.Description, vbExclamation
End Sub

' ---------------------------------------------------------------------
' 2. Full-screen icon: click once for full screen, click again to go back
' ---------------------------------------------------------------------
Sub ToggleFullScreen()
    Dim ws As Worksheet

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Activate

    If Application.DisplayFullScreen = False Then
        origZoom = ActiveWindow.Zoom
        Application.DisplayFullScreen = True
        Application.DisplayFormulaBar = False
        With ActiveWindow
            .DisplayGridlines = False
            .DisplayHeadings = False
            .DisplayWorkbookTabs = False
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
        End With
        ws.Range("DashboardArea").Select
        ActiveWindow.Zoom = True          ' fit the selected area to the screen
        ws.Range("A1").Select
    Else
        Application.DisplayFullScreen = False
        Application.DisplayFormulaBar = True
        With ActiveWindow
            .DisplayGridlines = True
            .DisplayHeadings = True
            .DisplayWorkbookTabs = True
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            If origZoom > 0 Then .Zoom = origZoom Else .Zoom = 100
        End With
    End If
    Exit Sub

ErrHandler:
    MsgBox "Full screen failed: " & Err.Description, vbExclamation
End Sub

' ---------------------------------------------------------------------
' 3. PDF icon: refresh, then export the dashboard as a one-page PDF
' ---------------------------------------------------------------------
Sub ExportDashboardPDF()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim filePath As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("Dashboard")

    ' Refresh first so the PDF always has the latest data
    RefreshAllData

    ' Where to save (falls back to Documents if the file is on OneDrive)
    folderPath = ThisWorkbook.Path
    If folderPath = "" Or Left(folderPath, 4) = "http" Then
        folderPath = Environ("USERPROFILE") & "\Documents"
    End If
    filePath = folderPath & "\Sales_Dashboard_" & Format(Now, "yyyy-mm-dd_hhmm") & ".pdf"

    ' Page setup: dashboard area only, one landscape A4 page
    With ws.PageSetup
        .PrintArea = ThisWorkbook.Names("DashboardArea").RefersToRange.Address
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
    End With

    ' Hide the icons, export, then show them again
    SetIconsVisible ws, False
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    SetIconsVisible ws, True

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    If Not ws Is Nothing Then SetIconsVisible ws, True
    Application.ScreenUpdating = True
    MsgBox "PDF export failed: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' OPTIONAL (goes in the "ThisWorkbook" module, not in this module):
' restores the normal view if the file is closed while in full screen.
'
' Private Sub Workbook_BeforeClose(Cancel As Boolean)
'     Application.DisplayFullScreen = False
'     Application.DisplayFormulaBar = True
' End Sub
' =====================================================================

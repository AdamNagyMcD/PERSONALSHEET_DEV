Attribute VB_Name = "mod_PerformanceBaseline"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Private Const PID_PERF_LOG_SHEET As String = "PID_PERFORMANCE_LOG"
Private Const PID_PERF_BENCH_MONTH As String = "Februar"
Private Const PID_PERF_BENCH_ROW As Long = 3


Public Sub RunPerformanceBaseline()
    PID_RunPerformanceBaseline
End Sub


Public Sub PID_RunPerformanceBaseline()
    Dim report As String
    Dim logWs As Worksheet
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    
    If Not PID_ConfirmAdminAction( _
        "Misst nicht-destruktive Performance-Schritte (FP-010 / TR-05)." & vbCrLf & _
        "CopyData und Cold Open: manuell laut docs/PERFORMANCE_BASELINE.md", _
        "Performance Baseline") Then
        Exit Sub
    End If
    
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    
    report = PID_PerfBuildEnvironmentHeader()
    report = report & PID_PerfMeasureKvMonthTabRefreshScoped()
    report = report & PID_PerfMeasureKvMonthTabRefresh()
    report = report & PID_PerfMeasureMonatslohnRecalc()
    report = report & PID_PerfMeasureFinanzSync()
    report = report & PID_PerfMeasureFluktuationSaveRefresh()
    report = report & PID_PerfMeasureFluktuationTabRefresh()
    report = report & PID_PerfMeasureFullSystemRefresh()
    report = report & PID_PerfManualStepsFooter()
    
    Set logWs = PID_PerfAppendLogReport(report)
    
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation
    
    MsgBox report & vbCrLf & vbCrLf & _
           "Log: " & IIf(logWs Is Nothing, "(nicht geschrieben)", PID_PERF_LOG_SHEET), _
           vbInformation, "FP-010 Performance Baseline"
End Sub


Private Function PID_PerfBuildEnvironmentHeader() As String
    Dim header As String
    
    header = "FP-010 Performance Baseline" & vbCrLf
    header = header & String$(40, "-") & vbCrLf
    header = header & "Datum: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    header = header & "Workbook: " & ThisWorkbook.Name & vbCrLf
    header = header & "Excel: " & Application.Version & vbCrLf
    header = header & "OS: " & Application.OperatingSystem & vbCrLf
    header = header & "Calculation: " & PID_GetCalculationModeText() & vbCrLf
    header = header & String$(40, "-") & vbCrLf & vbCrLf
    
    PID_PerfBuildEnvironmentHeader = header
End Function


Private Function PID_PerfMeasureKvMonthTabRefresh() As String
    Dim ws As Worksheet
    Dim started As Single
    Dim elapsed As Single
    
    On Error GoTo SafeExit
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_PERF_BENCH_MONTH)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then
        PID_PerfMeasureKvMonthTabRefresh = "2 KV-Refresh " & PID_PERF_BENCH_MONTH & ": SKIP (Blatt fehlt)" & vbCrLf
        Exit Function
    End If
    
    MarkAllKVDropdownsDirty
    
    started = Timer
    RefreshKVDropdownsIfDirtyForSheet ws
    ws.Activate
    elapsed = Timer - started
    
    PID_PerfMeasureKvMonthTabRefresh = _
        "2 KV dirty -> Refresh F-Dropdown (" & PID_PERF_BENCH_MONTH & "): " & _
        PID_PerfFormatSeconds(elapsed) & vbCrLf

SafeExit:
End Function


Private Function PID_PerfMeasureKvMonthTabRefreshScoped() As String
    Dim ws As Worksheet
    Dim started As Single
    Dim elapsed As Single
    Const PID_PERF_BENCH_KV_CODE As String = "BG1"
    
    On Error GoTo SafeExit
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_PERF_BENCH_MONTH)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then
        PID_PerfMeasureKvMonthTabRefreshScoped = _
            "2b KV scoped dirty -> Refresh F (" & PID_PERF_BENCH_MONTH & ", " & PID_PERF_BENCH_KV_CODE & "): SKIP (Blatt fehlt)" & vbCrLf
        Exit Function
    End If
    
    MarkKVDropdownDirtyForKVCode PID_PERF_BENCH_KV_CODE
    
    started = Timer
    RefreshKVStundenDropdownForSheet ws
    ws.Activate
    elapsed = Timer - started
    
    PID_PerfMeasureKvMonthTabRefreshScoped = _
        "2b KV scoped dirty -> Refresh F (" & PID_PERF_BENCH_MONTH & ", " & PID_PERF_BENCH_KV_CODE & "): " & _
        PID_PerfFormatSeconds(elapsed) & vbCrLf

SafeExit:
End Function


Private Function PID_PerfMeasureMonatslohnRecalc() As String
    Dim ws As Worksheet
    Dim started As Single
    Dim elapsed As Single
    Dim benchCell As Range
    
    On Error GoTo SafeExit
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Januar")
    On Error GoTo SafeExit
    
    If ws Is Nothing Then
        PID_PerfMeasureMonatslohnRecalc = "3 G-Recalc (Januar): SKIP (Blatt fehlt)" & vbCrLf
        Exit Function
    End If
    
    Set benchCell = ws.Range("E" & PID_PERF_BENCH_ROW)
    
    started = Timer
    PID_RecalculateMonatslohnForChangedRows ws, benchCell
    benchCell.Calculate
    ws.Range("G" & PID_PERF_BENCH_ROW).Calculate
    elapsed = Timer - started
    
    PID_PerfMeasureMonatslohnRecalc = _
        "3 E-Zeile -> G-Recalc (Januar Z" & PID_PERF_BENCH_ROW & "): " & _
        PID_PerfFormatSeconds(elapsed) & vbCrLf

SafeExit:
End Function


Private Function PID_PerfMeasureFinanzSync() As String
    Dim started As Single
    Dim elapsed As Single
    
    On Error GoTo SafeExit
    
    started = Timer
    PID_SyncFinanzSummaryToUbersicht
    elapsed = Timer - started
    
    PID_PerfMeasureFinanzSync = _
        "5 FINANZ-Sync -> UEBERSICHT: " & PID_PerfFormatSeconds(elapsed) & vbCrLf

SafeExit:
End Function


Private Function PID_PerfMeasureFluktuationSaveRefresh() As String
    Dim started As Single
    Dim elapsed As Single
    
    On Error GoTo SafeExit
    
    MarkFluktuationDirtyForMonth 2
    
    started = Timer
    RefreshFluktuationDataIfDirty
    elapsed = Timer - started
    
    PID_PerfMeasureFluktuationSaveRefresh = _
        "6 Save-Pfad: Fluktuation Daten only (Februar dirty): " & PID_PerfFormatSeconds(elapsed) & vbCrLf

SafeExit:
End Function


Private Function PID_PerfMeasureFluktuationTabRefresh() As String
    Dim started As Single
    Dim elapsed As Single
    
    On Error GoTo SafeExit
    
    MarkFluktuationDirtyForMonth 2
    
    started = Timer
    RefreshFluktuationIfDirty
    elapsed = Timer - started
    
    PID_PerfMeasureFluktuationTabRefresh = _
        "6b Tab-Pfad: Fluktuation Daten+Analyse (Februar dirty): " & PID_PerfFormatSeconds(elapsed) & vbCrLf

SafeExit:
End Function


Private Function PID_PerfMeasureFullSystemRefresh() As String
    Dim started As Single
    Dim elapsed As Single
    
    On Error GoTo SafeExit
    
    started = Timer
    PID_PerfRunFullSystemRefreshTimed
    elapsed = Timer - started
    
    PID_PerfMeasureFullSystemRefresh = _
        "7 FullSystemRefresh (ohne Dialog): " & PID_PerfFormatSeconds(elapsed) & vbCrLf

SafeExit:
End Function


Private Sub PID_PerfRunFullSystemRefreshTimed()
    On Error GoTo SafeExit
    
    PID_SetupSheetProtectionForMacros
    RefreshAllMonthKVStundenDropdowns
    PID_RestoreMonatslohnFormulasSilent
    PID_RestoreAktuelleStundenFormulasSilent
    PID_RestoreUrlaubGeldFormulasSilent
    PID_RestoreLetztesGehaltFormulasSilent
    PID_RestoreKVCodeDropdownValidationSilent
    ClearAllKVLohnDirty
    RefreshFluktuationAll
    PID_RestoreFinanzSummaryOnUbersicht
    PID_RecalculateAllMonthMergedFormulas
    PID_FormatAllMoneyColumns
    PID_ApplyCopyrightToAllSheets
    PID_EnableCalculationForAllSheets
    
    On Error Resume Next
    Application.CalculateFull
    Err.Clear

SafeExit:
End Sub


Private Function PID_PerfManualStepsFooter() As String
    Dim footer As String
    
    footer = vbCrLf & "MANUELL (Stoppuhr, docs/PERFORMANCE_BASELINE.md — TR-05):" & vbCrLf
    footer = footer & "1 Cold Open -> Monats-Tab" & vbCrLf
    footer = footer & "2 LOHNTABELLE -> Eigene Stunden -> Monats-Tab (real)" & vbCrLf
    footer = footer & "4 CopyData Januar -> Dezember (Testkopie!)" & vbCrLf
    footer = footer & "6 Save mit Fluktuation dirty (Strg+S)" & vbCrLf
    footer = footer & vbCrLf & "Tester-Frage: Was f" & ChrW(252) & "hlt sich langsam an?" & vbCrLf
    footer = footer & "(Open / CopyData / KV / Fluktuation-Tab / Full Refresh)" & vbCrLf
    
    PID_PerfManualStepsFooter = footer
End Function


Private Function PID_PerfFormatSeconds(ByVal elapsed As Single) As String
    If elapsed < 0! Then elapsed = elapsed + 86400!
    PID_PerfFormatSeconds = Format$(elapsed, "0.00") & " s"
End Function


Private Function PID_PerfAppendLogReport(ByVal report As String) As Worksheet
    Dim ws As Worksheet
    Dim nextRow As Long
    
    On Error GoTo SafeExit
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_PERF_LOG_SHEET)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PID_PERF_LOG_SHEET
        ws.Range("A1").Value = "FP-010 Log"
        ws.Range("A2").Value = "Zeitstempel"
        ws.Range("B2").Value = "Report"
    End If
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 3 Then nextRow = 3
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = report
    
    ws.Visible = xlSheetVeryHidden
    Set PID_PerfAppendLogReport = ws

SafeExit:
End Function

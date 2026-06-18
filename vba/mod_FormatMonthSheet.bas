Attribute VB_Name = "mod_FormatMonthSheet"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Design reference: manually formatted sheet "Januar" in Personalsheet.xlsm (2026-05-24).
Private Const PID_MS_PILOT_SHEET As String = "Januar"
Private Const PID_MS_PANEL_FIRST_ROW As Long = 3
Private Const PID_MS_PANEL_LAST_ROW As Long = 50
Private Const PID_MS_PANEL_TAIL_RANGE As String = "W3:Y50"

Private Const PID_MS_STYLE_INPUT As Long = 1
Private Const PID_MS_STYLE_READONLY As Long = 2
Private Const PID_MS_STYLE_LABEL As Long = 3
Private Const PID_MS_STYLE_HEADER As Long = 4
Private Const PID_MS_STYLE_MESSAGE As Long = 5
Private Const PID_MS_COPYDATA_BUTTON_NAME As String = "btn_CopyDataMonth"
Private Const PID_MS_COPYDATA_BUTTON_TEXT As String = "Aktualisierung des restlichen Jahres"
Private Const PID_MS_COPYDATA_BUTTON_WIDTH As Double = 275
Private Const PID_MS_COPYDATA_BUTTON_HEIGHT As Double = 24
Private Const PID_MS_COPYDATA_BUTTON_OFFSET_LEFT As Double = 6
Private Const PID_MS_COPYDATA_BUTTON_OFFSET_TOP As Double = 4
' FP-017: dd.mm.yyyy in D (Eintritt) und I (Austritt) — breiter als Default (~11).
Private Const PID_MS_DATE_COLUMN_WIDTH As Double = 13
Private Const PID_MS_AUSTRITTSGRUND_COL_WIDTH As Double = 34
Private Const PID_MS_AUSTRITTSGRUND_MIN_ROW_HEIGHT As Single = 18
Private Const PID_MS_AUSTRITTSGRUND_MAX_ROW_HEIGHT As Single = 72
Private Const PID_MS_ROW_INDEX_COL_WIDTH As Double = 3.82
Private Const PID_MS_ROW_INDEX_NUMBER_FORMAT As String = "@"
Private Const PID_MS_EMPLOYEE_HEADER_FIRST_COL As Long = 1
Private Const PID_MS_EMPLOYEE_HEADER_LAST_COL As Long = 14
Private Const PID_MS_PANEL_MESSAGE_ROW As Long = 15


Public Sub PID_FixDurchrechnungStartMonthPanels()
    Dim monthName As Variant
    Dim ws As Worksheet
    Dim countDone As Long
    
    On Error GoTo CleanFail
    
    For Each monthName In PID_MS_DurchrechnungPeriodStartMonthNames()
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthName))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If PID_IsWorkerMonthSheet(ws) Then
                PID_MSApplyDurchrechnungStartMonthPanelRow15 ws
                countDone = countDone + 1
            End If
        End If
    Next monthName
    
    MsgBox countDone & " Monatsblaetter: Panel O15 als O15:R15 (wie Februar).", _
           vbInformation, "Durchrechnung Panel"
    Exit Sub

CleanFail:
    MsgBox "Fehler bei PID_FixDurchrechnungStartMonthPanels:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, "Durchrechnung Panel"
End Sub


Private Function PID_MS_DurchrechnungPeriodStartMonthNames() As Variant
    PID_MS_DurchrechnungPeriodStartMonthNames = Array("Februar", "Mai", "August", "November")
End Function


Private Function PID_IsDurchrechnungPeriodStartMonth(ByVal ws As Worksheet) As Boolean
    PID_IsDurchrechnungPeriodStartMonth = PID_IsDurchrechnungPeriodStartMonthSheet(ws)
End Function


Private Sub PID_MSApplyDurchrechnungStartMonthPanelRow15(ByVal ws As Worksheet)
    Dim msgRange As Range
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    On Error Resume Next
    ws.Range("O15:R15").UnMerge
    ws.Range("O15:P15").UnMerge
    ws.Range("Q15:R15").UnMerge
    Err.Clear
    On Error GoTo 0
    
    Set msgRange = ws.Range("O15:R15")
    msgRange.Merge
    PID_MSApplyStyleToRangeMergedOnce msgRange, PID_MS_STYLE_MESSAGE
End Sub


Public Sub EnsureMonthSheetCopyDataButtons()
    Dim monthName As Variant
    Dim ws As Worksheet
    Dim countDone As Long
    
    On Error GoTo CleanFail
    
    For Each monthName In PID_MonthNames()
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthName))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If PID_IsWorkerMonthSheet(ws) Then
                If PID_MSEnsureAktualisierungButtonOnSheet(ws) Then
                    countDone = countDone + 1
                End If
            End If
        End If
    Next monthName
    
    MsgBox countDone & " Monatsblaetter: CopyData-Button im LOHNTABELLE-Stil.", _
           vbInformation, "Monatsblatt Button"
    Exit Sub

CleanFail:
    MsgBox "Fehler bei EnsureMonthSheetCopyDataButtons:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, "Monatsblatt Button"
End Sub


Public Function PID_EnsureMonthSheetCopyDataButton(ByVal ws As Worksheet) As Boolean
    PID_EnsureMonthSheetCopyDataButton = PID_MSEnsureAktualisierungButtonOnSheet(ws)
End Function


' FP-017: Spalten D und I auf allen Monatsblaettern (FullSystemRefresh, Format-Lauf).
Public Sub PID_ApplyMonthSheetDateColumnWidths(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    ws.Columns("D").ColumnWidth = PID_MS_DATE_COLUMN_WIDTH
    ws.Columns("I").ColumnWidth = PID_MS_DATE_COLUMN_WIDTH
End Sub


' Spalte N (Austrittsgrund): Umbruch + Zeilenhoehe, damit lange Dropdown-Texte lesbar bleiben.
Public Sub PID_ApplyMonthSheetAustrittsgrundLayout(ByVal ws As Worksheet)
    Dim r As Long
    Dim dataRange As Range
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    ws.Columns("N").ColumnWidth = PID_MS_AUSTRITTSGRUND_COL_WIDTH
    
    Set dataRange = ws.Range("N3:N82")
    dataRange.WrapText = True
    dataRange.VerticalAlignment = xlTop
    
    For r = PID_FIRST_ROW To PID_LAST_ROW
        If Len(Trim$(CStr(ws.Cells(r, "N").Value))) > 0 Then
            ws.Rows(r).AutoFit
            If ws.Rows(r).RowHeight < PID_MS_AUSTRITTSGRUND_MIN_ROW_HEIGHT Then
                ws.Rows(r).RowHeight = PID_MS_AUSTRITTSGRUND_MIN_ROW_HEIGHT
            ElseIf ws.Rows(r).RowHeight > PID_MS_AUSTRITTSGRUND_MAX_ROW_HEIGHT Then
                ws.Rows(r).RowHeight = PID_MS_AUSTRITTSGRUND_MAX_ROW_HEIGHT
            End If
        End If
    Next r
    
SafeExit:
End Sub


Public Sub FormatJanuarMonthSheet()
    PID_FormatMonthSheetByName PID_MS_PILOT_SHEET
End Sub


Public Sub FormatAllMonthSheets()
    Dim wsRef As Worksheet
    Dim monthName As Variant
    Dim ws As Worksheet
    Dim countDone As Long
    
    On Error GoTo CleanFail
    
    Set wsRef = ThisWorkbook.Worksheets(PID_MS_PILOT_SHEET)
    
    If Not PID_IsWorkerMonthSheet(wsRef) Then
        MsgBox "Referenzblatt '" & PID_MS_PILOT_SHEET & "' ist kein gueltiges Monatsblatt.", _
               vbExclamation, "Monatsblatt Format"
        Exit Sub
    End If
    
    For Each monthName In PID_MonthNames()
        If StrComp(CStr(monthName), PID_MS_PILOT_SHEET, vbTextCompare) <> 0 Then
            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(CStr(monthName))
            On Error GoTo CleanFail
            
            If Not ws Is Nothing Then
                If PID_IsWorkerMonthSheet(ws) Then
                    PID_MSCopyMonthFormatsFromReference wsRef, ws
                    PID_ApplyMonthSheetEmployeeRowLayout ws
                    PID_MSEnsureAktualisierungButtonOnSheet ws
                    countDone = countDone + 1
                End If
            End If
        End If
    Next monthName
    
    PID_ApplyCopyrightToAllSheets
    
    MsgBox countDone & " Monatsblaetter von '" & PID_MS_PILOT_SHEET & "' formatiert.", _
           vbInformation, "Monatsblatt Format"
    Exit Sub

CleanFail:
    MsgBox "Fehler bei FormatAllMonthSheets:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, "Monatsblatt Format"
End Sub


Private Sub PID_FormatMonthSheetByName(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim wasProtected As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean
    
    On Error GoTo CleanFail
    
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo CleanFail
    
    If ws Is Nothing Then
        MsgBox "Monatsblatt '" & sheetName & "' wurde nicht gefunden.", vbExclamation, "Monatsblatt Format"
        GoTo CleanExit
    End If
    
    If Not PID_IsWorkerMonthSheet(ws) Then
        MsgBox "'" & sheetName & "' ist kein gueltiges Monatsblatt.", vbExclamation, "Monatsblatt Format"
        GoTo CleanExit
    End If
    
    wasProtected = ws.ProtectContents
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    PID_MSApplyReferenceLayout ws
    PID_MSEnsureAktualisierungButtonOnSheet ws
    PID_MSRemoveLegacyPanelShapes ws
    
    Application.DisplayAlerts = oldDisplayAlerts
    MsgBox "Monatsblatt '" & sheetName & "' formatiert.", vbInformation, "Monatsblatt Format"
    
    GoTo CleanExit

CleanFail:
    Application.DisplayAlerts = oldDisplayAlerts
    MsgBox "Fehler bei Monatsblatt-Format (" & sheetName & "):" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, "Monatsblatt Format"

CleanExit:
    On Error Resume Next
    If Not ws Is Nothing Then
        If wasProtected Then
            PID_ReprotectWorksheet ws
        End If
    End If
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Private Sub PID_MSCopyMonthFormatsFromReference(ByVal wsRef As Worksheet, ByVal wsTarget As Worksheet)
    Dim refProtected As Boolean
    Dim tgtProtected As Boolean
    
    If wsRef Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub
    
    refProtected = wsRef.ProtectContents
    tgtProtected = wsTarget.ProtectContents
    
    On Error Resume Next
    wsRef.Unprotect Password:=PID_WORKBOOK_PASSWORD
    wsTarget.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo 0
    
    PID_MSCopyMonthEmployeeHeaderFromReference wsRef, wsTarget
    
    ' Nur Datenzeilen: Zeile 1-2 sind bereits per Header-Kopie gesetzt (Merge!).
    ' E3:F82 auslassen: xlPasteFormats kopiert Locked und loescht Data Validation.
    PID_MSCopyRangeFormatsMergeSafe wsRef.Range("A3:A82"), wsTarget.Range("A3:A82")
    PID_MSCopyRangeFormatsMergeSafe wsRef.Range("B3:D82"), wsTarget.Range("B3:D82")
    PID_MSCopyRangeFormatsMergeSafe wsRef.Range("G3:N82"), wsTarget.Range("G3:N82")
    PID_MSApplyRightPanelReferenceStyles wsTarget
    PID_MSCopyPanelTailFormatsFromReference wsRef, wsTarget
    PID_MSRestoreMonthSheetDropdowns wsTarget
    PID_MSApplyEmployeeBlockStyles wsTarget
    PID_MSEnsureAktualisierungButtonOnSheet wsTarget
    PID_ApplyMonthSheetDateColumnWidths wsTarget
    
    On Error Resume Next
    If refProtected Then
        If PID_IsWorkerMonthSheet(wsRef) Then
            PID_ReprotectWorksheet wsRef
        Else
            wsRef.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, _
                         AllowFiltering:=True, AllowSorting:=True
        End If
    End If
    If tgtProtected Then
        PID_ReprotectWorksheet wsTarget
    End If
    On Error GoTo 0
End Sub


Private Sub PID_MSCopyRangeFormats(ByVal sourceRange As Range, ByVal targetRange As Range)
    If sourceRange Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub


Private Sub PID_MSCopyRangeFormatsMergeSafe(ByVal sourceRange As Range, ByVal targetRange As Range)
    If sourceRange Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    
    On Error Resume Next
    targetRange.UnMerge
    Err.Clear
    On Error GoTo 0
    
    PID_MSCopyRangeFormats sourceRange, targetRange
End Sub


Private Sub PID_MSCopyPanelTailFormatsFromReference(ByVal wsRef As Worksheet, ByVal wsTarget As Worksheet)
    Dim tailRef As Range
    Dim tailTarget As Range
    
    Set tailRef = wsRef.Range(PID_MS_PANEL_TAIL_RANGE)
    Set tailTarget = wsTarget.Range(PID_MS_PANEL_TAIL_RANGE)
    
    On Error Resume Next
    tailTarget.UnMerge
    tailTarget.ClearFormats
    Err.Clear
    On Error GoTo 0
    
    PID_MSCopyRangeFormatsMergeSafe tailRef, tailTarget
End Sub


Private Sub PID_MSRestoreMonthSheetDropdowns(ByVal wsTarget As Worksheet)
    ' E/F werden beim Format-Kopieren nicht mehr ueberschrieben; Lock-Policy stellt
    ' Whitelist wieder her, dann Dropdowns neu aufbauen.
    If wsTarget Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsTarget) Then Exit Sub
    
    PID_ApplyMonthSheetLockPolicy wsTarget
    
    PID_ApplyKVCodeDropdownValidation wsTarget
    RefreshKVStundenDropdownForSheet wsTarget
End Sub


Private Sub PID_MSApplyReferenceLayout(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    
    PID_MSApplyEmployeeBlockStyles ws
    PID_MSApplyRightPanelReferenceStyles ws
    PID_ApplyMonthSheetDateColumnWidths ws
    PID_ApplyMonthSheetEmployeeRowLayout ws
End Sub


' Sorszahlen in Spalte A (1., 2., ...) und einheitliche Datenzeilenhoehe.
Public Sub PID_ApplyMonthSheetEmployeeRowLayout(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    PID_ApplyMonthSheetRowIndexColumnLayout ws
    PID_MSRestoreEmployeeRowIndexColumn ws
    ws.Rows(PID_FIRST_ROW & ":" & PID_LAST_ROW).RowHeight = PID_STYLE_COMPACT_DATA_ROW_HEIGHT
    PID_ApplyMonthSheetAustrittsgrundLayout ws
End Sub


' Spalte A: schmal, Textformat (@), kein Datum — Referenz Januar (OOXML).
Public Sub PID_ApplyMonthSheetRowIndexColumnLayout(ByVal ws As Worksheet)
    Dim indexRange As Range
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    ws.Columns("A").ColumnWidth = PID_MS_ROW_INDEX_COL_WIDTH
    
    ' A1 (Monatsindex) Schriftfarbe #DDEBF7 — auf der Kopfband-Farbe unauffaellig.
    ws.Range("A1").Font.Color = RGB(221, 235, 247)
    
    Set indexRange = ws.Range("A" & PID_FIRST_ROW & ":A" & PID_LAST_ROW)
    indexRange.NumberFormat = PID_MS_ROW_INDEX_NUMBER_FORMAT
    indexRange.HorizontalAlignment = xlCenter
    indexRange.VerticalAlignment = xlCenter
End Sub


Public Sub PID_MSRestoreEmployeeRowIndexColumn(ByVal ws As Worksheet)
    Dim r As Long
    Dim indexRange As Range
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    Set indexRange = ws.Range("A" & PID_FIRST_ROW & ":A" & PID_LAST_ROW)
    indexRange.NumberFormat = PID_MS_ROW_INDEX_NUMBER_FORMAT
    
    For r = PID_FIRST_ROW To PID_LAST_ROW
        ws.Cells(r, PID_MS_EMPLOYEE_HEADER_FIRST_COL).Value = _
            CStr(r - PID_FIRST_ROW + 1) & "."
    Next r
End Sub


Private Sub PID_MSCopyMonthEmployeeHeaderFromReference(ByVal wsRef As Worksheet, ByVal wsTarget As Worksheet)
    Dim savedMonthIndex As Variant
    
    If wsRef Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub
    
    savedMonthIndex = wsTarget.Range("A1").Value2
    
    On Error Resume Next
    wsTarget.Range("B1:N2").UnMerge
    Err.Clear
    On Error GoTo 0
    
    wsRef.Range("B1:N2").Copy
    wsTarget.Range("B1:N2").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    
    PID_MSCopyRangeFormatsMergeSafe wsRef.Range("A1:A2"), wsTarget.Range("A1:A2")
    wsTarget.Range("A1").Value2 = savedMonthIndex
End Sub


Private Sub PID_MSEnsureEmployeeColumnHeaderMerges(ByVal ws As Worksheet)
    Dim col As Long
    Dim headerRange As Range
    Dim topText As String
    Dim bottomText As String
    Dim savedMonthIndex As Variant
    
    If ws Is Nothing Then Exit Sub
    
    savedMonthIndex = ws.Range("A1").Value2
    
    On Error Resume Next
    ws.Range("A1:N2").UnMerge
    On Error GoTo 0
    
    For col = PID_MS_EMPLOYEE_HEADER_FIRST_COL To PID_MS_EMPLOYEE_HEADER_LAST_COL
        topText = Trim$(CStr(ws.Cells(1, col).Value2))
        bottomText = Trim$(CStr(ws.Cells(2, col).Value2))
        
        Set headerRange = ws.Range(ws.Cells(1, col), ws.Cells(2, col))
        headerRange.Merge
        
        If col = PID_MS_EMPLOYEE_HEADER_FIRST_COL Then
            ws.Cells(1, col).Value2 = savedMonthIndex
        ElseIf Len(topText) > 0 And Len(bottomText) > 0 Then
            If Right$(topText, 1) = "-" Then
                ws.Cells(1, col).Value = topText & bottomText
            Else
                ws.Cells(1, col).Value = topText & " " & bottomText
            End If
        ElseIf Len(topText) = 0 And Len(bottomText) > 0 Then
            ws.Cells(1, col).Value = bottomText
        End If
        
        ws.Cells(1, col).HorizontalAlignment = xlCenter
        ws.Cells(1, col).VerticalAlignment = xlCenter
    Next col
End Sub


Private Sub PID_MSApplyEmployeeBlockStyles(ByVal ws As Worksheet)
    Dim col As Long
    Dim headerCell As Range
    
    ws.Rows("1:2").RowHeight = PID_STYLE_COMPACT_HEADER_ROW_HEIGHT
    
    PID_MSEnsureEmployeeColumnHeaderMerges ws
    
    For col = PID_MS_EMPLOYEE_HEADER_FIRST_COL To PID_MS_EMPLOYEE_HEADER_LAST_COL
        Set headerCell = ws.Cells(1, col)
        PID_StyleApplyCompactHeaderBand headerCell.MergeArea
    Next col
    
    PID_MSRestoreEmployeeRowIndexColumn ws
    ws.Rows(PID_FIRST_ROW & ":" & PID_LAST_ROW).RowHeight = PID_STYLE_COMPACT_DATA_ROW_HEIGHT
    
    PID_ApplyMonthSheetRowIndexColumnLayout ws
    PID_ApplyMonthEmployeeZebraRows ws
    
    ws.Range("A3:A82").HorizontalAlignment = xlCenter
    ws.Range("B3:N82").HorizontalAlignment = xlCenter
    ws.Range("B3:C82").HorizontalAlignment = xlLeft
    ws.Range("M3:N82").HorizontalAlignment = xlLeft
    
    PID_MSApplyBlockBorders ws.Range("A1:N2")
    PID_MSApplyBlockBorders ws.Range("A3:N82")
End Sub


' Zebra-Hintergrund fuer Mitarbeiterblock (Zeile 3=grau, 4=weiss, ...).
' Nach CopyData/Sort erneut aufrufen, weil Sort Formatierung mit verschiebt.
Public Sub PID_ApplyMonthEmployeeZebraRows(ByVal ws As Worksheet)
    Dim r As Long
    
    If ws Is Nothing Then Exit Sub
    
    ws.Range("B3:N82").Interior.Pattern = xlSolid
    PID_StyleApplyInputCell ws.Range("B3:F82")
    PID_StyleApplyInputCell ws.Range("I3:J82")
    PID_StyleApplyInputCell ws.Range("M3:N82")
    PID_StyleApplyReadOnlyGuideCell ws.Range("G3:G82")
    PID_StyleApplyReadOnlyGuideCell ws.Range("H3:H82")
    PID_StyleApplyReadOnlyGuideCell ws.Range("K3:K82")
    PID_StyleApplyReadOnlyGuideCell ws.Range("L3:L82")
    
    For r = PID_FIRST_ROW To PID_LAST_ROW Step 2
        ws.Range("B" & r & ":F" & r).Interior.Color = PID_StyleColorZebra()
        ws.Range("I" & r & ":J" & r).Interior.Color = PID_StyleColorZebra()
        ws.Range("M" & r & ":N" & r).Interior.Color = PID_StyleColorZebra()
    Next r
End Sub


' Hintergrund + Rahmen fuer B3:N82 nach Formel-Aenderungen in Spalte L (AutoFill-Kollateralschaden).
Public Sub PID_RestoreMonthSheetEmployeeBlockStyles(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    PID_ApplyMonthEmployeeZebraRows ws
    PID_MSApplyBlockBorders ws.Range("B3:N82")
End Sub


Private Sub PID_MSApplyRightPanelReferenceStyles(ByVal ws As Worksheet)
    Dim r As Long
    Dim panelRange As Range
    
    Set panelRange = ws.Range("O" & PID_MS_PANEL_FIRST_ROW & ":V" & PID_MS_PANEL_LAST_ROW)
    
    PID_MSResetPanelRange panelRange
    PID_MSApplyTopSummaryHeader ws
    
    For r = 3 To 5
        PID_MSApplyStyleToRangeMergedOnce ws.Range("O" & r & ":R" & r), PID_MS_STYLE_HEADER
        PID_MSApplyStyleToRangeMergedOnce ws.Range("S" & r & ":V" & r), PID_MS_STYLE_INPUT
    Next r
    
    PID_MSApplyStyleToRangeMergedOnce ws.Range("O7:V7"), PID_MS_STYLE_INPUT
    
    For r = 8 To 14
        PID_MSApplyStyleToRangeMergedOnce ws.Range("O" & r & ":R" & r), PID_MS_STYLE_READONLY
        PID_MSApplyStyleToRangeMergedOnce ws.Range("S" & r & ":V" & r), PID_MS_STYLE_INPUT
    Next r
    
    If PID_IsDurchrechnungPeriodStartMonth(ws) Then
        PID_MSApplyDurchrechnungStartMonthPanelRow15 ws
    Else
        PID_MSApplyStyleToRangeMergedOnce ws.Range("O15:P15"), PID_MS_STYLE_READONLY
        PID_MSApplyStyleToRangeMergedOnce ws.Range("Q15:R15"), PID_MS_STYLE_READONLY
        PID_MSApplyStyleToRangeMergedOnce ws.Range("S15:V15"), PID_MS_STYLE_INPUT
    End If
    
    If PID_IsDurchrechnungPeriodStartMonth(ws) Then
        PID_MSApplyStyleToRangeMergedOnce ws.Range("Q12:R12"), PID_MS_STYLE_READONLY
    Else
        PID_MSApplyStyleToRangeMergedOnce ws.Range("Q12:R12"), PID_MS_STYLE_INPUT
    End If
    
    PID_MSApplyStyleToRangeMergedOnce ws.Range("O16:V16"), PID_MS_STYLE_MESSAGE
    
    PID_MSApplyStyleToRangeMergedOnce ws.Range("O17:R17"), PID_MS_STYLE_READONLY
    
    For r = 18 To 28
        PID_MSApplyStyleToRangeMergedOnce ws.Range("O" & r & ":P" & r), PID_MS_STYLE_LABEL
        PID_MSApplyStyleToRangeMergedOnce ws.Range("Q" & r & ":R" & r), PID_MS_STYLE_INPUT
    Next r
    
    PID_MSApplyStyleToRangeMergedOnce ws.Range("O29:R29"), PID_MS_STYLE_READONLY
    
    PID_MSApplyStyleToRangeMergedOnce ws.Range("O30:V30"), PID_MS_STYLE_INPUT
    
    PID_MSApplyStyleToRangeMergedOnce ws.Range("O31:R31"), PID_MS_STYLE_READONLY
    ' FP-FLUKT: Label fuer die Fluktuationsanzeige persistent setzen (Wert steht in Q31,
    ' der Sync-Pfad ergaenzt das Label zusaetzlich als Sicherheitsnetz).
    ws.Range("O31").Value = "Fluktuation:"
    PID_MSApplyStyleToRangeMergedOnce ws.Range("S31:V31"), PID_MS_STYLE_INPUT
    
    PID_MSApplyStyleToRangeMergedOnce ws.Range("O32:V32"), PID_MS_STYLE_INPUT
    
    For r = 33 To 34
        PID_MSApplyStyleToRangeMergedOnce ws.Range("O" & r & ":P" & r), PID_MS_STYLE_INPUT
        PID_MSApplyStyleToRangeMergedOnce ws.Range("Q" & r & ":V" & r), PID_MS_STYLE_HEADER
    Next r
    
    For r = 35 To 42
        PID_MSApplyStyleToRangeMergedOnce ws.Range("O" & r & ":V" & r), PID_MS_STYLE_READONLY
    Next r
    
    PID_MSApplyEintrittsdatumSection ws
    
    For r = 46 To PID_MS_PANEL_LAST_ROW
        PID_MSApplyStyleToRangeMergedOnce ws.Range("O" & r & ":V" & r), PID_MS_STYLE_INPUT
    Next r
    
    PID_MSApplyBlockBorders ws.Range("O8:R16")
    PID_MSApplyBlockBorders ws.Range("O17:R29")
    PID_MSApplyBlockBorders ws.Range("O31:R31")
    PID_MSApplyBlockBorders ws.Range("O33:V42")
    
    PID_MSReinforceEdgeBorder ws.Range("R17:R29"), xlEdgeRight
    PID_MSClearAllBorders ws.Range("O43:V50")
    PID_MSClearAllBorders ws.Range("S16")
End Sub


Private Sub PID_MSApplyBlockBorders(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    PID_StyleApplyOuterBorder target
    PID_StyleApplyTableBorders target
End Sub


Private Sub PID_MSResetPanelRange(ByVal target As Range)
    Dim c As Range
    Dim handled As Collection
    Dim areaKey As String
    Dim resetTarget As Range
    
    If target Is Nothing Then Exit Sub
    
    Set handled = New Collection
    
    For Each c In target.Cells
        If c.MergeCells Then
            areaKey = c.MergeArea.Address(False, False)
        Else
            areaKey = c.Address(False, False)
        End If
        
        If PID_CollectionHasKey(handled, areaKey) Then GoTo NextCell
        
        handled.Add areaKey, areaKey
        
        If c.MergeCells Then
            Set resetTarget = c.MergeArea
        Else
            Set resetTarget = c
        End If
        
        With resetTarget
            .Interior.Pattern = xlSolid
            .Interior.Color = vbWhite
            .Font.Color = vbBlack
            .Font.Bold = False
        End With
        
NextCell:
    Next c
End Sub


Private Sub PID_MSApplyTopSummaryHeader(ByVal ws As Worksheet)
    Dim col As Long
    Dim headerCell As Range
    
    For col = 15 To 17
        Set headerCell = ws.Cells(1, col)
        If headerCell.MergeCells Then
            If Len(PID_MSGetRangeText(headerCell)) > 0 Then
                PID_StyleApplyCompactHeaderBand headerCell.MergeArea
            End If
        ElseIf Len(PID_MSGetRangeText(headerCell)) > 0 Then
            PID_StyleApplyCompactHeaderBand ws.Range(ws.Cells(1, col), ws.Cells(2, col))
        End If
    Next col
End Sub


Private Sub PID_MSApplyEintrittsdatumSection(ByVal ws As Worksheet)
    Dim r As Long
    Dim c As Range
    Dim hintRange As Range
    Dim cellText As String
    
    PID_MSClearAllBorders ws.Range("O43:V50")
    
    Set hintRange = Nothing
    
    For r = 43 To 46
        For Each c In ws.Range("O" & r & ":Q" & r).Cells
            cellText = UCase$(PID_MSGetRangeText(c))
            If InStr(1, cellText, "WAHL", vbTextCompare) > 0 _
               And InStr(1, cellText, "EINTRITT", vbTextCompare) > 0 Then
                If c.MergeCells Then
                    Set hintRange = c.MergeArea
                Else
                    Set hintRange = ws.Range("O" & r & ":Q" & r)
                End If
                Exit For
            End If
        Next c
        If Not hintRange Is Nothing Then Exit For
    Next r
    
    If Not hintRange Is Nothing Then
        With hintRange
            .Interior.Color = vbWhite
            .Font.Color = PID_StyleColorNavy()
            .Font.Bold = False
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = True
        End With
        PID_StyleApplyOuterBorder hintRange
    End If
    
    With ws.Range("O45:Q45")
        .Interior.Color = vbWhite
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .WrapText = True
    End With
End Sub


Private Sub PID_MSRemoveLegacyPanelShapes(ByVal ws As Worksheet)
    Dim shapeIndex As Long
    Dim shp As Shape
    
    For shapeIndex = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(shapeIndex)
        
        If InStr(1, shp.Name, "Klammer", vbTextCompare) > 0 _
               Or InStr(1, shp.Name, "Brace", vbTextCompare) > 0 _
               Or InStr(1, shp.Name, "Geschweifte", vbTextCompare) > 0 Then
            shp.Delete
        End If
    Next shapeIndex
End Sub


Private Function PID_MSEnsureAktualisierungButtonOnSheet(ByVal ws As Worksheet) As Boolean
    Dim btn As Shape
    Dim wasProtected As Boolean
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    Dim shapeIndex As Long
    Dim shp As Shape
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Function
    
    oldScreenUpdating = Application.ScreenUpdating
    
    wasProtected = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    ' Position von O1: bei ScreenUpdating=False auf Mac sonst falsche Left/Top-Werte.
    Application.ScreenUpdating = True
    PID_MSGetCopyDataButtonTargetGeometry ws, btnLeft, btnTop, btnWidth, btnHeight
    Application.ScreenUpdating = oldScreenUpdating
    
    For shapeIndex = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(shapeIndex)
        If PID_MSIsLegacyCopyDataButton(shp) Then shp.Delete
    Next shapeIndex
    
    Set btn = Nothing
    On Error Resume Next
    Set btn = ws.Shapes(PID_MS_COPYDATA_BUTTON_NAME)
    On Error GoTo SafeExit
    
    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                     Left:=btnLeft, _
                                     Top:=btnTop, _
                                     Width:=btnWidth, _
                                     Height:=btnHeight)
        btn.Name = PID_MS_COPYDATA_BUTTON_NAME
    End If
    
    PID_MSApplyCopyDataButtonRuntimeState btn, btnLeft, btnTop, btnWidth, btnHeight
    
    PID_MSEnsureAktualisierungButtonOnSheet = True

ReprotectSheet:
    On Error Resume Next
    If wasProtected Then
        PID_ReprotectWorksheet ws
    End If
    Exit Function

SafeExit:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    If wasProtected Then
        PID_ReprotectWorksheet ws
    End If
End Function


Private Sub PID_MSApplyCopyDataButtonRuntimeState(ByVal btn As Shape, _
                                                  ByVal btnLeft As Double, _
                                                  ByVal btnTop As Double, _
                                                  ByVal btnWidth As Double, _
                                                  ByVal btnHeight As Double)
    If btn Is Nothing Then Exit Sub
    
    On Error Resume Next
    btn.Placement = xlFreeFloating
    btn.Left = btnLeft
    btn.Top = btnTop
    btn.Width = btnWidth
    btn.Height = btnHeight
    btn.Visible = msoTrue
    btn.TextFrame.Characters.Text = PID_MS_COPYDATA_BUTTON_TEXT
    btn.OnAction = "CopyData"
    btn.ZOrder msoBringToFront
    On Error GoTo 0
    
    On Error Resume Next
    PID_StyleApplyToolbarButton btn, PID_StyleColorNavy(), PID_StyleColorBtnPrimaryLine(), RGB(255, 255, 255)
    On Error GoTo 0
End Sub


Private Sub PID_MSGetCopyDataButtonTargetGeometry(ByVal ws As Worksheet, _
                                                  ByRef btnLeft As Double, _
                                                  ByRef btnTop As Double, _
                                                  ByRef btnWidth As Double, _
                                                  ByRef btnHeight As Double)
    btnLeft = ws.Range("O1").Left + PID_MS_COPYDATA_BUTTON_OFFSET_LEFT
    btnTop = ws.Range("O1").Top + PID_MS_COPYDATA_BUTTON_OFFSET_TOP
    btnWidth = PID_MS_COPYDATA_BUTTON_WIDTH
    btnHeight = PID_MS_COPYDATA_BUTTON_HEIGHT
End Sub


Private Function PID_MSCopyDataButtonGeometryMatches(ByVal btn As Shape, _
                                                     ByVal expectedLeft As Double, _
                                                     ByVal expectedTop As Double, _
                                                     ByVal expectedWidth As Double, _
                                                     ByVal expectedHeight As Double) As Boolean
    If btn Is Nothing Then Exit Function
    
    PID_MSCopyDataButtonGeometryMatches = _
        (Abs(btn.Left - expectedLeft) <= 1.5) And _
        (Abs(btn.Top - expectedTop) <= 1.5) And _
        (Abs(btn.Width - expectedWidth) <= 1.5) And _
        (Abs(btn.Height - expectedHeight) <= 1.5)
End Function


Private Function PID_MSIsLegacyCopyDataButton(ByVal shp As Shape) As Boolean
    Dim actionText As String
    Dim shpText As String
    
    On Error GoTo SafeExit
    
    If shp Is Nothing Then Exit Function
    If shp.Name = PID_MS_COPYDATA_BUTTON_NAME Then Exit Function
    
    actionText = LCase$(Trim$(Replace$(shp.OnAction, "'", "")))
    If InStr(1, actionText, "copydata", vbTextCompare) > 0 Then
        PID_MSIsLegacyCopyDataButton = True
        Exit Function
    End If
    
    shpText = PID_MSGetShapeText(shp)
    If InStr(1, shpText, "Aktualisierung", vbTextCompare) > 0 Then
        PID_MSIsLegacyCopyDataButton = True
    End If

SafeExit:
End Function


Private Sub PID_MSClearAllBorders(ByVal target As Range)
    Dim edgeIndex As Long
    If target Is Nothing Then Exit Sub
    On Error Resume Next
    For edgeIndex = 7 To 12
        target.Borders(edgeIndex).LineStyle = xlLineStyleNone
    Next edgeIndex
    On Error GoTo 0
End Sub


Private Sub PID_MSReinforceEdgeBorder(ByVal target As Range, ByVal edgeIndex As Long)
    If target Is Nothing Then Exit Sub
    On Error Resume Next
    With target.Borders(edgeIndex)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_StyleColorNavy()
    End With
    On Error GoTo 0
End Sub


Private Function PID_MSGetShapeText(ByVal shp As Shape) As String
    On Error GoTo SafeExit
    If shp Is Nothing Then Exit Function
    
    PID_MSGetShapeText = shp.TextFrame.Characters.Text
    If Len(Trim$(PID_MSGetShapeText)) > 0 Then Exit Function
    
    On Error Resume Next
    PID_MSGetShapeText = shp.AlternativeText
    Err.Clear
SafeExit:
End Function


Private Function PID_MSGetRangeText(ByVal targetCell As Range) As String
    If targetCell Is Nothing Then Exit Function
    If targetCell.MergeCells Then
        PID_MSGetRangeText = Trim$(CStr(targetCell.MergeArea.Cells(1, 1).Value2))
    Else
        PID_MSGetRangeText = Trim$(CStr(targetCell.Value2))
    End If
End Function


Private Sub PID_MSApplyStyleToRangeMergedOnce(ByVal targetRange As Range, ByVal styleMode As Long)
    Dim c As Range
    Dim handled As Collection
    Dim areaKey As String
    Dim styleTarget As Range
    
    If targetRange Is Nothing Then Exit Sub
    
    Set handled = New Collection
    
    For Each c In targetRange.Cells
        If c.MergeCells Then
            areaKey = c.MergeArea.Address(False, False)
        Else
            areaKey = c.Address(False, False)
        End If
        
        If PID_CollectionHasKey(handled, areaKey) Then GoTo NextCell
        
        handled.Add areaKey, areaKey
        
        If c.MergeCells Then
            Set styleTarget = c.MergeArea
        Else
            Set styleTarget = c
        End If
        
        Select Case styleMode
            Case PID_MS_STYLE_INPUT
                PID_StyleApplyInputCell styleTarget
            Case PID_MS_STYLE_READONLY
                PID_StyleApplyReadOnlyGuideCell styleTarget
            Case PID_MS_STYLE_LABEL
                PID_StyleApplyInputGuideLabel styleTarget
            Case PID_MS_STYLE_HEADER
                PID_StyleApplyCompactHeaderBand styleTarget
            Case PID_MS_STYLE_MESSAGE
                With styleTarget
                    .Interior.Color = vbWhite
                    .Font.Color = PID_StyleColorNavy()
                    .Font.Bold = True
                    .HorizontalAlignment = xlLeft
                    .WrapText = True
                End With
        End Select
        
NextCell:
    Next c
End Sub

Attribute VB_Name = "Modul1"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Public Const PID_FIRST_ROW As Long = 3
Public Const PID_LAST_ROW As Long = 82
Public Const PID_WORKBOOK_PASSWORD As String = "company"
Public Const PID_EINSTELLUNG_SHEET As String = "EINSTELLUNG"
Public Const PID_LOHNTABELLE_SHEET As String = "LOHNTABELLE"
Public Const PID_FLUKTUATION_SHEET As String = "FLUKTUATION"
Public Const PID_WORKBOOK_YEAR_CELL As String = "C35"
Public Const PID_FLUKTUATION_REASON_FIRST_ROW As Long = 38
Public Const PID_FLUKTUATION_REASON_LAST_ROW As Long = 49
Public Const PID_FLUKTUATION_TIME_FIRST_ROW As Long = 53
Public Const PID_FLUKTUATION_TIME_LAST_ROW As Long = 59

Private mHeavyMaintOldCalculation As XlCalculation
Private mHeavyMaintDepth As Long

Private mViewPreserveSheet As Worksheet
Private mViewPreserveScrollRow As Long
Private mViewPreserveScrollCol As Long
Private mViewPreserveDepth As Long

Private mStaleValueFormattingChecked As Boolean


' Application.FormatStaleValues gibt es erst ab Microsoft 365. Frueh gebunden
' (Application.FormatStaleValues = ...) meldet Excel 2016 beim Kompilieren
' "Methode oder Datenobjekt nicht gefunden" — On Error hilft dort nicht, weil der
' Fehler schon beim Kompilieren entsteht. Ueber ein Object laeuft die Aufloesung
' erst zur Laufzeit; in Excel 2016 schlaegt der Aufruf still fehl.
' Die Eigenschaft gilt fuer die ganze Excel-Anwendung, deshalb genuegt einmal
' je Sitzung statt bei jeder Zellaenderung.
Public Sub PID_DisableStaleValueFormatting()
    Dim excelApp As Object
    
    If mStaleValueFormattingChecked Then Exit Sub
    
    On Error Resume Next
    mStaleValueFormattingChecked = True
    Set excelApp = Application
    excelApp.FormatStaleValues = False
    Err.Clear
End Sub


' Frueher: Manual fuer schnelles Oeffnen — Endanwender sahen leere H/K/L-Formeln.
' Jetzt: Automatisch + EnableCalculation; Open bleibt kurz Manual nur in Workbook_Open.
Public Sub PID_ConfigureDeferredWorkbookCalculationOnOpen()
    On Error Resume Next
    PID_DisableStaleValueFormatting
    PID_EnableCalculationForAllSheets
    Application.Calculation = xlCalculationAutomatic
End Sub


' Schwere Wartung (KV-Refresh): Manual waehrenddessen, danach H/K/L einmal — FP-018.
Public Sub PID_BeginHeavyMaintenance()
    On Error Resume Next
    
    If mHeavyMaintDepth <= 0 Then
        mHeavyMaintOldCalculation = Application.Calculation
        Application.Calculation = xlCalculationManual
    End If
    
    mHeavyMaintDepth = mHeavyMaintDepth + 1
    Err.Clear
End Sub


Public Sub PID_EndHeavyMaintenance(Optional ByVal wsMonth As Worksheet = Nothing)
    On Error Resume Next
    
    If mHeavyMaintDepth > 0 Then mHeavyMaintDepth = mHeavyMaintDepth - 1
    If mHeavyMaintDepth > 0 Then Exit Sub
    
    If Not wsMonth Is Nothing Then
        If PID_IsWorkerMonthSheet(wsMonth) Then
            PID_RecalculateMonthFormulaColumns wsMonth
        End If
    End If
    
    ' Ein ungueltiger Merkwert (z. B. 0, wenn End ohne passendes Begin laeuft) wuerde
    ' lautlos fehlschlagen und die Mappe im Manual-Modus stehen lassen.
    Select Case mHeavyMaintOldCalculation
        Case xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic
            Application.Calculation = mHeavyMaintOldCalculation
        Case Else
            Application.Calculation = xlCalculationAutomatic
    End Select
    
    Err.Clear
End Sub


' Aktives Blatt + Scrollposition waehrend Hintergrund-Refresh (z. B. Monatsloop) sichern.
Public Sub PID_BeginPreserveWorkbookView()
    On Error Resume Next
    
    If mViewPreserveDepth <= 0 Then
        Set mViewPreserveSheet = ActiveSheet
        mViewPreserveScrollRow = 0
        mViewPreserveScrollCol = 0
        
        If Not ActiveWindow Is Nothing Then
            mViewPreserveScrollRow = ActiveWindow.ScrollRow
            mViewPreserveScrollCol = ActiveWindow.ScrollColumn
        End If
    End If
    
    mViewPreserveDepth = mViewPreserveDepth + 1
    Err.Clear
End Sub


Public Sub PID_EndPreserveWorkbookView()
    On Error Resume Next
    
    If mViewPreserveDepth > 0 Then mViewPreserveDepth = mViewPreserveDepth - 1
    If mViewPreserveDepth > 0 Then Exit Sub
    
    If Not mViewPreserveSheet Is Nothing Then
        mViewPreserveSheet.Activate
        
        If Not ActiveWindow Is Nothing Then
            If mViewPreserveScrollRow > 0 Then ActiveWindow.ScrollRow = mViewPreserveScrollRow
            If mViewPreserveScrollCol > 0 Then ActiveWindow.ScrollColumn = mViewPreserveScrollCol
        End If
    End If
    
    Set mViewPreserveSheet = Nothing
    Err.Clear
End Sub


Public Sub PID_RestoreLOHNTABELLEView(ByVal wsKV As Worksheet, _
                                      ByVal scrollRow As Long, _
                                      ByVal scrollCol As Long, _
                                      Optional ByVal focusRow As Long = 0, _
                                      Optional ByVal focusCol As String = "G")
    On Error Resume Next
    
    If wsKV Is Nothing Then Exit Sub
    
    wsKV.Activate
    
    If Not ActiveWindow Is Nothing Then
        If scrollRow > 0 Then ActiveWindow.ScrollRow = scrollRow
        If scrollCol > 0 Then ActiveWindow.ScrollColumn = scrollCol
    End If
    
    If focusRow >= 4 Then
        wsKV.Range(focusCol & CStr(focusRow)).Select
    End If
    
    Err.Clear
End Sub


Public Sub PID_RecalculateMonthFormulaColumns(ByVal wsMonth As Worksheet)
    On Error Resume Next
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    
    wsMonth.Range("H3:H82").Calculate
    wsMonth.Range("K3:K82").Calculate
    wsMonth.Range("L3:L82").Calculate
    Err.Clear
End Sub


Public Sub PID_EnableCalculationForAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.EnableCalculation = True
        Err.Clear
    Next ws
End Sub


Public Sub PID_FormatAllMoneyColumns()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            On Error Resume Next
            ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
            On Error GoTo CleanFail
            
            PID_ApplyEuroNumberFormat ws.Range("G3:G82")
            PID_ApplyEuroNumberFormat ws.Range("J3:J82")
            PID_ApplyEuroNumberFormat ws.Range("K3:K82")
            PID_ApplyLetztesGehaltNumberFormat ws.Range("L3:L82")
            
            ws.Columns("G").ColumnWidth = 13
            ws.Columns("J").ColumnWidth = 13
            ws.Columns("K").ColumnWidth = 14
            ws.Columns("L").ColumnWidth = 14
            PID_ApplyMonthSheetDateColumnWidths ws
            PID_ApplyMonthSheetAustrittsgrundLayout ws
            
            PID_ProtectWorkerMonthSheet ws
        End If
    Next i
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    On Error GoTo CleanFail
    
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
        On Error GoTo CleanFail
        
        PID_ApplyEuroNumberFormat ws.Range("H:H")
        ws.Columns("H").ColumnWidth = 14
        
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_FormatAllMoneyColumns:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Formatierung"
End Sub


Public Sub PID_ResetExcelState()
    On Error Resume Next
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    
    ' Verwaiste Tiefenzaehler zuruecksetzen: bleibt einer nach einem abgebrochenen
    ' Durchlauf stehen, wird der Rechenmodus danach nie wieder zurueckgestellt.
    mHeavyMaintDepth = 0
    mHeavyMaintOldCalculation = xlCalculationAutomatic
    mViewPreserveDepth = 0
    Set mViewPreserveSheet = Nothing
    
    ' TR-02: bleibt der Zaehler nach einem Abbruch stehen, greift der Einfuege-Schutz
    ' nicht mehr. Gleichzeitig die Strg+V-Belegung neu setzen.
    PID_ResetManagedPasteState
    PID_InstallPasteHooks
    
    On Error GoTo 0
    
    MsgBox "Excel wurde " & PID_UTxtZurueckgesetzt() & ".", _
           vbInformation, "Excel Reset"
End Sub


Public Sub PID_HideTechnicalSheets()
    On Error Resume Next
    
    ThisWorkbook.Worksheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("KV_DROPDOWN_HELPER").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(PID_ADMIN_SHEET_NAME).Visible = xlSheetVeryHidden
    
    On Error GoTo 0
    
    MsgBox "Technische " & PID_UTxtBlaetter() & " wurden ausgeblendet.", _
           vbInformation, "Technische " & PID_UTxtBlaetter()
End Sub


Public Sub PID_ShowTechnicalSheets()
    On Error Resume Next
    
    ThisWorkbook.Worksheets("FLUKTUATION_DATEN").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("KV_DROPDOWN_HELPER").Visible = xlSheetVisible
    
    On Error GoTo 0
    
    MsgBox "Technische " & PID_UTxtBlaetter() & " wurden sichtbar gemacht.", _
           vbInformation, "Technische " & PID_UTxtBlaetter()
End Sub


Public Function PID_GetWorkbookYear() As Long
    Dim wsConfig As Worksheet
    
    On Error GoTo Fallback
    
    Set wsConfig = ThisWorkbook.Worksheets(PID_EINSTELLUNG_SHEET)
    
    If IsNumeric(wsConfig.Range(PID_WORKBOOK_YEAR_CELL).Value) Then
        PID_GetWorkbookYear = CLng(wsConfig.Range(PID_WORKBOOK_YEAR_CELL).Value)
        Exit Function
    End If

Fallback:
    PID_GetWorkbookYear = Year(Date)
End Function


Public Function PID_WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    PID_WorksheetExists = Not ws Is Nothing
End Function


Public Function PID_CountMonthSheets() As Long
    Dim monthNames As Variant
    Dim i As Long
    Dim countValue As Long
    
    monthNames = PID_MonthNames()
    countValue = 0
    
    For i = LBound(monthNames) To UBound(monthNames)
        If PID_WorksheetExists(CStr(monthNames(i))) Then
            countValue = countValue + 1
        End If
    Next i
    
    PID_CountMonthSheets = countValue
End Function


Public Function PID_YesNoText(ByVal valueToCheck As Boolean) As String
    If valueToCheck Then
        PID_YesNoText = "OK"
    Else
        PID_YesNoText = "FEHLT"
    End If
End Function


Public Function PID_GetCalculationModeText() As String
    Select Case Application.Calculation
        Case xlCalculationAutomatic
            PID_GetCalculationModeText = "Automatisch"
        Case xlCalculationManual
            PID_GetCalculationModeText = "Manuell"
        Case xlCalculationSemiautomatic
            PID_GetCalculationModeText = "Teilautomatisch"
        Case Else
            PID_GetCalculationModeText = "Unbekannt"
    End Select
End Function


Public Sub PID_ApplyEuroNumberFormat(ByVal targetRange As Range)
    Dim euroSymbol As String
    
    If targetRange Is Nothing Then Exit Sub
    
    euroSymbol = ChrW(8364)
    
    On Error GoTo TryEnglishFormat
    
    targetRange.NumberFormatLocal = euroSymbol & " #.##0,00"
    Exit Sub

TryEnglishFormat:
    On Error GoTo SafeExit
    
    targetRange.NumberFormat = euroSymbol & " #,##0.00"

SafeExit:
End Sub


Public Sub PID_ApplyLetztesGehaltNumberFormat(ByVal targetRange As Range)
    Dim euroSymbol As String
    
    If targetRange Is Nothing Then Exit Sub
    
    euroSymbol = ChrW(8364)
    
    On Error GoTo TryEnglishFormat
    
    ' Positiv;Negativ;Null (leer);Text — 0 wird nicht als €0,00 angezeigt.
    targetRange.NumberFormatLocal = euroSymbol & " #.##0,00;-" & euroSymbol & " #.##0,00;;@"
    Exit Sub

TryEnglishFormat:
    On Error GoTo SafeExit
    
    targetRange.NumberFormat = euroSymbol & " #,##0.00;-" & euroSymbol & " #,##0.00;;@"

SafeExit:
End Sub


Public Sub PID_RestoreAktuelleStundenFormulas()
    Dim updatedCount As Long
    
    updatedCount = PID_RestoreAktuelleStundenFormulasSilent
    
    MsgBox "Aktuelle-Stunden-Formeln wurden wiederhergestellt." & vbCrLf & vbCrLf & _
           PID_UTxtMonatsblaetter() & " aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
           "Bereich: H" & PID_FIRST_ROW & ":H" & PID_LAST_ROW & vbCrLf & _
           "Jahr aus: " & PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL, _
           vbInformation, "Aktuelle Stunden"
End Sub


Public Function PID_RestoreAktuelleStundenFormulasSilent() As Long
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim formulaR1C1 As String
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    formulaR1C1 = PID_GetAktuelleStundenFormulaR1C1()
    monthNames = PID_MonthNames()
    PID_RestoreAktuelleStundenFormulasSilent = 0
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If PID_RestoreAktuelleStundenFormulasOnSheet(ws, formulaR1C1) Then
                PID_RestoreAktuelleStundenFormulasSilent = PID_RestoreAktuelleStundenFormulasSilent + 1
            End If
        End If
    Next i
    
    GoTo CleanExit

CleanFail:
    PID_RestoreAktuelleStundenFormulasSilent = 0

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Function


Public Sub PID_RecalculateAktuelleStundenForChangedRows(ByVal wsMonth As Worksheet, ByVal changedRange As Range)
    Dim rowsToCheck As Range
    Dim c As Range
    Dim checkedRows As Collection
    Dim rowKey As String
    Dim rowNumber As Long
    Dim calcRange As Range
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    
    Set rowsToCheck = Intersect(changedRange, wsMonth.Range("D3:F82,I3:I82"))
    If rowsToCheck Is Nothing Then Exit Sub
    
    Set checkedRows = New Collection
    
    For Each c In rowsToCheck.Cells
        If c.Row >= PID_FIRST_ROW And c.Row <= PID_LAST_ROW Then
            rowKey = CStr(c.Row)
            
            If Not PID_CollectionHasKey(checkedRows, rowKey) Then
                checkedRows.Add rowKey, rowKey
            End If
        End If
    Next c
    
    For Each c In checkedRows
        rowNumber = CLng(c)
        
        If calcRange Is Nothing Then
            Set calcRange = wsMonth.Cells(rowNumber, "H")
        Else
            Set calcRange = Union(calcRange, wsMonth.Cells(rowNumber, "H"))
        End If
    Next c
    
    If Not calcRange Is Nothing Then
        On Error Resume Next
        calcRange.Calculate
        Err.Clear
    End If

SafeExit:
End Sub


Private Function PID_RestoreAktuelleStundenFormulasOnSheet(ByVal ws As Worksheet, _
                                                           ByVal formulaR1C1 As String) As Boolean
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Function
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Function
    
    wasProtected = ws.ProtectContents
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    If Trim$(CStr(ws.Range("H1").Value)) = "" Then
        ws.Range("H1").Value = "Aktuelle Stunden"
    End If
    
    ws.Range("H" & PID_FIRST_ROW & ":H" & PID_LAST_ROW).FormulaR1C1 = formulaR1C1
    ws.Range("H" & PID_FIRST_ROW & ":H" & PID_LAST_ROW).NumberFormat = "0.00"
    
    PID_RestoreAktuelleStundenFormulasOnSheet = True

SafeExit:
    On Error Resume Next
    If wasProtected Then
        PID_ReprotectWorksheet ws
    End If
End Function


Private Function PID_GetEinstellungYearRefR1C1() As String
    ' PID_WORKBOOK_YEAR_CELL = C35 -> R35C3 in R1C1 notation.
    PID_GetEinstellungYearRefR1C1 = "'" & PID_EINSTELLUNG_SHEET & "'!R35C3"
End Function


Private Function PID_GetEinstellungYearRefA1() As String
    PID_GetEinstellungYearRefA1 = PID_EINSTELLUNG_SHEET & "!$C$35"
End Function


Private Function PID_MonthSheetHasLetztesGehaltRefError(ByVal wsMonth As Worksheet) As Boolean
    Dim formulaText As String
    
    If wsMonth Is Nothing Then Exit Function
    
    On Error Resume Next
    formulaText = UCase$(CStr(wsMonth.Range("L" & PID_FIRST_ROW).Formula))
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    PID_MonthSheetHasLetztesGehaltRefError = (InStr(1, formulaText, "#REF", vbTextCompare) > 0)
End Function


Private Function PID_GetAktuelleStundenFormulaR1C1() As String
    Dim yearRef As String
    
    yearRef = PID_GetEinstellungYearRefR1C1()
    
    PID_GetAktuelleStundenFormulaR1C1 = _
        "=IF(AND(ISNUMBER(RC[-4]),ISNUMBER(RC[-2]))," & _
        "ROUNDDOWN(MIN(" & _
        "IF(OR(ISBLANK(RC[1]),RC[1]="""")," & _
        "EOMONTH(DATE(" & yearRef & ",R1C1,1),0)," & _
        "RC[1])-" & _
        "MAX(RC[-4],DATE(" & yearRef & ",R1C1,1))+1)/" & _
        "DAY(EOMONTH(DATE(" & yearRef & ",R1C1,1),0))*RC[-2],2),"""")"
End Function


Public Sub RestoreKVStundenDropdownValidation()
    RefreshAllMonthKVStundenDropdowns
    MarkKVDropdownsClean
    
    MsgBox "Stunden-Dropdown (Spalte F) wurde wiederhergestellt." & vbCrLf & vbCrLf & _
           "Bereich: F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW & " auf allen " & PID_UTxtMonatsblaettern() & ".", _
           vbInformation, "Spalte F"
End Sub


Public Sub PID_RestoreLetztesGehaltFormulas()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim updatedCount As Long
    
    On Error GoTo CleanFail
    
    PID_RestoreLetztesGehaltFormulasSilent
    
    monthNames = PID_MonthNames()
    updatedCount = 0
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If Not PID_MonthSheetHasLetztesGehaltRefError(ws) Then
                updatedCount = updatedCount + 1
            End If
        End If
    Next i
    
    MsgBox "Letztes-Gehalt-Formeln (Spalte L) wurden wiederhergestellt." & vbCrLf & vbCrLf & _
           PID_UTxtMonatsblaetter() & " aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
           "Bereich: L" & PID_FIRST_ROW & ":L" & PID_LAST_ROW & vbCrLf & _
           "Jahr aus: " & PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL & vbCrLf & vbCrLf & _
           PID_GetLetztesGehaltRestoreStatusText() & vbCrLf & vbCrLf & _
           "AVG Bruttolohn (Q42) sollte danach wieder berechnet werden.", _
           vbInformation, "Spalte L"
    
    GoTo CleanExit

CleanExit:
    Exit Sub

CleanFail:
    MsgBox "Fehler bei PID_RestoreLetztesGehaltFormulas:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Spalte L"
End Sub


Public Sub PID_EnsureLetztesGehaltFormulas()
    If Not PID_AnyMonthNeedsLetztesGehaltRestore() Then Exit Sub
    PID_RestoreLetztesGehaltFormulasSilent
End Sub


Private Function PID_AnyMonthNeedsLetztesGehaltRestore() As Boolean
    Dim ws As Worksheet
    
    Set ws = Nothing
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Januar")
    Err.Clear
    
    If ws Is Nothing Then Exit Function
    
    If Not PID_MonthSheetHasLetztesGehaltFormula(ws) Then
        PID_AnyMonthNeedsLetztesGehaltRestore = True
        Exit Function
    End If
    
    If PID_MonthSheetHasLetztesGehaltRefError(ws) Then
        PID_AnyMonthNeedsLetztesGehaltRestore = True
        Exit Function
    End If
    
    If PID_MonthSheetHasLetztesGehaltStaticValues(ws) Then
        PID_AnyMonthNeedsLetztesGehaltRestore = True
        Exit Function
    End If
    
    If Not PID_MonthSheetHasLetztesGehaltEmptyZeroFix(ws) Then
        PID_AnyMonthNeedsLetztesGehaltRestore = True
    End If
End Function


Public Sub PID_RestoreLetztesGehaltFormulasSilent()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim formulaR1C1 As String
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo SafeExit
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    formulaR1C1 = PID_GetLetztesGehaltFormulaR1C1()
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        Err.Clear
        
        If Not ws Is Nothing Then
            PID_RestoreLetztesGehaltFormulasOnSheet ws, formulaR1C1
            Err.Clear
        End If
    Next i
    
    On Error Resume Next
    Application.CalculateFull
    Err.Clear

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Public Function PID_MonthSheetHasLetztesGehaltFormula(ByVal wsMonth As Worksheet) As Boolean
    Dim formulaText As String
    
    If wsMonth Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheetName(wsMonth.Name) Then Exit Function
    
    On Error Resume Next
    formulaText = UCase$(CStr(wsMonth.Range("L" & PID_FIRST_ROW).FormulaR1C1))
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    If Left$(formulaText, 1) <> "=" Then Exit Function
    If InStr(1, formulaText, "#REF", vbTextCompare) > 0 Then Exit Function
    If InStr(1, formulaText, UCase$(PID_EINSTELLUNG_SHEET), vbTextCompare) = 0 Then Exit Function
    If InStr(1, formulaText, "R35C3", vbTextCompare) = 0 Then
        If InStr(1, formulaText, "$C$35", vbTextCompare) = 0 Then Exit Function
    End If
    
    PID_MonthSheetHasLetztesGehaltFormula = True
End Function


Private Function PID_RestoreLetztesGehaltFormulasOnSheet(ByVal ws As Worksheet, _
                                                         ByVal formulaR1C1 As String) As Boolean
    Dim wasProtected As Boolean
    Dim targetRange As Range
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheetName(ws.Name) Then Exit Function
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Function
    
    wasProtected = ws.ProtectContents
    
    If Not PID_TryUnprotectMonthSheetForMacro(ws) Then Exit Function
    
    If Trim$(CStr(ws.Range("L1").Value)) = "" Then
        ws.Range("L1").Value = "Letztes Gehalt"
    End If
    
    Set targetRange = ws.Range("L" & PID_FIRST_ROW & ":L" & PID_LAST_ROW)
    PID_ApplyLetztesGehaltFormulaToSheet ws, targetRange, formulaR1C1
    PID_ApplyLetztesGehaltNumberFormat targetRange
    PID_RestoreMonthSheetEmployeeBlockStyles ws
    
    On Error Resume Next
    targetRange.Calculate
    Err.Clear
    On Error GoTo SafeExit
    
    PID_RestoreLetztesGehaltFormulasOnSheet = Not PID_MonthSheetHasLetztesGehaltRefError(ws)

SafeExit:
    On Error Resume Next
    If Not ws Is Nothing Then
        If wasProtected Then
            PID_ReprotectWorksheet ws
        End If
    End If
End Function


Private Function PID_TryUnprotectMonthSheetForMacro(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If ws.ProtectContents Then ws.Unprotect
    PID_TryUnprotectMonthSheetForMacro = Not ws.ProtectContents
End Function


Private Sub PID_ApplyLetztesGehaltFormulaToSheet(ByVal ws As Worksheet, _
                                                 ByVal targetRange As Range, _
                                                 ByVal formulaR1C1 As String)
    Dim firstCell As Range
    Dim existingFormula As String
    Dim wrappedFormula As String
    Dim rowNum As Long
    
    If ws Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    
    rowNum = targetRange.Row
    Set firstCell = targetRange.Cells(1, 1)
    existingFormula = Trim$(CStr(firstCell.Formula))
    
    If PID_FormulaHasLetztesGehaltEmployeeGuard(existingFormula) Then
        firstCell.FormulaR1C1 = formulaR1C1
    ElseIf Len(existingFormula) > 3 And InStr(1, existingFormula, "#REF", vbTextCompare) = 0 Then
        If Left$(existingFormula, 1) = "=" Then existingFormula = Mid$(existingFormula, 2)
        ' Vier Anfuehrungszeichen ergeben in der Formel den leeren Text "".
        ' Mit zwei Zeichen entstand frueher OR($B3=",$C3=") - ein Textvergleich,
        ' der immer FALSCH ist, statt des gewollten B/C-Schutzes.
        wrappedFormula = "=IF(OR($B" & CStr(rowNum) & "="""",$C" & CStr(rowNum) & "=""""),""""," & _
            "IF(" & existingFormula & "=0,""""," & existingFormula & "))"
        firstCell.Formula = wrappedFormula
    Else
        firstCell.FormulaR1C1 = formulaR1C1
    End If
    
    If targetRange.Rows.Count > 1 Then
        PID_FillFormulaDownWithoutFormats firstCell, targetRange
    End If
End Sub


Private Sub PID_FillFormulaDownWithoutFormats(ByVal firstCell As Range, ByVal targetRange As Range)
    Dim fillRange As Range
    
    If firstCell Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    If targetRange.Rows.Count <= 1 Then Exit Sub
    
    Set fillRange = targetRange.Offset(1, 0).Resize(targetRange.Rows.Count - 1, 1)
    
    On Error Resume Next
    firstCell.Copy
    fillRange.PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    Err.Clear
End Sub


Private Function PID_FormulaHasLetztesGehaltEmployeeGuard(ByVal formulaText As String) As Boolean
    Dim compactFormula As String
    
    compactFormula = UCase$(Replace(Replace(Replace(formulaText, " ", ""), vbLf, ""), vbTab, ""))
    
    If InStr(1, compactFormula, "RC[-10]", vbTextCompare) > 0 Then
        PID_FormulaHasLetztesGehaltEmployeeGuard = True
        Exit Function
    End If
    
    PID_FormulaHasLetztesGehaltEmployeeGuard = _
        (InStr(1, compactFormula, "$B", vbTextCompare) > 0) And _
        (InStr(1, compactFormula, "$C", vbTextCompare) > 0) And _
        (InStr(1, compactFormula, "=""", vbTextCompare) > 0 Or InStr(1, compactFormula, "="";", vbTextCompare) > 0)
End Function


Private Function PID_GetLetztesGehaltRestoreStatusText() As String
    Dim ws As Worksheet
    Dim probeFormula As String
    
    On Error GoTo SafeExit
    
    Set ws = ThisWorkbook.Worksheets("Februar")
    probeFormula = Left$(Trim$(CStr(ws.Range("L" & PID_FIRST_ROW).Formula)), 72)
    
    If PID_FormulaHasLetztesGehaltEmployeeGuard(CStr(ws.Range("L" & PID_FIRST_ROW).Formula)) Then
        PID_GetLetztesGehaltRestoreStatusText = "Pruefung Februar L" & PID_FIRST_ROW & ": B/C-Schutz aktiv."
    Else
        PID_GetLetztesGehaltRestoreStatusText = "WARNUNG: Februar L" & PID_FIRST_ROW & " noch alt:" & vbCrLf & probeFormula
    End If
    Exit Function

SafeExit:
    PID_GetLetztesGehaltRestoreStatusText = "Pruefung Februar L" & PID_FIRST_ROW & " nicht moeglich."
End Function


Public Function PID_GetLetztesGehaltFormulaR1C1() As String
    Dim yearRef As String
    Dim coreFormula As String
    
    yearRef = PID_GetEinstellungYearRefR1C1()
    
    ' Spalte L = Letztes Gehalt / Laborcost fuer AVG Bruttolohn (Q42 = Q17/Q8).
    ' Jahr aus EINSTELLUNG!C35 statt legacy LOHNTABELLE!G3 (#REF!).
    coreFormula = "IFERROR(IF(OR(ISBLANK(RC[-8]),ISBLANK(RC[-6]),ISBLANK(RC[-5]),ISBLANK(RC[-1])),0," & _
        "IF(YEAR(RC[-8])<" & yearRef & "," & _
        "IF(ISNUMBER(RC[-3]),IF(YEAR(RC[-3])<" & yearRef & ",0," & _
        "IF(OR(YEAR(RC[-3])>" & yearRef & ",MONTH(RC[-3])>R1C1),RC[-5]," & _
        "IF(MONTH(RC[-3])<R1C1,0,(RC[-5]/DAY(EOMONTH(DATE(" & yearRef & ",R1C1,1),0)))*DAY(RC[-3])+RC[-1]))))," & _
        "RC[-5]+RC[-1])," & _
        "IF(YEAR(RC[-8])=" & yearRef & "," & _
        "IF(MONTH(RC[-8])<R1C1,IF(ISNUMBER(RC[-3]),IF(YEAR(RC[-3])<" & yearRef & ",0," & _
        "IF(OR(YEAR(RC[-3])>" & yearRef & ",MONTH(RC[-3])>R1C1),RC[-5]," & _
        "IF(MONTH(RC[-3])<R1C1,0,(RC[-5]/DAY(EOMONTH(DATE(" & yearRef & ",R1C1,1),0)))*DAY(RC[-3])+RC[-1])))),RC[-5])," & _
        "IF(MONTH(RC[-8])=R1C1,IF(ISNUMBER(RC[-3]),IF(OR(YEAR(RC[-3])>" & yearRef & ",MONTH(RC[-3])>R1C1),RC[-5]," & _
        "(RC[-5]/DAY(EOMONTH(RC[-8],0)))*(RC[-3]-RC[-8]+1)+RC[-1])," & _
        "(RC[-5]/DAY(EOMONTH(RC[-8],0)))*(DAY(EOMONTH(RC[-8],0))-DAY(RC[-8])+1)+RC[-1]),0)),0))),0)"
    
    ' Kein Mitarbeiter (B/C leer) oder Ergebnis 0 -> leer, analog Spalte G ohne KV/Stunden.
    PID_GetLetztesGehaltFormulaR1C1 = _
        "=IF(OR(RC[-10]="""",RC[-9]=""""),""""," & _
        "IF(" & coreFormula & "=0,""""," & coreFormula & "))"
End Function


Private Function PID_MonthSheetHasLetztesGehaltEmptyZeroFix(ByVal wsMonth As Worksheet) As Boolean
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Function
    If Not PID_MonthSheetHasLetztesGehaltFormula(wsMonth) Then Exit Function
    
    PID_MonthSheetHasLetztesGehaltEmptyZeroFix = _
        PID_FormulaHasLetztesGehaltEmployeeGuard(CStr(wsMonth.Range("L" & PID_FIRST_ROW).Formula))

SafeExit:
End Function


Public Sub PID_RecalculateLetztesGehaltForChangedRows(ByVal wsMonth As Worksheet, ByVal changedRange As Range)
    Dim watchRange As Range
    Dim rowsToCheck As Range
    Dim c As Range
    Dim checkedRows As Collection
    Dim rowKey As String
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    
    ' E/F: Monatslohn-Handler aktualisiert L bereits — hier nur D, G, I, K.
    Set watchRange = Union(wsMonth.Range("D3:D82"), wsMonth.Range("G3:G82"), _
                           wsMonth.Range("I3:I82"), wsMonth.Range("K3:K82"))
    Set rowsToCheck = Intersect(changedRange, watchRange)
    If rowsToCheck Is Nothing Then Exit Sub
    
    Set checkedRows = New Collection
    
    For Each c In rowsToCheck.Cells
        rowKey = CStr(c.Row)
        
        If c.Row >= PID_FIRST_ROW And c.Row <= PID_LAST_ROW Then
            If Not PID_CollectionHasKey(checkedRows, rowKey) Then
                checkedRows.Add rowKey, rowKey
            End If
        End If
    Next c
    
    PID_RecalculateLetztesGehaltForRows wsMonth, checkedRows

SafeExit:
End Sub


Public Sub PID_RecalculateLetztesGehaltForRows(ByVal wsMonth As Worksheet, ByVal rows As Collection)
    Dim i As Long
    Dim rowNumber As Long
    Dim calcRange As Range
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then Exit Sub
    
    For i = 1 To rows.Count
        rowNumber = CLng(rows(i))
        
        If rowNumber >= PID_FIRST_ROW And rowNumber <= PID_LAST_ROW Then
            If calcRange Is Nothing Then
                Set calcRange = wsMonth.Cells(rowNumber, "L")
            Else
                Set calcRange = Union(calcRange, wsMonth.Cells(rowNumber, "L"))
            End If
        End If
    Next i
    
    If Not calcRange Is Nothing Then
        On Error Resume Next
        calcRange.Calculate
        Err.Clear
    End If

SafeExit:
End Sub


Public Sub PID_RecalculateLetztesGehaltForRow(ByVal wsMonth As Worksheet, ByVal rowNumber As Long)
    Dim lCell As Range
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Sub
    
    Set lCell = wsMonth.Cells(rowNumber, "L")
    If Not lCell.HasFormula Then Exit Sub
    
    On Error Resume Next
    lCell.Calculate
    Err.Clear

SafeExit:
End Sub


Private Function PID_MonthSheetHasLetztesGehaltStaticValues(ByVal wsMonth As Worksheet) As Boolean
    Dim r As Long
    Dim cellValue As Variant
    
    If wsMonth Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheetName(wsMonth.Name) Then Exit Function
    
    For r = PID_FIRST_ROW To PID_LAST_ROW
        If Not wsMonth.Cells(r, "L").HasFormula Then
            cellValue = wsMonth.Cells(r, "L").Value2
            If Not IsEmpty(cellValue) And cellValue <> 0 Then
                PID_MonthSheetHasLetztesGehaltStaticValues = True
                Exit Function
            End If
            
            If Len(Trim$(CStr(wsMonth.Cells(r, "D").Value))) > 0 _
               Or Len(Trim$(CStr(wsMonth.Cells(r, "E").Value))) > 0 Then
                PID_MonthSheetHasLetztesGehaltStaticValues = True
                Exit Function
            End If
        End If
    Next r
End Function


Private Function PID_IsWorkerMonthSheetName(ByVal sheetName As String) As Boolean
    Select Case sheetName
        Case "Januar", "Februar", "Marz", "April", "Mai", "Juni", _
             "Juli", "August", "September", "Oktober", "November", "Dezember"
            PID_IsWorkerMonthSheetName = True
        Case Else
            PID_IsWorkerMonthSheetName = False
    End Select
End Function


Private Function PID_GetEuroSymbol() As String
    Dim configuredSymbol As String
    
    configuredSymbol = Application.International(xlCurrencyCode)
    
    If Len(configuredSymbol) = 0 Then
        PID_GetEuroSymbol = ChrW$(8364)
    Else
        PID_GetEuroSymbol = configuredSymbol
    End If
End Function


Public Sub PID_RestoreAustrittsdatumValidation()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim updatedCount As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    monthNames = PID_MonthNames()
    updatedCount = 0
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If PID_RestoreAustrittsdatumValidationOnSheet(ws) Then
                updatedCount = updatedCount + 1
            End If
        End If
    Next i
    
    MsgBox "Austrittsdatum-" & PID_UTxtPruefung() & " wurde wiederhergestellt." & vbCrLf & vbCrLf & _
           PID_UTxtMonatsblaetter() & " aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
           "Spalte I: Datum zwischen AB1 und AB2" & vbCrLf & _
           "Jahr aus: " & PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL, _
           vbInformation, "Austrittsdatum"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_RestoreAustrittsdatumValidation:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Austrittsdatum"
End Sub


Private Function PID_RestoreAustrittsdatumValidationOnSheet(ByVal ws As Worksheet) As Boolean
    Dim wasProtected As Boolean
    Dim targetRange As Range
    Dim yearRef As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Function
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Function
    
    wasProtected = ws.ProtectContents
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    yearRef = PID_GetEinstellungYearRefR1C1()
    
    ws.Range("AB1").FormulaR1C1 = "=DATE(" & yearRef & ",R1C1,1)"
    ws.Range("AB2").FormulaR1C1 = "=EOMONTH(DATE(" & yearRef & ",R1C1,1),0)"
    
    Set targetRange = ws.Range("I" & PID_FIRST_ROW & ":I" & PID_LAST_ROW)
    
    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo SafeExit
    
    With targetRange.Validation
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=$AB$1", Formula2:="=$AB$2"
        .IgnoreBlank = True
        .InCellDropdown = False
        .ShowInput = False
        .ShowError = True
    End With
    
    PID_RestoreAustrittsdatumValidationOnSheet = True

SafeExit:
    On Error Resume Next
    If wasProtected Then
        PID_ReprotectWorksheet ws
    End If
End Function


' ===========================================================================
' FP-029 — Spalte K (Urlaub in EUR): kanonische Formel mit B/C- und J-Guard
' ===========================================================================

Public Function PID_GetUrlaubGeldFormulaR1C1() As String
    ' K = Spalte 11; Offsets: B=RC[-9], C=RC[-8], D=RC[-7], G=RC[-4], J=RC[-1]
    ' Kern: Tagessatz (G / Tage-im-Monat) * J; Monatstage via MONTH(D) analog Originalformel.
    ' Nur B/C-Guard: leere Zeile -> leer. IFERROR gibt 0 zurueck (nicht ""),
    ' da L-Spalte K arithmetisch addiert (G+K) — K="" wuerde dort zu Fehler fuehren.
    PID_GetUrlaubGeldFormulaR1C1 = _
        "=IF(OR(RC[-9]="""",RC[-8]=""""),""""," & _
        "IFERROR(IF(OR(MONTH(RC[-7])=1,MONTH(RC[-7])=3,MONTH(RC[-7])=5," & _
        "MONTH(RC[-7])=7,MONTH(RC[-7])=8,MONTH(RC[-7])=10,MONTH(RC[-7])=12)," & _
        "(RC[-4]/31)*RC[-1]," & _
        "IF(OR(MONTH(RC[-7])=4,MONTH(RC[-7])=6,MONTH(RC[-7])=9,MONTH(RC[-7])=11)," & _
        "(RC[-4]/30)*RC[-1]," & _
        "(RC[-4]/DAY(EOMONTH(RC[-7],0)))*RC[-1])),0))"
End Function


Public Function PID_RestoreUrlaubGeldFormulasSilent() As Long
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim formulaR1C1 As String
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo SafeExit

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    formulaR1C1 = PID_GetUrlaubGeldFormulaR1C1()
    monthNames = PID_MonthNames()
    PID_RestoreUrlaubGeldFormulasSilent = 0

    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        Err.Clear
        On Error GoTo SafeExit

        If Not ws Is Nothing Then
            If PID_RestoreUrlaubGeldFormulasOnSheet(ws, formulaR1C1) Then
                PID_RestoreUrlaubGeldFormulasSilent = PID_RestoreUrlaubGeldFormulasSilent + 1
            End If
        End If
    Next i

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Function


' TR-10: pro Blatt eigene Fehlerbehandlung. Vorher hat ein Fehler auf Januar die
' restlichen elf Monate uebersprungen und das Blatt entsperrt zurueckgelassen.
Private Function PID_RestoreUrlaubGeldFormulasOnSheet(ByVal ws As Worksheet, _
                                                      ByVal formulaR1C1 As String) As Boolean
    Dim wasProtected As Boolean
    Dim unprotected As Boolean

    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Function

    wasProtected = ws.ProtectContents
    If Not PID_TryUnprotectMonthSheetForMacro(ws) Then Exit Function
    unprotected = wasProtected

    ws.Range("K" & PID_FIRST_ROW & ":K" & PID_LAST_ROW).FormulaR1C1 = formulaR1C1

    On Error Resume Next
    ws.Range("K" & PID_FIRST_ROW & ":K" & PID_LAST_ROW).Calculate
    Err.Clear
    On Error GoTo SafeExit

    PID_RestoreUrlaubGeldFormulasOnSheet = True

SafeExit:
    On Error Resume Next
    If unprotected Then PID_ReprotectWorksheet ws
End Function


Public Sub PID_RestoreUrlaubGeldFormulas()
    Dim updatedCount As Long

    updatedCount = PID_RestoreUrlaubGeldFormulasSilent

    MsgBox "Urlaub-Euro-Formeln (Spalte K) wurden wiederhergestellt." & vbCrLf & vbCrLf & _
           PID_UTxtMonatsblaetter() & " aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
           "Bereich: K" & PID_FIRST_ROW & ":K" & PID_LAST_ROW, _
           vbInformation, "Spalte K"
End Sub


' Eingabespalten einer Mitarbeiterzeile: B-F, I-J und M-N.
' G (Monatslohn), H (Aktuelle Stunden), K (Urlaub Euro) und L (Letztes Gehalt) enthalten
' Formeln und gehoeren zur Struktur des Blattes - sie duerfen beim Loeschen eines
' Mitarbeiters nicht mitgeleert werden. Ihr B/C-Guard sorgt von selbst dafuer, dass die
' Zelle leer aussieht, sobald kein Mitarbeiter mehr in der Zeile steht.
Public Function PID_GetEmployeeInputCellsForRows(ByVal ws As Worksheet, _
                                                 ByVal firstRow As Long, _
                                                 ByVal lastRow As Long) As Range
    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Function

    If firstRow < PID_FIRST_ROW Then firstRow = PID_FIRST_ROW
    If lastRow > PID_LAST_ROW Then lastRow = PID_LAST_ROW
    If lastRow < firstRow Then Exit Function

    Set PID_GetEmployeeInputCellsForRows = ws.Range( _
        "B" & firstRow & ":F" & lastRow & "," & _
        "I" & firstRow & ":J" & lastRow & "," & _
        "M" & firstRow & ":N" & lastRow)

SafeExit:
End Function


' Reparaturnetz fuer Zeilen, in denen eine aeltere Version die Formeln beim Loeschen
' mitgenommen hat: fehlt in G, H, K oder L die Formel, wird sie fuer diese Zeile neu
' gesetzt. Vorhandene Formeln bleiben unangetastet. Das Blatt muss entsperrt sein.
Public Sub PID_RestoreFormulaColumnsForRows(ByVal ws As Worksheet, _
                                            ByVal firstRow As Long, _
                                            ByVal lastRow As Long)
    Dim r As Long
    Dim formulaG As String
    Dim formulaH As String
    Dim formulaK As String
    Dim formulaL As String

    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub

    If firstRow < PID_FIRST_ROW Then firstRow = PID_FIRST_ROW
    If lastRow > PID_LAST_ROW Then lastRow = PID_LAST_ROW
    If lastRow < firstRow Then Exit Sub

    formulaG = PID_GetMonatslohnFormulaR1C1()
    formulaH = PID_GetAktuelleStundenFormulaR1C1()
    formulaK = PID_GetUrlaubGeldFormulaR1C1()
    formulaL = PID_GetLetztesGehaltFormulaR1C1()

    For r = firstRow To lastRow
        PID_EnsureCellFormula ws.Cells(r, "G"), formulaG
        PID_EnsureCellFormula ws.Cells(r, "H"), formulaH
        PID_EnsureCellFormula ws.Cells(r, "K"), formulaK
        PID_EnsureCellFormula ws.Cells(r, "L"), formulaL
    Next r

SafeExit:
End Sub


Private Sub PID_EnsureCellFormula(ByVal targetCell As Range, ByVal formulaR1C1 As String)
    On Error GoTo SafeExit

    If targetCell Is Nothing Then Exit Sub
    If targetCell.HasFormula Then Exit Sub

    targetCell.FormulaR1C1 = formulaR1C1

SafeExit:
End Sub


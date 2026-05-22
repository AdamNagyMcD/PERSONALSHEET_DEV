Attribute VB_Name = "Modul1"
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


Public Sub FullSystemRefresh()
    PID_FullSystemRefresh
End Sub


Public Sub PID_FullSystemRefresh()
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Personalsheet wird aktualisiert..."
    
    PID_SetupSheetProtectionForMacros
    
    RefreshAllMonthKVStundenDropdowns
    PID_RestoreMonatslohnFormulasSilent
    PID_RestoreLetztesGehaltFormulasSilent
    PID_RestoreKVCodeDropdownValidationSilent
    ClearAllKVLohnDirty
    
    RefreshFluktuationAll
    
    PID_RecalculateAllMonthMergedFormulas
    
    PID_FormatAllMoneyColumns
    
    On Error Resume Next
    Application.CalculateFull
    On Error GoTo CleanFail
    
    MsgBox "Personalsheet wurde vollstaendig aktualisiert.", _
           vbInformation, "System Refresh"

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_FullSystemRefresh:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Personalsheet"
End Sub


Public Sub PID_QuickSystemCheck()
    Dim msg As String
    
    msg = "Personalsheet Systemcheck" & vbCrLf & vbCrLf
    
    msg = msg & "Workbook: " & ThisWorkbook.Name & vbCrLf
    msg = msg & "Excel Version: " & Application.Version & vbCrLf
    msg = msg & "Operating System: " & Application.OperatingSystem & vbCrLf & vbCrLf
    
    msg = msg & "Events aktiv: " & CStr(Application.EnableEvents) & vbCrLf
    msg = msg & "ScreenUpdating aktiv: " & CStr(Application.ScreenUpdating) & vbCrLf
    msg = msg & "Calculation: " & PID_GetCalculationModeText() & vbCrLf & vbCrLf
    
    msg = msg & "Pflichtblaetter:" & vbCrLf
    msg = msg & "- EINSTELLUNG: " & PID_YesNoText(PID_WorksheetExists(PID_EINSTELLUNG_SHEET)) & vbCrLf
    msg = msg & "- LOHNTABELLE: " & PID_YesNoText(PID_WorksheetExists(PID_LOHNTABELLE_SHEET)) & vbCrLf
    msg = msg & "- FLUKTUATION: " & PID_YesNoText(PID_WorksheetExists(PID_FLUKTUATION_SHEET)) & vbCrLf
    msg = msg & "- FLUKTUATION_DATEN: " & PID_YesNoText(PID_WorksheetExists("FLUKTUATION_DATEN")) & vbCrLf
    msg = msg & "- KV_DROPDOWN_HELPER: " & PID_YesNoText(PID_WorksheetExists("KV_DROPDOWN_HELPER")) & vbCrLf & vbCrLf
    
    msg = msg & "Monatsblaetter gefunden: " & CStr(PID_CountMonthSheets()) & " / 12" & vbCrLf
    
    MsgBox msg, vbInformation, "Personalsheet Systemcheck"
End Sub


Public Sub PID_RecalculateAllMonthFluctuation()
    RefreshFluktuationAll
End Sub


Public Sub RecalculateFinanzSummaryChain()
    PID_SyncFinanzSummaryToUbersicht
End Sub


Public Sub SyncFinanzSummaryToUbersicht()
    PID_SyncFinanzSummaryToUbersicht
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
            
            ws.Columns("G").ColumnWidth = 13
            ws.Columns("J").ColumnWidth = 13
            ws.Columns("K").ColumnWidth = 14
            
            ws.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
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


Public Sub FormatAllLohnColumns()
    PID_FormatAllMoneyColumns
    
    MsgBox "Lohn- und Geldspalten wurden formatiert.", _
           vbInformation, "Formatierung"
End Sub


Public Sub PID_ResetExcelState()
    On Error Resume Next
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    
    On Error GoTo 0
    
    MsgBox "Excel wurde zurueckgesetzt.", _
           vbInformation, "Excel Reset"
End Sub


Public Sub PID_EventReset()
    On Error Resume Next
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.CutCopyMode = False
    
    On Error GoTo 0
End Sub


Public Sub PID_HideTechnicalSheets()
    On Error Resume Next
    
    ThisWorkbook.Worksheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("KV_DROPDOWN_HELPER").Visible = xlSheetVeryHidden
    
    On Error GoTo 0
    
    MsgBox "Technische Blaetter wurden ausgeblendet.", _
           vbInformation, "Technische Blaetter"
End Sub


Public Sub PID_ShowTechnicalSheets()
    On Error Resume Next
    
    ThisWorkbook.Worksheets("FLUKTUATION_DATEN").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("KV_DROPDOWN_HELPER").Visible = xlSheetVisible
    
    On Error GoTo 0
    
    MsgBox "Technische Blaetter wurden sichtbar gemacht.", _
           vbInformation, "Technische Blaetter"
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


Public Sub PID_ForceRefreshActiveRowLohn()
    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim monthNumber As Long
    
    On Error GoTo SafeExit
    
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Sub
    
    rowNumber = ActiveCell.Row
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Sub
    
    monthNumber = CLng(ws.Range("A1").Value)
    RefreshKVLohnForRow ws, rowNumber, monthNumber
    
SafeExit:
End Sub


Public Sub RestoreAktuelleStundenFormulas()
    PID_RestoreAktuelleStundenFormulas
End Sub


Public Sub PID_RestoreAktuelleStundenFormulas()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim updatedCount As Long
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
    updatedCount = 0
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If PID_RestoreAktuelleStundenFormulasOnSheet(ws, formulaR1C1) Then
                updatedCount = updatedCount + 1
            End If
        End If
    Next i
    
    MsgBox "Aktuelle-Stunden-Formeln wurden wiederhergestellt." & vbCrLf & vbCrLf & _
           "Monatsblaetter aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
           "Bereich: H" & PID_FIRST_ROW & ":H" & PID_LAST_ROW & vbCrLf & _
           "Jahr aus: " & PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL, _
           vbInformation, "Aktuelle Stunden"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_RestoreAktuelleStundenFormulas:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Aktuelle Stunden"
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
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
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


Public Sub RestoreLetztesGehaltFormulas()
    PID_RestoreLetztesGehaltFormulas
End Sub


Public Sub RestoreMonatslohnFormulas()
    PID_RestoreMonatslohnFormulas
End Sub


Public Sub RestoreKVCodeDropdownValidation()
    PID_RestoreKVCodeDropdownValidation
End Sub


Public Sub RestoreKVStundenDropdownValidation()
    RefreshAllMonthKVStundenDropdowns
    MarkKVDropdownsClean
    
    MsgBox "Stunden-Dropdown (Spalte F) wurde wiederhergestellt." & vbCrLf & vbCrLf & _
           "Bereich: F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW & " auf allen Monatsblaettern.", _
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
           "Monatsblaetter aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
           "Bereich: L" & PID_FIRST_ROW & ":L" & PID_LAST_ROW & vbCrLf & _
           "Jahr aus: " & PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL & vbCrLf & vbCrLf & _
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
    PID_RestoreLetztesGehaltFormulasSilent
    
    On Error Resume Next
    Application.CalculateFull
    On Error GoTo 0
End Sub


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
    
    On Error GoTo SafeExit
    
    On Error Resume Next
    Application.CalculateFull
    Err.Clear
    On Error GoTo SafeExit

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
    Dim euroSymbol As String
    Dim sampleFormula As String
    Dim fixedFormula As String
    Dim sourceCell As Range
    Dim targetRange As Range
    Dim yearRefA1 As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheetName(ws.Name) Then Exit Function
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Function
    
    wasProtected = ws.ProtectContents
    
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If ws.ProtectContents Then Exit Function
    
    If Trim$(CStr(ws.Range("L1").Value)) = "" Then
        ws.Range("L1").Value = "Letztes Gehalt"
    End If
    
    Set sourceCell = ws.Range("L" & PID_FIRST_ROW)
    Set targetRange = ws.Range("L" & PID_FIRST_ROW & ":L" & PID_LAST_ROW)
    yearRefA1 = PID_GetEinstellungYearRefA1()
    
    sampleFormula = CStr(sourceCell.Formula)
    
    If InStr(1, sampleFormula, "#REF", vbTextCompare) > 0 Then
        fixedFormula = Replace(sampleFormula, "#REF!", yearRefA1)
        fixedFormula = Replace(fixedFormula, "#REF", yearRefA1)
        sourceCell.Formula = fixedFormula
        
        If PID_LAST_ROW > PID_FIRST_ROW Then
            sourceCell.AutoFill Destination:=targetRange
        End If
    ElseIf Not PID_MonthSheetHasLetztesGehaltFormula(ws) Then
        targetRange.FormulaR1C1 = formulaR1C1
    End If
    
    euroSymbol = PID_GetEuroSymbol()
    targetRange.NumberFormat = euroSymbol & " #,##0.00"
    
    PID_RestoreLetztesGehaltFormulasOnSheet = Not PID_MonthSheetHasLetztesGehaltRefError(ws)

SafeExit:
    On Error Resume Next
    If Not ws Is Nothing Then
        If wasProtected Then
            ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                       UserInterfaceOnly:=True, _
                       AllowFiltering:=True, _
                       AllowSorting:=True
        End If
    End If
End Function


Public Function PID_GetLetztesGehaltFormulaR1C1() As String
    Dim yearRef As String
    
    yearRef = PID_GetEinstellungYearRefR1C1()
    
    ' Spalte L = Letztes Gehalt / Laborcost fuer AVG Bruttolohn (Q42 = Q17/Q8).
    ' Jahr aus EINSTELLUNG!C35 statt legacy LOHNTABELLE!G3 (#REF!).
    PID_GetLetztesGehaltFormulaR1C1 = _
        "=IFERROR(IF(OR(ISBLANK(RC[-8]),ISBLANK(RC[-6]),ISBLANK(RC[-5]),ISBLANK(RC[-1])),0," & _
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


Public Sub RestoreAustrittsdatumValidation()
    PID_RestoreAustrittsdatumValidation
End Sub


Public Sub BuildDurchrechnungUebersicht()
    PID_BuildDurchrechnungUebersicht
End Sub


Public Sub RefreshDurchrechnungUebersicht()
    PID_RefreshDurchrechnungUebersicht
End Sub


Public Sub FormatDurchrechnungUebersicht()
    PID_FormatDurchrechnungUebersicht
End Sub


Public Sub FormatFinanzUebersicht()
    PID_FormatFinanzUebersicht
End Sub


Public Sub SyncDieseArbeitsmappeFromExport()
    Dim syncOk As Boolean
    Dim syncDetails As String
    
    PID_SyncDieseArbeitsmappeFromExport syncOk, syncDetails
    
    If syncOk Then
        MsgBox "DieseArbeitsmappe wurde aus vba/DieseArbeitsmappe.cls synchronisiert.", vbInformation, "VBA Sync"
    Else
        MsgBox "Synchronisation fehlgeschlagen:" & vbCrLf & syncDetails, vbExclamation, "VBA Sync"
    End If
End Sub


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
    
    MsgBox "Austrittsdatum-Pruefung wurde wiederhergestellt." & vbCrLf & vbCrLf & _
           "Monatsblaetter aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
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
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
    End If
End Function


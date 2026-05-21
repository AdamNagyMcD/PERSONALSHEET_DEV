Attribute VB_Name = "Modul1"
Option Explicit

Public Const PID_FIRST_ROW As Long = 3
Public Const PID_LAST_ROW As Long = 82
Public Const PID_WORKBOOK_PASSWORD As String = "company"
Public Const PID_EINSTELLUNG_SHEET As String = "EINSTELLUNG"
Public Const PID_LOHNTABELLE_SHEET As String = "LOHNTABELLE"
Public Const PID_WORKBOOK_YEAR_CELL As String = "C35"
Public Const PID_FLUKTUATION_REASON_FIRST_ROW As Long = 38
Public Const PID_FLUKTUATION_REASON_LAST_ROW As Long = 49
Public Const PID_FLUKTUATION_TIME_FIRST_ROW As Long = 53
Public Const PID_FLUKTUATION_TIME_LAST_ROW As Long = 59


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
    RefreshAllMonthKVLohn
    PID_RecalculateAllMonthFluctuation
    
    MarkFluktuationDirty
    RefreshFluktuationAll
    
    PID_FormatAllMoneyColumns
    
    PID_SetupSheetProtectionForMacros
    
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
    msg = msg & "- Fluktuation: " & PID_YesNoText(PID_WorksheetExists("Fluktuation")) & vbCrLf
    msg = msg & "- FLUKTUATION_DATEN: " & PID_YesNoText(PID_WorksheetExists("FLUKTUATION_DATEN")) & vbCrLf
    msg = msg & "- KV_DROPDOWN_HELPER: " & PID_YesNoText(PID_WorksheetExists("KV_DROPDOWN_HELPER")) & vbCrLf & vbCrLf
    
    msg = msg & "Monatsblaetter gefunden: " & CStr(PID_CountMonthSheets()) & " / 12" & vbCrLf
    
    MsgBox msg, vbInformation, "Personalsheet Systemcheck"
End Sub


Public Sub PID_RecalculateAllMonthFluctuation()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    
    On Error GoTo CleanFail
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            PID_CalculateFluctuation ws
        End If
    Next i
    
    Exit Sub

CleanFail:
    MsgBox "Fehler bei PID_RecalculateAllMonthFluctuation:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Fluktuation"
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


Private Function PID_GetAktuelleStundenFormulaR1C1() As String
    Dim yearRef As String
    
    ' PID_WORKBOOK_YEAR_CELL = C35 -> R35C3 in R1C1 notation.
    yearRef = "'" & PID_EINSTELLUNG_SHEET & "'!R35C3"
    
    PID_GetAktuelleStundenFormulaR1C1 = _
        "=IF(AND(ISNUMBER(RC[-4]),ISNUMBER(RC[-2]))," & _
        "ROUNDDOWN(MIN(" & _
        "IF(OR(ISBLANK(RC[1]),RC[1]="""")," & _
        "EOMONTH(DATE(" & yearRef & ",R1C1,1),0)," & _
        "RC[1])-" & _
        "MAX(RC[-4],DATE(" & yearRef & ",R1C1,1))+1)/" & _
        "DAY(EOMONTH(DATE(" & yearRef & ",R1C1,1),0))*RC[-2],2),"""")"
End Function


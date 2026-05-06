Attribute VB_Name = "Modul1"
Option Explicit

Private Const PID_PASSWORD As String = "company"


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
    msg = msg & "- LOHNTABELLE: " & PID_YesNoText(PID_WorksheetExists("LOHNTABELLE")) & vbCrLf
    msg = msg & "- LOHNTABELLE_TEST: " & PID_YesNoText(PID_WorksheetExists("LOHNTABELLE_TEST")) & vbCrLf
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
            ws.Unprotect Password:=PID_PASSWORD
            On Error GoTo CleanFail
            
            PID_ApplyEuroNumberFormat ws.Range("G3:G82")
            PID_ApplyEuroNumberFormat ws.Range("J3:J82")
            PID_ApplyEuroNumberFormat ws.Range("K3:K82")
            
            ws.Columns("G").ColumnWidth = 13
            ws.Columns("J").ColumnWidth = 13
            ws.Columns("K").ColumnWidth = 14
            
            ws.Protect Password:=PID_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
        End If
    Next i
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    On Error GoTo CleanFail
    
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Unprotect Password:=PID_PASSWORD
        On Error GoTo CleanFail
        
        PID_ApplyEuroNumberFormat ws.Range("H:H")
        ws.Columns("H").ColumnWidth = 14
        
        ws.Protect Password:=PID_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOHNTABELLE")
    On Error GoTo CleanFail
    
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Unprotect Password:=PID_PASSWORD
        On Error GoTo CleanFail
        
        PID_ApplyEuroNumberFormat ws.Range("H:H")
        ws.Columns("H").ColumnWidth = 14
        
        ws.Protect Password:=PID_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
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


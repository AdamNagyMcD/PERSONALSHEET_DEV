Attribute VB_Name = "mod_AddNewKVPeriodOnTop"
Option Explicit
Private Const PID_ADD_PERIOD_BUTTON_NAME As String = "btn_AddNewKVPeriodOnTop"

Public Sub AddNewKVPeriodOnTop()
    Dim wsKV As Worksheet
    Dim newPeriod As String
    Dim templatePeriod As String
    Dim templateStartYear As Long
    Dim newStartYear As Long
    Dim currentSchemaCount As Long
    Dim newSchemaCount As Long
    Dim newPeriodData As Variant
    Dim newRowCount As Long
    
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim templateFirstRow As Long
    Dim templateLastRow As Long
    Dim templateRowCount As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    Dim answer As VbMsgBoxResult
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    
    firstDataRow = 4
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    If lastRow < firstDataRow Then
        MsgBox "Keine KV-Daten in LOHNTABELLE_TEST gefunden.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templatePeriod = GetTopKVPeriod(wsKV, firstDataRow)
    
    If templatePeriod = "" Then
        MsgBox "Es wurde kein bestehender KV-Zeitraum als Vorlage gefunden.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templateStartYear = GetStartYearFromKVPeriod(templatePeriod)
    If templateStartYear = 0 Then templateStartYear = Year(Date)
    
    newStartYear = AskForKVStartYear(templateStartYear + 1)
    If newStartYear = 0 Then GoTo CleanExit
    
    newPeriod = BuildKVPeriodName(newStartYear)
    
    If KVPeriodExists(wsKV, newPeriod, firstDataRow) Then
        MsgBox "Dieser KV-Zeitraum existiert bereits:" & vbCrLf & vbCrLf & _
               newPeriod, _
               vbInformation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templateFirstRow = FindFirstRowOfPeriod(wsKV, templatePeriod, firstDataRow)
    templateLastRow = FindLastRowOfPeriod(wsKV, templatePeriod, firstDataRow)
    
    If templateFirstRow = 0 Or templateLastRow = 0 Then
        MsgBox "Die Vorlage fuer den neuen KV-Zeitraum konnte nicht gefunden werden.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templateRowCount = templateLastRow - templateFirstRow + 1
    
    If templateRowCount <= 0 Then
        MsgBox "Die Vorlage enthaelt keine gueltigen Datenzeilen.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    currentSchemaCount = GetSchemaCountFromPeriod(wsKV, templateFirstRow, templateLastRow)
    If currentSchemaCount <= 0 Then currentSchemaCount = 13
    
    newSchemaCount = AskForSchemaCount(currentSchemaCount)
    If newSchemaCount <= 0 Then GoTo CleanExit
    
    newPeriodData = BuildNewPeriodDataFromTemplate(wsKV, templateFirstRow, templateLastRow, newPeriod, newStartYear, newSchemaCount)
    If Not IsArray(newPeriodData) Then
        MsgBox "Die neue KV-Vorlage konnte nicht erstellt werden.", vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    newRowCount = UBound(newPeriodData, 1)
    If newRowCount <= 0 Then GoTo CleanExit
    
    answer = MsgBox( _
        "Neuen KV-Zeitraum erstellen?" & vbCrLf & vbCrLf & _
        "Zeitraum: " & newPeriod & vbCrLf & _
        "Vertraege pro KV-Code: " & newSchemaCount, _
        vbQuestion + vbYesNo, _
        "Neuer KV-Zeitraum" _
    )
    
    If answer <> vbYes Then GoTo CleanExit
    
    InsertNewKVPeriodRows wsKV, firstDataRow, newPeriodData
    
    PID_NormalizeKVTableHeader wsKV
    PID_NormalizeKVWarningText wsKV
    FormatKVPeriodArea wsKV
    EnsureAddNewKVPeriodButton
    
    MarkKVDropdownsDirty
    
    MsgBox "Der neue KV-Zeitraum wurde erfolgreich oben eingefuegt:" & vbCrLf & vbCrLf & _
           newPeriod & vbCrLf & vbCrLf & _
           "Bitte jetzt die Lohnwerte kontrollieren und bei Bedarf anpassen.", _
           vbInformation, "Neuer KV-Zeitraum"

CleanExit:
    On Error Resume Next
    wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    On Error GoTo 0
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not wsKV Is Nothing Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei AddNewKVPeriodOnTop:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Neuer KV-Zeitraum"
End Sub


Public Sub EnsureAddNewKVPeriodButton()
    Dim wsKV As Worksheet
    Dim btn As Shape
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    wsKV.Shapes(PID_ADD_PERIOD_BUTTON_NAME).Delete
    On Error GoTo SafeExit
    
    Set btn = wsKV.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                   Left:=wsKV.Range("J1").Left, _
                                   Top:=wsKV.Range("J1").Top + 2, _
                                   Width:=220, _
                                   Height:=22)
    
    btn.Name = PID_ADD_PERIOD_BUTTON_NAME
    btn.TextFrame.Characters.Text = "Neuen KV-Zeitraum oben einfuegen"
    btn.OnAction = "AddNewKVPeriodOnTop"
    
    btn.Fill.ForeColor.RGB = RGB(54, 96, 146)
    btn.Line.ForeColor.RGB = RGB(33, 64, 99)
    btn.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    btn.TextFrame.Characters.Font.Bold = True
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Public Sub FixLOHNTABELLE_TEST_HeaderText()
    Dim wsKV As Worksheet
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    PID_NormalizeKVWarningText wsKV
    PID_NormalizeKVTableHeader wsKV
    wsKV.Range("A2").WrapText = True
    
    ' Status/Pruefung duerfen keine statischen OK-Werte sein.
    PID_EnsureKVStatusFormulas wsKV
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Public Sub FixLOHNTABELLE_TEST_StatusFormulas()
    Dim wsKV As Worksheet
    Dim wasProtected As Boolean
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    FormatKVPeriodArea wsKV
    
    MsgBox "Status- und Pruefungsformeln in LOHNTABELLE_TEST wurden wiederhergestellt.", _
           vbInformation, "LOHNTABELLE_TEST"
    
CleanExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei FixLOHNTABELLE_TEST_StatusFormulas:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "LOHNTABELLE_TEST"
    Resume CleanExit
End Sub


Public Sub RebuildLOHNTABELLE_TEST()
    Dim wsKV As Worksheet
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim cleanupLastRow As Long
    Dim keepPeriod As String
    Dim periodFirstRow As Long
    Dim periodLastRow As Long
    Dim periodRowCount As Long
    Dim periodData As Variant
    Dim answer As VbMsgBoxResult
    Dim wasProtected As Boolean
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    firstDataRow = 4
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    cleanupLastRow = PID_GetSheetCleanupLastRow(wsKV, lastRow)
    
    If lastRow < firstDataRow Then
        MsgBox "Keine Daten in LOHNTABELLE_TEST gefunden.", vbExclamation, "LOHNTABELLE_TEST neu aufbauen"
        Exit Sub
    End If
    
    keepPeriod = GetBottomKVPeriod(wsKV, firstDataRow)
    If keepPeriod = "" Then
        MsgBox "Kein gueltiger KV-Zeitraum in Spalte A gefunden.", vbExclamation, "LOHNTABELLE_TEST neu aufbauen"
        Exit Sub
    End If
    
    periodFirstRow = FindFirstRowOfPeriod(wsKV, keepPeriod, firstDataRow)
    periodLastRow = FindLastRowOfPeriod(wsKV, keepPeriod, firstDataRow)
    
    If periodFirstRow = 0 Or periodLastRow = 0 Then
        MsgBox "Der unterste KV-Zeitraum konnte nicht gelesen werden.", vbExclamation, "LOHNTABELLE_TEST neu aufbauen"
        Exit Sub
    End If
    
    periodRowCount = periodLastRow - periodFirstRow + 1
    If periodRowCount <= 0 Then Exit Sub
    
    answer = MsgBox( _
        "LOHNTABELLE_TEST wird neu aufgebaut." & vbCrLf & vbCrLf & _
        "Behalten wird nur der unterste Zeitraum:" & vbCrLf & _
        keepPeriod & vbCrLf & vbCrLf & _
        "Alle weiteren Test-Zeitraeume werden geloescht." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        "LOHNTABELLE_TEST neu aufbauen" _
    )
    
    If answer <> vbYes Then Exit Sub
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    wasProtected = wsKV.ProtectContents
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    periodData = wsKV.Range("A" & periodFirstRow & ":I" & periodLastRow).Value
    
    PID_ClearKVDataArea wsKV, firstDataRow, cleanupLastRow
    wsKV.Range("A" & firstDataRow + 1 & ":I" & firstDataRow + periodRowCount).Value = periodData
    PID_WriteKVPeriodTitleRow wsKV, firstDataRow, keepPeriod, periodData(1, 2), periodData(1, 3)
    PID_ClearTrailingKVArea wsKV, firstDataRow + periodRowCount + 1, cleanupLastRow
    PID_NormalizeKVWarningText wsKV
    PID_NormalizeKVTableHeader wsKV
    FormatKVPeriodArea wsKV
    MarkKVDropdownsDirty
    
    On Error Resume Next
    PID_ResetHourOverrideLog
    On Error GoTo CleanFail
    
    EnsureAddNewKVPeriodButton
    
    MsgBox "LOHNTABELLE_TEST wurde neu aufgebaut. Zeitraum aktiv: " & keepPeriod, _
           vbInformation, "LOHNTABELLE_TEST neu aufgebaut"
    
CleanExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    On Error GoTo 0
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei RebuildLOHNTABELLE_TEST:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "LOHNTABELLE_TEST neu aufbauen"
    Resume CleanExit
End Sub


Public Sub RestoreLOHNTABELLE_TESTBase2025_2026()
    Dim wsKV As Worksheet
    Dim targetPeriod As String
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim cleanupLastRow As Long
    Dim periodFirstRow As Long
    Dim periodLastRow As Long
    Dim periodRowCount As Long
    Dim periodData As Variant
    Dim answer As VbMsgBoxResult
    Dim wasProtected As Boolean
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    targetPeriod = "KV 2025/2026"
    firstDataRow = 4
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    cleanupLastRow = PID_GetSheetCleanupLastRow(wsKV, lastRow)
    
    periodFirstRow = FindFirstRowOfPeriod(wsKV, targetPeriod, firstDataRow)
    periodLastRow = FindLastRowOfPeriod(wsKV, targetPeriod, firstDataRow)
    
    If periodFirstRow = 0 Or periodLastRow = 0 Then
        MsgBox "Der Basiszeitraum '" & targetPeriod & "' wurde nicht gefunden.", _
               vbExclamation, "Basis wiederherstellen"
        Exit Sub
    End If
    
    periodRowCount = periodLastRow - periodFirstRow + 1
    If periodRowCount <= 0 Then Exit Sub
    
    answer = MsgBox( _
        "LOHNTABELLE_TEST wird auf den Basiszeitraum zurueckgesetzt:" & vbCrLf & vbCrLf & _
        targetPeriod & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        "Basis wiederherstellen" _
    )
    
    If answer <> vbYes Then Exit Sub
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    wasProtected = wsKV.ProtectContents
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    periodData = wsKV.Range("A" & periodFirstRow & ":I" & periodLastRow).Value
    periodData = PID_FilterValidKVRows(periodData, targetPeriod)
    periodData = PID_EnsureBG1Basis173Row(periodData, targetPeriod)
    periodRowCount = UBound(periodData, 1)
    
    PID_ClearKVDataArea wsKV, firstDataRow, cleanupLastRow
    wsKV.Range("A" & firstDataRow + 1 & ":I" & firstDataRow + periodRowCount).Value = periodData
    PID_WriteKVPeriodTitleRow wsKV, firstDataRow, targetPeriod, periodData(1, 2), periodData(1, 3)
    PID_ClearTrailingKVArea wsKV, firstDataRow + periodRowCount + 1, cleanupLastRow
    
    PID_NormalizeKVWarningText wsKV
    PID_NormalizeKVTableHeader wsKV
    FormatKVPeriodArea wsKV
    EnsureAddNewKVPeriodButton
    
    On Error Resume Next
    PID_ResetHourOverrideLog
    On Error GoTo CleanFail
    
    MsgBox "Basiszeitraum wurde wiederhergestellt: " & targetPeriod, _
           vbInformation, "Basis wiederherstellen"
    
CleanExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    On Error GoTo 0
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei RestoreLOHNTABELLE_TESTBase2025_2026:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Basis wiederherstellen"
    Resume CleanExit
End Sub


Private Function PID_FilterValidKVRows(ByVal sourceData As Variant, ByVal periodName As String) As Variant
    Dim r As Long
    Dim outRow As Long
    Dim validCount As Long
    Dim resultData As Variant
    
    On Error GoTo Fallback
    
    validCount = 0
    
    For r = 1 To UBound(sourceData, 1)
        If PID_IsValidKVDataRow(sourceData, r) Then
            validCount = validCount + 1
        End If
    Next r
    
    If validCount = 0 Then GoTo Fallback
    
    ReDim resultData(1 To validCount, 1 To 9)
    outRow = 0
    
    For r = 1 To UBound(sourceData, 1)
        If PID_IsValidKVDataRow(sourceData, r) Then
            outRow = outRow + 1
            
            resultData(outRow, 1) = periodName
            resultData(outRow, 2) = sourceData(r, 2)
            resultData(outRow, 3) = sourceData(r, 3)
            resultData(outRow, 4) = PID_GetNearestTextValue(sourceData, r, 4)
            resultData(outRow, 5) = PID_GetNearestTextValue(sourceData, r, 5)
            resultData(outRow, 6) = PID_GetNearestTextValue(sourceData, r, 6)
            resultData(outRow, 7) = sourceData(r, 7)
            resultData(outRow, 8) = sourceData(r, 8)
            resultData(outRow, 9) = ""
        End If
    Next r
    
    PID_FilterValidKVRows = resultData
    Exit Function
    
Fallback:
    PID_FilterValidKVRows = sourceData
End Function


Private Function PID_IsValidKVDataRow(ByVal sourceData As Variant, ByVal rowIndex As Long) As Boolean
    If rowIndex < 1 Or rowIndex > UBound(sourceData, 1) Then Exit Function
    
    If Trim$(CStr(PID_GetNearestTextValue(sourceData, rowIndex, 4))) = "" Then Exit Function
    If Trim$(CStr(PID_GetNearestTextValue(sourceData, rowIndex, 5))) = "" Then Exit Function
    If Trim$(CStr(PID_GetNearestTextValue(sourceData, rowIndex, 6))) = "" Then Exit Function
    
    PID_IsValidKVDataRow = True
End Function


Private Function PID_EnsureBG1Basis173Row(ByVal sourceData As Variant, ByVal periodName As String) As Variant
    Dim r As Long
    Dim outRow As Long
    Dim has173 As Boolean
    Dim first151Row As Long
    Dim hoursValue As Double
    Dim resultData As Variant
    
    On Error GoTo Fallback
    
    For r = 1 To UBound(sourceData, 1)
        If UCase$(Trim$(CStr(sourceData(r, 4)))) = "BG1_BASIS" Then
            If PID_TryReadDouble(sourceData(r, 7), hoursValue) Then
                If Abs(hoursValue - 173#) < 0.01 Then
                    has173 = True
                End If
                
                If first151Row = 0 Then
                    If Abs(hoursValue - 151.38) < 0.01 Then
                        first151Row = r
                    End If
                End If
            End If
        End If
    Next r
    
    If has173 Then
        PID_EnsureBG1Basis173Row = sourceData
        Exit Function
    End If
    
    If first151Row = 0 Then
        PID_EnsureBG1Basis173Row = sourceData
        Exit Function
    End If
    
    ReDim resultData(1 To UBound(sourceData, 1) + 1, 1 To 9)
    
    outRow = 0
    For r = 1 To UBound(sourceData, 1)
        outRow = outRow + 1
        
        If outRow = first151Row Then
            resultData(outRow, 1) = periodName
            resultData(outRow, 2) = sourceData(first151Row, 2)
            resultData(outRow, 3) = sourceData(first151Row, 3)
            resultData(outRow, 4) = "BG1_Basis"
            resultData(outRow, 5) = PID_GetNearestTextValue(sourceData, first151Row, 5)
            resultData(outRow, 6) = PID_GetNearestTextValue(sourceData, first151Row, 6)
            resultData(outRow, 7) = 173#
            resultData(outRow, 8) = 2021#
            resultData(outRow, 9) = ""
            outRow = outRow + 1
        End If
        
        resultData(outRow, 1) = sourceData(r, 1)
        resultData(outRow, 2) = sourceData(r, 2)
        resultData(outRow, 3) = sourceData(r, 3)
        resultData(outRow, 4) = sourceData(r, 4)
        resultData(outRow, 5) = sourceData(r, 5)
        resultData(outRow, 6) = sourceData(r, 6)
        resultData(outRow, 7) = sourceData(r, 7)
        resultData(outRow, 8) = sourceData(r, 8)
        resultData(outRow, 9) = sourceData(r, 9)
    Next r
    
    PID_EnsureBG1Basis173Row = resultData
    Exit Function
    
Fallback:
    PID_EnsureBG1Basis173Row = sourceData
End Function


Private Function PID_TryReadDouble(ByVal valueToRead As Variant, ByRef resultValue As Double) As Boolean
    Dim s As String
    
    On Error GoTo SafeExit
    
    If IsNumeric(valueToRead) Then
        resultValue = CDbl(valueToRead)
        PID_TryReadDouble = True
        Exit Function
    End If
    
    s = Trim$(CStr(valueToRead))
    If s = "" Then GoTo SafeExit
    
    s = Replace(s, ".", ",")
    
    If IsNumeric(s) Then
        resultValue = CDbl(s)
        PID_TryReadDouble = True
    End If
    
SafeExit:
End Function


Private Function PID_GetNearestTextValue(ByVal sourceData As Variant, _
                                         ByVal rowIndex As Long, _
                                         ByVal colIndex As Long) As String
    Dim r As Long
    Dim valueText As String
    
    If rowIndex < 1 Or rowIndex > UBound(sourceData, 1) Then Exit Function
    
    valueText = Trim$(CStr(sourceData(rowIndex, colIndex)))
    If valueText <> "" Then
        PID_GetNearestTextValue = valueText
        Exit Function
    End If
    
    For r = rowIndex - 1 To 1 Step -1
        valueText = Trim$(CStr(sourceData(r, colIndex)))
        If valueText <> "" Then
            PID_GetNearestTextValue = valueText
            Exit Function
        End If
    Next r
    
    For r = rowIndex + 1 To UBound(sourceData, 1)
        valueText = Trim$(CStr(sourceData(r, colIndex)))
        If valueText <> "" Then
            PID_GetNearestTextValue = valueText
            Exit Function
        End If
    Next r
End Function


Private Function AskForKVStartYear(ByVal defaultYear As Long) As Long
    Dim inputText As String
    
    inputText = InputBox( _
        Prompt:="Bitte Startjahr fuer den neuen KV-Zeitraum eingeben." & vbCrLf & vbCrLf & _
                "Beispiel: 2026 -> KV 2026/2027", _
        Title:="Neuen KV-Zeitraum hinzufuegen", _
        Default:=CStr(defaultYear) _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Ungueltige Eingabe. Bitte nur das Startjahr eingeben (z.B. 2026).", _
               vbExclamation, "Neuer KV-Zeitraum"
        Exit Function
    End If
    
    AskForKVStartYear = CLng(inputText)
    
    If AskForKVStartYear < 2000 Or AskForKVStartYear > 2100 Then
        MsgBox "Das Startjahr liegt ausserhalb des erlaubten Bereichs (2000-2100).", _
               vbExclamation, "Neuer KV-Zeitraum"
        AskForKVStartYear = 0
    End If
End Function


Private Function AskForSchemaCount(ByVal defaultCount As Long) As Long
    Dim inputText As String
    
    inputText = InputBox( _
        Prompt:="Wie viele Vertraege/Monatsstunden-Zeilen pro KV-Code sollen erzeugt werden?" & vbCrLf & vbCrLf & _
                "Beispiel: 13 (wie bisher).", _
        Title:="Vertragsanzahl pro KV-Code", _
        Default:=CStr(defaultCount) _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Ungueltige Eingabe. Bitte eine ganze Zahl eingeben (z.B. 13).", _
               vbExclamation, "Vertragsanzahl"
        Exit Function
    End If
    
    AskForSchemaCount = CLng(inputText)
    
    If AskForSchemaCount < 1 Or AskForSchemaCount > 50 Then
        MsgBox "Die Vertragsanzahl muss zwischen 1 und 50 liegen.", _
               vbExclamation, "Vertragsanzahl"
        AskForSchemaCount = 0
    End If
End Function


Private Function BuildKVPeriodName(ByVal startYear As Long) As String
    BuildKVPeriodName = "KV " & CStr(startYear) & "/" & CStr(startYear + 1)
End Function


Private Function GetStartYearFromKVPeriod(ByVal periodName As String) As Long
    Dim s As String
    Dim parts As Variant
    
    s = NormalizeKVPeriodText(periodName)
    If Left$(s, 3) <> "KV " Then Exit Function
    
    s = Mid$(s, 4)
    parts = Split(s, "/")
    
    If UBound(parts) <> 1 Then Exit Function
    If Not IsNumeric(parts(0)) Then Exit Function
    
    GetStartYearFromKVPeriod = CLng(parts(0))
End Function


Private Function GetSchemaCountFromPeriod(ByVal wsKV As Worksheet, _
                                          ByVal periodFirstRow As Long, _
                                          ByVal periodLastRow As Long) As Long
    Dim firstCode As String
    Dim firstDataRowInPeriod As Long
    Dim r As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    If periodFirstRow <= 0 Or periodLastRow < periodFirstRow Then Exit Function
    
    firstDataRowInPeriod = 0
    
    For r = periodFirstRow To periodLastRow
        If Not wsKV.Range("A" & r).MergeCells Then
            firstCode = Trim$(CStr(wsKV.Cells(r, "D").Value))
            If firstCode <> "" Then
                firstDataRowInPeriod = r
                Exit For
            End If
        End If
    Next r
    
    If firstDataRowInPeriod = 0 Then Exit Function
    
    firstCode = Trim$(CStr(wsKV.Cells(firstDataRowInPeriod, "D").Value))
    
    For r = firstDataRowInPeriod To periodLastRow
        If Trim$(CStr(wsKV.Cells(r, "D").Value)) = firstCode Then
            GetSchemaCountFromPeriod = GetSchemaCountFromPeriod + 1
        Else
            Exit Function
        End If
    Next r
    
SafeExit:
End Function


Private Function BuildNewPeriodDataFromTemplate(ByVal wsKV As Worksheet, _
                                                ByVal templateFirstRow As Long, _
                                                ByVal templateLastRow As Long, _
                                                ByVal newPeriod As String, _
                                                ByVal newStartYear As Long, _
                                                ByVal newSchemaCount As Long) As Variant
    Dim templateData As Variant
    Dim blockStarts As Collection
    Dim blockEnds As Collection
    Dim totalRows As Long
    Dim r As Long
    Dim outRow As Long
    Dim blockIndex As Long
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim blockRows As Long
    Dim schemaIndex As Long
    Dim sourceOffset As Long
    Dim validFrom As Date
    Dim validTo As Date
    Dim resultData As Variant
    Dim currentCode As String
    Dim rowCode As String
    
    On Error GoTo BuildFail
    
    If wsKV Is Nothing Then Exit Function
    If templateLastRow < templateFirstRow Then Exit Function
    If newSchemaCount <= 0 Then Exit Function
    
    templateData = wsKV.Range("A" & templateFirstRow & ":I" & templateLastRow).Value
    
    Set blockStarts = New Collection
    Set blockEnds = New Collection
    currentCode = ""
    
    ' Build blocks only from real KV codes (ignore empty separator/title rows).
    For r = 1 To UBound(templateData, 1)
        rowCode = Trim$(CStr(templateData(r, 4)))
        
        If rowCode <> "" Then
            If currentCode = "" Then
                blockStarts.Add r
                currentCode = rowCode
            ElseIf rowCode <> currentCode Then
                blockEnds.Add r - 1
                blockStarts.Add r
                currentCode = rowCode
            End If
        Else
            If currentCode <> "" Then
                blockEnds.Add r - 1
                currentCode = ""
            End If
        End If
    Next r
    
    If currentCode <> "" Then
        blockEnds.Add UBound(templateData, 1)
    End If
    
    If blockStarts.Count = 0 Or blockEnds.Count = 0 Then GoTo BuildFail
    
    totalRows = blockStarts.Count * newSchemaCount
    ReDim resultData(1 To totalRows, 1 To 9)
    
    validFrom = DateSerial(newStartYear, 5, 1)
    validTo = DateSerial(newStartYear + 1, 4, 30)
    
    outRow = 0
    
    For blockIndex = 1 To blockStarts.Count
        blockStart = CLng(blockStarts(blockIndex))
        blockEnd = CLng(blockEnds(blockIndex))
        blockRows = blockEnd - blockStart + 1
        
        For schemaIndex = 1 To newSchemaCount
            outRow = outRow + 1
            
            resultData(outRow, 1) = newPeriod
            resultData(outRow, 2) = validFrom
            resultData(outRow, 3) = validTo
            resultData(outRow, 4) = templateData(blockStart, 4)
            resultData(outRow, 5) = templateData(blockStart, 5)
            resultData(outRow, 6) = templateData(blockStart, 6)
            
            sourceOffset = schemaIndex - 1
            If sourceOffset < blockRows Then
                resultData(outRow, 7) = templateData(blockStart + sourceOffset, 7)
                resultData(outRow, 8) = templateData(blockStart + sourceOffset, 8)
                resultData(outRow, 9) = ""
            Else
                resultData(outRow, 7) = ""
                resultData(outRow, 8) = ""
                resultData(outRow, 9) = ""
            End If
        Next schemaIndex
    Next blockIndex
    
    BuildNewPeriodDataFromTemplate = resultData
    Exit Function
    
BuildFail:
    Erase resultData
End Function


Private Function GetBottomKVPeriod(ByVal wsKV As Worksheet, ByVal firstDataRow As Long) As String
    Dim r As Long
    Dim lastRow As Long
    Dim valueText As String
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    
    For r = lastRow To firstDataRow Step -1
        valueText = PID_GetRowKVPeriod(wsKV, r)
        If valueText <> "" Then
            GetBottomKVPeriod = valueText
            Exit Function
        End If
    Next r
    
SafeExit:
End Function


Private Sub PID_NormalizeKVWarningText(ByVal wsKV As Worksheet)
    If wsKV Is Nothing Then Exit Sub
    
    wsKV.Range("A2").Value = _
        "Wichtig: Alte KV-Perioden niemals loeschen oder ueberschreiben. " & _
        "Wenn ab Mai neue Werte gueltig sind, immer eine neue KV-Periode hinzufuegen. " & _
        "In den Monatsblaettern wird spaeter nur der KV-Code ausgewaehlt. " & _
        "Nur Zeilen mit Status OK verwenden."
    wsKV.Range("A2").WrapText = True
End Sub


Private Sub PID_NormalizeKVTableHeader(ByVal wsKV As Worksheet)
    If wsKV Is Nothing Then Exit Sub
    
    wsKV.Range("A3").Value = "KV-Periode"
    wsKV.Range("B3").Value = "Gueltig ab"
    wsKV.Range("C3").Value = "Gueltig bis"
    wsKV.Range("D3").Value = "KV-Code"
    wsKV.Range("E3").Value = "KV-Gruppe"
    wsKV.Range("F3").Value = "Beschaeftigungsdauer"
    wsKV.Range("G3").Value = "Monatsstunden"
    wsKV.Range("H3").Value = "Monatslohn"
    wsKV.Range("I3").Value = "Status"
    wsKV.Range("J3").Value = "Pruefung"
End Sub


Private Sub PID_ClearKVDataArea(ByVal wsKV As Worksheet, ByVal firstDataRow As Long, ByVal lastRow As Long)
    Dim targetRange As Range
    
    On Error GoTo TryUnmerge
    
    Set targetRange = wsKV.Range("A" & firstDataRow & ":J" & lastRow)
    targetRange.ClearContents
    Exit Sub
    
TryUnmerge:
    On Error Resume Next
    targetRange.UnMerge
    targetRange.ClearContents
    On Error GoTo 0
End Sub


Private Sub PID_ClearTrailingKVArea(ByVal wsKV As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim trailingRange As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If startRow > endRow Then Exit Sub
    
    Set trailingRange = wsKV.Range("A" & startRow & ":J" & endRow)
    
    On Error Resume Next
    trailingRange.UnMerge
    On Error GoTo SafeExit
    
    trailingRange.ClearContents
    trailingRange.Validation.Delete
    trailingRange.ClearFormats
    
SafeExit:
End Sub


Private Function PID_GetSheetCleanupLastRow(ByVal ws As Worksheet, ByVal fallbackRow As Long) As Long
    Dim usedLastRow As Long
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then
        PID_GetSheetCleanupLastRow = fallbackRow
        Exit Function
    End If
    
    usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    
    If usedLastRow > fallbackRow Then
        PID_GetSheetCleanupLastRow = usedLastRow
    Else
        PID_GetSheetCleanupLastRow = fallbackRow
    End If
    
    Exit Function
    
SafeExit:
    PID_GetSheetCleanupLastRow = fallbackRow
End Function


Public Sub CleanupLOHNTABELLE_TESTTrailingArea()
    Dim wsKV As Worksheet
    Dim firstDataRow As Long
    Dim dataLastRow As Long
    Dim cleanupLastRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    firstDataRow = 4
    dataLastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    cleanupLastRow = PID_GetSheetCleanupLastRow(wsKV, dataLastRow)
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    PID_ClearTrailingKVArea wsKV, dataLastRow + 1, cleanupLastRow
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Sub InsertNewKVPeriodRows(ByVal wsKV As Worksheet, _
                                  ByVal firstDataRow As Long, _
                                  ByVal newPeriodData As Variant)
    Dim newRowCount As Long
    Dim insertRowCount As Long
    Dim dataStartRow As Long
    Dim insertEndRow As Long
    Dim newArea As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If Not IsArray(newPeriodData) Then Exit Sub
    
    newRowCount = UBound(newPeriodData, 1)
    insertRowCount = newRowCount + 1
    
    If newRowCount <= 0 Then Exit Sub
    
    insertEndRow = firstDataRow + insertRowCount - 1
    dataStartRow = firstDataRow + 1
    
    wsKV.Rows(firstDataRow & ":" & insertEndRow).Insert Shift:=xlDown
    
    Set newArea = wsKV.Range("A" & firstDataRow & ":J" & insertEndRow)
    
    On Error Resume Next
    newArea.UnMerge
    On Error GoTo SafeExit
    
    newArea.Clear
    
    wsKV.Range("A" & dataStartRow & ":I" & (dataStartRow + newRowCount - 1)).Value = newPeriodData
    
    PID_WriteKVPeriodTitleRow wsKV, firstDataRow, CStr(newPeriodData(1, 1)), newPeriodData(1, 2), newPeriodData(1, 3)
    
    Application.CutCopyMode = False
    
SafeExit:
    Application.CutCopyMode = False
End Sub


Private Sub PID_WriteKVPeriodTitleRow(ByVal wsKV As Worksheet, _
                                      ByVal rowNumber As Long, _
                                      ByVal periodName As String, _
                                      ByVal validFrom As Variant, _
                                      ByVal validTo As Variant)
    Dim titleText As String
    
    If wsKV Is Nothing Then Exit Sub
    If rowNumber < 1 Then Exit Sub
    
    titleText = periodName
    
    If IsDate(validFrom) And IsDate(validTo) Then
        titleText = titleText & "   |   gueltig von " & Format$(CDate(validFrom), "dd.mm.yyyy") & _
                    " bis " & Format$(CDate(validTo), "dd.mm.yyyy")
    End If
    
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).UnMerge
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).ClearContents
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).Merge
    
    wsKV.Cells(rowNumber, "A").Value = titleText
    
    With wsKV.Range("A" & rowNumber & ":J" & rowNumber)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(235, 235, 235)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
End Sub


Private Function GetTopKVPeriod(ByVal wsKV As Worksheet, ByVal firstDataRow As Long) As String
    Dim r As Long
    Dim lastRow As Long
    Dim rowPeriod As String
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    For r = firstDataRow To lastRow
        rowPeriod = PID_GetRowKVPeriod(wsKV, r)
        
        If rowPeriod <> "" Then
            GetTopKVPeriod = rowPeriod
            Exit Function
        End If
    Next r

SafeExit:
End Function


Private Function FindFirstRowOfPeriod(ByVal wsKV As Worksheet, _
                                      ByVal periodName As String, _
                                      ByVal firstDataRow As Long) As Long
    Dim boundsFirst As Long
    Dim boundsLast As Long
    
    periodName = NormalizeKVPeriodText(periodName)
    
    If PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, boundsFirst, boundsLast) Then
        FindFirstRowOfPeriod = boundsFirst
    End If
End Function


Private Function FindLastRowOfPeriod(ByVal wsKV As Worksheet, _
                                     ByVal periodName As String, _
                                     ByVal firstDataRow As Long) As Long
    Dim boundsFirst As Long
    Dim boundsLast As Long
    
    periodName = NormalizeKVPeriodText(periodName)
    
    If PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, boundsFirst, boundsLast) Then
        FindLastRowOfPeriod = boundsLast
    End If
End Function


Private Function PID_GetPeriodRowBounds(ByVal wsKV As Worksheet, _
                                        ByVal periodName As String, _
                                        ByVal firstDataRow As Long, _
                                        ByRef outFirstRow As Long, _
                                        ByRef outLastRow As Long) As Boolean
    Dim r As Long
    Dim lastRow As Long
    Dim rowPeriod As String
    Dim inPeriod As Boolean
    
    On Error GoTo SafeExit
    
    outFirstRow = 0
    outLastRow = 0
    inPeriod = False
    
    If wsKV Is Nothing Then Exit Function
    If periodName = "" Then Exit Function
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    
    For r = firstDataRow To lastRow
        rowPeriod = PID_GetRowKVPeriod(wsKV, r)
        
        If rowPeriod = periodName Then
            If outFirstRow = 0 Then outFirstRow = r
            outLastRow = r
            inPeriod = True
        ElseIf inPeriod And rowPeriod <> "" Then
            Exit For
        End If
    Next r
    
    PID_GetPeriodRowBounds = (outFirstRow > 0 And outLastRow >= outFirstRow)
    
SafeExit:
End Function


Private Function KVPeriodExists(ByVal wsKV As Worksheet, _
                                ByVal periodName As String, _
                                ByVal firstDataRow As Long) As Boolean
    Dim boundsFirst As Long
    Dim boundsLast As Long
    
    periodName = NormalizeKVPeriodText(periodName)
    KVPeriodExists = PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, boundsFirst, boundsLast)
End Function


Private Function NormalizeKVPeriodText(ByVal periodText As String) As String
    Dim s As String
    
    s = Trim$(CStr(periodText))
    
    If s = "" Then
        NormalizeKVPeriodText = ""
        Exit Function
    End If
    
    If InStr(1, s, "|", vbTextCompare) > 0 Then
        s = Trim$(Left$(s, InStr(1, s, "|", vbTextCompare) - 1))
    End If
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    s = Replace(s, "kv ", "KV ")
    s = Replace(s, "Kv ", "KV ")
    s = Replace(s, "kV ", "KV ")
    
    NormalizeKVPeriodText = s
End Function


Private Function PID_GetRowKVPeriod(ByVal wsKV As Worksheet, ByVal rowNumber As Long) As String
    Dim r As Long
    Dim cellText As String
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    If rowNumber < 1 Then Exit Function
    
    cellText = Trim$(CStr(wsKV.Cells(rowNumber, "A").Value))
    
    If cellText <> "" Then
        PID_GetRowKVPeriod = NormalizeKVPeriodText(cellText)
        Exit Function
    End If
    
    For r = rowNumber - 1 To 4 Step -1
        cellText = Trim$(CStr(wsKV.Cells(r, "A").Value))
        If cellText <> "" Then
            PID_GetRowKVPeriod = NormalizeKVPeriodText(cellText)
            Exit Function
        End If
    Next r
    
SafeExit:
End Function


Private Function IsValidKVPeriodName(ByVal periodName As String) As Boolean
    Dim s As String
    Dim yearPart As String
    Dim parts As Variant
    Dim y1 As Long
    Dim y2 As Long
    
    On Error GoTo InvalidPeriod
    
    s = NormalizeKVPeriodText(periodName)
    
    If Left$(s, 3) <> "KV " Then GoTo InvalidPeriod
    
    yearPart = Mid$(s, 4)
    
    If InStr(yearPart, "/") = 0 Then GoTo InvalidPeriod
    
    parts = Split(yearPart, "/")
    
    If UBound(parts) <> 1 Then GoTo InvalidPeriod
    If Not IsNumeric(parts(0)) Then GoTo InvalidPeriod
    If Not IsNumeric(parts(1)) Then GoTo InvalidPeriod
    
    y1 = CLng(parts(0))
    y2 = CLng(parts(1))
    
    If y1 < 2000 Or y1 > 2100 Then GoTo InvalidPeriod
    If y2 < 2000 Or y2 > 2100 Then GoTo InvalidPeriod
    If y2 <> y1 + 1 Then GoTo InvalidPeriod
    
    IsValidKVPeriodName = True
    Exit Function

InvalidPeriod:
    IsValidKVPeriodName = False
End Function


Private Function GetNextKVPeriodName(ByVal currentPeriodName As String) As String
    Dim s As String
    Dim yearPart As String
    Dim parts As Variant
    Dim y1 As Long
    Dim y2 As Long
    
    On Error GoTo Fallback
    
    s = NormalizeKVPeriodText(currentPeriodName)
    
    If Left$(s, 3) <> "KV " Then GoTo Fallback
    
    yearPart = Mid$(s, 4)
    parts = Split(yearPart, "/")
    
    If UBound(parts) <> 1 Then GoTo Fallback
    If Not IsNumeric(parts(0)) Then GoTo Fallback
    If Not IsNumeric(parts(1)) Then GoTo Fallback
    
    y1 = CLng(parts(0))
    y2 = CLng(parts(1))
    
    GetNextKVPeriodName = "KV " & CStr(y1 + 1) & "/" & CStr(y2 + 1)
    Exit Function

Fallback:
    GetNextKVPeriodName = "KV " & CStr(Year(Date)) & "/" & CStr(Year(Date) + 1)
End Function


Private Sub FormatKVPeriodArea(ByVal wsKV As Worksheet)
    Dim lastRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    If lastRow < 4 Then GoTo SafeExit
    
    With wsKV
        .Range("A4:I" & lastRow).VerticalAlignment = xlCenter
        .Range("A4:I" & lastRow).HorizontalAlignment = xlCenter
        .Range("B4:C" & lastRow).NumberFormat = "dd.mm.yyyy"
        
        ' Gesamte Datenflaeche einheitlich mit duennem Raster versehen.
        .Range("A4:J" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A4:J" & lastRow).Borders.Weight = xlThin
        .Range("A4:J" & lastRow).Borders.Color = RGB(150, 150, 150)
        
        .Range("A4:A" & lastRow).NumberFormat = "@"
        .Range("G4:G" & lastRow).NumberFormatLocal = "0,00"
        
        PID_ApplyEuroNumberFormat .Range("H4:H" & lastRow)
        
        .Columns("A").ColumnWidth = 16
        .Columns("D").ColumnWidth = 14
        .Columns("G").ColumnWidth = 13
        .Columns("H").ColumnWidth = 14
        PID_ConfigureKVStatusColumnWidths wsKV, 4, lastRow
    End With
    
    ' Status- und Pruefungsformeln auf allen gueltigen Datenzeilen wiederherstellen.
    PID_ApplyKVStatusFormulas wsKV, 4, lastRow
    
    ' Eingabefelder fuer Monatsstunden/Monatslohn muessen editierbar bleiben.
    PID_ConfigureKVInputCellLocks wsKV, 4, lastRow
    
    PID_ApplyKVVisualGrouping wsKV, 4, lastRow
    PID_FormatKVRowTypography wsKV, 4, lastRow

SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Sub PID_EnsureKVStatusFormulas(ByVal wsKV As Worksheet)
    Dim lastRow As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Sub
    
    PID_ApplyKVStatusFormulas wsKV, 4, lastRow
    PID_ConfigureKVStatusColumnWidths wsKV, 4, lastRow
    PID_ConfigureKVInputCellLocks wsKV, 4, lastRow
    
SafeExit:
End Sub


Private Sub PID_FormatKVRowTypography(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    Dim rowRange As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    wsKV.Range("A3:J3").Font.Bold = True
    wsKV.Range("A3:J3").Font.Size = 10
    
    For r = firstRow To lastRow
        Set rowRange = wsKV.Range("A" & r & ":J" & r)
        
        If wsKV.Range("A" & r).MergeCells Then
            ' Titelzeile wird in PID_WriteKVPeriodTitleRow formatiert.
        ElseIf Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "" Then
            rowRange.Font.Bold = False
            rowRange.Font.Size = 10
            rowRange.VerticalAlignment = xlCenter
            rowRange.HorizontalAlignment = xlCenter
        End If
    Next r
    
SafeExit:
End Sub


Private Sub PID_ConfigureKVStatusColumnWidths(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    wsKV.Columns("I").ColumnWidth = 22
    wsKV.Columns("J").ColumnWidth = 24
    
    If firstRow <= lastRow Then
        wsKV.Range("I" & firstRow & ":J" & lastRow).WrapText = False
    End If
    
SafeExit:
End Sub


Private Sub PID_ApplyKVStatusFormulas(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    Dim hasKeyData As Boolean
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    For r = firstRow To lastRow
        On Error Resume Next
        Err.Clear
        
        If wsKV.Range("A" & r).MergeCells Then
            wsKV.Cells(r, "I").ClearContents
            wsKV.Cells(r, "J").ClearContents
        Else
            hasKeyData = (Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "")
            
            If hasKeyData Then
                wsKV.Cells(r, "I").FormulaR1C1 = _
                    "=IF(RC1="""","""",IF(OR(RC2="""",RC3="""",RC4="""",RC5="""",RC6=""""),""Stammdaten fehlen"",IF(AND(RC7="""",RC8=""""),""Werte fehlen"",IF(RC7="""",""Monatsstunden fehlen"",IF(RC8="""",""Monatslohn fehlt"",""OK"")))))"
                
                wsKV.Cells(r, "J").FormulaR1C1 = _
                    "=IF(RC7="""","""",IF(COUNTIFS(C1,RC1,C4,RC4,C7,RC7)>1,""Doppelte Monatsstunden"",""""))"
            Else
                wsKV.Cells(r, "I").ClearContents
                wsKV.Cells(r, "J").ClearContents
            End If
        End If
        
        On Error GoTo 0
    Next r
End Sub


Private Sub PID_ConfigureKVInputCellLocks(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    wsKV.Range("A" & firstRow & ":J" & lastRow).Locked = True
    
    For r = firstRow To lastRow
        ' Nur echte Datenzeilen freigeben (keine zusammengefuehrten Titelzeilen).
        If Not wsKV.Range("A" & r).MergeCells Then
            If Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "" Then
                wsKV.Range("G" & r & ":H" & r).Locked = False
            End If
        End If
    Next r
    
SafeExit:
End Sub


Public Sub ApplyKVVisualGrouping()
    Dim wsKV As Worksheet
    Dim lastRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    PID_ApplyKVVisualGrouping wsKV, 4, lastRow
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Sub PID_ApplyKVVisualGrouping(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    Dim currentPeriod As String
    Dim prevPeriod As String
    Dim currentCode As String
    Dim prevCode As String
    Dim groupText As String
    Dim rowRange As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    For r = firstRow To lastRow
        Set rowRange = wsKV.Range("A" & r & ":J" & r)
        
        ' Titelzeilen nicht ueberschreiben (bleiben grau/formatiert).
        If wsKV.Range("A" & r).MergeCells Then GoTo NextRow
        
        currentPeriod = PID_GetRowKVPeriod(wsKV, r)
        currentCode = UCase$(Trim$(CStr(wsKV.Cells(r, "D").Value)))
        groupText = UCase$(Trim$(CStr(wsKV.Cells(r, "E").Value)))
        
        ' Reset per-row visual baseline first.
        rowRange.Interior.Pattern = xlSolid
        rowRange.Interior.PatternColorIndex = xlAutomatic
        rowRange.Interior.TintAndShade = 0
        rowRange.Interior.PatternTintAndShade = 0
        rowRange.Interior.ColorIndex = xlNone
        
        rowRange.Borders(xlEdgeTop).LineStyle = xlContinuous
        rowRange.Borders(xlEdgeTop).Weight = xlThin
        rowRange.Borders(xlEdgeTop).Color = RGB(180, 180, 180)
        
        ' Soft BG color blocks.
        Select Case True
            Case InStr(1, groupText, "BG1", vbTextCompare) > 0
                rowRange.Interior.Color = RGB(242, 248, 255)
            Case InStr(1, groupText, "BG2", vbTextCompare) > 0
                rowRange.Interior.Color = RGB(241, 250, 241)
            Case InStr(1, groupText, "BG3", vbTextCompare) > 0
                rowRange.Interior.Color = RGB(255, 246, 237)
        End Select
        
        ' Strong separator when a new KV period starts.
        If currentPeriod <> "" Then
            If currentPeriod <> prevPeriod Then
                rowRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                rowRange.Borders(xlEdgeTop).Weight = xlMedium
                rowRange.Borders(xlEdgeTop).Color = RGB(90, 90, 90)
                
                ' Bei neuem Zeitraum die KV-Code-Referenz zuruecksetzen.
                prevCode = ""
            End If
            prevPeriod = currentPeriod
        End If
        
        ' Zusaetzliche Trennlinie zwischen KV-Untergruppen (Basis / 5 / 10 / 15).
        If currentCode <> "" Then
            If prevCode <> "" And currentCode <> prevCode Then
                rowRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                rowRange.Borders(xlEdgeTop).Weight = xlMedium
                rowRange.Borders(xlEdgeTop).Color = RGB(120, 120, 120)
            End If
            prevCode = currentCode
        End If
        
NextRow:
    Next r
    
    ' Aussenrahmen der Tabelle am Ende explizit verstaerken.
    With wsKV.Range("A" & firstRow & ":J" & lastRow)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = RGB(90, 90, 90)
        
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeRight).Color = RGB(90, 90, 90)
        
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = RGB(90, 90, 90)
    End With
    
SafeExit:
End Sub


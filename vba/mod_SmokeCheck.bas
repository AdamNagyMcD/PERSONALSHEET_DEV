Attribute VB_Name = "mod_SmokeCheck"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Private Const PID_SMOKE_SHEET_NAME As String = "SYSTEM_CHECK"


Public Sub PID_RunSystemSmokeCheck()
    Dim ws As Worksheet
    Dim nextRow As Long
    
    Dim statusText As String
    Dim details As String
    Dim manualSteps As String
    
    Dim passCount As Long
    Dim failCount As Long
    Dim reviewCount As Long
    
    On Error GoTo CleanFail
    
    Set ws = PID_GetOrCreateSmokeCheckSheet()
    PID_PrepareSmokeCheckSheet ws
    
    nextRow = 2
    
    statusText = PID_EvaluateTest1(details)
    manualSteps = PID_GetManualStepsForTest(1, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 1 - Future Hour Change", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest2(details)
    manualSteps = PID_GetManualStepsForTest(2, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 2 - Exit Employee", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest3(details)
    manualSteps = PID_GetManualStepsForTest(3, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 3 - Future Employee Survival", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest4(details)
    manualSteps = PID_GetManualStepsForTest(4, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 4 - Override Survival", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest5(details)
    manualSteps = PID_GetManualStepsForTest(5, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 5 - O18:Q25 Propagation", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest6(details)
    manualSteps = PID_GetManualStepsForTest(6, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 6 - Column L", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest7(details)
    manualSteps = PID_GetManualStepsForTest(7, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 7 - Excel 2016 Compatibility", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest8(details)
    manualSteps = PID_GetManualStepsForTest(8, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 8 - Windows Platform", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest9(details)
    manualSteps = PID_GetManualStepsForTest(9, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 9 - Workbook Year (EINSTELLUNG C35)", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest10(details)
    manualSteps = PID_GetManualStepsForTest(10, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 10 - Month Sheet A1 Index", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest11(details)
    manualSteps = PID_GetManualStepsForTest(11, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 11 - UEBERSICHT FINANZIELL Block", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest12(details)
    manualSteps = PID_GetManualStepsForTest(12, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 12 - Durchrechnung Inputs E30/I30", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest13(details)
    manualSteps = PID_GetManualStepsForTest(13, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 13 - LOHNTABELLE KV Data", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest14(details)
    manualSteps = PID_GetManualStepsForTest(14, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 14 - KV_CODE_LIST Named Range", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest15(details)
    manualSteps = PID_GetManualStepsForTest(15, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 15 - Monatslohn Formula", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest16(details)
    manualSteps = PID_GetManualStepsForTest(16, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 16 - VBA Project Access", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest17(details)
    manualSteps = PID_GetManualStepsForTest(17, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 17 - Month Sheet Lock Policy", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    statusText = PID_EvaluateTest18(details)
    manualSteps = PID_GetManualStepsForTest(18, statusText)
    PID_WriteSmokeResult ws, nextRow, "TEST 18 - Q12 Vormonat Lock Rules", statusText, details, manualSteps
    PID_TrackSmokeStatus statusText, passCount, failCount, reviewCount
    nextRow = nextRow + 1
    
    PID_FinalizeSmokeSheet ws, nextRow - 1
    
    MsgBox "System Smoke Check abgeschlossen." & vbCrLf & vbCrLf & _
           "PASS: " & passCount & vbCrLf & _
           "FAIL: " & failCount & vbCrLf & _
           "REVIEW: " & reviewCount & vbCrLf & vbCrLf & _
           "Details im Blatt '" & PID_SMOKE_SHEET_NAME & "'.", _
           vbInformation, "System Smoke Check"
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei PID_RunSystemSmokeCheck:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "System Smoke Check"
End Sub


Public Sub PID_FilterSmokeReviewOnly()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filterRange As Range
    
    On Error GoTo CleanFail
    
    Set ws = PID_GetOrCreateSmokeCheckSheet()
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Keine Smoke-Check Daten vorhanden. Bitte zuerst PID_RunSystemSmokeCheck " & PID_UTxtAusfuehren() & ".", _
               vbInformation, "System Smoke Check"
        Exit Sub
    End If
    
    Set filterRange = ws.Range("A1:E" & lastRow)
    
    If ws.AutoFilterMode Then
        If ws.FilterMode Then ws.ShowAllData
    End If
    
    filterRange.AutoFilter Field:=2, Criteria1:="REVIEW"
    
    MsgBox "Filter aktiv: Es werden nur REVIEW-F" & PID_UTxtAe() & "lle angezeigt.", _
           vbInformation, "System Smoke Check"
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei PID_FilterSmokeReviewOnly:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "System Smoke Check"
End Sub


Public Sub PID_ClearSmokeFilter()
    Dim ws As Worksheet
    
    On Error GoTo CleanFail
    
    Set ws = PID_GetOrCreateSmokeCheckSheet()
    
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    End If
    
    MsgBox "Filter wurde " & PID_UTxtZurueckgesetzt() & ". Alle Smoke-Check Zeilen sind sichtbar.", _
           vbInformation, "System Smoke Check"
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei PID_ClearSmokeFilter:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "System Smoke Check"
End Sub


Private Function PID_EvaluateTest1(ByRef details As String) As String
    If Not PID_BasicMonthStructureOk() Then
        details = "Monatsblatt-Struktur unvollstaendig (A1 oder D/E/F Spaltenbereich)."
        PID_EvaluateTest1 = "FAIL"
        Exit Function
    End If
    
    details = "Basisstruktur vorhanden. Verhaltenspruefung (Mai/Juni unveraendert, Juli+ angepasst) manuell ausfuehren."
    PID_EvaluateTest1 = "REVIEW"
End Function


Private Function PID_EvaluateTest2(ByRef details As String) As String
    If Not PID_AllMonthSheetsHaveRange("I" & PID_FIRST_ROW & ":I" & PID_LAST_ROW) Then
        details = "Austrittsbereich I3:I82 fehlt in mindestens einem Monatsblatt."
        PID_EvaluateTest2 = "FAIL"
        Exit Function
    End If
    
    details = "Austrittsfelder vorhanden. Sichtbarkeit im Austrittsmonat und Entfernung ab Folgemonat manuell pruefen."
    PID_EvaluateTest2 = "REVIEW"
End Function


Private Function PID_EvaluateTest3(ByRef details As String) As String
    If PID_CountMonthSheets() <> 12 Then
        details = "Es sind nicht alle 12 Monatsblaetter vorhanden."
        PID_EvaluateTest3 = "FAIL"
        Exit Function
    End If
    
    details = "Monatskette vorhanden. Ueberleben kuenftiger MA nach Rueckwaertskopie manuell pruefen."
    PID_EvaluateTest3 = "REVIEW"
End Function


Private Function PID_EvaluateTest4(ByRef details As String) As String
    If Not PID_AllMonthSheetsHaveRange("D" & PID_FIRST_ROW & ":E" & PID_LAST_ROW) Then
        details = "Override-Bereich D:E fehlt in mindestens einem Monatsblatt."
        PID_EvaluateTest4 = "FAIL"
        Exit Function
    End If
    
    details = "Override-Bereiche vorhanden. Ueberlebenspruefung fuer D/E und Stabilitaet frueherer Monate manuell pruefen."
    PID_EvaluateTest4 = "REVIEW"
End Function


Private Function PID_EvaluateTest5(ByRef details As String) As String
    If Not PID_AllMonthSheetsHaveRange("O18:Q25") Then
        details = "Bereich O18:Q25 fehlt in mindestens einem Monatsblatt."
        PID_EvaluateTest5 = "FAIL"
        Exit Function
    End If
    
    details = "Bereich O18:Q25 in allen Monatsblaettern vorhanden. Propagationsverhalten manuell pruefen."
    PID_EvaluateTest5 = "REVIEW"
End Function


Private Function PID_EvaluateTest6(ByRef details As String) As String
    If Not PID_AllMonthSheetsHaveRange("L" & PID_FIRST_ROW & ":L" & PID_LAST_ROW) Then
        details = "Info-Spalte L fehlt in mindestens einem Monatsblatt."
        PID_EvaluateTest6 = "FAIL"
        Exit Function
    End If
    
    details = "Info-Spalte L vorhanden. Sicherstellen, dass L nicht propagiert wird (manuelle Pruefung)."
    PID_EvaluateTest6 = "REVIEW"
End Function


Private Function PID_EvaluateTest7(ByRef details As String) As String
    If Val(Application.Version) < 16# Then
        details = "Excel-Version < 16 erkannt (" & Application.Version & ")."
        PID_EvaluateTest7 = "FAIL"
        Exit Function
    End If
    
    If PID_WorkbookUsesBlockedFunctions() Then
        details = "Mindestens eine nicht erlaubte Excel-Funktion (z.B. XLOOKUP) wurde gefunden."
        PID_EvaluateTest7 = "FAIL"
        Exit Function
    End If
    
    details = "Excel-Version >= 16 und keine blockierten Funktionen gefunden."
    PID_EvaluateTest7 = "PASS"
End Function


Private Function PID_EvaluateTest8(ByRef details As String) As String
    ' Die Datei ist ausschliesslich fuer Excel unter Windows freigegeben.
    If InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0 Then
        details = "Windows Excel erkannt (" & Application.OperatingSystem & ")."
        PID_EvaluateTest8 = "PASS"
    Else
        details = "Kein Windows Excel erkannt (" & Application.OperatingSystem & "). Diese Datei ist nur fuer Windows freigegeben."
        PID_EvaluateTest8 = "FAIL"
    End If
End Function


Private Function PID_EvaluateTest9(ByRef details As String) As String
    Dim wsConfig As Worksheet
    Dim yearCell As Variant
    Dim yearFromFn As Long
    
    On Error GoTo Fail
    
    If Not PID_WorksheetExists(PID_EINSTELLUNG_SHEET) Then
        details = "Blatt " & PID_EINSTELLUNG_SHEET & " fehlt."
        PID_EvaluateTest9 = "FAIL"
        Exit Function
    End If
    
    Set wsConfig = ThisWorkbook.Worksheets(PID_EINSTELLUNG_SHEET)
    yearCell = wsConfig.Range(PID_WORKBOOK_YEAR_CELL).Value2
    
    If Not IsNumeric(yearCell) Then
        details = PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL & " ist nicht numerisch."
        PID_EvaluateTest9 = "FAIL"
        Exit Function
    End If
    
    If CLng(yearCell) < 2000 Or CLng(yearCell) > 2100 Then
        details = PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL & " ausserhalb 2000-2100 (" & yearCell & ")."
        PID_EvaluateTest9 = "FAIL"
        Exit Function
    End If
    
    yearFromFn = PID_GetWorkbookYear()
    If yearFromFn <> CLng(yearCell) Then
        details = "PID_GetWorkbookYear=" & yearFromFn & ", Zelle=" & yearCell & " — Abweichung."
        PID_EvaluateTest9 = "FAIL"
        Exit Function
    End If
    
    details = "Arbeitsjahr " & yearCell & " in " & PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL & " konsistent."
    PID_EvaluateTest9 = "PASS"
    Exit Function

Fail:
    details = "Fehler beim Lesen von " & PID_EINSTELLUNG_SHEET & "!" & PID_WORKBOOK_YEAR_CELL & "."
    PID_EvaluateTest9 = "FAIL"
End Function


Private Function PID_EvaluateTest10(ByRef details As String) As String
    Dim monthNames As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim monthIndex As Long
    Dim a1Value As Variant
    Dim badSheets As String
    
    monthNames = PID_MonthNames()
    badSheets = ""
    
    For i = LBound(monthNames) To UBound(monthNames)
        monthIndex = PID_GetMonthIndexFromSheetName(CStr(monthNames(i)))
        If monthIndex < 1 Then
            badSheets = badSheets & CStr(monthNames(i)) & " (unbekannter Name); "
            GoTo NextMonth
        End If
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo 0
        
        If ws Is Nothing Then
            badSheets = badSheets & CStr(monthNames(i)) & " (fehlt); "
            GoTo NextMonth
        End If
        
        a1Value = ws.Range("A1").Value2
        If Not IsNumeric(a1Value) Then
            badSheets = badSheets & ws.Name & " (A1 nicht numerisch); "
        ElseIf CLng(a1Value) <> monthIndex Then
            badSheets = badSheets & ws.Name & " (A1=" & a1Value & ", erwartet " & monthIndex & "); "
        End If
NextMonth:
    Next i
    
    If Len(badSheets) > 0 Then
        details = "Monatsindex A1 fehlerhaft: " & badSheets
        PID_EvaluateTest10 = "FAIL"
        Exit Function
    End If
    
    details = "Alle 12 Monatsblaetter: A1 = 1..12 passend zum Blattnamen."
    PID_EvaluateTest10 = "PASS"
End Function


Private Function PID_EvaluateTest11(ByRef details As String) As String
    Dim ws As Worksheet
    Dim headerText As String
    
    On Error GoTo Fail
    
    If Not PID_WorksheetExists("UBERSICHT") Then
        details = "Blatt UEBERSICHT fehlt."
        PID_EvaluateTest11 = "FAIL"
        Exit Function
    End If
    
    Set ws = ThisWorkbook.Worksheets("UBERSICHT")
    headerText = Trim$(CStr(ws.Cells(4, 3).Text))
    
    If InStr(1, headerText, "SALES", vbTextCompare) > 0 Then
        details = "FINANZIELL-Block erkannt (Header SALES in C4)."
        PID_EvaluateTest11 = "PASS"
        Exit Function
    End If
    
    If Trim$(CStr(ws.Cells(6, 3).Value)) = "BUDGET" Then
        details = "FINANZIELL-Block erkannt (BUDGET in C6)."
        PID_EvaluateTest11 = "PASS"
        Exit Function
    End If
    
    details = "Kein FINANZIELL-Block auf UEBERSICHT (C4/C6 Pruefung)."
    PID_EvaluateTest11 = "FAIL"
    Exit Function

Fail:
    details = "Fehler beim Pruefen des FINANZIELL-Blocks auf UEBERSICHT."
    PID_EvaluateTest11 = "FAIL"
End Function


Private Function PID_EvaluateTest12(ByRef details As String) As String
    Dim ws As Worksheet
    Dim titleText As String
    
    On Error GoTo Fail
    
    If Not PID_WorksheetExists("UBERSICHT") Then
        details = "Blatt UEBERSICHT fehlt."
        PID_EvaluateTest12 = "FAIL"
        Exit Function
    End If
    
    Set ws = ThisWorkbook.Worksheets("UBERSICHT")
    titleText = CStr(ws.Cells(28, 2).Text)
    
    If InStr(1, titleText, "DURCHRECHNUNGSSTUNDEN", vbTextCompare) = 0 Then
        If Trim$(CStr(ws.Cells(31, 2).Value)) <> "Zeitraum" Then
            details = "Durchrechnungsblock auf UEBERSICHT nicht gefunden (Zeile 28/31)."
            PID_EvaluateTest12 = "FAIL"
            Exit Function
        End If
    End If
    
    If ws.Range("E30").Locked Then
        details = "E30 (Jaenner Verfuegbar Plan) ist gesperrt — soll editierbar sein."
        PID_EvaluateTest12 = "FAIL"
        Exit Function
    End If
    
    If ws.Range("I30").Locked Then
        details = "I30 (Jaenner Muster Plan) ist gesperrt — soll editierbar sein."
        PID_EvaluateTest12 = "FAIL"
        Exit Function
    End If
    
    details = "Durchrechnungsblock vorhanden; E30 und I30 sind entsperrt (editierbar)."
    PID_EvaluateTest12 = "PASS"
    Exit Function

Fail:
    details = "Fehler beim Pruefen des Durchrechnungsblocks."
    PID_EvaluateTest12 = "FAIL"
End Function


Private Function PID_EvaluateTest13(ByRef details As String) As String
    Dim wsKV As Worksheet
    Dim lastRow As Long
    
    On Error GoTo Fail
    
    If Not PID_WorksheetExists(PID_LOHNTABELLE_SHEET) Then
        details = "Blatt " & PID_LOHNTABELLE_SHEET & " fehlt."
        PID_EvaluateTest13 = "FAIL"
        Exit Function
    End If
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    lastRow = wsKV.Cells(wsKV.Rows.Count, "D").End(xlUp).Row
    
    If lastRow < 4 Then
        details = "LOHNTABELLE: keine KV-Daten in Spalte D (lastRow=" & lastRow & ")."
        PID_EvaluateTest13 = "FAIL"
        Exit Function
    End If
    
    If Len(Trim$(CStr(wsKV.Cells(4, 4).Value2))) = 0 And lastRow = 4 Then
        details = "LOHNTABELLE: Spalte D ab Zeile 4 leer."
        PID_EvaluateTest13 = "FAIL"
        Exit Function
    End If
    
    details = "LOHNTABELLE enthaelt KV-Daten (Spalte D bis Zeile " & lastRow & ")."
    PID_EvaluateTest13 = "PASS"
    Exit Function

Fail:
    details = "Fehler beim Pruefen der LOHNTABELLE."
    PID_EvaluateTest13 = "FAIL"
End Function


Private Function PID_EvaluateTest14(ByRef details As String) As String
    Dim listRange As Range
    
    On Error GoTo Fail
    
    Set listRange = Nothing
    On Error Resume Next
    Set listRange = ThisWorkbook.Names(PID_KV_CODE_LIST_NAME).RefersToRange
    On Error GoTo Fail
    
    If listRange Is Nothing Then
        details = "Named Range '" & PID_KV_CODE_LIST_NAME & "' fehlt oder ungueltig."
        PID_EvaluateTest14 = "FAIL"
        Exit Function
    End If
    
    If listRange.Cells.Count < 1 Then
        details = "Named Range '" & PID_KV_CODE_LIST_NAME & "' ist leer."
        PID_EvaluateTest14 = "FAIL"
        Exit Function
    End If
    
    details = "Named Range '" & PID_KV_CODE_LIST_NAME & "' vorhanden (" & listRange.Address(False, False) & ")."
    PID_EvaluateTest14 = "PASS"
    Exit Function

Fail:
    details = "Named Range '" & PID_KV_CODE_LIST_NAME & "' nicht lesbar. FullSystemRefresh ausfuehren."
    PID_EvaluateTest14 = "FAIL"
End Function


Private Function PID_EvaluateTest15(ByRef details As String) As String
    Dim ws As Worksheet
    Dim monthNames As Variant
    Dim i As Long
    Dim missingSheets As String
    Dim checkedCount As Long
    
    monthNames = PID_MonthNames()
    missingSheets = ""
    checkedCount = 0
    
    For i = LBound(monthNames) To UBound(monthNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            checkedCount = checkedCount + 1
            If Not PID_MonthSheetHasMonatslohnFormula(ws) Then
                If Len(missingSheets) > 0 Then missingSheets = missingSheets & ", "
                missingSheets = missingSheets & ws.Name
            End If
        End If
    Next i
    
    If checkedCount = 0 Then
        details = "Kein Monatsblatt zum Pruefen der Monatslohn-Formel gefunden."
        PID_EvaluateTest15 = "FAIL"
        Exit Function
    End If
    
    If Len(missingSheets) > 0 Then
        details = "Monatslohn (Spalte G) unvollstaendig auf: " & missingSheets & " — E/F-Zeile ohne G-Wert. FullSystemRefresh empfohlen."
        PID_EvaluateTest15 = "FAIL"
        Exit Function
    End If
    
    details = "Monatslohn auf allen " & checkedCount & " Monatsblaettern in Ordnung (Formel oder VBA-Wert in G)."
    PID_EvaluateTest15 = "PASS"
End Function


Private Function PID_EvaluateTest16(ByRef details As String) As String
    Dim vbProj As Object
    
    On Error GoTo NoAccess
    
    Set vbProj = ThisWorkbook.VBProject
    If vbProj Is Nothing Then GoTo NoAccess
    
    details = "VBProject zugaenglich (Trust access to VBA project object model aktiv)."
    PID_EvaluateTest16 = "PASS"
    Set vbProj = Nothing
    Exit Function

NoAccess:
    details = "VBProject nicht zugaenglich. Fuer Dev-Import: Excel-Optionen -> Vertrauensstellung -> VBA-Projektobjektmodell."
    PID_EvaluateTest16 = "REVIEW"
    Set vbProj = Nothing
End Function


Private Function PID_EvaluateTest17(ByRef details As String) As String
    Dim ws As Worksheet
    Dim failList As String
    
    On Error GoTo Fail
    
    Set ws = ThisWorkbook.Worksheets("Januar")
    
    If ws.Range("G" & PID_FIRST_ROW).Locked = False Then
        details = "G3 auf Januar ist entsperrt — Lock-Policy nicht aktiv."
        PID_EvaluateTest17 = "FAIL"
        Exit Function
    End If
    
    If ws.Range("E" & PID_FIRST_ROW).Locked Then
        details = "E3 auf Januar ist gesperrt — Whitelist fehlt."
        PID_EvaluateTest17 = "FAIL"
        Exit Function
    End If
    
    If ws.Range("B" & PID_FIRST_ROW).Locked Then
        details = "B3 auf Januar ist gesperrt — Whitelist (B/C) fehlt."
        PID_EvaluateTest17 = "FAIL"
        Exit Function
    End If
    
    If ws.Range("I" & PID_FIRST_ROW).Locked Then
        details = "I3 auf Januar ist gesperrt — Whitelist (I/J) fehlt."
        PID_EvaluateTest17 = "FAIL"
        Exit Function
    End If
    
    details = "Januar: G3 gesperrt, E3/B3/I3 entsperrt. Lock-Policy aktiv."
    PID_EvaluateTest17 = "PASS"
    Exit Function

Fail:
    details = "Blatt Januar nicht lesbar. FullSystemRefresh ausfuehren."
    PID_EvaluateTest17 = "FAIL"
End Function


Private Function PID_EvaluateTest18(ByRef details As String) As String
    Dim wsStart As Worksheet
    Dim wsNormal As Worksheet
    
    On Error GoTo Fail
    
    Set wsStart = ThisWorkbook.Worksheets("Februar")
    Set wsNormal = ThisWorkbook.Worksheets("Januar")
    
    If Not wsStart.Range("Q12").Locked Then
        details = "Q12 auf Februar ist entsperrt — muss gesperrt sein (Durchrechnung-Formel)."
        PID_EvaluateTest18 = "FAIL"
        Exit Function
    End If
    
    If wsNormal.Range("Q12").Locked Then
        details = "Q12 auf Januar ist gesperrt — muss entsperrt sein (Vormonat-Eingabe)."
        PID_EvaluateTest18 = "FAIL"
        Exit Function
    End If
    
    details = "Februar Q12 gesperrt (Formel), Januar Q12 entsperrt (Eingabe). Korrekt."
    PID_EvaluateTest18 = "PASS"
    Exit Function

Fail:
    details = "Januar oder Februar nicht lesbar. FullSystemRefresh ausfuehren."
    PID_EvaluateTest18 = "FAIL"
End Function


Private Function PID_BasicMonthStructureOk() As Boolean
    If PID_CountMonthSheets() <> 12 Then Exit Function
    If Not PID_AllMonthSheetsHaveRange("A1") Then Exit Function
    If Not PID_AllMonthSheetsHaveRange("D" & PID_FIRST_ROW & ":F" & PID_LAST_ROW) Then Exit Function
    
    PID_BasicMonthStructureOk = True
End Function


Private Function PID_AllMonthSheetsHaveRange(ByVal rangeAddress As String) As Boolean
    Dim monthNames As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim testRange As Range
    
    On Error GoTo NotOk
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        If ws Is Nothing Then GoTo NotOk
        
        Set testRange = Nothing
        Set testRange = ws.Range(rangeAddress)
        If testRange Is Nothing Then GoTo NotOk
    Next i
    
    PID_AllMonthSheetsHaveRange = True
    Exit Function
    
NotOk:
    PID_AllMonthSheetsHaveRange = False
End Function


Private Function PID_WorkbookUsesBlockedFunctions() As Boolean
    Dim ws As Worksheet
    Dim formulaCells As Range
    Dim c As Range
    Dim formulaText As String
    
    On Error GoTo SafeExit
    
    For Each ws In ThisWorkbook.Worksheets
        Set formulaCells = Nothing
        
        On Error Resume Next
        Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo SafeExit
        
        If Not formulaCells Is Nothing Then
            For Each c In formulaCells.Cells
                formulaText = UCase$(CStr(c.Formula))
                
                If InStr(formulaText, "XLOOKUP(") > 0 Then
                    PID_WorkbookUsesBlockedFunctions = True
                    Exit Function
                End If
                
                If InStr(formulaText, "XVERWEIS(") > 0 Then
                    PID_WorkbookUsesBlockedFunctions = True
                    Exit Function
                End If
            Next c
        End If
    Next ws
    
SafeExit:
End Function


Private Function PID_GetOrCreateSmokeCheckSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_SMOKE_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PID_SMOKE_SHEET_NAME
    End If
    
    Set PID_GetOrCreateSmokeCheckSheet = ws
End Function


Private Sub PID_PrepareSmokeCheckSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    
    ws.Cells.Clear
    
    ws.Range("A1").Value = "Test Case"
    ws.Range("B1").Value = "Status"
    ws.Range("C1").Value = "Details"
    ws.Range("D1").Value = "Checked At"
    ws.Range("E1").Value = "Manual Steps"
    
    ws.Range("A1:E1").Font.Bold = True
End Sub


Private Sub PID_WriteSmokeResult(ByVal ws As Worksheet, _
                                 ByVal rowNumber As Long, _
                                 ByVal testName As String, _
                                 ByVal statusText As String, _
                                 ByVal details As String, _
                                 ByVal manualSteps As String)
    If ws Is Nothing Then Exit Sub
    
    ws.Cells(rowNumber, "A").Value = testName
    ws.Cells(rowNumber, "B").Value = statusText
    ws.Cells(rowNumber, "C").Value = details
    ws.Cells(rowNumber, "D").Value = Now
    ws.Cells(rowNumber, "E").Value = manualSteps
    
    Select Case UCase$(statusText)
        Case "PASS"
            ws.Cells(rowNumber, "B").Interior.Color = RGB(198, 239, 206)
            ws.Cells(rowNumber, "B").Font.Color = RGB(0, 97, 0)
        Case "FAIL"
            ws.Cells(rowNumber, "B").Interior.Color = RGB(255, 199, 206)
            ws.Cells(rowNumber, "B").Font.Color = RGB(156, 0, 6)
        Case Else
            ws.Cells(rowNumber, "B").Interior.Color = RGB(255, 235, 156)
            ws.Cells(rowNumber, "B").Font.Color = RGB(156, 101, 0)
    End Select
End Sub


Private Function PID_GetManualStepsForTest(ByVal testNumber As Long, ByVal statusText As String) As String
    If UCase$(statusText) <> "REVIEW" Then
        PID_GetManualStepsForTest = "-"
        Exit Function
    End If
    
    Select Case testNumber
        Case 1
            PID_GetManualStepsForTest = "1) Im Blatt Juli Feld E oder F aendern. 2) CopyData ab Juli starten. 3) Pruefen: Mai/Juni unveraendert, ab Juli aktualisiert."
        Case 2
            PID_GetManualStepsForTest = "1) Im Blatt August ein Austrittsdatum in Spalte I setzen. 2) CopyData ab August starten. 3) Pruefen: In August sichtbar, ab September entfernt."
        Case 3
            PID_GetManualStepsForTest = "1) Im Blatt Oktober neuen Mitarbeiter in B/C erfassen. 2) Rueckwaertskopie von frueherem Monat starten. 3) Pruefen: Oktober-Eintrag bleibt erhalten."
        Case 4
            PID_GetManualStepsForTest = "1) In einem spaeteren Monat D/E-Werte manuell anpassen. 2) Propagation starten. 3) Pruefen: Overrides bleiben erhalten, fruehere Monate bleiben unveraendert."
        Case 5
            PID_GetManualStepsForTest = "1) Im Quellmonat O18:Q25 aendern. 2) CopyData starten. 3) In Zielmonaten pruefen: O18:Q25 wurde korrekt uebernommen."
        Case 6
            PID_GetManualStepsForTest = "1) In einem Zielmonat einen individuellen Wert in Spalte L setzen. 2) Propagation von anderem Monat starten. 3) Pruefen: Spalte L wird nicht ueberschrieben."
        Case 7
            PID_GetManualStepsForTest = "1) Workbook in Excel 2016 oeffnen. 2) PID_QuickSystemCheck und PID_FullSystemRefresh ausfuehren. 3) Pruefen: keine Compile- oder Formel-Fehler."
        Case 8
            PID_GetManualStepsForTest = "1) Workbook in Excel fuer Windows oeffnen. 2) PID_QuickSystemCheck, CopyData und PID_RunSystemSmokeCheck ausfuehren. 3) Pruefen: keine plattformspezifischen Fehler."
        Case 16
            PID_GetManualStepsForTest = "1) Excel-Optionen -> Vertrauensstellungscenter -> Makroeinstellungen. 2) 'Zugriff auf das VBA-Projektobjektmodell' aktivieren. 3) Nur fuer Dev/Import noetig."
        Case Else
            PID_GetManualStepsForTest = "-"
    End Select
End Function


Private Sub PID_TrackSmokeStatus(ByVal statusText As String, _
                                 ByRef passCount As Long, _
                                 ByRef failCount As Long, _
                                 ByRef reviewCount As Long)
    Select Case UCase$(statusText)
        Case "PASS"
            passCount = passCount + 1
        Case "FAIL"
            failCount = failCount + 1
        Case Else
            reviewCount = reviewCount + 1
    End Select
End Sub


Private Sub PID_FinalizeSmokeSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    If ws Is Nothing Then Exit Sub
    If lastRow < 2 Then Exit Sub
    
    ws.Range("A1:E" & lastRow).Borders.LineStyle = xlContinuous
    ws.Range("A1:E" & lastRow).Borders.Weight = xlThin
    
    ws.Columns("A").ColumnWidth = 40
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 58
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 70
    
    ws.Range("C2:C" & lastRow).WrapText = True
    ws.Range("E2:E" & lastRow).WrapText = True
    ws.Range("D2:D" & lastRow).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    
    ws.Range("A1:E1").HorizontalAlignment = xlCenter
    ws.Range("A1:E1").VerticalAlignment = xlCenter
    
    ws.Range("A2:B" & lastRow).VerticalAlignment = xlCenter
    ws.Range("C2:C" & lastRow).VerticalAlignment = xlTop
    ws.Range("E2:E" & lastRow).VerticalAlignment = xlTop
    ws.Range("D2:D" & lastRow).HorizontalAlignment = xlCenter
    
    ws.Rows("1:" & lastRow).EntireRow.AutoFit
End Sub


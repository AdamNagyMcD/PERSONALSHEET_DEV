Attribute VB_Name = "mod_SmokeCheck"
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
    PID_WriteSmokeResult ws, nextRow, "TEST 8 - Mac Compatibility", statusText, details, manualSteps
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


Public Sub PID_SystemSmokeCheck()
    PID_RunSystemSmokeCheck
End Sub


Public Sub PID_FilterSmokeReviewOnly()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filterRange As Range
    
    On Error GoTo CleanFail
    
    Set ws = PID_GetOrCreateSmokeCheckSheet()
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Keine Smoke-Check Daten vorhanden. Bitte zuerst PID_RunSystemSmokeCheck ausfuehren.", _
               vbInformation, "System Smoke Check"
        Exit Sub
    End If
    
    Set filterRange = ws.Range("A1:E" & lastRow)
    
    If ws.AutoFilterMode Then
        If ws.FilterMode Then ws.ShowAllData
    End If
    
    filterRange.AutoFilter Field:=2, Criteria1:="REVIEW"
    
    MsgBox "Filter aktiv: Es werden nur REVIEW-Faelle angezeigt.", _
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
    
    MsgBox "Filter wurde zurueckgesetzt. Alle Smoke-Check Zeilen sind sichtbar.", _
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
    If PID_HasWindowsOnlyApiDeclarations() Then
        details = "Windows-only API Declare-Anweisungen in exportierten VBA-Dateien gefunden."
        PID_EvaluateTest8 = "FAIL"
        Exit Function
    End If
    
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        details = "Makro auf Mac gestartet. Keine Windows-only API Muster gefunden."
        PID_EvaluateTest8 = "PASS"
    Else
        details = "Keine Windows-only API Muster gefunden. Finale Laufzeitpruefung auf Mac bleibt manuell."
        PID_EvaluateTest8 = "REVIEW"
    End If
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


Private Function PID_HasWindowsOnlyApiDeclarations() As Boolean
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim lineText As String
    Dim fileNumber As Integer
    
    On Error GoTo SafeExit
    
    folderPath = ThisWorkbook.Path
    If folderPath = "" Then Exit Function
    
    fileName = Dir(folderPath & "\vba\*.bas")
    
    Do While fileName <> ""
        fullPath = folderPath & "\vba\" & fileName
        
        fileNumber = FreeFile
        Open fullPath For Input As #fileNumber
        
        Do While Not EOF(fileNumber)
            Line Input #fileNumber, lineText
            
            If InStr(1, lineText, "Declare ", vbTextCompare) > 0 Then
                PID_HasWindowsOnlyApiDeclarations = True
                Close #fileNumber
                Exit Function
            End If
        Loop
        
        Close #fileNumber
        fileName = Dir()
    Loop
    
SafeExit:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
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
            PID_GetManualStepsForTest = "1) Workbook in Excel fuer Mac oeffnen. 2) PID_QuickSystemCheck, CopyData und PID_RunSystemSmokeCheck ausfuehren. 3) Pruefen: keine plattformspezifischen Fehler."
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


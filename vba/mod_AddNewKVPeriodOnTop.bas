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
    
    InsertNewKVPeriodRows wsKV, firstDataRow, templateFirstRow, templateLastRow, newPeriodData
    
    FormatKVPeriodArea wsKV
    PID_NormalizeKVTableHeader wsKV
    
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
    wsKV.Range("A2").WrapText = False
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
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
    wsKV.Range("A" & firstDataRow & ":I" & firstDataRow + periodRowCount - 1).Value = periodData
    PID_ClearTrailingKVArea wsKV, firstDataRow + periodRowCount, cleanupLastRow
    PID_NormalizeKVWarningText wsKV
    PID_NormalizeKVTableHeader wsKV
    
    If firstDataRow + periodRowCount <= wsKV.Rows.Count Then
        wsKV.Range("I" & firstDataRow & ":I" & firstDataRow + periodRowCount - 1).Replace What:="", Replacement:="OK", LookAt:=xlWhole
    End If
    
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
    Dim r As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    If periodFirstRow <= 0 Or periodLastRow < periodFirstRow Then Exit Function
    
    firstCode = Trim$(CStr(wsKV.Cells(periodFirstRow, "D").Value))
    If firstCode = "" Then Exit Function
    
    For r = periodFirstRow To periodLastRow
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
                resultData(outRow, 9) = templateData(blockStart + sourceOffset, 9)
                
                If Trim$(CStr(resultData(outRow, 9))) = "" Then
                    resultData(outRow, 9) = "OK"
                End If
            Else
                resultData(outRow, 7) = ""
                resultData(outRow, 8) = ""
                resultData(outRow, 9) = "Werte fehlen"
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
        valueText = NormalizeKVPeriodText(CStr(wsKV.Cells(r, "A").Value))
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
                                  ByVal templateFirstRow As Long, _
                                  ByVal templateLastRow As Long, _
                                  ByVal newPeriodData As Variant)
    Dim templateRowCount As Long
    Dim newRowCount As Long
    Dim insertRowCount As Long
    Dim copyRowCount As Long
    Dim sourceRange As Range
    Dim sourceLastRow As Long
    Dim pasteStartRow As Long
    Dim validFrom As Variant
    Dim validTo As Variant
    
    On Error GoTo SafeExit
    
    templateRowCount = templateLastRow - templateFirstRow + 1
    newRowCount = UBound(newPeriodData, 1)
    insertRowCount = newRowCount + 1
    
    If templateRowCount <= 0 Then Exit Sub
    If newRowCount <= 0 Then Exit Sub
    
    Set sourceRange = wsKV.Rows(templateFirstRow & ":" & templateLastRow)
    
    wsKV.Rows(firstDataRow & ":" & firstDataRow + insertRowCount - 1).Insert Shift:=xlDown
    
    copyRowCount = templateRowCount
    If copyRowCount > newRowCount Then copyRowCount = newRowCount
    
    If copyRowCount > 0 Then
        wsKV.Rows(templateFirstRow & ":" & templateFirstRow + copyRowCount - 1).Copy
        wsKV.Rows(firstDataRow + 1 & ":" & firstDataRow + copyRowCount).PasteSpecial Paste:=xlPasteFormats
        wsKV.Rows(firstDataRow + 1 & ":" & firstDataRow + copyRowCount).PasteSpecial Paste:=xlPasteValidation
        wsKV.Rows(firstDataRow + 1 & ":" & firstDataRow + copyRowCount).PasteSpecial Paste:=xlPasteColumnWidths
    End If
    
    If newRowCount > copyRowCount Then
        sourceLastRow = templateLastRow
        pasteStartRow = firstDataRow + copyRowCount + 1
        
        wsKV.Rows(sourceLastRow & ":" & sourceLastRow).Copy
        wsKV.Rows(pasteStartRow & ":" & firstDataRow + newRowCount).PasteSpecial Paste:=xlPasteFormats
        wsKV.Rows(pasteStartRow & ":" & firstDataRow + newRowCount).PasteSpecial Paste:=xlPasteValidation
    End If
    
    wsKV.Range("A" & firstDataRow + 1 & ":I" & firstDataRow + newRowCount).Value = newPeriodData
    
    validFrom = newPeriodData(1, 2)
    validTo = newPeriodData(1, 3)
    PID_WriteKVPeriodTitleRow wsKV, firstDataRow, CStr(newPeriodData(1, 1)), validFrom, validTo
    
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
    Dim valueText As String
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    For r = firstDataRow To lastRow
        valueText = NormalizeKVPeriodText(CStr(wsKV.Cells(r, "A").Value))
        
        If valueText <> "" Then
            GetTopKVPeriod = valueText
            Exit Function
        End If
    Next r

SafeExit:
End Function


Private Function FindFirstRowOfPeriod(ByVal wsKV As Worksheet, _
                                      ByVal periodName As String, _
                                      ByVal firstDataRow As Long) As Long
    Dim r As Long
    Dim lastRow As Long
    Dim checkPeriod As String
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    periodName = NormalizeKVPeriodText(periodName)
    
    For r = firstDataRow To lastRow
        checkPeriod = NormalizeKVPeriodText(CStr(wsKV.Cells(r, "A").Value))
        
        If checkPeriod = periodName Then
            FindFirstRowOfPeriod = r
            Exit Function
        End If
    Next r

SafeExit:
End Function


Private Function FindLastRowOfPeriod(ByVal wsKV As Worksheet, _
                                     ByVal periodName As String, _
                                     ByVal firstDataRow As Long) As Long
    Dim r As Long
    Dim lastRow As Long
    Dim checkPeriod As String
    Dim firstFound As Boolean
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    periodName = NormalizeKVPeriodText(periodName)
    
    For r = firstDataRow To lastRow
        checkPeriod = NormalizeKVPeriodText(CStr(wsKV.Cells(r, "A").Value))
        
        If checkPeriod = periodName Then
            firstFound = True
            FindLastRowOfPeriod = r
        ElseIf firstFound Then
            Exit Function
        End If
    Next r

SafeExit:
End Function


Private Function KVPeriodExists(ByVal wsKV As Worksheet, _
                                ByVal periodName As String, _
                                ByVal firstDataRow As Long) As Boolean
    Dim r As Long
    Dim lastRow As Long
    Dim checkPeriod As String
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    periodName = NormalizeKVPeriodText(periodName)
    
    For r = firstDataRow To lastRow
        checkPeriod = NormalizeKVPeriodText(CStr(wsKV.Cells(r, "A").Value))
        
        If checkPeriod = periodName Then
            KVPeriodExists = True
            Exit Function
        End If
    Next r

SafeExit:
End Function


Private Function NormalizeKVPeriodText(ByVal periodText As String) As String
    Dim s As String
    
    s = Trim$(CStr(periodText))
    
    If s = "" Then
        NormalizeKVPeriodText = ""
        Exit Function
    End If
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    s = Replace(s, "kv ", "KV ")
    s = Replace(s, "Kv ", "KV ")
    s = Replace(s, "kV ", "KV ")
    
    NormalizeKVPeriodText = s
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
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    If lastRow < 4 Then Exit Sub
    
    With wsKV
        .Range("A4:I" & lastRow).VerticalAlignment = xlCenter
        .Range("A4:I" & lastRow).HorizontalAlignment = xlCenter
        .Range("B4:C" & lastRow).NumberFormat = "dd.mm.yyyy"
        
        .Range("A4:A" & lastRow).NumberFormat = "@"
        .Range("G4:G" & lastRow).NumberFormatLocal = "0,00"
        
        PID_ApplyEuroNumberFormat .Range("H4:H" & lastRow)
        
        .Columns("A").ColumnWidth = 16
        .Columns("D").ColumnWidth = 14
        .Columns("G").ColumnWidth = 13
        .Columns("H").ColumnWidth = 14
        .Columns("I").ColumnWidth = 10
    End With
    
    PID_ApplyKVVisualGrouping wsKV, 4, lastRow

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
    Dim groupText As String
    Dim rowRange As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    For r = firstRow To lastRow
        Set rowRange = wsKV.Range("A" & r & ":J" & r)
        currentPeriod = NormalizeKVPeriodText(CStr(wsKV.Cells(r, "A").Value))
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
                
                ' Slightly emphasize first row of each period.
                If rowRange.Interior.ColorIndex = xlNone Then
                    rowRange.Interior.Color = RGB(245, 245, 245)
                Else
                    rowRange.Interior.TintAndShade = -0.04
                End If
            End If
            prevPeriod = currentPeriod
        End If
    Next r
    
SafeExit:
End Sub


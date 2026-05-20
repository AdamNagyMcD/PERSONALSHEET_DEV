Attribute VB_Name = "mod_AddNewKVPeriodOnTop"
Option Explicit
Private Const PID_ADD_PERIOD_BUTTON_NAME As String = "btn_AddNewKVPeriodOnTop"

Public Sub AddNewKVPeriodOnTop()
    Dim wsKV As Worksheet
    Dim newPeriod As String
    Dim templatePeriod As String
    Dim defaultPeriod As String
    
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
    
    defaultPeriod = GetNextKVPeriodName(templatePeriod)
    
    newPeriod = InputBox( _
        Prompt:="Bitte neuen KV-Zeitraum eingeben." & vbCrLf & vbCrLf & _
                "Beispiel: KV 2025/2026", _
        Title:="Neuen KV-Zeitraum hinzufuegen", _
        Default:=defaultPeriod _
    )
    
    newPeriod = NormalizeKVPeriodText(newPeriod)
    
    If newPeriod = "" Then GoTo CleanExit
    
    If Not IsValidKVPeriodName(newPeriod) Then
        MsgBox "Der eingegebene KV-Zeitraum ist ungueltig." & vbCrLf & vbCrLf & _
               "Bitte Format verwenden, zum Beispiel:" & vbCrLf & _
               "KV 2025/2026", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
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
    
    answer = MsgBox( _
        "Neuer KV-Zeitraum wird oben eingefuegt:" & vbCrLf & vbCrLf & _
        newPeriod & vbCrLf & vbCrLf & _
        "Vorlage:" & vbCrLf & _
        templatePeriod & vbCrLf & vbCrLf & _
        "Anzahl Zeilen: " & templateRowCount & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        "Neuer KV-Zeitraum" _
    )
    
    If answer <> vbYes Then GoTo CleanExit
    
    InsertNewKVPeriodRows wsKV, firstDataRow, templateFirstRow, templateLastRow, newPeriod
    
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
    
    PID_ClearKVDataArea wsKV, firstDataRow, lastRow
    wsKV.Range("A" & firstDataRow & ":I" & firstDataRow + periodRowCount - 1).Value = periodData
    PID_ClearTrailingKVArea wsKV, firstDataRow + periodRowCount, lastRow
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


Private Sub InsertNewKVPeriodRows(ByVal wsKV As Worksheet, _
                                  ByVal firstDataRow As Long, _
                                  ByVal templateFirstRow As Long, _
                                  ByVal templateLastRow As Long, _
                                  ByVal newPeriod As String)
    Dim templateRowCount As Long
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim r As Long
    Dim newRow As Long
    
    On Error GoTo SafeExit
    
    templateRowCount = templateLastRow - templateFirstRow + 1
    If templateRowCount <= 0 Then Exit Sub
    
    Set sourceRange = wsKV.Rows(templateFirstRow & ":" & templateLastRow)
    
    wsKV.Rows(firstDataRow & ":" & firstDataRow + templateRowCount - 1).Insert Shift:=xlDown
    
    sourceRange.Copy
    wsKV.Rows(firstDataRow & ":" & firstDataRow + templateRowCount - 1).PasteSpecial Paste:=xlPasteFormats
    wsKV.Rows(firstDataRow & ":" & firstDataRow + templateRowCount - 1).PasteSpecial Paste:=xlPasteValidation
    wsKV.Rows(firstDataRow & ":" & firstDataRow + templateRowCount - 1).PasteSpecial Paste:=xlPasteColumnWidths
    
    Application.CutCopyMode = False
    
    Set targetRange = wsKV.Range("A" & firstDataRow & ":I" & firstDataRow + templateRowCount - 1)
    
    wsKV.Range("A" & templateFirstRow + templateRowCount & ":I" & templateLastRow + templateRowCount).Copy
    wsKV.Range("A" & firstDataRow).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    For r = 0 To templateRowCount - 1
        newRow = firstDataRow + r
        
        wsKV.Cells(newRow, "A").Value = newPeriod
        
        If Trim$(CStr(wsKV.Cells(newRow, "I").Value)) = "" Then
            wsKV.Cells(newRow, "I").Value = "OK"
        End If
    Next r
    
SafeExit:
    Application.CutCopyMode = False
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
        
        .Range("A4:A" & lastRow).NumberFormat = "@"
        .Range("G4:G" & lastRow).NumberFormatLocal = "0,00"
        
        PID_ApplyEuroNumberFormat .Range("H4:H" & lastRow)
        
        .Columns("A").ColumnWidth = 16
        .Columns("D").ColumnWidth = 14
        .Columns("G").ColumnWidth = 13
        .Columns("H").ColumnWidth = 14
        .Columns("I").ColumnWidth = 10
    End With

SafeExit:
End Sub


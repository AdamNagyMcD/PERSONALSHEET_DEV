Attribute VB_Name = "mod_AddNewKVPeriodOnTop"
Option Explicit

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
    wsKV.Unprotect Password:="company"
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
    
    MarkKVDropdownsDirty
    
    MsgBox "Der neue KV-Zeitraum wurde erfolgreich oben eingefuegt:" & vbCrLf & vbCrLf & _
           newPeriod & vbCrLf & vbCrLf & _
           "Bitte jetzt die Lohnwerte kontrollieren und bei Bedarf anpassen.", _
           vbInformation, "Neuer KV-Zeitraum"

CleanExit:
    On Error Resume Next
    wsKV.Protect Password:="company", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    On Error GoTo 0
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not wsKV Is Nothing Then
        wsKV.Protect Password:="company", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei AddNewKVPeriodOnTop:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Neuer KV-Zeitraum"
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


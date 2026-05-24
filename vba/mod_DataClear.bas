Attribute VB_Name = "mod_DataClear"
Option Explicit


Public Sub DataClear()
    PID_ClearCurrentMonthData
End Sub


Public Sub PID_ClearCurrentMonthData()
    Dim ws As Worksheet
    Dim answer As VbMsgBoxResult
    Dim monthIndex As Long
    Dim monthHint As String
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    
    On Error GoTo CleanFail
    
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Bitte zuerst ein Monatsblatt auswaehlen (z.B. Januar, Juli).", _
               vbExclamation, "Daten loeschen"
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    If Not PID_ValidateWorkerMonthSheet(ws, monthIndex, "Daten loeschen") Then Exit Sub
    
    monthHint = " (Monat " & monthIndex & ")"
    
    answer = MsgBox( _
        "Alle Eingabedaten auf dem Monatsblatt '" & ws.Name & "'" & monthHint & " werden geloescht." & vbCrLf & vbCrLf & _
        "Geloescht werden:" & vbCrLf & _
        "- Mitarbeiterdaten B:N" & vbCrLf & _
        "- Monatsinfo O18:Q25" & vbCrLf & _
        "- Hinweis O45" & vbCrLf & _
        "- Fluktuation Q31" & vbCrLf & vbCrLf & _
        "Nicht geloescht werden:" & vbCrLf & _
        "- Formate" & vbCrLf & _
        "- Kopfzeilen" & vbCrLf & _
        "- Grundstruktur" & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        "Monatsdaten loeschen" _
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
    
    PID_TryUnprotectMonthSheet ws
    
    PID_ClearMonthInputAreas ws
    PID_ApplyMonthSheetFormatsAfterClear ws
    
    RefreshKVStundenDropdownForSheet ws
    PID_EnsureMonatslohnFormulasOnSheet ws

    MarkFluktuationDirty
    MarkFinanzSummaryDirtyForMonth ws
    MarkAllKVDropdownsDirty
    
    PID_TryProtectMonthSheet ws
    
    MsgBox "Die Monatsdaten wurden geloescht.", _
           vbInformation, "Daten loeschen"

CleanExit:
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    
    PID_TryProtectMonthSheet ws
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_ClearCurrentMonthData:" & vbCrLf & _
           originalErrNumber & " - " & originalErrDescription, _
           vbExclamation, "Daten loeschen"
End Sub


Private Sub PID_ClearMonthInputAreas(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    ' Hauptbereich:
    ' B:N = Mitarbeiterdaten.
    ' L ist Monatsinfo, wird hier ebenfalls geloescht, weil es nur diesen Monat betrifft.
    ws.Range("B" & PID_FIRST_ROW & ":N" & PID_LAST_ROW).ClearContents
    
    ' Monatsinfo / Zusatzbereich.
    ws.Range("O18:Q25").ClearContents
    
    ' Hinweis Betriebszugehoerigkeit.
    ws.Range("O45").ClearContents
    
    ' Fluktuation Kennzahl.
    ws.Range("Q31").ClearContents

SafeExit:
End Sub


Private Sub PID_ApplyMonthSheetFormatsAfterClear(ByVal ws As Worksheet)
    Dim euroSymbol As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    euroSymbol = ChrW(8364)
    
    ws.Range("D" & PID_FIRST_ROW & ":D" & PID_LAST_ROW).NumberFormat = "dd.mm.yyyy"
    ws.Range("I" & PID_FIRST_ROW & ":I" & PID_LAST_ROW).NumberFormat = "dd.mm.yyyy"
    
    On Error GoTo TryEnglishNumberFormat
    
    ws.Range("F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW).NumberFormatLocal = "0,00"
    ws.Range("G" & PID_FIRST_ROW & ":G" & PID_LAST_ROW).NumberFormatLocal = euroSymbol & " #.##0,00"
    ws.Range("J" & PID_FIRST_ROW & ":J" & PID_LAST_ROW).NumberFormatLocal = euroSymbol & " #.##0,00"
    ws.Range("K" & PID_FIRST_ROW & ":K" & PID_LAST_ROW).NumberFormatLocal = euroSymbol & " #.##0,00"
    ws.Range("Q31").NumberFormatLocal = "0,00%"
    GoTo SafeExit

TryEnglishNumberFormat:
    On Error GoTo SafeExit
    
    ws.Range("F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW).NumberFormat = "0.00"
    ws.Range("G" & PID_FIRST_ROW & ":G" & PID_LAST_ROW).NumberFormat = euroSymbol & " #,##0.00"
    ws.Range("J" & PID_FIRST_ROW & ":J" & PID_LAST_ROW).NumberFormat = euroSymbol & " #,##0.00"
    ws.Range("K" & PID_FIRST_ROW & ":K" & PID_LAST_ROW).NumberFormat = euroSymbol & " #,##0.00"
    ws.Range("Q31").NumberFormat = "0.00%"

SafeExit:
End Sub


Public Sub PID_ClearOnlySelectedEmployeeRows()
    Dim ws As Worksheet
    Dim selectedRows As Collection
    Dim area As Range
    Dim c As Range
    Dim rowKey As String
    Dim rowNumber As Variant
    Dim answer As VbMsgBoxResult
    Dim monthIndex As Long
    Dim monthHint As String
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    
    On Error GoTo CleanFail
    
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Bitte zuerst ein Monatsblatt auswaehlen (z.B. Januar, Juli).", _
               vbExclamation, "Zeilen loeschen"
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    If Not PID_ValidateWorkerMonthSheet(ws, monthIndex, "Zeilen loeschen") Then Exit Sub
    
    monthHint = " (Monat " & monthIndex & ")"
    
    If Selection Is Nothing Then Exit Sub
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Bitte zuerst einen gueltigen Zellbereich markieren.", _
               vbExclamation, "Zeilen loeschen"
        Exit Sub
    End If
    
    Set area = Intersect(Selection, ws.Range("B" & PID_FIRST_ROW & ":N" & PID_LAST_ROW))
    
    If area Is Nothing Then
        MsgBox "Bitte zuerst eine oder mehrere Mitarbeiterzeilen im Bereich B3:N82 markieren.", _
               vbExclamation, "Zeilen loeschen"
        Exit Sub
    End If
    
    Set selectedRows = New Collection
    
    For Each c In area.Cells
        rowKey = CStr(c.Row)
        If Not PID_CollectionHasKey(selectedRows, rowKey) Then
            selectedRows.Add c.Row, rowKey
        End If
    Next c
    
    If selectedRows.count = 0 Then Exit Sub
    
    answer = MsgBox( _
        "Es werden " & selectedRows.count & " Mitarbeiterzeile(n) auf '" & ws.Name & "'" & monthHint & " geloescht." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        "Ausgewaehlte Zeilen loeschen" _
    )
    
    If answer <> vbYes Then Exit Sub
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    PID_TryUnprotectMonthSheet ws
    
    For Each rowNumber In selectedRows
        ws.Range("B" & CLng(rowNumber) & ":N" & CLng(rowNumber)).ClearContents
    Next rowNumber
    
    RefreshKVStundenDropdownForSheet ws
    PID_EnsureMonatslohnFormulasOnSheet ws

    MarkFluktuationDirty
    MarkFinanzSummaryDirtyForMonth ws
    MarkAllKVDropdownsDirty
    
    PID_TryProtectMonthSheet ws
    
    MsgBox "Ausgewaehlte Mitarbeiterzeile(n) wurden geloescht.", _
           vbInformation, "Zeilen loeschen"

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    
    PID_TryProtectMonthSheet ws
    
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_ClearOnlySelectedEmployeeRows:" & vbCrLf & _
           originalErrNumber & " - " & originalErrDescription, _
           vbExclamation, "Zeilen loeschen"
End Sub


Private Sub PID_TryUnprotectMonthSheet(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    
SafeExit:
End Sub


Private Sub PID_TryProtectMonthSheet(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
SafeExit:
End Sub


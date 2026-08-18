Attribute VB_Name = "mod_DataClear"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' TR-03: Beim Loeschen aller Daten wird das komplette Freitext-Panel geleert.
' Die Zeilen 18 bis 28 sind auf den Monatsblaettern entsperrt (PID_MONTH_PANEL_FREITEXT_RANGE
' in mod_SchutzHinzufugen), der Monatsclear und CopyData arbeiten dagegen mit O18:Q25.
' Bei "alle Daten loeschen" darf nichts stehen bleiben, deshalb hier der volle Bereich.
Private Const PID_CLEAR_ALL_PANEL_RANGE As String = "O18:Q28"


Public Sub DataClear()
    PID_ClearCurrentMonthData
End Sub


' TR-03 — Alt+F8-Einstieg, analog zu DataClear.
Public Sub AlleDatenLoeschen()
    PID_ClearAllWorkbookData
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
        MsgBox "Bitte zuerst ein Monatsblatt " & PID_UTxtAuswaehlen() & " (z.B. Januar, Juli).", _
               vbExclamation, PID_UTxtDatenLoeschen()
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    If Not PID_ValidateWorkerMonthSheet(ws, monthIndex, PID_UTxtDatenLoeschen()) Then Exit Sub
    
    monthHint = " (Monat " & monthIndex & ")"
    
    answer = MsgBox( _
        "Alle Eingabedaten auf dem Monatsblatt '" & ws.Name & "'" & monthHint & " werden " & PID_UTxtGeloescht() & "." & vbCrLf & vbCrLf & _
        PID_UTxtGeloeschtWerdenLabel() & vbCrLf & _
        "- Mitarbeiterdaten B:F, I:J, M:N" & vbCrLf & _
        "- Monatsinfo O18:Q25" & vbCrLf & _
        "- Hinweis O45" & vbCrLf & _
        "- Fluktuation Q31" & vbCrLf & vbCrLf & _
        "Nicht " & PID_UTxtGeloescht() & " werden:" & vbCrLf & _
        "- Formelspalten G, H, K, L" & vbCrLf & _
        "- Formate" & vbCrLf & _
        "- Kopfzeilen" & vbCrLf & _
        "- Grundstruktur" & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        PID_UTxtMonatsdatenLoeschen() _
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
    
    PID_TrackAction "DataClear", ws.Name & ": Monatsdaten " & PID_UTxtGeloescht()
    
    MsgBox "Die Monatsdaten wurden " & PID_UTxtGeloescht() & ".", _
           vbInformation, PID_UTxtDatenLoeschen()

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
           vbExclamation, PID_UTxtDatenLoeschen()
End Sub


Private Sub PID_ClearMonthInputAreas(ByVal ws As Worksheet)
    Dim inputCells As Range
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    ' Hauptbereich: nur die Eingabespalten (B-F, I-J, M-N).
    ' G, H, K und L bleiben stehen - dort stehen die Formeln.
    Set inputCells = PID_GetEmployeeInputCellsForRows(ws, PID_FIRST_ROW, PID_LAST_ROW)
    If Not inputCells Is Nothing Then inputCells.ClearContents
    
    PID_RestoreFormulaColumnsForRows ws, PID_FIRST_ROW, PID_LAST_ROW
    
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
    Dim inputCells As Range
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
        MsgBox "Bitte zuerst ein Monatsblatt " & PID_UTxtAuswaehlen() & " (z.B. Januar, Juli).", _
               vbExclamation, PID_UTxtZeilenLoeschen()
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    If Not PID_ValidateWorkerMonthSheet(ws, monthIndex, PID_UTxtZeilenLoeschen()) Then Exit Sub
    
    monthHint = " (Monat " & monthIndex & ")"
    
    If Selection Is Nothing Then Exit Sub
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Bitte zuerst einen " & PID_UTxtGueltigen() & " Zellbereich markieren.", _
               vbExclamation, PID_UTxtZeilenLoeschen()
        Exit Sub
    End If
    
    Set area = Intersect(Selection, ws.Range("B" & PID_FIRST_ROW & ":N" & PID_LAST_ROW))
    
    If area Is Nothing Then
        MsgBox "Bitte zuerst eine oder mehrere Mitarbeiterzeilen im Bereich B3:N82 markieren.", _
               vbExclamation, PID_UTxtZeilenLoeschen()
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
        "Es werden " & selectedRows.count & " Mitarbeiterzeile(n) auf '" & ws.Name & "'" & monthHint & " " & PID_UTxtGeloescht() & "." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        PID_UTxtAusgewaehlteZeilenLoeschen() _
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
        Set inputCells = PID_GetEmployeeInputCellsForRows(ws, CLng(rowNumber), CLng(rowNumber))
        If Not inputCells Is Nothing Then inputCells.ClearContents
        
        ' Repariert Zeilen, in denen eine aeltere Version die Formeln mitgeloescht hat.
        PID_RestoreFormulaColumnsForRows ws, CLng(rowNumber), CLng(rowNumber)
    Next rowNumber
    
    RefreshKVStundenDropdownForSheet ws
    PID_EnsureMonatslohnFormulasOnSheet ws

    MarkFluktuationDirty
    MarkFinanzSummaryDirtyForMonth ws
    MarkAllKVDropdownsDirty
    
    PID_TryProtectMonthSheet ws
    
    PID_TrackAction "Zeilen loeschen", ws.Name & ": " & selectedRows.count & " Zeile(n)"
    
    MsgBox PID_UTxtAusgewaehlte() & " Mitarbeiterzeile(n) wurden " & PID_UTxtGeloescht() & ".", _
           vbInformation, PID_UTxtZeilenLoeschen()

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
           vbExclamation, PID_UTxtZeilenLoeschen()
End Sub


'==============================================================================
' TR-03 — Alle Daten loeschen (alle zwoelf Monatsblaetter)
'==============================================================================

' Setzt die Arbeitsmappe auf den leeren Ausgangszustand zurueck: alle Eingaben der
' zwoelf Monatsblaetter und der Stunden-Override-Log werden geleert.
' Erhalten bleiben Struktur, Formate, Blattschutz, die Formelspalten G/H/K/L sowie
' EINSTELLUNG, LOHNTABELLE und UEBERSICHT.
Public Sub PID_ClearAllWorkbookData()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim clearedSheets As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    
    On Error GoTo CleanFail
    
    If Not PID_ConfirmClearAllFirstStep() Then Exit Sub
    If Not PID_ConfirmClearAllSecondStep() Then Exit Sub
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Alle Monatsdaten werden " & PID_UTxtGeloescht() & "..."
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        Err.Clear
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            ' Pro Blatt gekapselt: ein Fehler auf einem Monat darf die restlichen
            ' Monate nicht ueberspringen.
            If PID_ClearAllDataOnMonthSheet(ws) Then clearedSheets = clearedSheets + 1
        End If
    Next i
    
    ' Ohne dieses Leeren wuerden alte Stunden-Overrides bei der naechsten Eingabe
    ' desselben Mitarbeiters wieder auftauchen.
    PID_ResetHourOverrideLog
    
    MarkFluktuationDirty
    MarkAllKVDropdownsDirty
    
    PID_TrackAction "Alle Daten loeschen", clearedSheets & " " & PID_UTxtMonatsblaetter()
    
    Application.StatusBar = False
    
    MsgBox "Alle Monatsdaten wurden " & PID_UTxtGeloescht() & "." & vbCrLf & vbCrLf & _
           PID_UTxtMonatsblaetter() & " geleert: " & clearedSheets & " / 12" & vbCrLf & _
           "Stunden-Log geleert" & vbCrLf & vbCrLf & _
           "Die Datei ist jetzt bereit " & PID_UTxtFuer() & " neue Eingaben.", _
           vbInformation, PID_UTxtAlleDatenLoeschen()

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    
    PID_TryProtectMonthSheet ws
    
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_ClearAllWorkbookData:" & vbCrLf & _
           originalErrNumber & " - " & originalErrDescription, _
           vbExclamation, PID_UTxtAlleDatenLoeschen()
End Sub


Private Function PID_ConfirmClearAllFirstStep() As Boolean
    Dim answer As VbMsgBoxResult
    
    answer = MsgBox( _
        "ACHTUNG: Alle Eingabedaten in allen 12 Monaten werden " & PID_UTxtGeloescht() & "." & vbCrLf & vbCrLf & _
        PID_UTxtGeloeschtWerdenLabel() & vbCrLf & _
        "- Mitarbeiterdaten B:F, I:J, M:N (Januar bis Dezember)" & vbCrLf & _
        "- Monatsinfo " & PID_CLEAR_ALL_PANEL_RANGE & vbCrLf & _
        "- Hinweis O45 und Fluktuation Q31" & vbCrLf & _
        "- gespeicherte Stunden-" & PID_UTxtAenderung() & "en (Stunden-Log)" & vbCrLf & vbCrLf & _
        "Nicht " & PID_UTxtGeloescht() & " werden:" & vbCrLf & _
        "- Formelspalten G, H, K, L" & vbCrLf & _
        "- Formate, Kopfzeilen, Blattschutz" & vbCrLf & _
        "- EINSTELLUNG (Jahr), LOHNTABELLE, UEBERSICHT" & vbCrLf & _
        "- Vormonat-Wert Q12" & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbExclamation + vbYesNo + vbDefaultButton2, _
        PID_UTxtAlleDatenLoeschen())
    
    PID_ConfirmClearAllFirstStep = (answer = vbYes)
End Function


Private Function PID_ConfirmClearAllSecondStep() As Boolean
    Dim answer As VbMsgBoxResult
    
    answer = MsgBox( _
        "Letzte " & PID_UTxtPruefung() & ":" & vbCrLf & vbCrLf & _
        "Wirklich ALLE Mitarbeiterdaten aus allen 12 Monaten " & PID_UTxtLoeschen() & "?" & vbCrLf & vbCrLf & _
        "Das kann nicht " & PID_UTxtRueckgaengig() & " gemacht werden." & vbCrLf & _
        "Tipp: vorher eine Kopie der Datei speichern.", _
        vbCritical + vbYesNo + vbDefaultButton2, _
        PID_UTxtAlleDatenLoeschen())
    
    PID_ConfirmClearAllSecondStep = (answer = vbYes)
End Function


' Ein Monatsblatt leeren. Die Formelspalten G, H, K und L bleiben stehen (TR-10);
' fehlende Formeln werden dabei gleich ergaenzt.
Private Function PID_ClearAllDataOnMonthSheet(ByVal ws As Worksheet) As Boolean
    Dim inputCells As Range
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    
    PID_TryUnprotectMonthSheet ws
    
    Set inputCells = PID_GetEmployeeInputCellsForRows(ws, PID_FIRST_ROW, PID_LAST_ROW)
    If Not inputCells Is Nothing Then inputCells.ClearContents
    
    PID_RestoreFormulaColumnsForRows ws, PID_FIRST_ROW, PID_LAST_ROW
    
    ws.Range(PID_CLEAR_ALL_PANEL_RANGE).ClearContents
    ws.Range("O45").ClearContents
    ws.Range("Q31").ClearContents
    
    PID_ApplyMonthSheetFormatsAfterClear ws
    
    RefreshKVStundenDropdownForSheet ws
    PID_EnsureMonatslohnFormulasOnSheet ws
    
    MarkFinanzSummaryDirtyForMonth ws
    
    PID_ClearAllDataOnMonthSheet = True

SafeExit:
    PID_TryProtectMonthSheet ws
End Function


Private Sub PID_TryUnprotectMonthSheet(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    
SafeExit:
End Sub


Private Sub PID_TryProtectMonthSheet(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    PID_ProtectWorkerMonthSheet ws
    
SafeExit:
End Sub

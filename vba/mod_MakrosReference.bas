Attribute VB_Name = "mod_MakrosReference"
Option Explicit

Public Const PID_MAKROS_SHEET As String = "MAKROS"


Public Sub PID_BuildMakrosReferenceSheet()
    Dim ws As Worksheet
    Dim nextRow As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set ws = PID_GetOrCreateMakrosSheet()
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    ws.Cells.Clear
    
    ws.Range("A1").Value = "Makro-Uebersicht Personalsheet"
    ws.Range("A2").Value = "Stand"
    ws.Range("B2").Value = Format(Date, "dd.mm.yyyy")
    ws.Range("A3").Value = "Alt+F8 oder Entwicklertools > Makros. Legacy-Namen mit _TEST sind Aliase und muessen nicht extra gestartet werden."
    ws.Range("A3:D3").Merge
    
    ws.Range("A4").Value = "Kategorie"
    ws.Range("B4").Value = "Makro"
    ws.Range("C4").Value = "Wann benutzen"
    ws.Range("D4").Value = "Hinweis / Alternative"
    
    nextRow = 5
    
    PID_AddMacroReferenceRow ws, nextRow, "Alltag", "PID_FullSystemRefresh", _
        "Alles neu berechnen, wenn Dropdown, Lohn oder Fluktuation insgesamt falsch wirken.", _
        "Haupt-Makro fuer Vollrefresh."
    
    PID_AddMacroReferenceRow ws, nextRow, "Alltag", "CopyData", _
        "Daten ab dem aktiven Monatsblatt in die folgenden Monate kopieren.", _
        "Gleich wie DatenInFolgendeMonateKopieren."
    
    PID_AddMacroReferenceRow ws, nextRow, "Alltag", "DataClear", _
        "Aktuellen Monat leeren (mit Sicherheitsabfrage).", _
        "Gleich wie ClearData / DatenLoeschen."
    
    PID_AddMacroReferenceRow ws, nextRow, "Fluktuation", "RefreshFluktuationNow", _
        "Fluktuation-Analyse manuell neu aufbauen.", _
        "Passiert automatisch beim Oeffnen des FLUKTUATION-Blatts."
    
    PID_AddMacroReferenceRow ws, nextRow, "LOHNTABELLE", "AddNewKVPeriodOnTop", _
        "Neue KV-Periode oben einfuegen.", _
        "Normalerweise ueber den Button auf LOHNTABELLE."
    
    PID_AddMacroReferenceRow ws, nextRow, "LOHNTABELLE", "AddCustomKVMonatsstunden", _
        "Eigene Monatsstunden in einen KV-Block einfuegen.", _
        "Normalerweise ueber den gruenen Button auf LOHNTABELLE."
    
    PID_AddMacroReferenceRow ws, nextRow, "LOHNTABELLE", "DeleteSelectedKVPeriods", _
        "Ausgewaehlte KV-Periode loeschen.", _
        "Normalerweise ueber den Button auf LOHNTABELLE."
    
    PID_AddMacroReferenceRow ws, nextRow, "LOHNTABELLE", "RestoreLOHNTABELLEBase2025_2026", _
        "LOHNTABELLE auf definierte Basis zuruecksetzen.", _
        "Nur Admin. _TEST-Name ist identischer Alias."
    
    PID_AddMacroReferenceRow ws, nextRow, "Reparatur", "RestoreAktuelleStundenFormulas", _
        "Spalte H auf allen Monatsblaettern reparieren (Jahr aus EINSTELLUNG).", _
        "Einmalig nach Migration."
    
    PID_AddMacroReferenceRow ws, nextRow, "Reparatur", "RestoreAustrittsdatumValidation", _
        "Spalte I Validierung und AB1/AB2 Grenzen reparieren.", _
        "Einmalig nach Migration."
    
    PID_AddMacroReferenceRow ws, nextRow, "Reparatur", "FixLOHNTABELLE_HeaderTextIfNeeded", _
        "A2-Text und Header auf LOHNTABELLE pruefen/reparieren.", _
        "Laeuft auch automatisch beim Oeffnen."
    
    PID_AddMacroReferenceRow ws, nextRow, "Schutz", "SchutzHinzufugen", _
        "Blattschutz fuer Makros neu setzen.", _
        "Gleich wie ProtectEverything."
    
    PID_AddMacroReferenceRow ws, nextRow, "Schutz", "UnprotectEverything", _
        "Alle Blaetter entsperren und sichtbar machen.", _
        "Gleich wie UnlockEverything."
    
    PID_AddMacroReferenceRow ws, nextRow, "Excel", "PID_ResetExcelState", _
        "Wenn Events/ScreenUpdating haengen geblieben sind.", _
        "Statt ResetExcelEvents oder ResetExcelApplicationState."
    
    PID_AddMacroReferenceRow ws, nextRow, "Check", "PID_QuickSystemCheck", _
        "Kurzer Systemcheck: Pflichtblaetter, Events, Monatsblaetter.", _
        "Nur Diagnose."
    
    PID_AddMacroReferenceRow ws, nextRow, "Check", "PID_RunSystemSmokeCheck", _
        "Automatische Tests 1-8 auf Blatt SYSTEM_CHECK.", _
        "Entwicklung / QA."
    
    PID_AddMacroReferenceRow ws, nextRow, "Entwicklung", "ResetAndImportVBAFiles", _
        "VBA aus dem vba-Ordner neu importieren.", _
        "Nur Entwickler-PC mit VBProject-Zugriff."
    
    PID_AddMacroReferenceRow ws, nextRow, "Automatisch", "(kein Start noetig)", _
        "Monatsblatt oeffnen: Dropdown + Lohn. FLUKTUATION oeffnen: Analyse. Speichern: Fluktuation wenn noetig.", _
        "BuildFluktuationAnalyse und aehnliche internen Makros nicht manuell starten."
    
    PID_FormatMakrosReferenceSheet ws, nextRow - 1
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
    
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Private Function PID_GetOrCreateMakrosSheet() As Worksheet
    Dim ws As Worksheet
    Dim wsEinstellung As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_MAKROS_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PID_MAKROS_SHEET
    End If
    
    On Error Resume Next
    Set wsEinstellung = ThisWorkbook.Worksheets(PID_EINSTELLUNG_SHEET)
    If Not wsEinstellung Is Nothing Then
        ws.Move After:=wsEinstellung
    End If
    On Error GoTo 0
    
    ws.Visible = xlSheetVisible
    Set PID_GetOrCreateMakrosSheet = ws
End Function


Private Sub PID_AddMacroReferenceRow(ByVal ws As Worksheet, _
                                     ByRef nextRow As Long, _
                                     ByVal category As String, _
                                     ByVal macroName As String, _
                                     ByVal whenToUse As String, _
                                     ByVal hint As String)
    ws.Cells(nextRow, 1).Value = category
    ws.Cells(nextRow, 2).Value = macroName
    ws.Cells(nextRow, 3).Value = whenToUse
    ws.Cells(nextRow, 4).Value = hint
    nextRow = nextRow + 1
End Sub


Private Sub PID_FormatMakrosReferenceSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws
        .Range("A1:D1").Merge
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A2:B2").Font.Bold = True
        .Range("A3").WrapText = True
        .Rows(3).RowHeight = 36
        
        .Range("A4:D4").Font.Bold = True
        .Range("A4:D" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A4:D" & lastRow).Borders.Weight = xlThin
        
        .Range("A5:A" & lastRow).Font.Bold = True
        .Range("B5:B" & lastRow).Font.Name = "Courier New"
        
        .Range("C5:D" & lastRow).WrapText = True
        .Range("A4:D" & lastRow).VerticalAlignment = xlCenter
        .Range("C5:D" & lastRow).HorizontalAlignment = xlLeft
        
        .Columns("A").ColumnWidth = 14
        .Columns("B").ColumnWidth = 34
        .Columns("C").ColumnWidth = 44
        .Columns("D").ColumnWidth = 44
        
        .Rows("5:" & lastRow).AutoFit
        
        Dim r As Long
        For r = 5 To lastRow
            If .Rows(r).RowHeight < 28 Then .Rows(r).RowHeight = 28
            If .Rows(r).RowHeight > 120 Then .Rows(r).RowHeight = 120
        Next r
    End With
End Sub

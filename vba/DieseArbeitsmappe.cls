VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Obergrenze fuer das zellweise Nachschreiben beim Einfuege-Schutz (TR-02).
Private Const PID_MAX_PASTE_CELLS As Long = 20000

Private mInternalChange As Boolean


Private mWorkerSheetRuntimeReady As Boolean


Private mSelPrevSheet As String
Private mSelPrevRow As Long
Private mSelPrevCol As Long
Private mSelPrevEValueAtEntry As String
Private mFStundenRefreshRanThisSelection As Boolean


Private Sub PID_EnsureWorkerSheetRuntimeReady()
    If mWorkerSheetRuntimeReady Then Exit Sub
    
    PID_EnsureMonatslohnFormulas
    PID_EnsureLetztesGehaltFormulas
    mWorkerSheetRuntimeReady = True
End Sub


Private Sub Workbook_Open()
    Dim wsUbersicht As Worksheet
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo SafeExit
    
    mInternalChange = True
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    PID_DisableStaleValueFormatting
    
    PID_EnsureKVCodeListNamedRange
    
    PID_SetupSheetProtectionForOpen
    
    PID_EnsureAdminSheet
    
    ' Wichtig:
    ' Beim Oeffnen wird die vollstaendige Fluktuation-Analyse NICHT aktualisiert.
    ' Das ermoeglicht schnelleres Oeffnen - aktualisiert wird nur wenn wirklich noetig.
    MarkFluktuationDirty
    
    On Error Resume Next
    ThisWorkbook.Sheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("KV_DROPDOWN_HELPER").Visible = xlSheetVeryHidden
    PID_HideAdminSheet False
    On Error GoTo SafeExit

SafeExit:
    Set wsUbersicht = PID_GetUbersichtWorksheet()
    If Not wsUbersicht Is Nothing Then
        PID_ApplyUbersichtSheetProtection wsUbersicht
    End If

    If Not ActiveSheet Is Nothing Then
        If PID_IsWorkerMonthSheet(ActiveSheet) Then
            PID_EnsureWorkerSheetRuntimeReady
            PID_RecalculateMonthFormulaColumns ActiveSheet
        End If
    End If

    ' FP-027: Copyright mit Skip-Guard (kein Unprotect/Write wenn schon aktuell).
    PID_ApplyCopyrightToAllSheets

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False

    ' FP-027: AutoCalc erst nach ScreenUpdating=True — Benutzer sieht das Blatt sofort,
    ' das "Berechnet…" laeuft danach im Hintergrund statt das Open zu blockieren.
    PID_ConfigureDeferredWorkbookCalculationOnOpen

    ' FP-FLUKT: Q31 (+ O31-Label) aller Monatsblaetter deferred synchronisieren - nach
    ' ScreenUpdating=True, damit das Open nicht blockiert. Events kurz aus, damit die
    ' Q31-Schreibzugriffe keine Change-/Activate-Events ausloesen.
    ' ScreenUpdating kurz aus waehrend des Syncs, um sichtbares Flackern zu reduzieren.
    On Error Resume Next
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mInternalChange = True
    PID_SyncAllMonthSheetsFluctuationQ31
    mInternalChange = False
    Application.EnableEvents = True
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0

    ' TR-02: Strg+V auf "nur Werte" umlenken.
    PID_ResetManagedPasteState
    PID_InstallPasteHooks
End Sub


' Die Tastenbelegung aus PID_InstallPasteHooks gilt fuer die ganze Excel-Anwendung.
' Deshalb wird sie nur gesetzt, solange diese Mappe aktiv ist - andere Dateien sollen
' ihr normales Strg+V behalten.
Private Sub Workbook_Activate()
    On Error Resume Next
    PID_ResetManagedPasteState
    PID_InstallPasteHooks
End Sub


Private Sub Workbook_Deactivate()
    On Error Resume Next
    PID_RemovePasteHooks
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    PID_RemovePasteHooks
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo SafeExit
    
    If mInternalChange Then Exit Sub
    
    mInternalChange = True
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Vor dem Speichern: nur FLUKTUATION_DATEN (schnell); Analyse beim Tab-Oeffnen.
    RefreshFluktuationDataIfDirty
    
    On Error Resume Next
    ThisWorkbook.Sheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("KV_DROPDOWN_HELPER").Visible = xlSheetVeryHidden
    PID_HideAdminSheet False
    On Error GoTo SafeExit

SafeExit:
    ' Nie im Manual-Modus speichern: Excel merkt sich den Rechenmodus in der Datei
    ' und die Mappe wuerde beim naechsten Oeffnen wieder manuell rechnen.
    On Error Resume Next
    If Application.Calculation <> xlCalculationAutomatic Then
        Application.Calculation = xlCalculationAutomatic
    End If
    Err.Clear
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Dim heavyStarted As Boolean
    
    On Error GoTo SafeExit
    
    If mInternalChange Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    
    mInternalChange = True
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Fluktuation-Analyse nur bei Bedarf aktualisieren (gFluktuationDirty).
    If Sh.Name = PID_FLUKTUATION_SHEET Then
        PID_EnsureSheetProtectionForMacros Sh
        RefreshFluktuationIfDirty
        Application.ScreenUpdating = True
        PID_RemoveFluktuationPdfExportButtonIfNeeded
        GoTo SafeExit
    End If
    
    If Sh.Name = PID_LOHNTABELLE_SHEET Then
        PID_EnsureSheetProtectionForMacros Sh
        EnsureAddNewKVPeriodButton
        FixLOHNTABELLE_HeaderTextIfNeeded
        PID_RecalculateKVStatusColumns Sh
        PID_ConfigureLohntabelleView
        GoTo SafeExit
    End If
    
    If PID_IsUbersichtWorksheet(Sh) Then
        PID_ApplyUbersichtSheetProtection Sh
        RefreshFinanzSummaryIfDirty
        GoTo SafeExit
    End If
    
    PID_EnsureSheetProtectionForMacros Sh
    
    ' Dropdown/Lohn nur bei Bedarf (nach KV-Tabelle-Aenderung), pro Monat einmal.
    If PID_IsWorkerMonthSheet(Sh) Then
        PID_ApplyMonthSheetFillHandleGuard
        PID_EnsureWorkerSheetRuntimeReady
        PID_EnsureMonthSheetCopyDataButton Sh
        
        ' FP-FLUKT: Q31 dieses Monatsblatts beim Aktivieren mit der zentralen Logik
        ' (PID_ComputeFluctuationForPeriod) synchronisieren -> identisch zu UBERSICHT/FLUKTUATION.
        PID_SyncMonthSheetFluctuationQ31 Sh
        
        If gKVDropdownsDirty Or IsKVLohnDirty() Then
            ' heavyStarted stellt sicher, dass SafeExit die Wartung auch bei einem
            ' Fehler beendet — sonst bleibt der Rechenmodus dauerhaft auf Manuell.
            heavyStarted = True
            PID_BeginHeavyMaintenance
            RefreshKVDropdownsIfDirtyForSheet Sh
            RefreshKVLohnIfDirty Sh
            heavyStarted = False
            PID_EndHeavyMaintenance Sh
        Else
            PID_RecalculateMonthFormulaColumns Sh
        End If
        
        ' FINANZIELL: deferred dirty sync (E3:L82) auch beim Monatsblatt-Tab,
        ' nicht nur beim Oeffnen von UEBERSICHT — haelt G/J dort aktuell wie S36/S37.
        RefreshFinanzSummaryIfDirty
        
        GoTo SafeExit
    End If
    
    PID_ClearMonthSheetFillHandleGuard

SafeExit:
    If heavyStarted Then
        heavyStarted = False
        PID_EndHeavyMaintenance Sh
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub PID_ConfigureLohntabelleView()
    On Error GoTo SafeExit
    
    If ActiveWindow Is Nothing Then Exit Sub
    ActiveWindow.DisplayGridlines = False
    
SafeExit:
End Sub


Private Sub PID_ResetMonthView(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
    
SafeExit:
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim pasteHandled As Boolean
    
    On Error GoTo SafeExit
    
    If mInternalChange Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    
    ' Einfuege-Schutz (TR-02): beim Einfuegen aus anderen Dateien oder Programmen soll nur
    ' der Wert uebernommen werden, ohne Formatierung - auf jedem Blatt der Mappe.
    ' Muss als Erstes laufen: EnforcePasteValuesOnly setzt auf Application.Undo auf, und
    ' jeder VBA-Schreibzugriff davor (z. B. der F-Dropdown-Neuaufbau) leert die Undo-Historie.
    mInternalChange = True
    Application.EnableEvents = False
    pasteHandled = False
    EnforcePasteValuesOnly Sh, Target, pasteHandled
    Application.EnableEvents = True
    mInternalChange = False
    
    ' Wenn die KV-Basistabelle geaendert wird, wird die Dropdown-Liste spaeter neu aufgebaut.
    If Sh.Name = PID_LOHNTABELLE_SHEET Then
        If Not Intersect(Target, Sh.Range("D4:J500")) Is Nothing Then
            MarkAllKVLohnDirty
            
            ' F-Dropdown nur bei Struktur/Stunden-Aenderung (D:G), nicht bei Monatslohn (H) allein.
            If Not Intersect(Target, Sh.Range("D4:G500")) Is Nothing Then
                MarkKVDropdownDirtyFromLOHNTABELLERange Sh, Target
            End If
        End If
        
        ' Status/Pruefung (I/J) bei Manual Calculation explizit neu berechnen.
        If Not Intersect(Target, Sh.Range("A4:H500")) Is Nothing Then
            PID_RecalculateKVStatusColumns Sh
        End If
        Exit Sub
    End If
    
    If Not PID_IsWorkerMonthSheet(Sh) Then Exit Sub
    
    mInternalChange = True
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    ' Recalc-Abschnitt mit Resume Next: ein Fehler in Dropdown/Lohn-Recalc darf
    ' das E/F-Logging nicht verhindern.
    
    ' E zuerst: Dropdown in F muss sofort passen.
    If Not Intersect(Target, Sh.Range("E3:E82")) Is Nothing Then
        PID_RefreshFStundenDropdownForERows Sh, Target
        ' Snapshot nachziehen, damit das spaetere Verlassen der Zelle denselben
        ' Neuaufbau nicht ein zweites Mal ausloest.
        PID_SyncPrevEValueSnapshot Sh
    End If
    Err.Clear
    
    ' TR-07: Personal-ID (B) muss pro Monatsblatt eindeutig bleiben. Erst nach dem
    ' Einfuege-Schutz pruefen, damit der gepruefte Wert schon der endgueltige ist.
    If Not Intersect(Target, Sh.Range("B3:B82")) Is Nothing Then
        PID_CheckPersonalIdUniqueness Sh, Target
    End If
    Err.Clear
    
    ' Wenn Spalte E sich aendert:
    ' E = KV-Code / BG -> Dropdown F in der betreffenden Zeile.
    If Not Intersect(Target, Sh.Range("E3:F82")) Is Nothing Then
        PID_RecalculateMonatslohnForChangedRows Sh, Target
    End If
    Err.Clear
    
    If Not Intersect(Target, Sh.Range("D3:F82,I3:I82")) Is Nothing Then
        PID_RecalculateAktuelleStundenForChangedRows Sh, Target
    End If
    Err.Clear
    
    If Not Intersect(Target, Sh.Range("D3:D82,I3:I82,K3:K82")) Is Nothing Then
        PID_RecalculateLetztesGehaltForChangedRows Sh, Target
    End If
    Err.Clear
    
    ' Logging ist datenkritisch: ab hier wieder GoTo SafeExit.
    On Error GoTo SafeExit
    
    If Not Intersect(Target, Sh.Range("E3:F82")) Is Nothing Then
        PID_LogEFAenderungForSheet Sh, Intersect(Target, Sh.Range("E3:F82"))
    End If
    
    ' Fluktuation-Markierung fuer Aktualisierung:
    ' D = Eintrittsdatum
    ' I = Austrittsdatum
    ' N = Austrittsgrund
    If Not Intersect(Target, Sh.Range("D3:D82,I3:I82,N3:N82")) Is Nothing Then
        MarkFluktuationDirtyForMonthSheet Sh
    End If
    
    ' Q31 (Fluktuation %) sofort auf diesem Monatsblatt, wenn Austrittsdatum (I) ODER
    ' Austrittsgrund (N) geaendert wird. FP-Flukt FIX 2: der Grund entscheidet, ob ein
    ' Austritt in die Rate zaehlt (neutrale Gruende wie "Befoerderung" -> 0), daher muss
    ' eine reine N-Aenderung dieselbe Neuberechnung ausloesen wie eine I-Aenderung.
    If Not Intersect(Target, Sh.Range("I3:I82,N3:N82")) Is Nothing Then
        PID_CalculateFluctuation Sh
    End If
    
    If PID_MonthChangeNeedsImmediateFinanzSync(Sh, Target) Then
        PID_SyncFinanzSummaryForMonth Sh
    ElseIf PID_MonthChangeDefersFinanzSummarySync(Sh, Target) Then
        MarkFinanzSummaryDirtyForMonth Sh
    End If

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim eDropdownInvalid As Boolean
    Dim fDropdownInvalid As Boolean
    Dim needsDropdownRepair As Boolean
    Dim appStateTouched As Boolean
    
    On Error GoTo SafeExit
    
    mFStundenRefreshRanThisSelection = False
    
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If PID_IsWorkerMonthSheet(Sh) And Target.CountLarge = 1 Then
        PID_RefreshFStundenOnSelectionChange Sh, Target
    End If
    
    If mInternalChange Then GoTo UpdateSelectionPrev
    If Not PID_IsWorkerMonthSheet(Sh) Then GoTo UpdateSelectionPrev
    If Target.CountLarge > 1 Then GoTo UpdateSelectionPrev
    
    ' Ausserhalb D3:F82: O45 nur loeschen wenn noetig (leichter Pfad).
    If Intersect(Target, Sh.Range("D3:F82")) Is Nothing Then
        If Len(Trim$(CStr(Sh.Range("O45").Value))) = 0 Then GoTo UpdateSelectionPrev
        
        mInternalChange = True
        appStateTouched = True
        Application.EnableEvents = False
        Sh.Range("O45").Value = ""
        GoTo UpdateSelectionPrev
    End If
    
    If Target.Row < PID_FIRST_ROW Or Target.Row > PID_LAST_ROW Then
        If Len(Trim$(CStr(Sh.Range("O45").Value))) > 0 Then
            mInternalChange = True
            appStateTouched = True
            Application.EnableEvents = False
            Sh.Range("O45").Value = ""
        End If
        GoTo UpdateSelectionPrev
    End If
    
    ' D: Beschaeftigungsdauer — ohne ScreenUpdating-Off.
    If Not Intersect(Target, Sh.Range("D3:D82")) Is Nothing Then
        mInternalChange = True
        appStateTouched = True
        Application.EnableEvents = False
        PID_ShowEmploymentDuration Sh, Target
        GoTo UpdateSelectionPrev
    End If
    
    eDropdownInvalid = False
    fDropdownInvalid = False
    
    If Not Intersect(Target, Sh.Range("E3:E82")) Is Nothing Then
        eDropdownInvalid = Not PID_MonthSheetHasValidKVCodeDropdown(Sh)
    End If
    
    If Not Intersect(Target, Sh.Range("F3:F82")) Is Nothing Then
        fDropdownInvalid = Not PID_RowHasValidFStundenDropdown(Sh, Target.Row)
    End If
    
    needsDropdownRepair = eDropdownInvalid Or fDropdownInvalid
    
    ' O45 leeren beim Wechsel von D nach E/F (leichter Pfad).
    If Len(Trim$(CStr(Sh.Range("O45").Value))) > 0 Then
        mInternalChange = True
        appStateTouched = True
        Application.EnableEvents = False
        Sh.Range("O45").Value = ""
        Application.EnableEvents = True
        mInternalChange = False
    End If
    
    ' E/F ungueltig: reparieren.
    If Not needsDropdownRepair Then GoTo UpdateSelectionPrev
    If mFStundenRefreshRanThisSelection Then GoTo UpdateSelectionPrev
    
    mInternalChange = True
    appStateTouched = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If eDropdownInvalid Then
        On Error Resume Next
        PID_RestoreKVCodeDropdownOnSheet Sh
        On Error GoTo SafeExit
    End If
    
    If fDropdownInvalid Then
        On Error Resume Next
        PID_ForceRefreshFStundenDropdownForSheet Sh
        On Error GoTo SafeExit
    End If

UpdateSelectionPrev:
    If TypeOf Sh Is Worksheet Then
        If PID_IsWorkerMonthSheet(Sh) And Target.CountLarge = 1 Then
            PID_UpdateMonthSelectionPrev Sh, Target
        End If
    End If

SafeExit:
    On Error Resume Next
    ' Die beiden Application-Properties nur zurueckstellen, wenn dieser Handler sie
    ' auch wirklich abgeschaltet hat. Der haeufigste Fall — ein simpler Klick, der
    ' keinen der Zweige oben trifft — schrieb sie bisher bei jeder Auswahl neu, und
    ' ein Schreibzugriff auf ScreenUpdating erzwingt jedes Mal ein Neuzeichnen.
    If appStateTouched Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    ElseIf Not Application.ScreenUpdating Then
        ' Heilnetz beibehalten: hat ein anderes Makro ScreenUpdating haengen lassen,
        ' stellt die naechste Auswahl es weiterhin zurueck. Der Lesezugriff ist
        ' billig, nur der bisher bedingungslose Schreibzugriff entfaellt.
        Application.ScreenUpdating = True
    End If
    mInternalChange = False
End Sub


Private Sub PID_RefreshFStundenOnSelectionChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim shouldRefresh As Boolean
    Dim currentEValue As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    If Target.CountLarge <> 1 Then Exit Sub
    
    shouldRefresh = False
    
    If mSelPrevSheet = ws.Name And mSelPrevCol = 5 Then
        If mSelPrevRow >= PID_FIRST_ROW And mSelPrevRow <= PID_LAST_ROW Then
            If Not (Target.Row = mSelPrevRow And Target.Column = 5) Then
                ' Nur neu aufbauen, wenn der KV-Code beim Verlassen der Zelle wirklich
                ' anders ist als beim Betreten. Reines Navigieren durch Spalte E hat
                ' sonst jedes Mal alle 80 F-Dropdowns neu aufgebaut. Der Neuaufbau
                ' bleibt sheetweit, weil sich die Helper-Spalten-Slots verschieben
                ' koennen, sobald sich die Menge der KV-Codes auf dem Blatt aendert.
                currentEValue = NormalizeKVCodeForLookup(CStr(ws.Cells(mSelPrevRow, "E").Value))
                If StrComp(currentEValue, mSelPrevEValueAtEntry, vbBinaryCompare) <> 0 Then
                    shouldRefresh = True
                End If
            End If
        End If
    End If
    
    If Not shouldRefresh Then
        If Target.Column = 6 And Target.Row >= PID_FIRST_ROW And Target.Row <= PID_LAST_ROW Then
            shouldRefresh = Not PID_RowHasValidFStundenDropdown(ws, Target.Row)
        End If
    End If
    
    If Not shouldRefresh Then Exit Sub
    
    PID_ExecuteFStundenSheetRefresh ws

SafeExit:
End Sub


Private Sub PID_ExecuteFStundenSheetRefresh(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    mInternalChange = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    PID_ForceRefreshFStundenDropdownForSheet ws
    mFStundenRefreshRanThisSelection = True

SafeExit:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub PID_SyncPrevEValueSnapshot(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If mSelPrevSheet <> ws.Name Then Exit Sub
    If mSelPrevCol <> 5 Then Exit Sub
    If mSelPrevRow < PID_FIRST_ROW Or mSelPrevRow > PID_LAST_ROW Then Exit Sub
    
    mSelPrevEValueAtEntry = NormalizeKVCodeForLookup(CStr(ws.Cells(mSelPrevRow, "E").Value))

SafeExit:
End Sub


Private Sub PID_UpdateMonthSelectionPrev(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Target.CountLarge <> 1 Then Exit Sub
    
    mSelPrevSheet = ws.Name
    mSelPrevRow = Target.Row
    mSelPrevCol = Target.Column
    
    If Target.Column = 5 And Target.Row >= PID_FIRST_ROW And Target.Row <= PID_LAST_ROW Then
        mSelPrevEValueAtEntry = NormalizeKVCodeForLookup(CStr(Target.Value))
    Else
        mSelPrevEValueAtEntry = ""
    End If

SafeExit:
End Sub


Public Sub PID_ResetInternalChangeFlag()
    ' Wird von CopyData am Anfang aufgerufen, um einen moeglichen haengengebliebenen
    ' mInternalChange=True-Zustand zu bereinigen (z. B. wenn ein Event-Handler durch
    ' einen Laufzeitfehler ohne SafeExit beendet wurde).
    mInternalChange = False
End Sub


Private Sub EnforcePasteValuesOnly(ByVal Sh As Object, ByVal Target As Range, ByRef pasteHandled As Boolean)
    Dim affectedRange As Range
    Dim pastedValues As Variant
    Dim c As Range
    Dim relRow As Long
    Dim relCol As Long
    Dim valueToWrite As Variant
    Dim handledMerges As Collection
    Dim mergeKey As String
    
    On Error GoTo SafeExit
    
    pasteHandled = False
    
    ' Strg+V laeuft schon ueber PID_PasteValuesOnly - dort sind die Werte bereits sauber.
    If PID_IsManagedPasteRunning() Then Exit Sub
    
    ' Ein echtes Einfuegen betrifft immer genau einen zusammenhaengenden Bereich.
    ' Bei mehreren Bereichen (z. B. Loeschen einer Mehrfachauswahl) liefert Target.Value
    ' nur den ersten Bereich - dann lieber gar nicht eingreifen.
    If Target.Areas.Count > 1 Then Exit Sub
    
    ' Sicherheitsnetz gegen das Einfuegen ganzer Spalten: der Wert-fuer-Wert-Aufbau
    ' waere dort zu langsam. Vor Application.Undo abbrechen, damit nichts verloren geht.
    If Target.CountLarge > PID_MAX_PASTE_CELLS Then Exit Sub
    
    Set affectedRange = Target
    
    If Not IsProbablyPaste(Target) Then Exit Sub
    
    pastedValues = Target.Value
    
    Application.Undo
    
    Set handledMerges = New Collection
    
    For Each c In affectedRange.Cells
        relRow = c.Row - Target.Row + 1
        relCol = c.Column - Target.Column + 1
        
        If Target.CountLarge = 1 Then
            valueToWrite = pastedValues
        Else
            valueToWrite = pastedValues(relRow, relCol)
        End If
        
        If c.MergeCells Then
            mergeKey = c.mergeArea.Address
            
            If Not PID_CollectionHasKey(handledMerges, mergeKey) Then
                handledMerges.Add mergeKey, mergeKey
                c.mergeArea.Cells(1, 1).Value = valueToWrite
            End If
        Else
            c.Value = valueToWrite
        End If
    Next c
    
    pasteHandled = True
    
SafeExit:
End Sub


Private Function IsProbablyPaste(ByVal Target As Range) As Boolean
    Dim undoText As String
    
    On Error Resume Next
    undoText = Application.CommandBars("Standard").Controls("&Undo").List(1)
    On Error GoTo 0
    
    undoText = LCase$(Trim$(CStr(undoText)))
    
    If Target.CountLarge > 1 Then
        IsProbablyPaste = True
        Exit Function
    End If
    
    If InStr(undoText, "paste") > 0 Then
        IsProbablyPaste = True
        Exit Function
    End If
    
    If InStr(undoText, "einf" & Chr(252) & "gen") > 0 Then
        IsProbablyPaste = True
        Exit Function
    End If
    
    If InStr(undoText, "einfuegen") > 0 Then
        IsProbablyPaste = True
        Exit Function
    End If
    
    If InStr(undoText, "beilleszt") > 0 Then
        IsProbablyPaste = True
        Exit Function
    End If
    
    If Application.CutCopyMode <> False Then
        IsProbablyPaste = True
        Exit Function
    End If
    
    IsProbablyPaste = False
End Function


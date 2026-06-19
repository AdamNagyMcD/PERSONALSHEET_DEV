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

Private mInternalChange As Boolean


Private mWorkerSheetRuntimeReady As Boolean


Private mSelPrevSheet As String
Private mSelPrevRow As Long
Private mSelPrevCol As Long
Private mSelPrevEValueAtEntry As String
Private mSelPrevFValueAtEntry As String
Private mFStundenRefreshRanThisSelection As Boolean


Private Sub PID_EnsureWorkerSheetRuntimeReady()
    If mWorkerSheetRuntimeReady Then Exit Sub
    
    PID_EnsureMonatslohnFormulas
    PID_EnsureLetztesGehaltFormulas
    mWorkerSheetRuntimeReady = True
End Sub


Private Sub Workbook_Open()
    Dim wsUbersicht As Worksheet
    
    On Error GoTo SafeExit
    
    mInternalChange = True
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Application.FormatStaleValues = False
    Err.Clear
    
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
    On Error Resume Next
    Application.EnableEvents = False
    mInternalChange = True
    PID_SyncAllMonthSheetsFluctuationQ31
    mInternalChange = False
    Application.EnableEvents = True
    On Error GoTo 0
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
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
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub Workbook_SheetActivate(ByVal Sh As Object)
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
            PID_BeginHeavyMaintenance
            RefreshKVDropdownsIfDirtyForSheet Sh
            RefreshKVLohnIfDirty Sh
            PID_EndHeavyMaintenance Sh
        Else
            PID_RecalculateMonthFormulaColumns Sh
        End If
        
        GoTo SafeExit
    End If
    
    PID_ClearMonthSheetFillHandleGuard

SafeExit:
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
    Application.FormatStaleValues = False
    Err.Clear
    
    ' Recalc-Abschnitt mit Resume Next: ein Mac-Fehler in Dropdown/Lohn-Recalc darf
    ' das E/F-Logging nicht verhindern. Windows: kein Unterschied, da keine Fehler auftreten.
    
    ' E zuerst: Dropdown in F muss sofort passen (Mac meldet Change oft nicht).
    If Not Intersect(Target, Sh.Range("E3:E82")) Is Nothing Then
        PID_RefreshFStundenDropdownForERows Sh, Target
    End If
    Err.Clear
    
    ' Einfuege-Schutz:
    ' Beim Einfuegen aus alten Dateien oder anderen Quellen soll nur der Wert uebernommen werden.
    pasteHandled = False
    EnforcePasteValuesOnly Sh, Target, pasteHandled
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
    
    On Error GoTo SafeExit
    
    mFStundenRefreshRanThisSelection = False
    
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If PID_IsWorkerMonthSheet(Sh) And Target.CountLarge = 1 Then
        PID_RefreshFStundenOnSelectionChange Sh, Target
    End If
    
    If mInternalChange Then GoTo UpdateSelectionPrev
    If Not PID_IsWorkerMonthSheet(Sh) Then GoTo UpdateSelectionPrev
    If Target.CountLarge > 1 Then GoTo UpdateSelectionPrev
    
    ' Mac-Backup: F-Stunden-Logging via SelectionChange.
    ' SheetChange kann auf Mac bei Dropdown-Aenderungen ausfallen oder vor dem Logging abbrechen.
    ' Wenn der Nutzer eine F-Zelle verlaesst und der Wert sich geaendert hat, wird hier geloggt.
    ' Windows: keine Auswirkung — SheetChange funktioniert zuverlaessig.
    If PID_IsMacExcel() And Not gCopyDataRunning Then
        PID_LogFChangeOnMacIfNeeded Sh, Target
    End If
    
    ' Ausserhalb D3:F82: O45 nur loeschen wenn noetig (leichter Pfad).
    If Intersect(Target, Sh.Range("D3:F82")) Is Nothing Then
        If Len(Trim$(CStr(Sh.Range("O45").Value))) = 0 Then GoTo UpdateSelectionPrev
        
        mInternalChange = True
        Application.EnableEvents = False
        Sh.Range("O45").Value = ""
        GoTo UpdateSelectionPrev
    End If
    
    If Target.Row < PID_FIRST_ROW Or Target.Row > PID_LAST_ROW Then
        If Len(Trim$(CStr(Sh.Range("O45").Value))) > 0 Then
            mInternalChange = True
            Application.EnableEvents = False
            Sh.Range("O45").Value = ""
        End If
        GoTo UpdateSelectionPrev
    End If
    
    ' D: Beschaeftigungsdauer — ohne ScreenUpdating-Off.
    If Not Intersect(Target, Sh.Range("D3:D82")) Is Nothing Then
        mInternalChange = True
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
        Application.EnableEvents = False
        Sh.Range("O45").Value = ""
        Application.EnableEvents = True
        mInternalChange = False
    End If
    
    ' E/F ungueltig: reparieren.
    If Not needsDropdownRepair Then GoTo UpdateSelectionPrev
    If mFStundenRefreshRanThisSelection Then GoTo UpdateSelectionPrev
    
    mInternalChange = True
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
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub PID_RefreshFStundenOnSelectionChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim shouldRefresh As Boolean
    Dim currentE As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    If Target.CountLarge <> 1 Then Exit Sub
    
    ' Mac: SheetChange bei E-Dropdown oft aus — Refresh bei E-Wechsel / F-Fokus erzwingen.
    If PID_IsMacExcel() Then
        shouldRefresh = False
        
        If mSelPrevSheet = ws.Name And mSelPrevCol = 5 Then
            If mSelPrevRow >= PID_FIRST_ROW And mSelPrevRow <= PID_LAST_ROW Then
                If Not (Target.Row = mSelPrevRow And Target.Column = 5) Then
                    shouldRefresh = True
                    
                    currentE = NormalizeKVCodeForLookup(CStr(ws.Cells(mSelPrevRow, "E").Value))
                    If Len(mSelPrevEValueAtEntry) > 0 Then
                        If StrComp(currentE, mSelPrevEValueAtEntry, vbBinaryCompare) <> 0 Then
                            shouldRefresh = True
                        End If
                    End If
                End If
            End If
        End If
        
        If Target.Column = 6 And Target.Row >= PID_FIRST_ROW And Target.Row <= PID_LAST_ROW Then
            shouldRefresh = True
        End If
    Else
        shouldRefresh = False
        
        If mSelPrevSheet = ws.Name And mSelPrevCol = 5 Then
            If mSelPrevRow >= PID_FIRST_ROW And mSelPrevRow <= PID_LAST_ROW Then
                If Not (Target.Row = mSelPrevRow And Target.Column = 5) Then
                    shouldRefresh = True
                End If
            End If
        End If
        
        If Not shouldRefresh Then
            If Target.Column = 6 And Target.Row >= PID_FIRST_ROW And Target.Row <= PID_LAST_ROW Then
                shouldRefresh = Not PID_RowHasValidFStundenDropdown(ws, Target.Row)
            End If
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


Private Sub PID_UpdateMonthSelectionPrev(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Target.CountLarge <> 1 Then Exit Sub
    
    mSelPrevSheet = ws.Name
    mSelPrevRow = Target.Row
    mSelPrevCol = Target.Column
    
    If Target.Column = 5 And Target.Row >= PID_FIRST_ROW And Target.Row <= PID_LAST_ROW Then
        mSelPrevEValueAtEntry = NormalizeKVCodeForLookup(CStr(ws.Cells(Target.Row, "E").Value))
    End If
    
    If Target.Column = 6 And Target.Row >= PID_FIRST_ROW And Target.Row <= PID_LAST_ROW Then
        mSelPrevFValueAtEntry = Trim$(CStr(ws.Cells(Target.Row, "F").Value))
    End If

SafeExit:
End Sub


Public Sub PID_ResetInternalChangeFlag()
    ' Wird von CopyData am Anfang aufgerufen, um einen moeglichen haengengebliebenen
    ' mInternalChange=True-Zustand zu bereinigen (kann auf Mac passieren wenn ein
    ' Event-Handler durch einen Mac-spezifischen Fehler ohne SafeExit beendet wurde).
    ' Windows: harmlos (mInternalChange ist bereits False).
    mInternalChange = False
End Sub


Public Sub PID_FlushPendingEFLog()
    ' Wird von CopyData vor dem Starten aufgerufen.
    ' Problem: Wenn der Nutzer auf Mac eine F-Zelle aendert und sofort den CopyData-Button
    ' drueckt (ohne eine andere Zelle zu klicken), feuert SelectionChange nicht mehr.
    ' => Die Aenderung waere nie in den Log gelangt.
    ' Loesung: CopyData ruft diese Funktion auf, die den aktuellen F-Wert mit dem
    ' Eintritts-Snapshot (mSelPrevFValueAtEntry) vergleicht und bei Abweichung loggt.
    ' Windows: PID_IsMacExcel()=False => sofortiger Exit, kein Unterschied.
    
    Dim ws As Worksheet
    Dim currentVal As String
    
    On Error GoTo SafeExit
    
    If Not PID_IsMacExcel() Then Exit Sub
    If mSelPrevCol <> 6 And mSelPrevCol <> 5 Then Exit Sub
    If mSelPrevSheet = "" Then Exit Sub
    If mSelPrevRow < PID_FIRST_ROW Or mSelPrevRow > PID_LAST_ROW Then Exit Sub
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(mSelPrevSheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    currentVal = Trim$(CStr(ws.Cells(mSelPrevRow, mSelPrevCol).Value))
    
    If mSelPrevCol = 6 Then
        If currentVal = mSelPrevFValueAtEntry Then Exit Sub
    ElseIf mSelPrevCol = 5 Then
        If NormalizeKVCodeForLookup(currentVal) = mSelPrevEValueAtEntry Then Exit Sub
    End If
    
    If currentVal = "" Then Exit Sub
    
    PID_LogEFAenderungForSheet ws, ws.Cells(mSelPrevRow, mSelPrevCol)

SafeExit:
End Sub


Private Sub PID_LogFChangeOnMacIfNeeded(ByVal ws As Worksheet, ByVal Target As Range)
    Dim currentF As String
    
    On Error GoTo SafeExit
    
    ' Nur aktiv wenn vorherige Selektion eine F-Zelle in der Mitarbeiter-Range war.
    If mSelPrevCol <> 6 Then Exit Sub
    If mSelPrevSheet <> ws.Name Then Exit Sub
    If mSelPrevRow < PID_FIRST_ROW Or mSelPrevRow > PID_LAST_ROW Then Exit Sub
    
    ' Wenn der Nutzer noch auf der gleichen F-Zelle ist, nichts tun.
    If Target.Row = mSelPrevRow And Target.Column = 6 Then Exit Sub
    
    currentF = Trim$(CStr(ws.Cells(mSelPrevRow, "F").Value))
    
    ' Nur loggen wenn der Wert sich gegenueber dem Eintritts-Snapshot geaendert hat.
    If currentF = mSelPrevFValueAtEntry Then Exit Sub
    If currentF = "" Then Exit Sub
    
    PID_LogEFAenderungForSheet ws, ws.Cells(mSelPrevRow, "F")

SafeExit:
End Sub


Private Sub EnforcePasteValuesOnly(ByVal Sh As Object, ByVal Target As Range, ByRef pasteHandled As Boolean)
    Dim inputRange As Range
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
    
    Set inputRange = Union(Sh.Range("B3:C82"), Sh.Range("D3:D82"), Sh.Range("E3:F82"), Sh.Range("I3:J82"), Sh.Range("M3:N82"), Sh.Range("O18:Q25"))
    
    If PID_IsWorkerMonthSheet(Sh) And PID_IsMonthSheetVormonatInputEditable(Sh) Then
        Set inputRange = Union(inputRange, Sh.Range("Q12:R12"))
    End If
    
    Set affectedRange = Intersect(Target, inputRange)
    If affectedRange Is Nothing Then Exit Sub
    
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


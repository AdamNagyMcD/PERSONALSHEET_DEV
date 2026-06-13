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
    
    ' Wichtig:
    ' Beim Oeffnen wird die vollstaendige Fluktuation-Analyse NICHT aktualisiert.
    ' Das ermoeglicht schnelleres Oeffnen - aktualisiert wird nur wenn wirklich noetig.
    MarkFluktuationDirty
    
    On Error Resume Next
    ThisWorkbook.Sheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("KV_DROPDOWN_HELPER").Visible = xlSheetVeryHidden
    On Error GoTo SafeExit

SafeExit:
    PID_ConfigureDeferredWorkbookCalculationOnOpen
    
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
    
    PID_ApplyCopyrightToAllSheets
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
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
    On Error GoTo SafeExit
    
    ' Einfuege-Schutz:
    ' Beim Einfuegen aus alten Dateien oder anderen Quellen soll nur der Wert uebernommen werden.
    pasteHandled = False
    EnforcePasteValuesOnly Sh, Target, pasteHandled
    
    ' Wenn Spalte E sich aendert:
    ' E = KV-Code / BG -> Dropdown F in der betreffenden Zeile.
    If Not Intersect(Target, Sh.Range("E3:F82")) Is Nothing Then
        PID_RecalculateMonatslohnForChangedRows Sh, Target
    End If
    
    If Not Intersect(Target, Sh.Range("D3:F82,I3:I82")) Is Nothing Then
        PID_RecalculateAktuelleStundenForChangedRows Sh, Target
    End If
    
    If Not Intersect(Target, Sh.Range("D3:D82,I3:I82,K3:K82")) Is Nothing Then
        PID_RecalculateLetztesGehaltForChangedRows Sh, Target
    End If
    
    If Not Intersect(Target, Sh.Range("F3:F82")) Is Nothing Then
        PID_LogEFAenderungForSheet Sh, Target
    End If
    
    If Not Intersect(Target, Sh.Range("E3:E82")) Is Nothing Then
        PID_RefreshFStundenDropdownForERows Sh, Target
    End If
    
    ' Fluktuation-Markierung fuer Aktualisierung:
    ' D = Eintrittsdatum
    ' I = Austrittsdatum
    ' N = Austrittsgrund
    If Not Intersect(Target, Sh.Range("D3:D82,I3:I82,N3:N82")) Is Nothing Then
        MarkFluktuationDirtyForMonthSheet Sh
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
    
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If PID_IsWorkerMonthSheet(Sh) And Target.CountLarge = 1 Then
        PID_RefreshFStundenAfterLeavingE Sh, Target
    End If
    
    If mInternalChange Then GoTo UpdateSelectionPrev
    If Not PID_IsWorkerMonthSheet(Sh) Then GoTo UpdateSelectionPrev
    If Target.CountLarge > 1 Then GoTo UpdateSelectionPrev
    
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
        If mSelPrevSheet = Sh.Name And mSelPrevCol = 5 And mSelPrevRow = Target.Row Then
            fDropdownInvalid = True
        End If
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
    
    ' E->F: leave-E hat bereits ganzes Blatt refreshed; sonst nur E-Dropdown reparieren.
    If fDropdownInvalid And mSelPrevSheet = Sh.Name And mSelPrevCol = 5 And mSelPrevRow = Target.Row Then
        GoTo UpdateSelectionPrev
    End If
    
    ' E/F gueltig: sofort beenden — kein ScreenUpdating-Off.
    If Not needsDropdownRepair Then GoTo UpdateSelectionPrev
    
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
        PID_ClearStundenValuesCache
        RefreshKVStundenDropdownForSheet Sh, , True
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


Private Sub PID_RefreshFStundenAfterLeavingE(ByVal ws As Worksheet, ByVal Target As Range)
    Dim leftERow As Long
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If mInternalChange Then Exit Sub
    
    If mSelPrevSheet <> ws.Name Then Exit Sub
    If mSelPrevCol <> 5 Then Exit Sub
    
    leftERow = mSelPrevRow
    If leftERow < PID_FIRST_ROW Or leftERow > PID_LAST_ROW Then Exit Sub
    
    If Target.Row = leftERow And Target.Column = 5 Then Exit Sub
    
    mInternalChange = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    PID_ClearStundenValuesCache
    RefreshKVStundenDropdownForSheet ws, , True

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


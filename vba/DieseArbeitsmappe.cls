VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private mInternalChange As Boolean


Private mWorkerSheetRuntimeReady As Boolean


Private Sub PID_EnsureWorkerSheetRuntimeReady()
    If mWorkerSheetRuntimeReady Then Exit Sub
    
    PID_EnsureMonatslohnFormulas
    PID_EnsureLetztesGehaltFormulas
    mWorkerSheetRuntimeReady = True
End Sub


Private Sub Workbook_Open()
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
    
    If Not ActiveSheet Is Nothing Then
        If PID_IsWorkerMonthSheet(ActiveSheet) Then
            PID_EnsureWorkerSheetRuntimeReady
        End If
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo SafeExit
    
    If mInternalChange Then Exit Sub
    
    mInternalChange = True
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Vor dem Speichern wird die Fluktuation-Analyse nur dann aktualisiert,
    ' wenn sich in der Zwischenzeit relevante Daten geaendert haben.
    RefreshFluktuationIfDirty
    
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
    
    PID_EnsureSheetProtectionForMacros Sh
    
    ' Fluktuation beim Oeffnen des Analyse-Blatts immer neu aufbauen.
    If Sh.Name = PID_FLUKTUATION_SHEET Then
        RefreshFluktuationIfDirty
        GoTo SafeExit
    End If
    
    If Sh.Name = PID_LOHNTABELLE_SHEET Then
        EnsureAddNewKVPeriodButton
        FixLOHNTABELLE_HeaderTextIfNeeded
        PID_ConfigureLohntabelleView
        GoTo SafeExit
    End If
    
    If Sh.Name = "UBERSICHT" Then
        RefreshFinanzSummaryIfDirty
        GoTo SafeExit
    End If
    
    ' Dropdown/Lohn nur bei Bedarf (nach KV-Tabelle-Aenderung), pro Monat einmal.
    If PID_IsWorkerMonthSheet(Sh) Then
        PID_EnsureWorkerSheetRuntimeReady
        RefreshKVDropdownsIfDirtyForSheet Sh
        RefreshKVLohnIfDirty Sh
        GoTo SafeExit
    End If

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
            MarkAllKVDropdownsDirty
            MarkAllKVLohnDirty
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
        PID_PreloadKVLohnCaches
        PID_RecalculateMonatslohnForChangedRows Sh, Target
    End If
    
    If Not Intersect(Target, Sh.Range("D3:D82,I3:I82,K3:K82")) Is Nothing Then
        PID_RecalculateLetztesGehaltForChangedRows Sh, Target
    End If
    
    If Not Intersect(Target, Sh.Range("F3:F82")) Is Nothing Then
        PID_LogEFAenderungForSheet Sh, Target
    End If
    
    If Not Intersect(Target, Sh.Range("E3:E82")) Is Nothing Then
        PID_InvalidateFStundenDropdownForRows Sh, Target
    End If
    
    ' Fluktuation-Markierung fuer Aktualisierung:
    ' D = Eintrittsdatum
    ' I = Austrittsdatum
    ' N = Austrittsgrund
    If Not Intersect(Target, Sh.Range("D3:D82,I3:I82,N3:N82")) Is Nothing Then
        MarkFluktuationDirty
    End If
    
    If PID_MonthChangeNeedsImmediateFinanzSync(Sh, Target) Then
        PID_SyncFinanzSummaryForMonth Sh
    ElseIf PID_MonthChangeDefersFinanzSummarySync(Sh, Target) Then
        MarkFinanzSummaryDirty
    End If

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    
    On Error GoTo SafeExit
    
    If mInternalChange Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If Not PID_IsWorkerMonthSheet(Sh) Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub
    
    mInternalChange = True
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' If user selects E and dropdown is broken/missing, rebuild column E validation.
    If Not Intersect(Target, Sh.Range("E3:E82")) Is Nothing Then
        If Not PID_MonthSheetHasValidKVCodeDropdown(Sh) Then
            On Error Resume Next
            PID_RestoreKVCodeDropdownOnSheet Sh
            On Error GoTo SafeExit
        End If
    End If
    
    ' If user selects F and dropdown is missing/broken, rebuild only that row.
    If Not Intersect(Target, Sh.Range("F3:F82")) Is Nothing Then
        If Not PID_RowHasValidFStundenDropdown(Sh, Target.Row) Then
            If Trim$(CStr(Sh.Cells(Target.Row, "E").Value)) <> "" Then
                On Error Resume Next
                RefreshKVStundenDropdownForSingleRow Sh, Target.Row
                On Error GoTo SafeExit
            End If
        End If
    End If
    
    ' Nur Eintrittsdatum Spalte D.
    If Intersect(Target, Sh.Range("D3:D82")) Is Nothing Then
        
        ' Nur loeschen, wenn tatsaechlich etwas drin ist.
        If Sh.Range("O45").Value <> "" Then
            Sh.Range("O45").Value = ""
        End If
        
        GoTo SafeExit
    End If
    
    PID_ShowEmploymentDuration Sh, Target

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
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
    
    Set inputRange = Union(Sh.Range("B3:F82"), Sh.Range("I3:J82"), Sh.Range("M3:N82"), Sh.Range("O18:Q25"))
    
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
            
            If Not CollectionHasKeyLocal(handledMerges, mergeKey) Then
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


Private Function CollectionHasKeyLocal(ByVal col As Collection, ByVal key As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    tmp = col.item(key)
    
    CollectionHasKeyLocal = True
    Exit Function

NotFound:
    CollectionHasKeyLocal = False
End Function


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


Private Sub Workbook_Open()
    On Error GoTo SafeExit
    
    mInternalChange = True
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    PID_SetupSheetProtectionForMacros
    
    ' Wichtig:
    ' Beim Oeffnen wird die vollstaendige Fluktuation-Analyse NICHT aktualisiert.
    ' Das ermoeglicht schnelleres Oeffnen - aktualisiert wird nur wenn wirklich noetig.
    gFluktuationDirty = True
    
    On Error Resume Next
    ThisWorkbook.Sheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("KV_DROPDOWN_HELPER").Visible = xlSheetVeryHidden
    On Error GoTo SafeExit

SafeExit:
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
    If gFluktuationDirty Then
        RefreshFluktuationAll
    End If
    
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
    
    ' Fluktuation wird nur dann aktualisiert, wenn der Benutzer das Fluktuation-Blatt oeffnet.
    If Sh.Name = "Fluktuation" Then
        RefreshFluktuationIfDirty
        GoTo SafeExit
    End If
    
    ' Bei Aktivierung eines Monatsblatts wird KEINE vollstaendige KV-Lohn-Aktualisierung ausgefuehrt.
    ' Das war einer der Leistungsengpaesse.
    '
    ' Wenn die KV-Tabelle geaendert wurde, werden die Dropdowns nur per Dirty-Flag aktualisiert.
    If PID_IsWorkerMonthSheet(Sh) Then
        RefreshKVDropdownsIfDirty
        GoTo SafeExit
    End If

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mInternalChange = False
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim pasteHandled As Boolean
    
    On Error GoTo SafeExit
    
    If mInternalChange Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    
    ' Wenn die KV-Basistabelle geaendert wird, wird die Dropdown-Liste spaeter neu aufgebaut.
    If Sh.Name = "LOHNTABELLE_TEST" Then
        If Not Intersect(Target, Sh.Range("G:H")) Is Nothing Then
            MarkKVDropdownsDirty
        End If
        Exit Sub
    End If
    
    If Not PID_IsWorkerMonthSheet(Sh) Then Exit Sub
    
    mInternalChange = True
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Einfuege-Schutz:
    ' Beim Einfuegen aus alten Dateien oder anderen Quellen soll nur der Wert uebernommen werden.
    pasteHandled = False
    EnforcePasteValuesOnly Sh, Target, pasteHandled
    
    ' Wenn Spalte E sich aendert:
    ' E = KV-Code / BG
    ' Dropdown der Spalte F wird in der betreffenden Zeile aktualisiert.
    If Not Intersect(Target, Sh.Range("E3:E82")) Is Nothing Then
        RefreshKVStundenDropdownForSheet Sh, Target
    End If
    
    ' Wenn E oder F sich aendert:
    ' G = Monatslohn wird nur in der betreffenden Zeile aktualisiert.
    If Not Intersect(Target, Sh.Range("E3:F82")) Is Nothing Then
        RefreshKVLohnForSheet Sh, Target
    End If
    
    ' Fluktuation-Markierung fuer Aktualisierung:
    ' D = Eintrittsdatum
    ' I = Austrittsdatum
    ' N = Austrittsgrund
    If Not Intersect(Target, Sh.Range("D3:D82,I3:I82,N3:N82")) Is Nothing Then
        PID_CalculateFluctuation Sh
        MarkFluktuationDirty
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
    
    Set inputRange = Union( _
        Sh.Range("B3:F82"), _
        Sh.Range("I3:J82"), _
        Sh.Range("M3:N82"), _
        Sh.Range("O18:Q25") _
    )
    
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


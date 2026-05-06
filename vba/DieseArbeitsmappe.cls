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
    
    ' Fontos:
    ' Megnyitaskor NEM frissitunk teljes Fluktuation elemzest.
    ' Ez gyorsabb megnyitast ad, es csak akkor frissitunk, ha tenyleg kell.
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
    
    ' Mentes elott csak akkor frissitjuk a Fluktuation elemzest,
    ' ha kozben relevans adat valtozott.
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
    
    ' Fluktuation csak akkor frissul, amikor a felhasznalo megnyitja a Fluktuation lapot.
    If Sh.Name = "Fluktuation" Then
        RefreshFluktuationIfDirty
        GoTo SafeExit
    End If
    
    ' Honaplap aktivacional NEM futtatunk teljes KV-lohn frissitest.
    ' Ez volt az egyik lassito pont.
    '
    ' Ha a KV tabla valtozott, a dropdownokat csak dirty flag alapjan frissitjuk.
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
    
    ' Ha a KV alap tabla valtozik, akkor a dropdown lista kesobb ujraepul.
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
    
    ' Beillesztes-vedelem:
    ' regi fajlbol vagy mas forrasbol torteno masolasnal csak ertek maradjon.
    pasteHandled = False
    EnforcePasteValuesOnly Sh, Target, pasteHandled
    
    ' Ha E oszlop valtozik:
    ' E = KV-Code / BG
    ' F oszlop dropdownja frissul az adott sorban.
    If Not Intersect(Target, Sh.Range("E3:E82")) Is Nothing Then
        RefreshKVStundenDropdownForSheet Sh, Target
    End If
    
    ' Ha E vagy F valtozik:
    ' G = Monatslohn frissul csak az adott sorban.
    If Not Intersect(Target, Sh.Range("E3:F82")) Is Nothing Then
        RefreshKVLohnForSheet Sh, Target
    End If
    
    ' Fluktuation jeloles frissitesre:
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
    
    ' Csak Eintrittsdatum oszlop D.
    If Intersect(Target, Sh.Range("D3:D82")) Is Nothing Then
        
        ' Csak akkor toroljuk, ha tenyleg van benne valami.
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
    
    If InStr(undoText, "einfugen") > 0 Then
        IsProbablyPaste = True
        Exit Function
    End If
    
    If InStr(undoText, "einfŸgen") > 0 Then
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


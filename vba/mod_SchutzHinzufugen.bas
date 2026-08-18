Attribute VB_Name = "mod_SchutzHinzufugen"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Private mSessionProtectedSheets As Collection
Private mSavedEnableFillHandle As Variant
Private mSavedCellDragAndDrop As Variant
Private mMonthSheetFillHandleGuardActive As Boolean

Private Const PID_MONTH_PANEL_VORMONAT_RANGE As String = "Q12:R12"
Private Const PID_MONTH_PANEL_FREITEXT_RANGE As String = "O18:Q28"
Private Const PID_UBERSICHT_SHEET As String = "UBERSICHT"
Private Const PID_UBERSICHT_JAEN_VERF_CELL As String = "E30"
Private Const PID_UBERSICHT_JAEN_MUST_CELL As String = "I30"
Private Const PID_UBERSICHT_AER_VERF_TITLE As String = "PID_Plan_Verfuegbar"
Private Const PID_UBERSICHT_AER_MUST_TITLE As String = "PID_Plan_Muster"
' XlEnableSelection — bewusst numerisch (xlUnlockedCell = 1).
Private Const PID_XL_ENABLE_SELECTION_UNLOCKED As Long = 1


Public Sub PID_FixUbersichtPlanInputsEditable()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Set ws = PID_GetUbersichtWorksheet()
    If ws Is Nothing Then
        MsgBox "Blatt UEBERSICHT wurde nicht gefunden.", vbExclamation, "UEBERSICHT"
        Exit Sub
    End If
    
    PID_ApplyUbersichtSheetProtection ws
    
    MsgBox "UEBERSICHT " & PID_UTxtGeschuetzt() & ": nur E30 und I30 sind editierbar." & vbCrLf & vbCrLf & _
           "Wei" & PID_UTxtSs() & "e Plan-Zellen: Spalte E und I, Zeile 30 — Formeln darunter sind gesperrt.", _
           vbInformation, "UEBERSICHT"

SafeExit:
End Sub


Public Sub PID_ApplyMonthSheetLockPolicy(ByVal ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheetSafe(ws) Then Exit Sub
    
    ws.Cells.Locked = True
    
    ws.Range("B" & PID_FIRST_ROW & ":C" & PID_LAST_ROW).Locked = False
    ws.Range("D" & PID_FIRST_ROW & ":D" & PID_LAST_ROW).Locked = False
    ws.Range("E" & PID_FIRST_ROW & ":F" & PID_LAST_ROW).Locked = False
    ws.Range("I" & PID_FIRST_ROW & ":J" & PID_LAST_ROW).Locked = False
    ws.Range("M" & PID_FIRST_ROW & ":N" & PID_LAST_ROW).Locked = False
    PID_UnlockMonthPanelFreitextRange ws
    If PID_IsMonthSheetVormonatInputEditable(ws) Then
        ws.Range(PID_MONTH_PANEL_VORMONAT_RANGE).Locked = False
    End If
    
    Err.Clear
End Sub


Public Function PID_IsDurchrechnungPeriodStartMonthSheet(ByVal ws As Worksheet) As Boolean
    Dim sheetName As String
    
    If ws Is Nothing Then Exit Function
    
    sheetName = Trim$(CStr(ws.Name))
    
    Select Case sheetName
        Case "Februar", "Mai", "August", "November"
            PID_IsDurchrechnungPeriodStartMonthSheet = True
    End Select
End Function


Public Function PID_IsMonthSheetVormonatInputEditable(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheetSafe(ws) Then Exit Function
    
    PID_IsMonthSheetVormonatInputEditable = Not PID_IsDurchrechnungPeriodStartMonthSheet(ws)
End Function


Public Sub PID_UnlockSheetEditRanges(ByVal ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    
    PID_ForceUnprotectWorksheet ws
    
    If PID_IsWorkerMonthSheetSafe(ws) Then
        PID_ApplyMonthSheetLockPolicy ws
    ElseIf PID_IsUbersichtSheet(ws) Then
        PID_ApplyUbersichtEditableCells ws
    End If
    
    Err.Clear
End Sub


Public Sub PID_ApplyUbersichtSheetProtection(ByVal ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsUbersichtWorksheet(ws) Then Exit Sub
    
    PID_ForceUnprotectWorksheet ws
    
    ws.ScrollArea = ""
    ws.Cells.Locked = True
    
    If PID_DurchrechnungBlockExists(ws) Then
        PID_EnsureUbersichtDurchrechnungInputsEditable ws
    End If
    
    ws.Range(PID_UBERSICHT_JAEN_VERF_CELL).Locked = False
    ws.Range(PID_UBERSICHT_JAEN_MUST_CELL).Locked = False
    
    PID_ClearUbersichtAllowEditRanges ws
    PID_SetupUbersichtAllowEditRanges ws
    
    ' Nur E30/I30 anklickbar — Formeln (FINANZ, Durchrechnung) nicht auswaehlbar.
    ws.EnableSelection = PID_XL_ENABLE_SELECTION_UNLOCKED
    
    ' AllowSelectingLockedCells wird bewusst nicht gesetzt — Auswahl steuert EnableSelection.
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=False
    
    Err.Clear
End Sub


Public Sub PID_ReprotectWorksheet(ByVal ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    
    If PID_IsUbersichtWorksheet(ws) Then
        PID_ApplyUbersichtSheetProtection ws
        Exit Sub
    End If
    
    PID_ForceUnprotectWorksheet ws
    PID_UnlockSheetEditRanges ws
    PID_ApplySheetProtectionForMacros ws
End Sub


Public Sub PID_SetupSheetProtectionForMacros()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Application.ScreenUpdating = False
    Set mSessionProtectedSheets = New Collection
    
    For Each ws In ThisWorkbook.Worksheets
        If PID_IsUbersichtWorksheet(ws) Then
            PID_ApplyUbersichtSheetProtection ws
        Else
            PID_ApplySheetProtectionForMacros ws
        End If
        PID_MarkSheetProtectionReady ws.Name
    Next ws

SafeExit:
    Application.ScreenUpdating = True
End Sub


Public Sub PID_SetupSheetProtectionForOpen()
    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo SafeExit
    
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set mSessionProtectedSheets = New Collection
    
    ' Monatsblaetter gehoeren mit dazu. Frueher wurden sie erst beim ersten Besuch des
    ' jeweiligen Tabs geschuetzt (PID_EnsureSheetProtectionForMacros im
    ' Workbook_SheetActivate). Wer die Mappe also auf UEBERSICHT oder EINSTELLUNG
    ' liegend geoeffnet hat, konnte in allen zwoelf Monatsblaettern alles ueberschreiben -
    ' auch die Formelspalten G, H, K und L - und Zeilen einfuegen oder loeschen. Der
    ' Schutz selbst kostet je Blatt nur die Sperrliste plus ein Protect; dafuer sind die
    ' Blaetter beim ersten Tabwechsel bereits fertig und die Sitzungsmarkierung
    ' verhindert ein zweites Setzen.
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "FLUKTUATION_DATEN" Or ws.Name = "KV_DROPDOWN_HELPER" _
           Or PID_IsWorkerMonthSheetSafe(ws) Then
            PID_ApplySheetProtectionForMacros ws
            PID_MarkSheetProtectionReady ws.Name
        End If
    Next ws
    
    Set ws = PID_GetUbersichtWorksheet()
    If Not ws Is Nothing Then
        PID_ApplyUbersichtSheetProtection ws
        PID_MarkSheetProtectionReady ws.Name
    End If
    
    If Not ActiveSheet Is Nothing Then
        PID_EnsureSheetProtectionForMacros ActiveSheet
    End If

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
End Sub


Public Sub PID_EnsureSheetProtectionForMacros(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If PID_SheetProtectionReady(ws.Name) Then Exit Sub
    
    PID_ApplySheetProtectionForMacros ws
    PID_MarkSheetProtectionReady ws.Name

SafeExit:
End Sub


Private Sub PID_ApplySheetProtectionForMacros(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    PID_ForceUnprotectWorksheet ws
    PID_UnlockSheetEditRanges ws
    
    If PID_IsUbersichtSheet(ws) Then
        PID_ApplyUbersichtSheetProtection ws
        GoTo SafeExit
    
    ElseIf PID_IsWorkerMonthSheetSafe(ws) Then
        
        PID_ProtectWorkerMonthSheet ws
        
    ElseIf ws.Name = PID_LOHNTABELLE_SHEET Then
        
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
        
    ElseIf ws.Name = "FLUKTUATION_DATEN" Then
        
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True
        
        ws.Visible = xlSheetVeryHidden
        
    ElseIf ws.Name = "KV_DROPDOWN_HELPER" Then
        
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True
        
        ws.Visible = xlSheetVeryHidden
        
    ElseIf StrComp(ws.Name, PID_ADMIN_SHEET_NAME, vbTextCompare) = 0 Then
        
        PID_ForceUnprotectWorksheet ws
        
    Else
        
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
        
    End If

SafeExit:
End Sub


Public Sub PID_ProtectWorkerMonthSheet(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheetSafe(ws) Then Exit Sub
    
    PID_ApplyMonthSheetLockPolicy ws
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=False

SafeExit:
End Sub


Public Sub PID_ApplyMonthSheetFillHandleGuard()
    On Error Resume Next
    
    If mMonthSheetFillHandleGuardActive Then Exit Sub
    
    mSavedEnableFillHandle = Application.EnableFillHandle
    mSavedCellDragAndDrop = Application.CellDragAndDrop
    
    Application.EnableFillHandle = False
    Application.CellDragAndDrop = False
    
    mMonthSheetFillHandleGuardActive = True
    Err.Clear
End Sub


' forceRestore ist fuer PID_ResetExcelState gedacht: nach einem Reset des VBA-Projekts
' sind die Modulvariablen leer, die Einstellungen der Anwendung aber noch aus - dann
' meldet der Merker "kein Schutz aktiv", waehrend Ausfuellkaestchen und Ziehen weiterhin
' fehlen. Im Normalbetrieb (Blattwechsel) bleibt das Verhalten unveraendert.
Public Sub PID_ClearMonthSheetFillHandleGuard(Optional ByVal forceRestore As Boolean = False)
    On Error Resume Next
    
    If Not mMonthSheetFillHandleGuardActive And Not forceRestore Then Exit Sub
    
    If Not IsEmpty(mSavedEnableFillHandle) Then
        Application.EnableFillHandle = CBool(mSavedEnableFillHandle)
    Else
        Application.EnableFillHandle = True
    End If
    
    If Not IsEmpty(mSavedCellDragAndDrop) Then
        Application.CellDragAndDrop = CBool(mSavedCellDragAndDrop)
    Else
        Application.CellDragAndDrop = True
    End If
    
    mMonthSheetFillHandleGuardActive = False
    mSavedEnableFillHandle = Empty
    mSavedCellDragAndDrop = Empty
    Err.Clear
End Sub


Private Sub PID_ApplyUbersichtEditableCells(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not PID_IsUbersichtSheet(ws) Then Exit Sub
    
    ws.Cells.Locked = True
    PID_EnsureUbersichtDurchrechnungInputsEditable ws
    ws.Range(PID_UBERSICHT_JAEN_VERF_CELL).Locked = False
    ws.Range(PID_UBERSICHT_JAEN_MUST_CELL).Locked = False
End Sub


Private Sub PID_ClearUbersichtAllowEditRanges(ByVal ws As Worksheet)
    Dim i As Long
    
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    
    For i = ws.Protection.AllowEditRanges.Count To 1 Step -1
        ws.Protection.AllowEditRanges(i).Delete
    Next i
    
    Err.Clear
End Sub


Private Sub PID_SetupUbersichtAllowEditRanges(ByVal ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsUbersichtSheet(ws) Then Exit Sub
    
    ws.Protection.AllowEditRanges.Add Title:=PID_UBERSICHT_AER_VERF_TITLE, _
        Range:=ws.Range(PID_UBERSICHT_JAEN_VERF_CELL)
    ws.Protection.AllowEditRanges.Add Title:=PID_UBERSICHT_AER_MUST_TITLE, _
        Range:=ws.Range(PID_UBERSICHT_JAEN_MUST_CELL)
    
    Err.Clear
End Sub


Private Sub PID_ForceUnprotectWorksheet(ByVal ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    
    If ws.ProtectContents Then
        Err.Clear
        ws.Unprotect
    End If
    
    Err.Clear
End Sub


Public Function PID_IsUbersichtWorksheet(ByVal ws As Worksheet) As Boolean
    Dim sheetKey As String
    
    If ws Is Nothing Then Exit Function
    
    sheetKey = PID_NormalizeSheetNameKey(ws.Name)
    PID_IsUbersichtWorksheet = (sheetKey = "UBERSICHT")
End Function


Public Function PID_GetUbersichtWorksheet() As Worksheet
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If PID_IsUbersichtWorksheet(ws) Then
            Set PID_GetUbersichtWorksheet = ws
            Exit Function
        End If
    Next ws
End Function


Private Function PID_NormalizeSheetNameKey(ByVal sheetName As String) As String
    Dim s As String
    
    s = Trim$(CStr(sheetName))
    s = Replace$(s, ChrW$(220), "U")
    s = Replace$(s, ChrW$(252), "u")
    PID_NormalizeSheetNameKey = UCase$(s)
End Function


Private Function PID_IsUbersichtSheet(ByVal ws As Worksheet) As Boolean
    PID_IsUbersichtSheet = PID_IsUbersichtWorksheet(ws)
End Function


Private Function PID_SheetProtectionReady(ByVal sheetName As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    If mSessionProtectedSheets Is Nothing Then Exit Function
    If sheetName = "" Then Exit Function
    
    tmp = mSessionProtectedSheets.item(sheetName)
    
    PID_SheetProtectionReady = True
    Exit Function

NotFound:
    PID_SheetProtectionReady = False
End Function


Private Sub PID_MarkSheetProtectionReady(ByVal sheetName As String)
    On Error Resume Next
    
    If sheetName = "" Then Exit Sub
    If mSessionProtectedSheets Is Nothing Then Set mSessionProtectedSheets = New Collection
    
    mSessionProtectedSheets.Add sheetName, sheetName
End Sub


Private Function PID_IsWorkerMonthSheetSafe(ByVal ws As Worksheet) As Boolean
    Dim monthName As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    
    monthName = Trim$(CStr(ws.Name))
    
    Select Case monthName
        Case "Januar", "Februar", "Marz", "April", "Mai", "Juni", _
             "Juli", "August", "September", "Oktober", "November", "Dezember"
            
            PID_IsWorkerMonthSheetSafe = True
        
        Case Else
            PID_IsWorkerMonthSheetSafe = False
    End Select
    
    Exit Function

SafeExit:
    PID_IsWorkerMonthSheetSafe = False
End Function


Private Sub PID_UnlockMonthPanelFreitextRange(ByVal ws As Worksheet)
    Dim r As Long
    
    ' O18:Q28 — merge-safe (O:P und Q:R je Zeile), siehe PID_MONTH_PANEL_FREITEXT_RANGE.
    For r = 18 To 28
        ws.Range("O" & r & ":P" & r).Locked = False
        ws.Range("Q" & r & ":R" & r).Locked = False
    Next r
End Sub

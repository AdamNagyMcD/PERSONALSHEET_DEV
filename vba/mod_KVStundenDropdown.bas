Attribute VB_Name = "mod_KVStundenDropdown"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Public Const PID_KV_CODE_LIST_NAME As String = "PID_KV_CODE_LIST"
Private Const PID_KV_CODE_HELPER_COL As Long = 961
Private Const PID_KV_TEMPLATE_KV_CODE As String = "BG2"

Public gKVDropdownsDirty As Boolean
Private mKVDropdownDirtyScopeAll As Boolean
Private mDirtyKVCodes As Collection
Private mKVDropdownRefreshedSheets As Collection
Private mStundenValuesCache As Collection
Private mKVCodeDropdownValidSheets As Collection


Public Sub MarkAllKVDropdownsDirty()
    gKVDropdownsDirty = True
    mKVDropdownDirtyScopeAll = True
    Set mDirtyKVCodes = New Collection
    Set mKVDropdownRefreshedSheets = New Collection
    PID_ClearStundenValuesCache
    PID_InvalidateKVCodeDropdownValidCache
End Sub


Public Sub MarkKVDropdownDirtyForKVCode(ByVal kvCode As String)
    kvCode = NormalizeKVCodeForLookup(kvCode)
    If kvCode = "" Then Exit Sub
    
    gKVDropdownsDirty = True
    mKVDropdownDirtyScopeAll = False
    
    If mDirtyKVCodes Is Nothing Then Set mDirtyKVCodes = New Collection
    
    If Not CollectionHasKey_KVDropdown(mDirtyKVCodes, kvCode) Then
        mDirtyKVCodes.Add kvCode, kvCode
    End If
    
    ' Auch beim erneuten Markieren desselben KV-Codes (z. B. Eigene Stunde zuerst
    ' in alter, dann in neuer Periode) muessen Cache und Refresh-Tracking geleert
    ' werden, sonst zeigt das Monatsblatt (z. B. Mai) die neue Stunde nicht.
    Set mKVDropdownRefreshedSheets = New Collection
    PID_ClearStundenValuesCache
End Sub


Public Sub MarkKVDropdownDirtyFromLOHNTABELLERange(ByVal wsKV As Worksheet, ByVal changedRange As Range)
    Dim intersectRange As Range
    Dim cell As Range
    Dim kvCode As String
    Dim kvCodes As Collection
    Dim i As Long
    
    If wsKV Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    
    Set intersectRange = Intersect(changedRange, wsKV.Range("D4:G500"))
    If intersectRange Is Nothing Then Exit Sub
    
    Set kvCodes = New Collection
    
    For Each cell In intersectRange.Cells
        If cell.Row >= 4 Then
            kvCode = NormalizeKVCodeForLookup(CStr(wsKV.Cells(cell.Row, "D").Value))
            If kvCode <> "" Then
                If Not CollectionHasKey_KVDropdown(kvCodes, kvCode) Then
                    kvCodes.Add kvCode, kvCode
                End If
            End If
        End If
    Next cell
    
    If kvCodes.Count = 0 Then
        MarkAllKVDropdownsDirty
        Exit Sub
    End If
    
    For i = 1 To kvCodes.Count
        MarkKVDropdownDirtyForKVCode CStr(kvCodes(i))
    Next i
End Sub


Public Sub MarkKVDropdownsClean()
    gKVDropdownsDirty = False
    mKVDropdownDirtyScopeAll = False
    Set mDirtyKVCodes = New Collection
End Sub


Public Function AreKVDropdownsDirty() As Boolean
    AreKVDropdownsDirty = gKVDropdownsDirty
End Function


Public Sub RefreshKVDropdownsIfDirty()
    If gKVDropdownsDirty Then
        RefreshAllMonthKVStundenDropdowns
        MarkKVDropdownsClean
    End If
End Sub


' Mac-only: nach LOHNTABELLE-Aenderung betroffene Monatsblaetter sofort neu aufbauen
' (scoped dirty refresh — Windows-Pfad unveraendert lazy via SheetActivate).
Public Sub PID_MacRefreshKVDropdownsForKVPeriodChange(Optional ByVal targetPeriod As String = "")
    Dim monthNames As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim monthNum As Long
    Dim workbookYear As Long
    Dim periodForMonth As String
    Dim normalizedTarget As String
    
    If Not PID_IsMacExcel() Then Exit Sub
    If Not gKVDropdownsDirty Then Exit Sub
    
    normalizedTarget = NormalizeKVPeriodForLookup(targetPeriod)
    workbookYear = PID_GetWorkbookYear()
    monthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    PID_BeginPreserveWorkbookView
    PID_BeginHeavyMaintenance
    
    For i = LBound(monthNames) To UBound(monthNames)
        monthNum = i + 1
        
        If normalizedTarget <> "" Then
            periodForMonth = NormalizeKVPeriodForLookup(GetKVPeriodForWorkbookYear(workbookYear, monthNum))
            If periodForMonth <> normalizedTarget Then GoTo NextMonthMacRefresh
        End If
        
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            RefreshKVStundenDropdownForSheet ws, , True
            
            On Error Resume Next
            If mKVDropdownRefreshedSheets Is Nothing Then Set mKVDropdownRefreshedSheets = New Collection
            If Not CollectionHasKey_KVDropdown(mKVDropdownRefreshedSheets, ws.Name) Then
                mKVDropdownRefreshedSheets.Add ws.Name, ws.Name
            End If
            Err.Clear
        End If
        
NextMonthMacRefresh:
    Next i
    
    PID_EndHeavyMaintenance
    PID_EndPreserveWorkbookView
    
    ' Betroffene Monate sind neu aufgebaut — globaler Dirty-Zustand nicht offen lassen.
    MarkKVDropdownsClean
End Sub


Public Sub RefreshKVDropdownsIfDirtyForSheet(ByVal wsMonth As Worksheet)
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If Not gKVDropdownsDirty Then Exit Sub
    
    ' Windows: pro Monat einmal pro Dirty-Zyklus (Performance). Mac: SheetActivate
    ' ist unzuverlaessig — nie ueber "bereits frisch" ueberspringen.
    If Not PID_IsMacExcel() Then
        If Not mKVDropdownRefreshedSheets Is Nothing Then
            If CollectionHasKey_KVDropdown(mKVDropdownRefreshedSheets, wsMonth.Name) Then Exit Sub
        End If
        RefreshKVStundenDropdownForSheet wsMonth
    Else
        RefreshKVStundenDropdownForSheet wsMonth, , True
    End If
    
    On Error Resume Next
    If mKVDropdownRefreshedSheets Is Nothing Then Set mKVDropdownRefreshedSheets = New Collection
    mKVDropdownRefreshedSheets.Add wsMonth.Name, wsMonth.Name
End Sub


Public Sub RefreshAllMonthKVStundenDropdowns()
    Dim monthNames As Variant
    Dim i As Long
    Dim ws As Worksheet
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    PID_BeginPreserveWorkbookView
    PID_BeginHeavyMaintenance
    
    monthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            RefreshKVStundenDropdownForSheet ws, , PID_IsMacExcel()
        End If
    Next i
    
    PID_RemoveLegacyKVDDNamedRanges
    MarkKVDropdownsClean
    PID_EndHeavyMaintenance
    PID_EndPreserveWorkbookView

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    gKVDropdownsDirty = True
    
    PID_EndHeavyMaintenance
    PID_EndPreserveWorkbookView
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei RefreshAllMonthKVStundenDropdowns:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "KV Stunden Dropdowns"
End Sub


Public Sub RefreshKVStundenDropdownForSingleRow(ByVal wsMonth As Worksheet, ByVal rowNumber As Long)
    Dim wsHelper As Worksheet
    Dim monthNumber As Long
    Dim oldScreenUpdating As Boolean
    Dim monthWasProtected As Boolean
    
    On Error GoTo CleanExit
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Sub
    If Not IsNumeric(wsMonth.Range("A1").Value) Then Exit Sub
    
    monthNumber = CLng(wsMonth.Range("A1").Value)
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Set wsHelper = GetOrCreateKVDropdownHelperSheet()
    
    monthWasProtected = wsMonth.ProtectContents
    
    On Error Resume Next
    If monthWasProtected Then wsMonth.Unprotect Password:=PID_WORKBOOK_PASSWORD
    wsHelper.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanExit
    
    RefreshKVStundenDropdownForRow wsMonth, wsHelper, rowNumber, monthNumber
    
    wsHelper.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
    
    If monthWasProtected Then
        PID_ProtectWorkerMonthSheet wsMonth
    End If

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
End Sub


Public Sub RefreshKVStundenDropdownForSheet(ByVal wsMonth As Worksheet, _
                                           Optional ByVal changedRange As Range, _
                                           Optional ByVal forceFullRebuild As Boolean = False)
    Dim wsHelper As Worksheet
    Dim monthNumber As Long
    Dim r As Long
    Dim rowsToCheck As Range
    Dim c As Range
    Dim checkedRows As Collection
    Dim rowKey As String
    
    Dim oldDisplayAlerts As Boolean
    Dim oldScreenUpdating As Boolean
    Dim monthWasProtected As Boolean
    
    On Error GoTo CleanFail
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If Not IsNumeric(wsMonth.Range("A1").Value) Then Exit Sub
    
    monthNumber = CLng(wsMonth.Range("A1").Value)
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    oldDisplayAlerts = Application.DisplayAlerts
    oldScreenUpdating = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set wsHelper = GetOrCreateKVDropdownHelperSheet()
    
    monthWasProtected = wsMonth.ProtectContents
    
    On Error Resume Next
    wsMonth.Unprotect Password:=PID_WORKBOOK_PASSWORD
    wsHelper.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    wsMonth.Range("F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW).Locked = False
    
    If changedRange Is Nothing Then
        If forceFullRebuild Or PID_ShouldRefreshAllKVDropdownKeys() Then
            ClearHelperColumnsForSheet wsHelper, wsMonth.Name
        End If
        RefreshKVStundenDropdownForSheetBulk wsMonth, wsHelper, monthNumber, forceFullRebuild
    Else
        
        Set rowsToCheck = Intersect(changedRange, wsMonth.Range("E3:E82"))
        
        If Not rowsToCheck Is Nothing Then
            Set checkedRows = New Collection
            
            For Each c In rowsToCheck.Cells
                rowKey = CStr(c.Row)
                
                If Not CollectionHasKey_KVDropdown(checkedRows, rowKey) Then
                    checkedRows.Add rowKey, rowKey
                    RefreshKVStundenDropdownForRow wsMonth, wsHelper, c.Row, monthNumber
                End If
            Next c
        End If
        
    End If
    
    wsHelper.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
    wsHelper.Visible = xlSheetVeryHidden
    
    If monthWasProtected Then
        PID_ProtectWorkerMonthSheet wsMonth
    End If

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not wsHelper Is Nothing Then
        wsHelper.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
        wsHelper.Visible = xlSheetVeryHidden
    End If
    
    If Not wsMonth Is Nothing Then
        If monthWasProtected Then
            PID_ProtectWorkerMonthSheet wsMonth
        End If
    End If
    
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    
    MsgBox "Fehler bei RefreshKVStundenDropdownForSheet:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "KV Stunden Dropdown"
End Sub


Public Function PID_GetTemplateKVCodeForStundenDropdown() As String
    PID_GetTemplateKVCodeForStundenDropdown = PID_KV_TEMPLATE_KV_CODE
End Function


Private Function PID_GetDropdownKeyForRow(ByVal wsMonth As Worksheet, ByVal rowNumber As Long) As String
    Dim kvCode As String
    
    kvCode = NormalizeKVCodeForLookup(CStr(wsMonth.Cells(rowNumber, "E").Value))
    
    If kvCode = "" Then
        PID_GetDropdownKeyForRow = "__TEMPLATE__"
    Else
        PID_GetDropdownKeyForRow = kvCode
    End If
End Function


Private Function PID_BuildAllDropdownKeysForSheet(ByVal wsMonth As Worksheet) As Collection
    Dim allKeys As Collection
    Dim r As Long
    Dim key As String
    
    Set allKeys = New Collection
    
    For r = PID_FIRST_ROW To PID_LAST_ROW
        key = PID_GetDropdownKeyForRow(wsMonth, r)
        
        If Not CollectionHasKey_KVDropdown(allKeys, key) Then
            allKeys.Add key, key
        End If
    Next r
    
    Set PID_BuildAllDropdownKeysForSheet = allKeys
End Function


Private Function GetDropdownNameForKVCode(ByVal sheetName As String, ByVal kvCodeKey As String) As String
    Dim safeName As String
    Dim safeKey As String
    
    safeName = UCase$(Trim$(CStr(sheetName)))
    safeName = Replace(safeName, " ", "_")
    safeName = Replace(safeName, "-", "_")
    safeName = Replace(safeName, ".", "_")
    safeName = Replace(safeName, "/", "_")
    safeName = Replace(safeName, "\", "_")
    safeName = Replace(safeName, ChrW(196), "AE")
    safeName = Replace(safeName, ChrW(214), "OE")
    safeName = Replace(safeName, ChrW(220), "UE")
    safeName = Replace(safeName, ChrW(228), "AE")
    safeName = Replace(safeName, ChrW(246), "OE")
    safeName = Replace(safeName, ChrW(252), "UE")
    safeName = Replace(safeName, ChrW(223), "SS")
    
    safeKey = UCase$(Trim$(CStr(kvCodeKey)))
    safeKey = Replace(safeKey, " ", "_")
    safeKey = Replace(safeKey, "/", "_")
    safeKey = Replace(safeKey, "\", "_")
    
    GetDropdownNameForKVCode = "KV_DG_" & safeName & "_" & safeKey
End Function


Private Function GetHelperColumnForKVCodeSlot(ByVal sheetName As String, ByVal slotIndex As Long) As Long
    Dim monthIndex As Long
    Dim firstCol As Long
    
    monthIndex = GetMonthIndexForHelper(sheetName)
    If monthIndex < 1 Then monthIndex = 1
    
    If slotIndex < 1 Then slotIndex = 1
    If slotIndex > 25 Then slotIndex = 25
    
    firstCol = ((monthIndex - 1) * 80) + 1
    GetHelperColumnForKVCodeSlot = firstCol + slotIndex - 1
End Function


Private Sub PID_ForceDeleteAllFStundenValidations(ByVal wsMonth As Worksheet)
    Dim r As Long
    
    On Error Resume Next
    
    wsMonth.Range("F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW).Validation.Delete
    DoEvents
    
    For r = PID_FIRST_ROW To PID_LAST_ROW
        wsMonth.Cells(r, "F").Validation.Delete
    Next r
    
    If PID_IsMacExcel() Then
        DoEvents
        wsMonth.Range("F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW).Validation.Delete
        For r = PID_FIRST_ROW To PID_LAST_ROW
            wsMonth.Cells(r, "F").Validation.Delete
        Next r
    End If
    
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub PID_ApplyFStundenListValidation(ByVal wsMonth As Worksheet, _
                                            ByVal rowNumber As Long, _
                                            ByVal listName As String, _
                                            ByVal listRange As Range)
    If wsMonth Is Nothing Then Exit Sub
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Sub
    
    On Error Resume Next
    wsMonth.Cells(rowNumber, "F").Validation.Delete
    Err.Clear
    DoEvents
    wsMonth.Cells(rowNumber, "F").Validation.Delete
    Err.Clear
    
    If listRange Is Nothing Then Exit Sub
    
    If PID_IsMacExcel() Then GoTo ApplyDirectAddress
    
    On Error GoTo ApplyFailed
    With wsMonth.Cells(rowNumber, "F").Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    Exit Sub

ApplyFailed:
    On Error Resume Next
    wsMonth.Cells(rowNumber, "F").Validation.Delete
    Err.Clear

ApplyDirectAddress:
    On Error GoTo SafeExit
    With wsMonth.Cells(rowNumber, "F").Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=" & listRange.Address(External:=True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

SafeExit:
End Sub


Private Sub RefreshKVStundenDropdownForSheetBulk(ByVal wsMonth As Worksheet, _
                                                 ByVal wsHelper As Worksheet, _
                                                 ByVal monthNumber As Long, _
                                                 Optional ByVal forceFullRebuild As Boolean = False)
    Dim allKeys As Collection
    Dim refreshKeys As Collection
    Dim key As String
    Dim lookupCode As String
    Dim r As Long
    Dim i As Long
    Dim slotIndex As Long
    Dim values As Collection
    Dim helperCol As Long
    Dim helperLastRow As Long
    Dim listRange As Range
    Dim listName As String
    Dim refreshAllKeys As Boolean
    
    refreshAllKeys = forceFullRebuild Or PID_ShouldRefreshAllKVDropdownKeys()
    
    Set allKeys = PID_BuildAllDropdownKeysForSheet(wsMonth)
    Set refreshKeys = New Collection
    
    For r = PID_FIRST_ROW To PID_LAST_ROW
        key = PID_GetDropdownKeyForRow(wsMonth, r)
        
        If refreshAllKeys Or PID_IsKVCodeDirtyRefreshTarget(key) Then
            If Not CollectionHasKey_KVDropdown(refreshKeys, key) Then
                refreshKeys.Add key, key
            End If
        End If
    Next r
    
    If refreshKeys.Count = 0 Then Exit Sub
    
    If refreshAllKeys Then
        PID_ForceDeleteAllFStundenValidations wsMonth
    Else
        For r = PID_FIRST_ROW To PID_LAST_ROW
            key = PID_GetDropdownKeyForRow(wsMonth, r)
            If PID_IsKVCodeDirtyRefreshTarget(key) Then
                On Error Resume Next
                wsMonth.Cells(r, "F").Validation.Delete
                Err.Clear
            End If
        Next r
    End If
    
    For i = 1 To refreshKeys.Count
        key = CStr(refreshKeys(i))
        slotIndex = PID_GetKVCodeSlotIndexInSheet(allKeys, key)
        
        If key = "__TEMPLATE__" Then
            lookupCode = PID_GetTemplateKVCodeForStundenDropdown()
        Else
            lookupCode = key
        End If
        
        Set values = GetKVMonatsstundenValues(monthNumber, lookupCode)
        
        helperCol = GetHelperColumnForKVCodeSlot(wsMonth.Name, slotIndex)
        wsHelper.Range(wsHelper.Cells(1, helperCol), wsHelper.Cells(30, helperCol)).ClearContents
        
        helperLastRow = 0
        listName = ""
        Set listRange = Nothing
        
        If Not values Is Nothing Then
            If values.Count > 0 Then
                helperLastRow = WriteDropdownValuesToHelper(wsHelper, helperCol, values)
                
                If helperLastRow > 0 Then
                    Set listRange = wsHelper.Range(wsHelper.Cells(1, helperCol), wsHelper.Cells(helperLastRow, helperCol))
                    listName = GetDropdownNameForKVCode(wsMonth.Name, key)
                    PID_EnsureWorkbookNameRefersTo listName, listRange, (gKVDropdownsDirty Or forceFullRebuild)
                End If
            End If
        End If
        
        For r = PID_FIRST_ROW To PID_LAST_ROW
            If PID_GetDropdownKeyForRow(wsMonth, r) <> key Then GoTo NextRow
            
            If helperLastRow <= 0 Then
                On Error Resume Next
                wsMonth.Cells(r, "F").Validation.Delete
                If key <> "__TEMPLATE__" Then wsMonth.Cells(r, "F").ClearContents
                Err.Clear
            Else
                If forceFullRebuild Then
                    PID_ApplyFStundenListValidation wsMonth, r, listName, listRange
                ElseIf refreshAllKeys Then
                    If gKVDropdownsDirty Or Not PID_RowHasValidFStundenDropdown(wsMonth, r) Then
                        PID_ApplyFStundenListValidation wsMonth, r, listName, listRange
                    End If
                Else
                    PID_ApplyFStundenListValidation wsMonth, r, listName, listRange
                End If
            End If
            
NextRow:
        Next r
    Next i
End Sub


Public Sub RefreshKVStundenDropdownForRow(ByVal wsMonth As Worksheet, _
                                          ByVal wsHelper As Worksheet, _
                                          ByVal rowNumber As Long, _
                                          ByVal monthNumber As Long)
    Dim dropdownKey As String
    Dim lookupCode As String
    Dim useTemplateKV As Boolean
    Dim allKeys As Collection
    Dim slotIndex As Long
    Dim values As Collection
    Dim helperCol As Long
    Dim helperLastRow As Long
    Dim listRange As Range
    Dim listName As String
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If wsHelper Is Nothing Then Exit Sub
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Sub
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    dropdownKey = PID_GetDropdownKeyForRow(wsMonth, rowNumber)
    useTemplateKV = (dropdownKey = "__TEMPLATE__")
    
    If useTemplateKV Then
        lookupCode = PID_GetTemplateKVCodeForStundenDropdown()
    Else
        lookupCode = dropdownKey
    End If
    
    wsMonth.Cells(rowNumber, "F").Locked = False
    
    Set values = GetKVMonatsstundenValues(monthNumber, lookupCode)
    
    If values.Count = 0 Then
        On Error Resume Next
        wsMonth.Cells(rowNumber, "F").Validation.Delete
        If Not useTemplateKV Then wsMonth.Cells(rowNumber, "F").ClearContents
        Err.Clear
        Exit Sub
    End If
    
    Set allKeys = PID_BuildAllDropdownKeysForSheet(wsMonth)
    slotIndex = PID_GetKVCodeSlotIndexInSheet(allKeys, dropdownKey)
    helperCol = GetHelperColumnForKVCodeSlot(wsMonth.Name, slotIndex)
    
    wsHelper.Range(wsHelper.Cells(1, helperCol), wsHelper.Cells(30, helperCol)).ClearContents
    
    helperLastRow = WriteDropdownValuesToHelper(wsHelper, helperCol, values)
    
    If helperLastRow <= 0 Then Exit Sub
    
    Set listRange = wsHelper.Range(wsHelper.Cells(1, helperCol), wsHelper.Cells(helperLastRow, helperCol))
    listName = GetDropdownNameForKVCode(wsMonth.Name, dropdownKey)
    
    PID_EnsureWorkbookNameRefersTo listName, listRange, True
    
    PID_ApplyFStundenListValidation wsMonth, rowNumber, listName, listRange

SafeExit:
End Sub


Public Function GetOrCreateKVDropdownHelperSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("KV_DROPDOWN_HELPER")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "KV_DROPDOWN_HELPER"
    End If
    
    Set GetOrCreateKVDropdownHelperSheet = ws
End Function


Public Function GetKVMonatsstundenValues(ByVal monthNumber As Long, ByVal kvCode As String) As Collection
    Dim wsKV As Worksheet
    Dim currentYear As Long
    Dim targetPeriod As String
    Dim previousPeriod As String
    Dim values As Collection
    Dim cacheKey As String
    Dim cachedValues As Collection
    
    Set values = New Collection
    
    On Error GoTo SafeExit
    
    kvCode = NormalizeKVCodeForLookup(kvCode)
    If kvCode = "" Then GoTo SafeExit
    
    cacheKey = CStr(monthNumber) & "|" & kvCode
    
    ' Mac: kein Stunden-Cache — LOHNTABELLE-Aenderungen (Eigene Stunden) sonst oft veraltet.
    If Not PID_IsMacExcel() Then
        If Not mStundenValuesCache Is Nothing Then
            On Error Resume Next
            Set cachedValues = mStundenValuesCache(cacheKey)
            If Err.Number = 0 Then
                Set GetKVMonatsstundenValues = PID_CloneStundenValuesCollection(cachedValues)
                Exit Function
            End If
            Err.Clear
            On Error GoTo SafeExit
        End If
    End If
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    
    currentYear = PID_GetWorkbookYear()
    If currentYear <= 0 Then GoTo SafeExit
    
    targetPeriod = NormalizeKVPeriodForLookup(GetKVPeriodForWorkbookYear(currentYear, monthNumber))
    previousPeriod = NormalizeKVPeriodForLookup(GetPreviousKVPeriodForWorkbookYear(currentYear, monthNumber))
    
    AddMonatsstundenValuesFromPeriod wsKV, targetPeriod, kvCode, values
    
    If values.Count = 0 Then
        AddMonatsstundenValuesFromPeriod wsKV, previousPeriod, kvCode, values
    End If
    
    If Not PID_IsMacExcel() Then
        PID_StoreStundenValuesInCache cacheKey, values
    End If
    Set GetKVMonatsstundenValues = PID_CloneStundenValuesCollection(values)
    Exit Function

SafeExit:
    Set GetKVMonatsstundenValues = values
End Function


Public Sub PID_RefreshFStundenDropdownForERows(ByVal wsMonth As Worksheet, ByVal changedRange As Range)
    Dim rowsToCheck As Range
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    
    Set rowsToCheck = Intersect(changedRange, wsMonth.Range("E3:E82"))
    If rowsToCheck Is Nothing Then Exit Sub
    
    DoEvents
    PID_ForceRefreshFStundenDropdownForSheet wsMonth

SafeExit:
End Sub


Public Sub PID_ForceRefreshFStundenDropdownForSheet(ByVal wsMonth As Worksheet)
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    
    PID_BeginHeavyMaintenance
    PID_ClearStundenValuesCache
    RefreshKVStundenDropdownForSheet wsMonth, , True
    PID_EndHeavyMaintenance wsMonth
    
SafeExit:
    On Error Resume Next
    PID_EndHeavyMaintenance
End Sub


Public Sub PID_InvalidateFStundenDropdownForRows(ByVal wsMonth As Worksheet, ByVal changedRange As Range)
    Dim rowsToCheck As Range
    Dim c As Range
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    
    Set rowsToCheck = Intersect(changedRange, wsMonth.Range("E3:E82"))
    If rowsToCheck Is Nothing Then Exit Sub
    
    For Each c In rowsToCheck.Cells
        On Error Resume Next
        wsMonth.Cells(c.Row, "F").Validation.Delete
        Err.Clear
    Next c

SafeExit:
End Sub


Public Sub PID_ClearStundenValuesCache()
    Set mStundenValuesCache = Nothing
End Sub


Private Function PID_ShouldRefreshAllKVDropdownKeys() As Boolean
    PID_ShouldRefreshAllKVDropdownKeys = (Not gKVDropdownsDirty) Or mKVDropdownDirtyScopeAll
End Function


Private Function PID_IsKVCodeDirtyRefreshTarget(ByVal dropdownKey As String) As Boolean
    Dim lookupCode As String
    
    If mKVDropdownDirtyScopeAll Then
        PID_IsKVCodeDirtyRefreshTarget = True
        Exit Function
    End If
    
    If mDirtyKVCodes Is Nothing Then Exit Function
    If mDirtyKVCodes.Count = 0 Then Exit Function
    
    If dropdownKey = "__TEMPLATE__" Then
        lookupCode = NormalizeKVCodeForLookup(PID_GetTemplateKVCodeForStundenDropdown())
    Else
        lookupCode = dropdownKey
    End If
    
    PID_IsKVCodeDirtyRefreshTarget = CollectionHasKey_KVDropdown(mDirtyKVCodes, lookupCode)
End Function


Private Function PID_GetKVCodeSlotIndexInSheet(ByVal allKeys As Collection, ByVal key As String) As Long
    Dim i As Long
    
    For i = 1 To allKeys.Count
        If CStr(allKeys(i)) = key Then
            PID_GetKVCodeSlotIndexInSheet = i
            Exit Function
        End If
    Next i
    
    PID_GetKVCodeSlotIndexInSheet = 1
End Function


Private Sub PID_StoreStundenValuesInCache(ByVal cacheKey As String, ByVal values As Collection)
    On Error Resume Next
    
    If mStundenValuesCache Is Nothing Then Set mStundenValuesCache = New Collection
    mStundenValuesCache.Remove cacheKey
    mStundenValuesCache.Add PID_CloneStundenValuesCollection(values), cacheKey
End Sub


Private Function PID_CloneStundenValuesCollection(ByVal sourceValues As Collection) As Collection
    Dim cloneValues As Collection
    Dim i As Long
    
    Set cloneValues = New Collection
    
    If sourceValues Is Nothing Then
        Set PID_CloneStundenValuesCollection = cloneValues
        Exit Function
    End If
    
    For i = 1 To sourceValues.Count
        cloneValues.Add sourceValues(i)
    Next i
    
    Set PID_CloneStundenValuesCollection = cloneValues
End Function


Public Sub AddMonatsstundenValuesFromPeriod(ByVal wsKV As Worksheet, _
                                            ByVal periodName As String, _
                                            ByVal kvCode As String, _
                                            ByRef values As Collection)
    Dim lastRow As Long
    Dim r As Long
    Dim rowPeriod As String
    Dim rowKVCode As String
    Dim rowMonatsstunden As Variant
    Dim keyText As String
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If Trim$(periodName) = "" Then Exit Sub
    If Trim$(kvCode) = "" Then Exit Sub
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    periodName = NormalizeKVPeriodForLookup(periodName)
    If periodName = "" Then Exit Sub
    
    For r = 4 To lastRow
        If wsKV.Range("A" & r).MergeCells Then GoTo NextRow
        
        rowPeriod = GetRowKVPeriodForDropdown(wsKV, r)
        If rowPeriod <> periodName Then GoTo NextRow
        
        rowKVCode = NormalizeKVCodeForLookup(CStr(wsKV.Cells(r, "D").Value))
        If rowKVCode <> kvCode Then GoTo NextRow
        
        rowMonatsstunden = wsKV.Cells(r, "G").Value
        
        If IsNumeric(rowMonatsstunden) Then
            If CDbl(rowMonatsstunden) > 0# Then
                keyText = CStr(CDbl(rowMonatsstunden))
                
                If Not CollectionHasKey_KVDropdown(values, keyText) Then
                    values.Add CDbl(rowMonatsstunden), keyText
                End If
            End If
        End If
        
NextRow:
    Next r

SafeExit:
End Sub


Public Function WriteDropdownValuesToHelper(ByVal wsHelper As Worksheet, _
                                            ByVal helperCol As Long, _
                                            ByVal values As Collection) As Long
    Dim i As Long
    
    On Error GoTo SafeExit
    
    If wsHelper Is Nothing Then Exit Function
    If values Is Nothing Then Exit Function
    If values.count = 0 Then Exit Function
    If helperCol < 1 Then Exit Function
    
    For i = 1 To values.count
        wsHelper.Cells(i, helperCol).Value = CDbl(values.item(i))
        wsHelper.Cells(i, helperCol).NumberFormat = "0.00"
    Next i
    
    WriteDropdownValuesToHelper = values.count
    Exit Function

SafeExit:
    WriteDropdownValuesToHelper = 0
End Function


Public Sub CreateOrReplaceWorkbookName(ByVal nameText As String, ByVal targetRange As Range)
    PID_EnsureWorkbookNameRefersTo nameText, targetRange, True
End Sub


Public Sub PID_EnsureWorkbookNameRefersTo(ByVal nameText As String, _
                                          ByVal targetRange As Range, _
                                          Optional ByVal replaceIfDifferent As Boolean = False)
    Dim refersText As String
    Dim existingRefersTo As String
    
    On Error GoTo SafeExit
    
    If targetRange Is Nothing Then Exit Sub
    If Trim$(nameText) = "" Then Exit Sub
    
    refersText = "='" & targetRange.Worksheet.Name & "'!" & targetRange.Address(True, True)
    
    On Error Resume Next
    existingRefersTo = CStr(ThisWorkbook.Names(nameText).RefersTo)
    
    If Err.Number = 0 Then
        If StrComp(existingRefersTo, "=" & refersText, vbTextCompare) = 0 Then Exit Sub
        If Not replaceIfDifferent Then Exit Sub
        ThisWorkbook.Names(nameText).Delete
        Err.Clear
    End If
    
    On Error GoTo SafeExit
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:="=" & refersText

SafeExit:
End Sub


Public Sub PID_RemoveLegacyKVDDNamedRanges()
    Dim i As Long
    Dim nameText As String
    
    On Error Resume Next
    
    For i = ThisWorkbook.Names.Count To 1 Step -1
        nameText = CStr(ThisWorkbook.Names(i).Name)
        
        If InStr(1, nameText, "KV_DD_", vbTextCompare) > 0 Then
            ThisWorkbook.Names(i).Delete
        End If
    Next i
    
    On Error GoTo 0
End Sub


Public Function PID_CountKVDDNamedRanges() As Long
    Dim i As Long
    Dim nameText As String
    Dim countNames As Long
    
    countNames = 0
    
    On Error Resume Next
    
    For i = 1 To ThisWorkbook.Names.Count
        nameText = CStr(ThisWorkbook.Names(i).Name)
        
        If InStr(1, nameText, "KV_DD_", vbTextCompare) > 0 Then
            countNames = countNames + 1
        End If
    Next i
    
    On Error GoTo 0
    PID_CountKVDDNamedRanges = countNames
End Function


' Legacy row-based name — only for detecting old F-validations until migrated (FP-006).
Public Function GetDropdownNameForMonthRow(ByVal sheetName As String, ByVal rowNumber As Long) As String
    Dim safeName As String
    
    safeName = UCase$(Trim$(CStr(sheetName)))
    
    safeName = Replace(safeName, " ", "_")
    safeName = Replace(safeName, "-", "_")
    safeName = Replace(safeName, ".", "_")
    safeName = Replace(safeName, "/", "_")
    safeName = Replace(safeName, "\", "_")
    safeName = Replace(safeName, ChrW(196), "AE")
    safeName = Replace(safeName, ChrW(214), "OE")
    safeName = Replace(safeName, ChrW(220), "UE")
    safeName = Replace(safeName, ChrW(228), "AE")
    safeName = Replace(safeName, ChrW(246), "OE")
    safeName = Replace(safeName, ChrW(252), "UE")
    safeName = Replace(safeName, ChrW(223), "SS")
    
    GetDropdownNameForMonthRow = "KV_DD_" & safeName & "_" & CStr(rowNumber)
End Function


Public Function GetHelperColumnForMonthRow(ByVal sheetName As String, ByVal rowNumber As Long) As Long
    Dim monthIndex As Long
    
    monthIndex = GetMonthIndexForHelper(sheetName)
    
    If monthIndex < 1 Then monthIndex = 1
    If rowNumber < 3 Then rowNumber = 3
    If rowNumber > 82 Then rowNumber = 82
    
    GetHelperColumnForMonthRow = ((monthIndex - 1) * 80) + (rowNumber - 2)
End Function


Public Function GetMonthIndexForHelper(ByVal sheetName As String) As Long
    Select Case Trim$(CStr(sheetName))
        Case "Januar"
            GetMonthIndexForHelper = 1
        Case "Februar"
            GetMonthIndexForHelper = 2
        Case "Marz"
            GetMonthIndexForHelper = 3
        Case "April"
            GetMonthIndexForHelper = 4
        Case "Mai"
            GetMonthIndexForHelper = 5
        Case "Juni"
            GetMonthIndexForHelper = 6
        Case "Juli"
            GetMonthIndexForHelper = 7
        Case "August"
            GetMonthIndexForHelper = 8
        Case "September"
            GetMonthIndexForHelper = 9
        Case "Oktober"
            GetMonthIndexForHelper = 10
        Case "November"
            GetMonthIndexForHelper = 11
        Case "Dezember"
            GetMonthIndexForHelper = 12
        Case Else
            GetMonthIndexForHelper = 0
    End Select
End Function


Public Sub ClearHelperColumnsForSheet(ByVal wsHelper As Worksheet, ByVal sheetName As String)
    Dim monthIndex As Long
    Dim firstCol As Long
    Dim lastCol As Long
    
    On Error GoTo SafeExit
    
    If wsHelper Is Nothing Then Exit Sub
    
    monthIndex = GetMonthIndexForHelper(sheetName)
    If monthIndex < 1 Then Exit Sub
    
    firstCol = ((monthIndex - 1) * 80) + 1
    lastCol = firstCol + 79
    
    wsHelper.Range(wsHelper.Columns(firstCol), wsHelper.Columns(lastCol)).Clear

SafeExit:
End Sub


Private Function GetRowKVPeriodForDropdown(ByVal wsKV As Worksheet, ByVal rowNumber As Long) As String
    Dim r As Long
    Dim cellText As String
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    If rowNumber < 1 Then Exit Function
    
    cellText = Trim$(CStr(wsKV.Cells(rowNumber, "A").Value))
    
    If cellText <> "" Then
        GetRowKVPeriodForDropdown = NormalizeKVPeriodForLookup(cellText)
        Exit Function
    End If
    
    For r = rowNumber - 1 To 4 Step -1
        cellText = Trim$(CStr(wsKV.Cells(r, "A").Value))
        If cellText <> "" Then
            GetRowKVPeriodForDropdown = NormalizeKVPeriodForLookup(cellText)
            Exit Function
        End If
    Next r
    
SafeExit:
End Function


Public Function CollectionHasKey_KVDropdown(ByVal col As Collection, ByVal key As String) As Boolean
    CollectionHasKey_KVDropdown = PID_CollectionHasKey(col, key)
End Function


Public Function PID_GetStandardKVCodeValidationList() As String
    Dim sep As String
    Dim codes As Variant
    Dim i As Long
    Dim resultText As String
    
    sep = CStr(Application.International(xlListSeparator))
    If sep = "" Then sep = ","
    
    codes = PID_GetStandardKVCodeArray()
    resultText = CStr(codes(LBound(codes)))
    
    For i = LBound(codes) + 1 To UBound(codes)
        resultText = resultText & sep & CStr(codes(i))
    Next i
    
    PID_GetStandardKVCodeValidationList = resultText
End Function


Public Function PID_GetKVCodeValidationFormula() As String
    Dim wsHelper As Worksheet
    Dim listRange As Range
    Dim codes As Variant
    Dim i As Long
    Dim codeCount As Long
    
    On Error GoTo UseNamedRange
    
    Set wsHelper = GetOrCreateKVDropdownHelperSheet()
    
    On Error Resume Next
    wsHelper.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo UseNamedRange
    
    codes = PID_GetStandardKVCodeArray()
    codeCount = UBound(codes) - LBound(codes) + 1
    
    For i = LBound(codes) To UBound(codes)
        wsHelper.Cells(i - LBound(codes) + 1, PID_KV_CODE_HELPER_COL).Value = CStr(codes(i))
    Next i
    
    Set listRange = wsHelper.Range(wsHelper.Cells(1, PID_KV_CODE_HELPER_COL), _
                                   wsHelper.Cells(codeCount, PID_KV_CODE_HELPER_COL))
    
    PID_GetKVCodeValidationFormula = "='" & listRange.Worksheet.Name & "'!" & listRange.Address(True, True)
    Exit Function

UseNamedRange:
    PID_EnsureKVCodeListNamedRange
    PID_GetKVCodeValidationFormula = "=" & PID_KV_CODE_LIST_NAME
End Function


Private Function PID_KVCodeListNamedRangeExists() As Boolean
    On Error Resume Next
    PID_KVCodeListNamedRangeExists = (Len(ThisWorkbook.Names(PID_KV_CODE_LIST_NAME).Name) > 0)
    Err.Clear
End Function


Public Sub PID_EnsureKVCodeListNamedRange()
    Dim wsHelper As Worksheet
    Dim codes As Variant
    Dim i As Long
    Dim listRange As Range
    
    On Error GoTo SafeExit
    
    If PID_KVCodeListNamedRangeExists() Then Exit Sub
    
    Set wsHelper = GetOrCreateKVDropdownHelperSheet()
    
    On Error Resume Next
    wsHelper.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    codes = PID_GetStandardKVCodeArray()
    
    For i = LBound(codes) To UBound(codes)
        wsHelper.Cells(i - LBound(codes) + 1, PID_KV_CODE_HELPER_COL).Value = CStr(codes(i))
    Next i
    
    Set listRange = wsHelper.Range(wsHelper.Cells(1, PID_KV_CODE_HELPER_COL), _
                                   wsHelper.Cells(UBound(codes) - LBound(codes) + 1, PID_KV_CODE_HELPER_COL))
    
    CreateOrReplaceWorkbookName PID_KV_CODE_LIST_NAME, listRange
    PID_InvalidateKVCodeDropdownValidCache

SafeExit:
End Sub


Private Function PID_GetStandardKVCodeArray() As Variant
    PID_GetStandardKVCodeArray = Array( _
        "BG1", "BG1_5", "BG1_10", "BG1_15", _
        "BG2", "BG2_5", "BG2_10", "BG2_15", _
        "BG3", "BG3_5", "BG3_10", "BG3_15")
End Function


Public Sub PID_ApplyKVCodeDropdownValidation(ByVal ws As Worksheet)
    Dim targetRange As Range
    Dim validationFormula As String
    Dim applied As Boolean
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    
    Set targetRange = ws.Range("E" & PID_FIRST_ROW & ":E" & PID_LAST_ROW)
    targetRange.Locked = False
    
    On Error Resume Next
    targetRange.Validation.Delete
    Err.Clear
    On Error GoTo SafeExit
    
    validationFormula = PID_GetKVCodeValidationFormula()
    If Left$(validationFormula, 1) <> "=" Then
        validationFormula = "=" & validationFormula
    End If
    
    applied = PID_TryApplyKVCodeListValidation(targetRange, validationFormula)
    
    If Not applied Then
        applied = PID_TryApplyKVCodeListValidation(targetRange, PID_GetStandardKVCodeValidationList())
    End If

SafeExit:
End Sub


Private Function PID_TryApplyKVCodeListValidation(ByVal targetRange As Range, _
                                                  ByVal validationFormula As String) As Boolean
    On Error GoTo Failed
    
    If targetRange Is Nothing Then Exit Function
    If Trim$(validationFormula) = "" Then Exit Function
    
    With targetRange.Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=validationFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    PID_TryApplyKVCodeListValidation = True
    Exit Function

Failed:
    On Error Resume Next
    targetRange.Validation.Delete
    Err.Clear
End Function


Public Sub PID_InvalidateKVCodeDropdownValidCache()
    Set mKVCodeDropdownValidSheets = New Collection
End Sub


Private Sub PID_MarkKVCodeDropdownValidForSheet(ByVal sheetName As String)
    If sheetName = "" Then Exit Sub
    If mKVCodeDropdownValidSheets Is Nothing Then Set mKVCodeDropdownValidSheets = New Collection
    
    On Error Resume Next
    mKVCodeDropdownValidSheets.Remove sheetName
    Err.Clear
    mKVCodeDropdownValidSheets.Add sheetName, sheetName
    On Error GoTo 0
End Sub


Public Function PID_MonthSheetHasValidKVCodeDropdown(ByVal wsMonth As Worksheet) As Boolean
    Dim validationFormula As String
    Dim validationType As Long
    Dim validationCell As Range
    
    If wsMonth Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Function
    
    If Not mKVCodeDropdownValidSheets Is Nothing Then
        If CollectionHasKey_KVDropdown(mKVCodeDropdownValidSheets, wsMonth.Name) Then
            PID_MonthSheetHasValidKVCodeDropdown = True
            Exit Function
        End If
    End If
    
    Set validationCell = wsMonth.Range("E" & PID_FIRST_ROW)
    
    ' Broken #REF! validation throws on read; Resume Next avoids debugger breaks.
    On Error Resume Next
    validationType = validationCell.Validation.Type
    If Err.Number <> 0 Then Exit Function
    
    If validationType <> xlValidateList Then Exit Function
    
    validationFormula = UCase$(CStr(validationCell.Validation.Formula1))
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    If InStr(1, validationFormula, "#REF", vbTextCompare) > 0 Then Exit Function
    If InStr(1, validationFormula, PID_KV_CODE_LIST_NAME, vbTextCompare) = 0 Then
        If InStr(1, validationFormula, "KV_DROPDOWN_HELPER", vbTextCompare) = 0 Then Exit Function
    End If
    
    PID_MarkKVCodeDropdownValidForSheet wsMonth.Name
    PID_MonthSheetHasValidKVCodeDropdown = True
End Function


Public Sub PID_RestoreKVCodeDropdownValidation()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim updatedCount As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    monthNames = PID_MonthNames()
    updatedCount = 0
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If PID_RestoreKVCodeDropdownOnSheet(ws) Then
                updatedCount = updatedCount + 1
            End If
        End If
    Next i
    
    MsgBox "KV-Code Dropdown (Spalte E) wurde wiederhergestellt." & vbCrLf & vbCrLf & _
           PID_UTxtMonatsblaetter() & " aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
           "Bereich: E" & PID_FIRST_ROW & ":E" & PID_LAST_ROW, _
           vbInformation, "Spalte E"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_RestoreKVCodeDropdownValidation:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Spalte E"
End Sub


Public Sub PID_EnsureKVCodeDropdownValidation()
    PID_RestoreKVCodeDropdownValidationSilent
End Sub


Public Sub PID_EnsureKVCodeDropdownOnSheet(ByVal wsMonth As Worksheet)
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If PID_MonthSheetHasValidKVCodeDropdown(wsMonth) Then Exit Sub
    
    PID_RestoreKVCodeDropdownOnSheet wsMonth

SafeExit:
End Sub


Public Sub PID_RestoreKVCodeDropdownValidationSilent()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo SafeExit
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo SafeExit
        
        If Not ws Is Nothing Then
            If Not PID_MonthSheetHasValidKVCodeDropdown(ws) Then
                PID_RestoreKVCodeDropdownOnSheet ws
            End If
        End If
    Next i

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Public Function PID_RowHasValidFStundenDropdown(ByVal wsMonth As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim validationFormula As String
    Dim validationType As Long
    Dim dropdownKey As String
    Dim expectedListName As String
    Dim normalizedFormula As String
    Dim allKeys As Collection
    Dim slotIndex As Long
    Dim expectedHelperCol As Long
    
    If wsMonth Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Function
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Function
    
    On Error Resume Next
    validationType = wsMonth.Cells(rowNumber, "F").Validation.Type
    If Err.Number <> 0 Then Exit Function
    
    If validationType <> xlValidateList Then Exit Function
    
    validationFormula = CStr(wsMonth.Cells(rowNumber, "F").Validation.Formula1)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    
    normalizedFormula = PID_NormalizeValidationListFormula(validationFormula)
    If InStr(1, normalizedFormula, "#REF", vbTextCompare) > 0 Then Exit Function
    
    dropdownKey = PID_GetDropdownKeyForRow(wsMonth, rowNumber)
    expectedListName = UCase$(GetDropdownNameForKVCode(wsMonth.Name, dropdownKey))
    
    ' Exakter Named-Range-Vergleich (inkl. Workbook-Praefix auf Mac).
    If PID_ValidationFormulaRefersToListName(normalizedFormula, expectedListName) Then
        PID_RowHasValidFStundenDropdown = True
        Exit Function
    End If
    
    If InStr(1, UCase$(validationFormula), "KV_DROPDOWN_HELPER", vbTextCompare) > 0 Then
        Set allKeys = PID_BuildAllDropdownKeysForSheet(wsMonth)
        slotIndex = PID_GetKVCodeSlotIndexInSheet(allKeys, dropdownKey)
        expectedHelperCol = GetHelperColumnForKVCodeSlot(wsMonth.Name, slotIndex)
        
        If PID_ValidationFormulaUsesHelperColumn(UCase$(validationFormula), expectedHelperCol) Then
            PID_RowHasValidFStundenDropdown = True
        End If
    End If
End Function


Private Function PID_NormalizeValidationListFormula(ByVal formula As String) As String
    Dim normalized As String
    
    normalized = UCase$(Trim$(CStr(formula)))
    If Left$(normalized, 1) = "=" Then normalized = Mid$(normalized, 2)
    PID_NormalizeValidationListFormula = Replace(normalized, "$", "")
End Function


Private Function PID_ValidationFormulaRefersToListName(ByVal normalizedFormula As String, ByVal listName As String) As Boolean
    Dim pos As Long
    
    If normalizedFormula = listName Then
        PID_ValidationFormulaRefersToListName = True
        Exit Function
    End If
    
    If Len(listName) = 0 Then Exit Function
    If Len(normalizedFormula) < Len(listName) Then Exit Function
    
    If Right$(normalizedFormula, Len(listName)) <> listName Then Exit Function
    
    pos = Len(normalizedFormula) - Len(listName) + 1
    If pos = 1 Then
        PID_ValidationFormulaRefersToListName = True
    ElseIf Mid$(normalizedFormula, pos - 1, 1) = "!" Then
        PID_ValidationFormulaRefersToListName = True
    End If
End Function


Private Function PID_ValidationFormulaUsesHelperColumn(ByVal formulaUpper As String, ByVal helperCol As Long) As Boolean
    Dim colLetters As String
    
    colLetters = PID_GetExcelColumnLetters(helperCol)
    If colLetters = "" Then Exit Function
    
    PID_ValidationFormulaUsesHelperColumn = (InStr(1, formulaUpper, "$" & colLetters & "$", vbTextCompare) > 0)
End Function


Private Function PID_GetExcelColumnLetters(ByVal columnNumber As Long) As String
    Dim n As Long
    Dim remainder As Long
    Dim letters As String
    
    n = columnNumber
    letters = ""
    
    Do While n > 0
        remainder = (n - 1) Mod 26
        letters = Chr$(65 + remainder) & letters
        n = (n - 1) \ 26
    Loop
    
    PID_GetExcelColumnLetters = letters
End Function


Public Sub PID_RestoreFStundenDropdownOnSheet(ByVal wsMonth As Worksheet)
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    
    RefreshKVStundenDropdownForSheet wsMonth

SafeExit:
End Sub


Public Function PID_RestoreKVCodeDropdownOnSheet(ByVal ws As Worksheet) As Boolean
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Function
    
    wasProtected = ws.ProtectContents
    If wasProtected Then
        On Error Resume Next
        ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
        On Error GoTo SafeExit
    End If
    
    PID_ApplyKVCodeDropdownValidation ws
    ws.Range("F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW).Locked = False
    PID_MarkKVCodeDropdownValidForSheet ws.Name
    PID_RestoreKVCodeDropdownOnSheet = True

SafeExit:
    On Error Resume Next
    If wasProtected Then
        PID_ReprotectWorksheet ws
    End If
End Function


Public Sub PID_RestoreMonthSheetDropdownsAfterFormat()
    PID_RestoreMonthSheetDropdownsAfterFormatSilent
    
    MsgBox "Monatsblatt-Dropdowns (E/F) wiederhergestellt." & vbCrLf & vbCrLf & _
           "E = KV-Code, F = Stunden (auch leere Zeilen)." & vbCrLf & _
           "Nach E-" & PID_UTxtAenderung() & " aktualisiert sich F automatisch.", _
           vbInformation, "Dropdowns"
End Sub


Public Sub PID_RestoreMonthSheetDropdownsAfterFormatSilent()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If PID_IsWorkerMonthSheet(ws) Then
                PID_RestoreKVCodeDropdownOnSheet ws
                RefreshKVStundenDropdownForSheet ws
            End If
        End If
    Next i
    
    MarkKVDropdownsClean
    GoTo CleanExit

CleanFail:
    Err.Raise Err.Number, "PID_RestoreMonthSheetDropdownsAfterFormatSilent", Err.Description

CleanExit:
    PID_RestoreAktuelleStundenFormulasSilent
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


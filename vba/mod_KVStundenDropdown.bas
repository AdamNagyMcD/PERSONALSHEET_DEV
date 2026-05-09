Attribute VB_Name = "mod_KVStundenDropdown"
Option Explicit

Public gKVDropdownsDirty As Boolean


Public Sub MarkKVDropdownsDirty()
    gKVDropdownsDirty = True
End Sub


Public Sub MarkKVDropdownsClean()
    gKVDropdownsDirty = False
End Sub


Public Function AreKVDropdownsDirty() As Boolean
    AreKVDropdownsDirty = gKVDropdownsDirty
End Function


Public Sub RefreshKVDropdownsIfDirty()
    If gKVDropdownsDirty Then
        RefreshAllMonthKVStundenDropdowns
        gKVDropdownsDirty = False
    End If
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
    
    monthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            RefreshKVStundenDropdownForSheet ws
        End If
    Next i
    
    gKVDropdownsDirty = False

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    gKVDropdownsDirty = True
    
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei RefreshAllMonthKVStundenDropdowns:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "KV Stunden Dropdowns"
End Sub


Public Sub RefreshKVStundenDropdownForSheet(ByVal wsMonth As Worksheet, Optional ByVal changedRange As Range)
    Dim wsHelper As Worksheet
    Dim monthNumber As Long
    Dim r As Long
    Dim rowsToCheck As Range
    Dim c As Range
    Dim checkedRows As Collection
    Dim rowKey As String
    
    Dim oldDisplayAlerts As Boolean
    
    On Error GoTo CleanFail
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If Not IsNumeric(wsMonth.Range("A1").Value) Then Exit Sub
    
    monthNumber = CLng(wsMonth.Range("A1").Value)
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    oldDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    Set wsHelper = GetOrCreateKVDropdownHelperSheet()
    
    On Error Resume Next
    wsMonth.Unprotect Password:=PID_WORKBOOK_PASSWORD
    wsHelper.Visible = xlSheetVisible
    wsHelper.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    If changedRange Is Nothing Then
        
        ClearHelperColumnsForSheet wsHelper, wsMonth.Name
        
        For r = 3 To 82
            RefreshKVStundenDropdownForRow wsMonth, wsHelper, r, monthNumber
        Next r
        
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
    
    wsMonth.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not wsHelper Is Nothing Then
        wsHelper.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
        wsHelper.Visible = xlSheetVeryHidden
    End If
    
    If Not wsMonth Is Nothing Then
        wsMonth.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.DisplayAlerts = oldDisplayAlerts
    
    MsgBox "Fehler bei RefreshKVStundenDropdownForSheet:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "KV Stunden Dropdown"
End Sub


Public Sub RefreshKVStundenDropdownForRow(ByVal wsMonth As Worksheet, _
                                          ByVal wsHelper As Worksheet, _
                                          ByVal rowNumber As Long, _
                                          ByVal monthNumber As Long)
    Dim kvCode As String
    Dim values As Collection
    Dim helperCol As Long
    Dim helperLastRow As Long
    Dim listRange As Range
    Dim listName As String
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If wsHelper Is Nothing Then Exit Sub
    If rowNumber < 3 Or rowNumber > 82 Then Exit Sub
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    kvCode = NormalizeKVCodeForLookup(CStr(wsMonth.Cells(rowNumber, "E").Value))
    
    On Error Resume Next
    wsMonth.Cells(rowNumber, "F").Validation.Delete
    On Error GoTo SafeExit
    
    If kvCode = "" Then
        wsMonth.Cells(rowNumber, "F").ClearContents
        Exit Sub
    End If
    
    Set values = GetKVMonatsstundenValues(monthNumber, kvCode)
    
    If values.count = 0 Then
        wsMonth.Cells(rowNumber, "F").ClearContents
        Exit Sub
    End If
    
    helperCol = GetHelperColumnForMonthRow(wsMonth.Name, rowNumber)
    
    wsHelper.Columns(helperCol).Clear
    
    helperLastRow = WriteDropdownValuesToHelper(wsHelper, helperCol, values)
    
    If helperLastRow <= 0 Then Exit Sub
    
    Set listRange = wsHelper.Range(wsHelper.Cells(1, helperCol), wsHelper.Cells(helperLastRow, helperCol))
    
    listName = GetDropdownNameForMonthRow(wsMonth.Name, rowNumber)
    
    CreateOrReplaceWorkbookName listName, listRange
    
    With wsMonth.Cells(rowNumber, "F").Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

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
    Dim wsLohn As Worksheet
    Dim currentYear As Long
    Dim targetPeriod As String
    Dim previousPeriod As String
    Dim values As Collection
    
    Set values = New Collection
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    Set wsLohn = ThisWorkbook.Worksheets("LOHNTABELLE")
    
    If Not IsNumeric(wsLohn.Range("G3").Value) Then GoTo SafeExit
    
    currentYear = CLng(wsLohn.Range("G3").Value)
    
    targetPeriod = GetKVPeriodForWorkbookYear(currentYear, monthNumber)
    previousPeriod = GetPreviousKVPeriodForWorkbookYear(currentYear, monthNumber)
    
    AddMonatsstundenValuesFromPeriod wsKV, targetPeriod, kvCode, values
    
    If values.count = 0 Then
        AddMonatsstundenValuesFromPeriod wsKV, previousPeriod, kvCode, values
    End If

SafeExit:
    Set GetKVMonatsstundenValues = values
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
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    For r = 4 To lastRow
        rowPeriod = Trim$(CStr(wsKV.Cells(r, "A").Value))
        
        If rowPeriod = periodName Then
            rowKVCode = NormalizeKVCodeForLookup(CStr(wsKV.Cells(r, "D").Value))
            
            If rowKVCode = kvCode Then
                rowMonatsstunden = wsKV.Cells(r, "G").Value
                
                If IsNumeric(rowMonatsstunden) Then
                    keyText = CStr(CDbl(rowMonatsstunden))
                    
                    If Not CollectionHasKey_KVDropdown(values, keyText) Then
                        values.Add CDbl(rowMonatsstunden), keyText
                    End If
                End If
            End If
        End If
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
    Dim refersText As String
    
    On Error GoTo SafeExit
    
    If targetRange Is Nothing Then Exit Sub
    If Trim$(nameText) = "" Then Exit Sub
    
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo SafeExit
    
    refersText = "='" & targetRange.Worksheet.Name & "'!" & targetRange.Address(True, True)
    
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:=refersText

SafeExit:
End Sub


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


Public Function CollectionHasKey_KVDropdown(ByVal col As Collection, ByVal key As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    tmp = col.item(key)
    
    CollectionHasKey_KVDropdown = True
    Exit Function

NotFound:
    CollectionHasKey_KVDropdown = False
End Function


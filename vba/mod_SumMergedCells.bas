Attribute VB_Name = "mod_SumMergedCells"
Option Explicit

Public gFinanzSummaryDirty As Boolean
Private gFinanzSummaryDirtyMonth As Long


Public Function SumMergedCells(ByVal targetRange As Range) As Double
    Dim c As Range
    Dim mergeKey As String
    Dim handledMerges As Collection
    Dim valueToAdd As Variant
    Dim resultValue As Double
    
    ' Excel erkennt Aenderungen in zusammengefuehrten Zellen sonst oft nicht.
    Application.Volatile
    
    On Error GoTo SafeExit
    
    If targetRange Is Nothing Then
        SumMergedCells = 0
        Exit Function
    End If
    
    Set handledMerges = New Collection
    resultValue = 0
    
    For Each c In targetRange.Cells
        
        If c.MergeCells Then
            
            mergeKey = c.mergeArea.Address(External:=True)
            
            If Not SumMergedCellsCollectionHasKey(handledMerges, mergeKey) Then
                handledMerges.Add mergeKey, mergeKey
                
                valueToAdd = c.mergeArea.Cells(1, 1).Value
                
                If IsNumeric(valueToAdd) Then
                    resultValue = resultValue + CDbl(valueToAdd)
                End If
            End If
            
        Else
            
            valueToAdd = c.Value
            
            If IsNumeric(valueToAdd) Then
                resultValue = resultValue + CDbl(valueToAdd)
            End If
            
        End If
        
    Next c
    
    SumMergedCells = resultValue
    Exit Function

SafeExit:
    SumMergedCells = 0
End Function


Public Function CountMergedCellsValues(ByVal targetRange As Range) As Long
    Dim c As Range
    Dim mergeKey As String
    Dim handledMerges As Collection
    Dim valueToCheck As Variant
    Dim resultValue As Long
    
    On Error GoTo SafeExit
    
    If targetRange Is Nothing Then
        CountMergedCellsValues = 0
        Exit Function
    End If
    
    Set handledMerges = New Collection
    resultValue = 0
    
    For Each c In targetRange.Cells
        
        If c.MergeCells Then
            
            mergeKey = c.mergeArea.Address(External:=True)
            
            If Not SumMergedCellsCollectionHasKey(handledMerges, mergeKey) Then
                handledMerges.Add mergeKey, mergeKey
                
                valueToCheck = c.mergeArea.Cells(1, 1).Value
                
                If Trim$(CStr(valueToCheck)) <> "" Then
                    resultValue = resultValue + 1
                End If
            End If
            
        Else
            
            valueToCheck = c.Value
            
            If Trim$(CStr(valueToCheck)) <> "" Then
                resultValue = resultValue + 1
            End If
            
        End If
        
    Next c
    
    CountMergedCellsValues = resultValue
    Exit Function

SafeExit:
    CountMergedCellsValues = 0
End Function


Public Function AverageMergedCells(ByVal targetRange As Range) As Double
    Dim totalValue As Double
    Dim countValue As Long
    
    On Error GoTo SafeExit
    
    totalValue = SumMergedCells(targetRange)
    countValue = CountNumericMergedCellsValues(targetRange)
    
    If countValue > 0 Then
        AverageMergedCells = totalValue / countValue
    Else
        AverageMergedCells = 0
    End If
    
    Exit Function

SafeExit:
    AverageMergedCells = 0
End Function


Public Function CountNumericMergedCellsValues(ByVal targetRange As Range) As Long
    Dim c As Range
    Dim mergeKey As String
    Dim handledMerges As Collection
    Dim valueToCheck As Variant
    Dim resultValue As Long
    
    On Error GoTo SafeExit
    
    If targetRange Is Nothing Then
        CountNumericMergedCellsValues = 0
        Exit Function
    End If
    
    Set handledMerges = New Collection
    resultValue = 0
    
    For Each c In targetRange.Cells
        
        If c.MergeCells Then
            
            mergeKey = c.mergeArea.Address(External:=True)
            
            If Not SumMergedCellsCollectionHasKey(handledMerges, mergeKey) Then
                handledMerges.Add mergeKey, mergeKey
                
                valueToCheck = c.mergeArea.Cells(1, 1).Value
                
                If IsNumeric(valueToCheck) Then
                    resultValue = resultValue + 1
                End If
            End If
            
        Else
            
            valueToCheck = c.Value
            
            If IsNumeric(valueToCheck) Then
                resultValue = resultValue + 1
            End If
            
        End If
        
    Next c
    
    CountNumericMergedCellsValues = resultValue
    Exit Function

SafeExit:
    CountNumericMergedCellsValues = 0
End Function


Public Function SumVisibleMergedCells(ByVal targetRange As Range) As Double
    Dim c As Range
    Dim mergeKey As String
    Dim handledMerges As Collection
    Dim valueToAdd As Variant
    Dim resultValue As Double
    
    On Error GoTo SafeExit
    
    If targetRange Is Nothing Then
        SumVisibleMergedCells = 0
        Exit Function
    End If
    
    Set handledMerges = New Collection
    resultValue = 0
    
    For Each c In targetRange.Cells
        
        If Not c.EntireRow.Hidden Then
            If Not c.EntireColumn.Hidden Then
                
                If c.MergeCells Then
                    
                    mergeKey = c.mergeArea.Address(External:=True)
                    
                    If Not SumMergedCellsCollectionHasKey(handledMerges, mergeKey) Then
                        handledMerges.Add mergeKey, mergeKey
                        
                        valueToAdd = c.mergeArea.Cells(1, 1).Value
                        
                        If IsNumeric(valueToAdd) Then
                            resultValue = resultValue + CDbl(valueToAdd)
                        End If
                    End If
                    
                Else
                    
                    valueToAdd = c.Value
                    
                    If IsNumeric(valueToAdd) Then
                        resultValue = resultValue + CDbl(valueToAdd)
                    End If
                    
                End If
                
            End If
        End If
        
    Next c
    
    SumVisibleMergedCells = resultValue
    Exit Function

SafeExit:
    SumVisibleMergedCells = 0
End Function


Private Function SumMergedCellsCollectionHasKey(ByVal col As Collection, ByVal key As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    tmp = col.item(key)
    
    SumMergedCellsCollectionHasKey = True
    Exit Function

NotFound:
    SumMergedCellsCollectionHasKey = False
End Function


Public Sub MarkFinanzSummaryDirty()
    gFinanzSummaryDirty = True
    gFinanzSummaryDirtyMonth = 0
End Sub


Public Sub MarkFinanzSummaryDirtyForMonth(ByVal wsMonth As Worksheet)
    Dim monthIndex As Long
    
    gFinanzSummaryDirty = True
    gFinanzSummaryDirtyMonth = 0
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If Not IsNumeric(wsMonth.Range("A1").Value) Then Exit Sub
    
    monthIndex = CLng(wsMonth.Range("A1").Value)
    If monthIndex >= 1 And monthIndex <= 12 Then
        gFinanzSummaryDirtyMonth = monthIndex
    End If

SafeExit:
End Sub


Public Sub ClearFinanzSummaryDirty()
    gFinanzSummaryDirty = False
    gFinanzSummaryDirtyMonth = 0
End Sub


Public Sub RefreshFinanzSummaryIfDirty()
    If gFinanzSummaryDirty Then
        PID_SyncFinanzSummaryToUbersicht
    End If
End Sub


Public Sub PID_RecalculateAllMonthMergedFormulas()
    PID_SyncFinanzSummaryToUbersicht
End Sub


Public Sub PID_RecalculateFinanzSummaryChain()
    PID_SyncFinanzSummaryToUbersicht
End Sub


Public Sub PID_SyncFinanzSummaryToUbersicht()
    Dim monthNames As Variant
    Dim wsMonth As Worksheet
    Dim ubersichtWs As Worksheet
    Dim einstellungWs As Worksheet
    Dim monthRows As Variant
    Dim i As Long
    Dim monthIndex As Long
    Dim ubersichtProtected As Boolean
    Dim einstellungProtected As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    Dim syncSingleMonthOnly As Boolean
    
    On Error GoTo SafeExit
    
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    monthRows = Array(7, 8, 9, 11, 12, 13, 15, 16, 17, 19, 20, 21)
    monthNames = PID_MonthNames()
    syncSingleMonthOnly = (gFinanzSummaryDirtyMonth >= 1 And gFinanzSummaryDirtyMonth <= 12)
    
    Set ubersichtWs = Nothing
    Set einstellungWs = Nothing
    On Error Resume Next
    Set ubersichtWs = ThisWorkbook.Worksheets("UBERSICHT")
    Set einstellungWs = ThisWorkbook.Worksheets("EINSTELLUNG")
    On Error GoTo SafeExit
    
    If ubersichtWs Is Nothing Then GoTo RestoreSettings
    
    ubersichtProtected = ubersichtWs.ProtectContents
    PID_UnprotectWorksheet ubersichtWs
    
    If Not einstellungWs Is Nothing Then
        einstellungProtected = einstellungWs.ProtectContents
        PID_UnprotectWorksheet einstellungWs
    End If
    
    If syncSingleMonthOnly Then
        monthIndex = gFinanzSummaryDirtyMonth
        
        On Error Resume Next
        Set wsMonth = Nothing
        Set wsMonth = ThisWorkbook.Worksheets(CStr(monthNames(monthIndex - 1)))
        On Error GoTo SafeExit
        
        If Not wsMonth Is Nothing Then
            PID_WriteFinanzSummaryMonthRow ubersichtWs, einstellungWs, wsMonth, monthIndex, _
                                           CLng(monthRows(monthIndex - 1))
        End If
    Else
        For i = LBound(monthNames) To UBound(monthNames)
            monthIndex = i - LBound(monthNames) + 1
            
            On Error Resume Next
            Set wsMonth = Nothing
            Set wsMonth = ThisWorkbook.Worksheets(CStr(monthNames(i)))
            On Error GoTo SafeExit
            
            If Not wsMonth Is Nothing Then
                PID_WriteFinanzSummaryMonthRow ubersichtWs, einstellungWs, wsMonth, monthIndex, _
                                               CLng(monthRows(monthIndex - 1))
            End If
        Next i
    End If
    
    PID_SyncFinanzSummaryQuarterAndTotalRows ubersichtWs
    PID_RecalculateFinanzDiffColumns ubersichtWs
    
    If Not einstellungWs Is Nothing Then
        PID_ReprotectWorksheet einstellungWs, einstellungProtected
    End If
    
    PID_ReprotectWorksheet ubersichtWs, ubersichtProtected
    ClearFinanzSummaryDirty

RestoreSettings:
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating

SafeExit:
    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating
End Sub


Public Sub PID_RecalculateFinanzSummaryForMonth(ByVal ws As Worksheet)
    Dim monthIndex As Long
    Dim ubersichtWs As Worksheet
    Dim einstellungWs As Worksheet
    Dim monthRows As Variant
    Dim ubersichtRow As Long
    Dim crewLabor As Double
    Dim crewPct As Double
    Dim ubersichtProtected As Boolean
    Dim einstellungProtected As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Sub
    
    monthIndex = CLng(ws.Range("A1").Value)
    If monthIndex < 1 Or monthIndex > 12 Then Exit Sub
    
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    monthRows = Array(7, 8, 9, 11, 12, 13, 15, 16, 17, 19, 20, 21)
    ubersichtRow = CLng(monthRows(monthIndex - 1))
    
    crewLabor = PID_GetMonthCrewLaborValue(ws)
    crewPct = PID_GetMonthCrewLaborPct(ws, crewLabor)
    PID_RefreshMonthFinanzSummaryCells ws
    
    Set ubersichtWs = Nothing
    Set einstellungWs = Nothing
    On Error Resume Next
    Set ubersichtWs = ThisWorkbook.Worksheets("UBERSICHT")
    Set einstellungWs = ThisWorkbook.Worksheets("EINSTELLUNG")
    On Error GoTo RestoreSettings
    
    If ubersichtWs Is Nothing Then GoTo RestoreSettings
    
    ubersichtProtected = ubersichtWs.ProtectContents
    PID_UnprotectWorksheet ubersichtWs
    
    ubersichtWs.Cells(ubersichtRow, "G").Value2 = crewLabor
    ubersichtWs.Cells(ubersichtRow, "J").Value2 = crewPct
    
    If Not einstellungWs Is Nothing Then
        einstellungProtected = einstellungWs.ProtectContents
        PID_UnprotectWorksheet einstellungWs
        einstellungWs.Cells(21 + monthIndex, "E").Value2 = crewPct
        PID_ReprotectWorksheet einstellungWs, einstellungProtected
    End If
    
    PID_SyncFinanzSummaryQuarterAndTotalRows ubersichtWs
    PID_RecalculateFinanzDiffColumns ubersichtWs
    PID_ReprotectWorksheet ubersichtWs, ubersichtProtected
    ClearFinanzSummaryDirty

RestoreSettings:
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating

SafeExit:
    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating
End Sub


Public Sub PID_SyncFinanzSummaryForMonth(ByVal ws As Worksheet)
    PID_RecalculateFinanzSummaryForMonth ws
End Sub


Public Function PID_MonthChangeAffectsFinanzSummary(ByVal ws As Worksheet, ByVal changedRange As Range) As Boolean
    If PID_MonthChangeNeedsImmediateFinanzSync(ws, changedRange) Then
        PID_MonthChangeAffectsFinanzSummary = True
        Exit Function
    End If
    
    If PID_MonthChangeDefersFinanzSummarySync(ws, changedRange) Then
        PID_MonthChangeAffectsFinanzSummary = True
    End If
End Function


Public Function PID_MonthChangeNeedsImmediateFinanzSync(ByVal ws As Worksheet, ByVal changedRange As Range) As Boolean
    Dim watchRange As Range
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If changedRange Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Function
    
    Set watchRange = Union(ws.Range("Q17:R29"), ws.Range("S35"), ws.Range("O18:Q25"))
    
    If Not Intersect(changedRange, watchRange) Is Nothing Then
        PID_MonthChangeNeedsImmediateFinanzSync = True
    End If

SafeExit:
End Function


Public Function PID_MonthChangeDefersFinanzSummarySync(ByVal ws As Worksheet, ByVal changedRange As Range) As Boolean
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If changedRange Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Function
    
    If Not Intersect(changedRange, ws.Range("E3:L82")) Is Nothing Then
        PID_MonthChangeDefersFinanzSummarySync = True
    End If

SafeExit:
End Function


Private Sub PID_WriteFinanzSummaryMonthRow(ByVal ubersichtWs As Worksheet, _
                                           ByVal einstellungWs As Worksheet, _
                                           ByVal wsMonth As Worksheet, _
                                           ByVal monthIndex As Long, _
                                           ByVal ubersichtRow As Long)
    Dim crewLabor As Double
    Dim crewPct As Double
    
    On Error GoTo SafeExit
    
    If ubersichtWs Is Nothing Then Exit Sub
    If wsMonth Is Nothing Then Exit Sub
    If monthIndex < 1 Or monthIndex > 12 Then Exit Sub
    
    crewLabor = PID_GetMonthCrewLaborValue(wsMonth)
    crewPct = PID_GetMonthCrewLaborPct(wsMonth, crewLabor)
    
    ubersichtWs.Cells(ubersichtRow, "G").Value2 = crewLabor
    ubersichtWs.Cells(ubersichtRow, "J").Value2 = crewPct
    
    If Not einstellungWs Is Nothing Then
        einstellungWs.Cells(21 + monthIndex, "E").Value2 = crewPct
    End If

SafeExit:
End Sub


Private Function PID_GetMonthCrewLaborValue(ByVal ws As Worksheet) As Double
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    
    PID_GetMonthCrewLaborValue = SumMergedCells(ws.Range("Q17:R29"))

SafeExit:
End Function


Private Function PID_GetMonthCrewLaborPct(ByVal ws As Worksheet, ByVal crewLabor As Double) As Double
    Dim salesValue As Double
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    
    salesValue = CDbl(ws.Range("S35").Value2)
    
    If salesValue <> 0 Then
        PID_GetMonthCrewLaborPct = crewLabor / salesValue
    End If

SafeExit:
End Function


Private Sub PID_SyncFinanzSummaryQuarterAndTotalRows(ByVal ubersichtWs As Worksheet)
    Dim quarterRows As Variant
    Dim monthGroups As Variant
    Dim q As Long
    Dim m As Long
    Dim sumG As Double
    Dim sumJ As Double
    Dim totalG As Double
    Dim totalJ As Double
    
    On Error GoTo SafeExit
    
    If ubersichtWs Is Nothing Then Exit Sub
    
    quarterRows = Array(10, 14, 18, 22)
    monthGroups = Array(Array(7, 8, 9), Array(11, 12, 13), Array(15, 16, 17), Array(19, 20, 21))
    
    totalG = 0
    totalJ = 0
    
    For q = 0 To 3
        sumG = 0
        sumJ = 0
        
        For m = 0 To 2
            sumG = sumG + CDbl(ubersichtWs.Cells(CLng(monthGroups(q)(m)), "G").Value2)
            sumJ = sumJ + CDbl(ubersichtWs.Cells(CLng(monthGroups(q)(m)), "J").Value2)
        Next m
        
        ubersichtWs.Cells(CLng(quarterRows(q)), "G").Value2 = sumG
        ubersichtWs.Cells(CLng(quarterRows(q)), "J").Value2 = sumJ / 3
        
        totalG = totalG + sumG
        totalJ = totalJ + (sumJ / 3)
    Next q
    
    ubersichtWs.Cells(23, "G").Value2 = totalG
    ubersichtWs.Cells(23, "J").Value2 = totalJ / 4

SafeExit:
End Sub


Private Sub PID_RecalculateFinanzDiffColumns(ByVal ubersichtWs As Worksheet)
    Dim dataRow As Long
    
    On Error GoTo SafeExit
    
    If ubersichtWs Is Nothing Then Exit Sub
    
    For dataRow = 7 To 23
        ubersichtWs.Cells(dataRow, "H").Value2 = CDbl(ubersichtWs.Cells(dataRow, "G").Value2) - CDbl(ubersichtWs.Cells(dataRow, "F").Value2)
        ubersichtWs.Cells(dataRow, "K").Value2 = CDbl(ubersichtWs.Cells(dataRow, "J").Value2) - CDbl(ubersichtWs.Cells(dataRow, "I").Value2)
    Next dataRow

SafeExit:
End Sub


Private Sub PID_RefreshMonthFinanzSummaryCells(ByVal ws As Worksheet)
    Dim wasProtected As Boolean
    Dim formulaText As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    wasProtected = ws.ProtectContents
    PID_UnprotectWorksheet ws
    
    ws.Range("Q17:R29").Calculate
    
    formulaText = CStr(ws.Range("S36").Formula)
    If Len(formulaText) > 0 Then
        ws.Range("S36").Formula = formulaText
    End If
    
    formulaText = CStr(ws.Range("S37").Formula)
    If Len(formulaText) > 0 Then
        ws.Range("S37").Formula = formulaText
    End If
    
    ws.Range("S35").Calculate
    ws.Range("S36").Calculate
    ws.Range("S37").Calculate
    
    PID_ReprotectWorksheet ws, wasProtected

SafeExit:
End Sub


Private Sub PID_UnprotectWorksheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
    On Error GoTo 0
End Sub


Private Sub PID_ReprotectWorksheet(ByVal ws As Worksheet, ByVal wasProtected As Boolean)
    On Error Resume Next
    
    If wasProtected Then
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
    End If
    
    On Error GoTo 0
End Sub

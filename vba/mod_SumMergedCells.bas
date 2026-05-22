Attribute VB_Name = "mod_SumMergedCells"
Option Explicit


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


Public Sub PID_RecalculateAllMonthMergedFormulas()
    PID_SyncFinanzSummaryToUbersicht
End Sub


Public Sub PID_RecalculateFinanzSummaryChain()
    PID_SyncFinanzSummaryToUbersicht
End Sub


Public Sub PID_SyncFinanzSummaryToUbersicht()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    
    On Error GoTo SafeExit
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        
        If Not ws Is Nothing Then
            PID_SyncFinanzSummaryMonthToDisplay ws, i - LBound(monthNames) + 1
        End If
        
        On Error GoTo SafeExit
    Next i
    
    PID_SyncFinanzSummaryQuarterAndTotalRows
    PID_RecalculateFinanzDiffColumns

SafeExit:
End Sub


Public Sub PID_RecalculateFinanzSummaryForMonth(ByVal ws As Worksheet)
    Dim monthIndex As Long
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Sub
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Sub
    
    monthIndex = CLng(ws.Range("A1").Value)
    If monthIndex < 1 Or monthIndex > 12 Then Exit Sub
    
    PID_SyncFinanzSummaryMonthToDisplay ws, monthIndex
    PID_SyncFinanzSummaryQuarterAndTotalRows
    PID_RecalculateFinanzDiffColumns

SafeExit:
End Sub


Public Sub PID_SyncFinanzSummaryForMonth(ByVal ws As Worksheet)
    PID_RecalculateFinanzSummaryForMonth ws
End Sub


Public Function PID_MonthChangeAffectsFinanzSummary(ByVal ws As Worksheet, ByVal changedRange As Range) As Boolean
    Dim watchRange As Range
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If changedRange Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Function
    
    Set watchRange = Union(ws.Range("E3:L82"), ws.Range("Q17:R29"), ws.Range("S35"), ws.Range("O18:Q25"))
    
    If Not Intersect(changedRange, watchRange) Is Nothing Then
        PID_MonthChangeAffectsFinanzSummary = True
    End If

SafeExit:
End Function


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


Private Sub PID_SyncFinanzSummaryMonthToDisplay(ByVal ws As Worksheet, ByVal monthIndex As Long)
    Dim crewLabor As Double
    Dim crewPct As Double
    Dim ubersichtWs As Worksheet
    Dim einstellungWs As Worksheet
    Dim monthRows As Variant
    Dim ubersichtRow As Long
    Dim ubersichtWasProtected As Boolean
    Dim einstellungWasProtected As Boolean
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If monthIndex < 1 Or monthIndex > 12 Then Exit Sub
    
    monthRows = Array(7, 8, 9, 11, 12, 13, 15, 16, 17, 19, 20, 21)
    ubersichtRow = CLng(monthRows(monthIndex - 1))
    
    crewLabor = PID_GetMonthCrewLaborValue(ws)
    crewPct = PID_GetMonthCrewLaborPct(ws, crewLabor)
    
    PID_RefreshMonthFinanzSummaryCells ws
    
    On Error Resume Next
    Set einstellungWs = ThisWorkbook.Worksheets("EINSTELLUNG")
    If Not einstellungWs Is Nothing Then
        einstellungWasProtected = einstellungWs.ProtectContents
        PID_UnprotectWorksheet einstellungWs
        einstellungWs.Cells(21 + monthIndex, "E").Value2 = crewPct
        PID_ReprotectWorksheet einstellungWs, einstellungWasProtected
    End If
    
    Set ubersichtWs = ThisWorkbook.Worksheets("UBERSICHT")
    If Not ubersichtWs Is Nothing Then
        ubersichtWasProtected = ubersichtWs.ProtectContents
        PID_UnprotectWorksheet ubersichtWs
        ubersichtWs.Cells(ubersichtRow, "G").Value2 = crewLabor
        ubersichtWs.Cells(ubersichtRow, "J").Value2 = crewPct
        PID_ReprotectWorksheet ubersichtWs, ubersichtWasProtected
    End If
    On Error GoTo SafeExit

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


Private Sub PID_SyncFinanzSummaryQuarterAndTotalRows()
    Dim ubersichtWs As Worksheet
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set ubersichtWs = ThisWorkbook.Worksheets("UBERSICHT")
    If ubersichtWs Is Nothing Then Exit Sub
    
    wasProtected = ubersichtWs.ProtectContents
    PID_UnprotectWorksheet ubersichtWs
    
    ubersichtWs.Range("G10,G14,G18,G22,J10,J14,J18,J22,G23,J23").Calculate
    
    PID_ReprotectWorksheet ubersichtWs, wasProtected

SafeExit:
End Sub


Private Sub PID_RecalculateFinanzDiffColumns()
    Dim ubersichtWs As Worksheet
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set ubersichtWs = ThisWorkbook.Worksheets("UBERSICHT")
    If ubersichtWs Is Nothing Then Exit Sub
    
    wasProtected = ubersichtWs.ProtectContents
    PID_UnprotectWorksheet ubersichtWs
    
    ubersichtWs.Range("H7:H23,K7:K23").Calculate
    
    PID_ReprotectWorksheet ubersichtWs, wasProtected

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

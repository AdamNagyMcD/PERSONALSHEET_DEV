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
    PID_RecalculateFinanzSummaryChain
End Sub


Public Sub PID_RecalculateFinanzSummaryChain()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim einstellungWs As Worksheet
    Dim ubersichtWs As Worksheet
    Dim i As Long
    
    On Error GoTo SafeExit
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        
        If Not ws Is Nothing Then
            PID_RecalculateFinanzSummaryMonthCells ws
        End If
        
        On Error GoTo SafeExit
    Next i
    
    PID_RecalculateFinanzSummaryUbersichtCells

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
    
    PID_RecalculateFinanzSummaryMonthCells ws
    PID_RecalculateFinanzSummaryUbersichtForMonth monthIndex

SafeExit:
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


Private Sub PID_RecalculateFinanzSummaryMonthCells(ByVal ws As Worksheet)
    On Error Resume Next
    
    ws.Range("S35").Calculate
    ws.Range("S36").Calculate
    ws.Range("S37").Calculate
    ws.Range("Q37").Calculate
    ws.Range("Q42").Calculate
    
    On Error GoTo 0
End Sub


Private Sub PID_RecalculateFinanzSummaryUbersichtCells()
    Dim einstellungWs As Worksheet
    Dim ubersichtWs As Worksheet
    
    On Error Resume Next
    
    Set einstellungWs = ThisWorkbook.Worksheets("EINSTELLUNG")
    If Not einstellungWs Is Nothing Then
        einstellungWs.Range("E22:E33").Calculate
    End If
    
    Set ubersichtWs = ThisWorkbook.Worksheets("UBERSICHT")
    If Not ubersichtWs Is Nothing Then
        ubersichtWs.Range("G7:G23").Calculate
        ubersichtWs.Range("J7:J23").Calculate
        ubersichtWs.Range("H7:H23").Calculate
        ubersichtWs.Range("K7:K23").Calculate
    End If
    
    On Error GoTo 0
End Sub


Private Sub PID_RecalculateFinanzSummaryUbersichtForMonth(ByVal monthIndex As Long)
    Dim einstellungWs As Worksheet
    Dim ubersichtWs As Worksheet
    Dim ubersichtRow As Long
    Dim quarterRow As Long
    
    On Error Resume Next
    
    Set einstellungWs = ThisWorkbook.Worksheets("EINSTELLUNG")
    If Not einstellungWs Is Nothing Then
        einstellungWs.Cells(21 + monthIndex, "E").Calculate
    End If
    
    Set ubersichtWs = ThisWorkbook.Worksheets("UBERSICHT")
    If Not ubersichtWs Is Nothing Then
        ubersichtRow = 6 + monthIndex
        quarterRow = 10 + 4 * Int((monthIndex - 1) / 3)
        
        ubersichtWs.Range("G" & ubersichtRow & ",J" & ubersichtRow & ",H" & ubersichtRow & ",K" & ubersichtRow).Calculate
        ubersichtWs.Range("G" & quarterRow & ",J" & quarterRow & ",H" & quarterRow & ",K" & quarterRow).Calculate
        ubersichtWs.Range("G23,J23,H23,K23").Calculate
    End If
    
    On Error GoTo 0
End Sub

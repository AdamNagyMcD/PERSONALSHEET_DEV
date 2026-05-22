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
            ws.Range("S36").Calculate
        End If
        
        On Error GoTo SafeExit
    Next i

SafeExit:
End Sub

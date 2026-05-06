Attribute VB_Name = "mod_KVLohnLookup"
Option Explicit

Public Sub RefreshKVLohnForSheet(ByVal wsMonth As Worksheet, Optional ByVal changedRange As Range)
    Dim r As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim monthNumber As Long
    Dim rowsToCheck As Range
    Dim c As Range
    Dim checkedRows As Collection
    Dim rowKey As String
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If Not IsNumeric(wsMonth.Range("A1").Value) Then Exit Sub
    
    monthNumber = CLng(wsMonth.Range("A1").Value)
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    firstRow = 3
    lastRow = 82
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set checkedRows = New Collection
    
    If changedRange Is Nothing Then
        
        For r = firstRow To lastRow
            RefreshKVLohnForRow wsMonth, r, monthNumber
        Next r
        
    Else
        
        Set rowsToCheck = Intersect(changedRange, wsMonth.Range("E3:F82"))
        
        If rowsToCheck Is Nothing Then GoTo CleanExit
        
        For Each c In rowsToCheck.Cells
            rowKey = CStr(c.Row)
            
            If Not CollectionHasKey_KVLohn(checkedRows, rowKey) Then
                checkedRows.Add rowKey, rowKey
                RefreshKVLohnForRow wsMonth, c.Row, monthNumber
            End If
        Next c
        
    End If

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei RefreshKVLohnForSheet:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "KV Lohn"
End Sub


Public Sub RefreshKVLohnForRow(ByVal wsMonth As Worksheet, ByVal rowNumber As Long, ByVal monthNumber As Long)
    Dim kvCode As String
    Dim monatsstunden As Variant
    Dim lohnValue As Variant
    Dim usedFallback As Boolean
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If rowNumber < 3 Or rowNumber > 82 Then Exit Sub
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    kvCode = NormalizeKVCodeForLookup(CStr(wsMonth.Cells(rowNumber, "E").Value))
    monatsstunden = wsMonth.Cells(rowNumber, "F").Value
    
    If kvCode <> "" And IsNumeric(monatsstunden) Then
        
        usedFallback = False
        lohnValue = GetKVLohnByPeriod(monthNumber, kvCode, CDbl(monatsstunden), usedFallback)
        
        If IsError(lohnValue) Then
            wsMonth.Cells(rowNumber, "G").Value = "Nicht gefunden"
            wsMonth.Cells(rowNumber, "G").NumberFormat = "General"
        Else
            wsMonth.Cells(rowNumber, "G").Value = CDbl(lohnValue)
            ApplyEuroNumberFormatToRange wsMonth.Cells(rowNumber, "G")
        End If
        
    Else
        
        wsMonth.Cells(rowNumber, "G").ClearContents
        ApplyEuroNumberFormatToRange wsMonth.Cells(rowNumber, "G")
        
    End If

SafeExit:
End Sub


Public Sub RefreshAllMonthKVLohn()
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
            RefreshKVLohnForSheet ws
        End If
    Next i

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei RefreshAllMonthKVLohn:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "KV Lohn"
End Sub


Public Function GetKVLohnByPeriod(ByVal monthNumber As Long, _
                                  ByVal kvCode As String, _
                                  ByVal monatsstunden As Double, _
                                  Optional ByRef usedFallback As Boolean = False) As Variant
    Dim wsLohn As Worksheet
    Dim currentYear As Long
    Dim targetPeriod As String
    Dim previousPeriod As String
    Dim resultValue As Variant
    
    On Error GoTo NotFound
    
    Set wsLohn = ThisWorkbook.Worksheets("LOHNTABELLE")
    
    If Not IsNumeric(wsLohn.Range("G3").Value) Then GoTo NotFound
    
    currentYear = CLng(wsLohn.Range("G3").Value)
    
    If monthNumber < 1 Or monthNumber > 12 Then GoTo NotFound
    
    targetPeriod = GetKVPeriodForWorkbookYear(currentYear, monthNumber)
    previousPeriod = GetPreviousKVPeriodForWorkbookYear(currentYear, monthNumber)
    
    usedFallback = False
    
    resultValue = FindKVLohnInPeriod(targetPeriod, kvCode, monatsstunden)
    
    If Not IsError(resultValue) Then
        If IsNumeric(resultValue) Then
            GetKVLohnByPeriod = CDbl(resultValue)
            Exit Function
        End If
    End If
    
    resultValue = FindKVLohnInPeriod(previousPeriod, kvCode, monatsstunden)
    
    If Not IsError(resultValue) Then
        If IsNumeric(resultValue) Then
            usedFallback = True
            GetKVLohnByPeriod = CDbl(resultValue)
            Exit Function
        End If
    End If

NotFound:
    GetKVLohnByPeriod = CVErr(xlErrNA)
End Function


Public Function FindKVLohnInPeriod(ByVal periodName As String, _
                                   ByVal kvCode As String, _
                                   ByVal monatsstunden As Double) As Variant
    Dim wsKV As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    Dim rowPeriod As String
    Dim rowKVCode As String
    Dim rowMonatsstunden As Variant
    Dim rowLohn As Variant
    Dim rowStatus As String
    
    On Error GoTo NotFound
    
    If Trim$(periodName) = "" Then GoTo NotFound
    If Trim$(kvCode) = "" Then GoTo NotFound
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    For r = 4 To lastRow
        rowPeriod = Trim$(CStr(wsKV.Cells(r, "A").Value))
        
        If rowPeriod = periodName Then
            
            rowKVCode = NormalizeKVCodeForLookup(CStr(wsKV.Cells(r, "D").Value))
            
            If rowKVCode = kvCode Then
                
                rowMonatsstunden = wsKV.Cells(r, "G").Value
                
                If IsNumeric(rowMonatsstunden) Then
                    
                    If Abs(CDbl(rowMonatsstunden) - monatsstunden) < 0.001 Then
                        
                        rowLohn = wsKV.Cells(r, "H").Value
                        rowStatus = Trim$(CStr(wsKV.Cells(r, "I").Value))
                        
                        If rowStatus = "OK" Then
                            If IsNumeric(rowLohn) Then
                                FindKVLohnInPeriod = CDbl(rowLohn)
                                Exit Function
                            End If
                        End If
                        
                        GoTo NotFound
                    End If
                    
                End If
                
            End If
            
        End If
    Next r

NotFound:
    FindKVLohnInPeriod = CVErr(xlErrNA)
End Function


Public Function NormalizeKVCodeForLookup(ByVal valueToNormalize As String) As String
    Dim s As String
    
    s = Trim$(CStr(valueToNormalize))
    
    If s = "" Then
        NormalizeKVCodeForLookup = ""
        Exit Function
    End If
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    s = Replace(s, "-", "_")
    s = Replace(s, " ", "_")
    
    Select Case UCase$(s)
        Case "BG1"
            NormalizeKVCodeForLookup = "BG1_Basis"
        
        Case "BG2"
            NormalizeKVCodeForLookup = "BG2_Basis"
        
        Case "BG3"
            NormalizeKVCodeForLookup = "BG3_Basis"
        
        Case "BG1_BASIS"
            NormalizeKVCodeForLookup = "BG1_Basis"
        
        Case "BG1_5"
            NormalizeKVCodeForLookup = "BG1_5"
        
        Case "BG1_10"
            NormalizeKVCodeForLookup = "BG1_10"
        
        Case "BG1_15"
            NormalizeKVCodeForLookup = "BG1_15"
        
        Case "BG2_BASIS"
            NormalizeKVCodeForLookup = "BG2_Basis"
        
        Case "BG2_5"
            NormalizeKVCodeForLookup = "BG2_5"
        
        Case "BG2_10"
            NormalizeKVCodeForLookup = "BG2_10"
        
        Case "BG2_15"
            NormalizeKVCodeForLookup = "BG2_15"
        
        Case "BG3_BASIS"
            NormalizeKVCodeForLookup = "BG3_Basis"
        
        Case "BG3_5"
            NormalizeKVCodeForLookup = "BG3_5"
        
        Case "BG3_10"
            NormalizeKVCodeForLookup = "BG3_10"
        
        Case "BG3_15"
            NormalizeKVCodeForLookup = "BG3_15"
        
        Case Else
            NormalizeKVCodeForLookup = s
    End Select
End Function


Public Function GetKVPeriodForWorkbookYear(ByVal workbookYear As Long, ByVal monthNumber As Long) As String
    If monthNumber >= 1 And monthNumber <= 4 Then
        GetKVPeriodForWorkbookYear = "KV " & CStr(workbookYear - 1) & "/" & CStr(workbookYear)
    Else
        GetKVPeriodForWorkbookYear = "KV " & CStr(workbookYear) & "/" & CStr(workbookYear + 1)
    End If
End Function


Public Function GetPreviousKVPeriodForWorkbookYear(ByVal workbookYear As Long, ByVal monthNumber As Long) As String
    If monthNumber >= 1 And monthNumber <= 4 Then
        GetPreviousKVPeriodForWorkbookYear = "KV " & CStr(workbookYear - 2) & "/" & CStr(workbookYear - 1)
    Else
        GetPreviousKVPeriodForWorkbookYear = "KV " & CStr(workbookYear - 1) & "/" & CStr(workbookYear)
    End If
End Function


Public Function CollectionHasKey_KVLohn(ByVal col As Collection, ByVal key As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    tmp = col.item(key)
    
    CollectionHasKey_KVLohn = True
    Exit Function

NotFound:
    CollectionHasKey_KVLohn = False
End Function


Private Sub ApplyEuroNumberFormatToRange(ByVal targetRange As Range)
    Dim euroSymbol As String
    
    If targetRange Is Nothing Then Exit Sub
    
    euroSymbol = ChrW(8364)
    
    On Error GoTo TryEnglishFormat
    
    ' Deutsch / Oesterreich Excel: Û 2.328,00
    targetRange.NumberFormatLocal = euroSymbol & " #.##0,00"
    Exit Sub

TryEnglishFormat:
    On Error GoTo SafeExit
    
    ' Fallback fuer andere Excel-Sprachen: Û 2,328.00
    targetRange.NumberFormat = euroSymbol & " #,##0.00"

SafeExit:
End Sub


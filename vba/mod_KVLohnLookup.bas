Attribute VB_Name = "mod_KVLohnLookup"
Option Explicit

' Module-level cache for LOHNTABELLE data.
' Populated once per RefreshKVLohnForSheet call, cleared afterwards.
' Columns in cache match Range("A4:I<lastRow>"):
'   1=A(Period), 2=B, 3=C, 4=D(KVCode), 5=E, 6=F, 7=G(Stunden), 8=H(Lohn), 9=I(Status)
Public gKVLohnAllMonthsDirty As Boolean
Private mKVLohnRefreshedSheets As Collection
Private mLohnTableCache As Variant
Private mLohnTableCacheLoaded As Boolean
Private mBatchKVLohnRefresh As Boolean
Private mCachedWorkbookYear As Long
Private mWorkbookYearCached As Boolean


Private Sub PID_EnsureLohnTableCacheLoaded()
    If mLohnTableCacheLoaded Then Exit Sub
    PID_LoadLohnTableCache
End Sub


Private Sub PID_LoadLohnTableCache()
    Dim wsKV As Worksheet
    Dim lastRow As Long
    
    mLohnTableCacheLoaded = False
    
    On Error Resume Next
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    On Error GoTo 0
    
    If wsKV Is Nothing Then
        mLohnTableCacheLoaded = True
        Exit Sub
    End If
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    
    If lastRow >= 4 Then
        mLohnTableCache = wsKV.Range("A4:I" & lastRow).Value
    End If
    
    mLohnTableCacheLoaded = True
End Sub


Private Sub PID_ClearLohnTableCache()
    mLohnTableCacheLoaded = False
    mWorkbookYearCached = False
    On Error Resume Next
    Erase mLohnTableCache
    On Error GoTo 0
End Sub


Private Sub PID_EnsureWorkbookYearCached()
    If mWorkbookYearCached Then Exit Sub
    
    mCachedWorkbookYear = PID_GetWorkbookYear()
    mWorkbookYearCached = True
End Sub

Public Sub RefreshKVLohnForSheet(ByVal wsMonth As Worksheet, _
                                 Optional ByVal changedRange As Range, _
                                 Optional ByVal preserveGOnMiss As Boolean = False)
    Dim r As Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim monthNumber As Long
    Dim rowsToCheck As Range
    Dim c As Range
    Dim checkedRows As Collection
    Dim rowKey As String
    Dim wasProtected As Boolean
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If Not IsNumeric(wsMonth.Range("A1").Value) Then Exit Sub
    
    monthNumber = CLng(wsMonth.Range("A1").Value)
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    firstRow = PID_FIRST_ROW
    lastRow = PID_LAST_ROW
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set checkedRows = New Collection
    
    wasProtected = wsMonth.ProtectContents
    On Error Resume Next
    wsMonth.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    PID_LoadLohnTableCache
    PID_EnsureWorkbookYearCached
    mBatchKVLohnRefresh = True
    
    If changedRange Is Nothing Then
        
        For r = firstRow To lastRow
            RefreshKVLohnForRow wsMonth, r, monthNumber, preserveGOnMiss
        Next r
        
        PID_ApplyEuroNumberFormat wsMonth.Range("G" & firstRow & ":G" & lastRow)
        
    Else
        
        Set rowsToCheck = Intersect(changedRange, wsMonth.Range("E3:F82"))
        
        If rowsToCheck Is Nothing Then GoTo CleanExit
        
        For Each c In rowsToCheck.Cells
            rowKey = CStr(c.Row)
            
            If Not CollectionHasKey_KVLohn(checkedRows, rowKey) Then
                checkedRows.Add rowKey, rowKey
                RefreshKVLohnForRow wsMonth, c.Row, monthNumber, preserveGOnMiss
            End If
        Next c
        
        PID_ApplyEuroNumberFormat wsMonth.Range("G" & firstRow & ":G" & lastRow)
        
    End If

CleanExit:
    mBatchKVLohnRefresh = False
    PID_ClearLohnTableCache
    
    On Error Resume Next
    If wasProtected Then
        wsMonth.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    mBatchKVLohnRefresh = False
    PID_ClearLohnTableCache
    
    On Error Resume Next
    If wasProtected Then
        wsMonth.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei RefreshKVLohnForSheet:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "KV Lohn"
End Sub


Public Sub RefreshKVLohnForRow(ByVal wsMonth As Worksheet, _
                               ByVal rowNumber As Long, _
                               ByVal monthNumber As Long, _
                               Optional ByVal preserveGOnMiss As Boolean = False)
    Dim kvCode As String
    Dim monatsstunden As Variant
    Dim monatsstundenValue As Double
    Dim lohnValue As Variant
    Dim usedFallback As Boolean
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Sub
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    If Not mBatchKVLohnRefresh Then
        wasProtected = wsMonth.ProtectContents
        If wasProtected Then
            On Error Resume Next
            wsMonth.Unprotect Password:=PID_WORKBOOK_PASSWORD
            On Error GoTo SafeExit
        End If
    End If
    
    kvCode = NormalizeKVCodeForLookup(CStr(wsMonth.Cells(rowNumber, "E").Value))
    monatsstunden = wsMonth.Cells(rowNumber, "F").Value
    
    If kvCode <> "" And PID_TryGetDouble(monatsstunden, monatsstundenValue) Then
        
        usedFallback = False
        lohnValue = GetKVLohnByPeriod(monthNumber, kvCode, monatsstundenValue, usedFallback)
        
        If IsError(lohnValue) Then
            If Not preserveGOnMiss Then
                wsMonth.Cells(rowNumber, "G").Value = "Nicht gefunden"
                wsMonth.Cells(rowNumber, "G").NumberFormat = "General"
            End If
        Else
            wsMonth.Cells(rowNumber, "G").Value = CDbl(lohnValue)
        End If
        
    Else
        
        If Not preserveGOnMiss Then
            wsMonth.Cells(rowNumber, "G").ClearContents
        End If
        
    End If

SafeExit:
    If Not mBatchKVLohnRefresh Then
        On Error Resume Next
        If wasProtected Then
            wsMonth.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
        End If
    End If
End Sub


Private Function PID_TryGetDouble(ByVal valueToParse As Variant, ByRef resultValue As Double) As Boolean
    Dim s As String
    Dim decimalSeparator As String
    Dim thousandSeparator As String
    
    On Error GoTo ParseFail
    
    If IsNumeric(valueToParse) Then
        resultValue = CDbl(valueToParse)
        PID_TryGetDouble = True
        Exit Function
    End If
    
    s = Trim$(CStr(valueToParse))
    If s = "" Then Exit Function
    
    decimalSeparator = CStr(Application.International(xlDecimalSeparator))
    thousandSeparator = CStr(Application.International(xlThousandsSeparator))
    
    If thousandSeparator <> "" Then
        s = Replace(s, thousandSeparator, "")
    End If
    
    If decimalSeparator = "," Then
        s = Replace(s, ".", ",")
    Else
        s = Replace(s, ",", ".")
    End If
    
    If IsNumeric(s) Then
        resultValue = CDbl(s)
        PID_TryGetDouble = True
    End If
    
    Exit Function
    
ParseFail:
    PID_TryGetDouble = False
End Function


Public Sub RefreshAllMonthKVLohn()
    PID_RestoreMonatslohnFormulasSilent
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
    
    PID_EnsureWorkbookYearCached
    If Not mWorkbookYearCached Then GoTo NotFound
    
    currentYear = mCachedWorkbookYear
    
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
    Dim rowCount As Long
    
    Dim rowPeriod As String
    Dim rowKVCode As String
    Dim rowMonatsstunden As Variant
    Dim rowLohn As Variant
    Dim rowStatus As String
    Dim normalizedTargetPeriod As String
    
    On Error GoTo NotFound
    
    normalizedTargetPeriod = NormalizeKVPeriodForLookup(periodName)
    If normalizedTargetPeriod = "" Then GoTo NotFound
    If Trim$(kvCode) = "" Then GoTo NotFound
    
    If mLohnTableCacheLoaded And Not IsEmpty(mLohnTableCache) Then
        
        rowCount = UBound(mLohnTableCache, 1)
        
        For r = 1 To rowCount
            rowPeriod = NormalizeKVPeriodForLookup(CStr(mLohnTableCache(r, 1)))
            
            If rowPeriod = normalizedTargetPeriod Then
                rowKVCode = NormalizeKVCodeForLookup(CStr(mLohnTableCache(r, 4)))
                
                If rowKVCode = kvCode Then
                    rowMonatsstunden = mLohnTableCache(r, 7)
                    
                    If IsNumeric(rowMonatsstunden) Then
                        If Abs(CDbl(rowMonatsstunden) - monatsstunden) < 0.001 Then
                            rowLohn = mLohnTableCache(r, 8)
                            rowStatus = Trim$(CStr(mLohnTableCache(r, 9)))
                            
                            If PID_IsUsableKVLookupLohn(rowLohn, rowStatus) Then
                                FindKVLohnInPeriod = CDbl(rowLohn)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next r
        
    Else
        
        Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
        
        lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
        
        For r = 4 To lastRow
            rowPeriod = NormalizeKVPeriodForLookup(CStr(wsKV.Cells(r, "A").Value))
            
            If rowPeriod = normalizedTargetPeriod Then
                rowKVCode = NormalizeKVCodeForLookup(CStr(wsKV.Cells(r, "D").Value))
                
                If rowKVCode = kvCode Then
                    rowMonatsstunden = wsKV.Cells(r, "G").Value
                    
                    If IsNumeric(rowMonatsstunden) Then
                        If Abs(CDbl(rowMonatsstunden) - monatsstunden) < 0.001 Then
                            rowLohn = wsKV.Cells(r, "H").Value
                            rowStatus = Trim$(CStr(wsKV.Cells(r, "I").Value))
                            
                            If PID_IsUsableKVLookupLohn(rowLohn, rowStatus) Then
                                FindKVLohnInPeriod = CDbl(rowLohn)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next r
        
    End If

NotFound:
    FindKVLohnInPeriod = CVErr(xlErrNA)
End Function


Public Function NormalizeKVPeriodForLookup(ByVal periodText As String) As String
    Dim s As String
    
    s = Trim$(CStr(periodText))
    If s = "" Then Exit Function
    
    If InStr(1, s, "|", vbTextCompare) > 0 Then
        s = Trim$(Left$(s, InStr(1, s, "|", vbTextCompare) - 1))
    End If
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeKVPeriodForLookup = s
End Function


Private Function PID_IsUsableKVLookupLohn(ByVal rowLohn As Variant, ByVal rowStatus As String) As Boolean
    On Error GoTo SafeExit
    
    If Not IsNumeric(rowLohn) Then Exit Function
    If CDbl(rowLohn) <= 0# Then Exit Function
    
    If UCase$(Trim$(rowStatus)) = "OK" Then
        PID_IsUsableKVLookupLohn = True
        Exit Function
    End If
    
    ' Leerer Lohn in neuer Periode: trotzdem gueltigen numerischen Lohn aus Vorperiode zulassen.
    PID_IsUsableKVLookupLohn = True
    
SafeExit:
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


Public Function PID_KVLohnLookup(ByVal monthNumber As Variant, _
                                 ByVal kvCode As Variant, _
                                 ByVal monatsstunden As Variant, _
                                 Optional ByVal lohnTableTouch As Variant) As Variant
    Dim monthNum As Long
    Dim stundenValue As Double
    Dim normalizedCode As String
    Dim usedFallback As Boolean
    Dim resultValue As Variant
    
    On Error GoTo SafeExit
    
    If IsEmpty(monthNumber) Or IsEmpty(kvCode) Or IsEmpty(monatsstunden) Then
        PID_KVLohnLookup = ""
        Exit Function
    End If
    
    If Len(Trim$(CStr(kvCode))) = 0 Or Len(Trim$(CStr(monatsstunden))) = 0 Then
        PID_KVLohnLookup = ""
        Exit Function
    End If
    
    If Not IsNumeric(monthNumber) Then
        PID_KVLohnLookup = ""
        Exit Function
    End If
    
    monthNum = CLng(monthNumber)
    If monthNum < 1 Or monthNum > 12 Then
        PID_KVLohnLookup = ""
        Exit Function
    End If
    
    normalizedCode = NormalizeKVCodeForLookup(CStr(kvCode))
    If normalizedCode = "" Then
        PID_KVLohnLookup = ""
        Exit Function
    End If
    
    If Not PID_TryGetDouble(monatsstunden, stundenValue) Then
        PID_KVLohnLookup = ""
        Exit Function
    End If
    
    PID_EnsureLohnTableCacheLoaded
    PID_EnsureWorkbookYearCached
    
    resultValue = GetKVLohnByPeriod(monthNum, normalizedCode, stundenValue, usedFallback)
    
    If IsError(resultValue) Then
        PID_KVLohnLookup = "Nicht gefunden"
    Else
        PID_KVLohnLookup = CDbl(resultValue)
    End If
    Exit Function

SafeExit:
    PID_KVLohnLookup = "Nicht gefunden"
End Function


Public Function PID_GetMonatslohnFormulaR1C1() As String
    Dim kvSheet As String
    
    kvSheet = "'" & PID_LOHNTABELLE_SHEET & "'"
    
    ' UDF mirrors VBA lookup (BG3_15 etc.) and LOHNTABELLE anchor keeps recalc on table edits.
    PID_GetMonatslohnFormulaR1C1 = _
        "=IF(OR(RC[-2]="""",RC[-1]=""""),""""," & _
        "PID_KVLohnLookup(R1C1,RC[-2],RC[-1]," & kvSheet & "!R4C8))"
End Function


Public Function PID_MonthSheetHasMonatslohnFormula(ByVal wsMonth As Worksheet) As Boolean
    Dim formulaText As String
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Function
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Function
    
    formulaText = CStr(wsMonth.Range("G" & PID_FIRST_ROW).FormulaR1C1)
    PID_MonthSheetHasMonatslohnFormula = (InStr(1, formulaText, "PID_KVLohnLookup", vbTextCompare) > 0)

SafeExit:
End Function


Public Sub PID_RestoreMonatslohnFormulas()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim updatedCount As Long
    Dim formulaR1C1 As String
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    formulaR1C1 = PID_GetMonatslohnFormulaR1C1()
    monthNames = PID_MonthNames()
    updatedCount = 0
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If PID_RestoreMonatslohnFormulasOnSheet(ws, formulaR1C1) Then
                updatedCount = updatedCount + 1
            End If
        End If
    Next i
    
    ClearAllKVLohnDirty
    
    MsgBox "Monatslohn-Formeln (Spalte G) wurden wiederhergestellt." & vbCrLf & vbCrLf & _
           "Monatsblaetter aktualisiert: " & CStr(updatedCount) & " / 12" & vbCrLf & _
           "Bereich: G" & PID_FIRST_ROW & ":G" & PID_LAST_ROW & vbCrLf & _
           "Lookup aus: " & PID_LOHNTABELLE_SHEET & vbCrLf & vbCrLf & _
           "Spalte G aktualisiert sich jetzt automatisch bei Aenderung von E/F.", _
           vbInformation, "Spalte G"
    
    GoTo CleanExit

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_RestoreMonatslohnFormulas:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Spalte G"
End Sub


Public Sub PID_EnsureMonatslohnFormulas(Optional ByVal showMessage As Boolean = False)
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Januar")
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    If PID_MonthSheetHasMonatslohnFormula(ws) Then Exit Sub
    
    If showMessage Then
        PID_RestoreMonatslohnFormulas
    Else
        PID_RestoreMonatslohnFormulasSilent
    End If

SafeExit:
End Sub


Public Sub PID_EnsureMonatslohnFormulasOnSheet(ByVal wsMonth As Worksheet)
    Dim formulaR1C1 As String
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If PID_MonthSheetHasMonatslohnFormula(wsMonth) Then Exit Sub
    
    formulaR1C1 = PID_GetMonatslohnFormulaR1C1()
    PID_RestoreMonatslohnFormulasOnSheet wsMonth, formulaR1C1

SafeExit:
End Sub


Public Sub PID_RestoreMonatslohnFormulasSilent()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim formulaR1C1 As String
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo SafeExit
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    formulaR1C1 = PID_GetMonatslohnFormulaR1C1()
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo SafeExit
        
        If Not ws Is Nothing Then
            PID_RestoreMonatslohnFormulasOnSheet ws, formulaR1C1
        End If
    Next i
    
    ClearAllKVLohnDirty

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Private Function PID_RestoreMonatslohnFormulasOnSheet(ByVal ws As Worksheet, _
                                                      ByVal formulaR1C1 As String) As Boolean
    Dim targetRange As Range
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
    
    Set targetRange = ws.Range("G" & PID_FIRST_ROW & ":G" & PID_LAST_ROW)
    targetRange.FormulaR1C1 = formulaR1C1
    PID_ApplyEuroNumberFormat targetRange
    
    PID_RestoreMonatslohnFormulasOnSheet = True

SafeExit:
    On Error Resume Next
    If wasProtected Then
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
    End If
End Function


Public Sub MarkAllKVLohnDirty()
    gKVLohnAllMonthsDirty = True
    Set mKVLohnRefreshedSheets = New Collection
    PID_ClearLohnTableCache
End Sub


Public Sub MarkKVLohnDirty()
    MarkAllKVLohnDirty
End Sub


Public Sub ClearAllKVLohnDirty()
    gKVLohnAllMonthsDirty = False
    Set mKVLohnRefreshedSheets = New Collection
End Sub


Public Sub ClearKVLohnDirty()
    ClearAllKVLohnDirty
End Sub


Public Function IsKVLohnDirty() As Boolean
    IsKVLohnDirty = gKVLohnAllMonthsDirty
End Function


Public Sub PID_MarkKVLohnSheetRefreshed(ByVal sheetName As String)
    On Error Resume Next
    If sheetName = "" Then Exit Sub
    If mKVLohnRefreshedSheets Is Nothing Then Set mKVLohnRefreshedSheets = New Collection
    mKVLohnRefreshedSheets.Add sheetName, sheetName
End Sub


Public Sub PID_RecalculateMonatslohnForChangedRows(ByVal wsMonth As Worksheet, ByVal changedRange As Range)
    Dim rowsToCheck As Range
    Dim c As Range
    Dim checkedRows As Collection
    Dim rowKey As String
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    
    Set rowsToCheck = Intersect(changedRange, wsMonth.Range("E3:F82"))
    If rowsToCheck Is Nothing Then Exit Sub
    
    Set checkedRows = New Collection
    
    For Each c In rowsToCheck.Cells
        rowKey = CStr(c.Row)
        
        If c.Row >= PID_FIRST_ROW And c.Row <= PID_LAST_ROW Then
            If Not CollectionHasKey_KVLohn(checkedRows, rowKey) Then
                checkedRows.Add rowKey, rowKey
                PID_ForceMonatslohnRecalcForRow wsMonth, c.Row
            End If
        End If
    Next c

SafeExit:
End Sub


Public Sub PID_RecalculateMonatslohnForUsedRows(ByVal wsMonth As Worksheet)
    Dim r As Long
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    
    For r = PID_FIRST_ROW To PID_LAST_ROW
        If Len(Trim$(CStr(wsMonth.Cells(r, "E").Value))) > 0 Then
            If Len(Trim$(CStr(wsMonth.Cells(r, "F").Value))) > 0 Then
                PID_ForceMonatslohnRecalcForRow wsMonth, r
            End If
        End If
    Next r

SafeExit:
End Sub


Public Sub PID_ForceMonatslohnRecalcForRow(ByVal wsMonth As Worksheet, ByVal rowNumber As Long)
    Dim gCell As Range
    Dim formulaR1C1 As String
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Sub
    
    Set gCell = wsMonth.Cells(rowNumber, "G")
    
    If Not gCell.HasFormula Then
        If Not PID_MonthSheetHasMonatslohnFormula(wsMonth) Then
            PID_RestoreMonatslohnFormulasOnSheet wsMonth, PID_GetMonatslohnFormulaR1C1()
        End If
        Exit Sub
    End If
    
    formulaR1C1 = gCell.FormulaR1C1
    gCell.FormulaR1C1 = ""
    gCell.FormulaR1C1 = formulaR1C1

SafeExit:
End Sub


Public Sub RefreshKVLohnIfDirty(ByVal wsMonth As Worksheet)
    If wsMonth Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    
    If Not PID_MonthSheetHasMonatslohnFormula(wsMonth) Then
        PID_RestoreMonatslohnFormulasOnSheet wsMonth, PID_GetMonatslohnFormulaR1C1()
    End If
    
    If gKVLohnAllMonthsDirty Then
        If Not mKVLohnRefreshedSheets Is Nothing Then
            If CollectionHasKey_KVLohn(mKVLohnRefreshedSheets, wsMonth.Name) Then Exit Sub
        End If
        
        PID_RecalculateMonatslohnForUsedRows wsMonth
        PID_MarkKVLohnSheetRefreshed wsMonth.Name
        
        If PID_AllMonthSheetsKVLohnRefreshed() Then
            ClearAllKVLohnDirty
        End If
    End If
End Sub


Private Function PID_AllMonthSheetsKVLohnRefreshed() As Boolean
    Dim monthNames As Variant
    Dim i As Long
    
    If mKVLohnRefreshedSheets Is Nothing Then Exit Function
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        If Not CollectionHasKey_KVLohn(mKVLohnRefreshedSheets, CStr(monthNames(i))) Then
            Exit Function
        End If
    Next i
    
    PID_AllMonthSheetsKVLohnRefreshed = True
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


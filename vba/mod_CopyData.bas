Attribute VB_Name = "mod_CopyData"
Option Explicit

' Geschwindigkeit / Verhalten:
' True  = Jedes Monatsblatt springt am Ende des Makros auf A1 zurueck, langsamer, aber saubere Ansicht.
' False = Nur das urspruengliche Blatt kehrt zu A1 zurueck, schneller.
Private Const PID_RESET_ALL_MONTH_SELECTIONS As Boolean = False

' True  = CopyData setzt waehrend der Ausfuehrung die Geld-/Zahlenformate aller betroffenen Monatsblaetter neu, langsamer.
' False = Formatiert nicht alle Monate neu, schneller. Vorhandene Formate bleiben erhalten.
Private Const PID_APPLY_FORMATS_DURING_COPY As Boolean = False

' True  = Die monatliche Fluktuation in Q31 wird auch waehrend CopyData aktualisiert.
' False = Nur das Dirty-Flag bleibt gesetzt, Fluktuation wird spaeter aktualisiert.
Private Const PID_CALCULATE_FLUCTUATION_DURING_COPY As Boolean = False
Private Const PID_HOUR_OVERRIDE_LOG_SHEET As String = "PID_HOUR_OVERRIDES"


Public Sub CopyData()
    PID_CopyDataToFollowingMonths
End Sub


Public Sub DatenInFolgendeMonateKopieren()
    PID_CopyDataToFollowingMonths
End Sub


Public Sub PID_CopyDataToFollowingMonths()
    Dim wsSource As Worksheet
    Dim sourceSheetName As String
    Dim sourceMonthIndex As Long
    Dim workbookYear As Long
    Dim monthNames As Variant
    
    Dim sourceData As Variant
    Dim currentData As Variant
    
    Dim futureOverrides As Collection
    Dim futureNewStarts As Collection
    
    Dim formulaH As Variant
    Dim formulaK As Variant
    Dim formulaL As Variant
    Dim infoOQ As Variant
    
    Dim i As Long
    Dim targetSheetName As String
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant
    
    On Error GoTo CleanFail
    
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    
    Set wsSource = ActiveSheet
    sourceSheetName = wsSource.Name
    
    If Not PID_IsWorkerMonthSheet(wsSource) Then Exit Sub
    If Not IsNumeric(wsSource.Range("A1").Value) Then Exit Sub
    
    sourceMonthIndex = CLng(wsSource.Range("A1").Value)
    If sourceMonthIndex < 1 Or sourceMonthIndex > 12 Then Exit Sub
    
    workbookYear = PID_GetWorkbookYear()
    monthNames = PID_MonthNames()
    
    ' Self-heal: make sure workbook events are active before top-level copy run.
    Application.EnableEvents = True
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldStatusBar = Application.StatusBar
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Daten werden kopiert..."
    
    sourceData = PID_ReadMonthData(wsSource)
    currentData = sourceData
    
    formulaH = wsSource.Range("H3:H82").FormulaR1C1
    formulaK = wsSource.Range("K3:K82").FormulaR1C1
    formulaL = wsSource.Range("L3:L82").FormulaR1C1
    infoOQ = wsSource.Range("O18:Q25").FormulaR1C1
    
    Set futureOverrides = New Collection
    Set futureNewStarts = New Collection
    
    PID_CollectFutureOverrides sourceData, sourceMonthIndex, monthNames, futureOverrides, futureNewStarts
    PID_ApplyLoggedHourOverrides futureOverrides, sourceMonthIndex, workbookYear
    
    For i = sourceMonthIndex + 1 To 12
        targetSheetName = CStr(monthNames(i - 1))
        
        currentData = PID_BuildTargetMonthData(currentData, futureOverrides, futureNewStarts, workbookYear, i)
        
        PID_WriteMonthData targetSheetName, currentData, formulaH, formulaK, formulaL, infoOQ
    Next i
    
    MarkFluktuationDirty
    PID_HideUnwantedTechnicalSheets

CleanExit:
    Application.CutCopyMode = False
    Application.StatusBar = oldStatusBar
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    If PID_RESET_ALL_MONTH_SELECTIONS Then
        PID_ResetMonthSelections sourceSheetName
    Else
        PID_ReturnToSourceSheet sourceSheetName
    End If
    
    Exit Sub

CleanFail:
    On Error Resume Next
    
    PID_HideUnwantedTechnicalSheets
    
    Application.CutCopyMode = False
    Application.StatusBar = oldStatusBar
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    If PID_RESET_ALL_MONTH_SELECTIONS Then
        PID_ResetMonthSelections sourceSheetName
    Else
        PID_ReturnToSourceSheet sourceSheetName
    End If
    
    MsgBox "Fehler bei CopyData:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Daten kopieren"
End Sub


Private Sub PID_CollectFutureOverrides(ByVal sourceData As Variant, _
                                      ByVal sourceMonthIndex As Long, _
                                      ByVal monthNames As Variant, _
                                      ByRef futureOverrides As Collection, _
                                      ByRef futureNewStarts As Collection)
    Dim baseValues As Collection
    Dim currentValues As Collection
    
    Dim ws As Worksheet
    Dim monthIndex As Long
    Dim r As Long
    Dim keyText As String
    Dim targetData As Variant
    
    Set baseValues = New Collection
    Set currentValues = New Collection
    
    For r = 1 To UBound(sourceData, 1)
        keyText = PID_BuildEmployeeKey(sourceData(r, 1), sourceData(r, 2))
        
        If keyText <> "" Then
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "EXISTS"), True
            
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "B"), CStr(sourceData(r, 1))
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "C"), CStr(sourceData(r, 2))
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "D"), CStr(sourceData(r, 3))
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "E"), CStr(sourceData(r, 4))
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "F"), CStr(sourceData(r, 5))
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "I"), CStr(sourceData(r, 8))
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "J"), CStr(sourceData(r, 9))
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "M"), CStr(sourceData(r, 12))
            PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "N"), CStr(sourceData(r, 13))
            
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "B"), CStr(sourceData(r, 1))
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "C"), CStr(sourceData(r, 2))
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "D"), CStr(sourceData(r, 3))
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "E"), CStr(sourceData(r, 4))
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "F"), CStr(sourceData(r, 5))
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "I"), CStr(sourceData(r, 8))
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "J"), CStr(sourceData(r, 9))
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "M"), CStr(sourceData(r, 12))
            PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "N"), CStr(sourceData(r, 13))
        End If
    Next r
    
    For monthIndex = sourceMonthIndex + 1 To 12
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(monthIndex - 1)))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            targetData = PID_ReadMonthData(ws)
    
            For r = 1 To UBound(targetData, 1)
                keyText = PID_BuildEmployeeKey(targetData(r, 1), targetData(r, 2))
                
                If keyText <> "" Then
                    
                    If Not PID_CollectionHasKey(baseValues, PID_BaseKey(keyText, "EXISTS")) Then
                        
                        If Not PID_CollectionHasKey(futureNewStarts, keyText) Then
                            futureNewStarts.Add CStr(monthIndex) & PID_Sep() & keyText, keyText
                        End If
                        
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "EXISTS"), True
                        
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "B", targetData(r, 1)
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "C", targetData(r, 2)
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "D", targetData(r, 3)
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "E", targetData(r, 4)
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "F", targetData(r, 5)
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "I", targetData(r, 8)
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "J", targetData(r, 9)
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "M", targetData(r, 12)
                        PID_AddOverrideValue futureOverrides, monthIndex, keyText, "N", targetData(r, 13)
                        
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "B"), CStr(targetData(r, 1))
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "C"), CStr(targetData(r, 2))
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "D"), CStr(targetData(r, 3))
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "E"), CStr(targetData(r, 4))
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "F"), CStr(targetData(r, 5))
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "I"), CStr(targetData(r, 8))
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "J"), CStr(targetData(r, 9))
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "M"), CStr(targetData(r, 12))
                        PID_AddOrReplaceCollectionValue baseValues, PID_BaseKey(keyText, "N"), CStr(targetData(r, 13))
                        
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "B"), CStr(targetData(r, 1))
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "C"), CStr(targetData(r, 2))
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "D"), CStr(targetData(r, 3))
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "E"), CStr(targetData(r, 4))
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "F"), CStr(targetData(r, 5))
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "I"), CStr(targetData(r, 8))
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "J"), CStr(targetData(r, 9))
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "M"), CStr(targetData(r, 12))
                        PID_AddOrReplaceCollectionValue currentValues, PID_BaseKey(keyText, "N"), CStr(targetData(r, 13))
                        
                    Else
                        If Trim$(CStr(targetData(r, 8))) <> "" Then
                            PID_AddOverrideValue futureOverrides, monthIndex, keyText, "I", targetData(r, 8)
                        End If
                        
                        If Trim$(CStr(targetData(r, 9))) <> "" Then
                            PID_AddOverrideValue futureOverrides, monthIndex, keyText, "J", targetData(r, 9)
                        End If
                        
                        If Trim$(CStr(targetData(r, 12))) <> "" Then
                            PID_AddOverrideValue futureOverrides, monthIndex, keyText, "M", targetData(r, 12)
                        End If
                        
                        If Trim$(CStr(targetData(r, 13))) <> "" Then
                            PID_AddOverrideValue futureOverrides, monthIndex, keyText, "N", targetData(r, 13)
                        End If
                        
                    End If
                    
                End If
            Next r
        End If
    Next monthIndex
End Sub


Public Sub PID_LogEFAenderungForSheet(ByVal wsMonth As Worksheet, ByVal changedRange As Range)
    Dim rowsToCheck As Range
    Dim c As Range
    Dim keyText As String
    Dim fieldCode As String
    Dim monthIndex As Long
    Dim workbookYear As Long
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    If Not PID_IsWorkerMonthSheet(wsMonth) Then Exit Sub
    If Not IsNumeric(wsMonth.Range("A1").Value) Then Exit Sub
    
    monthIndex = CLng(wsMonth.Range("A1").Value)
    If monthIndex < 1 Or monthIndex > 12 Then Exit Sub
    
    Set rowsToCheck = Intersect(changedRange, wsMonth.Range("E3:F82"))
    If rowsToCheck Is Nothing Then Exit Sub
    
    workbookYear = PID_GetWorkbookYear()
    
    For Each c In rowsToCheck.Cells
        keyText = PID_BuildEmployeeKey(wsMonth.Cells(c.Row, "B").Value, wsMonth.Cells(c.Row, "C").Value)
        
        If keyText <> "" Then
            If c.Column = 5 Then
                fieldCode = "E"
            ElseIf c.Column = 6 Then
                fieldCode = "F"
            Else
                fieldCode = ""
            End If
            
            If fieldCode <> "" Then
                PID_UpsertHourOverride workbookYear, monthIndex, keyText, fieldCode, c.Value
            End If
        End If
    Next c
    
SafeExit:
End Sub


Private Sub PID_ApplyLoggedHourOverrides(ByRef futureOverrides As Collection, _
                                         ByVal sourceMonthIndex As Long, _
                                         ByVal workbookYear As Long)
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim data As Variant
    Dim r As Long
    Dim logYear As Long
    Dim logMonth As Long
    Dim employeeKey As String
    Dim fieldCode As String
    
    On Error GoTo SafeExit
    
    Set wsLog = Nothing
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(PID_HOUR_OVERRIDE_LOG_SHEET)
    On Error GoTo SafeExit
    
    If wsLog Is Nothing Then Exit Sub
    
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    data = wsLog.Range("A2:E" & lastRow).Value
    
    For r = 1 To UBound(data, 1)
        If IsNumeric(data(r, 1)) And IsNumeric(data(r, 2)) Then
            logYear = CLng(data(r, 1))
            logMonth = CLng(data(r, 2))
            
            If logYear = workbookYear Then
                If logMonth > sourceMonthIndex And logMonth <= 12 Then
                    employeeKey = Trim$(CStr(data(r, 3)))
                    fieldCode = Trim$(CStr(data(r, 4)))
                    
                    If employeeKey <> "" Then
                        If fieldCode = "E" Or fieldCode = "F" Then
                            If Trim$(CStr(data(r, 5))) <> "" Then
                                PID_AddOverrideValue futureOverrides, logMonth, employeeKey, fieldCode, data(r, 5)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r
    
SafeExit:
End Sub


Private Sub PID_UpsertHourOverride(ByVal workbookYear As Long, _
                                   ByVal monthIndex As Long, _
                                   ByVal employeeKey As String, _
                                   ByVal fieldCode As String, _
                                   ByVal valueToStore As Variant)
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    On Error GoTo SafeExit
    
    If employeeKey = "" Then Exit Sub
    If fieldCode <> "E" And fieldCode <> "F" Then Exit Sub
    
    Set wsLog = PID_GetOrCreateHourOverrideLogSheet()
    If wsLog Is Nothing Then Exit Sub
    
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 1
    
    For r = 2 To lastRow
        If CLng(Val(wsLog.Cells(r, "A").Value)) = workbookYear _
           And CLng(Val(wsLog.Cells(r, "B").Value)) = monthIndex _
           And UCase$(Trim$(CStr(wsLog.Cells(r, "C").Value))) = UCase$(employeeKey) _
           And UCase$(Trim$(CStr(wsLog.Cells(r, "D").Value))) = fieldCode Then
            
            If Trim$(CStr(valueToStore)) = "" Then
                wsLog.Rows(r).Delete
            Else
                wsLog.Cells(r, "E").Value = valueToStore
            End If
            GoTo SafeExit
        End If
    Next r
    
    If Trim$(CStr(valueToStore)) = "" Then Exit Sub
    
    lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Cells(lastRow, "A").Value = workbookYear
    wsLog.Cells(lastRow, "B").Value = monthIndex
    wsLog.Cells(lastRow, "C").Value = employeeKey
    wsLog.Cells(lastRow, "D").Value = fieldCode
    wsLog.Cells(lastRow, "E").Value = valueToStore
    
SafeExit:
End Sub


Private Function PID_GetOrCreateHourOverrideLogSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_HOUR_OVERRIDE_LOG_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PID_HOUR_OVERRIDE_LOG_SHEET
        
        ws.Range("A1").Value = "Year"
        ws.Range("B1").Value = "Month"
        ws.Range("C1").Value = "EmployeeKey"
        ws.Range("D1").Value = "Field"
        ws.Range("E1").Value = "Value"
    End If
    
    ws.Visible = xlSheetVeryHidden
    
    Set PID_GetOrCreateHourOverrideLogSheet = ws
End Function


Public Sub PID_ResetHourOverrideLog()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Set ws = PID_GetOrCreateHourOverrideLogSheet()
    If ws Is Nothing Then Exit Sub
    
    ws.Range("A2:E" & ws.Rows.Count).ClearContents
    ws.Visible = xlSheetVeryHidden
    
SafeExit:
End Sub


Private Function PID_BuildTargetMonthData(ByVal currentData As Variant, _
                                          ByVal futureOverrides As Collection, _
                                          ByVal futureNewStarts As Collection, _
                                          ByVal workbookYear As Long, _
                                          ByVal targetMonthIndex As Long) As Variant
    Dim resultData As Variant
    Dim knownKeys As Collection
    
    Dim r As Long
    Dim c As Long
    Dim resultRow As Long
    Dim keyText As String
    Dim exitDate As Variant
    
    Set knownKeys = New Collection
    
    ReDim resultData(1 To PID_LAST_ROW - PID_FIRST_ROW + 1, 1 To 13)
    resultRow = 0
    
    For r = 1 To UBound(currentData, 1)
        keyText = PID_BuildEmployeeKey(currentData(r, 1), currentData(r, 2))
        
        If keyText <> "" Then
            exitDate = currentData(r, 8)
            
            If PID_ShouldEmployeeExistInMonth(exitDate, workbookYear, targetMonthIndex) Then
                resultRow = resultRow + 1
                
                If resultRow <= UBound(resultData, 1) Then
                    For c = 1 To 13
                        resultData(resultRow, c) = currentData(r, c)
                    Next c
                    
                    PID_ApplyOverridesUntilMonth resultData, resultRow, futureOverrides, targetMonthIndex, keyText
                    
                    resultData(resultRow, 6) = ""
                    resultData(resultRow, 7) = ""
                    resultData(resultRow, 10) = ""
                    resultData(resultRow, 11) = ""
                    
                    If Not PID_CollectionHasKey(knownKeys, keyText) Then
                        knownKeys.Add True, keyText
                    End If
                End If
            End If
        End If
    Next r
    
    PID_AddFutureNewEmployees resultData, resultRow, knownKeys, futureNewStarts, futureOverrides, workbookYear, targetMonthIndex
    
    PID_BuildTargetMonthData = resultData
End Function


Private Sub PID_AddFutureNewEmployees(ByRef resultData As Variant, _
                                      ByRef resultRow As Long, _
                                      ByRef knownKeys As Collection, _
                                      ByVal futureNewStarts As Collection, _
                                      ByVal futureOverrides As Collection, _
                                      ByVal workbookYear As Long, _
                                      ByVal targetMonthIndex As Long)
    Dim item As Variant
    Dim parts As Variant
    Dim startMonth As Long
    Dim keyText As String
    Dim exitDate As Variant
    
    For Each item In futureNewStarts
        parts = Split(CStr(item), PID_Sep())
        
        If UBound(parts) >= 1 Then
            startMonth = CLng(parts(0))
            keyText = CStr(parts(1))
            
            If startMonth <= targetMonthIndex Then
                If Not PID_CollectionHasKey(knownKeys, keyText) Then
                    exitDate = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "I")
                    
                    If PID_ShouldEmployeeExistInMonth(exitDate, workbookYear, targetMonthIndex) Then
                        resultRow = resultRow + 1
                        
                        If resultRow <= UBound(resultData, 1) Then
                            resultData(resultRow, 1) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "B")
                            resultData(resultRow, 2) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "C")
                            resultData(resultRow, 3) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "D")
                            resultData(resultRow, 4) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "E")
                            resultData(resultRow, 5) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "F")
                            resultData(resultRow, 6) = ""
                            resultData(resultRow, 7) = ""
                            resultData(resultRow, 8) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "I")
                            resultData(resultRow, 9) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "J")
                            resultData(resultRow, 10) = ""
                            resultData(resultRow, 11) = ""
                            resultData(resultRow, 12) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "M")
                            resultData(resultRow, 13) = PID_GetOverrideValue(futureOverrides, startMonth, keyText, "N")
                            
                            PID_ApplyOverridesUntilMonth resultData, resultRow, futureOverrides, targetMonthIndex, keyText
                            
                            If Not PID_CollectionHasKey(knownKeys, keyText) Then
                                knownKeys.Add True, keyText
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next item
End Sub


Private Sub PID_ApplyOverridesUntilMonth(ByRef resultData As Variant, _
                                        ByVal resultRow As Long, _
                                        ByVal futureOverrides As Collection, _
                                        ByVal targetMonthIndex As Long, _
                                        ByVal keyText As String)
    Dim m As Long
    
    For m = 1 To targetMonthIndex
        If PID_HasOverrideValue(futureOverrides, m, keyText, "E") Then
            resultData(resultRow, 4) = PID_GetOverrideValue(futureOverrides, m, keyText, "E")
        End If
        
        If PID_HasOverrideValue(futureOverrides, m, keyText, "F") Then
            resultData(resultRow, 5) = PID_GetOverrideValue(futureOverrides, m, keyText, "F")
        End If
        
        If PID_HasOverrideValue(futureOverrides, m, keyText, "I") Then
            resultData(resultRow, 8) = PID_GetOverrideValue(futureOverrides, m, keyText, "I")
        End If
        
        If PID_HasOverrideValue(futureOverrides, m, keyText, "J") Then
            resultData(resultRow, 9) = PID_GetOverrideValue(futureOverrides, m, keyText, "J")
        End If
        
        If PID_HasOverrideValue(futureOverrides, m, keyText, "M") Then
            resultData(resultRow, 12) = PID_GetOverrideValue(futureOverrides, m, keyText, "M")
        End If
        
        If PID_HasOverrideValue(futureOverrides, m, keyText, "N") Then
            resultData(resultRow, 13) = PID_GetOverrideValue(futureOverrides, m, keyText, "N")
        End If
    Next m
End Sub


Private Sub PID_WriteMonthData(ByVal targetSheetName As String, _
                              ByVal dataToWrite As Variant, _
                              ByVal formulaH As Variant, _
                              ByVal formulaK As Variant, _
                              ByVal formulaL As Variant, _
                              ByVal infoOQ As Variant)
    Dim ws As Worksheet
    Dim arrBG As Variant
    Dim arrIJ As Variant
    Dim arrMN As Variant
    Dim r As Long
    
    On Error GoTo SafeExit
    
    Set ws = ThisWorkbook.Worksheets(targetSheetName)
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    ReDim arrBG(1 To PID_LAST_ROW - PID_FIRST_ROW + 1, 1 To 6)
    ReDim arrIJ(1 To PID_LAST_ROW - PID_FIRST_ROW + 1, 1 To 2)
    ReDim arrMN(1 To PID_LAST_ROW - PID_FIRST_ROW + 1, 1 To 2)
    
    For r = 1 To UBound(dataToWrite, 1)
        arrBG(r, 1) = dataToWrite(r, 1)
        arrBG(r, 2) = dataToWrite(r, 2)
        arrBG(r, 3) = dataToWrite(r, 3)
        arrBG(r, 4) = dataToWrite(r, 4)
        arrBG(r, 5) = dataToWrite(r, 5)
        arrBG(r, 6) = ""
        
        arrIJ(r, 1) = dataToWrite(r, 8)
        arrIJ(r, 2) = dataToWrite(r, 9)
        
        arrMN(r, 1) = dataToWrite(r, 12)
        arrMN(r, 2) = dataToWrite(r, 13)
    Next r
    
    ws.Range("B3:G82").Value = arrBG
    ws.Range("I3:J82").Value = arrIJ
    ws.Range("M3:N82").Value = arrMN
    
    PID_SortMonthSheet ws
    
    PID_RestoreFormulas ws, formulaH, formulaK, formulaL, infoOQ
    
    RefreshKVLohnForSheet ws
    
    If PID_CALCULATE_FLUCTUATION_DURING_COPY Then
        PID_CalculateFluctuation ws
    End If
    
    If PID_APPLY_FORMATS_DURING_COPY Then
        PID_ApplyMonthSheetFormats ws
    End If
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True

SafeExit:
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Function PID_ReadMonthData(ByVal ws As Worksheet) As Variant
    PID_ReadMonthData = ws.Range("B3:N82").Value
End Function


Private Function PID_BuildEmployeeKey(ByVal keyPart1 As Variant, ByVal keyPart2 As Variant) As String
    Dim s1 As String
    Dim s2 As String
    
    s1 = Trim$(CStr(keyPart1))
    s2 = Trim$(CStr(keyPart2))
    
    If s1 = "" And s2 = "" Then
        PID_BuildEmployeeKey = ""
    Else
        PID_BuildEmployeeKey = UCase$(s1 & "|" & s2)
    End If
End Function


Private Function PID_ShouldEmployeeExistInMonth(ByVal exitDate As Variant, _
                                                ByVal workbookYear As Long, _
                                                ByVal targetMonthIndex As Long) As Boolean
    Dim targetFirstDay As Date
    
    PID_ShouldEmployeeExistInMonth = True
    
    If Not IsDate(exitDate) Then Exit Function
    
    targetFirstDay = DateSerial(workbookYear, targetMonthIndex, 1)
    
    If CDate(exitDate) < targetFirstDay Then
        PID_ShouldEmployeeExistInMonth = False
    End If
End Function


Private Sub PID_RestoreFormulas(ByVal ws As Worksheet, _
                               ByVal formulaH As Variant, _
                               ByVal formulaK As Variant, _
                               ByVal formulaL As Variant, _
                               ByVal infoOQ As Variant)
    ws.Range("H" & PID_FIRST_ROW & ":H" & PID_LAST_ROW).FormulaR1C1 = formulaH
    ws.Range("K" & PID_FIRST_ROW & ":K" & PID_LAST_ROW).FormulaR1C1 = formulaK
    ws.Range("L" & PID_FIRST_ROW & ":L" & PID_LAST_ROW).FormulaR1C1 = formulaL
    ws.Range("O18:Q25").FormulaR1C1 = infoOQ
End Sub


Private Sub PID_SortMonthSheet(ByVal ws As Worksheet)
    Dim sortRange As Range
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    Set sortRange = ws.Range("B3:N82")
    
    With ws.Sort
        .SortFields.Clear
        
        .SortFields.Add key:=ws.Range("B3:B82"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        .SortFields.Add key:=ws.Range("C3:C82"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        .SetRange sortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

SafeExit:
End Sub


Private Sub PID_ApplyMonthSheetFormats(ByVal ws As Worksheet)
    Dim euroSymbol As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    euroSymbol = ChrW(8364)
    
    ws.Range("D3:D82").NumberFormat = "dd.mm.yyyy"
    ws.Range("I3:I82").NumberFormat = "dd.mm.yyyy"
    
    On Error GoTo TryEnglishFormat
    
    ws.Range("F3:F82").NumberFormatLocal = "0,00"
    ws.Range("H3:H82").NumberFormatLocal = "0,00"
    ws.Range("J3:J82").NumberFormatLocal = "0,00"
    
    ws.Range("G3:G82").NumberFormatLocal = euroSymbol & " #.##0,00"
    ws.Range("K3:K82").NumberFormatLocal = euroSymbol & " #.##0,00"
    ws.Range("L3:L82").NumberFormatLocal = euroSymbol & " #.##0,00"
    
    GoTo SafeExit

TryEnglishFormat:
    On Error Resume Next
    
    ws.Range("F3:F82").NumberFormat = "0.00"
    ws.Range("H3:H82").NumberFormat = "0.00"
    ws.Range("J3:J82").NumberFormat = "0.00"
    
    ws.Range("G3:G82").NumberFormat = euroSymbol & " #,##0.00"
    ws.Range("K3:K82").NumberFormat = euroSymbol & " #,##0.00"
    ws.Range("L3:L82").NumberFormat = euroSymbol & " #,##0.00"

SafeExit:
End Sub


Private Function PID_Sep() As String
    PID_Sep = Chr$(30)
End Function


Private Function PID_BaseKey(ByVal employeeKey As String, ByVal fieldCode As String) As String
    PID_BaseKey = employeeKey & PID_Sep() & fieldCode
End Function


Private Function PID_OverrideKey(ByVal monthIndex As Long, ByVal employeeKey As String, ByVal fieldCode As String) As String
    PID_OverrideKey = CStr(monthIndex) & PID_Sep() & employeeKey & PID_Sep() & fieldCode
End Function


Private Sub PID_AddOverrideValue(ByRef col As Collection, _
                                ByVal monthIndex As Long, _
                                ByVal employeeKey As String, _
                                ByVal fieldCode As String, _
                                ByVal valueToStore As Variant)
    PID_AddOrReplaceCollectionValue col, PID_OverrideKey(monthIndex, employeeKey, fieldCode), valueToStore
End Sub


Private Function PID_HasOverrideValue(ByVal col As Collection, _
                                      ByVal monthIndex As Long, _
                                      ByVal employeeKey As String, _
                                      ByVal fieldCode As String) As Boolean
    PID_HasOverrideValue = PID_CollectionHasKey(col, PID_OverrideKey(monthIndex, employeeKey, fieldCode))
End Function


Private Function PID_GetOverrideValue(ByVal col As Collection, _
                                      ByVal monthIndex As Long, _
                                      ByVal employeeKey As String, _
                                      ByVal fieldCode As String) As Variant
    PID_GetOverrideValue = PID_GetCollectionValue(col, PID_OverrideKey(monthIndex, employeeKey, fieldCode), "")
End Function


Private Sub PID_AddOrReplaceCollectionValue(ByRef col As Collection, ByVal keyText As String, ByVal valueToStore As Variant)
    On Error Resume Next
    col.Remove keyText
    On Error GoTo 0
    
    col.Add valueToStore, keyText
End Sub


Private Function PID_GetCollectionValue(ByVal col As Collection, ByVal keyText As String, ByVal defaultValue As Variant) As Variant
    On Error GoTo NotFound
    
    PID_GetCollectionValue = col.item(keyText)
    Exit Function

NotFound:
    PID_GetCollectionValue = defaultValue
End Function


Public Function PID_CollectionHasKey(ByVal col As Collection, ByVal key As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    tmp = col.item(key)
    
    PID_CollectionHasKey = True
    Exit Function

NotFound:
    PID_CollectionHasKey = False
End Function


Public Function PID_IsWorkerMonthSheet(ByVal ws As Worksheet) As Boolean
    Dim monthIndex As Long
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    
    monthIndex = PID_GetMonthIndexFromSheetName(ws.Name)
    
    If monthIndex < 1 Or monthIndex > 12 Then
        PID_IsWorkerMonthSheet = False
        Exit Function
    End If
    
    If IsNumeric(ws.Range("A1").Value) Then
        If CLng(ws.Range("A1").Value) = monthIndex Then
            PID_IsWorkerMonthSheet = True
        Else
            PID_IsWorkerMonthSheet = False
        End If
    Else
        PID_IsWorkerMonthSheet = True
    End If
    
    Exit Function

SafeExit:
    PID_IsWorkerMonthSheet = False
End Function


Public Function PID_GetMonthIndexFromSheetName(ByVal sheetName As String) As Long
    Select Case Trim$(CStr(sheetName))
        Case "Januar"
            PID_GetMonthIndexFromSheetName = 1
        Case "Februar"
            PID_GetMonthIndexFromSheetName = 2
        Case "Marz"
            PID_GetMonthIndexFromSheetName = 3
        Case "April"
            PID_GetMonthIndexFromSheetName = 4
        Case "Mai"
            PID_GetMonthIndexFromSheetName = 5
        Case "Juni"
            PID_GetMonthIndexFromSheetName = 6
        Case "Juli"
            PID_GetMonthIndexFromSheetName = 7
        Case "August"
            PID_GetMonthIndexFromSheetName = 8
        Case "September"
            PID_GetMonthIndexFromSheetName = 9
        Case "Oktober"
            PID_GetMonthIndexFromSheetName = 10
        Case "November"
            PID_GetMonthIndexFromSheetName = 11
        Case "Dezember"
            PID_GetMonthIndexFromSheetName = 12
        Case Else
            PID_GetMonthIndexFromSheetName = 0
    End Select
End Function


Public Function PID_MonthNames() As Variant
    PID_MonthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
End Function


Public Function PID_GetWorkbookYear() As Long
    Dim wsLohn As Worksheet
    
    On Error GoTo Fallback
    
    Set wsLohn = ThisWorkbook.Worksheets("LOHNTABELLE")
    
    If IsNumeric(wsLohn.Range("G3").Value) Then
        PID_GetWorkbookYear = CLng(wsLohn.Range("G3").Value)
        Exit Function
    End If

Fallback:
    PID_GetWorkbookYear = Year(Date)
End Function


Public Sub PID_ShowEmploymentDuration(ByVal ws As Worksheet, ByVal targetCell As Range)
    Dim entryDate As Variant
    Dim diffYears As Long
    Dim diffMonths As Long
    Dim totalMonths As Long
    Dim todayDate As Date
    Dim messageText As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If targetCell Is Nothing Then Exit Sub
    
    If targetCell.CountLarge > 1 Then Exit Sub
    If Intersect(targetCell, ws.Range("D3:D82")) Is Nothing Then Exit Sub
    
    entryDate = targetCell.Value
    
    If Not IsDate(entryDate) Then
        ws.Range("O45").Value = ""
        Exit Sub
    End If
    
    todayDate = Date
    totalMonths = DateDiff("m", CDate(entryDate), todayDate)
    
    If Day(todayDate) < Day(CDate(entryDate)) Then
        totalMonths = totalMonths - 1
    End If
    
    If totalMonths < 0 Then totalMonths = 0
    
    diffYears = totalMonths \ 12
    diffMonths = totalMonths Mod 12
    
    messageText = "Betriebszugehoerigkeit: "
    
    If diffYears > 0 Then
        messageText = messageText & diffYears & " Jahr"
        If diffYears <> 1 Then messageText = messageText & "e"
    End If
    
    If diffMonths > 0 Then
        If diffYears > 0 Then messageText = messageText & " und "
        messageText = messageText & diffMonths & " Monat"
        If diffMonths <> 1 Then messageText = messageText & "e"
    End If
    
    If diffYears = 0 And diffMonths = 0 Then
        messageText = messageText & "unter 1 Monat"
    End If
    
    ws.Range("O45").Value = messageText

SafeExit:
End Sub


Private Sub PID_HideUnwantedTechnicalSheets()
    On Error Resume Next
    
    ThisWorkbook.Worksheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("KV_DROPDOWN_HELPER").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets(PID_HOUR_OVERRIDE_LOG_SHEET).Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("Settings").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("Message").Visible = xlSheetVeryHidden
    
    On Error GoTo 0
End Sub


Private Sub PID_ReturnToSourceSheet(ByVal sourceSheetName As String)
    On Error Resume Next
    
    If sourceSheetName <> "" Then
        ThisWorkbook.Worksheets(sourceSheetName).Activate
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ThisWorkbook.Worksheets(sourceSheetName).Range("A1").Select
    End If
    
    Application.CutCopyMode = False
    
    On Error GoTo 0
End Sub


Private Sub PID_ResetMonthSelections(ByVal sourceSheetName As String)
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    
    On Error Resume Next
    
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        
        If Not ws Is Nothing Then
            If ws.Visible = xlSheetVisible Then
                ws.Activate
                ActiveWindow.ScrollRow = 1
                ActiveWindow.ScrollColumn = 1
                ws.Range("A1").Select
            End If
        End If
    Next i
    
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    PID_ReturnToSourceSheet sourceSheetName
    
    On Error GoTo 0
End Sub


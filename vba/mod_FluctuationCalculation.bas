Attribute VB_Name = "mod_FluctuationCalculation"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Private Const PID_FLUKTUATION_PERCENT_FORMAT As String = "0.00%"


Public Sub PID_CalculateFluctuation(ByVal ws As Worksheet)
    Dim monthNumber As Long
    Dim currentYear As Long
    Dim monthEndDate As Date
    
    Dim personalEnde As Long
    Dim austritte As Long
    Dim fluctuation As Double
    
    Dim r As Long
    Dim employeeID As Variant
    Dim employeeName As Variant
    Dim entryDate As Variant
    Dim exitDate As Variant
    
    Dim arrID As Variant
    Dim arrName As Variant
    Dim arrEntry As Variant
    Dim arrExit As Variant
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Sub
    
    monthNumber = CLng(ws.Range("A1").Value)
    currentYear = PID_GetWorkbookYear()
    
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    monthEndDate = DateSerial(currentYear, monthNumber + 1, 0)
    
    arrID = ws.Range("B3:B82").Value
    arrName = ws.Range("C3:C82").Value
    arrEntry = ws.Range("D3:D82").Value
    arrExit = ws.Range("I3:I82").Value
    
    austritte = 0
    personalEnde = 0
    
    For r = 1 To 80
        employeeID = arrID(r, 1)
        employeeName = arrName(r, 1)
        entryDate = arrEntry(r, 1)
        exitDate = arrExit(r, 1)
        
        If IsDate(exitDate) Then
            If Year(CDate(exitDate)) = currentYear Then
                If Month(CDate(exitDate)) = monthNumber Then
                    If PID_FluctuationRowHasEmployee(employeeID, employeeName) Then
                        austritte = austritte + 1
                    End If
                End If
            End If
        End If
        
        If PID_FluctuationRowHasEmployee(employeeID, employeeName) Then
            
            If IsDate(entryDate) Then
                If CDate(entryDate) <= monthEndDate Then
                    
                    If Not IsDate(exitDate) Then
                        personalEnde = personalEnde + 1
                    ElseIf CDate(exitDate) > monthEndDate Then
                        personalEnde = personalEnde + 1
                    End If
                    
                End If
            Else
                
                If Not IsDate(exitDate) Then
                    personalEnde = personalEnde + 1
                ElseIf CDate(exitDate) > monthEndDate Then
                    personalEnde = personalEnde + 1
                End If
                
            End If
            
        End If
    Next r
    
    If personalEnde > 0 Then
        fluctuation = austritte / personalEnde
    Else
        fluctuation = 0
    End If
    
    ws.Range("Q31").Value = fluctuation
    ws.Range("Q31").NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT

SafeExit:
End Sub


Public Sub PID_FillFluktuationRates(ByRef monthExit() As Long, _
                                    ByRef monthPersonalEnde() As Long, _
                                    ByRef monthFluctuation() As Double, _
                                    ByRef quarterFluctuation() As Double, _
                                    ByRef ytdFluctuation As Double, _
                                    ByVal currentYear As Long)
    Dim i As Long
    Dim ytdMonthLimit As Long
    Dim ytdExits As Long
    Dim ytdPersonalSum As Long
    Dim ytdPersonalMonths As Long
    Dim monthNames As Variant
    
    monthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    For i = 1 To 12
        monthPersonalEnde(i) = GetPersonalEndeForMonth(CStr(monthNames(i - 1)), currentYear, i)
        
        If monthPersonalEnde(i) > 0 Then
            monthFluctuation(i) = monthExit(i) / monthPersonalEnde(i)
        Else
            monthFluctuation(i) = 0
        End If
    Next i
    
    quarterFluctuation(1) = PID_ComputeFluktuationForMonthRange(monthExit, monthPersonalEnde, 1, 3)
    quarterFluctuation(2) = PID_ComputeFluktuationForMonthRange(monthExit, monthPersonalEnde, 4, 6)
    quarterFluctuation(3) = PID_ComputeFluktuationForMonthRange(monthExit, monthPersonalEnde, 7, 9)
    quarterFluctuation(4) = PID_ComputeFluktuationForMonthRange(monthExit, monthPersonalEnde, 10, 12)
    
    If currentYear = Year(Date) Then
        ytdMonthLimit = Month(Date)
    ElseIf currentYear < Year(Date) Then
        ytdMonthLimit = 12
    Else
        ytdMonthLimit = GetLastMonthWithExit(monthExit)
    End If
    
    ytdExits = 0
    ytdPersonalSum = 0
    ytdPersonalMonths = 0
    
    If ytdMonthLimit > 0 Then
        For i = 1 To ytdMonthLimit
            ytdExits = ytdExits + monthExit(i)
            
            If monthPersonalEnde(i) > 0 Then
                ytdPersonalSum = ytdPersonalSum + monthPersonalEnde(i)
                ytdPersonalMonths = ytdPersonalMonths + 1
            End If
        Next i
    End If
    
    If ytdPersonalMonths > 0 Then
        ytdFluctuation = ytdExits / (ytdPersonalSum / ytdPersonalMonths)
    Else
        ytdFluctuation = 0
    End If
End Sub


Public Sub PID_SyncFluktuationToDisplaySheets(ByRef monthFluctuation() As Double, _
                                              ByRef quarterFluctuation() As Double, _
                                              ByVal ytdFluctuation As Double)
    PID_SyncFluktuationToUbersicht monthFluctuation, quarterFluctuation, ytdFluctuation
    PID_SyncFluktuationToMonthSheets monthFluctuation
End Sub


Public Sub PID_SyncFluktuationToUbersicht(ByRef monthFluctuation() As Double, _
                                          ByRef quarterFluctuation() As Double, _
                                          ByVal ytdFluctuation As Double)
    Dim ws As Worksheet
    Dim monthRows As Variant
    Dim quarterRows As Variant
    Dim i As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set ws = ThisWorkbook.Worksheets("UBERSICHT")
    
    wasProtected = ws.ProtectContents
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
    On Error GoTo SafeExit
    
    monthRows = Array(7, 8, 9, 11, 12, 13, 15, 16, 17, 19, 20, 21)
    quarterRows = Array(10, 14, 18, 22)
    
    For i = 1 To 12
        ws.Cells(CLng(monthRows(i - 1)), 17).ClearContents
        ws.Cells(CLng(monthRows(i - 1)), 17).Value2 = monthFluctuation(i)
        ws.Cells(CLng(monthRows(i - 1)), 17).NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT
    Next i
    
    For i = 1 To 4
        ws.Cells(CLng(quarterRows(i - 1)), 17).ClearContents
        ws.Cells(CLng(quarterRows(i - 1)), 17).Value2 = quarterFluctuation(i)
        ws.Cells(CLng(quarterRows(i - 1)), 17).NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT
    Next i
    
    ws.Cells(23, 17).ClearContents
    ws.Cells(23, 17).Value2 = ytdFluctuation
    ws.Cells(23, 17).NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT
    
    If wasProtected Then
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
    End If

SafeExit:
End Sub


Public Sub PID_SyncFluktuationToMonthSheets(ByRef monthFluctuation() As Double)
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    
    monthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    For i = 1 To 12
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i - 1)))
        
        If Not ws Is Nothing Then
            ws.Range("Q31").Value2 = monthFluctuation(i)
            ws.Range("Q31").NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT
        End If
        
        On Error GoTo 0
    Next i
End Sub


Private Function PID_ComputeFluktuationForMonthRange(ByRef monthExit() As Long, _
                                                     ByRef monthPersonalEnde() As Long, _
                                                     ByVal startMonth As Long, _
                                                     ByVal endMonth As Long) As Double
    Dim i As Long
    Dim exits As Long
    Dim personalSum As Long
    Dim personalMonths As Long
    
    For i = startMonth To endMonth
        exits = exits + monthExit(i)
        
        If monthPersonalEnde(i) > 0 Then
            personalSum = personalSum + monthPersonalEnde(i)
            personalMonths = personalMonths + 1
        End If
    Next i
    
    If personalMonths > 0 Then
        PID_ComputeFluktuationForMonthRange = exits / (personalSum / personalMonths)
    Else
        PID_ComputeFluktuationForMonthRange = 0
    End If
End Function


Private Function PID_FluctuationRowHasEmployee(ByVal employeeID As Variant, ByVal employeeName As Variant) As Boolean
    PID_FluctuationRowHasEmployee = (Len(Trim$(CStr(employeeID))) > 0 Or Len(Trim$(CStr(employeeName))) > 0)
End Function

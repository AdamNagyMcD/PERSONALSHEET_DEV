Attribute VB_Name = "mod_FluctuationCalculation"
Option Explicit

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
                    austritte = austritte + 1
                End If
            End If
        End If
        
        If Trim$(CStr(employeeID)) <> "" Or Trim$(CStr(employeeName)) <> "" Then
            
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
    ws.Range("Q31").NumberFormat = "0.00%"

SafeExit:
End Sub

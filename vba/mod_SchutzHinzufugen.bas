Attribute VB_Name = "mod_SchutzHinzufugen"
Option Explicit


Public Sub PID_SetupSheetProtectionForMacros()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        
        On Error Resume Next
        ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
        On Error GoTo SafeExit
        
        If PID_IsWorkerMonthSheetSafe(ws) Then
            
            ws.Range("E" & PID_FIRST_ROW & ":E" & PID_LAST_ROW).Locked = False
            ws.Range("F" & PID_FIRST_ROW & ":F" & PID_LAST_ROW).Locked = False
            
            ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                       UserInterfaceOnly:=True, _
                       AllowFiltering:=True, _
                       AllowSorting:=True
            
        ElseIf ws.Name = PID_LOHNTABELLE_SHEET Then
            
            ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                       UserInterfaceOnly:=True, _
                       AllowFiltering:=True, _
                       AllowSorting:=True
            
        ElseIf ws.Name = "FLUKTUATION_DATEN" Then
            
            ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                       UserInterfaceOnly:=True
            
            ws.Visible = xlSheetVeryHidden
            
        ElseIf ws.Name = "KV_DROPDOWN_HELPER" Then
            
            ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                       UserInterfaceOnly:=True
            
            ws.Visible = xlSheetVeryHidden
            
        Else
            
            ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                       UserInterfaceOnly:=True, _
                       AllowFiltering:=True, _
                       AllowSorting:=True
            
        End If
        
    Next ws

SafeExit:
    Application.ScreenUpdating = True
End Sub


Public Sub SchutzHinzufugen()
    PID_SetupSheetProtectionForMacros
    
    MsgBox "Blattschutz wurde eingerichtet." & vbCrLf & _
           "Makros duerfen geschuetzte Blaetter bearbeiten.", _
           vbInformation, "Blattschutz"
End Sub


Public Sub ProtectEverything()
    PID_SetupSheetProtectionForMacros
    
    MsgBox "Alle Blaetter wurden geschuetzt.", _
           vbInformation, "Blattschutz"
End Sub


Public Sub PID_ProtectWorkbookSheets()
    PID_SetupSheetProtectionForMacros
End Sub


Private Function PID_IsWorkerMonthSheetSafe(ByVal ws As Worksheet) As Boolean
    Dim monthName As String
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    
    monthName = Trim$(CStr(ws.Name))
    
    Select Case monthName
        Case "Januar", "Februar", "Marz", "April", "Mai", "Juni", _
             "Juli", "August", "September", "Oktober", "November", "Dezember"
            
            PID_IsWorkerMonthSheetSafe = True
        
        Case Else
            PID_IsWorkerMonthSheetSafe = False
    End Select
    
    Exit Function

SafeExit:
    PID_IsWorkerMonthSheetSafe = False
End Function


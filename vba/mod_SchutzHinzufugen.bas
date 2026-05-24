Attribute VB_Name = "mod_SchutzHinzufugen"
Option Explicit

Private mSessionProtectedSheets As Collection


Public Sub PID_SetupSheetProtectionForMacros()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Application.ScreenUpdating = False
    Set mSessionProtectedSheets = New Collection
    
    For Each ws In ThisWorkbook.Worksheets
        PID_ApplySheetProtectionForMacros ws
        PID_MarkSheetProtectionReady ws.Name
    Next ws

SafeExit:
    Application.ScreenUpdating = True
End Sub


Public Sub PID_SetupSheetProtectionForOpen()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Application.ScreenUpdating = False
    Set mSessionProtectedSheets = New Collection
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "FLUKTUATION_DATEN" Or ws.Name = "KV_DROPDOWN_HELPER" Then
            PID_ApplySheetProtectionForMacros ws
            PID_MarkSheetProtectionReady ws.Name
        End If
    Next ws
    
    If Not ActiveSheet Is Nothing Then
        PID_EnsureSheetProtectionForMacros ActiveSheet
    End If

SafeExit:
    Application.ScreenUpdating = True
End Sub


Public Sub PID_EnsureSheetProtectionForMacros(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If PID_SheetProtectionReady(ws.Name) Then Exit Sub
    
    PID_ApplySheetProtectionForMacros ws
    PID_MarkSheetProtectionReady ws.Name

SafeExit:
End Sub


Private Sub PID_ApplySheetProtectionForMacros(ByVal ws As Worksheet)
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
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

SafeExit:
End Sub


Private Function PID_SheetProtectionReady(ByVal sheetName As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    If mSessionProtectedSheets Is Nothing Then Exit Function
    If sheetName = "" Then Exit Function
    
    tmp = mSessionProtectedSheets.item(sheetName)
    
    PID_SheetProtectionReady = True
    Exit Function

NotFound:
    PID_SheetProtectionReady = False
End Function


Private Sub PID_MarkSheetProtectionReady(ByVal sheetName As String)
    On Error Resume Next
    
    If sheetName = "" Then Exit Sub
    If mSessionProtectedSheets Is Nothing Then Set mSessionProtectedSheets = New Collection
    
    mSessionProtectedSheets.Add sheetName, sheetName
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

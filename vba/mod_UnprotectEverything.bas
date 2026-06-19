Attribute VB_Name = "mod_UnprotectEverything"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit


Public Sub UnprotectEverything()
    Dim ws As Worksheet
    
    On Error GoTo CleanFail
    
    If Not PID_ConfirmAdminAction( _
        "Alle Blatt-Schutz und Workbook-Schutz werden aufgehoben. Alle versteckten " & PID_UTxtBlaetter() & " werden sichtbar.", _
        "Schutz aufheben") Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Blattschutz wird aufgehoben..."
    
    For Each ws In ThisWorkbook.Worksheets
        
        On Error Resume Next
        
        ws.Visible = xlSheetVisible
        ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
        
        On Error GoTo CleanFail
        
    Next ws
    
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    Application.CutCopyMode = False

CleanExit:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Alle " & PID_UTxtBlaetter() & " wurden entsperrt und sichtbar gemacht." & vbCrLf & _
           "Excel-Events und ScreenUpdating sind wieder aktiv.", _
           vbInformation, "Schutz aufgehoben"
    
    Exit Sub

CleanFail:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Fehler bei UnprotectEverything:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Schutz aufheben"
End Sub


Public Sub ShowAllSheets()
    Dim ws As Worksheet
    
    On Error GoTo CleanFail
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Visible = xlSheetVisible
        On Error GoTo CleanFail
    Next ws

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Alle " & PID_UTxtBlaetter() & " wurden sichtbar gemacht.", _
           vbInformation, PID_UTxtBlaetter() & " sichtbar"
    
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Fehler bei ShowAllSheets:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, PID_UTxtBlaetter() & " sichtbar"
End Sub

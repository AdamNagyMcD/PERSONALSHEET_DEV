Attribute VB_Name = "mod_UnprotectEverything"
Option Explicit


Public Sub UnprotectEverything()
    Dim ws As Worksheet
    
    On Error GoTo CleanFail
    
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
    
    MsgBox "Alle Blaetter wurden entsperrt und sichtbar gemacht." & vbCrLf & _
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


Public Sub PID_UnprotectWorkbookSheets()
    UnprotectEverything
End Sub


Public Sub UnlockEverything()
    UnprotectEverything
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
    
    MsgBox "Alle Blaetter wurden sichtbar gemacht.", _
           vbInformation, "Blaetter sichtbar"
    
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Fehler bei ShowAllSheets:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Blaetter sichtbar"
End Sub


Public Sub ResetExcelApplicationState()
    On Error Resume Next
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    
    On Error GoTo 0
    
    MsgBox "Excel wurde zurueckgesetzt:" & vbCrLf & _
           "- Events aktiv" & vbCrLf & _
           "- ScreenUpdating aktiv" & vbCrLf & _
           "- DisplayAlerts aktiv" & vbCrLf & _
           "- StatusBar zurueckgesetzt" & vbCrLf & _
           "- Calculation automatisch", _
           vbInformation, "Excel Reset"
End Sub


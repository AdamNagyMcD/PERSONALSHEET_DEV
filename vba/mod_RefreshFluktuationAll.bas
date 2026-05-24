Attribute VB_Name = "mod_RefreshFluktuationAll"
Option Explicit

Public gFluktuationDirty As Boolean


Public Sub MarkFluktuationDirty()
    gFluktuationDirty = True
End Sub


Public Sub MarkFluktuationClean()
    gFluktuationDirty = False
End Sub


Public Function IsFluktuationDirty() As Boolean
    IsFluktuationDirty = gFluktuationDirty
End Function


Public Sub RefreshFluktuationIfDirty()
    If gFluktuationDirty Then
        RefreshFluktuationAll
    End If
End Sub


Public Sub RefreshFluktuationNow()
    RefreshFluktuationAll
End Sub


Public Sub RefreshFluktuationAll()
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Fluktuation wird aktualisiert..."
    
    BuildFluktuationDaten
    BuildFluktuationAnalyse
    
    On Error Resume Next
    ThisWorkbook.Worksheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    On Error GoTo CleanFail
    
    MarkFluktuationClean

CleanExit:
    Application.StatusBar = oldStatusBar
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    gFluktuationDirty = True
    
    MsgBox "Die Fluktuation konnte nicht vollstaendig aktualisiert werden." & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Fluktuation aktualisieren"
    
    Resume CleanExit
End Sub


Attribute VB_Name = "Modul2"
Option Explicit

Public Sub FormatAllLohnColumns()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim i As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    monthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    For i = LBound(monthNames) To UBound(monthNames)
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            On Error Resume Next
            ws.Unprotect Password:="company"
            On Error GoTo CleanFail
            
            ' G = Lohn
            ApplyEuroNumberFormat ws.Range("G3:G82")
            
            ' J = Urlaub in Euro, ha nalad ez is penz oszlop
            ApplyEuroNumberFormat ws.Range("J3:J82")
            
            ' K = Letztes Gehalt, ha nalad ez is penz oszlop
            ApplyEuroNumberFormat ws.Range("K3:K82")
            
            ' Ne legyen ####### a keskeny oszlop miatt
            ws.Columns("G").ColumnWidth = 13
            ws.Columns("J").ColumnWidth = 13
            ws.Columns("K").ColumnWidth = 14
            
            ws.Protect Password:="company", UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
        End If
    Next i
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    On Error GoTo CleanFail
    
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Unprotect Password:="company"
        On Error GoTo CleanFail
        
        ApplyEuroNumberFormat ws.Range("H:H")
        ws.Columns("H").ColumnWidth = 14
        
        ws.Protect Password:="company", UserInterfaceOnly:=True
    End If
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOHNTABELLE")
    On Error GoTo CleanFail
    
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Unprotect Password:="company"
        On Error GoTo CleanFail
        
        ApplyEuroNumberFormat ws.Range("H:H")
        ws.Columns("H").ColumnWidth = 14
        
        ws.Protect Password:="company", UserInterfaceOnly:=True
    End If

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

CleanFail:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    
    MsgBox "Fehler bei FormatAllLohnColumns:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Lohn Formatierung"
End Sub


Private Sub ApplyEuroNumberFormat(ByVal targetRange As Range)
    Dim euroSymbol As String
    
    euroSymbol = ChrW(8364)
    
    On Error GoTo TryEnglishFormat
    
    ' Nemet / osztrak Excel
    targetRange.NumberFormatLocal = euroSymbol & " #.##0,00"
    Exit Sub

TryEnglishFormat:
    On Error GoTo SafeExit
    
    ' Fallback mas nyelvu Excelre
    targetRange.NumberFormat = euroSymbol & " #,##0.00"

SafeExit:
End Sub


Attribute VB_Name = "mod_PIDEingabePruefung"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Plausibilitaetspruefungen bei der Eingabe auf den Monatsblaettern.
'
' Hier steht nur, was der Benutzer beim Tippen sofort erfahren soll. Die Pruefung
' korrigiert nichts und lehnt nichts ab - sie weist hin. Abgelehnt wird nur dort,
' wo ein falscher Wert stillen Datenschaden anrichtet (doppelte Personal-ID,
' siehe mod_PersonalIdUnique).
'
' Austrittsdatum vor Eintrittsdatum:
' Spalte H (Aktuelle Stunden) rechnet mit (Austritt - Eintritt + 1). Liegt der
' Austritt vor dem Eintritt, entstehen negative Stunden, die unbemerkt in
' "Gesamt Crew Stunden" und in die Fluktuation einfliessen. Ein Tippfehler im Jahr
' (2025 statt 2026) reicht dafuer aus.

' Kein Const mit ChrW: eine Konstante muss zur Kompilierzeit feststehen.
Private Function PID_EPTitle() As String
    PID_EPTitle = "Datum pr" & PID_UTxtUe() & "fen"
End Function


' Wird aus Workbook_SheetChange gerufen, wenn sich D (Eintritt) oder I (Austritt)
' geaendert hat. Meldet pro Aenderung hoechstens einmal.
Public Sub PID_CheckEintrittAustrittPlausibel(ByVal wsMonth As Worksheet, ByVal changedRange As Range)
    Dim watchRange As Range
    Dim c As Range
    Dim checkedRows As Collection
    Dim rowKey As String
    Dim badRows As String
    Dim badCount As Long
    
    On Error GoTo SafeExit
    
    If wsMonth Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    
    Set watchRange = Intersect(changedRange, _
        wsMonth.Range("D" & PID_FIRST_ROW & ":D" & PID_LAST_ROW & "," & _
                      "I" & PID_FIRST_ROW & ":I" & PID_LAST_ROW))
    If watchRange Is Nothing Then Exit Sub
    
    Set checkedRows = New Collection
    
    For Each c In watchRange.Cells
        rowKey = CStr(c.Row)
        
        If Not PID_CollectionHasKey(checkedRows, rowKey) Then
            checkedRows.Add rowKey, rowKey
            
            If PID_EPRowHasImpossibleDates(wsMonth, c.Row) Then
                badCount = badCount + 1
                If badCount <= 8 Then
                    badRows = badRows & "- Zeile " & c.Row & ": " & _
                              PID_EPFormatDate(wsMonth.Cells(c.Row, "D").Value) & " bis " & _
                              PID_EPFormatDate(wsMonth.Cells(c.Row, "I").Value) & vbCrLf
                End If
            End If
        End If
    Next c
    
    If badCount = 0 Then Exit Sub
    
    If badCount > 8 Then
        badRows = badRows & "- ... (" & (badCount - 8) & " weitere)" & vbCrLf
    End If
    
    PID_TrackAction "Datum unplausibel", wsMonth.Name & ": " & badCount & " Zeile(n)"
    
    MsgBox "Das Austrittsdatum liegt vor dem Eintrittsdatum:" & vbCrLf & vbCrLf & _
           badRows & vbCrLf & _
           "Bitte pr" & PID_UTxtUe() & "fen - sonst rechnet die Spalte " & _
           """Aktuelle Stunden"" mit einer negativen Zahl." & vbCrLf & vbCrLf & _
           "Der Wert bleibt stehen, es wurde nichts " & PID_UTxtGeaendert() & ".", _
           vbExclamation, PID_EPTitle()

SafeExit:
End Sub


Private Function PID_EPRowHasImpossibleDates(ByVal wsMonth As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim entryValue As Variant
    Dim exitValue As Variant
    
    On Error GoTo SafeExit
    
    If rowNumber < PID_FIRST_ROW Or rowNumber > PID_LAST_ROW Then Exit Function
    
    entryValue = wsMonth.Cells(rowNumber, "D").Value
    exitValue = wsMonth.Cells(rowNumber, "I").Value
    
    If Not IsDate(entryValue) Then Exit Function
    If Not IsDate(exitValue) Then Exit Function
    
    PID_EPRowHasImpossibleDates = (CDate(exitValue) < CDate(entryValue))

SafeExit:
End Function


Private Function PID_EPFormatDate(ByVal dateValue As Variant) As String
    On Error GoTo SafeExit
    
    If IsDate(dateValue) Then
        PID_EPFormatDate = Format$(CDate(dateValue), "dd.mm.yyyy")
    Else
        PID_EPFormatDate = "(leer)"
    End If
    
    Exit Function

SafeExit:
    PID_EPFormatDate = "(?)"
End Function

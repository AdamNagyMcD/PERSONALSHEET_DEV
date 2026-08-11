Attribute VB_Name = "mod_PIDActionLog"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Kleines Aktionsprotokoll fuer die Testphase.
'
' Zweck: Wenn ein Tester "es hat etwas geloescht" meldet, ist ohne Protokoll nicht
' rekonstruierbar, welches Makro wann gelaufen ist. Die Fehlermeldung (mod_PIDFeedback)
' haengt die letzten Eintraege automatisch an.
'
' Bewusst klein gehalten: nur Zeitpunkt, Aktion, Blatt, Detail; rollierend begrenzt.
' Das Protokoll darf NIE einen Aufrufer stoeren - jeder Fehler wird verschluckt.
'
' Nicht protokolliert wird CopyData: mod_CopyData ist ein Bootstrap-Modul und wird
' ohne ausdrueckliche Freigabe nicht angefasst.

Private Const PID_AL_SHEET As String = "PID_ACTION_LOG"
Private Const PID_AL_FIRST_ROW As Long = 2
Private Const PID_AL_MAX_ROWS As Long = 500
Private Const PID_AL_REPORT_COUNT As Long = 15


Public Sub PID_TrackAction(ByVal actionName As String, ByVal detailText As String)
    Dim wsLog As Worksheet
    Dim targetRow As Long
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo SafeExit

    If Trim$(actionName) = "" Then Exit Sub

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set wsLog = PID_ALGetOrCreateSheet()
    If wsLog Is Nothing Then GoTo SafeExit

    targetRow = PID_ALGetLastRow(wsLog) + 1
    If targetRow < PID_AL_FIRST_ROW Then targetRow = PID_AL_FIRST_ROW

    ' Zeitpunkt als Text: kein Datumsformat, das je nach Systemeinstellung anders aussieht.
    wsLog.Cells(targetRow, "A").Value = Format$(Now, "dd.mm.yyyy hh:nn:ss")
    wsLog.Cells(targetRow, "B").Value = actionName
    wsLog.Cells(targetRow, "C").Value = PID_ALActiveSheetName()
    wsLog.Cells(targetRow, "D").Value = detailText

    PID_ALTrimToMaxRows wsLog

SafeExit:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


' Letzte Eintraege als fertiger Textblock (fuer die Fehlermeldung).
Public Function PID_GetLastActionsText(ByVal maxCount As Long) As String
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim firstRow As Long
    Dim r As Long
    Dim resultText As String
    Dim countWanted As Long

    On Error GoTo SafeExit

    countWanted = maxCount
    If countWanted <= 0 Then countWanted = PID_AL_REPORT_COUNT

    Set wsLog = Nothing
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(PID_AL_SHEET)
    On Error GoTo SafeExit

    If wsLog Is Nothing Then Exit Function

    lastRow = PID_ALGetLastRow(wsLog)
    If lastRow < PID_AL_FIRST_ROW Then Exit Function

    firstRow = lastRow - countWanted + 1
    If firstRow < PID_AL_FIRST_ROW Then firstRow = PID_AL_FIRST_ROW

    For r = firstRow To lastRow
        resultText = resultText & "- " & _
                     CStr(wsLog.Cells(r, "A").Value) & " | " & _
                     CStr(wsLog.Cells(r, "B").Value) & " | " & _
                     CStr(wsLog.Cells(r, "C").Value) & " | " & _
                     CStr(wsLog.Cells(r, "D").Value) & vbCrLf
    Next r

    PID_GetLastActionsText = resultText

SafeExit:
End Function


Public Sub PID_AdminShowActionLog()
    Dim logText As String

    logText = PID_GetLastActionsText(20)

    If logText = "" Then
        MsgBox "Es wurden noch keine Aktionen aufgezeichnet.", vbInformation, PID_ALTitle()
        Exit Sub
    End If

    MsgBox "Die letzten Aktionen in dieser Datei:" & vbCrLf & vbCrLf & logText, _
           vbInformation, PID_ALTitle()
End Sub


Private Function PID_ALGetOrCreateSheet() As Worksheet
    Dim ws As Worksheet
    Dim oldActive As Object

    On Error GoTo SafeExit

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_AL_SHEET)
    On Error GoTo SafeExit

    If ws Is Nothing Then
        ' Worksheets.Add aktiviert das neue Blatt - vorher merken, danach zurueck.
        Set oldActive = ActiveSheet

        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PID_AL_SHEET

        ws.Range("A1").Value = "Zeitpunkt"
        ws.Range("B1").Value = "Aktion"
        ws.Range("C1").Value = "Blatt"
        ws.Range("D1").Value = "Detail"

        ws.Visible = xlSheetVeryHidden

        On Error Resume Next
        If Not oldActive Is Nothing Then oldActive.Activate
        On Error GoTo SafeExit
    End If

    Set PID_ALGetOrCreateSheet = ws

SafeExit:
End Function


Private Function PID_ALGetLastRow(ByVal wsLog As Worksheet) As Long
    On Error GoTo SafeExit

    PID_ALGetLastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row

SafeExit:
End Function


' Aelteste Eintraege abschneiden, damit die Datei nicht dauerhaft waechst.
Private Sub PID_ALTrimToMaxRows(ByVal wsLog As Worksheet)
    Dim lastRow As Long
    Dim deleteCount As Long

    On Error GoTo SafeExit

    lastRow = PID_ALGetLastRow(wsLog)
    deleteCount = (lastRow - PID_AL_FIRST_ROW + 1) - PID_AL_MAX_ROWS

    If deleteCount <= 0 Then Exit Sub

    wsLog.Rows(PID_AL_FIRST_ROW & ":" & (PID_AL_FIRST_ROW + deleteCount - 1)).Delete

SafeExit:
End Sub


Private Function PID_ALActiveSheetName() As String
    On Error GoTo SafeExit

    If TypeName(ActiveSheet) = "Worksheet" Then PID_ALActiveSheetName = ActiveSheet.Name

SafeExit:
End Function


Private Function PID_ALTitle() As String
    PID_ALTitle = "Aktionsprotokoll"
End Function

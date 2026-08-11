Attribute VB_Name = "mod_PersonalIdUnique"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' TR-07 - Ein Mitarbeiter = eine Personal-ID.
'
' Zwei Stufen, beide aus Workbook_SheetChange heraus:
'
' 1. HART: Dieselbe ID darf auf einem Monatsblatt nur einmal vorkommen. Die neu
'    eingegebene Zeile wird zurueckgewiesen (Zelle geleert), der Bestand bleibt
'    unangetastet.
' 2. WEICH: Wird eine ID eingetragen, die in anderen Monaten zu einem ANDEREN
'    Namen gehoert, gibt es einen Hinweis auf das Korrektur-Makro. Genau diese
'    Konstellation erzeugt sonst TR-01 - die Aktualisierung sieht "ID|NAME" als
'    zwei verschiedene Mitarbeiter und friert den falschen Stand ein.
'
' Wichtig: waehrend CopyData und der Pflege-Makros ist EnableEvents = False,
' die Pruefung blockiert also nur die manuelle Eingabe und das Einfuegen.

Private Const PID_PIU_ID_COL As String = "B"
Private Const PID_PIU_NAME_COL As String = "C"
Private Const PID_PIU_MAX_LIST As Long = 8


Public Sub PID_CheckPersonalIdUniqueness(ByVal wsMonth As Worksheet, ByVal changedRange As Range)
    Dim idCells As Range
    Dim c As Range
    Dim data As Variant
    Dim isChanged() As Boolean
    Dim seenRow As Collection
    Dim rowCount As Long
    Dim r As Long
    Dim sheetRow As Long
    Dim idText As String
    Dim keyText As String
    Dim firstRow As Long
    Dim conflictInfo As String
    Dim conflictCount As Long
    Dim changedCount As Long
    Dim lastChangedRow As Long

    On Error GoTo SafeExit

    If wsMonth Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub

    Set idCells = Intersect(changedRange, PID_PIUIdColumnRange(wsMonth))
    If idCells Is Nothing Then Exit Sub

    rowCount = PID_LAST_ROW - PID_FIRST_ROW + 1
    ReDim isChanged(1 To rowCount)

    For Each c In idCells.Cells
        isChanged(c.Row - PID_FIRST_ROW + 1) = True
        changedCount = changedCount + 1
        lastChangedRow = c.Row
    Next c

    data = wsMonth.Range(PID_PIU_ID_COL & PID_FIRST_ROW & ":" & PID_PIU_NAME_COL & PID_LAST_ROW).Value
    Set seenRow = New Collection

    ' Durchgang 1: Bestand aufnehmen (alles, was gerade NICHT eingegeben wurde).
    For r = 1 To rowCount
        If Not isChanged(r) Then
            idText = Trim$(CStr(data(r, 1)))

            If idText <> "" Then
                keyText = PID_PIUBuildIdKey(idText)

                If PID_PIULookupRow(seenRow, keyText) = 0 Then
                    seenRow.Add PID_PIUSheetRow(r), keyText
                End If
            End If
        End If
    Next r

    ' Durchgang 2: nur die neuen Eingaben pruefen und ggf. zurueckweisen.
    For r = 1 To rowCount
        If isChanged(r) Then
            idText = Trim$(CStr(data(r, 1)))

            If idText <> "" Then
                keyText = PID_PIUBuildIdKey(idText)
                firstRow = PID_PIULookupRow(seenRow, keyText)
                sheetRow = PID_PIUSheetRow(r)

                If firstRow = 0 Then
                    seenRow.Add sheetRow, keyText
                Else
                    If conflictCount < PID_PIU_MAX_LIST Then
                        conflictInfo = conflictInfo & _
                            "- Zeile " & sheetRow & ": ID " & idText & " ist bereits in Zeile " & firstRow & _
                            PID_PIUNameHint(wsMonth, firstRow) & vbCrLf
                    End If

                    wsMonth.Cells(sheetRow, PID_PIU_ID_COL).ClearContents
                    conflictCount = conflictCount + 1
                End If
            End If
        End If
    Next r

    If conflictCount > 0 Then
        If conflictCount > PID_PIU_MAX_LIST Then
            conflictInfo = conflictInfo & "- ... und " & (conflictCount - PID_PIU_MAX_LIST) & " weitere" & vbCrLf
        End If

        PID_TrackAction "ID-Doppelung abgewiesen", _
                        wsMonth.Name & ": " & conflictCount & " Eingabe(n) entfernt"

        MsgBox "Eine Personal-ID darf auf einem Monatsblatt nur einmal vorkommen." & vbCrLf & vbCrLf & _
               conflictInfo & vbCrLf & _
               "Die Eingabe wurde entfernt. Bitte eine eindeutige Personal-ID eintragen.", _
               vbExclamation, PID_PIUTitleTaken()
        Exit Sub
    End If

    ' Weicher Hinweis nur bei einer einzelnen Handeingabe - beim Einfuegen vieler
    ' Zeilen waeren Popups pro Zeile unbrauchbar.
    If changedCount = 1 Then
        PID_PIUWarnIfIdUsedWithOtherName wsMonth, lastChangedRow
    End If

SafeExit:
End Sub


' Dieselbe ID in einem anderen Monat mit anderem Namen: kein Fehler, aber fast immer
' entweder ein Tippfehler oder eine Namensaenderung, die zentral korrigiert gehoert.
Private Sub PID_PIUWarnIfIdUsedWithOtherName(ByVal wsMonth As Worksheet, ByVal sheetRow As Long)
    Dim ws As Worksheet
    Dim data As Variant
    Dim monthNames As Variant
    Dim monthIndex As Long
    Dim r As Long
    Dim idText As String
    Dim nameText As String
    Dim otherId As String
    Dim otherName As String
    Dim foundInfo As String
    Dim foundCount As Long
    Dim seenNames As Collection
    Dim nameKey As String

    On Error GoTo SafeExit

    idText = Trim$(CStr(wsMonth.Cells(sheetRow, PID_PIU_ID_COL).Value))
    nameText = Trim$(CStr(wsMonth.Cells(sheetRow, PID_PIU_NAME_COL).Value))

    If idText = "" Then Exit Sub
    ' Ohne Namen ist der Vergleich sinnlos - die Zeile wird gerade erst gefuellt.
    If nameText = "" Then Exit Sub

    Set seenNames = New Collection
    monthNames = PID_MonthNames()

    For monthIndex = LBound(monthNames) To UBound(monthNames)
        If CStr(monthNames(monthIndex)) <> wsMonth.Name Then
            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(CStr(monthNames(monthIndex)))
            On Error GoTo SafeExit

            If Not ws Is Nothing Then
                data = ws.Range(PID_PIU_ID_COL & PID_FIRST_ROW & ":" & PID_PIU_NAME_COL & PID_LAST_ROW).Value

                For r = 1 To UBound(data, 1)
                    otherId = Trim$(CStr(data(r, 1)))

                    If StrComp(otherId, idText, vbTextCompare) = 0 Then
                        otherName = Trim$(CStr(data(r, 2)))

                        If otherName <> "" And StrComp(otherName, nameText, vbTextCompare) <> 0 Then
                            nameKey = "N|" & UCase$(otherName)

                            If Not PID_CollectionHasKey(seenNames, nameKey) Then
                                seenNames.Add otherName, nameKey

                                If foundCount < PID_PIU_MAX_LIST Then
                                    foundInfo = foundInfo & "- " & ws.Name & ": " & otherName & vbCrLf
                                End If

                                foundCount = foundCount + 1
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    Next monthIndex

    If foundCount = 0 Then Exit Sub

    MsgBox "Die Personal-ID " & idText & " wird in anderen Monaten mit einem anderen Namen verwendet:" & vbCrLf & vbCrLf & _
           foundInfo & vbCrLf & _
           "Hier: " & nameText & vbCrLf & vbCrLf & _
           "Wenn es derselbe Mitarbeiter ist, bitte das Makro " & Chr$(34) & "Personal-ID korrigieren" & Chr$(34) & _
           " verwenden. Sonst behandelt die Aktualisierung die Zeilen als zwei verschiedene Mitarbeiter.", _
           vbInformation, PID_PIUTitleCheck()

SafeExit:
End Sub


Private Function PID_PIUIdColumnRange(ByVal wsMonth As Worksheet) As Range
    Set PID_PIUIdColumnRange = wsMonth.Range(PID_PIU_ID_COL & PID_FIRST_ROW & ":" & PID_PIU_ID_COL & PID_LAST_ROW)
End Function


Private Function PID_PIUBuildIdKey(ByVal idText As String) As String
    ' Praefix, damit der Collection-Key nie mit einer reinen Zahl kollidiert.
    PID_PIUBuildIdKey = "ID|" & UCase$(Trim$(idText))
End Function


Private Function PID_PIULookupRow(ByVal seenRow As Collection, ByVal keyText As String) As Long
    On Error Resume Next
    PID_PIULookupRow = CLng(seenRow.item(keyText))
End Function


Private Function PID_PIUSheetRow(ByVal arrayIndex As Long) As Long
    PID_PIUSheetRow = arrayIndex + PID_FIRST_ROW - 1
End Function


Private Function PID_PIUNameHint(ByVal wsMonth As Worksheet, ByVal sheetRow As Long) As String
    Dim nameText As String

    On Error GoTo SafeExit

    nameText = Trim$(CStr(wsMonth.Cells(sheetRow, PID_PIU_NAME_COL).Value))
    If nameText <> "" Then PID_PIUNameHint = " (" & nameText & ")"

SafeExit:
End Function


Private Function PID_PIUTitleTaken() As String
    PID_PIUTitleTaken = "Personal-ID bereits vergeben"
End Function


Private Function PID_PIUTitleCheck() As String
    PID_PIUTitleCheck = "Personal-ID pr" & PID_UTxtUe() & "fen"
End Function

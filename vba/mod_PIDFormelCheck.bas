Attribute VB_Name = "mod_PIDFormelCheck"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' TR-10 — Formelspalten G, H, K und L pruefen und ergaenzen.
'
' Hintergrund: Aeltere Versionen haben beim Loeschen eines Mitarbeiters die ganze
' Zeile B:N geleert und damit auch die vier Formelspalten mitgenommen. Solche Zeilen
' bleiben "tot": ein neuer Mitarbeiter in derselben Zeile bekommt keinen Lohn, keine
' aktuellen Stunden, kein Urlaubsgeld und kein letztes Gehalt.
'
' Die vier Wiederherstellungs-Routinen im Full Refresh schreiben zwar ganze Spalten,
' ueberspringen ein Monatsblatt aber still, wenn der Monatsindex in A1 fehlt oder
' nicht zum Blattnamen passt (PID_IsWorkerMonthSheet, IsNumeric(A1)-Guards). Dann
' bleiben genau die kaputten Zeilen leer und der Benutzer sieht keinen Hinweis.
'
' Dieses Modul schliesst die Luecke:
'   PID_PruefeFormelspalten          - reine Diagnose, aendert nichts
'   PID_FormelspaltenReparieren      - Reparatur mit Rueckmeldung
'   PID_RepairFormulaColumnsSilent   - Reparatur ohne Dialog (Full Refresh)
'
' Geschrieben wird nur, wo tatsaechlich eine Formel fehlt: PID_EnsureCellFormula
' (Modul1) prueft HasFormula und laesst vorhandene Formeln unangetastet.
'
' Ausnahme Spalte G: dort ist ein von PID_ForceMonatslohnRecalcForRow geschriebener
' Zahlenwert der gewollte Zustand (schneller als 960 UDF-Aufrufe). Zeilen mit KV-Gruppe,
' Stunden und Zahlenwert gelten deshalb als in Ordnung — Details in
' PID_RowNeedsMonatslohnFormula (Modul1).

Private Const PID_FC_TITLE As String = "Formelspalten"
Private Const PID_FC_COLUMNS As String = "G,H,K,L"


' Diagnose fuer den Tester: welche Monatsblaetter haben fehlende Formeln,
' einen falschen Monatsindex in A1 oder sind gesperrt.
Public Sub PID_PruefeFormelspalten()
    Dim monthIndex As Long
    Dim ws As Worksheet
    Dim missingOnSheet As Long
    Dim totalMissing As Long
    Dim sheetsFound As Long
    Dim detailText As String
    Dim indexText As String
    Dim protectedText As String
    Dim indexInCell As Long

    On Error GoTo CleanFail

    For monthIndex = 1 To 12
        Set ws = PID_FCGetMonthSheet(monthIndex)

        If Not ws Is Nothing Then
            sheetsFound = sheetsFound + 1

            missingOnSheet = PID_FCCountMissingFormulaCells(ws)
            totalMissing = totalMissing + missingOnSheet

            If missingOnSheet > 0 Then
                detailText = detailText & "- " & ws.Name & ": " & missingOnSheet & _
                             " Zellen ohne Formel" & vbCrLf
            End If

            indexInCell = PID_FCReadMonthIndexCell(ws)
            If indexInCell <> monthIndex Then
                indexText = indexText & "- " & ws.Name & vbCrLf
            End If

            If ws.ProtectContents Then
                protectedText = protectedText & ws.Name & " "
            End If
        End If
    Next monthIndex

    MsgBox PID_FCBuildReportText(sheetsFound, totalMissing, detailText, indexText, protectedText), _
           vbInformation, PID_FC_TITLE

    Exit Sub

CleanFail:
    MsgBox "Fehler bei der " & PID_UTxtPruefung() & " der Formelspalten:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, PID_FC_TITLE
End Sub


' Benutzer-Einstieg (Alt+F8 und _ADMIN-Schaltflaeche): fehlende Formeln ergaenzen.
Public Sub PID_FormelspaltenReparieren()
    Dim repairedCells As Long
    Dim repairedSheets As Long
    Dim fixedIndexSheets As Long
    Dim checkedSheets As Long

    checkedSheets = PID_RepairFormulaColumnsSilent(repairedCells, repairedSheets, fixedIndexSheets)

    PID_TrackAction "Formelspalten reparieren", _
                    repairedCells & " Zellen auf " & repairedSheets & " " & PID_UTxtMonatsblaettern() & _
                    ", A1: " & fixedIndexSheets

    If repairedCells = 0 And fixedIndexSheets = 0 Then
        MsgBox "Alle Formeln in den Spalten " & PID_FC_COLUMNS & " sind vorhanden." & vbCrLf & vbCrLf & _
               PID_UTxtGeprueft() & ": " & checkedSheets & " " & PID_UTxtMonatsblaetter() & ".", _
               vbInformation, PID_FC_TITLE
        Exit Sub
    End If

    MsgBox "Fehlende Formeln wurden " & PID_UTxtErgaenzt() & "." & vbCrLf & vbCrLf & _
           PID_UTxtGeprueft() & ": " & checkedSheets & " " & PID_UTxtMonatsblaetter() & vbCrLf & _
           "Neue Formeln: " & repairedCells & " Zellen auf " & repairedSheets & " " & PID_UTxtMonatsblaettern() & vbCrLf & _
           "Monatsindex A1 korrigiert: " & fixedIndexSheets, _
           vbInformation, PID_FC_TITLE
End Sub


' Reparatur ohne Dialog. Rueckgabe: Anzahl gepruefter Monatsblaetter.
' repairedCells / repairedSheets / fixedIndexSheets liefern die Details fuer die
' Abschlussmeldung des Aufrufers (Full Refresh).
Public Function PID_RepairFormulaColumnsSilent(ByRef repairedCells As Long, _
                                               ByRef repairedSheets As Long, _
                                               ByRef fixedIndexSheets As Long) As Long
    Dim monthIndex As Long
    Dim ws As Worksheet
    Dim checkedSheets As Long
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean

    repairedCells = 0
    repairedSheets = 0
    fixedIndexSheets = 0

    On Error GoTo SafeExit

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For monthIndex = 1 To 12
        Set ws = PID_FCGetMonthSheet(monthIndex)

        If Not ws Is Nothing Then
            checkedSheets = checkedSheets + 1
            PID_FCRepairSheet ws, monthIndex, repairedCells, repairedSheets, fixedIndexSheets
        End If
    Next monthIndex

SafeExit:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    PID_RepairFormulaColumnsSilent = checkedSheets
End Function


' Ein Monatsblatt reparieren. Eigene Fehlerbehandlung, damit ein Problem auf Januar
' die restlichen elf Monate nicht mitnimmt und das Blatt nicht entsperrt zurueckbleibt.
Private Sub PID_FCRepairSheet(ByVal ws As Worksheet, _
                              ByVal monthIndex As Long, _
                              ByRef repairedCells As Long, _
                              ByRef repairedSheets As Long, _
                              ByRef fixedIndexSheets As Long)
    Dim missingBefore As Long
    Dim missingAfter As Long
    Dim wasProtected As Boolean

    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Sub

    missingBefore = PID_FCCountMissingFormulaCells(ws)

    If missingBefore = 0 And PID_FCReadMonthIndexCell(ws) = monthIndex Then Exit Sub

    wasProtected = ws.ProtectContents
    PID_FCTryUnprotect ws

    If PID_FCEnsureMonthIndexCell(ws, monthIndex) Then
        fixedIndexSheets = fixedIndexSheets + 1
    End If

    If missingBefore > 0 Then
        PID_RestoreFormulaColumnsForRows ws, PID_FIRST_ROW, PID_LAST_ROW

        missingAfter = PID_FCCountMissingFormulaCells(ws)

        If missingAfter < missingBefore Then
            repairedCells = repairedCells + (missingBefore - missingAfter)
            repairedSheets = repairedSheets + 1
        End If
    End If

SafeExit:
    On Error Resume Next
    If wasProtected Then PID_ReprotectWorksheet ws
End Sub


' Anzahl der Zellen in G, H, K und L (Zeile 3 bis 82) ohne Formel.
' Range.HasFormula liefert fuer die ganze Spalte True/False und nur bei gemischtem
' Inhalt Null - so bleibt der teure Zeilen-Durchlauf die Ausnahme.
Private Function PID_FCCountMissingFormulaCells(ByVal ws As Worksheet) As Long
    Dim columnKeys As Variant
    Dim i As Long
    Dim missing As Long

    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Function

    columnKeys = Split(PID_FC_COLUMNS, ",")

    For i = LBound(columnKeys) To UBound(columnKeys)
        missing = missing + PID_FCCountMissingInColumn(ws, CStr(columnKeys(i)))
    Next i

SafeExit:
    PID_FCCountMissingFormulaCells = missing
End Function


Private Function PID_FCCountMissingInColumn(ByVal ws As Worksheet, ByVal columnKey As String) As Long
    Dim columnRange As Range
    Dim hasFormulaState As Variant
    Dim r As Long
    Dim missing As Long

    On Error GoTo SafeExit

    ' Spalte G darf statt der Formel den von PID_ForceMonatslohnRecalcForRow
    ' geschriebenen Zahlenwert tragen. Ohne diese Ausnahme meldete die Pruefung bis zu
    ' 50 Zeilen je Monatsblatt als "Formel fehlt" und die Reparatur haette die schnellen
    ' Werte durch 960 UDF-Formeln ersetzt.
    If StrComp(columnKey, "G", vbTextCompare) = 0 Then
        For r = PID_FIRST_ROW To PID_LAST_ROW
            If PID_RowNeedsMonatslohnFormula(ws, r) Then missing = missing + 1
        Next r
        GoTo SafeExit
    End If

    Set columnRange = ws.Range(columnKey & PID_FIRST_ROW & ":" & columnKey & PID_LAST_ROW)
    hasFormulaState = columnRange.HasFormula

    If IsNull(hasFormulaState) Then
        For r = PID_FIRST_ROW To PID_LAST_ROW
            If Not ws.Cells(r, columnKey).HasFormula Then missing = missing + 1
        Next r
    ElseIf hasFormulaState = False Then
        missing = PID_LAST_ROW - PID_FIRST_ROW + 1
    End If

SafeExit:
    PID_FCCountMissingInColumn = missing
End Function


' A1 traegt den Monatsindex. Fehlt er oder passt er nicht zum Blattnamen, halten
' PID_IsWorkerMonthSheet und die IsNumeric(A1)-Guards das Blatt fuer ungueltig und
' ueberspringen es bei jeder Wiederherstellung. Das Blatt muss entsperrt sein.
Private Function PID_FCEnsureMonthIndexCell(ByVal ws As Worksheet, ByVal monthIndex As Long) As Boolean
    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Function
    If monthIndex < 1 Or monthIndex > 12 Then Exit Function
    If PID_FCReadMonthIndexCell(ws) = monthIndex Then Exit Function

    ws.Range("A1").Value2 = monthIndex

    PID_FCEnsureMonthIndexCell = (PID_FCReadMonthIndexCell(ws) = monthIndex)

SafeExit:
End Function


' Monatsindex aus A1; 0 bedeutet leer, Text oder ausserhalb 1 bis 12.
Private Function PID_FCReadMonthIndexCell(ByVal ws As Worksheet) As Long
    Dim cellValue As Variant

    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Function

    cellValue = ws.Range("A1").Value2
    If Not IsNumeric(cellValue) Then Exit Function

    If CLng(cellValue) >= 1 And CLng(cellValue) <= 12 Then
        PID_FCReadMonthIndexCell = CLng(cellValue)
    End If

SafeExit:
End Function


Private Function PID_FCGetMonthSheet(ByVal monthIndex As Long) As Worksheet
    Dim monthNames As Variant

    On Error GoTo SafeExit

    If monthIndex < 1 Or monthIndex > 12 Then Exit Function

    monthNames = PID_MonthNames()

    On Error Resume Next
    Set PID_FCGetMonthSheet = ThisWorkbook.Worksheets(CStr(monthNames(monthIndex - 1)))
    Err.Clear

SafeExit:
End Function


Private Sub PID_FCTryUnprotect(ByVal ws As Worksheet)
    On Error Resume Next

    If ws Is Nothing Then Exit Sub

    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If ws.ProtectContents Then ws.Unprotect

    Err.Clear
End Sub


Private Function PID_FCBuildReportText(ByVal sheetsFound As Long, _
                                       ByVal totalMissing As Long, _
                                       ByVal detailText As String, _
                                       ByVal indexText As String, _
                                       ByVal protectedText As String) As String
    Dim reportText As String

    reportText = PID_UTxtGeprueft() & ": " & sheetsFound & " " & PID_UTxtMonatsblaetter() & vbCrLf & _
                 "Spalten: " & PID_FC_COLUMNS & " (Zeile " & PID_FIRST_ROW & " bis " & PID_LAST_ROW & ")" & vbCrLf & vbCrLf

    If totalMissing = 0 Then
        reportText = reportText & "Alle Formeln sind vorhanden." & vbCrLf
    Else
        reportText = reportText & "Fehlende Formeln: " & totalMissing & " Zellen" & vbCrLf & _
                     detailText & vbCrLf & _
                     "Bitte ""Formeln reparieren"" " & PID_UTxtAusfuehren() & "." & vbCrLf
    End If

    If Len(indexText) > 0 Then
        reportText = reportText & vbCrLf & "Monatsindex in A1 fehlt oder passt nicht:" & vbCrLf & indexText
    End If

    If Len(protectedText) > 0 Then
        reportText = reportText & vbCrLf & "Gesperrte " & PID_UTxtBlaetter() & ": " & _
                     Trim$(protectedText) & vbCrLf
    End If

    PID_FCBuildReportText = reportText
End Function

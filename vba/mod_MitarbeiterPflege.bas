Attribute VB_Name = "mod_MitarbeiterPflege"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Stammdatenpflege fuer einzelne Mitarbeiter ueber alle Monatsblaetter hinweg.
'
' TR-06  PersonalIdKorrigieren  - Personal-ID und/oder Name korrigieren
' TR-08  MitarbeiterEntfernen   - Mitarbeiter aus allen oder ab einem Monat entfernen
'
' Warum zentral und nicht von Hand:
' Der Mitarbeiter-Schluessel ist "ID|NAME" (PID_BuildEmployeeKey in mod_CopyData).
' Wird nur auf einzelnen Blaettern von Hand korrigiert oder geleert, behandelt
' CopyData die abweichenden Zeilen als eigenstaendige Neuzugaenge und stellt den
' alten Stand wieder her. Zusaetzlich bleiben Zeilen im versteckten Stunden-Log
' PID_HOUR_OVERRIDES zurueck, weil sie denselben Schluessel speichern - die
' gespeicherten Monats-Overrides verwaisen oder tauchen spaeter wieder auf.
'
' mod_CopyData wird bewusst NICHT veraendert.

Private Const PID_MP_LOG_SHEET As String = "PID_HOUR_OVERRIDES"
Private Const PID_MP_LOG_FIRST_ROW As Long = 2
Private Const PID_MP_LOG_YEAR_COL As String = "A"
Private Const PID_MP_LOG_MONTH_COL As String = "B"
Private Const PID_MP_LOG_KEY_COL As String = "C"


'==============================================================================
' Einstiegspunkte
'==============================================================================

' Alt+F8 (Muster wie DataClear / CopyData).
Public Sub PersonalIdKorrigieren()
    PID_KorrigierePersonalIdUndName
End Sub


Public Sub MitarbeiterEntfernen()
    PID_EntferneMitarbeiterAusMonaten
End Sub


' Buttons auf dem Admin-Panel.
Public Sub PID_AdminKorrigierePersonalId()
    PID_KorrigierePersonalIdUndName
End Sub


Public Sub PID_AdminMitarbeiterEntfernen()
    PID_EntferneMitarbeiterAusMonaten
End Sub


'==============================================================================
' TR-06 - Personal-ID / Name korrigieren
'==============================================================================

Public Sub PID_KorrigierePersonalIdUndName()
    Dim oldId As String
    Dim oldName As String
    Dim newId As String
    Dim newName As String
    Dim oldKey As String
    Dim newKey As String

    Dim matchCount As Long
    Dim matchInfo As String
    Dim conflictInfo As String
    Dim changedRows As Long
    Dim changedLogRows As Long

    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim stateChanged As Boolean
    Dim originalErrNumber As Long
    Dim originalErrDescription As String

    On Error GoTo CleanFail

    If Not PID_MPAskForEmployee(oldId, oldName, PID_MPFixTitle()) Then Exit Sub
    If Not PID_MPAskForNewValues(oldId, oldName, newId, newName) Then Exit Sub

    oldKey = PID_MPBuildKey(oldId, oldName)
    newKey = PID_MPBuildKey(newId, newName)

    If oldKey = newKey Then
        MsgBox "Keine " & PID_UTxtAenderung() & " - ID und Name sind " & PID_UTxtUnveraendert() & ".", _
               vbInformation, PID_MPFixTitle()
        Exit Sub
    End If

    ' Erst pruefen, dann schreiben: bei einem Konflikt bleibt alles unveraendert.
    PID_MPAnalyzeForRename oldKey, newKey, newId, matchCount, matchInfo, conflictInfo

    If matchCount = 0 Then
        MsgBox "Auf keinem Monatsblatt gefunden:" & vbCrLf & vbCrLf & _
               oldId & " / " & oldName, _
               vbExclamation, PID_MPFixTitle()
        Exit Sub
    End If

    If conflictInfo <> "" Then
        MsgBox "Korrektur nicht " & PID_UTxtMoeglich() & " - es wurde nichts " & PID_UTxtGeaendert() & ":" & vbCrLf & vbCrLf & _
               conflictInfo, _
               vbExclamation, PID_MPFixTitle()
        Exit Sub
    End If

    If MsgBox( _
        "Mitarbeiter wird auf allen " & PID_UTxtMonatsblaettern() & " korrigiert:" & vbCrLf & vbCrLf & _
        "ALT:  " & oldId & " / " & oldName & vbCrLf & _
        "NEU:  " & newId & " / " & newName & vbCrLf & vbCrLf & _
        "Betroffen (" & matchCount & " Zeile(n)):" & vbCrLf & _
        matchInfo & vbCrLf & _
        "Das Stunden-Log wird mit korrigiert." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        PID_MPFixTitle()) <> vbYes Then
        Exit Sub
    End If

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    stateChanged = True

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    changedRows = PID_MPApplyRenameToMonthSheets(oldKey, newId, newName)
    changedLogRows = PID_MPRenameKeyInOverrideLog(oldKey, newKey)

    MarkFluktuationDirty

    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    stateChanged = False

    PID_TrackAction "PersonalIdKorrigieren", _
                    oldId & " / " & oldName & " -> " & newId & " / " & newName & _
                    " (" & changedRows & " Zeilen, " & changedLogRows & " Log)"

    MsgBox changedRows & " Mitarbeiterzeile(n) auf den " & PID_UTxtMonatsblaettern() & " " & PID_UTxtGeaendert() & "." & vbCrLf & _
           changedLogRows & " " & PID_UTxtEintraege() & " im Stunden-Log " & PID_UTxtGeaendert() & "." & vbCrLf & vbCrLf & _
           "NEU:  " & newId & " / " & newName, _
           vbInformation, PID_MPFixTitle()

    Exit Sub

CleanFail:
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description

    If stateChanged Then
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
    End If

    MsgBox "Fehler bei PID_KorrigierePersonalIdUndName:" & vbCrLf & _
           originalErrNumber & " - " & originalErrDescription, _
           vbExclamation, PID_MPFixTitle()
End Sub


'==============================================================================
' TR-08 - Mitarbeiter entfernen
'==============================================================================

Public Sub PID_EntferneMitarbeiterAusMonaten()
    Dim empId As String
    Dim empName As String
    Dim empKey As String
    Dim startMonth As Long
    Dim defaultMonth As Long

    Dim matchCount As Long
    Dim matchInfo As String
    Dim clearedRows As Long
    Dim removedLogRows As Long

    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim stateChanged As Boolean
    Dim originalErrNumber As Long
    Dim originalErrDescription As String

    On Error GoTo CleanFail

    defaultMonth = PID_MPGetActiveMonthIndex()

    If Not PID_MPAskForEmployee(empId, empName, PID_MPRemoveTitle()) Then Exit Sub

    startMonth = PID_MPAskForStartMonth(defaultMonth)
    If startMonth = 0 Then Exit Sub

    empKey = PID_MPBuildKey(empId, empName)
    If empKey = "" Then Exit Sub

    PID_MPAnalyzeForRemoval empKey, startMonth, matchCount, matchInfo

    If matchCount = 0 Then
        MsgBox "In diesem Zeitraum nicht gefunden:" & vbCrLf & vbCrLf & _
               empId & " / " & empName, _
               vbExclamation, PID_MPRemoveTitle()
        Exit Sub
    End If

    If MsgBox( _
        "Mitarbeiter wird aus den Monatsdaten entfernt:" & vbCrLf & vbCrLf & _
        empId & " / " & empName & vbCrLf & vbCrLf & _
        "Zeitraum:  " & PID_MPMonthName(startMonth) & " bis Dezember" & vbCrLf & vbCrLf & _
        "Betroffen (" & matchCount & " Zeile(n)):" & vbCrLf & _
        matchInfo & vbCrLf & _
        "Die Eingaben B:F, I:J und M:N werden " & PID_UTxtGeloescht() & "; " & _
        "die Formelspalten G, H, K, L sowie Formate und Struktur bleiben erhalten." & vbCrLf & _
        "Die Stunden-" & PID_UTxtEintraege() & " dieses Mitarbeiters werden " & PID_UTxtGeloescht() & "." & vbCrLf & vbCrLf & _
        "HINWEIS: Bei einem echten Austritt bitte das Austrittsdatum (Spalte I) " & _
        "verwenden statt zu " & PID_UTxtLoeschen() & "." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbExclamation + vbYesNo, _
        PID_MPRemoveTitle()) <> vbYes Then
        Exit Sub
    End If

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    stateChanged = True

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    clearedRows = PID_MPClearEmployeeRows(empKey, startMonth)
    removedLogRows = PID_MPDeleteFromOverrideLog(empKey, startMonth)

    MarkFluktuationDirty
    MarkAllKVDropdownsDirty

    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    stateChanged = False

    PID_TrackAction "MitarbeiterEntfernen", _
                    empId & " / " & empName & " ab " & PID_MPMonthName(startMonth) & _
                    " (" & clearedRows & " Zeilen, " & removedLogRows & " Log)"

    MsgBox clearedRows & " Mitarbeiterzeile(n) " & PID_UTxtGeloescht() & " (" & _
           PID_MPMonthName(startMonth) & " bis Dezember)." & vbCrLf & _
           removedLogRows & " " & PID_UTxtEintraege() & " im Stunden-Log " & PID_UTxtGeloescht() & ".", _
           vbInformation, PID_MPRemoveTitle()

    Exit Sub

CleanFail:
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description

    If stateChanged Then
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
    End If

    MsgBox "Fehler bei PID_EntferneMitarbeiterAusMonaten:" & vbCrLf & _
           originalErrNumber & " - " & originalErrDescription, _
           vbExclamation, PID_MPRemoveTitle()
End Sub


'==============================================================================
' Mitarbeiter bestimmen
'==============================================================================

' Bevorzugt aus der markierten Zeile auf einem Monatsblatt (keine Tippfehler).
' Fallback InputBox, wenn das Makro nicht von einem Monatsblatt aus gestartet wird
' (z. B. ueber den Button auf dem Admin-Panel).
Private Function PID_MPAskForEmployee(ByRef outId As String, _
                                      ByRef outName As String, _
                                      ByVal dialogTitle As String) As Boolean
    If PID_MPGetEmployeeFromSelection(outId, outName) Then
        PID_MPAskForEmployee = True
        Exit Function
    End If

    PID_MPAskForEmployee = PID_MPGetEmployeeFromInputBox(outId, outName, dialogTitle)
End Function


Private Function PID_MPGetEmployeeFromSelection(ByRef outId As String, ByRef outName As String) As Boolean
    Dim ws As Worksheet
    Dim targetRow As Long

    On Error GoTo SafeExit

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function

    Set ws = ActiveSheet
    If Not PID_IsWorkerMonthSheet(ws) Then Exit Function

    If Selection Is Nothing Then Exit Function
    If TypeName(Selection) <> "Range" Then Exit Function

    targetRow = Selection.Cells(1, 1).Row
    If targetRow < PID_FIRST_ROW Or targetRow > PID_LAST_ROW Then Exit Function

    outId = Trim$(CStr(ws.Cells(targetRow, "B").Value))
    outName = Trim$(CStr(ws.Cells(targetRow, "C").Value))

    If outId = "" And outName = "" Then Exit Function

    PID_MPGetEmployeeFromSelection = True

SafeExit:
End Function


' Sucht die ID auf allen Monatsblaettern. Nur eindeutig, wenn genau ein Name zu
' dieser ID gehoert - sonst muss der Benutzer die Zeile markieren.
Private Function PID_MPGetEmployeeFromInputBox(ByRef outId As String, _
                                               ByRef outName As String, _
                                               ByVal dialogTitle As String) As Boolean
    Dim answer As Variant
    Dim searchId As String
    Dim foundNames As Collection
    Dim nameList As String
    Dim item As Variant

    On Error GoTo SafeExit

    answer = Application.InputBox( _
        "Bitte die Personal-ID des Mitarbeiters eingeben." & vbCrLf & vbCrLf & _
        "Tipp: Einfacher ist es, vorher die Mitarbeiterzeile auf einem Monatsblatt zu markieren.", _
        dialogTitle, "", , , , , 2)

    If VarType(answer) = vbBoolean Then Exit Function

    searchId = Trim$(CStr(answer))
    If searchId = "" Then Exit Function

    Set foundNames = PID_MPCollectNamesForId(searchId)

    If foundNames.Count = 0 Then
        MsgBox "Diese Personal-ID wurde auf keinem Monatsblatt gefunden:" & vbCrLf & searchId, _
               vbExclamation, dialogTitle
        Exit Function
    End If

    If foundNames.Count > 1 Then
        For Each item In foundNames
            nameList = nameList & "- " & CStr(item) & vbCrLf
        Next item

        MsgBox "Diese Personal-ID geh" & PID_UTxtOe() & "rt zu mehreren Namen:" & vbCrLf & vbCrLf & _
               nameList & vbCrLf & _
               "Bitte die gew" & PID_UTxtUe() & "nschte Mitarbeiterzeile auf einem Monatsblatt markieren " & _
               "und das Makro erneut " & PID_UTxtAusfuehren() & ".", _
               vbExclamation, dialogTitle
        Exit Function
    End If

    outId = searchId
    outName = CStr(foundNames.item(1))

    PID_MPGetEmployeeFromInputBox = True

SafeExit:
End Function


Private Function PID_MPCollectNamesForId(ByVal searchId As String) As Collection
    Dim result As Collection
    Dim ws As Worksheet
    Dim data As Variant
    Dim monthIndex As Long
    Dim r As Long
    Dim rowId As String
    Dim rowName As String
    Dim nameKey As String

    Set result = New Collection
    Set PID_MPCollectNamesForId = result

    On Error GoTo SafeExit

    For monthIndex = 1 To 12
        Set ws = PID_MPGetMonthSheet(monthIndex)

        If Not ws Is Nothing Then
            data = PID_MPReadIdNameBlock(ws)

            For r = 1 To UBound(data, 1)
                rowId = Trim$(CStr(data(r, 1)))
                rowName = Trim$(CStr(data(r, 2)))

                If StrComp(rowId, searchId, vbTextCompare) = 0 Then
                    ' Praefix, damit der Collection-Key auch bei leerem Namen gueltig bleibt.
                    nameKey = "N|" & UCase$(rowName)

                    If Not PID_CollectionHasKey(result, nameKey) Then
                        result.Add rowName, nameKey
                    End If
                End If
            Next r
        End If
    Next monthIndex

SafeExit:
End Function


Private Function PID_MPAskForNewValues(ByVal oldId As String, _
                                       ByVal oldName As String, _
                                       ByRef outNewId As String, _
                                       ByRef outNewName As String) As Boolean
    Dim answer As Variant

    On Error GoTo SafeExit

    answer = Application.InputBox( _
        "Aktuell:  " & oldId & " / " & oldName & vbCrLf & vbCrLf & _
        "Neue Personal-ID eingeben." & vbCrLf & _
        "Leer lassen = ID bleibt " & PID_UTxtUnveraendert() & ".", _
        PID_MPFixTitle(), oldId, , , , , 2)

    If VarType(answer) = vbBoolean Then Exit Function

    outNewId = Trim$(CStr(answer))
    If outNewId = "" Then outNewId = oldId

    answer = Application.InputBox( _
        "Aktuell:  " & oldId & " / " & oldName & vbCrLf & vbCrLf & _
        "Neuer Name eingeben." & vbCrLf & _
        "Leer lassen = Name bleibt " & PID_UTxtUnveraendert() & ".", _
        PID_MPFixTitle(), oldName, , , , , 2)

    If VarType(answer) = vbBoolean Then Exit Function

    outNewName = Trim$(CStr(answer))
    If outNewName = "" Then outNewName = oldName

    If outNewId = "" And outNewName = "" Then
        MsgBox "ID und Name d" & PID_UTxtUe() & "rfen nicht beide leer sein.", _
               vbExclamation, PID_MPFixTitle()
        Exit Function
    End If

    PID_MPAskForNewValues = True

SafeExit:
End Function


' Rueckgabe: 0 = abgebrochen, sonst Startmonat 1..12 (1 = alle Monate).
Private Function PID_MPAskForStartMonth(ByVal defaultMonth As Long) As Long
    Dim answer As VbMsgBoxResult
    Dim inputValue As Variant
    Dim monthText As String
    Dim monthIndex As Long

    On Error GoTo SafeExit

    If defaultMonth < 1 Or defaultMonth > 12 Then defaultMonth = 1

    answer = MsgBox( _
        "Aus welchen Monaten soll der Mitarbeiter entfernt werden?" & vbCrLf & vbCrLf & _
        "JA          = alle 12 Monate (Januar bis Dezember)" & vbCrLf & _
        "NEIN        = erst ab einem bestimmten Monat" & vbCrLf & _
        "ABBRECHEN   = nichts tun", _
        vbQuestion + vbYesNoCancel, _
        PID_MPRemoveTitle())

    If answer = vbYes Then
        PID_MPAskForStartMonth = 1
        Exit Function
    End If

    If answer <> vbNo Then Exit Function

    inputValue = Application.InputBox( _
        "Ab welchem Monat entfernen?" & vbCrLf & vbCrLf & _
        "Monatsnummer 1-12 eingeben (1 = Januar, 12 = Dezember)." & vbCrLf & _
        "Alle fr" & PID_UTxtUe() & "heren Monate bleiben " & PID_UTxtUnveraendert() & ".", _
        PID_MPRemoveTitle(), CStr(defaultMonth), , , , , 2)

    If VarType(inputValue) = vbBoolean Then Exit Function

    monthText = Trim$(CStr(inputValue))
    If monthText = "" Then Exit Function

    If Not IsNumeric(monthText) Then
        MsgBox "Bitte eine " & PID_UTxtGueltig() & "e Monatsnummer 1-12 eingeben.", _
               vbExclamation, PID_MPRemoveTitle()
        Exit Function
    End If

    monthIndex = CLng(Val(monthText))

    If monthIndex < 1 Or monthIndex > 12 Then
        MsgBox "Bitte eine " & PID_UTxtGueltig() & "e Monatsnummer 1-12 eingeben.", _
               vbExclamation, PID_MPRemoveTitle()
        Exit Function
    End If

    PID_MPAskForStartMonth = monthIndex

SafeExit:
End Function


'==============================================================================
' Pruefen
'==============================================================================

' Zaehlt die Treffer und sammelt Konflikte. Konflikt = die Ziel-ID gehoert bereits
' zu einem anderen Mitarbeiter, oder die Zielzeile existiert im selben Monat schon.
Private Sub PID_MPAnalyzeForRename(ByVal oldKey As String, _
                                   ByVal newKey As String, _
                                   ByVal newId As String, _
                                   ByRef outMatchCount As Long, _
                                   ByRef outMatchInfo As String, _
                                   ByRef outConflictInfo As String)
    Dim ws As Worksheet
    Dim data As Variant
    Dim monthIndex As Long
    Dim r As Long
    Dim rowKey As String
    Dim rowId As String
    Dim rowName As String
    Dim monthMatches As Long

    outMatchCount = 0
    outMatchInfo = ""
    outConflictInfo = ""

    On Error GoTo SafeExit

    For monthIndex = 1 To 12
        Set ws = PID_MPGetMonthSheet(monthIndex)

        If Not ws Is Nothing Then
            data = PID_MPReadIdNameBlock(ws)
            monthMatches = 0

            For r = 1 To UBound(data, 1)
                rowId = Trim$(CStr(data(r, 1)))
                rowName = Trim$(CStr(data(r, 2)))
                rowKey = PID_MPBuildKey(rowId, rowName)

                If rowKey <> "" Then
                    If rowKey = oldKey Then
                        monthMatches = monthMatches + 1
                    ElseIf rowKey = newKey Then
                        outConflictInfo = outConflictInfo & _
                            "- " & ws.Name & " Zeile " & PID_MPSheetRow(r) & _
                            ": Zielzeile existiert bereits" & vbCrLf
                    ElseIf StrComp(rowId, newId, vbTextCompare) = 0 Then
                        outConflictInfo = outConflictInfo & _
                            "- " & ws.Name & " Zeile " & PID_MPSheetRow(r) & _
                            ": ID " & newId & " geh" & PID_UTxtOe() & "rt zu " & rowName & vbCrLf
                    End If
                End If
            Next r

            If monthMatches > 0 Then
                outMatchCount = outMatchCount + monthMatches
                outMatchInfo = outMatchInfo & "- " & ws.Name & " (" & monthMatches & ")" & vbCrLf
            End If
        End If
    Next monthIndex

SafeExit:
End Sub


Private Sub PID_MPAnalyzeForRemoval(ByVal empKey As String, _
                                    ByVal startMonth As Long, _
                                    ByRef outMatchCount As Long, _
                                    ByRef outMatchInfo As String)
    Dim ws As Worksheet
    Dim data As Variant
    Dim monthIndex As Long
    Dim r As Long
    Dim monthMatches As Long

    outMatchCount = 0
    outMatchInfo = ""

    On Error GoTo SafeExit

    For monthIndex = startMonth To 12
        Set ws = PID_MPGetMonthSheet(monthIndex)

        If Not ws Is Nothing Then
            data = PID_MPReadIdNameBlock(ws)
            monthMatches = 0

            For r = 1 To UBound(data, 1)
                If PID_MPBuildKey(data(r, 1), data(r, 2)) = empKey Then
                    monthMatches = monthMatches + 1
                End If
            Next r

            If monthMatches > 0 Then
                outMatchCount = outMatchCount + monthMatches
                outMatchInfo = outMatchInfo & "- " & ws.Name & " (" & monthMatches & ")" & vbCrLf
            End If
        End If
    Next monthIndex

SafeExit:
End Sub


'==============================================================================
' Schreiben
'==============================================================================

Private Function PID_MPApplyRenameToMonthSheets(ByVal oldKey As String, _
                                                ByVal newId As String, _
                                                ByVal newName As String) As Long
    Dim ws As Worksheet
    Dim data As Variant
    Dim monthIndex As Long
    Dim r As Long
    Dim sheetRow As Long
    Dim changed As Long
    Dim sheetTouched As Boolean

    On Error GoTo SafeExit

    For monthIndex = 1 To 12
        Set ws = PID_MPGetMonthSheet(monthIndex)

        If Not ws Is Nothing Then
            data = PID_MPReadIdNameBlock(ws)
            sheetTouched = False

            For r = 1 To UBound(data, 1)
                If PID_MPBuildKey(data(r, 1), data(r, 2)) = oldKey Then
                    If Not sheetTouched Then
                        PID_MPTryUnprotectMonthSheet ws
                        sheetTouched = True
                    End If

                    sheetRow = PID_MPSheetRow(r)

                    PID_MPWriteIdCell ws.Cells(sheetRow, "B"), newId
                    ws.Cells(sheetRow, "C").Value = newName

                    changed = changed + 1
                End If
            Next r

            If sheetTouched Then
                PID_MPTryProtectMonthSheet ws
            End If
        End If
    Next monthIndex

SafeExit:
    PID_MPApplyRenameToMonthSheets = changed
End Function


' Wie PID_ClearOnlySelectedEmployeeRows: nur die Eingabespalten der Zeile leeren
' (B-F, I-J, M-N). Die Formelspalten G, H, K und L bleiben stehen.
' Zeilen werden NICHT geloescht, damit Struktur, Formate und Zebra erhalten bleiben.
Private Function PID_MPClearEmployeeRows(ByVal empKey As String, ByVal startMonth As Long) As Long
    Dim ws As Worksheet
    Dim data As Variant
    Dim inputCells As Range
    Dim monthIndex As Long
    Dim r As Long
    Dim sheetRow As Long
    Dim cleared As Long
    Dim sheetTouched As Boolean

    On Error GoTo SafeExit

    For monthIndex = startMonth To 12
        Set ws = PID_MPGetMonthSheet(monthIndex)

        If Not ws Is Nothing Then
            data = PID_MPReadIdNameBlock(ws)
            sheetTouched = False

            For r = 1 To UBound(data, 1)
                If PID_MPBuildKey(data(r, 1), data(r, 2)) = empKey Then
                    If Not sheetTouched Then
                        PID_MPTryUnprotectMonthSheet ws
                        sheetTouched = True
                    End If

                    sheetRow = PID_MPSheetRow(r)
                    
                    Set inputCells = PID_GetEmployeeInputCellsForRows(ws, sheetRow, sheetRow)
                    If Not inputCells Is Nothing Then inputCells.ClearContents
                    
                    ' Repariert Zeilen, in denen eine aeltere Version die Formeln mitgeloescht hat.
                    PID_RestoreFormulaColumnsForRows ws, sheetRow, sheetRow
                    
                    cleared = cleared + 1
                End If
            Next r

            If sheetTouched Then
                PID_EnsureMonatslohnFormulasOnSheet ws
                MarkFinanzSummaryDirtyForMonth ws
                PID_MPTryProtectMonthSheet ws
            End If
        End If
    Next monthIndex

SafeExit:
    PID_MPClearEmployeeRows = cleared
End Function


' Fuehrende Nullen wuerden in einer Standard-Zelle verloren gehen (00123 -> 123),
' deshalb fuer solche IDs vorher Textformat setzen. Andere IDs bleiben unberuehrt.
Private Sub PID_MPWriteIdCell(ByVal targetCell As Range, ByVal newId As String)
    On Error GoTo SafeExit

    If Len(newId) > 1 Then
        If Left$(newId, 1) = "0" And IsNumeric(newId) Then
            targetCell.NumberFormat = "@"
        End If
    End If

    targetCell.Value = newId

SafeExit:
End Sub


'==============================================================================
' Stunden-Override-Log
'==============================================================================

' Der Log speichert denselben Schluessel "ID|NAME" in Spalte C.
' Ohne diese Korrektur verlieren die gespeicherten Monats-Overrides ihren Bezug.
Private Function PID_MPRenameKeyInOverrideLog(ByVal oldKey As String, ByVal newKey As String) As Long
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim changed As Long

    On Error GoTo SafeExit

    Set wsLog = PID_MPGetOverrideLog()
    If wsLog Is Nothing Then Exit Function

    lastRow = PID_MPGetLogLastRow(wsLog)
    If lastRow < PID_MP_LOG_FIRST_ROW Then Exit Function

    ' Bewusst ohne Jahresfilter: der Mitarbeiter behaelt alle seine Overrides.
    For r = PID_MP_LOG_FIRST_ROW To lastRow
        If UCase$(Trim$(CStr(wsLog.Cells(r, PID_MP_LOG_KEY_COL).Value))) = oldKey Then
            wsLog.Cells(r, PID_MP_LOG_KEY_COL).Value = newKey
            changed = changed + 1
        End If
    Next r

SafeExit:
    PID_MPRenameKeyInOverrideLog = changed
End Function


' Stunden-Overrides des entfernten Mitarbeiters ab startMonth loeschen. Sonst
' wuerden sie wieder greifen, sobald derselbe Schluessel spaeter erneut auftaucht.
' Rueckwaerts loeschen, damit die Zeilenindizes gueltig bleiben.
Private Function PID_MPDeleteFromOverrideLog(ByVal empKey As String, ByVal startMonth As Long) As Long
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim workbookYear As Long
    Dim logMonth As Long
    Dim removed As Long

    On Error GoTo SafeExit

    Set wsLog = PID_MPGetOverrideLog()
    If wsLog Is Nothing Then Exit Function

    lastRow = PID_MPGetLogLastRow(wsLog)
    If lastRow < PID_MP_LOG_FIRST_ROW Then Exit Function

    workbookYear = PID_GetWorkbookYear()

    ' Jahr und Monat wie in PID_PruneHourOverrideLogForCopy erst auf Zahl pruefen.
    For r = lastRow To PID_MP_LOG_FIRST_ROW Step -1
        If UCase$(Trim$(CStr(wsLog.Cells(r, PID_MP_LOG_KEY_COL).Value))) = empKey Then
            If IsNumeric(wsLog.Cells(r, PID_MP_LOG_YEAR_COL).Value) Then
                If IsNumeric(wsLog.Cells(r, PID_MP_LOG_MONTH_COL).Value) Then
                    If CLng(wsLog.Cells(r, PID_MP_LOG_YEAR_COL).Value) = workbookYear Then
                        logMonth = CLng(wsLog.Cells(r, PID_MP_LOG_MONTH_COL).Value)

                        If logMonth >= startMonth Then
                            wsLog.Rows(r).Delete
                            removed = removed + 1
                        End If
                    End If
                End If
            End If
        End If
    Next r

SafeExit:
    PID_MPDeleteFromOverrideLog = removed
End Function


' Das Log-Blatt ist xlSheetVeryHidden, aber nicht geschuetzt - CopyData schreibt
' ebenfalls direkt darauf. Sichtbarkeit wird bewusst nicht veraendert.
Private Function PID_MPGetOverrideLog() As Worksheet
    Dim ws As Worksheet

    On Error GoTo SafeExit

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PID_MP_LOG_SHEET)
    On Error GoTo SafeExit

    Set PID_MPGetOverrideLog = ws

SafeExit:
End Function


Private Function PID_MPGetLogLastRow(ByVal wsLog As Worksheet) As Long
    On Error GoTo SafeExit

    PID_MPGetLogLastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row

SafeExit:
End Function


'==============================================================================
' Helfer
'==============================================================================

' Gleiche Logik wie PID_BuildEmployeeKey in mod_CopyData (dort Private, und das
' Bootstrap-Modul wird nicht angefasst).
Private Function PID_MPBuildKey(ByVal keyPart1 As Variant, ByVal keyPart2 As Variant) As String
    Dim s1 As String
    Dim s2 As String

    s1 = Trim$(CStr(keyPart1))
    s2 = Trim$(CStr(keyPart2))

    If s1 = "" And s2 = "" Then
        PID_MPBuildKey = ""
    Else
        PID_MPBuildKey = UCase$(s1 & "|" & s2)
    End If
End Function


Private Function PID_MPReadIdNameBlock(ByVal ws As Worksheet) As Variant
    PID_MPReadIdNameBlock = ws.Range("B" & PID_FIRST_ROW & ":C" & PID_LAST_ROW).Value
End Function


' Array-Index (1-basiert) in die echte Blattzeile umrechnen.
Private Function PID_MPSheetRow(ByVal arrayIndex As Long) As Long
    PID_MPSheetRow = arrayIndex + PID_FIRST_ROW - 1
End Function


Private Function PID_MPGetMonthSheet(ByVal monthIndex As Long) As Worksheet
    Dim monthNames As Variant
    Dim ws As Worksheet

    On Error GoTo SafeExit

    If monthIndex < 1 Or monthIndex > 12 Then Exit Function

    monthNames = PID_MonthNames()

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CStr(monthNames(monthIndex - 1)))
    On Error GoTo SafeExit

    Set PID_MPGetMonthSheet = ws

SafeExit:
End Function


Private Function PID_MPMonthName(ByVal monthIndex As Long) As String
    Dim monthNames As Variant

    On Error GoTo SafeExit

    If monthIndex < 1 Or monthIndex > 12 Then Exit Function

    monthNames = PID_MonthNames()
    PID_MPMonthName = CStr(monthNames(monthIndex - 1))

SafeExit:
End Function


Private Function PID_MPGetActiveMonthIndex() As Long
    On Error GoTo SafeExit

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function

    PID_MPGetActiveMonthIndex = PID_GetMonthIndexFromSheetName(ActiveSheet.Name)

SafeExit:
End Function


Private Sub PID_MPTryUnprotectMonthSheet(ByVal ws As Worksheet)
    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Sub
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD

SafeExit:
End Sub


Private Sub PID_MPTryProtectMonthSheet(ByVal ws As Worksheet)
    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Sub
    PID_ProtectWorkerMonthSheet ws

SafeExit:
End Sub


Private Function PID_MPFixTitle() As String
    PID_MPFixTitle = "Personal-ID korrigieren"
End Function


Private Function PID_MPRemoveTitle() As String
    PID_MPRemoveTitle = "Mitarbeiter entfernen"
End Function

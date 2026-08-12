Attribute VB_Name = "mod_AddNewKVPeriodOnTop"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit
Private Const PID_ADD_PERIOD_BUTTON_NAME As String = "btn_AddNewKVPeriodOnTop"
Private Const PID_ADD_CUSTOM_HOURS_BUTTON_NAME As String = "btn_AddCustomKVMonatsstunden"
Private Const PID_DELETE_PERIODS_BUTTON_NAME As String = "btn_DeleteKVPeriods"
Private Const PID_DELETE_CUSTOM_HOURS_BUTTON_NAME As String = "btn_DeleteCustomKVMonatsstunden"
Private Const PID_KV_CODE_COUNT As Long = 12
Private Const PID_LOHNTABELLE_BUTTON_HEIGHT As Double = 17
Private Const PID_LOHNTABELLE_BUTTON_GAP As Double = 5
Private Const PID_CUSTOM_KV_HOUR_MARKER_COL As String = "K"
Private Const PID_CUSTOM_KV_HOUR_MARKER_VALUE As String = "PID_EIGEN"


Private Function PID_KVTxtAe() As String
    PID_KVTxtAe = ChrW(228)
End Function


Private Function PID_KVTxtOe() As String
    PID_KVTxtOe = ChrW(246)
End Function


Private Function PID_KVTxtUe() As String
    PID_KVTxtUe = ChrW(252)
End Function


Private Function PID_KVTxtLoeschen() As String
    PID_KVTxtLoeschen = "l" & PID_KVTxtOe() & "schen"
End Function


Private Function PID_KVTxtGueltig() As String
    PID_KVTxtGueltig = "G" & PID_KVTxtUe() & "ltig"
End Function


Private Function PID_KVTxtBeschaeftigungsdauer() As String
    PID_KVTxtBeschaeftigungsdauer = "Besch" & PID_KVTxtAe() & "ftigungsdauer"
End Function


Private Function PID_KVTxtPruefung() As String
    PID_KVTxtPruefung = "Pr" & PID_KVTxtUe() & "fung"
End Function


Private Function PID_KVTxtSs() As String
    PID_KVTxtSs = ChrW(223)
End Function


Private Function PID_KVTxtStundeLoeschen() As String
    PID_KVTxtStundeLoeschen = "Stunde " & PID_KVTxtLoeschen()
End Function


Private Function PID_KVTxtSpaeter() As String
    PID_KVTxtSpaeter = "sp" & PID_KVTxtAe() & "ter"
End Function


Private Function PID_KVTxtWaehlen() As String
    PID_KVTxtWaehlen = "w" & PID_KVTxtAe() & "hlen"
End Function


Private Function PID_KVTxtMuessen() As String
    PID_KVTxtMuessen = "m" & PID_KVTxtUe() & "ssen"
End Function


Private Function PID_KVTxtGroesser() As String
    PID_KVTxtGroesser = "gr" & PID_KVTxtOe() & PID_KVTxtSs() & "er"
End Function


Private Function PID_KVBtnDeletePeriod() As String
    PID_KVBtnDeletePeriod = "3) Alte Periode " & PID_KVTxtLoeschen()
End Function


Private Function PID_KVBtnDeleteCustomHour() As String
    PID_KVBtnDeleteCustomHour = "4) Stunde " & PID_KVTxtLoeschen()
End Function


Public Sub AddNewKVPeriodOnTop()
    Dim wsKV As Worksheet
    Dim newPeriod As String
    Dim templatePeriod As String
    Dim templateStartYear As Long
    Dim newStartYear As Long
    Dim currentSchemaCount As Long
    Dim newSchemaCount As Long
    Dim newPeriodData As Variant
    Dim newRowCount As Long
    
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim templateFirstRow As Long
    Dim templateLastRow As Long
    Dim templateRowCount As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    
    firstDataRow = 4
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    If lastRow < firstDataRow Then
        MsgBox "Keine KV-Daten in LOHNTABELLE gefunden.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templatePeriod = GetTopKVPeriod(wsKV, firstDataRow)
    
    If templatePeriod = "" Then
        MsgBox "Es wurde kein bestehender KV-Zeitraum als Vorlage gefunden.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templateStartYear = GetStartYearFromKVPeriod(templatePeriod)
    If templateStartYear = 0 Then templateStartYear = Year(Date)
    
    newStartYear = AskForKVStartYear(templateStartYear + 1)
    If newStartYear = 0 Then GoTo CleanExit
    
    newPeriod = BuildKVPeriodName(newStartYear)
    
    If KVPeriodExists(wsKV, newPeriod, firstDataRow) Then
        MsgBox "Dieser KV-Zeitraum existiert bereits:" & vbCrLf & vbCrLf & _
               newPeriod, _
               vbInformation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templateFirstRow = FindFirstRowOfPeriod(wsKV, templatePeriod, firstDataRow)
    templateLastRow = FindLastRowOfPeriod(wsKV, templatePeriod, firstDataRow)
    
    If templateFirstRow = 0 Or templateLastRow = 0 Then
        MsgBox "Die Vorlage f" & PID_KVTxtUe() & "r den neuen KV-Zeitraum konnte nicht gefunden werden.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templateRowCount = templateLastRow - templateFirstRow + 1
    
    If templateRowCount <= 0 Then
        MsgBox "Die Vorlage enth" & PID_KVTxtAe() & "lt keine g" & PID_KVTxtUe() & "ltigen Datenzeilen.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    currentSchemaCount = GetSchemaCountFromPeriod(wsKV, templateFirstRow, templateLastRow)
    If currentSchemaCount <= 0 Then currentSchemaCount = 13
    
    newSchemaCount = AskForSchemaCount(currentSchemaCount)
    If newSchemaCount <= 0 Then GoTo CleanExit
    
    newPeriodData = BuildNewPeriodDataFromTemplate(wsKV, templateFirstRow, templateLastRow, newPeriod, newStartYear, newSchemaCount)
    If Not IsArray(newPeriodData) Then
        MsgBox "Die neue KV-Vorlage konnte nicht erstellt werden.", vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    newRowCount = UBound(newPeriodData, 1)
    If newRowCount <= 0 Then GoTo CleanExit
    
    InsertNewKVPeriodRows wsKV, firstDataRow, newPeriodData
    
    PID_NormalizeKVTableHeader wsKV
    PID_NormalizeKVWarningText wsKV
    FormatKVPeriodArea wsKV
    EnsureAddNewKVPeriodButton
    
    MarkAllKVDropdownsDirty
    MarkAllKVLohnDirty

CleanExit:
    On Error Resume Next
    wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    On Error GoTo 0
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not wsKV Is Nothing Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei AddNewKVPeriodOnTop:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Neuer KV-Zeitraum"
End Sub


Public Sub EnsureAddNewKVPeriodButton()
    PID_EnsureLOHNTABELLEButtons
End Sub


Private Function PID_LOHNTABELLEButtonsExist(ByVal wsKV As Worksheet) As Boolean
    On Error GoTo Missing
    
    If wsKV.Shapes(PID_ADD_PERIOD_BUTTON_NAME).Name = PID_ADD_PERIOD_BUTTON_NAME Then
        If wsKV.Shapes(PID_ADD_CUSTOM_HOURS_BUTTON_NAME).Name = PID_ADD_CUSTOM_HOURS_BUTTON_NAME Then
            If wsKV.Shapes(PID_DELETE_PERIODS_BUTTON_NAME).Name = PID_DELETE_PERIODS_BUTTON_NAME Then
                If wsKV.Shapes(PID_DELETE_CUSTOM_HOURS_BUTTON_NAME).Name = PID_DELETE_CUSTOM_HOURS_BUTTON_NAME Then
                    PID_LOHNTABELLEButtonsExist = True
                    Exit Function
                End If
            End If
        End If
    End If

Missing:
    PID_LOHNTABELLEButtonsExist = False
End Function


Private Function PID_LOHNTABELLEButtonsNeedRefresh(ByVal wsKV As Worksheet) As Boolean
    Dim shp As Shape
    
    On Error GoTo NeedRefresh
    
    If wsKV.Shapes(PID_DELETE_CUSTOM_HOURS_BUTTON_NAME).Height < (PID_LOHNTABELLE_BUTTON_HEIGHT - 1#) Then
        PID_LOHNTABELLEButtonsNeedRefresh = True
        Exit Function
    End If
    
    If wsKV.Shapes(PID_ADD_CUSTOM_HOURS_BUTTON_NAME).Top > _
       wsKV.Shapes(PID_ADD_PERIOD_BUTTON_NAME).Top + 2# Then
        PID_LOHNTABELLEButtonsNeedRefresh = True
        Exit Function
    End If
    
    If wsKV.Shapes(PID_ADD_PERIOD_BUTTON_NAME).Fill.ForeColor.RGB <> PID_StyleColorNavy() Then
        PID_LOHNTABELLEButtonsNeedRefresh = True
        Exit Function
    End If
    
    If PID_CountLOHNTABELLEToolbarShapes(wsKV) <> 4 Then
        PID_LOHNTABELLEButtonsNeedRefresh = True
        Exit Function
    End If
    
    For Each shp In wsKV.Shapes
        If PID_IsLOHNTABELLEToolbarShape(shp) Then
            If PID_LOHNTABELLEToolbarShapeIsStale(wsKV, shp) Then
                PID_LOHNTABELLEButtonsNeedRefresh = True
                Exit Function
            End If
        End If
    Next shp
    
    Exit Function

NeedRefresh:
    PID_LOHNTABELLEButtonsNeedRefresh = True
End Function


Private Function PID_IsLOHNTABELLEToolbarShape(ByVal shp As Shape) As Boolean
    Dim actionText As String
    Dim labelText As String
    
    On Error GoTo SafeExit
    
    If shp Is Nothing Then Exit Function
    
    Select Case shp.Name
        Case PID_ADD_PERIOD_BUTTON_NAME, PID_ADD_CUSTOM_HOURS_BUTTON_NAME, _
             PID_DELETE_PERIODS_BUTTON_NAME, PID_DELETE_CUSTOM_HOURS_BUTTON_NAME
            PID_IsLOHNTABELLEToolbarShape = True
            Exit Function
    End Select
    
    actionText = LCase$(Trim$(Replace$(shp.OnAction, "'", "")))
    
    If InStr(1, actionText, "addnewkvperiodontop", vbTextCompare) > 0 Then
        PID_IsLOHNTABELLEToolbarShape = True
        Exit Function
    End If
    If InStr(1, actionText, "addcustomkvmonatsstunden", vbTextCompare) > 0 Then
        PID_IsLOHNTABELLEToolbarShape = True
        Exit Function
    End If
    If InStr(1, actionText, "deleteselectedkvperiods", vbTextCompare) > 0 Then
        PID_IsLOHNTABELLEToolbarShape = True
        Exit Function
    End If
    If InStr(1, actionText, "deletecustomkvmonatsstunden", vbTextCompare) > 0 Then
        PID_IsLOHNTABELLEToolbarShape = True
        Exit Function
    End If
    If InStr(1, actionText, "showlohntabellebuttonhelp", vbTextCompare) > 0 Then
        PID_IsLOHNTABELLEToolbarShape = True
        Exit Function
    End If
    
    If shp.Type <> msoAutoShape Then Exit Function
    
    labelText = Trim$(shp.TextFrame.Characters.Text)
    
    If labelText = "Hilfe" Then
        PID_IsLOHNTABELLEToolbarShape = True
        Exit Function
    End If
    
    If Left$(labelText, 2) = "4)" Then
        PID_IsLOHNTABELLEToolbarShape = True
        Exit Function
    End If
    
    If Left$(labelText, 2) = "1)" Or Left$(labelText, 2) = "2)" Or Left$(labelText, 2) = "3)" Then
        PID_IsLOHNTABELLEToolbarShape = True
    End If
    
SafeExit:
End Function


Private Sub PID_DeleteAllLOHNTABELLEToolbarShapes(ByVal wsKV As Worksheet)
    Dim shp As Shape
    Dim i As Long
    
    If wsKV Is Nothing Then Exit Sub
    
    For i = wsKV.Shapes.Count To 1 Step -1
        Set shp = wsKV.Shapes(i)
        If PID_IsLOHNTABELLEToolbarShape(shp) Then
            shp.Delete
        End If
    Next i
End Sub


Private Function PID_LOHNTABELLEToolbarShapeIsStale(ByVal wsKV As Worksheet, ByVal shp As Shape) As Boolean
    Dim row2Bottom As Double
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Or shp Is Nothing Then Exit Function
    
    row2Bottom = wsKV.Range("A2").Top + wsKV.Rows(2).RowHeight
    
    If shp.Top > row2Bottom + 1# Then
        PID_LOHNTABELLEToolbarShapeIsStale = True
    End If
    
SafeExit:
End Function


Private Function PID_CountLOHNTABELLEToolbarShapes(ByVal wsKV As Worksheet) As Long
    Dim shp As Shape
    
    If wsKV Is Nothing Then Exit Function
    
    For Each shp In wsKV.Shapes
        If PID_IsLOHNTABELLEToolbarShape(shp) Then
            PID_CountLOHNTABELLEToolbarShapes = PID_CountLOHNTABELLEToolbarShapes + 1
        End If
    Next shp
End Function


' Einmalig: Zeile mit eigener Stunde markieren (z.B. vor K-Marker-Update eingefuegt).
Public Sub PID_MarkSelectedLOHNTABELLECustomHour()
    Dim wsKV As Worksheet
    Dim targetRow As Long
    
    On Error GoTo SafeExit
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Bitte zuerst die Zeile mit der eigenen Stunde anklicken (Feld ""Monatsstunden"").", _
               vbExclamation, "Eigene Stunde markieren"
        Exit Sub
    End If
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    targetRow = Selection.Row
    
    If targetRow < PID_GetKVTableFirstDataRow(wsKV) Then
        MsgBox "Bitte eine Datenzeile in LOHNTABELLE " & PID_KVTxtWaehlen() & ".", _
               vbExclamation, "Eigene Stunde markieren"
        Exit Sub
    End If
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    PID_MarkCustomKVHourRow wsKV, targetRow
    
    MsgBox "Zeile " & CStr(targetRow) & " ist als eigene Stunde markiert." & vbCrLf & _
           "Jetzt ""4) Stunde " & PID_KVTxtLoeschen() & """ verwenden.", _
           vbInformation, "Eigene Stunde markieren"

SafeExit:
    On Error Resume Next
    If Not wsKV Is Nothing Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Public Sub DeleteCustomKVMonatsstunden()
    Dim wsKV As Worksheet
    Dim firstDataRow As Long
    Dim selectedPeriod As String
    Dim selectedKVCode As String
    Dim selectedHours As Double
    Dim blockFirst As Long
    Dim blockLast As Long
    Dim schemaRowCount As Long
    Dim blockRowCount As Long
    Dim deleteRow As Long
    Dim usageCount As Long
    Dim confirmText As String
    Dim answer As VbMsgBoxResult
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim stateCaptured As Boolean
    Dim savedScrollRow As Long
    Dim savedScrollCol As Long
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    firstDataRow = PID_GetKVTableFirstDataRow(wsKV)
    
    If PID_GetKVTableLastRow(wsKV, firstDataRow) < firstDataRow Then
        MsgBox "Keine KV-Daten in LOHNTABELLE gefunden.", _
               vbExclamation, PID_KVTxtStundeLoeschen()
        Exit Sub
    End If
    
    selectedPeriod = AskForKVPeriodSelection( _
        wsKV, firstDataRow, _
        PID_KVTxtStundeLoeschen(), _
        "Schritt 1 von 4 - Welcher Zeitraum?", _
        "Nummer eingeben und OK klicken. Meist ist das ""1"" (oben).")
    If selectedPeriod = "" Then Exit Sub
    
    selectedKVCode = AskForKVCodeSelection(PID_KVTxtStundeLoeschen(), _
        "Schritt 2 von 4 - Welche Gruppe?", _
        "Beispiel: 1 = BG1_Basis (Standard-Vertrag).")
    If selectedKVCode = "" Then Exit Sub
    
    If Not PID_GetKVCodeBlockBounds(wsKV, selectedPeriod, selectedKVCode, firstDataRow, blockFirst, blockLast) Then
        MsgBox "Der KV-Code """ & selectedKVCode & """ wurde im Zeitraum """ & selectedPeriod & """ nicht gefunden.", _
               vbExclamation, PID_KVTxtStundeLoeschen()
        Exit Sub
    End If
    
    blockRowCount = blockLast - blockFirst + 1
    
    If Not PID_BlockHasDeletableCustomHours(wsKV, selectedPeriod, selectedKVCode, _
            firstDataRow, blockFirst, blockLast, blockRowCount) Then
        MsgBox "In """ & selectedKVCode & """ gibt es keine zus" & PID_KVTxtAe() & "tzliche eigene Stunde." & vbCrLf & vbCrLf & _
               "Nur Stunden, die mit ""2) Eigene Stunden"" hinzugef" & PID_KVTxtUe() & "gt wurden, k" & PID_KVTxtOe() & "nnen entfernt werden." & vbCrLf & _
               "Zeile mit Monatsstunden anklicken, Alt+F8 -> PID_MarkSelectedLOHNTABELLECustomHour, danach erneut " & PID_KVTxtLoeschen() & "." & vbCrLf & vbCrLf & _
               "Standard-Vertr" & PID_KVTxtAe() & "ge bleiben unver" & PID_KVTxtAe() & "ndert.", _
               vbInformation, PID_KVTxtStundeLoeschen()
        Exit Sub
    End If
    
    If Not AskForKVHoursInBlockSelection(wsKV, blockFirst, blockLast, selectedHours, _
            PID_KVTxtStundeLoeschen(), "Schritt 3 von 4 - Welche Stunden?", _
            PID_CollectDeletableHoursInBlock(wsKV, selectedPeriod, selectedKVCode, _
                firstDataRow, blockFirst, blockLast)) Then
        Exit Sub
    End If
    
    deleteRow = PID_FindDeletableHourRowInBlock(wsKV, blockFirst, blockLast, selectedHours)
    If deleteRow <= 0 Then
        MsgBox "Die gew" & PID_KVTxtAe() & "hlten Stunden konnten in der Gruppe nicht gefunden werden.", _
               vbExclamation, PID_KVTxtStundeLoeschen()
        Exit Sub
    End If
    
    usageCount = PID_CountMonthSheetUsageForKVHour(selectedKVCode, selectedHours)
    
    confirmText = "Wirklich " & PID_KVTxtLoeschen() & "?" & vbCrLf & vbCrLf & _
                  "Zeitraum: " & selectedPeriod & vbCrLf & _
                  "Gruppe: " & selectedKVCode & vbCrLf & _
                  "Stunden: " & PID_FormatHoursText(selectedHours) & vbCrLf
    
    If usageCount > 0 Then
        confirmText = confirmText & vbCrLf & _
            "Hinweis: Diese Stunden sind auf " & CStr(usageCount) & _
            " Monatszeile(n) noch eingetragen." & vbCrLf & _
            "Nach dem L" & PID_KVTxtOe() & "schen dort ggf. F anpassen." & vbCrLf
    End If
    
    confirmText = confirmText & vbCrLf & _
                  "NEIN = abbrechen." & vbCrLf & _
                  "JA = Zeile in LOHNTABELLE entfernen."
    
    answer = MsgBox(confirmText, vbCritical + vbYesNo, "Schritt 4 von 4 - Letzte R" & PID_KVTxtUe() & "ckfrage")
    If answer <> vbYes Then
        MsgBox PID_KVTxtLoeschen() & " abgebrochen. Es wurde nichts ge" & PID_KVTxtAe() & "ndert.", _
               vbInformation, PID_KVTxtStundeLoeschen()
        Exit Sub
    End If
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    stateCaptured = True
    
    If Not ActiveWindow Is Nothing Then
        savedScrollRow = ActiveWindow.ScrollRow
        savedScrollCol = ActiveWindow.ScrollColumn
    End If
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    wsKV.Rows(deleteRow).Delete Shift:=xlShiftUp
    
    FormatKVPeriodArea wsKV
    PID_EnsureLOHNTABELLEButtons
    
    MarkKVDropdownDirtyForKVCode selectedKVCode
    MarkAllKVLohnDirty
    
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
    If stateCaptured Then
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        PID_RestoreLOHNTABELLEView wsKV, savedScrollRow, savedScrollCol, deleteRow, "G"
    End If
    On Error GoTo 0
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not wsKV Is Nothing Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    If stateCaptured Then
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        PID_RestoreLOHNTABELLEView wsKV, savedScrollRow, savedScrollCol, deleteRow, "G"
    End If
    
    MsgBox "Fehler bei DeleteCustomKVMonatsstunden:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, PID_KVTxtStundeLoeschen()
End Sub


Public Sub DeleteSelectedKVPeriods()
    Dim wsKV As Worksheet
    Dim firstDataRow As Long
    Dim allPeriods As Collection
    Dim periodToDelete As String
    Dim confirmText As String
    Dim answer As VbMsgBoxResult
    Dim periodFirstRow As Long
    Dim periodLastRow As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim stateCaptured As Boolean
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    firstDataRow = PID_GetKVTableFirstDataRow(wsKV)
    
    Set allPeriods = PID_CollectKVPeriods(wsKV, firstDataRow)
    If allPeriods Is Nothing Or allPeriods.Count = 0 Then
        MsgBox "Keine KV-Zeitr" & PID_KVTxtAe() & "ume gefunden.", vbExclamation, "Alte Periode " & PID_KVTxtLoeschen()
        Exit Sub
    End If
    
    If allPeriods.Count <= 1 Then
        MsgBox "Es ist nur ein KV-Zeitraum vorhanden." & vbCrLf & vbCrLf & _
               "Mindestens ein Zeitraum muss erhalten bleiben. L" & PID_KVTxtOe() & "schen ist nicht m" & PID_KVTxtOe() & "glich.", _
               vbInformation, "Alte Periode " & PID_KVTxtLoeschen()
        Exit Sub
    End If
    
    periodToDelete = AskForKVPeriodSelection( _
        wsKV, firstDataRow, _
        "Alte Periode " & PID_KVTxtLoeschen(), _
        "Schritt 1 von 2 - Welche Periode?", _
        "Nur EINE Nummer eingeben (z.B. 2)." & vbCrLf & _
        "Nur sehr alte Perioden " & PID_KVTxtLoeschen() & ", die niemand mehr braucht." & vbCrLf & _
        "Abbrechen oder leer lassen = nichts " & PID_KVTxtAe() & "ndern.")
    
    If periodToDelete = "" Then Exit Sub
    
    confirmText = "Wirklich " & PID_KVTxtLoeschen() & "?" & vbCrLf & vbCrLf & _
                  periodToDelete & vbCrLf & vbCrLf & _
                  "NEIN = abbrechen (empfohlen bei Unsicherheit)." & vbCrLf & _
                  "JA = endg" & PID_KVTxtUe() & "ltig " & PID_KVTxtLoeschen() & "."
    
    answer = MsgBox(confirmText, vbCritical + vbYesNo, "Schritt 2 von 2 - Letzte R" & PID_KVTxtUe() & "ckfrage")
    If answer <> vbYes Then
        MsgBox "L" & PID_KVTxtOe() & "schen abgebrochen. Es wurde nichts ge" & PID_KVTxtAe() & "ndert.", vbInformation, "Alte Periode " & PID_KVTxtLoeschen()
        Exit Sub
    End If
    
    periodFirstRow = FindFirstRowOfPeriod(wsKV, periodToDelete, firstDataRow)
    periodLastRow = FindLastRowOfPeriod(wsKV, periodToDelete, firstDataRow)
    
    If periodFirstRow <= 0 Or periodLastRow < periodFirstRow Then
        MsgBox "Der gew" & PID_KVTxtAe() & "hlte Zeitraum konnte nicht gefunden werden.", vbExclamation, "Alte Periode " & PID_KVTxtLoeschen()
        Exit Sub
    End If
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    stateCaptured = True
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    wsKV.Rows(periodFirstRow & ":" & periodLastRow).Delete Shift:=xlShiftUp
    
    FormatKVPeriodArea wsKV
    PID_EnsureLOHNTABELLEButtons
    
    MarkAllKVDropdownsDirty
    MarkAllKVLohnDirty

    GoTo CleanExit

CleanExit:
    On Error Resume Next
    wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
    If stateCaptured Then
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
    End If
    On Error GoTo 0
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not wsKV Is Nothing Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    If stateCaptured Then
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
    End If
    
    MsgBox "Fehler beim L" & PID_KVTxtOe() & "schen:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Alte Periode " & PID_KVTxtLoeschen()
End Sub


Public Sub AddCustomKVMonatsstunden()
    Dim wsKV As Worksheet
    Dim firstDataRow As Long
    Dim selectedPeriod As String
    Dim selectedKVCode As String
    Dim newHours As Double
    Dim newLohn As Variant
    Dim hasLohn As Boolean
    Dim blockFirst As Long
    Dim blockLast As Long
    Dim insertRow As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim stateCaptured As Boolean
    Dim savedScrollRow As Long
    Dim savedScrollCol As Long
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    firstDataRow = PID_GetKVTableFirstDataRow(wsKV)
    
    If PID_GetKVTableLastRow(wsKV, firstDataRow) < firstDataRow Then
        MsgBox "Keine KV-Daten in LOHNTABELLE gefunden.", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Sub
    End If
    
    selectedPeriod = AskForKVPeriodSelection( _
        wsKV, firstDataRow, _
        "Eigene Stunden", _
        "Schritt 1 von 4 - Welcher Zeitraum?", _
        "Nummer eingeben und OK klicken. Meist ist das ""1"" (oben).")
    If selectedPeriod = "" Then Exit Sub
    
    selectedKVCode = AskForKVCodeSelection("Eigene Stunden", "Schritt 2 von 4 - Welche Gruppe?", _
        "Beispiel: 1 = BG1_Basis (Standard-Vertrag).")
    If selectedKVCode = "" Then Exit Sub
    
    If Not AskForCustomMonatsstunden(newHours) Then Exit Sub
    
    hasLohn = AskForOptionalMonatslohn(newLohn)
    If Not hasLohn Then newLohn = Empty
    
    If Not PID_GetKVCodeBlockBounds(wsKV, selectedPeriod, selectedKVCode, firstDataRow, blockFirst, blockLast) Then
        MsgBox "Der KV-Code """ & selectedKVCode & """ wurde im Zeitraum """ & selectedPeriod & """ nicht gefunden.", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Sub
    End If
    
    If PID_KVCodeBlockHasHours(wsKV, blockFirst, blockLast, newHours) Then
        MsgBox "Die Monatsstunden " & PID_FormatHoursText(newHours) & " existieren bereits in " & selectedKVCode & ".", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Sub
    End If
    
    insertRow = PID_FindSortedInsertRowInBlock(wsKV, blockFirst, blockLast, newHours)
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    stateCaptured = True
    
    If Not ActiveWindow Is Nothing Then
        savedScrollRow = ActiveWindow.ScrollRow
        savedScrollCol = ActiveWindow.ScrollColumn
    End If
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    wsKV.Rows(insertRow).Insert Shift:=xlDown
    
    PID_WriteCustomKVRow wsKV, insertRow, selectedPeriod, selectedKVCode, firstDataRow, newHours, newLohn, hasLohn
    
    FormatKVPeriodArea wsKV
    PID_MarkCustomKVHourRow wsKV, insertRow
    PID_EnsureLOHNTABELLEButtons
    
    MarkKVDropdownDirtyForKVCode selectedKVCode
    MarkAllKVLohnDirty
    
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
    If stateCaptured Then
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        PID_RestoreLOHNTABELLEView wsKV, savedScrollRow, savedScrollCol, insertRow, "G"
    End If
    On Error GoTo 0
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not wsKV Is Nothing Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    If stateCaptured Then
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        PID_RestoreLOHNTABELLEView wsKV, savedScrollRow, savedScrollCol, insertRow, "G"
    End If
    
    MsgBox "Fehler bei AddCustomKVMonatsstunden:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Eigene Stunden"
End Sub


Private Sub PID_EnsureLOHNTABELLEButtons(Optional ByVal forceRecreate As Boolean = False)
    Dim wsKV As Worksheet
    Dim btn As Shape
    Dim wasProtected As Boolean
    Dim areaLeft As Double
    Dim areaTop As Double
    Dim areaWidth As Double
    Dim areaHeight As Double
    Dim buttonLeft As Double
    Dim buttonWidth As Double
    Dim buttonTop As Double
    Dim buttonHeight As Double
    Dim buttonGap As Double
    Dim i As Long
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    If wsKV Is Nothing Then Exit Sub
    
    If Not forceRecreate Then
        If PID_LOHNTABELLEButtonsExist(wsKV) Then
            If Not PID_LOHNTABELLEButtonsNeedRefresh(wsKV) Then Exit Sub
        End If
    End If
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    PID_DeleteAllLOHNTABELLEToolbarShapes wsKV
    On Error GoTo SafeExit
    
    PID_ConfigureLOHNTABELLEHeaderLayout wsKV
    PID_NormalizeKVTableHeader wsKV
    
    areaLeft = wsKV.Range("A2").Left
    areaTop = wsKV.Range("A2").Top
    areaWidth = wsKV.Range("A2:J2").Width
    areaHeight = wsKV.Rows(2).RowHeight
    
    buttonGap = PID_LOHNTABELLE_BUTTON_GAP
    buttonHeight = PID_LOHNTABELLE_BUTTON_HEIGHT
    buttonWidth = (areaWidth - (5# * buttonGap)) / 4#
    buttonTop = areaTop + ((areaHeight - buttonHeight) / 2#)
    buttonLeft = areaLeft + buttonGap
    
    For i = 0 To 3
        Set btn = wsKV.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                       Left:=buttonLeft + (i * (buttonWidth + buttonGap)), _
                                       Top:=buttonTop, _
                                       Width:=buttonWidth, _
                                       Height:=buttonHeight)
        
        Select Case i
            Case 0
                btn.Name = PID_ADD_PERIOD_BUTTON_NAME
                btn.TextFrame.Characters.Text = "1) Neue Periode"
                btn.OnAction = "AddNewKVPeriodOnTop"
                PID_StyleApplyToolbarButton btn, PID_StyleColorNavy(), PID_StyleColorBtnPrimaryLine(), RGB(255, 255, 255)
            Case 1
                btn.Name = PID_ADD_CUSTOM_HOURS_BUTTON_NAME
                btn.TextFrame.Characters.Text = "2) Eigene Stunden"
                btn.OnAction = "AddCustomKVMonatsstunden"
                PID_StyleApplyToolbarButton btn, PID_StyleColorHeaderBg(), PID_StyleColorNavy(), PID_StyleColorNavy()
            Case 2
                btn.Name = PID_DELETE_PERIODS_BUTTON_NAME
                btn.TextFrame.Characters.Text = PID_KVBtnDeletePeriod()
                btn.OnAction = "DeleteSelectedKVPeriods"
                PID_StyleApplyToolbarButton btn, PID_StyleColorAccent(), PID_StyleColorNavy(), PID_StyleColorNavy()
            Case 3
                btn.Name = PID_DELETE_CUSTOM_HOURS_BUTTON_NAME
                btn.TextFrame.Characters.Text = PID_KVBtnDeleteCustomHour()
                btn.OnAction = "DeleteCustomKVMonatsstunden"
                PID_StyleApplyToolbarButton btn, PID_StyleColorZebra(), PID_StyleColorNavy(), PID_StyleColorNavy()
        End Select
    Next i
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Sub PID_ConfigureLOHNTABELLEHeaderLayout(ByVal wsKV As Worksheet)
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    On Error Resume Next
    wsKV.Range("A1:J2").UnMerge
    On Error GoTo SafeExit
    
    wsKV.Range("A1:J1").Merge
    PID_StyleApplyTitleBand wsKV.Range("A1")
    wsKV.Rows(1).RowHeight = PID_STYLE_COMPACT_BLOCK_TITLE_HEIGHT
    
    wsKV.Range("A2:J2").UnMerge
    wsKV.Range("A2:J2").ClearContents
    wsKV.Range("A2:J2").Interior.Pattern = xlNone
    
    wsKV.Rows(2).RowHeight = 24
    
SafeExit:
End Sub


Public Sub FixLOHNTABELLE_HeaderText()
    FixLOHNTABELLE_HeaderTextIfNeeded True
End Sub


Public Sub FixLOHNTABELLE_HeaderTextIfNeeded(Optional ByVal forceFormulaRepair As Boolean = False)
    Dim wsKV As Worksheet
    Dim wasProtected As Boolean
    Dim lastRow As Long
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    PID_NormalizeKVWarningText wsKV
    PID_NormalizeKVTableHeader wsKV
    
    If forceFormulaRepair Or PID_KVStatusFormulasNeedRepair(wsKV) Then
        PID_EnsureKVStatusFormulas wsKV
    End If
    
    If wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row >= 4 Then
        lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
        PID_ConfigureKVTableColumnWidths wsKV
        PID_ApplyKVTableRowHeights wsKV, 4, lastRow
        PID_RefreshKVPeriodTitleRowStyles wsKV, 4, lastRow
        PID_FormatKVRowTypography wsKV, 4, lastRow
        PID_ApplyKVVisualGrouping wsKV, 4, lastRow
    End If
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Function PID_KVStatusFormulasNeedRepair(ByVal wsKV As Worksheet) As Boolean
    Dim r As Long
    Dim lastRow As Long
    Dim checkedRows As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Function
    
    For r = 4 To lastRow
        If Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "" Then
            If Not wsKV.Cells(r, "I").HasFormula Then
                PID_KVStatusFormulasNeedRepair = True
                Exit Function
            End If
            
            checkedRows = checkedRows + 1
            If checkedRows >= 3 Then Exit Function
        End If
    Next r
    
SafeExit:
End Function


Public Sub FixLOHNTABELLE_StatusFormulas()
    Dim wsKV As Worksheet
    Dim wasProtected As Boolean
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    FormatKVPeriodArea wsKV
    
    MsgBox "Status- und " & PID_KVTxtPruefung() & "sformeln in LOHNTABELLE wurden wiederhergestellt.", _
           vbInformation, "LOHNTABELLE"
    
CleanExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei FixLOHNTABELLE_StatusFormulas:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "LOHNTABELLE"
    Resume CleanExit
End Sub


Public Sub RebuildLOHNTABELLE()
    Dim wsKV As Worksheet
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim cleanupLastRow As Long
    Dim keepPeriod As String
    Dim periodFirstRow As Long
    Dim periodLastRow As Long
    Dim periodRowCount As Long
    Dim periodData As Variant
    Dim answer As VbMsgBoxResult
    Dim wasProtected As Boolean
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    If wsKV Is Nothing Then Exit Sub
    
    firstDataRow = 4
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    cleanupLastRow = PID_GetSheetCleanupLastRow(wsKV, lastRow)
    
    If lastRow < firstDataRow Then
        MsgBox "Keine Daten in LOHNTABELLE gefunden.", vbExclamation, "LOHNTABELLE neu aufbauen"
        Exit Sub
    End If
    
    keepPeriod = GetBottomKVPeriod(wsKV, firstDataRow)
    If keepPeriod = "" Then
        MsgBox "Kein g" & PID_KVTxtUe() & "ltiger KV-Zeitraum in Spalte A gefunden.", vbExclamation, "LOHNTABELLE neu aufbauen"
        Exit Sub
    End If
    
    periodFirstRow = FindFirstRowOfPeriod(wsKV, keepPeriod, firstDataRow)
    periodLastRow = FindLastRowOfPeriod(wsKV, keepPeriod, firstDataRow)
    
    If periodFirstRow = 0 Or periodLastRow = 0 Then
        MsgBox "Der unterste KV-Zeitraum konnte nicht gelesen werden.", vbExclamation, "LOHNTABELLE neu aufbauen"
        Exit Sub
    End If
    
    periodRowCount = periodLastRow - periodFirstRow + 1
    If periodRowCount <= 0 Then Exit Sub
    
    answer = MsgBox( _
        "LOHNTABELLE wird neu aufgebaut." & vbCrLf & vbCrLf & _
        "Behalten wird nur der unterste Zeitraum:" & vbCrLf & _
        keepPeriod & vbCrLf & vbCrLf & _
        "Alle weiteren Test-Zeitr" & PID_KVTxtAe() & "ume werden gel" & PID_KVTxtOe() & "scht." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        "LOHNTABELLE neu aufbauen" _
    )
    
    If answer <> vbYes Then Exit Sub
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    wasProtected = wsKV.ProtectContents
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    periodData = wsKV.Range("A" & periodFirstRow & ":I" & periodLastRow).Value
    periodData = PID_FilterValidKVRows(periodData, keepPeriod)
    periodData = PID_EnsureBG1Basis173Row(periodData, keepPeriod)
    periodRowCount = UBound(periodData, 1)
    
    PID_ClearKVDataArea wsKV, firstDataRow, cleanupLastRow
    wsKV.Range("A" & firstDataRow + 1 & ":I" & firstDataRow + periodRowCount).Value = periodData
    PID_WriteKVPeriodTitleRow wsKV, firstDataRow, keepPeriod, periodData(1, 2), periodData(1, 3)
    PID_ClearTrailingKVArea wsKV, firstDataRow + periodRowCount + 1, cleanupLastRow
    PID_NormalizeKVWarningText wsKV
    PID_NormalizeKVTableHeader wsKV
    FormatKVPeriodArea wsKV
    PID_ClearTrailingKVArea wsKV, firstDataRow + periodRowCount + 1, cleanupLastRow
    MarkAllKVDropdownsDirty
    MarkAllKVLohnDirty
    
    On Error Resume Next
    PID_ResetHourOverrideLog
    On Error GoTo CleanFail
    
    EnsureAddNewKVPeriodButton
    
    MsgBox "LOHNTABELLE wurde neu aufgebaut. Zeitraum aktiv: " & keepPeriod, _
           vbInformation, "LOHNTABELLE neu aufgebaut"
    
CleanExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    On Error GoTo 0
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei RebuildLOHNTABELLE:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "LOHNTABELLE neu aufbauen"
    Resume CleanExit
End Sub


Public Sub RestoreLOHNTABELLEBase2025_2026()
    Dim wsKV As Worksheet
    Dim targetPeriod As String
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim cleanupLastRow As Long
    Dim periodFirstRow As Long
    Dim periodLastRow As Long
    Dim periodRowCount As Long
    Dim periodData As Variant
    Dim answer As VbMsgBoxResult
    Dim wasProtected As Boolean
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    If wsKV Is Nothing Then Exit Sub
    
    targetPeriod = "KV 2025/2026"
    firstDataRow = 4
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    cleanupLastRow = PID_GetSheetCleanupLastRow(wsKV, lastRow)
    
    periodFirstRow = FindFirstRowOfPeriod(wsKV, targetPeriod, firstDataRow)
    periodLastRow = FindLastRowOfPeriod(wsKV, targetPeriod, firstDataRow)
    
    If periodFirstRow = 0 Or periodLastRow = 0 Then
        MsgBox "Der Basiszeitraum '" & targetPeriod & "' wurde nicht gefunden.", _
               vbExclamation, "Basis wiederherstellen"
        Exit Sub
    End If
    
    periodRowCount = periodLastRow - periodFirstRow + 1
    If periodRowCount <= 0 Then Exit Sub
    
    answer = MsgBox( _
        "LOHNTABELLE wird auf den Basiszeitraum zur" & PID_KVTxtUe() & "ckgesetzt:" & vbCrLf & vbCrLf & _
        targetPeriod & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        "Basis wiederherstellen" _
    )
    
    If answer <> vbYes Then Exit Sub
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    wasProtected = wsKV.ProtectContents
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    periodData = wsKV.Range("A" & periodFirstRow & ":I" & periodLastRow).Value
    periodData = PID_FilterValidKVRows(periodData, targetPeriod)
    periodData = PID_EnsureBG1Basis173Row(periodData, targetPeriod)
    periodRowCount = UBound(periodData, 1)
    
    PID_ClearKVDataArea wsKV, firstDataRow, cleanupLastRow
    wsKV.Range("A" & firstDataRow + 1 & ":I" & firstDataRow + periodRowCount).Value = periodData
    PID_WriteKVPeriodTitleRow wsKV, firstDataRow, targetPeriod, periodData(1, 2), periodData(1, 3)
    PID_ClearTrailingKVArea wsKV, firstDataRow + periodRowCount + 1, cleanupLastRow
    
    PID_NormalizeKVWarningText wsKV
    PID_NormalizeKVTableHeader wsKV
    FormatKVPeriodArea wsKV
    PID_ClearTrailingKVArea wsKV, firstDataRow + periodRowCount + 1, cleanupLastRow
    EnsureAddNewKVPeriodButton
    MarkAllKVLohnDirty
    
    On Error Resume Next
    PID_ResetHourOverrideLog
    On Error GoTo CleanFail
    
    MsgBox "Basiszeitraum wurde wiederhergestellt: " & targetPeriod, _
           vbInformation, "Basis wiederherstellen"
    
CleanExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    On Error GoTo 0
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei RestoreLOHNTABELLEBase2025_2026:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Basis wiederherstellen"
    Resume CleanExit
End Sub


Private Function PID_FilterValidKVRows(ByVal sourceData As Variant, ByVal periodName As String) As Variant
    Dim r As Long
    Dim outRow As Long
    Dim validCount As Long
    Dim resultData As Variant
    
    On Error GoTo Fallback
    
    validCount = 0
    
    For r = 1 To UBound(sourceData, 1)
        If PID_IsValidKVDataRow(sourceData, r) Then
            validCount = validCount + 1
        End If
    Next r
    
    If validCount = 0 Then GoTo Fallback
    
    ReDim resultData(1 To validCount, 1 To 9)
    outRow = 0
    
    For r = 1 To UBound(sourceData, 1)
        If PID_IsValidKVDataRow(sourceData, r) Then
            outRow = outRow + 1
            
            resultData(outRow, 1) = periodName
            resultData(outRow, 2) = sourceData(r, 2)
            resultData(outRow, 3) = sourceData(r, 3)
            resultData(outRow, 4) = PID_GetNearestTextValue(sourceData, r, 4)
            resultData(outRow, 5) = PID_GetNearestTextValue(sourceData, r, 5)
            resultData(outRow, 6) = PID_GetNearestTextValue(sourceData, r, 6)
            resultData(outRow, 7) = sourceData(r, 7)
            resultData(outRow, 8) = sourceData(r, 8)
            resultData(outRow, 9) = ""
        End If
    Next r
    
    PID_FilterValidKVRows = resultData
    Exit Function
    
Fallback:
    PID_FilterValidKVRows = sourceData
End Function


Private Function PID_IsValidKVDataRow(ByVal sourceData As Variant, ByVal rowIndex As Long) As Boolean
    If rowIndex < 1 Or rowIndex > UBound(sourceData, 1) Then Exit Function
    
    If Trim$(CStr(PID_GetNearestTextValue(sourceData, rowIndex, 4))) = "" Then Exit Function
    If Trim$(CStr(PID_GetNearestTextValue(sourceData, rowIndex, 5))) = "" Then Exit Function
    If Trim$(CStr(PID_GetNearestTextValue(sourceData, rowIndex, 6))) = "" Then Exit Function
    
    ' Leere Restzeilen aus laengeren Testperioden ausschliessen.
    If Not PID_ArrayRowHasHoursOrWage(sourceData, rowIndex) Then Exit Function
    
    PID_IsValidKVDataRow = True
End Function


Private Function PID_ArrayRowHasHoursOrWage(ByVal sourceData As Variant, ByVal rowIndex As Long) As Boolean
    Dim hoursValue As Double
    Dim wageValue As Double
    
    If PID_TryReadDouble(sourceData(rowIndex, 7), hoursValue) Then
        If hoursValue > 0# Then
            PID_ArrayRowHasHoursOrWage = True
            Exit Function
        End If
    End If
    
    If PID_TryReadDouble(sourceData(rowIndex, 8), wageValue) Then
        If wageValue > 0# Then
            PID_ArrayRowHasHoursOrWage = True
        End If
    End If
End Function


Private Function PID_RowHasKVTableContent(ByVal wsKV As Worksheet, ByVal rowNumber As Long) As Boolean
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    If rowNumber < 1 Then Exit Function
    
    If wsKV.Range("A" & rowNumber).MergeCells Then
        PID_RowHasKVTableContent = True
        Exit Function
    End If
    
    If Trim$(CStr(wsKV.Cells(rowNumber, "D").Value)) <> "" Then
        PID_RowHasKVTableContent = True
        Exit Function
    End If
    
    If Trim$(CStr(wsKV.Cells(rowNumber, "G").Value)) <> "" Then
        PID_RowHasKVTableContent = True
        Exit Function
    End If
    
    If Trim$(CStr(wsKV.Cells(rowNumber, "H").Value)) <> "" Then
        PID_RowHasKVTableContent = True
    End If
    
SafeExit:
End Function


Private Function PID_GetKVTableLastRow(ByVal wsKV As Worksheet, ByVal firstDataRow As Long) As Long
    Dim r As Long
    Dim sheetLastRow As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    
    sheetLastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    If sheetLastRow < firstDataRow Then
        PID_GetKVTableLastRow = firstDataRow
        Exit Function
    End If
    
    For r = sheetLastRow To firstDataRow Step -1
        If PID_RowHasKVTableContent(wsKV, r) Then
            PID_GetKVTableLastRow = r
            Exit Function
        End If
    Next r
    
    PID_GetKVTableLastRow = firstDataRow
    
SafeExit:
End Function


Private Function PID_EnsureBG1Basis173Row(ByVal sourceData As Variant, ByVal periodName As String) As Variant
    Dim r As Long
    Dim outRow As Long
    Dim has173 As Boolean
    Dim first151Row As Long
    Dim hoursValue As Double
    Dim resultData As Variant
    
    On Error GoTo Fallback
    
    For r = 1 To UBound(sourceData, 1)
        If UCase$(Trim$(CStr(sourceData(r, 4)))) = "BG1_BASIS" Then
            If PID_TryReadDouble(sourceData(r, 7), hoursValue) Then
                If Abs(hoursValue - 173#) < 0.01 Then
                    has173 = True
                End If
                
                If first151Row = 0 Then
                    If Abs(hoursValue - 151.38) < 0.01 Then
                        first151Row = r
                    End If
                End If
            End If
        End If
    Next r
    
    If has173 Then
        PID_EnsureBG1Basis173Row = sourceData
        Exit Function
    End If
    
    If first151Row = 0 Then
        PID_EnsureBG1Basis173Row = sourceData
        Exit Function
    End If
    
    ReDim resultData(1 To UBound(sourceData, 1) + 1, 1 To 9)
    
    outRow = 0
    For r = 1 To UBound(sourceData, 1)
        outRow = outRow + 1
        
        If outRow = first151Row Then
            resultData(outRow, 1) = periodName
            resultData(outRow, 2) = sourceData(first151Row, 2)
            resultData(outRow, 3) = sourceData(first151Row, 3)
            resultData(outRow, 4) = "BG1_Basis"
            resultData(outRow, 5) = PID_GetNearestTextValue(sourceData, first151Row, 5)
            resultData(outRow, 6) = PID_GetNearestTextValue(sourceData, first151Row, 6)
            resultData(outRow, 7) = 173#
            resultData(outRow, 8) = 2021#
            resultData(outRow, 9) = ""
            outRow = outRow + 1
        End If
        
        resultData(outRow, 1) = sourceData(r, 1)
        resultData(outRow, 2) = sourceData(r, 2)
        resultData(outRow, 3) = sourceData(r, 3)
        resultData(outRow, 4) = sourceData(r, 4)
        resultData(outRow, 5) = sourceData(r, 5)
        resultData(outRow, 6) = sourceData(r, 6)
        resultData(outRow, 7) = sourceData(r, 7)
        resultData(outRow, 8) = sourceData(r, 8)
        resultData(outRow, 9) = sourceData(r, 9)
    Next r
    
    PID_EnsureBG1Basis173Row = resultData
    Exit Function
    
Fallback:
    PID_EnsureBG1Basis173Row = sourceData
End Function


Private Function PID_TryReadDouble(ByVal valueToRead As Variant, ByRef resultValue As Double) As Boolean
    Dim s As String
    
    On Error GoTo SafeExit
    
    If IsNumeric(valueToRead) Then
        resultValue = CDbl(valueToRead)
        PID_TryReadDouble = True
        Exit Function
    End If
    
    s = Trim$(CStr(valueToRead))
    If s = "" Then GoTo SafeExit
    
    s = Replace(s, ".", ",")
    
    If IsNumeric(s) Then
        resultValue = CDbl(s)
        PID_TryReadDouble = True
    End If
    
SafeExit:
End Function


Private Function PID_GetNearestTextValue(ByVal sourceData As Variant, _
                                         ByVal rowIndex As Long, _
                                         ByVal colIndex As Long) As String
    Dim r As Long
    Dim valueText As String
    
    If rowIndex < 1 Or rowIndex > UBound(sourceData, 1) Then Exit Function
    
    valueText = Trim$(CStr(sourceData(rowIndex, colIndex)))
    If valueText <> "" Then
        PID_GetNearestTextValue = valueText
        Exit Function
    End If
    
    For r = rowIndex - 1 To 1 Step -1
        valueText = Trim$(CStr(sourceData(r, colIndex)))
        If valueText <> "" Then
            PID_GetNearestTextValue = valueText
            Exit Function
        End If
    Next r
    
    For r = rowIndex + 1 To UBound(sourceData, 1)
        valueText = Trim$(CStr(sourceData(r, colIndex)))
        If valueText <> "" Then
            PID_GetNearestTextValue = valueText
            Exit Function
        End If
    Next r
End Function


Private Function AskForKVStartYear(ByVal defaultYear As Long) As Long
    Dim inputText As String
    
    inputText = InputBox( _
        Prompt:="Schritt 1 von 1 - Startjahr eingeben." & vbCrLf & vbCrLf & _
                "Beispiel: 2026 erzeugt die Periode KV 2026/2027." & vbCrLf & _
                "Nur die vier Ziffern des Jahres eingeben.", _
        Title:="Neue Periode", _
        Default:=CStr(defaultYear) _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Ung" & PID_KVTxtUe() & "ltige Eingabe. Bitte nur das Startjahr eingeben (z.B. 2026).", _
               vbExclamation, "Neuer KV-Zeitraum"
        Exit Function
    End If
    
    AskForKVStartYear = CLng(inputText)
    
    If AskForKVStartYear < 2000 Or AskForKVStartYear > 2100 Then
        MsgBox "Das Startjahr liegt au" & PID_KVTxtSs() & "erhalb des erlaubten Bereichs (2000-2100).", _
               vbExclamation, "Neuer KV-Zeitraum"
        AskForKVStartYear = 0
    End If
End Function


Private Function AskForSchemaCount(ByVal defaultCount As Long) As Long
    Dim inputText As String
    
    inputText = InputBox( _
        Prompt:="Wie viele Vertr" & PID_KVTxtAe() & "ge/Monatsstunden-Zeilen pro KV-Code sollen erzeugt werden?" & vbCrLf & vbCrLf & _
                "Beispiel: 13 (wie bisher).", _
        Title:="Vertragsanzahl pro KV-Code", _
        Default:=CStr(defaultCount) _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Ung" & PID_KVTxtUe() & "ltige Eingabe. Bitte eine ganze Zahl eingeben (z.B. 13).", _
               vbExclamation, "Vertragsanzahl"
        Exit Function
    End If
    
    AskForSchemaCount = CLng(inputText)
    
    If AskForSchemaCount < 1 Or AskForSchemaCount > 50 Then
        MsgBox "Die Vertragsanzahl muss zwischen 1 und 50 liegen.", _
               vbExclamation, "Vertragsanzahl"
        AskForSchemaCount = 0
    End If
End Function


Private Function BuildKVPeriodName(ByVal startYear As Long) As String
    BuildKVPeriodName = "KV " & CStr(startYear) & "/" & CStr(startYear + 1)
End Function


Private Function GetStartYearFromKVPeriod(ByVal periodName As String) As Long
    Dim s As String
    Dim parts As Variant
    
    s = NormalizeKVPeriodText(periodName)
    If Left$(s, 3) <> "KV " Then Exit Function
    
    s = Mid$(s, 4)
    parts = Split(s, "/")
    
    If UBound(parts) <> 1 Then Exit Function
    If Not IsNumeric(parts(0)) Then Exit Function
    
    GetStartYearFromKVPeriod = CLng(parts(0))
End Function


Private Function GetSchemaCountFromPeriod(ByVal wsKV As Worksheet, _
                                          ByVal periodFirstRow As Long, _
                                          ByVal periodLastRow As Long) As Long
    Dim firstCode As String
    Dim firstDataRowInPeriod As Long
    Dim r As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    If periodFirstRow <= 0 Or periodLastRow < periodFirstRow Then Exit Function
    
    firstDataRowInPeriod = 0
    
    For r = periodFirstRow To periodLastRow
        If Not wsKV.Range("A" & r).MergeCells Then
            firstCode = Trim$(CStr(wsKV.Cells(r, "D").Value))
            If firstCode <> "" Then
                firstDataRowInPeriod = r
                Exit For
            End If
        End If
    Next r
    
    If firstDataRowInPeriod = 0 Then Exit Function
    
    firstCode = Trim$(CStr(wsKV.Cells(firstDataRowInPeriod, "D").Value))
    
    For r = firstDataRowInPeriod To periodLastRow
        If Trim$(CStr(wsKV.Cells(r, "D").Value)) = firstCode Then
            GetSchemaCountFromPeriod = GetSchemaCountFromPeriod + 1
        Else
            Exit Function
        End If
    Next r
    
SafeExit:
End Function


Private Function BuildNewPeriodDataFromTemplate(ByVal wsKV As Worksheet, _
                                                ByVal templateFirstRow As Long, _
                                                ByVal templateLastRow As Long, _
                                                ByVal newPeriod As String, _
                                                ByVal newStartYear As Long, _
                                                ByVal newSchemaCount As Long) As Variant
    Dim templateData As Variant
    Dim blockStarts As Collection
    Dim blockEnds As Collection
    Dim totalRows As Long
    Dim r As Long
    Dim outRow As Long
    Dim blockIndex As Long
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim blockRows As Long
    Dim schemaIndex As Long
    Dim sourceOffset As Long
    Dim validFrom As Date
    Dim validTo As Date
    Dim resultData As Variant
    Dim currentCode As String
    Dim rowCode As String
    
    On Error GoTo BuildFail
    
    If wsKV Is Nothing Then Exit Function
    If templateLastRow < templateFirstRow Then Exit Function
    If newSchemaCount <= 0 Then Exit Function
    
    templateData = wsKV.Range("A" & templateFirstRow & ":I" & templateLastRow).Value
    
    Set blockStarts = New Collection
    Set blockEnds = New Collection
    currentCode = ""
    
    ' Build blocks only from real KV codes (ignore empty separator/title rows).
    For r = 1 To UBound(templateData, 1)
        rowCode = Trim$(CStr(templateData(r, 4)))
        
        If rowCode <> "" Then
            If currentCode = "" Then
                blockStarts.Add r
                currentCode = rowCode
            ElseIf rowCode <> currentCode Then
                blockEnds.Add r - 1
                blockStarts.Add r
                currentCode = rowCode
            End If
        Else
            If currentCode <> "" Then
                blockEnds.Add r - 1
                currentCode = ""
            End If
        End If
    Next r
    
    If currentCode <> "" Then
        blockEnds.Add UBound(templateData, 1)
    End If
    
    If blockStarts.Count = 0 Or blockEnds.Count = 0 Then GoTo BuildFail
    
    totalRows = blockStarts.Count * newSchemaCount
    ReDim resultData(1 To totalRows, 1 To 9)
    
    validFrom = DateSerial(newStartYear, 5, 1)
    validTo = DateSerial(newStartYear + 1, 4, 30)
    
    outRow = 0
    
    For blockIndex = 1 To blockStarts.Count
        blockStart = CLng(blockStarts(blockIndex))
        blockEnd = CLng(blockEnds(blockIndex))
        blockRows = blockEnd - blockStart + 1
        
        For schemaIndex = 1 To newSchemaCount
            outRow = outRow + 1
            
            resultData(outRow, 1) = newPeriod
            resultData(outRow, 2) = validFrom
            resultData(outRow, 3) = validTo
            resultData(outRow, 4) = templateData(blockStart, 4)
            resultData(outRow, 5) = templateData(blockStart, 5)
            resultData(outRow, 6) = templateData(blockStart, 6)
            
            sourceOffset = schemaIndex - 1
            If sourceOffset < blockRows Then
                resultData(outRow, 7) = templateData(blockStart + sourceOffset, 7)
                ' Monatslohn wird pro neuer KV-Periode neu erfasst (nicht aus Vorlage uebernehmen).
                resultData(outRow, 8) = ""
                resultData(outRow, 9) = ""
            Else
                resultData(outRow, 7) = ""
                resultData(outRow, 8) = ""
                resultData(outRow, 9) = ""
            End If
        Next schemaIndex
    Next blockIndex
    
    BuildNewPeriodDataFromTemplate = resultData
    Exit Function
    
BuildFail:
    Erase resultData
End Function


Private Function GetBottomKVPeriod(ByVal wsKV As Worksheet, ByVal firstDataRow As Long) As String
    Dim r As Long
    Dim lastRow As Long
    Dim valueText As String
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    
    For r = lastRow To firstDataRow Step -1
        valueText = PID_GetRowKVPeriod(wsKV, r)
        If valueText <> "" Then
            GetBottomKVPeriod = valueText
            Exit Function
        End If
    Next r
    
SafeExit:
End Function


Private Sub PID_NormalizeKVWarningText(ByVal wsKV As Worksheet)
    If wsKV Is Nothing Then Exit Sub
    
    PID_ConfigureLOHNTABELLEHeaderLayout wsKV
End Sub


Private Sub PID_NormalizeKVTableHeader(ByVal wsKV As Worksheet)
    If wsKV Is Nothing Then Exit Sub
    
    wsKV.Range("A3").Value = "KV-Periode"
    wsKV.Range("B3").Value = PID_KVTxtGueltig() & " ab"
    wsKV.Range("C3").Value = PID_KVTxtGueltig() & " bis"
    wsKV.Range("D3").Value = "KV-Code"
    wsKV.Range("E3").Value = "KV-Gruppe"
    wsKV.Range("F3").Value = PID_KVTxtBeschaeftigungsdauer()
    wsKV.Range("G3").Value = "Monatsstunden"
    wsKV.Range("H3").Value = "Monatslohn"
    wsKV.Range("I3").Value = "Status"
    wsKV.Range("J3").Value = PID_KVTxtPruefung()
    
    PID_StyleApplyCompactHeaderBand wsKV.Range("A3:J3")
    wsKV.Rows(3).RowHeight = PID_STYLE_COMPACT_HEADER_ROW_HEIGHT
End Sub


Private Sub PID_ClearKVDataArea(ByVal wsKV As Worksheet, ByVal firstDataRow As Long, ByVal lastRow As Long)
    Dim targetRange As Range
    
    On Error GoTo TryUnmerge
    
    Set targetRange = wsKV.Range("A" & firstDataRow & ":J" & lastRow)
    targetRange.Clear
    Exit Sub
    
TryUnmerge:
    On Error Resume Next
    targetRange.UnMerge
    targetRange.Clear
    On Error GoTo 0
End Sub


Private Sub PID_ClearTrailingKVArea(ByVal wsKV As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If startRow > endRow Then Exit Sub
    
    On Error Resume Next
    wsKV.Range("A" & startRow & ":XFD" & endRow).UnMerge
    wsKV.Range("A" & startRow & ":XFD" & endRow).Clear
    wsKV.Rows(startRow & ":" & endRow).Delete Shift:=xlShiftUp
    On Error GoTo 0
    
SafeExit:
End Sub


Private Sub PID_TrimKVSheetBelowTable(ByVal wsKV As Worksheet, ByVal firstDataRow As Long)
    Dim lastDataRow As Long
    Dim trimStart As Long
    Dim usedLastRow As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    lastDataRow = PID_GetKVTableLastRow(wsKV, firstDataRow)
    trimStart = lastDataRow + 1
    usedLastRow = PID_GetSheetCleanupLastRow(wsKV, lastDataRow)
    
    If trimStart <= usedLastRow Then
        PID_ClearTrailingKVArea wsKV, trimStart, usedLastRow
    End If
    
    ' Spalte K = PID_EIGEN-Markierung fuer eigene Stunden — nicht loeschen.
    On Error Resume Next
    If lastDataRow >= firstDataRow Then
        wsKV.Range("L" & firstDataRow & ":XFD" & lastDataRow).Clear
    End If
    
SafeExit:
End Sub


Private Function PID_GetSheetCleanupLastRow(ByVal ws As Worksheet, ByVal fallbackRow As Long) As Long
    Dim usedLastRow As Long
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then
        PID_GetSheetCleanupLastRow = fallbackRow
        Exit Function
    End If
    
    usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    
    If usedLastRow > fallbackRow Then
        PID_GetSheetCleanupLastRow = usedLastRow
    Else
        PID_GetSheetCleanupLastRow = fallbackRow
    End If
    
    Exit Function
    
SafeExit:
    PID_GetSheetCleanupLastRow = fallbackRow
End Function


Public Sub CleanupLOHNTABELLETrailingArea()
    Dim wsKV As Worksheet
    Dim firstDataRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    If wsKV Is Nothing Then Exit Sub
    
    firstDataRow = 4
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    PID_TrimKVSheetBelowTable wsKV, firstDataRow
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Sub InsertNewKVPeriodRows(ByVal wsKV As Worksheet, _
                                  ByVal firstDataRow As Long, _
                                  ByVal newPeriodData As Variant)
    Dim newRowCount As Long
    Dim insertRowCount As Long
    Dim dataStartRow As Long
    Dim insertEndRow As Long
    Dim newArea As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If Not IsArray(newPeriodData) Then Exit Sub
    
    newRowCount = UBound(newPeriodData, 1)
    insertRowCount = newRowCount + 1
    
    If newRowCount <= 0 Then Exit Sub
    
    insertEndRow = firstDataRow + insertRowCount - 1
    dataStartRow = firstDataRow + 1
    
    wsKV.Rows(firstDataRow & ":" & insertEndRow).Insert Shift:=xlDown
    
    Set newArea = wsKV.Range("A" & firstDataRow & ":J" & insertEndRow)
    
    On Error Resume Next
    newArea.UnMerge
    On Error GoTo SafeExit
    
    ' Nur Inhalte leeren — Zeilenhoehen/Spaltenbreiten der Vorlage bleiben bis FormatKVPeriodArea.
    newArea.ClearContents
    
    wsKV.Range("A" & dataStartRow & ":I" & (dataStartRow + newRowCount - 1)).Value = newPeriodData
    
    PID_WriteKVPeriodTitleRow wsKV, firstDataRow, CStr(newPeriodData(1, 1)), newPeriodData(1, 2), newPeriodData(1, 3)
    
    Application.CutCopyMode = False
    
SafeExit:
    Application.CutCopyMode = False
End Sub


Private Sub PID_WriteKVPeriodTitleRow(ByVal wsKV As Worksheet, _
                                      ByVal rowNumber As Long, _
                                      ByVal periodName As String, _
                                      ByVal validFrom As Variant, _
                                      ByVal validTo As Variant)
    Dim titleText As String
    
    If wsKV Is Nothing Then Exit Sub
    If rowNumber < 1 Then Exit Sub
    
    titleText = periodName
    
    If IsDate(validFrom) And IsDate(validTo) Then
        titleText = titleText & "   |   " & LCase$(PID_KVTxtGueltig()) & " von " & Format$(CDate(validFrom), "dd.mm.yyyy") & _
                    " bis " & Format$(CDate(validTo), "dd.mm.yyyy")
    End If
    
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).UnMerge
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).ClearContents
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).Merge
    
    wsKV.Cells(rowNumber, "A").Value = titleText
    
    PID_StyleApplySubsectionTitle wsKV.Range("A" & rowNumber & ":J" & rowNumber), False
    wsKV.Rows(rowNumber).RowHeight = PID_STYLE_COMPACT_YEAR_ROW_HEIGHT
    
    With wsKV.Range("A" & rowNumber & ":J" & rowNumber)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).Color = PID_StyleColorNavy()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
    End With
End Sub


Private Sub PID_RefreshKVPeriodTitleRowStyles(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    For r = firstRow To lastRow
        If wsKV.Range("A" & r).MergeCells Then
            PID_StyleApplySubsectionTitle wsKV.Range("A" & r & ":J" & r), False
            
            With wsKV.Range("A" & r & ":J" & r)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeTop).Color = PID_StyleColorNavy()
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
            End With
        End If
    Next r
End Sub


Private Function GetTopKVPeriod(ByVal wsKV As Worksheet, ByVal firstDataRow As Long) As String
    Dim r As Long
    Dim lastRow As Long
    Dim rowPeriod As String
    
    On Error GoTo SafeExit
    
    lastRow = wsKV.Cells(wsKV.Rows.count, "A").End(xlUp).Row
    
    For r = firstDataRow To lastRow
        rowPeriod = PID_GetRowKVPeriod(wsKV, r)
        
        If rowPeriod <> "" Then
            GetTopKVPeriod = rowPeriod
            Exit Function
        End If
    Next r

SafeExit:
End Function


Private Function FindFirstRowOfPeriod(ByVal wsKV As Worksheet, _
                                      ByVal periodName As String, _
                                      ByVal firstDataRow As Long) As Long
    Dim boundsFirst As Long
    Dim boundsLast As Long
    
    periodName = NormalizeKVPeriodText(periodName)
    
    If PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, boundsFirst, boundsLast) Then
        FindFirstRowOfPeriod = boundsFirst
    End If
End Function


Private Function FindLastRowOfPeriod(ByVal wsKV As Worksheet, _
                                     ByVal periodName As String, _
                                     ByVal firstDataRow As Long) As Long
    Dim boundsFirst As Long
    Dim boundsLast As Long
    
    periodName = NormalizeKVPeriodText(periodName)
    
    If PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, boundsFirst, boundsLast) Then
        FindLastRowOfPeriod = boundsLast
    End If
End Function


Private Function PID_GetPeriodRowBounds(ByVal wsKV As Worksheet, _
                                        ByVal periodName As String, _
                                        ByVal firstDataRow As Long, _
                                        ByRef outFirstRow As Long, _
                                        ByRef outLastRow As Long) As Boolean
    Dim r As Long
    Dim lastRow As Long
    Dim rowPeriod As String
    Dim inPeriod As Boolean
    
    On Error GoTo SafeExit
    
    outFirstRow = 0
    outLastRow = 0
    inPeriod = False
    
    If wsKV Is Nothing Then Exit Function
    If periodName = "" Then Exit Function
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    
    For r = firstDataRow To lastRow
        rowPeriod = PID_GetRowKVPeriod(wsKV, r)
        
        If rowPeriod = periodName Then
            If PID_RowHasKVTableContent(wsKV, r) Then
                If outFirstRow = 0 Then outFirstRow = r
                outLastRow = r
            End If
            inPeriod = True
        ElseIf inPeriod And rowPeriod <> "" Then
            Exit For
        End If
    Next r
    
    PID_GetPeriodRowBounds = (outFirstRow > 0 And outLastRow >= outFirstRow)
    
SafeExit:
End Function


Private Function KVPeriodExists(ByVal wsKV As Worksheet, _
                                ByVal periodName As String, _
                                ByVal firstDataRow As Long) As Boolean
    Dim boundsFirst As Long
    Dim boundsLast As Long
    
    periodName = NormalizeKVPeriodText(periodName)
    KVPeriodExists = PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, boundsFirst, boundsLast)
End Function


Private Function NormalizeKVPeriodText(ByVal periodText As String) As String
    Dim s As String
    
    s = Trim$(CStr(periodText))
    
    If s = "" Then
        NormalizeKVPeriodText = ""
        Exit Function
    End If
    
    If InStr(1, s, "|", vbTextCompare) > 0 Then
        s = Trim$(Left$(s, InStr(1, s, "|", vbTextCompare) - 1))
    End If
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    s = Replace(s, "kv ", "KV ")
    s = Replace(s, "Kv ", "KV ")
    s = Replace(s, "kV ", "KV ")
    
    NormalizeKVPeriodText = s
End Function


Private Function PID_GetRowKVPeriod(ByVal wsKV As Worksheet, ByVal rowNumber As Long) As String
    Dim r As Long
    Dim cellText As String
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    If rowNumber < 1 Then Exit Function
    
    cellText = Trim$(CStr(wsKV.Cells(rowNumber, "A").Value))
    
    If cellText <> "" Then
        PID_GetRowKVPeriod = NormalizeKVPeriodText(cellText)
        Exit Function
    End If
    
    ' Leere Restzeilen nicht dem oberen Zeitraum zuordnen.
    If Trim$(CStr(wsKV.Cells(rowNumber, "D").Value)) = "" Then Exit Function
    
    For r = rowNumber - 1 To 4 Step -1
        cellText = Trim$(CStr(wsKV.Cells(r, "A").Value))
        If cellText <> "" Then
            PID_GetRowKVPeriod = NormalizeKVPeriodText(cellText)
            Exit Function
        End If
    Next r
    
SafeExit:
End Function


Private Function IsValidKVPeriodName(ByVal periodName As String) As Boolean
    Dim s As String
    Dim yearPart As String
    Dim parts As Variant
    Dim y1 As Long
    Dim y2 As Long
    
    On Error GoTo InvalidPeriod
    
    s = NormalizeKVPeriodText(periodName)
    
    If Left$(s, 3) <> "KV " Then GoTo InvalidPeriod
    
    yearPart = Mid$(s, 4)
    
    If InStr(yearPart, "/") = 0 Then GoTo InvalidPeriod
    
    parts = Split(yearPart, "/")
    
    If UBound(parts) <> 1 Then GoTo InvalidPeriod
    If Not IsNumeric(parts(0)) Then GoTo InvalidPeriod
    If Not IsNumeric(parts(1)) Then GoTo InvalidPeriod
    
    y1 = CLng(parts(0))
    y2 = CLng(parts(1))
    
    If y1 < 2000 Or y1 > 2100 Then GoTo InvalidPeriod
    If y2 < 2000 Or y2 > 2100 Then GoTo InvalidPeriod
    If y2 <> y1 + 1 Then GoTo InvalidPeriod
    
    IsValidKVPeriodName = True
    Exit Function

InvalidPeriod:
    IsValidKVPeriodName = False
End Function


Private Function GetNextKVPeriodName(ByVal currentPeriodName As String) As String
    Dim s As String
    Dim yearPart As String
    Dim parts As Variant
    Dim y1 As Long
    Dim y2 As Long
    
    On Error GoTo Fallback
    
    s = NormalizeKVPeriodText(currentPeriodName)
    
    If Left$(s, 3) <> "KV " Then GoTo Fallback
    
    yearPart = Mid$(s, 4)
    parts = Split(yearPart, "/")
    
    If UBound(parts) <> 1 Then GoTo Fallback
    If Not IsNumeric(parts(0)) Then GoTo Fallback
    If Not IsNumeric(parts(1)) Then GoTo Fallback
    
    y1 = CLng(parts(0))
    y2 = CLng(parts(1))
    
    GetNextKVPeriodName = "KV " & CStr(y1 + 1) & "/" & CStr(y2 + 1)
    Exit Function

Fallback:
    GetNextKVPeriodName = "KV " & CStr(Year(Date)) & "/" & CStr(Year(Date) + 1)
End Function


Private Sub FormatKVPeriodArea(ByVal wsKV As Worksheet)
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    firstDataRow = PID_GetKVTableFirstDataRow(wsKV)
    lastRow = PID_GetKVTableLastRow(wsKV, firstDataRow)
    
    If lastRow < firstDataRow Then GoTo SafeExit
    
    With wsKV
        .Range("A" & firstDataRow & ":I" & lastRow).VerticalAlignment = xlCenter
        .Range("A" & firstDataRow & ":I" & lastRow).HorizontalAlignment = xlCenter
        .Range("B" & firstDataRow & ":C" & lastRow).NumberFormat = "dd.mm.yyyy"
        
        ' Gesamte Datenflaeche einheitlich mit duennem Raster versehen.
        .Range("A" & firstDataRow & ":J" & lastRow).Borders.LineStyle = xlContinuous
        .Range("A" & firstDataRow & ":J" & lastRow).Borders.Weight = xlThin
        .Range("A" & firstDataRow & ":J" & lastRow).Borders.Color = RGB(180, 180, 180)
        
        .Range("A" & firstDataRow & ":A" & lastRow).NumberFormat = "@"
        .Range("G" & firstDataRow & ":G" & lastRow).NumberFormatLocal = "0,00"
        
        PID_ApplyEuroNumberFormat .Range("H" & firstDataRow & ":H" & lastRow)
        
        PID_ConfigureKVTableColumnWidths wsKV
        PID_ConfigureKVStatusColumnWidths wsKV, firstDataRow, lastRow
        .Columns(PID_CUSTOM_KV_HOUR_MARKER_COL).Hidden = True
    End With
    
    ' Status- und Pruefungsformeln auf allen gueltigen Datenzeilen wiederherstellen.
    PID_ApplyKVStatusFormulas wsKV, firstDataRow, lastRow
    
    ' Eingabefelder fuer Monatsstunden/Monatslohn muessen editierbar bleiben.
    PID_ConfigureKVInputCellLocks wsKV, firstDataRow, lastRow
    
    PID_ApplyKVTableRowHeights wsKV, firstDataRow, lastRow
    PID_ApplyKVVisualGrouping wsKV, firstDataRow, lastRow
    PID_FormatKVRowTypography wsKV, firstDataRow, lastRow
    PID_TrimKVSheetBelowTable wsKV, firstDataRow

SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Sub PID_EnsureKVStatusFormulas(ByVal wsKV As Worksheet)
    Dim lastRow As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Sub
    
    PID_ApplyKVStatusFormulas wsKV, 4, lastRow
    PID_ConfigureKVStatusColumnWidths wsKV, 4, lastRow
    PID_ConfigureKVInputCellLocks wsKV, 4, lastRow
    
SafeExit:
End Sub


Private Sub PID_FormatKVRowTypography(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    Dim rowRange As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    wsKV.Range("A3:J3").Font.Bold = True
    wsKV.Range("A3:J3").Font.Size = 10
    wsKV.Range("A3:J3").Font.Color = PID_StyleColorNavy()
    
    For r = firstRow To lastRow
        Set rowRange = wsKV.Range("A" & r & ":J" & r)
        
        If wsKV.Range("A" & r).MergeCells Then
            ' Titelzeile: Navy in PID_WriteKVPeriodTitleRow / PID_RefreshKVPeriodTitleRowStyles.
        ElseIf Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "" Then
            rowRange.Font.Bold = False
            rowRange.Font.Size = 10
            rowRange.Font.Color = PID_StyleColorNavy()
            rowRange.VerticalAlignment = xlCenter
            rowRange.HorizontalAlignment = xlCenter
        End If
    Next r
    
SafeExit:
End Sub


Private Sub PID_ConfigureKVTableColumnWidths(ByVal wsKV As Worksheet)
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    ' Spaltenbreiten an bestehende LOHNTABELLE-Vorlage (OOXML-Baseline) angleichen.
    wsKV.Columns("A").ColumnWidth = 16
    wsKV.Columns("B").ColumnWidth = 14
    wsKV.Columns("C").ColumnWidth = 14
    wsKV.Columns("D").ColumnWidth = 14
    wsKV.Columns("E").ColumnWidth = 13
    wsKV.Columns("F").ColumnWidth = 25
    wsKV.Columns("G").ColumnWidth = 13
    wsKV.Columns("H").ColumnWidth = 14
    
SafeExit:
End Sub


Private Sub PID_ApplyKVTableRowHeights(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    ' Nach Insert+ClearContents sind neue Zeilen oft auf Excel-Standardhoehe (~15).
    wsKV.Rows(firstRow & ":" & lastRow).RowHeight = PID_STYLE_COMPACT_YEAR_ROW_HEIGHT
    
SafeExit:
End Sub


Private Sub PID_ConfigureKVStatusColumnWidths(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    wsKV.Columns("I").ColumnWidth = 22
    wsKV.Columns("J").ColumnWidth = 24
    
    If firstRow <= lastRow Then
        wsKV.Range("I" & firstRow & ":J" & lastRow).WrapText = False
    End If
    
SafeExit:
End Sub


Public Sub PID_RecalculateKVStatusColumns(Optional ByVal wsKV As Worksheet)
    Dim lastRow As Long
    Dim calcRange As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Set wsKV = ThisWorkbook.Worksheets(PID_LOHNTABELLE_SHEET)
    If wsKV Is Nothing Then Exit Sub
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Sub
    
    ' COUNTIFS in Spalte J kann mehrere Zeilen betreffen -> immer gesamten Statusblock berechnen.
    Set calcRange = wsKV.Range("I4:J" & lastRow)
    calcRange.Calculate
    
SafeExit:
End Sub


Private Sub PID_ApplyKVStatusFormulas(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    Dim hasKeyData As Boolean
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    For r = firstRow To lastRow
        On Error Resume Next
        Err.Clear
        
        If wsKV.Range("A" & r).MergeCells Then
            wsKV.Cells(r, "I").ClearContents
            wsKV.Cells(r, "J").ClearContents
        Else
            hasKeyData = (Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "")
            
            If hasKeyData Then
                wsKV.Cells(r, "I").FormulaR1C1 = _
                    "=IF(RC1="""","""",IF(OR(RC2="""",RC3="""",RC4="""",RC5="""",RC6=""""),""Stammdaten fehlen"",IF(AND(RC7="""",RC8=""""),""Werte fehlen"",IF(RC7="""",""Monatsstunden fehlen"",IF(RC8="""",""Monatslohn fehlt"",""OK"")))))"
                
                wsKV.Cells(r, "J").FormulaR1C1 = _
                    "=IF(RC7="""","""",IF(COUNTIFS(C1,RC1,C4,RC4,C7,RC7)>1,""Doppelte Monatsstunden"",""""))"
            Else
                wsKV.Cells(r, "I").ClearContents
                wsKV.Cells(r, "J").ClearContents
            End If
        End If
        
        On Error GoTo 0
    Next r
    
    PID_RecalculateKVStatusColumns wsKV
End Sub


Private Sub PID_ConfigureKVInputCellLocks(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    wsKV.Range("A" & firstRow & ":J" & lastRow).Locked = True
    
    For r = firstRow To lastRow
        ' Nur echte Datenzeilen freigeben (keine zusammengefuehrten Titelzeilen).
        If Not wsKV.Range("A" & r).MergeCells Then
            If Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "" Then
                wsKV.Range("G" & r & ":H" & r).Locked = False
            End If
        End If
    Next r
    
SafeExit:
End Sub


Private Sub PID_ApplyKVVisualGrouping(ByVal wsKV As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    Dim currentPeriod As String
    Dim prevPeriod As String
    Dim currentCode As String
    Dim prevCode As String
    Dim groupText As String
    Dim rowRange As Range
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    If firstRow > lastRow Then Exit Sub
    
    For r = firstRow To lastRow
        Set rowRange = wsKV.Range("A" & r & ":J" & r)
        
        ' Titelzeilen nicht ueberschreiben (bleiben grau/formatiert).
        If wsKV.Range("A" & r).MergeCells Then GoTo NextRow
        
        currentPeriod = PID_GetRowKVPeriod(wsKV, r)
        currentCode = UCase$(Trim$(CStr(wsKV.Cells(r, "D").Value)))
        groupText = UCase$(Trim$(CStr(wsKV.Cells(r, "E").Value)))
        
        ' Reset per-row visual baseline first.
        rowRange.Interior.Pattern = xlSolid
        rowRange.Interior.PatternColorIndex = xlAutomatic
        rowRange.Interior.TintAndShade = 0
        rowRange.Interior.PatternTintAndShade = 0
        rowRange.Interior.ColorIndex = xlNone
        
        rowRange.Borders(xlEdgeTop).LineStyle = xlContinuous
        rowRange.Borders(xlEdgeTop).Weight = xlThin
        rowRange.Borders(xlEdgeTop).Color = RGB(180, 180, 180)
        
        ' Soft BG color blocks (UEBERSICHT-Palette).
        Select Case True
            Case InStr(1, groupText, "BG1", vbTextCompare) > 0
                rowRange.Interior.Color = PID_StyleColorHeaderBg()
            Case InStr(1, groupText, "BG2", vbTextCompare) > 0
                rowRange.Interior.Color = vbWhite
            Case InStr(1, groupText, "BG3", vbTextCompare) > 0
                rowRange.Interior.Color = PID_StyleColorAccent()
        End Select
        
        ' Strong separator when a new KV period starts.
        If currentPeriod <> "" Then
            If currentPeriod <> prevPeriod Then
                rowRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                rowRange.Borders(xlEdgeTop).Weight = xlMedium
                rowRange.Borders(xlEdgeTop).Color = PID_StyleColorNavy()
                
                ' Bei neuem Zeitraum die KV-Code-Referenz zuruecksetzen.
                prevCode = ""
            End If
            prevPeriod = currentPeriod
        End If
        
        ' Zusaetzliche Trennlinie zwischen KV-Untergruppen (Basis / 5 / 10 / 15).
        If currentCode <> "" Then
            If prevCode <> "" And currentCode <> prevCode Then
                rowRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                rowRange.Borders(xlEdgeTop).Weight = xlMedium
                rowRange.Borders(xlEdgeTop).Color = PID_StyleColorNavy()
            End If
            prevCode = currentCode
        End If
        
NextRow:
    Next r
    
    ' Aussenrahmen der Tabelle am Ende explizit verstaerken.
    With wsKV.Range("A" & firstRow & ":J" & lastRow)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = PID_StyleColorNavy()
        
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeRight).Color = PID_StyleColorNavy()
        
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = PID_StyleColorNavy()
    End With
    
SafeExit:
End Sub


Private Function AskForKVPeriodSelection(ByVal wsKV As Worksheet, _
                                         ByVal firstDataRow As Long, _
                                         Optional ByVal dialogTitle As String = "KV-Zeitraum", _
                                         Optional ByVal stepText As String = "", _
                                         Optional ByVal extraHint As String = "") As String
    Dim periods As Collection
    Dim promptText As String
    Dim inputText As String
    Dim selectedIndex As Long
    Dim defaultIndex As Long
    Dim titleText As String
    
    Set periods = PID_CollectKVPeriods(wsKV, firstDataRow)
    
    If periods Is Nothing Then Exit Function
    If periods.Count = 0 Then
        MsgBox "Kein g" & PID_KVTxtUe() & "ltiger KV-Zeitraum gefunden.", vbExclamation, dialogTitle
        Exit Function
    End If
    
    defaultIndex = 1
    promptText = ""
    
    If stepText <> "" Then
        promptText = stepText & vbCrLf & vbCrLf
    End If
    
    promptText = promptText & "Liste:" & vbCrLf & vbCrLf
    promptText = promptText & PID_BuildNumberedListFromCollection(periods)
    promptText = promptText & vbCrLf & "Nur die Nummer eintippen (meist ""1"") und OK klicken."
    
    If extraHint <> "" Then
        promptText = promptText & vbCrLf & vbCrLf & extraHint
    End If
    
    titleText = dialogTitle
    If stepText <> "" Then titleText = dialogTitle & " - " & stepText
    
    inputText = InputBox(Prompt:=promptText, Title:=titleText, Default:=CStr(defaultIndex))
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Das war keine g" & PID_KVTxtUe() & "ltige Nummer." & vbCrLf & _
               "Bitte nur die Zahl aus der Liste eingeben (z.B. 1).", _
               vbExclamation, dialogTitle
        Exit Function
    End If
    
    selectedIndex = CLng(inputText)
    
    If selectedIndex < 1 Or selectedIndex > periods.Count Then
        MsgBox "Diese Nummer steht nicht in der Liste." & vbCrLf & _
               "Bitte eine Nummer zwischen 1 und " & CStr(periods.Count) & " eingeben.", _
               vbExclamation, dialogTitle
        Exit Function
    End If
    
    AskForKVPeriodSelection = CStr(periods(selectedIndex))
End Function


Private Function AskForKVCodeSelection(Optional ByVal dialogTitle As String = "Eigene Stunden", _
                                       Optional ByVal stepText As String = "Schritt 2 von 4 - Welche Gruppe?", _
                                       Optional ByVal extraHint As String = "") As String
    Dim promptText As String
    Dim inputText As String
    Dim selectedIndex As Long
    Dim i As Long
    Dim titleText As String
    
    promptText = ""
    
    If stepText <> "" Then
        promptText = stepText & vbCrLf & vbCrLf
    End If
    
    promptText = promptText & "Liste:" & vbCrLf & vbCrLf
    
    For i = 1 To PID_KV_CODE_COUNT
        promptText = promptText & CStr(i) & " = " & PID_GetStandardKVCodeByIndex(i) & vbCrLf
    Next i
    
    promptText = promptText & vbCrLf & "Nummer eintippen und OK klicken."
    
    If extraHint <> "" Then
        promptText = promptText & vbCrLf & vbCrLf & extraHint
    End If
    
    titleText = dialogTitle
    If stepText <> "" Then titleText = dialogTitle & " - " & stepText
    
    inputText = InputBox(Prompt:=promptText, Title:=titleText, Default:="1")
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Das war keine g" & PID_KVTxtUe() & "ltige Nummer." & vbCrLf & _
               "Bitte nur die Zahl aus der Liste eingeben (z.B. 1).", _
               vbExclamation, dialogTitle
        Exit Function
    End If
    
    selectedIndex = CLng(inputText)
    
    If selectedIndex < 1 Or selectedIndex > PID_KV_CODE_COUNT Then
        MsgBox "Diese Nummer steht nicht in der Liste." & vbCrLf & _
               "Bitte eine Nummer zwischen 1 und " & CStr(PID_KV_CODE_COUNT) & " eingeben.", _
               vbExclamation, dialogTitle
        Exit Function
    End If
    
    AskForKVCodeSelection = PID_GetStandardKVCodeByIndex(selectedIndex)
End Function


Private Function AskForCustomMonatsstunden(ByRef outHours As Double) As Boolean
    Dim inputText As String
    
    inputText = InputBox( _
        Prompt:="Schritt 3 von 4 - Wie viele Stunden pro Monat?" & vbCrLf & vbCrLf & _
                "Beispiel: 64 oder 160,00" & vbCrLf & _
                "Zahl eintippen und OK klicken.", _
        Title:="Eigene Stunden", _
        Default:="" _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not PID_TryReadDouble(inputText, outHours) Then
        MsgBox "Das war keine g" & PID_KVTxtUe() & "ltige Zahl." & vbCrLf & _
               "Bitte nur Stunden eingeben (z.B. 64).", _
               vbExclamation, "Eigene Stunden"
        Exit Function
    End If
    
    If outHours <= 0# Then
        MsgBox "Stunden " & PID_KVTxtMuessen() & " " & PID_KVTxtGroesser() & " als 0 sein.", _
               vbExclamation, "Eigene Stunden"
        Exit Function
    End If
    
    AskForCustomMonatsstunden = True
End Function


Private Function AskForOptionalMonatslohn(ByRef outLohn As Variant) As Boolean
    Dim inputText As String
    Dim lohnValue As Double
    
    inputText = InputBox( _
        Prompt:="Schritt 4 von 4 - Monatslohn (optional)" & vbCrLf & vbCrLf & _
                "Wenn unklar: einfach Abbrechen oder leer lassen." & vbCrLf & _
                "Monatslohn k" & PID_KVTxtOe() & "nnen Sie " & PID_KVTxtSpaeter() & " auch direkt in LOHNTABELLE eintragen.", _
        Title:="Eigene Stunden", _
        Default:="" _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not PID_TryReadDouble(inputText, lohnValue) Then
        MsgBox "Ung" & PID_KVTxtUe() & "ltiger Monatslohn. Bitte eine Zahl eingeben oder leer lassen.", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    If lohnValue < 0# Then
        MsgBox "Monatslohn darf nicht negativ sein.", vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    outLohn = lohnValue
    AskForOptionalMonatslohn = True
End Function


Private Function PID_CollectKVPeriods(ByVal wsKV As Worksheet, ByVal firstDataRow As Long) As Collection
    Dim periods As Collection
    Dim lastRow As Long
    Dim r As Long
    Dim rowPeriod As String
    
    Set periods = New Collection
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then
        Set PID_CollectKVPeriods = periods
        Exit Function
    End If
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    
    For r = firstDataRow To lastRow
        rowPeriod = PID_GetRowKVPeriod(wsKV, r)
        
        If rowPeriod <> "" Then
            If Not PID_CollectionContainsText(periods, rowPeriod) Then
                periods.Add rowPeriod, PID_MakeCollectionKey(rowPeriod)
            End If
        End If
    Next r
    
SafeExit:
    Set PID_CollectKVPeriods = periods
End Function


Private Function PID_BuildNumberedListFromCollection(ByVal items As Collection) As String
    Dim resultText As String
    Dim i As Long
    
    If items Is Nothing Then Exit Function
    
    For i = 1 To items.Count
        resultText = resultText & CStr(i) & " = " & CStr(items(i)) & vbCrLf
    Next i
    
    PID_BuildNumberedListFromCollection = resultText
End Function


Private Function PID_CollectionContainsText(ByVal items As Collection, ByVal valueText As String) As Boolean
    Dim i As Long
    
    On Error GoTo SafeExit
    
    If items Is Nothing Then Exit Function
    
    For i = 1 To items.Count
        If StrComp(CStr(items(i)), valueText, vbTextCompare) = 0 Then
            PID_CollectionContainsText = True
            Exit Function
        End If
    Next i
    
SafeExit:
End Function


Private Function PID_MakeCollectionKey(ByVal valueText As String) As String
    PID_MakeCollectionKey = Replace(Replace(Trim$(valueText), " ", "_"), "/", "_")
End Function


Private Function PID_GetStandardKVCodeByIndex(ByVal indexNumber As Long) As String
    Select Case indexNumber
        Case 1: PID_GetStandardKVCodeByIndex = "BG1_Basis"
        Case 2: PID_GetStandardKVCodeByIndex = "BG1_5"
        Case 3: PID_GetStandardKVCodeByIndex = "BG1_10"
        Case 4: PID_GetStandardKVCodeByIndex = "BG1_15"
        Case 5: PID_GetStandardKVCodeByIndex = "BG2_Basis"
        Case 6: PID_GetStandardKVCodeByIndex = "BG2_5"
        Case 7: PID_GetStandardKVCodeByIndex = "BG2_10"
        Case 8: PID_GetStandardKVCodeByIndex = "BG2_15"
        Case 9: PID_GetStandardKVCodeByIndex = "BG3_Basis"
        Case 10: PID_GetStandardKVCodeByIndex = "BG3_5"
        Case 11: PID_GetStandardKVCodeByIndex = "BG3_10"
        Case 12: PID_GetStandardKVCodeByIndex = "BG3_15"
        Case Else: PID_GetStandardKVCodeByIndex = ""
    End Select
End Function


Private Function PID_GetKVCodeBlockBounds(ByVal wsKV As Worksheet, _
                                            ByVal periodName As String, _
                                            ByVal kvCode As String, _
                                            ByVal firstDataRow As Long, _
                                            ByRef outBlockFirst As Long, _
                                            ByRef outBlockLast As Long) As Boolean
    Dim periodFirst As Long
    Dim periodLast As Long
    Dim r As Long
    Dim rowCode As String
    Dim normalizedTargetCode As String
    
    outBlockFirst = 0
    outBlockLast = 0
    
    periodName = NormalizeKVPeriodText(periodName)
    normalizedTargetCode = NormalizeKVCodeForLookup(kvCode)
    
    If normalizedTargetCode = "" Then Exit Function
    If Not PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, periodFirst, periodLast) Then Exit Function
    
    For r = periodFirst To periodLast
        If wsKV.Range("A" & r).MergeCells Then GoTo NextRow
        If Not PID_RowHasKVTableContent(wsKV, r) Then GoTo NextRow
        
        rowCode = NormalizeKVCodeForLookup(CStr(wsKV.Cells(r, "D").Value))
        
        If rowCode = normalizedTargetCode Then
            If outBlockFirst = 0 Then outBlockFirst = r
            outBlockLast = r
        ElseIf outBlockFirst > 0 Then
            Exit For
        End If
        
NextRow:
    Next r
    
    PID_GetKVCodeBlockBounds = (outBlockFirst > 0 And outBlockLast >= outBlockFirst)
End Function


Private Function PID_KVCodeBlockHasHours(ByVal wsKV As Worksheet, _
                                         ByVal blockFirst As Long, _
                                         ByVal blockLast As Long, _
                                         ByVal targetHours As Double) As Boolean
    Dim r As Long
    Dim rowHours As Double
    
    For r = blockFirst To blockLast
        If PID_TryReadDouble(wsKV.Cells(r, "G").Value, rowHours) Then
            If Abs(rowHours - targetHours) < 0.001 Then
                PID_KVCodeBlockHasHours = True
                Exit Function
            End If
        End If
    Next r
End Function


Private Function PID_FindSortedInsertRowInBlock(ByVal wsKV As Worksheet, _
                                               ByVal blockFirst As Long, _
                                               ByVal blockLast As Long, _
                                               ByVal newHours As Double) As Long
    Dim r As Long
    Dim rowHours As Double
    Dim sortDescending As Boolean
    
    sortDescending = PID_BlockUsesDescendingHours(wsKV, blockFirst, blockLast)
    
    For r = blockFirst To blockLast
        If PID_TryReadDouble(wsKV.Cells(r, "G").Value, rowHours) Then
            If sortDescending Then
                If newHours + 0.001 > rowHours Then
                    PID_FindSortedInsertRowInBlock = r
                    Exit Function
                End If
            Else
                If newHours + 0.001 < rowHours Then
                    PID_FindSortedInsertRowInBlock = r
                    Exit Function
                End If
            End If
        Else
            PID_FindSortedInsertRowInBlock = r
            Exit Function
        End If
    Next r
    
    PID_FindSortedInsertRowInBlock = blockLast + 1
End Function


Private Function PID_BlockUsesDescendingHours(ByVal wsKV As Worksheet, _
                                              ByVal blockFirst As Long, _
                                              ByVal blockLast As Long) As Boolean
    Dim firstHours As Double
    Dim lastHours As Double
    
    If blockFirst >= blockLast Then
        PID_BlockUsesDescendingHours = True
        Exit Function
    End If
    
    If PID_TryReadDouble(wsKV.Cells(blockFirst, "G").Value, firstHours) _
       And PID_TryReadDouble(wsKV.Cells(blockLast, "G").Value, lastHours) Then
        PID_BlockUsesDescendingHours = (firstHours >= lastHours)
    Else
        PID_BlockUsesDescendingHours = True
    End If
End Function


Private Sub PID_WriteCustomKVRow(ByVal wsKV As Worksheet, _
                               ByVal targetRow As Long, _
                               ByVal periodName As String, _
                               ByVal kvCode As String, _
                               ByVal firstDataRow As Long, _
                               ByVal newHours As Double, _
                               ByVal newLohn As Variant, _
                               ByVal hasLohn As Boolean)
    Dim validFrom As Variant
    Dim validTo As Variant
    
    If wsKV Is Nothing Then Exit Sub
    If targetRow < 1 Then Exit Sub
    
    PID_GetPeriodValidDates wsKV, periodName, firstDataRow, validFrom, validTo
    
    wsKV.Range("A" & targetRow & ":J" & targetRow).UnMerge
    wsKV.Range("A" & targetRow & ":J" & targetRow).ClearContents
    
    wsKV.Cells(targetRow, "A").Value = periodName
    
    If IsDate(validFrom) Then
        wsKV.Cells(targetRow, "B").Value = CDate(validFrom)
    End If
    
    If IsDate(validTo) Then
        wsKV.Cells(targetRow, "C").Value = CDate(validTo)
    End If
    
    wsKV.Cells(targetRow, "D").Value = kvCode
    wsKV.Cells(targetRow, "E").Value = PID_GetStandardKVGroupByCode(kvCode)
    wsKV.Cells(targetRow, "F").Value = PID_GetStandardKVDurationByCode(kvCode)
    wsKV.Cells(targetRow, "G").Value = newHours
    
    If hasLohn Then
        wsKV.Cells(targetRow, "H").Value = newLohn
    Else
        wsKV.Cells(targetRow, "H").ClearContents
    End If
    
    PID_MarkCustomKVHourRow wsKV, targetRow
End Sub


Private Sub PID_GetPeriodValidDates(ByVal wsKV As Worksheet, _
                                    ByVal periodName As String, _
                                    ByVal firstDataRow As Long, _
                                    ByRef validFrom As Variant, _
                                    ByRef validTo As Variant)
    Dim periodFirst As Long
    Dim periodLast As Long
    Dim r As Long
    
    validFrom = Empty
    validTo = Empty
    
    If wsKV Is Nothing Then Exit Sub
    
    periodName = NormalizeKVPeriodText(periodName)
    If periodName = "" Then Exit Sub
    
    If Not PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, periodFirst, periodLast) Then Exit Sub
    
    For r = periodFirst To periodLast
        If IsDate(wsKV.Cells(r, "B").Value) And IsDate(wsKV.Cells(r, "C").Value) Then
            validFrom = wsKV.Cells(r, "B").Value
            validTo = wsKV.Cells(r, "C").Value
            Exit Sub
        End If
    Next r
End Sub


Private Function PID_GetKVTableFirstDataRow(ByVal wsKV As Worksheet) As Long
    Dim r As Long
    Dim lastRow As Long
    
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then
        PID_GetKVTableFirstDataRow = 4
        Exit Function
    End If
    
    If StrComp(Trim$(CStr(wsKV.Cells(3, "A").Value)), "KV-Periode", vbTextCompare) = 0 Then
        PID_GetKVTableFirstDataRow = 4
        Exit Function
    End If
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    
    For r = 1 To lastRow
        If Not wsKV.Range("A" & r).MergeCells Then
            If Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "" Then
                PID_GetKVTableFirstDataRow = r
                Exit Function
            End If
        End If
    Next r
    
SafeExit:
    PID_GetKVTableFirstDataRow = 4
End Function


Private Function PID_GetStandardKVGroupByCode(ByVal kvCode As String) As String
    Dim normalizedCode As String
    
    normalizedCode = UCase$(NormalizeKVCodeForLookup(kvCode))
    
    If InStr(1, normalizedCode, "BG1", vbTextCompare) > 0 Then
        PID_GetStandardKVGroupByCode = "BG1"
    ElseIf InStr(1, normalizedCode, "BG2", vbTextCompare) > 0 Then
        PID_GetStandardKVGroupByCode = "BG2"
    ElseIf InStr(1, normalizedCode, "BG3", vbTextCompare) > 0 Then
        PID_GetStandardKVGroupByCode = "BG3"
    End If
End Function


Private Function PID_GetStandardKVDurationByCode(ByVal kvCode As String) As String
    Dim normalizedCode As String
    
    normalizedCode = UCase$(NormalizeKVCodeForLookup(kvCode))
    
    If InStr(1, normalizedCode, "_15", vbTextCompare) > 0 Then
        PID_GetStandardKVDurationByCode = "ueber 15 Jahre"
    ElseIf InStr(1, normalizedCode, "_10", vbTextCompare) > 0 Then
        PID_GetStandardKVDurationByCode = "10 bis 15 Jahre"
    ElseIf InStr(1, normalizedCode, "_5", vbTextCompare) > 0 Then
        PID_GetStandardKVDurationByCode = "5 bis 10 Jahre"
    ElseIf InStr(1, normalizedCode, "BASIS", vbTextCompare) > 0 Then
        PID_GetStandardKVDurationByCode = "Basis / bis 5 Jahre"
    End If
End Function


Private Function PID_FormatHoursText(ByVal hoursValue As Double) As String
    PID_FormatHoursText = Format$(hoursValue, "0.00")
End Function


Private Sub PID_MarkCustomKVHourRow(ByVal wsKV As Worksheet, ByVal targetRow As Long)
    On Error Resume Next
    
    If wsKV Is Nothing Then Exit Sub
    If targetRow < 1 Then Exit Sub
    
    wsKV.Cells(targetRow, PID_CUSTOM_KV_HOUR_MARKER_COL).Value = PID_CUSTOM_KV_HOUR_MARKER_VALUE
    wsKV.Columns(PID_CUSTOM_KV_HOUR_MARKER_COL).Hidden = True
    
    Err.Clear
End Sub


Private Function PID_IsCustomKVHourRow(ByVal wsKV As Worksheet, ByVal rowNumber As Long) As Boolean
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Function
    If rowNumber < 1 Then Exit Function
    
    PID_IsCustomKVHourRow = (StrComp(Trim$(CStr(wsKV.Cells(rowNumber, PID_CUSTOM_KV_HOUR_MARKER_COL).Value)), _
        PID_CUSTOM_KV_HOUR_MARKER_VALUE, vbTextCompare) = 0)
    
SafeExit:
End Function


Private Function PID_GetKVCodeBlockRowCountInPeriod(ByVal wsKV As Worksheet, _
                                                      ByVal periodName As String, _
                                                      ByVal kvCode As String, _
                                                      ByVal firstDataRow As Long) As Long
    Dim blockFirst As Long
    Dim blockLast As Long
    
    If PID_GetKVCodeBlockBounds(wsKV, periodName, kvCode, firstDataRow, blockFirst, blockLast) Then
        PID_GetKVCodeBlockRowCountInPeriod = blockLast - blockFirst + 1
    End If
End Function


Private Function PID_GetMinKVCodeBlockRowCountInPeriod(ByVal wsKV As Worksheet, _
                                                       ByVal periodName As String, _
                                                       ByVal firstDataRow As Long) As Long
    Dim periodFirst As Long
    Dim periodLast As Long
    Dim r As Long
    Dim rowCode As String
    Dim normalizedCode As String
    Dim blockFirst As Long
    Dim blockLast As Long
    Dim blockSize As Long
    Dim minSize As Long
    
    periodName = NormalizeKVPeriodText(periodName)
    If periodName = "" Then Exit Function
    
    If Not PID_GetPeriodRowBounds(wsKV, periodName, firstDataRow, periodFirst, periodLast) Then Exit Function
    
    minSize = 0
    blockFirst = 0
    blockLast = 0
    
    For r = periodFirst To periodLast
        If wsKV.Range("A" & r).MergeCells Then GoTo NextRow
        If Not PID_RowHasKVTableContent(wsKV, r) Then GoTo NextRow
        
        rowCode = Trim$(CStr(wsKV.Cells(r, "D").Value))
        normalizedCode = NormalizeKVCodeForLookup(rowCode)
        
        If normalizedCode = "" Then GoTo NextRow
        
        If blockFirst = 0 Then
            blockFirst = r
            blockLast = r
        ElseIf StrComp(NormalizeKVCodeForLookup(CStr(wsKV.Cells(blockFirst, "D").Value)), normalizedCode, vbTextCompare) = 0 Then
            blockLast = r
        Else
            blockSize = blockLast - blockFirst + 1
            If minSize = 0 Or blockSize < minSize Then minSize = blockSize
            blockFirst = r
            blockLast = r
        End If
        
NextRow:
    Next r
    
    If blockFirst > 0 Then
        blockSize = blockLast - blockFirst + 1
        If minSize = 0 Or blockSize < minSize Then minSize = blockSize
    End If
    
    PID_GetMinKVCodeBlockRowCountInPeriod = minSize
End Function


Private Function PID_BlockHasMarkedCustomHourRows(ByVal wsKV As Worksheet, _
                                                  ByVal blockFirst As Long, _
                                                  ByVal blockLast As Long) As Boolean
    Dim r As Long
    
    For r = blockFirst To blockLast
        If PID_IsCustomKVHourRow(wsKV, r) Then
            PID_BlockHasMarkedCustomHourRows = True
            Exit Function
        End If
    Next r
End Function


Private Function PID_BlockHasDeletableCustomHours(ByVal wsKV As Worksheet, _
                                                  ByVal periodName As String, _
                                                  ByVal kvCode As String, _
                                                  ByVal firstDataRow As Long, _
                                                  ByVal blockFirst As Long, _
                                                  ByVal blockLast As Long, _
                                                  ByVal blockRowCount As Long) As Boolean
    Dim deletableHours As Collection
    
    Set deletableHours = PID_CollectDeletableHoursInBlock(wsKV, periodName, kvCode, _
        firstDataRow, blockFirst, blockLast)
    
    PID_BlockHasDeletableCustomHours = (Not deletableHours Is Nothing And deletableHours.Count > 0)
End Function


Private Function PID_CollectDeletableHoursInBlock(ByVal wsKV As Worksheet, _
                                                  ByVal periodName As String, _
                                                  ByVal kvCode As String, _
                                                  ByVal firstDataRow As Long, _
                                                  ByVal blockFirst As Long, _
                                                  ByVal blockLast As Long) As Collection
    Dim deletableHours As Collection
    Dim baselinePeriod As String
    Dim baselineFirst As Long
    Dim baselineLast As Long
    Dim baselineHours As Collection
    Dim allHours As Collection
    Dim r As Long
    Dim rowHours As Double
    Dim hoursKey As String
    Dim i As Long
    Dim hourValue As Double
    Dim baselineCount As Long
    Dim currentCount As Long
    
    Set deletableHours = New Collection
    
    If PID_BlockHasMarkedCustomHourRows(wsKV, blockFirst, blockLast) Then
        For r = blockFirst To blockLast
            If PID_IsCustomKVHourRow(wsKV, r) Then
                If PID_TryReadDouble(wsKV.Cells(r, "G").Value, rowHours) Then
                    hoursKey = PID_FormatHoursText(rowHours)
                    On Error Resume Next
                    deletableHours.Add rowHours, hoursKey
                    Err.Clear
                End If
            End If
        Next r
        
        Set PID_CollectDeletableHoursInBlock = deletableHours
        Exit Function
    End If
    
    baselinePeriod = GetBottomKVPeriod(wsKV, firstDataRow)
    If baselinePeriod <> "" _
       And PID_GetKVCodeBlockBounds(wsKV, baselinePeriod, kvCode, firstDataRow, baselineFirst, baselineLast) Then
        
        Set baselineHours = PID_CollectHoursInBlock(wsKV, baselineFirst, baselineLast)
        Set allHours = PID_CollectHoursInBlock(wsKV, blockFirst, blockLast)
        
        If Not allHours Is Nothing Then
            For i = 1 To allHours.Count
                hourValue = CDbl(allHours(i))
                hoursKey = PID_FormatHoursText(hourValue)
                baselineCount = PID_CountHourOccurrencesInBlock(wsKV, baselineFirst, baselineLast, hourValue)
                currentCount = PID_CountHourOccurrencesInBlock(wsKV, blockFirst, blockLast, hourValue)
                
                If currentCount > baselineCount Then
                    On Error Resume Next
                    deletableHours.Add hourValue, hoursKey
                    Err.Clear
                End If
            Next i
        End If
    End If
    
    Set PID_CollectDeletableHoursInBlock = deletableHours
End Function


Private Function PID_CountHourOccurrencesInBlock(ByVal wsKV As Worksheet, _
                                                 ByVal blockFirst As Long, _
                                                 ByVal blockLast As Long, _
                                                 ByVal targetHours As Double) As Long
    Dim r As Long
    Dim rowHours As Double
    
    For r = blockFirst To blockLast
        If PID_TryReadDouble(wsKV.Cells(r, "G").Value, rowHours) Then
            If Abs(rowHours - targetHours) < 0.001 Then
                PID_CountHourOccurrencesInBlock = PID_CountHourOccurrencesInBlock + 1
            End If
        End If
    Next r
End Function


Private Function AskForKVHoursInBlockSelection(ByVal wsKV As Worksheet, _
                                               ByVal blockFirst As Long, _
                                               ByVal blockLast As Long, _
                                               ByRef outHours As Double, _
                                               Optional ByVal dialogTitle As String = "", _
                                               Optional ByVal stepText As String = "", _
                                               Optional ByVal hourOptions As Collection) As Boolean
    Dim hourList As Collection
    Dim promptText As String
    Dim inputText As String
    Dim selectedIndex As Long
    Dim titleText As String
    Dim selectedHours As Double
    
    If dialogTitle = "" Then dialogTitle = PID_KVTxtStundeLoeschen()
    
    Set hourList = hourOptions
    
    If hourList Is Nothing Or hourList.Count = 0 Then
        MsgBox "Keine l" & PID_KVTxtOe() & "schbare eigene Stunde in dieser Gruppe." & vbCrLf & vbCrLf & _
               "Nur mit ""2) Eigene Stunden"" hinzugef" & PID_KVTxtUe() & "gte Zeilen " & _
               "oder eine neuere Stundenzahl gegen" & PID_KVTxtUe() & "ber der " & PID_KVTxtAe() & "ltesten KV-Periode.", _
               vbExclamation, dialogTitle
        Exit Function
    End If
    
    promptText = ""
    If stepText <> "" Then
        promptText = stepText & vbCrLf & vbCrLf
    End If
    
    promptText = promptText & "Liste:" & vbCrLf & vbCrLf
    promptText = promptText & PID_BuildNumberedHoursListFromCollection(hourList)
    promptText = promptText & vbCrLf & "Nur die Nummer eintippen und OK klicken."
    
    titleText = dialogTitle
    If stepText <> "" Then titleText = dialogTitle & " - " & stepText
    
    inputText = InputBox(Prompt:=promptText, Title:=titleText, Default:="1")
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Das war keine g" & PID_KVTxtUe() & "ltige Nummer.", vbExclamation, dialogTitle
        Exit Function
    End If
    
    selectedIndex = CLng(inputText)
    
    If selectedIndex < 1 Or selectedIndex > hourList.Count Then
        MsgBox "Diese Nummer steht nicht in der Liste.", vbExclamation, dialogTitle
        Exit Function
    End If
    
    selectedHours = CDbl(hourList(selectedIndex))
    outHours = selectedHours
    AskForKVHoursInBlockSelection = True
End Function


Private Function PID_BuildNumberedHoursListFromCollection(ByVal hourOptions As Collection) As String
    Dim resultText As String
    Dim i As Long
    Dim hoursValue As Double
    
    If hourOptions Is Nothing Then Exit Function
    
    For i = 1 To hourOptions.Count
        hoursValue = CDbl(hourOptions(i))
        resultText = resultText & CStr(i) & " = " & PID_FormatHoursText(hoursValue) & " Stunden" & vbCrLf
    Next i
    
    PID_BuildNumberedHoursListFromCollection = resultText
End Function


Private Function PID_CollectHoursInBlock(ByVal wsKV As Worksheet, _
                                        ByVal blockFirst As Long, _
                                        ByVal blockLast As Long) As Collection
    Dim hoursList As Collection
    Dim r As Long
    Dim rowHours As Double
    Dim hoursKey As String
    
    Set hoursList = New Collection
    
    If wsKV Is Nothing Then Exit Function
    If blockFirst <= 0 Or blockLast < blockFirst Then Exit Function
    
    For r = blockFirst To blockLast
        If PID_TryReadDouble(wsKV.Cells(r, "G").Value, rowHours) Then
            hoursKey = PID_FormatHoursText(rowHours)
            On Error Resume Next
            hoursList.Add rowHours, hoursKey
            Err.Clear
        End If
    Next r
    
    Set PID_CollectHoursInBlock = hoursList
End Function


Private Function PID_FindHourRowInBlock(ByVal wsKV As Worksheet, _
                                       ByVal blockFirst As Long, _
                                       ByVal blockLast As Long, _
                                       ByVal targetHours As Double) As Long
    Dim r As Long
    Dim rowHours As Double
    
    For r = blockFirst To blockLast
        If PID_TryReadDouble(wsKV.Cells(r, "G").Value, rowHours) Then
            If Abs(rowHours - targetHours) < 0.001 Then
                PID_FindHourRowInBlock = r
                Exit Function
            End If
        End If
    Next r
End Function


Private Function PID_FindDeletableHourRowInBlock(ByVal wsKV As Worksheet, _
                                                 ByVal blockFirst As Long, _
                                                 ByVal blockLast As Long, _
                                                 ByVal targetHours As Double) As Long
    Dim r As Long
    Dim rowHours As Double
    
    For r = blockFirst To blockLast
        If PID_IsCustomKVHourRow(wsKV, r) Then
            If PID_TryReadDouble(wsKV.Cells(r, "G").Value, rowHours) Then
                If Abs(rowHours - targetHours) < 0.001 Then
                    PID_FindDeletableHourRowInBlock = r
                    Exit Function
                End If
            End If
        End If
    Next r
    
    For r = blockLast To blockFirst Step -1
        If PID_TryReadDouble(wsKV.Cells(r, "G").Value, rowHours) Then
            If Abs(rowHours - targetHours) < 0.001 Then
                PID_FindDeletableHourRowInBlock = r
                Exit Function
            End If
        End If
    Next r
End Function


Private Function PID_CountMonthSheetUsageForKVHour(ByVal kvCode As String, _
                                                   ByVal targetHours As Double) As Long
    Dim monthName As Variant
    Dim ws As Worksheet
    Dim r As Long
    Dim rowCode As String
    Dim rowHours As Double
    Dim normalizedTargetCode As String
    
    normalizedTargetCode = NormalizeKVCodeForLookup(kvCode)
    If normalizedTargetCode = "" Then Exit Function
    
    For Each monthName In PID_MonthNames()
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(monthName))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            If PID_IsWorkerMonthSheet(ws) Then
                For r = PID_FIRST_ROW To PID_LAST_ROW
                    rowCode = NormalizeKVCodeForLookup(CStr(ws.Cells(r, "E").Value))
                    If rowCode = normalizedTargetCode Then
                        If PID_TryReadDouble(ws.Cells(r, "F").Value, rowHours) Then
                            If Abs(rowHours - targetHours) < 0.001 Then
                                PID_CountMonthSheetUsageForKVHour = PID_CountMonthSheetUsageForKVHour + 1
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    Next monthName
End Function

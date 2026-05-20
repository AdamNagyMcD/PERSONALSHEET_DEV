Attribute VB_Name = "mod_AddNewKVPeriodOnTop"
Option Explicit
Private Const PID_ADD_PERIOD_BUTTON_NAME As String = "btn_AddNewKVPeriodOnTop"
Private Const PID_ADD_CUSTOM_HOURS_BUTTON_NAME As String = "btn_AddCustomKVMonatsstunden"
Private Const PID_DELETE_PERIODS_BUTTON_NAME As String = "btn_DeleteKVPeriods"
Private Const PID_HELP_BUTTON_NAME As String = "btn_LOHNTABELLE_TESTHelp"
Private Const PID_KV_CODE_COUNT As Long = 12

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
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    
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
        MsgBox "Keine KV-Daten in LOHNTABELLE_TEST gefunden.", _
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
        MsgBox "Die Vorlage fuer den neuen KV-Zeitraum konnte nicht gefunden werden.", _
               vbExclamation, "Neuer KV-Zeitraum"
        GoTo CleanExit
    End If
    
    templateRowCount = templateLastRow - templateFirstRow + 1
    
    If templateRowCount <= 0 Then
        MsgBox "Die Vorlage enthaelt keine gueltigen Datenzeilen.", _
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
    
    MsgBox "Der neue KV-Zeitraum wurde eingefuegt:" & vbCrLf & vbCrLf & _
           newPeriod & vbCrLf & vbCrLf & _
           "Bitte jetzt Monatslohn (Spalte H) in der neuen Periode erfassen." & vbCrLf & vbCrLf & _
           PID_GetKVTeamAfterChangeHint(), _
           vbInformation, "Neue Periode"

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
    PID_EnsureLOHNTABELLE_TESTButtons
End Sub


Public Sub ShowLOHNTABELLE_TESTButtonHelp()
    MsgBox PID_GetLOHNTABELLE_TESTTeamHelpText(), vbInformation, "Hilfe - LOHNTABELLE_TEST"
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
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    firstDataRow = PID_GetKVTableFirstDataRow(wsKV)
    
    Set allPeriods = PID_CollectKVPeriods(wsKV, firstDataRow)
    If allPeriods Is Nothing Or allPeriods.Count = 0 Then
        MsgBox "Keine KV-Zeitraeume gefunden.", vbExclamation, "Alte Periode loeschen"
        Exit Sub
    End If
    
    If allPeriods.Count <= 1 Then
        MsgBox "Es ist nur ein KV-Zeitraum vorhanden." & vbCrLf & vbCrLf & _
               "Mindestens ein Zeitraum muss erhalten bleiben. Loeschen ist nicht moeglich.", _
               vbInformation, "Alte Periode loeschen"
        Exit Sub
    End If
    
    periodToDelete = AskForKVPeriodSelection( _
        wsKV, firstDataRow, _
        "Alte Periode loeschen", _
        "Schritt 1 von 2 - Welche Periode?", _
        "Nur EINE Nummer eingeben (z.B. 2)." & vbCrLf & _
        "Nur sehr alte Perioden loeschen, die niemand mehr braucht." & vbCrLf & _
        "Abbrechen oder leer lassen = nichts aendern.")
    
    If periodToDelete = "" Then Exit Sub
    
    confirmText = "Wirklich loeschen?" & vbCrLf & vbCrLf & _
                  periodToDelete & vbCrLf & vbCrLf & _
                  "NEIN = abbrechen (empfohlen bei Unsicherheit)." & vbCrLf & _
                  "JA = endgueltig loeschen."
    
    answer = MsgBox(confirmText, vbCritical + vbYesNo, "Schritt 2 von 2 - Letzte Rueckfrage")
    If answer <> vbYes Then
        MsgBox "Loeschen abgebrochen. Es wurde nichts geaendert.", vbInformation, "Alte Periode loeschen"
        Exit Sub
    End If
    
    periodFirstRow = FindFirstRowOfPeriod(wsKV, periodToDelete, firstDataRow)
    periodLastRow = FindLastRowOfPeriod(wsKV, periodToDelete, firstDataRow)
    
    If periodFirstRow <= 0 Or periodLastRow < periodFirstRow Then
        MsgBox "Der gewaehlte Zeitraum konnte nicht gefunden werden.", vbExclamation, "Alte Periode loeschen"
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
    PID_EnsureLOHNTABELLE_TESTButtons
    
    MarkAllKVDropdownsDirty
    MarkAllKVLohnDirty
    
    MsgBox "Geloescht: " & periodToDelete & vbCrLf & vbCrLf & PID_GetKVTeamAfterChangeHint(), _
           vbInformation, "Alte Periode loeschen"
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
    
    MsgBox "Fehler beim Loeschen:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Alte Periode loeschen"
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
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    firstDataRow = PID_GetKVTableFirstDataRow(wsKV)
    
    If PID_GetKVTableLastRow(wsKV, firstDataRow) < firstDataRow Then
        MsgBox "Keine KV-Daten in LOHNTABELLE_TEST gefunden.", _
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
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    wsKV.Rows(insertRow).Insert Shift:=xlDown
    
    PID_WriteCustomKVRow wsKV, insertRow, selectedPeriod, selectedKVCode, firstDataRow, newHours, newLohn, hasLohn
    
    FormatKVPeriodArea wsKV
    PID_EnsureLOHNTABELLE_TESTButtons
    
    MarkAllKVDropdownsDirty
    MarkAllKVLohnDirty
    
    MsgBox "Fertig - eigene Stunden eingefuegt:" & vbCrLf & vbCrLf & _
           "Zeitraum: " & selectedPeriod & vbCrLf & _
           "Gruppe: " & selectedKVCode & vbCrLf & _
           "Stunden: " & PID_FormatHoursText(newHours) & vbCrLf & vbCrLf & _
           PID_GetKVTeamAfterChangeHint(), _
           vbInformation, "Eigene Stunden"
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
    
    MsgBox "Fehler bei AddCustomKVMonatsstunden:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Individuelle Monatsstunden"
End Sub


Private Sub PID_EnsureLOHNTABELLE_TESTButtons()
    Dim wsKV As Worksheet
    Dim btn As Shape
    Dim wasProtected As Boolean
    Dim buttonLeft As Double
    Dim buttonWidth As Double
    Dim buttonTop As Double
    Dim buttonHeight As Double
    Dim buttonGap As Double
    Dim rowHeight As Double
    Dim totalButtonsHeight As Double
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    wsKV.Shapes(PID_ADD_PERIOD_BUTTON_NAME).Delete
    wsKV.Shapes(PID_ADD_CUSTOM_HOURS_BUTTON_NAME).Delete
    wsKV.Shapes(PID_DELETE_PERIODS_BUTTON_NAME).Delete
    wsKV.Shapes(PID_HELP_BUTTON_NAME).Delete
    On Error GoTo SafeExit
    
    PID_ConfigureLOHNTABELLE_TESTHeaderLayout wsKV
    
    buttonLeft = wsKV.Range("I2").Left + 1
    buttonWidth = wsKV.Range("I2:J2").Width - 2
    buttonHeight = 15
    buttonGap = 2
    totalButtonsHeight = (4 * buttonHeight) + (3 * buttonGap)
    rowHeight = wsKV.Rows(2).RowHeight
    
    If rowHeight < totalButtonsHeight + 6 Then
        wsKV.Rows(2).RowHeight = totalButtonsHeight + 6
        rowHeight = wsKV.Rows(2).RowHeight
    End If
    
    buttonTop = wsKV.Range("I2").Top + ((rowHeight - totalButtonsHeight) / 2)
    
    Set btn = wsKV.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                   Left:=buttonLeft, _
                                   Top:=buttonTop, _
                                   Width:=buttonWidth, _
                                   Height:=buttonHeight)
    
    btn.Name = PID_ADD_PERIOD_BUTTON_NAME
    btn.TextFrame.Characters.Text = "1) Neue Periode"
    btn.OnAction = "AddNewKVPeriodOnTop"
    PID_ApplyLOHNTABELLE_TESTButtonStyle btn, RGB(54, 96, 146), RGB(33, 64, 99)
    
    Set btn = wsKV.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                   Left:=buttonLeft, _
                                   Top:=buttonTop + (buttonHeight + buttonGap), _
                                   Width:=buttonWidth, _
                                   Height:=buttonHeight)
    
    btn.Name = PID_ADD_CUSTOM_HOURS_BUTTON_NAME
    btn.TextFrame.Characters.Text = "2) Eigene Stunden"
    btn.OnAction = "AddCustomKVMonatsstunden"
    PID_ApplyLOHNTABELLE_TESTButtonStyle btn, RGB(84, 130, 53), RGB(56, 87, 35)
    
    Set btn = wsKV.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                   Left:=buttonLeft, _
                                   Top:=buttonTop + (2 * (buttonHeight + buttonGap)), _
                                   Width:=buttonWidth, _
                                   Height:=buttonHeight)
    
    btn.Name = PID_DELETE_PERIODS_BUTTON_NAME
    btn.TextFrame.Characters.Text = "3) Alte Periode loeschen"
    btn.OnAction = "DeleteSelectedKVPeriods"
    PID_ApplyLOHNTABELLE_TESTButtonStyle btn, RGB(192, 80, 77), RGB(132, 46, 43)
    
    Set btn = wsKV.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                   Left:=buttonLeft, _
                                   Top:=buttonTop + (3 * (buttonHeight + buttonGap)), _
                                   Width:=buttonWidth, _
                                   Height:=buttonHeight)
    
    btn.Name = PID_HELP_BUTTON_NAME
    btn.TextFrame.Characters.Text = "Hilfe"
    btn.OnAction = "ShowLOHNTABELLE_TESTButtonHelp"
    PID_ApplyLOHNTABELLE_TESTButtonStyle btn, RGB(120, 120, 120), RGB(80, 80, 80)
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
End Sub


Private Sub PID_ApplyLOHNTABELLE_TESTButtonStyle(ByVal btn As Shape, _
                                                 ByVal fillColor As Long, _
                                                 ByVal lineColor As Long)
    On Error GoTo SafeExit
    
    If btn Is Nothing Then Exit Sub
    
    btn.Fill.ForeColor.RGB = fillColor
    btn.Line.ForeColor.RGB = lineColor
    btn.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Size = 8
    btn.TextFrame.WordWrap = msoTrue
    btn.TextFrame.VerticalAnchor = msoAnchorMiddle
    btn.TextFrame.HorizontalAnchor = msoAnchorCenter
    btn.Placement = xlFreeFloating
    
SafeExit:
End Sub


Private Sub PID_ConfigureLOHNTABELLE_TESTHeaderLayout(ByVal wsKV As Worksheet)
    On Error GoTo SafeExit
    
    If wsKV Is Nothing Then Exit Sub
    
    On Error Resume Next
    wsKV.Range("A1:J2").UnMerge
    On Error GoTo SafeExit
    
    wsKV.Range("A1:J1").Merge
    wsKV.Range("A1").HorizontalAlignment = xlCenter
    wsKV.Range("A1").VerticalAlignment = xlCenter
    
    wsKV.Range("A2:H2").Merge
    wsKV.Range("A2").WrapText = True
    wsKV.Range("A2").VerticalAlignment = xlCenter
    wsKV.Range("A2").HorizontalAlignment = xlLeft
    
    wsKV.Range("I2:J2").ClearContents
    wsKV.Range("I2:J2").Interior.Pattern = xlNone
    
    wsKV.Rows(2).AutoFit
    If wsKV.Rows(2).RowHeight < 74 Then wsKV.Rows(2).RowHeight = 74
    
SafeExit:
End Sub


Public Sub FixLOHNTABELLE_TEST_HeaderText()
    FixLOHNTABELLE_TEST_HeaderTextIfNeeded True
End Sub


Public Sub FixLOHNTABELLE_TEST_HeaderTextIfNeeded(Optional ByVal forceFormulaRepair As Boolean = False)
    Dim wsKV As Worksheet
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
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


Public Sub FixLOHNTABELLE_TEST_StatusFormulas()
    Dim wsKV As Worksheet
    Dim wasProtected As Boolean
    
    On Error GoTo CleanFail
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    FormatKVPeriodArea wsKV
    
    MsgBox "Status- und Pruefungsformeln in LOHNTABELLE_TEST wurden wiederhergestellt.", _
           vbInformation, "LOHNTABELLE_TEST"
    
CleanExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    Exit Sub
    
CleanFail:
    MsgBox "Fehler bei FixLOHNTABELLE_TEST_StatusFormulas:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "LOHNTABELLE_TEST"
    Resume CleanExit
End Sub


Public Sub RebuildLOHNTABELLE_TEST()
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
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    firstDataRow = 4
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    cleanupLastRow = PID_GetSheetCleanupLastRow(wsKV, lastRow)
    
    If lastRow < firstDataRow Then
        MsgBox "Keine Daten in LOHNTABELLE_TEST gefunden.", vbExclamation, "LOHNTABELLE_TEST neu aufbauen"
        Exit Sub
    End If
    
    keepPeriod = GetBottomKVPeriod(wsKV, firstDataRow)
    If keepPeriod = "" Then
        MsgBox "Kein gueltiger KV-Zeitraum in Spalte A gefunden.", vbExclamation, "LOHNTABELLE_TEST neu aufbauen"
        Exit Sub
    End If
    
    periodFirstRow = FindFirstRowOfPeriod(wsKV, keepPeriod, firstDataRow)
    periodLastRow = FindLastRowOfPeriod(wsKV, keepPeriod, firstDataRow)
    
    If periodFirstRow = 0 Or periodLastRow = 0 Then
        MsgBox "Der unterste KV-Zeitraum konnte nicht gelesen werden.", vbExclamation, "LOHNTABELLE_TEST neu aufbauen"
        Exit Sub
    End If
    
    periodRowCount = periodLastRow - periodFirstRow + 1
    If periodRowCount <= 0 Then Exit Sub
    
    answer = MsgBox( _
        "LOHNTABELLE_TEST wird neu aufgebaut." & vbCrLf & vbCrLf & _
        "Behalten wird nur der unterste Zeitraum:" & vbCrLf & _
        keepPeriod & vbCrLf & vbCrLf & _
        "Alle weiteren Test-Zeitraeume werden geloescht." & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbQuestion + vbYesNo, _
        "LOHNTABELLE_TEST neu aufbauen" _
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
    MarkKVDropdownsDirty
    MarkKVLohnDirty
    
    On Error Resume Next
    PID_ResetHourOverrideLog
    On Error GoTo CleanFail
    
    EnsureAddNewKVPeriodButton
    
    MsgBox "LOHNTABELLE_TEST wurde neu aufgebaut. Zeitraum aktiv: " & keepPeriod, _
           vbInformation, "LOHNTABELLE_TEST neu aufgebaut"
    
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
    MsgBox "Fehler bei RebuildLOHNTABELLE_TEST:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "LOHNTABELLE_TEST neu aufbauen"
    Resume CleanExit
End Sub


Public Sub RestoreLOHNTABELLE_TESTBase2025_2026()
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
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
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
        "LOHNTABELLE_TEST wird auf den Basiszeitraum zurueckgesetzt:" & vbCrLf & vbCrLf & _
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
    MarkKVLohnDirty
    
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
    MsgBox "Fehler bei RestoreLOHNTABELLE_TESTBase2025_2026:" & vbCrLf & _
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
        MsgBox "Ungueltige Eingabe. Bitte nur das Startjahr eingeben (z.B. 2026).", _
               vbExclamation, "Neuer KV-Zeitraum"
        Exit Function
    End If
    
    AskForKVStartYear = CLng(inputText)
    
    If AskForKVStartYear < 2000 Or AskForKVStartYear > 2100 Then
        MsgBox "Das Startjahr liegt ausserhalb des erlaubten Bereichs (2000-2100).", _
               vbExclamation, "Neuer KV-Zeitraum"
        AskForKVStartYear = 0
    End If
End Function


Private Function AskForSchemaCount(ByVal defaultCount As Long) As Long
    Dim inputText As String
    
    inputText = InputBox( _
        Prompt:="Wie viele Vertraege/Monatsstunden-Zeilen pro KV-Code sollen erzeugt werden?" & vbCrLf & vbCrLf & _
                "Beispiel: 13 (wie bisher).", _
        Title:="Vertragsanzahl pro KV-Code", _
        Default:=CStr(defaultCount) _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Ungueltige Eingabe. Bitte eine ganze Zahl eingeben (z.B. 13).", _
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
    
    wsKV.Range("A2").Value = _
        "Wichtig: Alte Perioden nicht von Hand loeschen. Dafuer Button ""3) Alte Periode loeschen"" nutzen. " & _
        "Neue Werte ab Mai = Button ""1) Neue Periode"". " & _
        "Sondervertrag-Stunden = Button ""2) Eigene Stunden"". " & _
        "Bei Unsicherheit: Button ""Hilfe"". Nur Zeilen mit Status OK verwenden."
    
    PID_ConfigureLOHNTABELLE_TESTHeaderLayout wsKV
End Sub


Private Sub PID_NormalizeKVTableHeader(ByVal wsKV As Worksheet)
    If wsKV Is Nothing Then Exit Sub
    
    wsKV.Range("A3").Value = "KV-Periode"
    wsKV.Range("B3").Value = "Gueltig ab"
    wsKV.Range("C3").Value = "Gueltig bis"
    wsKV.Range("D3").Value = "KV-Code"
    wsKV.Range("E3").Value = "KV-Gruppe"
    wsKV.Range("F3").Value = "Beschaeftigungsdauer"
    wsKV.Range("G3").Value = "Monatsstunden"
    wsKV.Range("H3").Value = "Monatslohn"
    wsKV.Range("I3").Value = "Status"
    wsKV.Range("J3").Value = "Pruefung"
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
    
    On Error Resume Next
    If lastDataRow >= firstDataRow Then
        wsKV.Range("K" & firstDataRow & ":XFD" & lastDataRow).Clear
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


Public Sub CleanupLOHNTABELLE_TESTTrailingArea()
    Dim wsKV As Worksheet
    Dim firstDataRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
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
    
    newArea.Clear
    
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
        titleText = titleText & "   |   gueltig von " & Format$(CDate(validFrom), "dd.mm.yyyy") & _
                    " bis " & Format$(CDate(validTo), "dd.mm.yyyy")
    End If
    
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).UnMerge
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).ClearContents
    wsKV.Range("A" & rowNumber & ":J" & rowNumber).Merge
    
    wsKV.Cells(rowNumber, "A").Value = titleText
    
    With wsKV.Range("A" & rowNumber & ":J" & rowNumber)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(235, 235, 235)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
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
        .Range("A" & firstDataRow & ":J" & lastRow).Borders.Color = RGB(150, 150, 150)
        
        .Range("A" & firstDataRow & ":A" & lastRow).NumberFormat = "@"
        .Range("G" & firstDataRow & ":G" & lastRow).NumberFormatLocal = "0,00"
        
        PID_ApplyEuroNumberFormat .Range("H" & firstDataRow & ":H" & lastRow)
        
        .Columns("A").ColumnWidth = 16
        .Columns("D").ColumnWidth = 14
        .Columns("G").ColumnWidth = 13
        .Columns("H").ColumnWidth = 14
        PID_ConfigureKVStatusColumnWidths wsKV, firstDataRow, lastRow
    End With
    
    ' Status- und Pruefungsformeln auf allen gueltigen Datenzeilen wiederherstellen.
    PID_ApplyKVStatusFormulas wsKV, firstDataRow, lastRow
    
    ' Eingabefelder fuer Monatsstunden/Monatslohn muessen editierbar bleiben.
    PID_ConfigureKVInputCellLocks wsKV, firstDataRow, lastRow
    
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
    
    For r = firstRow To lastRow
        Set rowRange = wsKV.Range("A" & r & ":J" & r)
        
        If wsKV.Range("A" & r).MergeCells Then
            ' Titelzeile wird in PID_WriteKVPeriodTitleRow formatiert.
        ElseIf Trim$(CStr(wsKV.Cells(r, "D").Value)) <> "" Then
            rowRange.Font.Bold = False
            rowRange.Font.Size = 10
            rowRange.VerticalAlignment = xlCenter
            rowRange.HorizontalAlignment = xlCenter
        End If
    Next r
    
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


Public Sub ApplyKVVisualGrouping()
    Dim wsKV As Worksheet
    Dim lastRow As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set wsKV = ThisWorkbook.Worksheets("LOHNTABELLE_TEST")
    If wsKV Is Nothing Then Exit Sub
    
    lastRow = wsKV.Cells(wsKV.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Sub
    
    wasProtected = wsKV.ProtectContents
    
    On Error Resume Next
    wsKV.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo SafeExit
    
    PID_ApplyKVVisualGrouping wsKV, 4, lastRow
    
SafeExit:
    On Error Resume Next
    If wasProtected Then
        wsKV.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
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
        
        ' Soft BG color blocks.
        Select Case True
            Case InStr(1, groupText, "BG1", vbTextCompare) > 0
                rowRange.Interior.Color = RGB(242, 248, 255)
            Case InStr(1, groupText, "BG2", vbTextCompare) > 0
                rowRange.Interior.Color = RGB(241, 250, 241)
            Case InStr(1, groupText, "BG3", vbTextCompare) > 0
                rowRange.Interior.Color = RGB(255, 246, 237)
        End Select
        
        ' Strong separator when a new KV period starts.
        If currentPeriod <> "" Then
            If currentPeriod <> prevPeriod Then
                rowRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                rowRange.Borders(xlEdgeTop).Weight = xlMedium
                rowRange.Borders(xlEdgeTop).Color = RGB(90, 90, 90)
                
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
                rowRange.Borders(xlEdgeTop).Color = RGB(120, 120, 120)
            End If
            prevCode = currentCode
        End If
        
NextRow:
    Next r
    
    ' Aussenrahmen der Tabelle am Ende explizit verstaerken.
    With wsKV.Range("A" & firstRow & ":J" & lastRow)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = RGB(90, 90, 90)
        
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeRight).Color = RGB(90, 90, 90)
        
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = RGB(90, 90, 90)
    End With
    
SafeExit:
End Sub


Private Function AskForKVPeriodSelection(ByVal wsKV As Worksheet, ByVal firstDataRow As Long) As String
    Dim periods As Collection
    Dim promptText As String
    Dim inputText As String
    Dim selectedIndex As Long
    Dim defaultIndex As Long
    
    Set periods = PID_CollectKVPeriods(wsKV, firstDataRow)
    
    If periods Is Nothing Then Exit Function
    If periods.Count = 0 Then
        MsgBox "Kein gueltiger KV-Zeitraum gefunden.", vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    defaultIndex = 1
    promptText = "Bitte Nummer des KV-Zeitraums waehlen:" & vbCrLf & vbCrLf
    promptText = promptText & PID_BuildNumberedListFromCollection(periods)
    promptText = promptText & vbCrLf & "Standard: 1 = aktuellster Zeitraum oben"
    
    inputText = InputBox(Prompt:=promptText, Title:="Individuelle Monatsstunden", Default:=CStr(defaultIndex))
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Ungueltige Eingabe. Bitte die Nummer des KV-Zeitraums eingeben.", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    selectedIndex = CLng(inputText)
    
    If selectedIndex < 1 Or selectedIndex > periods.Count Then
        MsgBox "Die gewaehlte Nummer liegt ausserhalb der verfuegbaren Zeitraeume.", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    AskForKVPeriodSelection = CStr(periods(selectedIndex))
End Function


Private Function AskForKVCodeSelection() As String
    Dim promptText As String
    Dim inputText As String
    Dim selectedIndex As Long
    Dim i As Long
    
    promptText = "Bitte Nummer des KV-Codes waehlen:" & vbCrLf & vbCrLf
    
    For i = 1 To PID_KV_CODE_COUNT
        promptText = promptText & CStr(i) & " = " & PID_GetStandardKVCodeByIndex(i) & vbCrLf
    Next i
    
    inputText = InputBox(Prompt:=promptText, Title:="Individuelle Monatsstunden", Default:="1")
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not IsNumeric(inputText) Then
        MsgBox "Ungueltige Eingabe. Bitte die Nummer des KV-Codes eingeben.", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    selectedIndex = CLng(inputText)
    
    If selectedIndex < 1 Or selectedIndex > PID_KV_CODE_COUNT Then
        MsgBox "Die gewaehlte Nummer liegt ausserhalb der verfuegbaren KV-Codes.", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    AskForKVCodeSelection = PID_GetStandardKVCodeByIndex(selectedIndex)
End Function


Private Function AskForCustomMonatsstunden(ByRef outHours As Double) As Boolean
    Dim inputText As String
    
    inputText = InputBox( _
        Prompt:="Bitte individuelle Monatsstunden eingeben." & vbCrLf & vbCrLf & _
                "Beispiel: 160 oder 160,00", _
        Title:="Individuelle Monatsstunden", _
        Default:="" _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not PID_TryReadDouble(inputText, outHours) Then
        MsgBox "Ungueltige Monatsstunden. Bitte eine positive Zahl eingeben.", _
               vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    If outHours <= 0# Then
        MsgBox "Monatsstunden muessen groesser als 0 sein.", vbExclamation, "Individuelle Monatsstunden"
        Exit Function
    End If
    
    AskForCustomMonatsstunden = True
End Function


Private Function AskForOptionalMonatslohn(ByRef outLohn As Variant) As Boolean
    Dim inputText As String
    Dim lohnValue As Double
    
    inputText = InputBox( _
        Prompt:="Optional: Monatslohn eingeben." & vbCrLf & vbCrLf & _
                "Leer lassen, wenn der Lohn spaeter manuell erfasst wird.", _
        Title:="Individuelle Monatsstunden", _
        Default:="" _
    )
    
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    If Not PID_TryReadDouble(inputText, lohnValue) Then
        MsgBox "Ungueltiger Monatslohn. Bitte eine Zahl eingeben oder leer lassen.", _
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


Private Function AskForKVPeriodsToDelete(ByVal wsKV As Worksheet, ByVal firstDataRow As Long) As Collection
    Dim periods As Collection
    Dim promptText As String
    Dim inputText As String
    
    Set periods = PID_CollectKVPeriods(wsKV, firstDataRow)
    
    If periods Is Nothing Then Exit Function
    If periods.Count = 0 Then
        MsgBox "Kein gueltiger KV-Zeitraum gefunden.", vbExclamation, "KV-Zeitraeume loeschen"
        Exit Function
    End If
    
    promptText = "Welche KV-Zeitraeume sollen geloescht werden?" & vbCrLf & vbCrLf
    promptText = promptText & PID_BuildNumberedListFromCollection(periods)
    promptText = promptText & vbCrLf & "Mehrere Nummern mit Komma trennen (z.B. 2 oder 2,3)." & vbCrLf
    promptText = promptText & "Nummer 1 = oberster Zeitraum in der Tabelle."
    
    inputText = InputBox(Prompt:=promptText, Title:="KV-Zeitraeume loeschen", Default:="")
    inputText = Trim$(inputText)
    If inputText = "" Then Exit Function
    
    Set AskForKVPeriodsToDelete = PID_ParsePeriodIndexSelection(inputText, periods)
End Function


Private Function PID_ParsePeriodIndexSelection(ByVal inputText As String, ByVal periods As Collection) As Collection
    Dim result As Collection
    Dim parts As Variant
    Dim i As Long
    Dim token As String
    Dim selectedIndex As Long
    Dim periodName As String
    
    Set result = New Collection
    
    On Error GoTo InvalidInput
    
    inputText = Replace(inputText, ";", ",")
    parts = Split(inputText, ",")
    
    For i = LBound(parts) To UBound(parts)
        token = Trim$(CStr(parts(i)))
        If token = "" Then GoTo NextToken
        
        If Not IsNumeric(token) Then GoTo InvalidInput
        
        selectedIndex = CLng(token)
        
        If selectedIndex < 1 Or selectedIndex > periods.Count Then GoTo InvalidInput
        
        periodName = CStr(periods(selectedIndex))
        
        If Not PID_CollectionContainsText(result, periodName) Then
            result.Add periodName, PID_MakeCollectionKey(periodName)
        End If
        
NextToken:
    Next i
    
    Set PID_ParsePeriodIndexSelection = result
    Exit Function
    
InvalidInput:
    MsgBox "Ungueltige Eingabe. Bitte Nummern aus der Liste verwenden (z.B. 2 oder 2,3).", _
           vbExclamation, "KV-Zeitraeume loeschen"
End Function


Private Function PID_BuildPeriodDeleteSummary(ByVal periodsToDelete As Collection) As String
    Dim summaryText As String
    Dim i As Long
    
    If periodsToDelete Is Nothing Then Exit Function
    
    For i = 1 To periodsToDelete.Count
        summaryText = summaryText & "- " & CStr(periodsToDelete(i)) & vbCrLf
    Next i
    
    PID_BuildPeriodDeleteSummary = summaryText
End Function


Private Function PID_SortPeriodsByFirstRowDesc(ByVal wsKV As Worksheet, _
                                             ByVal periodsToDelete As Collection, _
                                             ByVal firstDataRow As Long) As Collection
    Dim result As Collection
    Dim countItems As Long
    Dim names() As String
    Dim rows() As Long
    Dim i As Long
    Dim j As Long
    Dim tmpName As String
    Dim tmpRow As Long
    
    Set result = New Collection
    
    If periodsToDelete Is Nothing Then
        Set PID_SortPeriodsByFirstRowDesc = result
        Exit Function
    End If
    
    countItems = periodsToDelete.Count
    If countItems = 0 Then
        Set PID_SortPeriodsByFirstRowDesc = result
        Exit Function
    End If
    
    ReDim names(1 To countItems)
    ReDim rows(1 To countItems)
    
    For i = 1 To countItems
        names(i) = CStr(periodsToDelete(i))
        rows(i) = FindFirstRowOfPeriod(wsKV, names(i), firstDataRow)
    Next i
    
    For i = 1 To countItems - 1
        For j = i + 1 To countItems
            If rows(j) > rows(i) Then
                tmpRow = rows(i)
                rows(i) = rows(j)
                rows(j) = tmpRow
                
                tmpName = names(i)
                names(i) = names(j)
                names(j) = tmpName
            End If
        Next j
    Next i
    
    For i = 1 To countItems
        result.Add names(i)
    Next i
    
    Set PID_SortPeriodsByFirstRowDesc = result
End Function


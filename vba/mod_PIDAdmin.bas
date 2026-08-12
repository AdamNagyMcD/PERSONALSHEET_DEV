Attribute VB_Name = "mod_PIDAdmin"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit


Public Function PID_ConfirmAdminAction(ByVal actionDescription As String, ByVal dialogTitle As String) As Boolean
    Dim answer As VbMsgBoxResult
    
    answer = MsgBox( _
        "ADMIN-MAKRO — nur f" & PID_UTxtUe() & "r Entwickler/Release." & vbCrLf & vbCrLf & _
        actionDescription & vbCrLf & vbCrLf & _
        "Fortfahren?", _
        vbExclamation + vbYesNo, _
        dialogTitle)
    
    PID_ConfirmAdminAction = (answer = vbYes)
End Function


Public Sub PID_ShowAdminMacroInfo()
    Dim msg As String
    
    msg = "Admin-/Entwickler-Makros (nicht f" & PID_UTxtUe() & "r Restaurant-User):" & vbCrLf & vbCrLf
    msg = msg & "- PID_ToggleAdminSheet (Admin-Panel ein/aus)" & vbCrLf
    msg = msg & "- ResetAndImportVBAFiles" & vbCrLf
    msg = msg & "- FullSystemRefresh / PID_FullSystemRefresh" & vbCrLf
    msg = msg & "- PID_QuickSystemCheck" & vbCrLf
    msg = msg & "- PID_RunSystemSmokeCheck" & vbCrLf
    msg = msg & "- PID_RunPerformanceBaseline (FP-010)" & vbCrLf
    msg = msg & "- PID_AdminResetHourOverrideLog (Stunden-Log leeren)" & vbCrLf
    msg = msg & "- PID_AdminKorrigierePersonalId (Personal-ID/Name korrigieren)" & vbCrLf
    msg = msg & "- PID_AdminMitarbeiterEntfernen (Mitarbeiter aus Monaten entfernen)" & vbCrLf
    msg = msg & "- PID_AdminFehlerMelden (Fehlermeldung mit Kontext erstellen)" & vbCrLf
    msg = msg & "- PID_AdminShowActionLog (letzte Aktionen anzeigen)" & vbCrLf
    msg = msg & "- PID_PruefeFormelspalten (Formeln G/H/K/L pr" & PID_UTxtUe() & "fen)" & vbCrLf
    msg = msg & "- PID_FormelspaltenReparieren (fehlende Formeln " & PID_UTxtErgaenzt() & ")" & vbCrLf
    msg = msg & "- RebuildLOHNTABELLE" & vbCrLf
    msg = msg & "- UnprotectEverything" & vbCrLf & vbCrLf
    msg = msg & "Siehe docs/RELEASE.md"
    
    MsgBox msg, vbInformation, "Admin Makros"
End Sub


Public Sub PID_AdminResetHourOverrideLog()
    If Not PID_ConfirmAdminAction( _
        "Das versteckte Blatt PID_HOUR_OVERRIDES wird geleert." & vbCrLf & _
        "Nur n" & PID_UTxtOe() & "tig nach altem CopyData-Log oder FP-028-Tests.", _
        "Stunden-Log leeren") Then
        Exit Sub
    End If
    
    PID_ResetHourOverrideLog
    
    MsgBox "PID_HOUR_OVERRIDES geleert.", vbInformation, "Admin"
End Sub


Public Sub FullSystemRefresh()
    PID_FullSystemRefresh
End Sub


Public Sub PID_FullSystemRefresh()
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    Dim checkedSheets As Long
    Dim repairedCells As Long
    Dim repairedSheets As Long
    Dim fixedIndexSheets As Long
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Personalsheet wird aktualisiert..."
    
    PID_SetupSheetProtectionForMacros
    
    RefreshAllMonthKVStundenDropdowns
    PID_RestoreMonatslohnFormulasSilent
    PID_RestoreAktuelleStundenFormulasSilent
    PID_RestoreUrlaubGeldFormulasSilent
    PID_RestoreLetztesGehaltFormulasSilent
    PID_RestoreKVCodeDropdownValidationSilent
    ClearAllKVLohnDirty
    
    ' TR-10: Sicherheitsnetz nach den vier Spalten-Wiederherstellungen. Diese
    ' ueberspringen ein Monatsblatt still, wenn der Monatsindex in A1 fehlt oder
    ' nicht passt - hier werden A1 und einzelne fehlende Formeln zeilenweise ergaenzt.
    checkedSheets = PID_RepairFormulaColumnsSilent(repairedCells, repairedSheets, fixedIndexSheets)
    
    RefreshFluktuationAll
    
    PID_RestoreFinanzSummaryOnUbersicht
    
    PID_RecalculateAllMonthMergedFormulas
    
    PID_FormatAllMoneyColumns
    
    PID_ApplyCopyrightToAllSheets
    
    PID_EnsureAdminSheet
    PID_HideAdminSheet False
    
    PID_EnableCalculationForAllSheets
    
    On Error Resume Next
    Application.CalculateFull
    On Error GoTo CleanFail
    
    MsgBox "Personalsheet wurde " & PID_UTxtVollstaendig() & " aktualisiert." & vbCrLf & vbCrLf & _
           "Formelspalten G/H/K/L " & PID_UTxtGeprueft() & ": " & checkedSheets & " " & PID_UTxtMonatsblaetter() & vbCrLf & _
           "Fehlende Formeln " & PID_UTxtErgaenzt() & ": " & repairedCells & " Zellen" & vbCrLf & _
           "Monatsindex A1 korrigiert: " & fixedIndexSheets, _
           vbInformation, "System Refresh"

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler bei PID_FullSystemRefresh:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Personalsheet"
End Sub


Public Sub PID_QuickSystemCheck()
    Dim msg As String
    
    msg = "Personalsheet Systemcheck" & vbCrLf & vbCrLf
    
    msg = msg & "Workbook: " & ThisWorkbook.Name & vbCrLf
    msg = msg & "Excel Version: " & Application.Version & vbCrLf
    msg = msg & "Operating System: " & Application.OperatingSystem & vbCrLf & vbCrLf
    
    msg = msg & "Events aktiv: " & CStr(Application.EnableEvents) & vbCrLf
    msg = msg & "ScreenUpdating aktiv: " & CStr(Application.ScreenUpdating) & vbCrLf
    msg = msg & "Calculation: " & PID_GetCalculationModeText() & vbCrLf & vbCrLf
    
    msg = msg & "Pflichtbl" & PID_UTxtAe() & "tter:" & vbCrLf
    msg = msg & "- EINSTELLUNG: " & PID_YesNoText(PID_WorksheetExists(PID_EINSTELLUNG_SHEET)) & vbCrLf
    msg = msg & "- LOHNTABELLE: " & PID_YesNoText(PID_WorksheetExists(PID_LOHNTABELLE_SHEET)) & vbCrLf
    msg = msg & "- FLUKTUATION: " & PID_YesNoText(PID_WorksheetExists(PID_FLUKTUATION_SHEET)) & vbCrLf
    msg = msg & "- FLUKTUATION_DATEN: " & PID_YesNoText(PID_WorksheetExists("FLUKTUATION_DATEN")) & vbCrLf
    msg = msg & "- KV_DROPDOWN_HELPER: " & PID_YesNoText(PID_WorksheetExists("KV_DROPDOWN_HELPER")) & vbCrLf & vbCrLf
    
    msg = msg & PID_UTxtMonatsblaetter() & " gefunden: " & CStr(PID_CountMonthSheets()) & " / 12" & vbCrLf
    
    MsgBox msg, vbInformation, "Personalsheet Systemcheck"
End Sub

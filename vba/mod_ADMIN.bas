Attribute VB_Name = "mod_ADMIN"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Sammelstelle fuer alle Entwickler- und Wartungsmakros.
'
' Warum: Alt+F8 listet jede oeffentliche Prozedur ohne Argumente auf - das sind ueber
' 130 Eintraege, in denen die paar wirklich nutzbaren Werkzeuge untergehen. Die
' ADMIN_-Namen stehen dank Ziffernblock direkt am Anfang der Liste und sind nach
' Themen gruppiert:
'
'   ADMIN_01-05   Setup und VBA
'   ADMIN_10-18   Test und Diagnose
'   ADMIN_20-28   Reparatur
'   ADMIN_30-38   Format und LOHNTABELLE
'   ADMIN_40-41   Daten loeschen (Vorsicht)
'
' Die Eintraege sind bewusst duenne Weiterleitungen: die eigentliche Logik bleibt in
' ihrem Fachmodul, alle bestehenden Aufrufe (Schaltflaechen, Ereignisse, Doku)
' funktionieren unveraendert weiter. Nur die beiden Werkzeuge, die es ausschliesslich
' fuer den Entwickler gibt - VBA-Export und das Aufraeumen kaputter Namen - stehen
' hier direkt.
'
' Alles was mit PID_ beginnt, ist interne Technik und wird nicht von Hand gestartet.
' Die Makros fuer die Restaurant-Manager heissen weiterhin CopyData, DataClear,
' AlleDatenLoeschen, PersonalIdKorrigieren, MitarbeiterEntfernen, FehlerMelden und
' RefreshFluktuationNow.

Private Const PID_ADMIN_EXPORT_FOLDER As String = "vba_export"


'==============================================================================
' 00-05  Setup und VBA
'==============================================================================

Public Sub ADMIN_00_Hilfe()
    Dim msg As String
    
    msg = "ADMIN-Makros " & ChrW(8212) & " " & PID_UTxtUebersicht() & vbCrLf & vbCrLf
    msg = msg & "Alt+F8 zeigt alle ADMIN_-Makros zusammen am Anfang der Liste." & vbCrLf & vbCrLf
    msg = msg & "ADMIN_01-05   Setup und VBA (Import, Export, Admin-Panel)" & vbCrLf
    msg = msg & "ADMIN_10-18   Test und Diagnose (Smoke, Schnellcheck," & vbCrLf
    msg = msg & "              Performance, Formelspalten, Protokolle)" & vbCrLf
    msg = msg & "ADMIN_20-28   Reparatur (Full Refresh, Formeln, Dropdowns, Schutz)" & vbCrLf
    msg = msg & "ADMIN_30-38   Format und LOHNTABELLE" & vbCrLf
    msg = msg & "ADMIN_40-41   Daten " & PID_UTxtLoeschen() & " " & ChrW(8212) & " VORSICHT" & vbCrLf & vbCrLf
    msg = msg & "Makros " & PID_UTxtFuer() & " die Restaurant-Manager:" & vbCrLf
    msg = msg & "CopyData, DataClear, AlleDatenLoeschen, PersonalIdKorrigieren," & vbCrLf
    msg = msg & "MitarbeiterEntfernen, FehlerMelden, RefreshFluktuationNow" & vbCrLf & vbCrLf
    msg = msg & "Alles mit PID_ ist interne Technik und wird nicht direkt gestartet."
    
    MsgBox msg, vbInformation, "ADMIN"
End Sub


Public Sub ADMIN_01_VBA_Import()
    ResetAndImportVBAFiles
End Sub


' Schreibt alle Module in den Ordner "vba_export" neben der Arbeitsmappe.
' Bewusst NICHT in den Ordner "vba": der ist die Git-Quelle und darf nicht von einem
' aelteren Stand aus der Mappe ueberschrieben werden. Der Vergleich der beiden Ordner
' zeigt, was in der Mappe vom Repository abweicht.
Public Sub ADMIN_02_VBA_Export()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim targetFolder As String
    Dim fileName As String
    Dim extension As String
    Dim exported As Long
    Dim skipped As Long
    
    On Error GoTo VBProjectBlocked
    
    If ThisWorkbook.Path = "" Then
        MsgBox "Die Arbeitsmappe ist noch nicht gespeichert." & vbCrLf & vbCrLf & _
               "Bitte zuerst speichern, damit der Export-Ordner angelegt werden kann.", _
               vbExclamation, "VBA Export"
        Exit Sub
    End If
    
    Set vbProj = ThisWorkbook.VBProject
    
    On Error GoTo ExportError
    
    targetFolder = ThisWorkbook.Path & Application.PathSeparator & PID_ADMIN_EXPORT_FOLDER
    
    If Not PID_AdminFolderExists(targetFolder) Then
        MkDir targetFolder
    End If
    
    For Each vbComp In vbProj.VBComponents
        extension = PID_AdminExportExtension(CLng(vbComp.Type))
        
        If extension = "" Then
            skipped = skipped + 1
        Else
            fileName = targetFolder & Application.PathSeparator & vbComp.Name & extension
            vbComp.Export fileName
            exported = exported + 1
        End If
    Next vbComp
    
    MsgBox exported & " Module exportiert, " & skipped & " " & PID_UTxtUebersprungen() & "." & vbCrLf & vbCrLf & _
           "Ordner: " & PID_ADMIN_EXPORT_FOLDER & vbCrLf & vbCrLf & _
           "Der Ordner ""vba"" bleibt unangetastet " & ChrW(8212) & " er ist die Quelle " & PID_UTxtFuer() & " den Import.", _
           vbInformation, "VBA Export"
    Exit Sub

VBProjectBlocked:
    MsgBox "Zugriff auf das VBA-Projekt ist blockiert (Fehler " & Err.Number & ")." & vbCrLf & vbCrLf & _
           "Datei > Optionen > Trust Center > Trust Center-Einstellungen > " & _
           "Makroeinstellungen > ""Zugriff auf das VBA-Projektobjektmodell vertrauen""", _
           vbExclamation, "VBA Export"
    Exit Sub

ExportError:
    MsgBox "Fehler beim VBA-Export:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "VBA Export"
End Sub


Public Sub ADMIN_03_VBA_Reparatur_Nach_Import()
    PID_RepairWorkbookAfterVBAImport
    
    MsgBox "Formeln und Dropdowns wurden nach dem Import neu gesetzt." & vbCrLf & vbCrLf & _
           "Danach: Kompilieren, speichern und ADMIN_10_Test_Smoke_Check " & PID_UTxtAusfuehren() & ".", _
           vbInformation, "VBA Import"
End Sub


Public Sub ADMIN_04_Admin_Panel()
    PID_ToggleAdminSheet
End Sub


Public Sub ADMIN_05_Makro_Uebersicht()
    ADMIN_00_Hilfe
End Sub


'==============================================================================
' 10-18  Test und Diagnose
'==============================================================================

Public Sub ADMIN_10_Test_Smoke_Check()
    PID_RunSystemSmokeCheck
End Sub


Public Sub ADMIN_11_Test_Schnellcheck()
    PID_QuickSystemCheck
End Sub


Public Sub ADMIN_12_Test_Performance()
    PID_RunPerformanceBaseline
End Sub


Public Sub ADMIN_13_Test_Formelspalten()
    PID_PruefeFormelspalten
End Sub


Public Sub ADMIN_14_Test_Stunden_Log()
    PID_ShowHourOverrideLog
End Sub


Public Sub ADMIN_15_Test_Aktionsprotokoll()
    PID_AdminShowActionLog
End Sub


Public Sub ADMIN_16_Test_Ergebnisblatt()
    PID_AdminOpenSmokeSheet
End Sub


Public Sub ADMIN_17_Tech_Blaetter_Zeigen()
    PID_ShowTechnicalSheets
End Sub


Public Sub ADMIN_18_Tech_Blaetter_Verbergen()
    PID_HideTechnicalSheets
End Sub


'==============================================================================
' 20-28  Reparatur
'==============================================================================

Public Sub ADMIN_20_Reparatur_Full_Refresh()
    PID_FullSystemRefresh
End Sub


Public Sub ADMIN_21_Reparatur_Formelspalten()
    PID_FormelspaltenReparieren
End Sub


Public Sub ADMIN_22_Reparatur_Stunden_Dropdown()
    RestoreKVStundenDropdownValidation
End Sub


Public Sub ADMIN_23_Reparatur_KV_Dropdown()
    PID_RestoreKVCodeDropdownValidation
End Sub


Public Sub ADMIN_24_Reparatur_Austrittsdatum()
    PID_RestoreAustrittsdatumValidation
End Sub


' Entfernt Namen, die ins Leere zeigen (#REF!). In der Testdatei sind das die
' Altlasten BG1_Hours, BG1_Wages ... aus einer frueheren Dropdown-Loesung: sie werden
' nirgends mehr verwendet, tauchen aber im Namensmanager auf und wandern bei jedem
' Blattkopieren mit. Namen, die auf einen gueltigen Bereich zeigen, und interne
' Excel-Namen (beginnen mit _) bleiben unangetastet.
Public Sub ADMIN_25_Namen_Aufraeumen()
    Dim brokenNames As String
    Dim brokenCount As Long
    
    brokenCount = PID_AdminCollectBrokenNames(brokenNames)
    
    If brokenCount = 0 Then
        MsgBox "Es gibt keine Namen mit #REF!." & vbCrLf & vbCrLf & _
               "Namen gesamt: " & ThisWorkbook.Names.count, _
               vbInformation, "Namen aufr" & PID_UTxtAe() & "umen"
        Exit Sub
    End If
    
    If Not PID_ConfirmAdminAction( _
        brokenCount & " Namen zeigen auf #REF! und werden " & PID_UTxtGeloescht() & ":" & vbCrLf & _
        brokenNames, _
        "Namen aufr" & PID_UTxtAe() & "umen") Then
        Exit Sub
    End If
    
    PID_RemoveLegacyKVDDNamedRanges
    brokenCount = PID_AdminDeleteBrokenNames()
    
    PID_TrackAction "Namen aufraeumen", brokenCount & " Namen " & PID_UTxtGeloescht()
    
    MsgBox brokenCount & " Namen wurden " & PID_UTxtGeloescht() & "." & vbCrLf & vbCrLf & _
           "Namen gesamt: " & ThisWorkbook.Names.count, _
           vbInformation, "Namen aufr" & PID_UTxtAe() & "umen"
End Sub


Public Sub ADMIN_26_Schutz_AN()
    PID_AdminRunProtectAll
End Sub


Public Sub ADMIN_27_Schutz_AUS()
    UnprotectEverything
End Sub


Public Sub ADMIN_28_UEBERSICHT_Schutz()
    PID_FixUbersichtPlanInputsEditable
End Sub


'==============================================================================
' 30-38  Format und LOHNTABELLE
'==============================================================================

Public Sub ADMIN_30_Format_Alle_Monate()
    PID_AdminRunFormatMonths
End Sub


Public Sub ADMIN_31_Format_Januar()
    FormatJanuarMonthSheet
End Sub


Public Sub ADMIN_32_Format_EINSTELLUNG()
    PID_FormatEinstellungSheet
End Sub


Public Sub ADMIN_33_Format_UEBERSICHT_Finanz()
    PID_FormatFinanzUebersicht
End Sub


Public Sub ADMIN_34_Format_UEBERSICHT_Durchrechnung()
    PID_FormatDurchrechnungUebersicht
End Sub


Public Sub ADMIN_35_UEBERSICHT_Durchrechnung_Neu()
    PID_RefreshDurchrechnungUebersicht
End Sub


Public Sub ADMIN_36_LOHNTABELLE_Neu_Aufbauen()
    PID_AdminRunRebuildLohn
End Sub


Public Sub ADMIN_37_LOHNTABELLE_Kopf_Reparieren()
    FixLOHNTABELLE_HeaderText
End Sub


Public Sub ADMIN_38_LOHNTABELLE_Aufraeumen()
    CleanupLOHNTABELLETrailingArea
End Sub


'==============================================================================
' 40-41  Daten loeschen (Vorsicht)
'==============================================================================

Public Sub ADMIN_40_Daten_Stunden_Log_Leeren()
    PID_AdminResetHourOverrideLog
End Sub


Public Sub ADMIN_41_Daten_Alles_Loeschen()
    PID_ClearAllWorkbookData
End Sub


'==============================================================================
' Interne Helfer
'==============================================================================

Private Function PID_AdminFolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    PID_AdminFolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    Err.Clear
End Function


' 1 = Standardmodul, 2 = Klassenmodul, 3 = UserForm, 100 = Dokumentmodul
Private Function PID_AdminExportExtension(ByVal componentType As Long) As String
    Select Case componentType
        Case 1
            PID_AdminExportExtension = ".bas"
        Case 2, 100
            PID_AdminExportExtension = ".cls"
        Case 3
            PID_AdminExportExtension = ".frm"
        Case Else
            PID_AdminExportExtension = ""
    End Select
End Function


Private Function PID_AdminNameIsBroken(ByVal nameIndex As Long) As Boolean
    Dim nameText As String
    Dim refersTo As String
    
    On Error GoTo SafeExit
    
    nameText = CStr(ThisWorkbook.Names(nameIndex).Name)
    
    ' Interne Excel-Namen (_xlfn, _xlpm, Print_Area ...) nicht anfassen.
    If Left$(nameText, 1) = "_" Then Exit Function
    If ThisWorkbook.Names(nameIndex).BuiltIn Then Exit Function
    
    refersTo = CStr(ThisWorkbook.Names(nameIndex).refersTo)
    PID_AdminNameIsBroken = (InStr(1, refersTo, "#REF!", vbTextCompare) > 0)

SafeExit:
End Function


Private Function PID_AdminCollectBrokenNames(ByRef nameList As String) As Long
    Dim i As Long
    Dim brokenCount As Long
    
    On Error GoTo SafeExit
    
    nameList = ""
    
    For i = 1 To ThisWorkbook.Names.count
        If PID_AdminNameIsBroken(i) Then
            brokenCount = brokenCount + 1
            If brokenCount <= 15 Then
                nameList = nameList & "- " & CStr(ThisWorkbook.Names(i).Name) & vbCrLf
            End If
        End If
    Next i
    
    If brokenCount > 15 Then
        nameList = nameList & "- ... (" & (brokenCount - 15) & " weitere)" & vbCrLf
    End If

SafeExit:
    PID_AdminCollectBrokenNames = brokenCount
End Function


Private Function PID_AdminDeleteBrokenNames() As Long
    Dim i As Long
    Dim deletedCount As Long
    
    On Error Resume Next
    
    ' Rueckwaerts, damit das Loeschen die Indizes nicht verschiebt.
    For i = ThisWorkbook.Names.count To 1 Step -1
        If PID_AdminNameIsBroken(i) Then
            ThisWorkbook.Names(i).Delete
            If Err.Number = 0 Then deletedCount = deletedCount + 1
            Err.Clear
        End If
    Next i
    
    PID_AdminDeleteBrokenNames = deletedCount
End Function

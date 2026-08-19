Attribute VB_Name = "mod_PIDAdminSheet"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Public Const PID_ADMIN_SHEET_NAME As String = "_ADMIN"

Private Const PID_ADMIN_BTN_PREFIX As String = "btn_Admin_"
Private Const PID_ADMIN_BTN_WIDTH As Double = 148#
Private Const PID_ADMIN_BTN_HEIGHT As Double = 30#
Private Const PID_ADMIN_BTN_GAP_H As Double = 14#
Private Const PID_ADMIN_BTN_GAP_V As Double = 10#
Private Const PID_ADMIN_BTN_COLS As Long = 2
Private Const PID_ADMIN_BTN_COUNT As Long = 22
' Weit rechts und ausserhalb jedes Layoutbereichs: hier steht die Signatur der zuletzt
' gebauten Button-Liste, damit PID_EnsureAdminSheet den Neuaufbau ueberspringen kann.
Private Const PID_ADMIN_SIGNATURE_CELL As String = "BZ1"


' Laeuft bei jedem Oeffnen (Workbook_Open) und vor jedem Anzeigen des Panels.
' Frueher wurde dabei jedes Mal das komplette Blatt neu beschriftet und alle 22 Buttons
' geloescht und neu angelegt - 22 Shape-Loeschungen plus 22 Shape-Erzeugungen mit je
' Text, OnAction und Stil, obwohl das Blatt danach sofort wieder auf xlSheetVeryHidden
' geht und ein normaler Benutzer es nie sieht. Der Neuaufbau passiert jetzt nur noch,
' wenn er wirklich noetig ist: bei fehlenden Buttons oder wenn sich die Button-Liste im
' Code geaendert hat (Signatur). Damit bleibt die Selbstheilung erhalten.
Public Sub PID_EnsureAdminSheet()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Set ws = PID_GetAdminWorksheet()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PID_ADMIN_SHEET_NAME
    Else
        PID_AdminMoveSheetToFront ws
        
        If PID_AdminSheetIsUpToDate(ws) Then
            PID_AdminRefreshHeaderValues ws
            PID_HideAdminSheet False
            Exit Sub
        End If
    End If
    
    PID_BuildAdminSheetLayout ws
    PID_EnsureAdminSheetButtons ws
    PID_AdminWriteSpecSignature ws
    PID_HideAdminSheet False

SafeExit:
End Sub


' Nur dann "aktuell", wenn alle Buttons vorhanden sind UND die gespeicherte Signatur zur
' aktuellen Button-Liste im Code passt. Nach einem VBA-Import mit geaenderten Beschriftungen
' oder Makronamen stimmt die Signatur nicht mehr und das Blatt wird neu gebaut.
Private Function PID_AdminSheetIsUpToDate(ByVal ws As Worksheet) As Boolean
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If PID_AdminToolbarShapeCount(ws) <> PID_ADMIN_BTN_COUNT Then Exit Function
    
    PID_AdminSheetIsUpToDate = _
        (StrComp(CStr(ws.Range(PID_ADMIN_SIGNATURE_CELL).Value2), _
                 PID_AdminSheetSpecSignature(), vbBinaryCompare) = 0)
SafeExit:
End Function


Private Function PID_AdminToolbarShapeCount(ByVal ws As Worksheet) As Long
    Dim shp As Shape
    Dim found As Long
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(PID_ADMIN_BTN_PREFIX)) = PID_ADMIN_BTN_PREFIX Then
            found = found + 1
        End If
    Next shp
    
SafeExit:
    PID_AdminToolbarShapeCount = found
End Function


' Reine VBA-Stringarbeit, kein Zellzugriff je Button.
Private Function PID_AdminSheetSpecSignature() As String
    Dim i As Long
    Dim parts As String
    Dim btnKey As String
    Dim btnLabel As String
    Dim btnMacro As String
    Dim btnStyle As Long
    
    For i = 0 To PID_ADMIN_BTN_COUNT - 1
        PID_AdminGetButtonSpec i, btnKey, btnLabel, btnMacro, btnStyle
        parts = parts & btnKey & Chr$(31) & btnLabel & Chr$(31) & btnMacro & _
                Chr$(31) & CStr(btnStyle) & Chr$(30)
    Next i
    
    PID_AdminSheetSpecSignature = "PID_ADMIN_V1/" & CStr(PID_ADMIN_BTN_COUNT) & "/" & parts
End Function


' Jahr und Excel-Version stehen als Werte auf dem Blatt und muessen auch stimmen, wenn der
' Neuaufbau uebersprungen wird. Nur bei echter Abweichung schreiben, damit das Oeffnen die
' Mappe nicht ohne Not als geaendert markiert.
Private Sub PID_AdminRefreshHeaderValues(ByVal ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    
    If StrComp(CStr(ws.Range("C3").Value2), CStr(PID_GetWorkbookYear()), vbTextCompare) <> 0 Then
        ws.Range("C3").Value = PID_GetWorkbookYear()
    End If
    
    ' Val statt Textvergleich: Excel legt "16.0" als Zahl 16 ab, ein Textvergleich haette
    ' die Zelle bei jedem Oeffnen erneut geschrieben und die Mappe als geaendert markiert.
    If Val(CStr(ws.Range("E3").Value2)) <> Val(Application.Version) Then
        ws.Range("E3").Value = Application.Version
    End If
    
    Err.Clear
End Sub


Private Sub PID_AdminWriteSpecSignature(ByVal ws As Worksheet)
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    
    ws.Range(PID_ADMIN_SIGNATURE_CELL).Value2 = PID_AdminSheetSpecSignature()
    Err.Clear
End Sub


Public Sub PID_ShowAdminSheet()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    PID_EnsureAdminSheet
    Set ws = PID_GetAdminWorksheet()
    If ws Is Nothing Then Exit Sub
    
    ws.Visible = xlSheetVisible
    PID_AdminMoveSheetToFront ws
    ws.Activate
    ws.Range("A1").Select

SafeExit:
End Sub


Public Sub PID_HideAdminSheet(Optional ByVal activateUserSheet As Boolean = True)
    Dim ws As Worksheet
    Dim fallbackWs As Worksheet
    
    On Error GoTo SafeExit
    
    Set ws = PID_GetAdminWorksheet()
    If ws Is Nothing Then Exit Sub
    
    If ws.Visible = xlSheetVisible And activateUserSheet Then
        Set fallbackWs = PID_GetAdminFallbackWorksheet()
        If Not fallbackWs Is Nothing Then
            fallbackWs.Activate
            fallbackWs.Range("A1").Select
        End If
    End If
    
    ws.Visible = xlSheetVeryHidden

SafeExit:
End Sub


Public Sub PID_ToggleAdminSheet()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    PID_EnsureAdminSheet
    Set ws = PID_GetAdminWorksheet()
    If ws Is Nothing Then Exit Sub
    
    If ws.Visible = xlSheetVisible Then
        PID_HideAdminSheet True
    Else
        PID_ShowAdminSheet
    End If

SafeExit:
End Sub


Public Sub PID_AdminRunSmokeCheck()
    PID_RunSystemSmokeCheck
End Sub


Public Sub PID_AdminRunFullRefresh()
    PID_FullSystemRefresh
End Sub


Public Sub PID_AdminRunQuickCheck()
    PID_QuickSystemCheck
End Sub


Public Sub PID_AdminRunFormatMonths()
    If Not PID_ConfirmAdminAction( _
        "FormatAllMonthSheets auf allen 12 " & PID_UTxtMonatsblaettern() & " " & PID_UTxtAusfuehren() & "?", _
        "Monatsformat") Then
        Exit Sub
    End If
    
    FormatAllMonthSheets
End Sub


Public Sub PID_AdminRunPerfBaseline()
    If Not PID_ConfirmAdminAction( _
        "Performance-Baseline messen und in PERFORMANCE_BASELINE schreiben?", _
        "Performance Baseline") Then
        Exit Sub
    End If
    
    PID_RunPerformanceBaseline
End Sub


Public Sub PID_AdminRunProtectAll()
    If Not PID_ConfirmAdminAction( _
        "Alle " & PID_UTxtBlaetter() & " werden " & PID_UTxtGeschuetzt() & " (Monatsblatt-Policy, UEBERSICHT, LOHNTABELLE, ...).", _
        "Schutz aktivieren") Then
        Exit Sub
    End If
    
    PID_SetupSheetProtectionForMacros
    
    MsgBox "Schutz auf allen " & PID_UTxtBlaetter() & "n aktiv.", vbInformation, "Admin"
End Sub


Public Sub PID_AdminRunUnprotectAll()
    UnprotectEverything
End Sub


Public Sub PID_AdminShowTechSheets()
    PID_ShowTechnicalSheets
End Sub


Public Sub PID_AdminHideTechSheets()
    PID_HideTechnicalSheets
End Sub


Public Sub PID_AdminRunRebuildLohn()
    RebuildLOHNTABELLE
End Sub


Public Sub PID_AdminOpenSmokeSheet()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("SYSTEM_CHECK")
    On Error GoTo SafeExit
    
    If ws Is Nothing Then
        MsgBox "Noch kein Smoke-Check gelaufen. Bitte zuerst ""Smoke Check"" starten.", _
               vbExclamation, "Admin"
        Exit Sub
    End If
    
    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range("A1").Select

SafeExit:
End Sub


Public Sub PID_AdminHidePanel()
    PID_HideAdminSheet True
End Sub


Private Function PID_GetAdminWorksheet() As Worksheet
    On Error Resume Next
    Set PID_GetAdminWorksheet = ThisWorkbook.Worksheets(PID_ADMIN_SHEET_NAME)
End Function


Private Function PID_GetAdminFallbackWorksheet() As Worksheet
    Dim ws As Worksheet
    
    Set PID_GetAdminFallbackWorksheet = PID_GetUbersichtWorksheet()
    If Not PID_GetAdminFallbackWorksheet Is Nothing Then Exit Function
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If StrComp(ws.Name, PID_ADMIN_SHEET_NAME, vbTextCompare) <> 0 Then
                Set PID_GetAdminFallbackWorksheet = ws
                Exit Function
            End If
        End If
    Next ws
End Function


Private Sub PID_BuildAdminSheetLayout(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    ws.Unprotect
    On Error GoTo 0
    
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
    
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    
    ws.Range("B1:E1").Merge
    ws.Range("B1").Value = "Personalsheet — Admin"
    ws.Range("B1").Font.Size = 16
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Color = RGB(31, 56, 100)
    
    ws.Range("B2:E2").Merge
    ws.Range("B2").Value = "Alt+F8 → PID_ToggleAdminSheet  |  Nur Entwickler / Release"
    ws.Range("B2").Font.Size = 9
    ws.Range("B2").Font.Color = RGB(90, 108, 125)
    
    ws.Range("B3").Value = "Jahr:"
    ws.Range("C3").Value = PID_GetWorkbookYear()
    ws.Range("D3").Value = "Excel:"
    ws.Range("E3").Value = Application.Version
    
    ws.Range("B3:E3").Font.Size = 9
    ws.Range("B3:E3").Font.Color = RGB(70, 85, 105)
    
    ws.Tab.Color = RGB(192, 0, 0)
End Sub


Private Sub PID_AdminMoveSheetToFront(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If ws.Index = 1 Then Exit Sub
    
    On Error Resume Next
    ws.Move Before:=ThisWorkbook.Worksheets(1)
    On Error GoTo 0
End Sub


Private Sub PID_EnsureAdminSheetButtons(ByVal ws As Worksheet)
    Dim i As Long
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim anchorLeft As Double
    Dim anchorTop As Double
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btn As Shape
    Dim btnKey As String
    Dim btnLabel As String
    Dim btnMacro As String
    Dim btnStyle As Long
    
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    ws.Unprotect
    On Error GoTo 0
    
    PID_DeleteAdminToolbarShapes ws
    
    anchorLeft = ws.Range("B5").Left
    anchorTop = ws.Range("B5").Top
    
    For i = 0 To PID_ADMIN_BTN_COUNT - 1
        PID_AdminGetButtonSpec i, btnKey, btnLabel, btnMacro, btnStyle
        
        colIndex = i Mod PID_ADMIN_BTN_COLS
        rowIndex = i \ PID_ADMIN_BTN_COLS
        
        btnLeft = anchorLeft + (colIndex * (PID_ADMIN_BTN_WIDTH + PID_ADMIN_BTN_GAP_H))
        btnTop = anchorTop + (rowIndex * (PID_ADMIN_BTN_HEIGHT + PID_ADMIN_BTN_GAP_V))
        
        Set btn = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                     Left:=btnLeft, _
                                     Top:=btnTop, _
                                     Width:=PID_ADMIN_BTN_WIDTH, _
                                     Height:=PID_ADMIN_BTN_HEIGHT)
        
        btn.Name = PID_ADMIN_BTN_PREFIX & btnKey
        btn.TextFrame.Characters.Text = btnLabel
        btn.OnAction = btnMacro
        
        PID_AdminApplyButtonStyle btn, btnStyle
    Next i
End Sub


Private Sub PID_AdminGetButtonSpec(ByVal index As Long, _
                                   ByRef btnKey As String, _
                                   ByRef btnLabel As String, _
                                   ByRef btnMacro As String, _
                                   ByRef btnStyle As Long)
    Select Case index
        Case 0
            btnKey = "Smoke": btnLabel = "Smoke Check": btnMacro = "PID_AdminRunSmokeCheck": btnStyle = 0
        Case 1
            btnKey = "Refresh": btnLabel = "Full Refresh": btnMacro = "PID_AdminRunFullRefresh": btnStyle = 0
        Case 2
            btnKey = "Quick": btnLabel = "Quick Check": btnMacro = "PID_AdminRunQuickCheck": btnStyle = 1
        Case 3
            btnKey = "Format": btnLabel = "Format Monate": btnMacro = "PID_AdminRunFormatMonths": btnStyle = 1
        Case 4
            btnKey = "Import": btnLabel = "VBA Import": btnMacro = "ResetAndImportVBAFiles": btnStyle = 2
        Case 5
            btnKey = "Perf": btnLabel = "Perf. Baseline": btnMacro = "PID_AdminRunPerfBaseline": btnStyle = 2
        Case 6
            btnKey = "Protect": btnLabel = "Schutz AN": btnMacro = "PID_AdminRunProtectAll": btnStyle = 3
        Case 7
            btnKey = "Unprotect": btnLabel = "Schutz AUS": btnMacro = "PID_AdminRunUnprotectAll": btnStyle = 2
        Case 8
            btnKey = "ShowTech": btnLabel = "Tech-Bl. zeigen": btnMacro = "PID_AdminShowTechSheets": btnStyle = 1
        Case 9
            btnKey = "HideTech": btnLabel = "Tech-Bl. verbergen": btnMacro = "PID_AdminHideTechSheets": btnStyle = 1
        Case 10
            btnKey = "Lohn": btnLabel = "LOHNTABELLE": btnMacro = "PID_AdminRunRebuildLohn": btnStyle = 2
        Case 11
            btnKey = "ResetHours": btnLabel = "Stunden-Log": btnMacro = "PID_AdminResetHourOverrideLog": btnStyle = 2
        Case 12
            btnKey = "SmokeSheet": btnLabel = "SYSTEM_CHECK": btnMacro = "PID_AdminOpenSmokeSheet": btnStyle = 3
        Case 13
            btnKey = "FixPID": btnLabel = "Personal-ID fix": btnMacro = "PID_AdminKorrigierePersonalId": btnStyle = 2
        Case 14
            btnKey = "DelMA": btnLabel = "MA entfernen": btnMacro = "PID_AdminMitarbeiterEntfernen": btnStyle = 2
        Case 15
            btnKey = "Feedback": btnLabel = "Fehler melden": btnMacro = "PID_AdminFehlerMelden": btnStyle = 1
        Case 16
            btnKey = "ActionLog": btnLabel = "Aktionsprotokoll": btnMacro = "PID_AdminShowActionLog": btnStyle = 3
        Case 17
            btnKey = "CheckFormulas": btnLabel = "Formeln pr" & PID_UTxtUe() & "fen": btnMacro = "PID_PruefeFormelspalten": btnStyle = 1
        Case 18
            btnKey = "FixFormulas": btnLabel = "Formeln reparieren": btnMacro = "PID_FormelspaltenReparieren": btnStyle = 2
        Case 19
            btnKey = "ClearAll": btnLabel = "Alle Daten l" & PID_UTxtOe() & "schen": btnMacro = "PID_ClearAllWorkbookData": btnStyle = 2
        Case 20
            btnKey = "AdminHelp": btnLabel = "ADMIN-Makros": btnMacro = "ADMIN_00_Hilfe": btnStyle = 1
        Case Else
            btnKey = "Hide": btnLabel = "Admin verbergen": btnMacro = "PID_AdminHidePanel": btnStyle = 3
    End Select
End Sub


Private Sub PID_AdminApplyButtonStyle(ByVal btn As Shape, ByVal styleId As Long)
    Select Case styleId
        Case 0
            PID_StyleApplyToolbarButton btn, PID_StyleColorNavy(), PID_StyleColorBtnPrimaryLine(), RGB(255, 255, 255)
        Case 1
            PID_StyleApplyToolbarButton btn, PID_StyleColorHeaderBg(), PID_StyleColorNavy(), PID_StyleColorNavy()
        Case 2
            PID_StyleApplyToolbarButton btn, PID_StyleColorAccent(), PID_StyleColorNavy(), PID_StyleColorNavy()
        Case Else
            PID_StyleApplyToolbarButton btn, PID_StyleColorZebra(), PID_StyleColorNavy(), PID_StyleColorNavy()
    End Select
End Sub


Private Sub PID_DeleteAdminToolbarShapes(ByVal ws As Worksheet)
    Dim shapeIndex As Long
    Dim shp As Shape
    
    If ws Is Nothing Then Exit Sub
    
    For shapeIndex = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(shapeIndex)
        If Left$(shp.Name, Len(PID_ADMIN_BTN_PREFIX)) = PID_ADMIN_BTN_PREFIX Then
            shp.Delete
        End If
    Next shapeIndex
End Sub

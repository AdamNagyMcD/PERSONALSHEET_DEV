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


Public Sub PID_EnsureAdminSheet()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    Set ws = PID_GetAdminWorksheet()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PID_ADMIN_SHEET_NAME
    Else
        ws.Move Before:=ThisWorkbook.Worksheets(1)
    End If
    
    PID_BuildAdminSheetLayout ws
    PID_EnsureAdminSheetButtons ws
    PID_HideAdminSheet False

SafeExit:
End Sub


Public Sub PID_ShowAdminSheet()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    PID_EnsureAdminSheet
    Set ws = PID_GetAdminWorksheet()
    If ws Is Nothing Then Exit Sub
    
    ws.Visible = xlSheetVisible
    ws.Move Before:=ThisWorkbook.Worksheets(1)
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
        "FormatAllMonthSheets auf allen 12 Monatsblaettern ausfuehren?", _
        "Monatsformat") Then
        Exit Sub
    End If
    
    FormatAllMonthSheets
End Sub


Public Sub PID_AdminRunVbaImport()
    ResetAndImportVBAFiles
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
        "Alle Blaetter werden geschuetzt (Monatsblatt-Policy, UEBERSICHT, LOHNTABELLE, ...).", _
        "Schutz aktivieren") Then
        Exit Sub
    End If
    
    PID_SetupSheetProtectionForMacros
    
    MsgBox "Schutz auf allen Blaettern aktiv.", vbInformation, "Admin"
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


Private Sub PID_EnsureAdminSheetButtons(ByVal ws As Worksheet)
    Dim specs As Variant
    Dim i As Long
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim anchorLeft As Double
    Dim anchorTop As Double
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btn As Shape
    
    If ws Is Nothing Then Exit Sub
    
    specs = PID_AdminButtonSpecs()
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    ws.Unprotect
    On Error GoTo 0
    
    PID_DeleteAdminToolbarShapes ws
    
    anchorLeft = ws.Range("B5").Left
    anchorTop = ws.Range("B5").Top
    
    For i = LBound(specs) To UBound(specs)
        colIndex = i Mod PID_ADMIN_BTN_COLS
        rowIndex = i \ PID_ADMIN_BTN_COLS
        
        btnLeft = anchorLeft + (colIndex * (PID_ADMIN_BTN_WIDTH + PID_ADMIN_BTN_GAP_H))
        btnTop = anchorTop + (rowIndex * (PID_ADMIN_BTN_HEIGHT + PID_ADMIN_BTN_GAP_V))
        
        Set btn = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                     Left:=btnLeft, _
                                     Top:=btnTop, _
                                     Width:=PID_ADMIN_BTN_WIDTH, _
                                     Height:=PID_ADMIN_BTN_HEIGHT)
        
        btn.Name = PID_ADMIN_BTN_PREFIX & CStr(specs(i)(0))
        btn.TextFrame.Characters.Text = CStr(specs(i)(1))
        btn.OnAction = CStr(specs(i)(2))
        
        PID_AdminApplyButtonStyle btn, CLng(specs(i)(3))
    Next i
End Sub


Private Function PID_AdminButtonSpecs() As Variant
    PID_AdminButtonSpecs = Array( _
        Array("Smoke", "Smoke Check", "PID_AdminRunSmokeCheck", 0), _
        Array("Refresh", "Full Refresh", "PID_AdminRunFullRefresh", 0), _
        Array("Quick", "Quick Check", "PID_AdminRunQuickCheck", 1), _
        Array("Format", "Format Monate", "PID_AdminRunFormatMonths", 1), _
        Array("Import", "VBA Import", "PID_AdminRunVbaImport", 2), _
        Array("Perf", "Perf. Baseline", "PID_AdminRunPerfBaseline", 2), _
        Array("Protect", "Schutz AN", "PID_AdminRunProtectAll", 3), _
        Array("Unprotect", "Schutz AUS", "PID_AdminRunUnprotectAll", 2), _
        Array("ShowTech", "Tech-Bl. zeigen", "PID_AdminShowTechSheets", 1), _
        Array("HideTech", "Tech-Bl. verbergen", "PID_AdminHideTechSheets", 1), _
        Array("Lohn", "LOHNTABELLE", "PID_AdminRunRebuildLohn", 2), _
        Array("SmokeSheet", "SYSTEM_CHECK", "PID_AdminOpenSmokeSheet", 3), _
        Array("Hide", "Admin verbergen", "PID_AdminHidePanel", 3) _
    )
End Function


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

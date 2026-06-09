Attribute VB_Name = "mod_PIDCopyright"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Public Const PID_COPYRIGHT_AUTHOR As String = "Adam Nagy"
Public Const PID_COPYRIGHT_ORG As String = "McOpCo"

Private Const PID_COPYRIGHT_LEGACY_CELL As String = "Z100"
Private Const PID_COPYRIGHT_SHAPE_NAME As String = "shp_PID_Copyright"
Private Const PID_COPYRIGHT_SKIP_SHEETS As String = "|FLUKTUATION_DATEN|KV_DROPDOWN_HELPER|"


Public Function PID_GetCopyrightText() As String
    PID_GetCopyrightText = ChrW(169) & " " & PID_COPYRIGHT_AUTHOR & PID_GetCopyrightSeparator() & _
                           PID_COPYRIGHT_ORG & PID_GetCopyrightSeparator() & _
                           "Personalsheet " & CStr(PID_GetWorkbookYear())
End Function


Private Function PID_GetCopyrightSeparator() As String
    PID_GetCopyrightSeparator = " " & ChrW(183) & " "
End Function


Public Sub PID_ApplyCopyrightToAllSheets()
    Dim ws As Worksheet
    
    On Error GoTo SafeExit
    
    For Each ws In ThisWorkbook.Worksheets
        If PID_SheetNeedsCopyrightNotice(ws) Then
            PID_ApplyCopyrightToWorksheet ws
        End If
    Next ws

SafeExit:
End Sub


Public Sub PID_ApplyCopyrightToWorksheet(ByVal ws As Worksheet)
    Dim wasProtected As Boolean
    Dim copyrightText As String
    Dim noticeRange As Range
    Dim rangeAddress As String
    Dim align As Long
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not PID_SheetNeedsCopyrightNotice(ws) Then Exit Sub
    
    copyrightText = PID_GetCopyrightText()
    wasProtected = ws.ProtectContents
    
    If wasProtected Then
        On Error Resume Next
        ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
        On Error GoTo SafeExit
    End If
    
    PID_ClearLegacyCopyrightArtifacts ws
    
    rangeAddress = PID_GetCopyrightRangeAddress(ws)
    align = PID_GetCopyrightHorizontalAlignment(ws)
    Set noticeRange = ws.Range(rangeAddress)
    PID_WriteCopyrightNotice noticeRange, copyrightText, align
    
    If wasProtected Then
        PID_ReprotectWorksheet ws
    End If

SafeExit:
End Sub


Private Function PID_GetCopyrightRangeAddress(ByVal ws As Worksheet) As String
    If PID_IsWorkerMonthSheet(ws) Then
        ' Rechts neben CopyData-Button (O1), nicht darunter.
        PID_GetCopyrightRangeAddress = "S2:V2"
        Exit Function
    End If
    
    If PID_IsUbersichtWorksheet(ws) Then
        PID_GetCopyrightRangeAddress = "B25:H25"
        Exit Function
    End If
    
    If StrComp(ws.Name, PID_EINSTELLUNG_SHEET, vbTextCompare) = 0 Then
        PID_GetCopyrightRangeAddress = "S2:U2"
        Exit Function
    End If
    
    If StrComp(ws.Name, PID_LOHNTABELLE_SHEET, vbTextCompare) = 0 Then
        PID_GetCopyrightRangeAddress = "L2:N2"
        Exit Function
    End If
    
    If StrComp(ws.Name, PID_FLUKTUATION_SHEET, vbTextCompare) = 0 Then
        PID_GetCopyrightRangeAddress = "A3:D3"
        Exit Function
    End If
    
    PID_GetCopyrightRangeAddress = "F2:H2"
End Function


Private Function PID_GetCopyrightHorizontalAlignment(ByVal ws As Worksheet) As Long
    If PID_IsUbersichtWorksheet(ws) Then
        PID_GetCopyrightHorizontalAlignment = xlHAlignLeft
        Exit Function
    End If
    
    If StrComp(ws.Name, PID_FLUKTUATION_SHEET, vbTextCompare) = 0 Then
        PID_GetCopyrightHorizontalAlignment = xlHAlignLeft
        Exit Function
    End If
    
    PID_GetCopyrightHorizontalAlignment = xlHAlignRight
End Function


Private Sub PID_WriteCopyrightNotice(ByVal noticeRange As Range, _
                                     ByVal copyrightText As String, _
                                     ByVal align As Long)
    On Error Resume Next
    noticeRange.UnMerge
    On Error GoTo 0
    
    noticeRange.Merge
    noticeRange.Value = copyrightText
    noticeRange.Font.Name = "Calibri"
    noticeRange.Font.Size = 8
    noticeRange.Font.Color = RGB(110, 128, 148)
    noticeRange.HorizontalAlignment = align
    noticeRange.VerticalAlignment = xlVAlignCenter
    noticeRange.WrapText = False
    noticeRange.Interior.Pattern = xlNone
    noticeRange.Locked = True
End Sub


Private Function PID_SheetNeedsCopyrightNotice(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If ws.Visible <> xlSheetVisible Then Exit Function
    If InStr(1, PID_COPYRIGHT_SKIP_SHEETS, "|" & ws.Name & "|", vbTextCompare) > 0 Then Exit Function
    PID_SheetNeedsCopyrightNotice = True
End Function


Private Sub PID_ClearLegacyCopyrightArtifacts(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim legacyAddress As Variant
    Dim legacyAddresses As Variant
    
    legacyAddresses = Array("O2:Q2", "S2:V2", "M2:O2", "E2:G2", "A3:D3", "B25:H25", "F2:H2", PID_COPYRIGHT_LEGACY_CELL)
    
    On Error Resume Next
    Set shp = ws.Shapes(PID_COPYRIGHT_SHAPE_NAME)
    If Not shp Is Nothing Then shp.Delete
    
    For Each legacyAddress In legacyAddresses
        ws.Range(CStr(legacyAddress)).UnMerge
        ws.Range(CStr(legacyAddress)).ClearContents
    Next legacyAddress
    
    Err.Clear
End Sub

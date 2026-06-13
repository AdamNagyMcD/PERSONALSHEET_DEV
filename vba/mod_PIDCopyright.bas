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
Private Const PID_COPYRIGHT_SKIP_SHEETS As String = "|FLUKTUATION_DATEN|KV_DROPDOWN_HELPER|_ADMIN|"


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
    Dim copyrightText As String

    On Error GoTo SafeExit

    copyrightText = PID_GetCopyrightText()

    For Each ws In ThisWorkbook.Worksheets
        If PID_SheetNeedsCopyrightNotice(ws) Then
            If Not PID_CopyrightAlreadyCurrent(ws, copyrightText) Then
                PID_ApplyCopyrightToWorksheet ws
            End If
        End If
    Next ws

SafeExit:
End Sub


Private Function PID_CopyrightAlreadyCurrent(ByVal ws As Worksheet, ByVal expected As String) As Boolean
    Dim rangeAddr As String
    Dim cellVal As String

    On Error Exit Function

    rangeAddr = PID_GetCopyrightRangeAddress(ws)
    cellVal = Trim$(CStr(ws.Range(rangeAddr).Cells(1, 1).Value))
    PID_CopyrightAlreadyCurrent = (cellVal = expected)
End Function


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
    
    If Not PID_CanWriteCopyrightToRange(noticeRange) Then Exit Sub
    
    PID_WriteCopyrightNotice noticeRange, copyrightText, align
    
    If wasProtected Then
        PID_ReprotectWorksheet ws
    End If

SafeExit:
End Sub


Private Function PID_GetCopyrightRangeAddress(ByVal ws As Worksheet) As String
    If PID_IsWorkerMonthSheet(ws) Then
        PID_GetCopyrightRangeAddress = "S2:V2"
        Exit Function
    End If
    
    If PID_IsUbersichtWorksheet(ws) Then
        PID_GetCopyrightRangeAddress = "B24"
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
        PID_GetCopyrightRangeAddress = "A7"
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
    
    ' S2 (Monatsblatt) und L2 (LOHNTABELLE) linksbuendig.
    If PID_IsWorkerMonthSheet(ws) Then
        PID_GetCopyrightHorizontalAlignment = xlHAlignLeft
        Exit Function
    End If
    
    If StrComp(ws.Name, PID_LOHNTABELLE_SHEET, vbTextCompare) = 0 Then
        PID_GetCopyrightHorizontalAlignment = xlHAlignLeft
        Exit Function
    End If
    
    PID_GetCopyrightHorizontalAlignment = xlHAlignRight
End Function


Private Function PID_CanWriteCopyrightToRange(ByVal noticeRange As Range) As Boolean
    Dim anchor As Range
    
    If noticeRange Is Nothing Then Exit Function
    
    Set anchor = noticeRange.Cells(1, 1)
    
    If Not anchor.MergeCells Then
        PID_CanWriteCopyrightToRange = True
        Exit Function
    End If
    
    If PID_RangeContainsCopyrightText(anchor.MergeArea) Then
        PID_CanWriteCopyrightToRange = True
        Exit Function
    End If
    
    If PID_MergeAreaMatchesRange(anchor.MergeArea, noticeRange) Then
        PID_CanWriteCopyrightToRange = True
        Exit Function
    End If
    
    PID_CanWriteCopyrightToRange = False
End Function


Private Sub PID_WriteCopyrightNotice(ByVal noticeRange As Range, _
                                     ByVal copyrightText As String, _
                                     ByVal align As Long)
    If noticeRange Is Nothing Then Exit Sub
    
    If noticeRange.Cells.Count = 1 Then
        noticeRange.Value = copyrightText
        PID_ApplyCopyrightNoticeStyle noticeRange, align
        Exit Sub
    End If
    
    On Error Resume Next
    If noticeRange.Cells(1, 1).MergeCells Then
        If Not PID_MergeAreaMatchesRange(noticeRange.Cells(1, 1).MergeArea, noticeRange) Then
            If Not PID_RangeContainsCopyrightText(noticeRange.Cells(1, 1).MergeArea) Then Exit Sub
        End If
    End If
    noticeRange.UnMerge
    On Error GoTo 0
    
    noticeRange.Merge
    noticeRange.Value = copyrightText
    PID_ApplyCopyrightNoticeStyle noticeRange, align
End Sub


Private Sub PID_ApplyCopyrightNoticeStyle(ByVal noticeRange As Range, ByVal align As Long)
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
    Dim rng As Range
    
    On Error Resume Next
    Set shp = ws.Shapes(PID_COPYRIGHT_SHAPE_NAME)
    If Not shp Is Nothing Then shp.Delete
    
    legacyAddresses = PID_GetLegacyCopyrightAddressesForSheet(ws)
    For Each legacyAddress In legacyAddresses
        Set rng = ws.Range(CStr(legacyAddress))
        PID_ClearCopyrightRangeIfOurs rng
    Next legacyAddress
    
    Err.Clear
End Sub


Private Function PID_GetLegacyCopyrightAddressesForSheet(ByVal ws As Worksheet) As Variant
    If PID_IsWorkerMonthSheet(ws) Then
        PID_GetLegacyCopyrightAddressesForSheet = Array("O2:Q2", "S2:V2", PID_COPYRIGHT_LEGACY_CELL)
        Exit Function
    End If
    
    If PID_IsUbersichtWorksheet(ws) Then
        PID_GetLegacyCopyrightAddressesForSheet = Array("B25:H25", "S2:U2", PID_COPYRIGHT_LEGACY_CELL)
        Exit Function
    End If
    
    If StrComp(ws.Name, PID_EINSTELLUNG_SHEET, vbTextCompare) = 0 Then
        PID_GetLegacyCopyrightAddressesForSheet = Array(PID_COPYRIGHT_LEGACY_CELL)
        Exit Function
    End If
    
    If StrComp(ws.Name, PID_LOHNTABELLE_SHEET, vbTextCompare) = 0 Then
        PID_GetLegacyCopyrightAddressesForSheet = Array("M2:O2", PID_COPYRIGHT_LEGACY_CELL)
        Exit Function
    End If
    
    If StrComp(ws.Name, PID_FLUKTUATION_SHEET, vbTextCompare) = 0 Then
        PID_GetLegacyCopyrightAddressesForSheet = Array("E2:G2", "A3:D3", "A19", PID_COPYRIGHT_LEGACY_CELL)
        Exit Function
    End If
    
    PID_GetLegacyCopyrightAddressesForSheet = Array("F2:H2", PID_COPYRIGHT_LEGACY_CELL)
End Function


Private Sub PID_ClearCopyrightRangeIfOurs(ByVal rng As Range)
    If rng Is Nothing Then Exit Sub
    If Not PID_RangeContainsCopyrightText(rng) Then Exit Sub
    
    On Error Resume Next
    rng.UnMerge
    rng.ClearContents
    Err.Clear
End Sub


Private Function PID_RangeContainsCopyrightText(ByVal rng As Range) As Boolean
    Dim cell As Range
    
    If rng Is Nothing Then Exit Function
    
    For Each cell In rng.Cells
        If InStr(1, CStr(cell.Value2), PID_COPYRIGHT_AUTHOR, vbTextCompare) > 0 Then
            PID_RangeContainsCopyrightText = True
            Exit Function
        End If
        If InStr(1, CStr(cell.Value2), PID_COPYRIGHT_ORG, vbTextCompare) > 0 Then
            PID_RangeContainsCopyrightText = True
            Exit Function
        End If
    Next cell
End Function


Private Function PID_MergeAreaMatchesRange(ByVal mergeArea As Range, ByVal noticeRange As Range) As Boolean
    If mergeArea Is Nothing Then Exit Function
    If noticeRange Is Nothing Then Exit Function
    
    PID_MergeAreaMatchesRange = _
        (UCase$(mergeArea.Address(False, False)) = UCase$(noticeRange.Address(False, False)))
End Function

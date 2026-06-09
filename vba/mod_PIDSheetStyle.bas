Attribute VB_Name = "mod_PIDSheetStyle"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Referenz: UEBERSICHT FINANZIELL-Block (Original-Design)
Public Const PID_STYLE_TITLE_ROW_HEIGHT As Single = 28
Public Const PID_STYLE_TITLE_SUB_ROW_HEIGHT As Single = 18
Public Const PID_STYLE_HEADER_TOP_ROW_HEIGHT As Single = 24
Public Const PID_STYLE_HEADER_BOTTOM_ROW_HEIGHT As Single = 30
Public Const PID_STYLE_DATA_ROW_HEIGHT As Single = 30
Public Const PID_STYLE_TOTAL_ROW_HEIGHT As Single = 32
Public Const PID_STYLE_COMPACT_BLOCK_TITLE_HEIGHT As Single = 20
Public Const PID_STYLE_COMPACT_HEADER_ROW_HEIGHT As Single = 18
Public Const PID_STYLE_COMPACT_DATA_ROW_HEIGHT As Single = 22
Public Const PID_STYLE_COMPACT_YEAR_ROW_HEIGHT As Single = 24


Public Function PID_StyleColorNavy() As Long
    PID_StyleColorNavy = RGB(31, 78, 121)
End Function


Public Function PID_StyleColorHeaderBg() As Long
    PID_StyleColorHeaderBg = RGB(221, 235, 247)
End Function


Public Function PID_StyleColorAccent() As Long
    PID_StyleColorAccent = RGB(255, 242, 204)
End Function


Public Function PID_StyleColorZebra() As Long
    PID_StyleColorZebra = RGB(248, 248, 248)
End Function


Public Function PID_StyleColorBtnPrimaryLine() As Long
    PID_StyleColorBtnPrimaryLine = RGB(24, 58, 90)
End Function


Public Sub PID_StyleApplyToolbarButton(ByVal btn As Shape, _
                                       ByVal fillColor As Long, _
                                       ByVal lineColor As Long, _
                                       ByVal fontColor As Long)
    Dim shapeRef As Object
    
    On Error GoTo SafeExit
    
    If btn Is Nothing Then Exit Sub
    
    btn.Fill.ForeColor.RGB = fillColor
    btn.Fill.Visible = msoTrue
    btn.Line.ForeColor.RGB = lineColor
    btn.Line.Weight = 0.75
    btn.Line.Visible = msoTrue
    
    On Error Resume Next
    btn.Adjustments(1) = 0.16
    btn.Shadow.Type = msoShadow21
    btn.Shadow.Visible = msoTrue
    btn.Shadow.Transparency = 0.6
    btn.Shadow.Blur = 2
    btn.TextFrame.MarginLeft = 3
    btn.TextFrame.MarginRight = 3
    btn.TextFrame.MarginTop = 1
    btn.TextFrame.MarginBottom = 1
    On Error GoTo SafeExit
    
    btn.TextFrame.Characters.Font.Name = "Calibri"
    btn.TextFrame.Characters.Font.Color = fontColor
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Size = 8
    
    ' TextFrame.WordWrap/Anchor per Early Binding auf Mac nicht verfuegbar (Compile-Fehler).
    On Error Resume Next
    Set shapeRef = btn
    shapeRef.TextFrame.WordWrap = msoTrue
    shapeRef.TextFrame.VerticalAnchor = msoAnchorMiddle
    shapeRef.TextFrame.HorizontalAnchor = msoAnchorCenter
    shapeRef.TextFrame.VerticalAlignment = xlVAlignCenter
    shapeRef.TextFrame.HorizontalAlignment = xlHAlignCenter
    shapeRef.Placement = xlFreeFloating
    Set shapeRef = Nothing
    On Error GoTo SafeExit
    
SafeExit:
End Sub


Public Sub PID_StyleApplyTitleBand(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorNavy()
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub


Public Sub PID_StyleApplyHeaderBand(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorHeaderBg()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub


Public Sub PID_StyleApplySubsectionTitle(ByVal target As Range, Optional ByVal alignLeft As Boolean = True)
    With target
        .Interior.Color = PID_StyleColorHeaderBg()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = True
        .Font.Size = 10
        If alignLeft Then
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        Else
            .HorizontalAlignment = xlCenter
            .IndentLevel = 0
        End If
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub


Public Sub PID_StyleApplyAccentSummaryBand(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorAccent()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub


Public Sub PID_StyleApplyInputHighlight(ByVal target As Range)
    PID_StyleApplyAccentSummaryBand target
End Sub


Public Sub PID_StyleApplyEditableCell(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorAccent()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = False
    End With
End Sub


Public Sub PID_StyleApplyInputCell(ByVal target As Range)
    With target
        .Interior.Color = vbWhite
        .Font.Color = vbBlack
        .Font.Bold = False
    End With
End Sub


Public Sub PID_StyleApplyInputGuideHeader(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorAccent()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub


Public Sub PID_StyleApplyInputGuideLabel(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorAccent()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
End Sub


Public Sub PID_StyleApplyReadOnlyGuideCell(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorAccent()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub


Public Sub PID_StyleApplyReadOnlyCell(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorHeaderBg()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = False
    End With
End Sub


Public Sub PID_StyleApplyCompactHeaderBand(ByVal target As Range)
    With target
        .Interior.Color = PID_StyleColorHeaderBg()
        .Font.Color = PID_StyleColorNavy()
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub


Public Sub PID_StyleApplyZebraRows(ByVal dataRange As Range)
    Dim r As Long
    
    If dataRange Is Nothing Then Exit Sub
    
    For r = 1 To dataRange.Rows.Count
        If ((r - 1) Mod 2) = 1 Then
            dataRange.Rows(r).Interior.Color = PID_StyleColorZebra()
        Else
            dataRange.Rows(r).Interior.Color = vbWhite
        End If
        dataRange.Rows(r).Font.Color = vbBlack
        dataRange.Rows(r).Font.Bold = False
    Next r
End Sub


Public Sub PID_StyleApplyOuterBorder(ByVal target As Range)
    On Error Resume Next
    With target
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = PID_StyleColorNavy()
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeRight).Color = PID_StyleColorNavy()
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).Color = PID_StyleColorNavy()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = PID_StyleColorNavy()
    End With
    On Error GoTo 0
End Sub


Public Sub PID_StyleApplyTableBorders(ByVal tableRange As Range)
    On Error Resume Next
    With tableRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_StyleColorNavy()
    End With
    With tableRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_StyleColorNavy()
    End With
    With tableRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_StyleColorNavy()
    End With
    With tableRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_StyleColorNavy()
    End With
    If tableRange.Rows.Count > 1 Then
        With tableRange.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 180)
        End With
    End If
    If tableRange.Columns.Count > 1 Then
        With tableRange.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 180)
        End With
    End If
    On Error GoTo 0
End Sub

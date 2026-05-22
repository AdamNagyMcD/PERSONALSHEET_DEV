Attribute VB_Name = "mod_FormatEinstellung"
Option Explicit


Public Sub PID_FormatEinstellungSheet()
    Dim ws As Worksheet
    Dim wasProtected As Boolean
    Dim oldScreenUpdating As Boolean
    Dim r As Long
    
    On Error GoTo CleanFail
    
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(PID_EINSTELLUNG_SHEET)
    wasProtected = ws.ProtectContents
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
    On Error GoTo CleanFail
    
    With ws
        On Error Resume Next
        .Range("B1:U1").Merge
        On Error GoTo CleanFail
        If Len(Trim$(CStr(.Range("B1").Value2))) = 0 Then
            .Range("B1").Value = "EINSTELLUNG - Zentrale Konfiguration"
        End If
        .Rows(1).RowHeight = PID_STYLE_TITLE_ROW_HEIGHT
        PID_StyleApplyTitleBand .Range("B1:U1")
        
        .Rows(3).RowHeight = PID_STYLE_COMPACT_BLOCK_TITLE_HEIGHT
        .Rows(4).RowHeight = PID_STYLE_COMPACT_BLOCK_TITLE_HEIGHT
        PID_StyleApplyTitleBand .Range("B3:F4")
        PID_StyleApplyTitleBand .Range("H3:I4")
        PID_StyleApplyTitleBand .Range("K3:L4")
        PID_StyleApplyTitleBand .Range("N3:O4")
        
        .Rows(5).RowHeight = PID_STYLE_COMPACT_HEADER_ROW_HEIGHT
        PID_StyleApplyCompactHeaderBand .Range("B5")
        PID_StyleApplyInputGuideHeader .Range("C5:F5")
        PID_StyleApplyInputGuideHeader .Range("H5:I5")
        PID_StyleApplyInputGuideHeader .Range("K5:L5")
        PID_StyleApplyInputGuideHeader .Range("N5:O5")
        
        PID_StyleApplyZebraRows .Range("B6:F17")
        PID_StyleApplyZebraRows .Range("H6:I17")
        PID_StyleApplyZebraRows .Range("K6:L17")
        PID_StyleApplyZebraRows .Range("N6:O17")
        PID_StyleApplyTableBorders .Range("C5:F17")
        PID_StyleApplyTableBorders .Range("H5:I17")
        PID_StyleApplyTableBorders .Range("K5:L17")
        PID_StyleApplyTableBorders .Range("N5:O17")
        
        .Range("C6:F17").HorizontalAlignment = xlCenter
        .Range("H6:I17").HorizontalAlignment = xlCenter
        .Range("K6:L17").HorizontalAlignment = xlCenter
        .Range("N6:O17").HorizontalAlignment = xlCenter
        
        .Rows(19).RowHeight = PID_STYLE_COMPACT_BLOCK_TITLE_HEIGHT
        .Rows(20).RowHeight = PID_STYLE_COMPACT_BLOCK_TITLE_HEIGHT
        PID_StyleApplyTitleBand .Range("B19:F20")
        PID_StyleApplyTitleBand .Range("H19:I20")
        PID_StyleApplyTitleBand .Range("K19:L20")
        
        .Rows(21).RowHeight = PID_STYLE_COMPACT_HEADER_ROW_HEIGHT
        PID_StyleApplyCompactHeaderBand .Range("B21")
        PID_StyleApplyInputGuideHeader .Range("C21:F21")
        PID_StyleApplyInputGuideHeader .Range("H21:I21")
        PID_StyleApplyInputGuideHeader .Range("K21:L21")
        
        PID_StyleApplyZebraRows .Range("B22:F33")
        PID_StyleApplyZebraRows .Range("H22:I33")
        PID_StyleApplyZebraRows .Range("K22:L33")
        PID_StyleApplyTableBorders .Range("C21:F33")
        PID_StyleApplyTableBorders .Range("H21:I33")
        PID_StyleApplyTableBorders .Range("K21:L33")
        
        .Range("C22:F33").HorizontalAlignment = xlCenter
        .Range("H22:I33").HorizontalAlignment = xlCenter
        .Range("K22:L33").HorizontalAlignment = xlCenter
        
        .Rows(35).RowHeight = PID_STYLE_COMPACT_YEAR_ROW_HEIGHT
        PID_StyleApplyInputGuideLabel .Range("B35")
        PID_StyleApplyInputCell .Range("C35")
        PID_StyleApplyTableBorders .Range("B35:C35")
        
        .Rows(37).RowHeight = PID_STYLE_COMPACT_HEADER_ROW_HEIGHT
        PID_StyleApplyInputGuideHeader .Range("B37")
        PID_StyleApplyInputGuideHeader .Range("C37")
        PID_StyleApplyZebraRows .Range("B38:C49")
        PID_StyleApplyTableBorders .Range("B37:C49")
        .Range("C38:C49").HorizontalAlignment = xlCenter
        
        .Rows(52).RowHeight = PID_STYLE_COMPACT_HEADER_ROW_HEIGHT
        PID_StyleApplyInputGuideHeader .Range("B52")
        PID_StyleApplyInputGuideHeader .Range("C52")
        PID_StyleApplyZebraRows .Range("B53:C59")
        PID_StyleApplyTableBorders .Range("B52:C59")
        .Range("C53:C59").HorizontalAlignment = xlCenter
        
        PID_ApplyEuroNumberFormat .Range("C6:D17")
        PID_ApplyEuroNumberFormat .Range("C22:D33")
        .Range("F6:F17").NumberFormat = "#,##0"
        .Range("F22:F33").NumberFormat = "#,##0"
        .Range("E22:E33").NumberFormat = "0.00%"
        .Range("H6:I17").NumberFormat = "0.00"
        .Range("K6:L17").NumberFormat = "0.00"
        .Range("N6:O17").NumberFormat = "0.00"
        .Range("H22:I33").NumberFormat = "0.00"
        .Range("K22:L33").NumberFormat = "0.00"
        .Range("C35").NumberFormat = "0"
        .Range("C38:C49").NumberFormat = "0.00"
        .Range("C53:C59").NumberFormat = "0.00"
        
        For r = 6 To 17
            .Rows(r).RowHeight = PID_STYLE_COMPACT_DATA_ROW_HEIGHT
        Next r
        For r = 22 To 33
            .Rows(r).RowHeight = PID_STYLE_COMPACT_DATA_ROW_HEIGHT
        Next r
        For r = 38 To 49
            .Rows(r).RowHeight = PID_STYLE_COMPACT_DATA_ROW_HEIGHT
        Next r
        For r = 53 To 59
            .Rows(r).RowHeight = PID_STYLE_COMPACT_DATA_ROW_HEIGHT
        Next r
        
        PID_ESApplyEinstellungInputZones ws
        
        .Range("B1:U59").VerticalAlignment = xlCenter
    End With
    
    PID_ESApplyCompactMonthSheetRowHeights
    
    GoTo CleanProtect

CleanFail:
    MsgBox "Fehler bei FormatEinstellung:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "EINSTELLUNG"

CleanProtect:
    On Error Resume Next
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    On Error GoTo 0

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
End Sub


Private Sub PID_ESApplyEinstellungInputZones(ByVal ws As Worksheet)
    ' Gelb = Hinweis (wo eintragen). Weiss = tatsaechliche Eingabe.
    PID_StyleApplyInputGuideLabel ws.Range("B6:B17")
    PID_StyleApplyInputGuideLabel ws.Range("B22:B33")
    PID_StyleApplyInputGuideLabel ws.Range("B38:B49")
    PID_StyleApplyInputGuideLabel ws.Range("B53:B59")
    
    PID_StyleApplyInputCell ws.Range("C6:F17")
    PID_StyleApplyInputCell ws.Range("H6:I17")
    PID_StyleApplyInputCell ws.Range("K6:L17")
    PID_StyleApplyInputCell ws.Range("N6:O17")
    
    PID_StyleApplyInputCell ws.Range("C22:F33")
    PID_StyleApplyInputCell ws.Range("H22:I33")
    PID_StyleApplyInputCell ws.Range("K22:L33")
    
    PID_StyleApplyInputCell ws.Range("C38:C49")
    PID_StyleApplyInputCell ws.Range("C53:C59")
End Sub


Private Sub PID_ESApplyCompactMonthSheetRowHeights()
    Dim monthNames As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim wasProtected As Boolean
    
    monthNames = PID_MonthNames()
    
    For i = LBound(monthNames) To UBound(monthNames)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(monthNames(i)))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            If IsNumeric(ws.Range("A1").Value) Then
                wasProtected = ws.ProtectContents
                On Error Resume Next
                ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
                On Error GoTo 0
                
                ws.Rows("1:2").RowHeight = PID_STYLE_COMPACT_HEADER_ROW_HEIGHT
                ws.Rows(PID_FIRST_ROW & ":" & PID_LAST_ROW).RowHeight = PID_STYLE_COMPACT_DATA_ROW_HEIGHT
                
                On Error Resume Next
                If wasProtected Then
                    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                               UserInterfaceOnly:=True, _
                               AllowFiltering:=True, _
                               AllowSorting:=True
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

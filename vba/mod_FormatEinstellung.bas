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
        .Rows(1).RowHeight = 28
        PID_ESApplyTitleBand .Range("B1:U1")
        
        .Rows(3).RowHeight = 24
        .Rows(4).RowHeight = 24
        PID_ESApplyTitleBand .Range("B3:F4")
        PID_ESApplyTitleBand .Range("H3:I4")
        PID_ESApplyTitleBand .Range("K3:L4")
        PID_ESApplyTitleBand .Range("N3:O4")
        
        .Rows(5).RowHeight = 24
        PID_ESApplyHeaderBand .Range("C5:F5")
        PID_ESApplyHeaderBand .Range("H5:I5")
        PID_ESApplyHeaderBand .Range("K5:L5")
        PID_ESApplyHeaderBand .Range("N5:O5")
        
        PID_ESApplyZebraRows .Range("B6:F17")
        PID_ESApplyZebraRows .Range("H6:I17")
        PID_ESApplyZebraRows .Range("K6:L17")
        PID_ESApplyZebraRows .Range("N6:O17")
        PID_ESApplyTableBorders .Range("C5:F17")
        PID_ESApplyTableBorders .Range("H5:I17")
        PID_ESApplyTableBorders .Range("K5:L17")
        PID_ESApplyTableBorders .Range("N5:O17")
        
        .Range("B6:B17").Font.Bold = True
        .Range("B6:B17").Font.Color = PID_ESColorNavy()
        .Range("B6:B17").HorizontalAlignment = xlLeft
        .Range("C6:F17").HorizontalAlignment = xlCenter
        .Range("H6:I17").HorizontalAlignment = xlCenter
        .Range("K6:L17").HorizontalAlignment = xlCenter
        .Range("N6:O17").HorizontalAlignment = xlCenter
        
        .Rows(19).RowHeight = 24
        .Rows(20).RowHeight = 24
        PID_ESApplyTitleBand .Range("B19:F20")
        PID_ESApplyTitleBand .Range("H19:I20")
        PID_ESApplyTitleBand .Range("K19:L20")
        
        .Rows(21).RowHeight = 24
        PID_ESApplyHeaderBand .Range("C21:F21")
        PID_ESApplyHeaderBand .Range("H21:I21")
        PID_ESApplyHeaderBand .Range("K21:L21")
        
        PID_ESApplyZebraRows .Range("B22:F33")
        PID_ESApplyZebraRows .Range("H22:I33")
        PID_ESApplyZebraRows .Range("K22:L33")
        PID_ESApplyTableBorders .Range("C21:F33")
        PID_ESApplyTableBorders .Range("H21:I33")
        PID_ESApplyTableBorders .Range("K21:L33")
        
        .Range("B22:B33").Font.Bold = True
        .Range("B22:B33").Font.Color = PID_ESColorNavy()
        .Range("B22:B33").HorizontalAlignment = xlLeft
        .Range("C22:F33").HorizontalAlignment = xlCenter
        .Range("H22:I33").HorizontalAlignment = xlCenter
        .Range("K22:L33").HorizontalAlignment = xlCenter
        
        .Rows(35).RowHeight = 28
        PID_ESApplyHeaderBand .Range("B35")
        .Range("B35").HorizontalAlignment = xlLeft
        .Range("B35").IndentLevel = 1
        With .Range("C35")
            .Interior.Color = PID_ESColorAccent()
            .Font.Color = PID_ESColorNavy()
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        PID_ESApplyTableBorders .Range("B35:C35")
        
        .Rows(37).RowHeight = 24
        PID_ESApplyHeaderBand .Range("B37:C37")
        PID_ESApplyZebraRows .Range("B38:C49")
        PID_ESApplyTableBorders .Range("B37:C49")
        .Range("B38:B49").HorizontalAlignment = xlLeft
        .Range("C38:C49").HorizontalAlignment = xlCenter
        
        .Rows(52).RowHeight = 24
        PID_ESApplyHeaderBand .Range("B52:C52")
        PID_ESApplyZebraRows .Range("B53:C59")
        PID_ESApplyTableBorders .Range("B52:C59")
        .Range("B53:B59").HorizontalAlignment = xlLeft
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
            .Rows(r).RowHeight = 22
        Next r
        For r = 22 To 33
            .Rows(r).RowHeight = 22
        Next r
        For r = 38 To 49
            .Rows(r).RowHeight = 22
        Next r
        For r = 53 To 59
            .Rows(r).RowHeight = 22
        Next r
        
        .Range("B1:U59").VerticalAlignment = xlCenter
    End With
    
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


Private Function PID_ESColorNavy() As Long
    PID_ESColorNavy = RGB(31, 78, 121)
End Function


Private Function PID_ESColorAccent() As Long
    PID_ESColorAccent = RGB(255, 242, 204)
End Function


Private Function PID_ESColorZebra() As Long
    PID_ESColorZebra = RGB(248, 248, 248)
End Function


Private Sub PID_ESApplyTitleBand(ByVal target As Range)
    With target
        .Interior.Color = PID_ESColorNavy()
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub


Private Sub PID_ESApplyHeaderBand(ByVal target As Range)
    With target
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = PID_ESColorNavy()
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub


Private Sub PID_ESApplyZebraRows(ByVal dataRange As Range)
    Dim r As Long
    
    For r = 1 To dataRange.Rows.Count
        If ((r - 1) Mod 2) = 1 Then
            dataRange.Rows(r).Interior.Color = PID_ESColorZebra()
        Else
            dataRange.Rows(r).Interior.Color = vbWhite
        End If
        dataRange.Rows(r).Font.Color = vbBlack
        dataRange.Rows(r).Font.Bold = False
    Next r
End Sub


Private Sub PID_ESApplyTableBorders(ByVal tableRange As Range)
    On Error Resume Next
    With tableRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_ESColorNavy()
    End With
    With tableRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_ESColorNavy()
    End With
    With tableRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_ESColorNavy()
    End With
    With tableRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = PID_ESColorNavy()
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

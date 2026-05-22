Attribute VB_Name = "mod_BuildDurchrechnungUebersicht"
Option Explicit

Private Const PID_DR_START_ROW As Long = 28
Private Const PID_DR_END_ROW As Long = 38
Private Const PID_DR_FIRST_COL As Long = 2   ' B
Private Const PID_DR_LAST_COL As Long = 10   ' J

Private Const PID_DR_JAEN_VERF_CELL As String = "E30"
Private Const PID_DR_JAEN_MUST_CELL As String = "G30"

Private Const PID_UBERSICHT_SHEET As String = "UBERSICHT"


Public Sub PID_BuildDurchrechnungUebersicht()
    PID_BuildDurchrechnungUebersichtInternal True
End Sub


Public Sub PID_RefreshDurchrechnungUebersicht()
    Dim ws As Worksheet
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(PID_UBERSICHT_SHEET)
    
    If Not PID_DurchrechnungBlockExists(ws) Then
        PID_BuildDurchrechnungUebersichtInternal False
        GoTo CleanExit
    End If
    
    If Not PID_DurchrechnungBlockHasLohnColumn(ws) Then
        PID_BuildDurchrechnungUebersichtInternal False
        GoTo CleanExit
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    ws.Calculate
    PID_UpdateDurchrechnungLohnFormulas ws
    PID_UpdateDurchrechnungEuroFormulas ws
    PID_UnlockDurchrechnungInputs ws
    PID_ApplyDurchrechnungFormats ws
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    
    GoTo CleanExit

CleanFail:
    ' Keine Meldung beim Blattwechsel.

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Private Sub PID_BuildDurchrechnungUebersichtInternal(ByVal showMessage As Boolean)
    Dim ws As Worksheet
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim savedJaenVerf As Variant
    Dim savedJaenMust As Variant
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ws = ThisWorkbook.Worksheets("UBERSICHT")
    
    savedJaenVerf = ws.Range(PID_DR_JAEN_VERF_CELL).Value2
    savedJaenMust = ws.Range(PID_DR_JAEN_MUST_CELL).Value2
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    PID_UnmergeDurchrechnungBlock ws
    PID_ClearDurchrechnungBlock ws
    PID_WriteDurchrechnungBlock ws
    
    If Len(Trim$(CStr(savedJaenVerf))) > 0 Then ws.Range(PID_DR_JAEN_VERF_CELL).Value2 = savedJaenVerf
    If Len(Trim$(CStr(savedJaenMust))) > 0 Then ws.Range(PID_DR_JAEN_MUST_CELL).Value2 = savedJaenMust
    
    PID_UnlockDurchrechnungInputs ws
    PID_ApplyDurchrechnungFormats ws
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    
    If showMessage Then
        MsgBox "Durchrechnungsblock auf UEBERSICHT wurde erstellt." & vbCrLf & vbCrLf & _
               "Gelbe Felder (nur Jaenner-Plan):" & vbCrLf & _
               "- Jaenner Verfuegbar Plan (E30)" & vbCrLf & _
               "- Jaenner Muster Plan (G30)" & vbCrLf & vbCrLf & _
               "Stundenlohn pro Zeile = AVG Bruttolohn/h aus dem Schlussmonat (Q42).", _
               vbInformation, "Durchrechnung"
    End If
    
    GoTo CleanExit

CleanFail:
    If showMessage Then
        MsgBox "Fehler bei PID_BuildDurchrechnungUebersicht:" & vbCrLf & _
               Err.Number & " - " & Err.Description, _
               vbExclamation, "Durchrechnung"
    End If

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Private Function PID_DurchrechnungBlockExists(ByVal ws As Worksheet) As Boolean
    Dim titleText As String
    
    On Error GoTo SafeExit
    
    titleText = CStr(ws.Cells(PID_DR_START_ROW, 2).Text)
    
    If InStr(1, titleText, "DURCHRECHNUNGSSTUNDEN", vbTextCompare) > 0 Then
        PID_DurchrechnungBlockExists = True
        Exit Function
    End If
    
    If Trim$(CStr(ws.Cells(PID_DR_START_ROW + 3, 2).Value)) = "Zeitraum" Then
        PID_DurchrechnungBlockExists = True
        Exit Function
    End If

SafeExit:
    PID_DurchrechnungBlockExists = False
End Function


Private Function PID_DurchrechnungBlockHasLohnColumn(ByVal ws As Worksheet) As Boolean
    Dim headerText As String
    
    On Error GoTo SafeExit
    
    headerText = Trim$(CStr(ws.Cells(PID_DR_START_ROW + 3, 7).Value))
    
    If InStr(1, headerText, "Lohn", vbTextCompare) > 0 Then
        PID_DurchrechnungBlockHasLohnColumn = True
        Exit Function
    End If

SafeExit:
    PID_DurchrechnungBlockHasLohnColumn = False
End Function


Private Sub PID_UnmergeDurchrechnungBlock(ByVal ws As Worksheet)
    Dim cell As Range
    Dim mergedAreas As Collection
    Dim areaKey As String
    
    Set mergedAreas = New Collection
    
    For Each cell In ws.Range("B" & PID_DR_START_ROW & ":J" & PID_DR_END_ROW).Cells
        If cell.MergeCells Then
            areaKey = cell.MergeArea.Address(False, False)
            
            If Not PID_DRCollectionHasKey(mergedAreas, areaKey) Then
                mergedAreas.Add areaKey, areaKey
                cell.MergeArea.UnMerge
            End If
        End If
    Next cell
End Sub


Private Function PID_DRCollectionHasKey(ByVal col As Collection, ByVal key As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    tmp = col.item(key)
    PID_DRCollectionHasKey = True
    Exit Function

NotFound:
    PID_DRCollectionHasKey = False
End Function


Private Sub PID_ClearDurchrechnungBlock(ByVal ws As Worksheet)
    Dim target As Range
    
    Set target = ws.Range(ws.Cells(PID_DR_START_ROW, PID_DR_FIRST_COL), ws.Cells(PID_DR_END_ROW, PID_DR_LAST_COL))
    
    target.ClearContents
    target.Interior.Pattern = xlNone
    target.Font.Bold = False
    target.Font.Color = vbBlack
    target.HorizontalAlignment = xlGeneral
    target.VerticalAlignment = xlCenter
    target.WrapText = False
    
    On Error Resume Next
    target.FormatConditions.Delete
    On Error GoTo 0
End Sub


Private Sub PID_WriteDurchrechnungBlock(ByVal ws As Worksheet)
    Dim titleRow As Long
    Dim hintRow As Long
    Dim inputRow As Long
    Dim headerRow As Long
    Dim dataRow As Long
    Dim noteRow As Long
    
    titleRow = PID_DR_START_ROW
    hintRow = titleRow + 1
    inputRow = titleRow + 2
    headerRow = titleRow + 3
    noteRow = headerRow + 5
    
    With ws
        .Range("B" & titleRow & ":J" & titleRow).Merge
        .Cells(titleRow, 2).Formula = "= ""DURCHRECHNUNGSSTUNDEN - "" & EINSTELLUNG!C35"
        .Cells(titleRow, 2).Font.Bold = True
        .Cells(titleRow, 2).Font.Size = 12
        
        .Range("B" & hintRow & ":J" & hintRow).Merge
        .Cells(hintRow, 2).Value = _
            "Jede Zeile = 3 Monate bis zum Schlussmonat. " & _
            "Differenz rot = zu wenig Stunden (Ueberstunden-Risiko), gelb = Reserve. " & _
            "Ueberstunden-EUR nutzt den AVG Bruttolohn/h aus dem Schlussmonat (Monatsblatt Q42). " & _
            "Nur Jaenner-Plan (naechstes Jahr) ist gelb und manuell."
        .Cells(hintRow, 2).Font.Size = 9
        .Cells(hintRow, 2).WrapText = True
        
        .Range("B" & inputRow & ":D" & inputRow).Merge
        .Cells(inputRow, 2).Value = "Jaenner Verfuegbar Plan (naechstes Jahr):"
        .Range("F" & inputRow & ":G" & inputRow).Merge
        .Cells(inputRow, 6).Value = "Jaenner Muster Plan (naechstes Jahr):"
        
        .Cells(headerRow, 2).Value = "Zeitraum"
        .Cells(headerRow, 3).Value = "Schlussmonat"
        .Cells(headerRow, 4).Value = "Verfuegbar Summe"
        .Cells(headerRow, 5).Value = "Muster Summe"
        .Cells(headerRow, 6).Value = "Differenz"
        .Cells(headerRow, 7).Value = "AVG Lohn/h"
        .Cells(headerRow, 8).Value = "Ueberstunden Std"
        .Cells(headerRow, 9).Value = "Ueberstunden EUR"
        .Cells(headerRow, 10).Value = "Status"
        .Rows(headerRow).Font.Bold = True
        
        dataRow = headerRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, "Feb-Maer-Apr", "April", _
            "=Februar!Q12+Marz!Q13+April!Q13", _
            "=EINSTELLUNG!L7+EINSTELLUNG!L8+EINSTELLUNG!L9", _
            "=April!Q15", _
            "=April!Q42"
        
        dataRow = dataRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, "Mai-Jun-Jul", "Juli", _
            "=Mai!Q12+Juni!Q13+Juli!Q13", _
            "=EINSTELLUNG!L10+EINSTELLUNG!L11+EINSTELLUNG!L12", _
            "=Juli!Q15", _
            "=Juli!Q42"
        
        dataRow = dataRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, "Aug-Sep-Okt", "Oktober", _
            "=August!Q12+September!Q13+Oktober!Q13", _
            "=EINSTELLUNG!L13+EINSTELLUNG!L14+EINSTELLUNG!L15", _
            "=Oktober!Q15", _
            "=Oktober!Q42"
        
        dataRow = dataRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, "Nov-Dez-Jaen", "Jaen (Plan)", _
            "=November!Q12+Dezember!Q13+" & PID_DR_JAEN_VERF_CELL, _
            "=EINSTELLUNG!L16+EINSTELLUNG!L17+" & PID_DR_JAEN_MUST_CELL, _
            "=D" & dataRow & "-E" & dataRow, _
            "=Dezember!Q42"
        
        .Range("B" & noteRow & ":J" & noteRow).Merge
        .Cells(noteRow, 2).Value = _
            "Ueberstunden EUR = Ueberstunden Std x AVG Lohn/h x 1,5 (Spalte G aus Schlussmonat Q42). " & _
            "Nur bei negativem Schluss (rote Differenz). Positive Differenz = Reserve ohne EUR-Kosten."
        .Cells(noteRow, 2).Font.Size = 9
        .Cells(noteRow, 2).WrapText = True
    End With
End Sub


Private Sub PID_WriteDurchrechnungDataRow(ByVal ws As Worksheet, _
                                          ByVal dataRow As Long, _
                                          ByVal periodLabel As String, _
                                          ByVal closingMonthLabel As String, _
                                          ByVal verfFormula As String, _
                                          ByVal musterFormula As String, _
                                          ByVal diffFormula As String, _
                                          ByVal lohnFormula As String)
    With ws
        .Cells(dataRow, 2).Value = periodLabel
        .Cells(dataRow, 3).Value = closingMonthLabel
        .Cells(dataRow, 4).Formula = verfFormula
        .Cells(dataRow, 5).Formula = musterFormula
        .Cells(dataRow, 6).Formula = diffFormula
        .Cells(dataRow, 7).Formula = lohnFormula
        .Cells(dataRow, 8).Formula = "=MAX(0,-F" & dataRow & ")"
        .Cells(dataRow, 9).Formula = PID_GetEuroFormula(dataRow)
        .Cells(dataRow, 10).Formula = "=IF(F" & dataRow & "<0,""ACHTUNG: Ueberstunden-Risiko"",IF(F" & dataRow & ">0,""Hinweis: Reserve / Minus-Stunden moeglich"",""OK""))"
    End With
End Sub


Private Function PID_GetEuroFormula(ByVal dataRow As Long) As String
    ' 3/2 statt 1,5 wegen Excel-2016/Mac Dezimal-Trennzeichen.
    PID_GetEuroFormula = "=IF(OR(G" & dataRow & "=0,H" & dataRow & "=0),"""",H" & dataRow & "*G" & dataRow & "*3/2)"
End Function


Private Sub PID_UpdateDurchrechnungLohnFormulas(ByVal ws As Worksheet)
    Dim dataRow As Long
    Dim headerRow As Long
    Dim lohnFormulas As Variant
    
    headerRow = PID_DR_START_ROW + 3
    lohnFormulas = Array("=April!Q42", "=Juli!Q42", "=Oktober!Q42", "=Dezember!Q42")
    
    For dataRow = headerRow + 1 To headerRow + 4
        ws.Cells(dataRow, 7).Formula = CStr(lohnFormulas(dataRow - headerRow - 1))
    Next dataRow
End Sub


Private Sub PID_UpdateDurchrechnungEuroFormulas(ByVal ws As Worksheet)
    Dim dataRow As Long
    Dim headerRow As Long
    
    headerRow = PID_DR_START_ROW + 3
    
    For dataRow = headerRow + 1 To headerRow + 4
        ws.Cells(dataRow, 9).Formula = PID_GetEuroFormula(dataRow)
    Next dataRow
End Sub


Private Sub PID_UnlockDurchrechnungInputs(ByVal ws As Worksheet)
    ws.Range(PID_DR_JAEN_VERF_CELL).Locked = False
    ws.Range(PID_DR_JAEN_MUST_CELL).Locked = False
End Sub


Private Sub PID_ApplyDurchrechnungFormats(ByVal ws As Worksheet)
    Dim headerRow As Long
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim inputRow As Long
    Dim hintRow As Long
    Dim noteRow As Long
    Dim titleRow As Long
    Dim dataRow As Long
    Dim diffRange As Range
    Dim statusRange As Range
    Dim ueberRange As Range
    Dim blockRange As Range
    Dim tableRange As Range
    
    titleRow = PID_DR_START_ROW
    hintRow = PID_DR_START_ROW + 1
    inputRow = PID_DR_START_ROW + 2
    headerRow = PID_DR_START_ROW + 3
    dataStartRow = headerRow + 1
    dataEndRow = dataStartRow + 3
    noteRow = headerRow + 5
    
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 11
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 15
    ws.Columns("J").ColumnWidth = 28
    
    ws.Rows(titleRow).RowHeight = 28
    ws.Rows(hintRow).RowHeight = 52
    ws.Rows(inputRow).RowHeight = 24
    ws.Rows(headerRow).RowHeight = 36
    ws.Rows(noteRow).RowHeight = 44
    
    For dataRow = dataStartRow To dataEndRow
        ws.Rows(dataRow).RowHeight = 36
    Next dataRow
    
    Set blockRange = ws.Range("B" & titleRow & ":J" & noteRow)
    Set tableRange = ws.Range("B" & headerRow & ":J" & dataEndRow)
    
    With ws.Range("B" & titleRow & ":J" & titleRow)
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("B" & hintRow & ":J" & hintRow)
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(89, 89, 89)
        .Font.Italic = True
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    With ws.Range("B" & inputRow & ":J" & inputRow)
        .Interior.Color = RGB(255, 242, 204)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("B" & inputRow).Font.Bold = True
    ws.Range("F" & inputRow).Font.Bold = True
    
    With ws.Range(PID_DR_JAEN_VERF_CELL)
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = RGB(255, 192, 0)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).Color = RGB(255, 192, 0)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = RGB(255, 192, 0)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeRight).Color = RGB(255, 192, 0)
    End With
    
    With ws.Range(PID_DR_JAEN_MUST_CELL)
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = RGB(255, 192, 0)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).Color = RGB(255, 192, 0)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = RGB(255, 192, 0)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeRight).Color = RGB(255, 192, 0)
    End With
    
    With ws.Range("B" & headerRow & ":J" & headerRow)
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    For dataRow = dataStartRow To dataEndRow
        If ((dataRow - dataStartRow) Mod 2) = 1 Then
            ws.Range("B" & dataRow & ":E" & dataRow).Interior.Color = RGB(248, 248, 248)
            ws.Range("G" & dataRow & ":J" & dataRow).Interior.Color = RGB(248, 248, 248)
        End If
    Next dataRow
    
    ws.Range("B" & dataStartRow & ":C" & dataEndRow).HorizontalAlignment = xlLeft
    ws.Range("D" & dataStartRow & ":I" & dataEndRow).HorizontalAlignment = xlRight
    ws.Range("J" & dataStartRow & ":J" & dataEndRow).HorizontalAlignment = xlLeft
    ws.Range("J" & dataStartRow & ":J" & dataEndRow).WrapText = True
    ws.Range("B" & dataStartRow & ":J" & dataEndRow).VerticalAlignment = xlCenter
    
    ws.Range("D" & dataStartRow & ":F" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range("G" & dataStartRow & ":G" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range("H" & dataStartRow & ":H" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range("I" & dataStartRow & ":I" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range(PID_DR_JAEN_VERF_CELL).NumberFormat = "#,##0.00"
    ws.Range(PID_DR_JAEN_MUST_CELL).NumberFormat = "#,##0.00"
    
    With ws.Range("B" & noteRow & ":J" & noteRow)
        .Interior.Color = RGB(245, 245, 245)
        .Font.Color = RGB(89, 89, 89)
        .Font.Italic = True
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    With blockRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With
    With blockRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With
    With blockRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With
    With blockRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 121)
    End With
    
    With tableRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With tableRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With tableRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With tableRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With tableRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With tableRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Set diffRange = ws.Range("F" & dataStartRow & ":F" & dataEndRow)
    On Error Resume Next
    diffRange.FormatConditions.Delete
    On Error GoTo 0
    
    diffRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    With diffRange.FormatConditions(diffRange.FormatConditions.Count)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With
    
    diffRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    With diffRange.FormatConditions(diffRange.FormatConditions.Count)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
        .Font.Bold = True
    End With
    
    diffRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
    With diffRange.FormatConditions(diffRange.FormatConditions.Count)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = True
    End With
    
    Set statusRange = ws.Range("J" & dataStartRow & ":J" & dataEndRow)
    On Error Resume Next
    statusRange.FormatConditions.Delete
    On Error GoTo 0
    
    statusRange.FormatConditions.Add Type:=xlTextString, String:="ACHTUNG", TextOperator:=xlContains
    With statusRange.FormatConditions(statusRange.FormatConditions.Count)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With
    
    statusRange.FormatConditions.Add Type:=xlTextString, String:="Hinweis", TextOperator:=xlContains
    With statusRange.FormatConditions(statusRange.FormatConditions.Count)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
        .Font.Bold = True
    End With
    
    statusRange.FormatConditions.Add Type:=xlTextString, String:="OK", TextOperator:=xlContains
    With statusRange.FormatConditions(statusRange.FormatConditions.Count)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = True
    End With
    
    Set ueberRange = ws.Range("H" & dataStartRow & ":I" & dataEndRow)
    On Error Resume Next
    ueberRange.FormatConditions.Delete
    On Error GoTo 0
    
    ueberRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    With ueberRange.FormatConditions(ueberRange.FormatConditions.Count)
        .Interior.Color = RGB(252, 228, 214)
        .Font.Color = RGB(132, 46, 43)
        .Font.Bold = True
    End With
End Sub


Public Sub PID_FormatDurchrechnungUebersicht()
    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(PID_UBERSICHT_SHEET)
    
    If Not PID_DurchrechnungBlockExists(ws) Then
        MsgBox "Kein Durchrechnungsblock auf UEBERSICHT gefunden. Bitte zuerst BuildDurchrechnungUebersicht ausfuehren.", _
               vbExclamation, "Durchrechnung"
        GoTo CleanExit
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    PID_UnlockDurchrechnungInputs ws
    PID_ApplyDurchrechnungFormats ws
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    
    GoTo CleanExit

CleanFail:
    MsgBox "Fehler bei FormatDurchrechnungUebersicht:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Durchrechnung"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
End Sub

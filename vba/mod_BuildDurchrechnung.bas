Attribute VB_Name = "mod_BuildDurchrechnung"
Option Explicit

Private Const PID_DR_START_ROW As Long = 28
Private Const PID_DR_END_ROW As Long = 38
Private Const PID_DR_FIRST_COL As Long = 2   ' B
Private Const PID_DR_LAST_COL As Long = 17   ' Q (align with FINANZIELL block above)
Private Const PID_DR_STATUS_FIRST_COL As Long = 10 ' J
Private Const PID_DR_STATUS_LAST_COL As Long = 17  ' Q

Private Const PID_DR_JAEN_VERF_CELL As String = "E30"
Private Const PID_DR_JAEN_MUST_CELL As String = "I30"

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
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
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
    If Len(Trim$(CStr(savedJaenMust))) = 0 Then
        savedJaenMust = ws.Range("G30").Value2
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
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
               "- Jaenner Muster Plan (I30)" & vbCrLf & vbCrLf & _
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
    
    For Each cell In ws.Range("B" & PID_DR_START_ROW & ":Q" & PID_DR_END_ROW).Cells
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
    
    PID_UnmergeDurchrechnungBlock ws
    
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
        .Cells(titleRow, 2).Formula = "= ""DURCHRECHNUNGSSTUNDEN - "" & EINSTELLUNG!C35"
        .Cells(titleRow, 2).Font.Bold = True
        .Cells(titleRow, 2).Font.Size = 12
        
        .Cells(hintRow, 2).Value = _
            "Jede Zeile = 3 Monate bis zum Schlussmonat. " & _
            "Differenz rot = zu wenig Stunden (Ueberstunden-Risiko), gelb = Reserve. " & _
            "Ueberstunden-EUR nutzt den AVG Bruttolohn/h aus dem Schlussmonat (Monatsblatt Q42). " & _
            "Nur Jaenner-Plan (naechstes Jahr) ist gelb und manuell."
        .Cells(hintRow, 2).Font.Size = 9
        .Cells(hintRow, 2).WrapText = True
        
        .Cells(inputRow, 2).Value = "Jaenner Verfuegbar Plan (naechstes Jahr):"
        .Cells(inputRow, 6).Value = "Jaenner Muster Plan (naechstes Jahr):"
        
        .Cells(headerRow, 2).Value = "Zeitraum"
        .Cells(headerRow, 3).Value = "Endmonat"
        .Cells(headerRow, 4).Value = "Verfuegbar"
        .Cells(headerRow, 5).Value = "Muster"
        .Cells(headerRow, 6).Value = "Differenz"
        .Cells(headerRow, 7).Value = "Lohn/h"
        .Cells(headerRow, 8).Value = "Std"
        .Cells(headerRow, 9).Value = "EUR"
        .Cells(headerRow, 10).Value = "Status / Hinweis"
        .Range("K" & headerRow & ":Q" & headerRow).ClearContents
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
        .Cells(dataRow, 10).Formula = "=IF(F" & dataRow & "<0,""ACHTUNG: Ueberstunden"",IF(F" & dataRow & ">0,""Reserve moeglich"",""OK""))"
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
    Dim tableRange As Range
    Dim blockRange As Range
    Dim inputBg As Long
    
    titleRow = PID_DR_START_ROW
    hintRow = PID_DR_START_ROW + 1
    inputRow = PID_DR_START_ROW + 2
    headerRow = PID_DR_START_ROW + 3
    dataStartRow = headerRow + 1
    dataEndRow = dataStartRow + 3
    noteRow = headerRow + 5
    inputBg = RGB(255, 242, 204)
    
    PID_UnmergeDurchrechnungBlock ws
    
    PID_DRMigrateJaennerMusterInput ws
    PID_DRMigrateJaennerMusterFormula ws
    PID_DRRefreshBlockLabels ws, headerRow
    
    ws.Rows(titleRow).RowHeight = 28
    ws.Rows(hintRow).RowHeight = 48
    ws.Rows(inputRow).RowHeight = 26
    ws.Rows(headerRow).RowHeight = 30
    ws.Rows(noteRow).RowHeight = 40
    
    For dataRow = dataStartRow To dataEndRow
        ws.Rows(dataRow).RowHeight = 42
    Next dataRow
    
    Set blockRange = ws.Range("B" & titleRow & ":Q" & noteRow)
    Set tableRange = ws.Range("B" & headerRow & ":I" & dataEndRow)
    
    With ws.Range("B" & titleRow & ":Q" & titleRow)
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("B" & hintRow & ":Q" & hintRow)
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(89, 89, 89)
        .Font.Italic = True
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    With ws.Range("B" & inputRow & ":Q" & inputRow)
        .Interior.Color = inputBg
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ws.Cells(inputRow, 2).Font.Bold = True
    ws.Cells(inputRow, 6).Font.Bold = True
    
    With ws.Range(PID_DR_JAEN_VERF_CELL)
        .Interior.Color = inputBg
        .Font.Bold = True
        PID_DRApplyInputBorder .Borders
    End With
    
    With ws.Range(PID_DR_JAEN_MUST_CELL)
        .Interior.Color = inputBg
        .Font.Bold = True
        PID_DRApplyInputBorder .Borders
    End With
    
    With ws.Range("B" & headerRow & ":I" & headerRow)
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    
    ws.Range("J" & headerRow & ":Q" & headerRow).ClearContents
    ws.Cells(headerRow, 10).Value = "Status / Hinweis"
    
    With ws.Range("J" & headerRow & ":Q" & headerRow)
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    
    For dataRow = dataStartRow To dataEndRow
        If ((dataRow - dataStartRow) Mod 2) = 1 Then
            ws.Range("B" & dataRow & ":I" & dataRow).Interior.Color = RGB(248, 248, 248)
            ws.Range("J" & dataRow & ":Q" & dataRow).Interior.Color = RGB(248, 248, 248)
        End If
    Next dataRow
    
    ws.Range("B" & dataStartRow & ":C" & dataEndRow).HorizontalAlignment = xlLeft
    ws.Range("D" & dataStartRow & ":I" & dataEndRow).HorizontalAlignment = xlRight
    ws.Range("B" & dataStartRow & ":I" & dataEndRow).VerticalAlignment = xlCenter
    
    ws.Range("D" & dataStartRow & ":F" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range("G" & dataStartRow & ":G" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range("H" & dataStartRow & ":H" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range("I" & dataStartRow & ":I" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range(PID_DR_JAEN_VERF_CELL).NumberFormat = "#,##0.00"
    ws.Range(PID_DR_JAEN_MUST_CELL).NumberFormat = "#,##0.00"
    
    With ws.Range("B" & noteRow & ":Q" & noteRow)
        .Interior.Color = RGB(245, 245, 245)
        .Font.Color = RGB(89, 89, 89)
        .Font.Italic = True
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    PID_DRApplyOuterBorder blockRange
    
    PID_DRApplyTableBorders tableRange
    PID_DRApplyOuterBorder ws.Range("J" & headerRow & ":Q" & dataEndRow)
    
    Set diffRange = ws.Range("F" & dataStartRow & ":F" & dataEndRow)
    PID_DRClearFormatConditions diffRange
    PID_DRAddDiffFormatConditions diffRange
    
    Set statusRange = ws.Range("J" & dataStartRow & ":J" & dataEndRow)
    PID_DRClearFormatConditions statusRange
    PID_DRAddStatusFormatConditions statusRange
    
    Set ueberRange = ws.Range("H" & dataStartRow & ":I" & dataEndRow)
    PID_DRClearFormatConditions ueberRange
    PID_DRAddUeberFormatConditions ueberRange
    
    PID_DRMergeDisplayRows ws, headerRow, dataStartRow, dataEndRow
End Sub


Private Sub PID_DRRefreshBlockLabels(ByVal ws As Worksheet, ByVal headerRow As Long)
    ws.Cells(headerRow, 2).Value = "Zeitraum"
    ws.Cells(headerRow, 3).Value = "Endmonat"
    ws.Cells(headerRow, 4).Value = "Verfuegbar"
    ws.Cells(headerRow, 5).Value = "Muster"
    ws.Cells(headerRow, 6).Value = "Differenz"
    ws.Cells(headerRow, 7).Value = "Lohn/h"
    ws.Cells(headerRow, 8).Value = "Std"
    ws.Cells(headerRow, 9).Value = "EUR"
    ws.Cells(headerRow, 10).Value = "Status / Hinweis"
End Sub


Private Sub PID_DRMigrateJaennerMusterFormula(ByVal ws As Worksheet)
    Dim musterCell As Range
    Dim headerRow As Long
    Dim dataRow As Long
    Dim formulaText As String
    
    headerRow = PID_DR_START_ROW + 3
    dataRow = headerRow + 4
    Set musterCell = ws.Cells(dataRow, 5)
    
    On Error Resume Next
    formulaText = CStr(musterCell.Formula)
    If InStr(1, formulaText, "G30", vbTextCompare) > 0 Then
        musterCell.Formula = Replace(formulaText, "G30", "I30", , , vbTextCompare)
    End If
End Sub


Private Sub PID_DRAddUeberFormatConditions(ByVal ueberRange As Range)
    Dim firstCell As String
    
    firstCell = ueberRange.Cells(1, 1).Address(False, False)
    
    ueberRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & firstCell & ")," & firstCell & ">0)"
    With ueberRange.FormatConditions(ueberRange.FormatConditions.Count)
        .Interior.Color = RGB(252, 228, 214)
        .Font.Color = RGB(132, 46, 43)
        .Font.Bold = True
    End With
End Sub


Private Sub PID_DRApplyOuterBorder(ByVal target As Range)
    On Error Resume Next
    With target
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = RGB(31, 78, 121)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeRight).Color = RGB(31, 78, 121)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeTop).Color = RGB(31, 78, 121)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = RGB(31, 78, 121)
    End With
    On Error GoTo 0
End Sub


Private Sub PID_DRMigrateJaennerMusterInput(ByVal ws As Worksheet)
    On Error Resume Next
    
    If Len(Trim$(CStr(ws.Range(PID_DR_JAEN_MUST_CELL).Value2))) = 0 Then
        If Len(Trim$(CStr(ws.Range("G30").Value2))) > 0 Then
            ws.Range(PID_DR_JAEN_MUST_CELL).Value2 = ws.Range("G30").Value2
        End If
    End If
End Sub


Private Sub PID_DRMergeDisplayRows(ByVal ws As Worksheet, _
                                   ByVal headerRow As Long, _
                                   ByVal dataStartRow As Long, _
                                   ByVal dataEndRow As Long)
    Dim titleRow As Long
    Dim hintRow As Long
    Dim inputRow As Long
    Dim noteRow As Long
    Dim dataRow As Long
    
    titleRow = PID_DR_START_ROW
    hintRow = titleRow + 1
    inputRow = titleRow + 2
    noteRow = PID_DR_START_ROW + 8
    
    On Error Resume Next
    
    ws.Range("B" & titleRow & ":Q" & titleRow).Merge
    ws.Range("B" & hintRow & ":Q" & hintRow).Merge
    ws.Range("B" & inputRow & ":D" & inputRow).Merge
    ws.Range("F" & inputRow & ":H" & inputRow).Merge
    ws.Range("B" & noteRow & ":Q" & noteRow).Merge
    ws.Range("J" & headerRow & ":Q" & headerRow).Merge
    
    For dataRow = dataStartRow To dataEndRow
        ws.Range("J" & dataRow & ":Q" & dataRow).Merge
        With ws.Range("J" & dataRow)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    Next dataRow
    
    With ws.Range("J" & headerRow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    
    On Error GoTo 0
End Sub


Private Sub PID_DRApplyTableBorders(ByVal tableRange As Range)
    On Error Resume Next
    With tableRange
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).Color = RGB(180, 180, 180)
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideVertical).Color = RGB(180, 180, 180)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=RGB(31, 78, 121)
    End With
    On Error GoTo 0
End Sub


Private Sub PID_DRClearFormatConditions(ByVal target As Range)
    On Error Resume Next
    target.FormatConditions.Delete
    On Error GoTo 0
End Sub


Private Sub PID_DRAddDiffFormatConditions(ByVal diffRange As Range)
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
End Sub


Private Sub PID_DRAddStatusFormatConditions(ByVal statusRange As Range)
    Dim firstCell As String
    
    firstCell = statusRange.Cells(1, 1).Address(False, False)
    
    statusRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=ISNUMBER(SEARCH(""ACHTUNG""," & firstCell & "))"
    With statusRange.FormatConditions(statusRange.FormatConditions.Count)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With
    
    statusRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=OR(ISNUMBER(SEARCH(""Reserve""," & firstCell & ")),ISNUMBER(SEARCH(""Hinweis""," & firstCell & ")))"
    With statusRange.FormatConditions(statusRange.FormatConditions.Count)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
        .Font.Bold = True
    End With
    
    statusRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=" & firstCell & "=""OK"""
    With statusRange.FormatConditions(statusRange.FormatConditions.Count)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = True
    End With
End Sub


Private Sub PID_DRApplyInputBorder(ByVal borders As Borders)
    borders(xlEdgeLeft).LineStyle = xlContinuous
    borders(xlEdgeLeft).Weight = xlMedium
    borders(xlEdgeLeft).Color = RGB(255, 192, 0)
    borders(xlEdgeTop).LineStyle = xlContinuous
    borders(xlEdgeTop).Weight = xlMedium
    borders(xlEdgeTop).Color = RGB(255, 192, 0)
    borders(xlEdgeBottom).LineStyle = xlContinuous
    borders(xlEdgeBottom).Weight = xlMedium
    borders(xlEdgeBottom).Color = RGB(255, 192, 0)
    borders(xlEdgeRight).LineStyle = xlContinuous
    borders(xlEdgeRight).Weight = xlMedium
    borders(xlEdgeRight).Color = RGB(255, 192, 0)
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
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
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

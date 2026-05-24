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

Private Const PID_FU_TITLE_TOP_ROW As Long = 2
Private Const PID_FU_TITLE_BOTTOM_ROW As Long = 3
Private Const PID_FU_HEADER_TOP_ROW As Long = 4
Private Const PID_FU_HEADER_BOTTOM_ROW As Long = 6
Private Const PID_FU_DATA_START_ROW As Long = 7
Private Const PID_FU_DATA_END_ROW As Long = 22
Private Const PID_FU_TOTAL_ROW As Long = 23


' Anzeigetexte mit Umlauten (ChrW = ASCII-sichere Quelle, Win/Mac Excel 2016).
Private Function PID_DRTxtVerfuegbar() As String
    PID_DRTxtVerfuegbar = "Verf" & ChrW(252) & "gbar"
End Function

Private Function PID_DRTxtUeberstunden() As String
    PID_DRTxtUeberstunden = ChrW(220) & "berstunden"
End Function

Private Function PID_DRTxtMoeglich() As String
    PID_DRTxtMoeglich = "m" & ChrW(246) & "glich"
End Function

Private Function PID_DRTxtJaenner() As String
    PID_DRTxtJaenner = "J" & ChrW(228) & "nner"
End Function

Private Function PID_DRTxtJaenPlan() As String
    PID_DRTxtJaenPlan = "J" & ChrW(228) & "n (Plan)"
End Function

Private Function PID_DRTxtNaechstes() As String
    PID_DRTxtNaechstes = "n" & ChrW(228) & "chstes"
End Function

Private Function PID_DRTxtFebMaerApr() As String
    PID_DRTxtFebMaerApr = "Feb-M" & ChrW(228) & "r-Apr"
End Function

Private Function PID_DRTxtNovDezJaen() As String
    PID_DRTxtNovDezJaen = "Nov-Dez-J" & ChrW(228) & "n"
End Function

Private Function PID_DRGetStatusFormula(ByVal dataRow As Long) As String
    PID_DRGetStatusFormula = "=IF(F" & dataRow & "<0,""ACHTUNG: " & PID_DRTxtUeberstunden() & """,IF(F" & dataRow & ">0,""Reserve " & PID_DRTxtMoeglich() & """,""OK""))"
End Function

Private Function PID_DRTxtFinanzTitle() As String
    PID_DRTxtFinanzTitle = "FINANZIELL " & ChrW(220) & "BERSICHT "
End Function


Public Sub PID_BuildDurchrechnungUebersicht()
    PID_BuildDurchrechnungUebersichtInternal True
End Sub


Public Sub PID_RefreshDurchrechnungUebersicht()
    PID_RefreshDurchrechnungUebersichtInternal False
End Sub


Private Sub PID_RefreshDurchrechnungUebersichtInternal(ByVal lightRefresh As Boolean)
    Dim ws As Worksheet
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(PID_UBERSICHT_SHEET)
    
    If lightRefresh Then
        If PID_DurchrechnungBlockExists(ws) Then
            ws.Range("B" & PID_DR_START_ROW & ":Q" & PID_DR_END_ROW).Calculate
        End If
        GoTo CleanExit
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
    On Error GoTo CleanFail
    
    If PID_FinanzUebersichtBlockExists(ws) Then
        ws.Calculate
        PID_ApplyFinanzUebersichtFormats ws, False
    End If
    
    If Not PID_DurchrechnungBlockExists(ws) Then
        PID_BuildDurchrechnungUebersichtInternal False
        GoTo CleanProtect
    End If
    
    If Not PID_DurchrechnungBlockHasLohnColumn(ws) Then
        PID_BuildDurchrechnungUebersichtInternal False
        GoTo CleanProtect
    End If
    
    ws.Calculate
    PID_UpdateDurchrechnungLohnFormulas ws
    PID_UpdateDurchrechnungEuroFormulas ws
    PID_UnlockDurchrechnungInputs ws
    PID_ApplyDurchrechnungFormats ws
    
CleanProtect:
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
    
    savedJaenVerf = PID_DRSafeCellValue(ws.Range(PID_DR_JAEN_VERF_CELL))
    savedJaenMust = PID_DRSafeCellValue(ws.Range(PID_DR_JAEN_MUST_CELL))
    If Len(Trim$(CStr(savedJaenMust))) = 0 Then
        savedJaenMust = PID_DRSafeCellValue(ws.Range("G30"))
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
    
    If PID_FinanzUebersichtBlockExists(ws) Then
        PID_ApplyFinanzUebersichtFormats ws, True
    End If
    
    If PID_DurchrechnungBlockExists(ws) Then
        PID_UnlockDurchrechnungInputs ws
        PID_ApplyDurchrechnungFormats ws
    End If
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    
    If showMessage Then
        MsgBox "Durchrechnungsblock auf UEBERSICHT wurde erstellt." & vbCrLf & vbCrLf & _
               "Weisse Felder (manuell editierbar):" & vbCrLf & _
               "- " & PID_DRTxtJaenner() & " " & PID_DRTxtVerfuegbar() & " Plan (E30)" & vbCrLf & _
               "- " & PID_DRTxtJaenner() & " Muster Plan (I30)" & vbCrLf & vbCrLf & _
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
            
            If Not PID_CollectionHasKey(mergedAreas, areaKey) Then
                mergedAreas.Add areaKey, areaKey
                cell.MergeArea.UnMerge
            End If
        End If
    Next cell
End Sub


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
            "Differenz rot = zu wenig Stunden (" & PID_DRTxtUeberstunden() & "-Risiko), gelb = Reserve. " & _
            PID_DRTxtUeberstunden() & "-EUR nutzt den AVG Bruttolohn/h aus dem Schlussmonat (Monatsblatt Q42). " & _
            "Nur " & PID_DRTxtJaenner() & "-Plan (" & PID_DRTxtNaechstes() & " Jahr) sind die weissen Felder (E30, I30) manuell editierbar."
        .Cells(hintRow, 2).Font.Size = 9
        .Cells(hintRow, 2).WrapText = True
        
        .Cells(inputRow, 2).Value = PID_DRTxtJaenner() & " " & PID_DRTxtVerfuegbar() & " Plan (" & PID_DRTxtNaechstes() & " Jahr):"
        .Cells(inputRow, 6).Value = PID_DRTxtJaenner() & " Muster Plan (" & PID_DRTxtNaechstes() & " Jahr):"
        
        .Cells(headerRow, 2).Value = "Zeitraum"
        .Cells(headerRow, 3).Value = "Endmonat"
        .Cells(headerRow, 4).Value = PID_DRTxtVerfuegbar()
        .Cells(headerRow, 5).Value = "Muster"
        .Cells(headerRow, 6).Value = "Differenz"
        .Cells(headerRow, 7).Value = "Lohn/h"
        .Cells(headerRow, 8).Value = "Std"
        .Cells(headerRow, 9).Value = "EUR"
        .Cells(headerRow, 10).Value = "Status / Hinweis"
        .Range("K" & headerRow & ":Q" & headerRow).ClearContents
        .Rows(headerRow).Font.Bold = True
        
        dataRow = headerRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, PID_DRTxtFebMaerApr(), "April", _
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
        PID_WriteDurchrechnungDataRow ws, dataRow, PID_DRTxtNovDezJaen(), PID_DRTxtJaenPlan(), _
            "=November!Q12+Dezember!Q13+" & PID_DR_JAEN_VERF_CELL, _
            "=EINSTELLUNG!L16+EINSTELLUNG!L17+" & PID_DR_JAEN_MUST_CELL, _
            "=D" & dataRow & "-E" & dataRow, _
            "=Dezember!Q42"
        
        .Cells(noteRow, 2).Value = _
            PID_DRTxtUeberstunden() & " EUR = " & PID_DRTxtUeberstunden() & " Std x AVG Lohn/h x 1,5 (Spalte G aus Schlussmonat Q42). " & _
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
        .Cells(dataRow, 10).Formula = PID_DRGetStatusFormula(dataRow)
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
    If Not PID_DurchrechnungBlockExists(ws) Then Exit Sub
    
    ws.Range("B" & PID_DR_START_ROW & ":Q" & PID_DR_END_ROW).Locked = True
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
    inputBg = PID_StyleColorAccent()
    
    PID_UnmergeDurchrechnungBlock ws
    
    PID_DRMigrateJaennerMusterInput ws
    PID_DRMigrateJaennerMusterFormula ws
    PID_DRRefreshDisplayTexts ws, headerRow, dataStartRow, dataEndRow
    
    ws.Rows(titleRow).RowHeight = PID_STYLE_TITLE_ROW_HEIGHT
    ws.Rows(hintRow).RowHeight = 48
    ws.Rows(inputRow).RowHeight = 40
    ws.Rows(headerRow).RowHeight = PID_STYLE_HEADER_BOTTOM_ROW_HEIGHT
    ws.Rows(noteRow).RowHeight = 40
    
    For dataRow = dataStartRow To dataEndRow
        ws.Rows(dataRow).RowHeight = 42
    Next dataRow
    
    Set blockRange = ws.Range("B" & titleRow & ":Q" & noteRow)
    Set tableRange = ws.Range("B" & headerRow & ":I" & dataEndRow)
    
    PID_StyleApplyTitleBand ws.Range("B" & titleRow & ":Q" & titleRow)
    
    With ws.Range("B" & hintRow & ":Q" & hintRow)
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(89, 89, 89)
        .Font.Italic = True
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ws.Cells(inputRow, 2).Font.Bold = True
    ws.Cells(inputRow, 6).Font.Bold = True
    
    PID_DRCleanInputRow ws, inputRow
    
    PID_StyleApplyHeaderBand ws.Range("B" & headerRow & ":I" & headerRow)
    ws.Range("B" & headerRow & ":I" & headerRow).WrapText = False
    
    ws.Range("J" & headerRow & ":Q" & headerRow).ClearContents
    ws.Cells(headerRow, 10).Value = "Status / Hinweis"
    
    PID_StyleApplyHeaderBand ws.Range("J" & headerRow & ":Q" & headerRow)
    ws.Range("J" & headerRow & ":Q" & headerRow).WrapText = False
    
    For dataRow = dataStartRow To dataEndRow
        If ((dataRow - dataStartRow) Mod 2) = 1 Then
            ws.Range("B" & dataRow & ":I" & dataRow).Interior.Color = PID_StyleColorZebra()
            ws.Range("J" & dataRow & ":Q" & dataRow).Interior.Color = PID_StyleColorZebra()
        End If
    Next dataRow
    
    ws.Range("B" & dataStartRow & ":I" & dataEndRow).HorizontalAlignment = xlCenter
    ws.Range("B" & dataStartRow & ":I" & dataEndRow).VerticalAlignment = xlCenter
    
    ws.Range("D" & dataStartRow & ":F" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range("H" & dataStartRow & ":H" & dataEndRow).NumberFormat = "#,##0.00"
    ws.Range(PID_DR_JAEN_VERF_CELL).NumberFormat = "#,##0.00"
    ws.Range(PID_DR_JAEN_MUST_CELL).NumberFormat = "#,##0.00"
    PID_ApplyEuroNumberFormat ws.Range("G" & dataStartRow & ":G" & dataEndRow)
    PID_ApplyEuroNumberFormat ws.Range("I" & dataStartRow & ":I" & dataEndRow)
    
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
    
    PID_DRMergeDisplayRows ws, headerRow, dataStartRow, dataEndRow
    PID_DRFormatInputRow ws, inputRow, inputBg
    
    Set diffRange = ws.Range("F" & dataStartRow & ":F" & dataEndRow)
    PID_DRClearFormatConditions diffRange
    PID_DRAddDiffFormatConditions diffRange
    
    Set statusRange = ws.Range("J" & dataStartRow & ":J" & dataEndRow)
    PID_DRClearFormatConditions statusRange
    PID_DRAddStatusFormatConditions statusRange
    
    Set ueberRange = ws.Range("H" & dataStartRow & ":I" & dataEndRow)
    PID_DRClearFormatConditions ueberRange
    PID_DRAddUeberFormatConditions ueberRange
End Sub


Private Sub PID_DRRefreshDisplayTexts(ByVal ws As Worksheet, _
                                      ByVal headerRow As Long, _
                                      ByVal dataStartRow As Long, _
                                      ByVal dataEndRow As Long)
    Dim hintRow As Long
    Dim inputRow As Long
    Dim noteRow As Long
    Dim dataRow As Long
    
    hintRow = PID_DR_START_ROW + 1
    inputRow = PID_DR_START_ROW + 2
    noteRow = headerRow + 5
    
    ws.Cells(headerRow, 2).Value = "Zeitraum"
    ws.Cells(headerRow, 3).Value = "Endmonat"
    ws.Cells(headerRow, 4).Value = PID_DRTxtVerfuegbar()
    ws.Cells(headerRow, 5).Value = "Muster"
    ws.Cells(headerRow, 6).Value = "Differenz"
    ws.Cells(headerRow, 7).Value = "Lohn/h"
    ws.Cells(headerRow, 8).Value = "Std"
    ws.Cells(headerRow, 9).Value = "EUR"
    ws.Cells(headerRow, 10).Value = "Status / Hinweis"
    
    ws.Cells(hintRow, 2).Value = _
        "Jede Zeile = 3 Monate bis zum Schlussmonat. " & _
        "Differenz rot = zu wenig Stunden (" & PID_DRTxtUeberstunden() & "-Risiko), gelb = Reserve. " & _
        PID_DRTxtUeberstunden() & "-EUR nutzt den AVG Bruttolohn/h aus dem Schlussmonat (Monatsblatt Q42). " & _
        "Nur " & PID_DRTxtJaenner() & "-Plan (" & PID_DRTxtNaechstes() & " Jahr) sind die weissen Felder (E30, I30) manuell editierbar."
    
    ws.Cells(inputRow, 2).Value = PID_DRTxtJaenner() & " " & PID_DRTxtVerfuegbar() & " Plan (" & PID_DRTxtNaechstes() & " Jahr):"
    ws.Cells(inputRow, 6).Value = PID_DRTxtJaenner() & " Muster Plan (" & PID_DRTxtNaechstes() & " Jahr):"
    
    ws.Cells(noteRow, 2).Value = _
        PID_DRTxtUeberstunden() & " EUR = " & PID_DRTxtUeberstunden() & " Std x AVG Lohn/h x 1,5 (Spalte G aus Schlussmonat Q42). " & _
        "Nur bei negativem Schluss (rote Differenz). Positive Differenz = Reserve ohne EUR-Kosten."
    
    ws.Cells(dataStartRow, 2).Value = PID_DRTxtFebMaerApr()
    ws.Cells(dataStartRow + 3, 2).Value = PID_DRTxtNovDezJaen()
    ws.Cells(dataStartRow + 3, 3).Value = PID_DRTxtJaenPlan()
    
    For dataRow = dataStartRow To dataEndRow
        ws.Cells(dataRow, 10).Formula = PID_DRGetStatusFormula(dataRow)
    Next dataRow
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
    
    On Error Resume Next
    
    ueberRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & firstCell & ")," & firstCell & ">0)"
    With ueberRange.FormatConditions(ueberRange.FormatConditions.Count)
        .Interior.Color = RGB(252, 228, 214)
        .Font.Color = RGB(132, 46, 43)
        .Font.Bold = True
    End With
    
    On Error GoTo 0
End Sub


Private Sub PID_DRApplyOuterBorder(ByVal target As Range)
    PID_StyleApplyOuterBorder target
End Sub


Private Function PID_DRSafeCellValue(ByVal target As Range) As Variant
    On Error Resume Next
    PID_DRSafeCellValue = target.Value2
    If Err.Number <> 0 Then
        Err.Clear
        PID_DRSafeCellValue = Empty
    End If
    On Error GoTo 0
End Function


Private Sub PID_DRMigrateJaennerMusterInput(ByVal ws As Worksheet)
    On Error Resume Next
    
    If Len(Trim$(CStr(ws.Range(PID_DR_JAEN_MUST_CELL).Value2))) = 0 Then
        If Len(Trim$(CStr(ws.Range("G30").Value2))) > 0 Then
            ws.Range(PID_DR_JAEN_MUST_CELL).Value2 = ws.Range("G30").Value2
            ws.Range("G30").ClearContents
        End If
    End If
    
    On Error Resume Next
    ws.Range("G30").ClearContents
    With ws.Range("G30").Borders
        .LineStyle = xlNone
    End With
    On Error GoTo 0
End Sub


Private Sub PID_DRCleanInputRow(ByVal ws As Worksheet, ByVal inputRow As Long)
    Dim labelVerf As String
    Dim labelMust As String
    Dim valVerf As Variant
    Dim valMust As Variant
    
    labelVerf = CStr(ws.Cells(inputRow, 2).Value)
    labelMust = CStr(ws.Cells(inputRow, 6).Value)
    valVerf = ws.Range(PID_DR_JAEN_VERF_CELL).Value2
    valMust = ws.Range(PID_DR_JAEN_MUST_CELL).Value2
    
    ws.Range("D" & inputRow & ",G" & inputRow & ",H" & inputRow & ",J" & inputRow & ":Q" & inputRow).ClearContents
    
    ws.Cells(inputRow, 2).Value = labelVerf
    ws.Cells(inputRow, 6).Value = labelMust
    
    If Len(Trim$(CStr(valVerf))) > 0 Then ws.Range(PID_DR_JAEN_VERF_CELL).Value2 = valVerf
    If Len(Trim$(CStr(valMust))) > 0 Then ws.Range(PID_DR_JAEN_MUST_CELL).Value2 = valMust
End Sub


Private Sub PID_DRFormatInputRow(ByVal ws As Worksheet, ByVal inputRow As Long, ByVal inputBg As Long)
    On Error Resume Next
    
    With ws.Range("B" & inputRow & ":Q" & inputRow)
        .VerticalAlignment = xlCenter
    End With
    
    PID_StyleApplyInputGuideLabel ws.Range("B" & inputRow & ":D" & inputRow)
    With ws.Range("B" & inputRow & ":D" & inputRow)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    
    PID_StyleApplyInputGuideLabel ws.Range("F" & inputRow & ":H" & inputRow)
    With ws.Range("F" & inputRow & ":H" & inputRow)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    
    PID_StyleApplyInputCell ws.Range(PID_DR_JAEN_VERF_CELL)
    With ws.Range(PID_DR_JAEN_VERF_CELL)
        .HorizontalAlignment = xlCenter
        .Font.Bold = False
        PID_DRApplyInputBorder .Borders
    End With
    
    PID_StyleApplyInputCell ws.Range(PID_DR_JAEN_MUST_CELL)
    With ws.Range(PID_DR_JAEN_MUST_CELL)
        .HorizontalAlignment = xlCenter
        .Font.Bold = False
        PID_DRApplyInputBorder .Borders
    End With
    
    PID_StyleApplyInputGuideLabel ws.Range("J" & inputRow & ":Q" & inputRow)
    With ws.Range("J" & inputRow & ":Q" & inputRow)
        .HorizontalAlignment = xlCenter
        .WrapText = False
    End With
    
    On Error GoTo 0
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
    ws.Range("J" & inputRow & ":Q" & inputRow).Merge
    ws.Range("B" & noteRow & ":Q" & noteRow).Merge
    ws.Range("J" & headerRow & ":Q" & headerRow).Merge
    
    For dataRow = dataStartRow To dataEndRow
        ws.Range("J" & dataRow & ":Q" & dataRow).Merge
        With ws.Range("J" & dataRow)
            .WrapText = True
            .HorizontalAlignment = xlCenter
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
    PID_StyleApplyTableBorders tableRange
End Sub


Private Sub PID_DRClearFormatConditions(ByVal target As Range)
    On Error Resume Next
    target.FormatConditions.Delete
    On Error GoTo 0
End Sub


Private Sub PID_DRAddDiffFormatConditions(ByVal diffRange As Range)
    Dim firstCell As String
    
    firstCell = diffRange.Cells(1, 1).Address(False, False)
    
    On Error Resume Next
    
    diffRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & firstCell & ")," & firstCell & "<0)"
    With diffRange.FormatConditions(diffRange.FormatConditions.Count)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With
    
    diffRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & firstCell & ")," & firstCell & ">0)"
    With diffRange.FormatConditions(diffRange.FormatConditions.Count)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
        .Font.Bold = True
    End With
    
    diffRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & firstCell & ")," & firstCell & "=0)"
    With diffRange.FormatConditions(diffRange.FormatConditions.Count)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = True
    End With
    
    On Error GoTo 0
End Sub


Private Sub PID_DRAddStatusFormatConditions(ByVal statusRange As Range)
    Dim firstCell As String
    
    firstCell = statusRange.Cells(1, 1).Address(False, False)
    
    On Error Resume Next
    
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
    
    On Error GoTo 0
End Sub


Private Sub PID_DRApplyInputBorder(ByVal borders As Borders)
    On Error Resume Next
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
    On Error GoTo 0
End Sub


Public Sub PID_FormatFinanzUebersicht()
    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(PID_UBERSICHT_SHEET)
    
    If Not PID_FinanzUebersichtBlockExists(ws) Then
        MsgBox "Kein FINANZIELL-Block auf UEBERSICHT gefunden.", vbExclamation, "UEBERSICHT"
        GoTo CleanExit
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
    On Error GoTo CleanFail
    
    PID_ApplyFinanzUebersichtFormats ws, True
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    
    GoTo CleanExit

CleanFail:
    MsgBox "Fehler bei FormatFinanzUebersicht:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "UEBERSICHT"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
End Sub


Private Function PID_FinanzUebersichtBlockExists(ByVal ws As Worksheet) As Boolean
    Dim headerText As String
    
    On Error GoTo SafeExit
    
    headerText = Trim$(CStr(ws.Cells(PID_FU_HEADER_TOP_ROW, 3).Text))
    
    If InStr(1, headerText, "SALES", vbTextCompare) > 0 Then
        PID_FinanzUebersichtBlockExists = True
        Exit Function
    End If
    
    If Trim$(CStr(ws.Cells(PID_FU_HEADER_BOTTOM_ROW, 3).Value)) = "BUDGET" Then
        PID_FinanzUebersichtBlockExists = True
        Exit Function
    End If

SafeExit:
    PID_FinanzUebersichtBlockExists = False
End Function


Private Sub PID_FixFinanzUebersichtFormulas(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range("B2").Formula = "= """ & PID_DRTxtFinanzTitle() & """ & EINSTELLUNG!C35"
    ws.Range("B23").Formula = "= ""GESAMT "" & EINSTELLUNG!C35"
    On Error GoTo 0
End Sub


Private Sub PID_ApplyFinanzUebersichtFormats(ByVal ws As Worksheet, Optional ByVal syncFinanzValues As Boolean = False)
    Dim dataRow As Long
    Dim qRows As Variant
    Dim i As Long
    Dim tableRange As Range
    
    qRows = Array(10, 14, 18, 22)
    
    PID_FixFinanzUebersichtFormulas ws
    
    If syncFinanzValues Then
        PID_SyncFinanzSummaryToUbersicht
    End If
    
    ws.Rows(PID_FU_TITLE_TOP_ROW).RowHeight = PID_STYLE_TITLE_ROW_HEIGHT
    ws.Rows(PID_FU_TITLE_BOTTOM_ROW).RowHeight = PID_STYLE_TITLE_SUB_ROW_HEIGHT
    ws.Rows(PID_FU_HEADER_TOP_ROW).RowHeight = PID_STYLE_HEADER_TOP_ROW_HEIGHT
    ws.Rows(PID_FU_HEADER_BOTTOM_ROW).RowHeight = PID_STYLE_HEADER_BOTTOM_ROW_HEIGHT
    
    For dataRow = PID_FU_DATA_START_ROW To PID_FU_DATA_END_ROW
        ws.Rows(dataRow).RowHeight = PID_STYLE_DATA_ROW_HEIGHT
    Next dataRow
    ws.Rows(PID_FU_TOTAL_ROW).RowHeight = PID_STYLE_TOTAL_ROW_HEIGHT
    
    PID_StyleApplyTitleBand ws.Range("B" & PID_FU_TITLE_TOP_ROW & ":Q" & PID_FU_TITLE_BOTTOM_ROW)
    PID_StyleApplyHeaderBand ws.Range("B" & PID_FU_HEADER_TOP_ROW & ":Q" & PID_FU_HEADER_BOTTOM_ROW)
    
    ws.Range("B" & PID_FU_DATA_START_ROW & ":Q" & PID_FU_DATA_END_ROW).Interior.Color = vbWhite
    ws.Range("B" & PID_FU_DATA_START_ROW & ":Q" & PID_FU_DATA_END_ROW).Font.Color = vbBlack
    ws.Range("B" & PID_FU_DATA_START_ROW & ":Q" & PID_FU_DATA_END_ROW).Font.Bold = False
    
    For dataRow = PID_FU_DATA_START_ROW To PID_FU_DATA_END_ROW
        If ((dataRow - PID_FU_DATA_START_ROW) Mod 2) = 1 Then
            ws.Range("B" & dataRow & ":Q" & dataRow).Interior.Color = PID_StyleColorZebra()
        End If
    Next dataRow
    
    For i = LBound(qRows) To UBound(qRows)
        dataRow = CLng(qRows(i))
        PID_StyleApplyAccentSummaryBand ws.Range("B" & dataRow & ":Q" & dataRow)
    Next i
    
    PID_StyleApplyAccentSummaryBand ws.Range("B" & PID_FU_TOTAL_ROW & ":Q" & PID_FU_TOTAL_ROW)
    
    ws.Range("B" & PID_FU_DATA_START_ROW & ":Q" & PID_FU_TOTAL_ROW).HorizontalAlignment = xlCenter
    ws.Range("B" & PID_FU_DATA_START_ROW & ":Q" & PID_FU_TOTAL_ROW).VerticalAlignment = xlCenter
    
    PID_ApplyEuroNumberFormat ws.Range("C" & PID_FU_DATA_START_ROW & ":H" & PID_FU_TOTAL_ROW)
    ws.Range("I" & PID_FU_DATA_START_ROW & ":K" & PID_FU_TOTAL_ROW).NumberFormat = "0.00%"
    ws.Range("L" & PID_FU_DATA_START_ROW & ":M" & PID_FU_TOTAL_ROW).NumberFormat = "#,##0.00"
    ws.Range("N" & PID_FU_DATA_START_ROW & ":P" & PID_FU_TOTAL_ROW).NumberFormat = "#,##0.00"
    ws.Range("Q" & PID_FU_DATA_START_ROW & ":Q" & PID_FU_TOTAL_ROW).NumberFormat = "0.00%"
    
    Set tableRange = ws.Range("B" & PID_FU_HEADER_TOP_ROW & ":Q" & PID_FU_TOTAL_ROW)
    PID_DRApplyOuterBorder tableRange
    PID_DRApplyTableBorders tableRange
    
    PID_FUApplyDiffFormatConditions ws.Range("E" & PID_FU_DATA_START_ROW & ":E" & PID_FU_TOTAL_ROW)
    PID_FUApplyDiffFormatConditions ws.Range("H" & PID_FU_DATA_START_ROW & ":H" & PID_FU_TOTAL_ROW)
    PID_FUApplyDiffFormatConditions ws.Range("K" & PID_FU_DATA_START_ROW & ":K" & PID_FU_TOTAL_ROW)
    PID_FUApplyDiffFormatConditions ws.Range("P" & PID_FU_DATA_START_ROW & ":P" & PID_FU_TOTAL_ROW)
End Sub


Private Sub PID_FUApplyDiffFormatConditions(ByVal diffRange As Range)
    PID_DRClearFormatConditions diffRange
    PID_DRAddDiffFormatConditions diffRange
End Sub


Public Sub PID_FormatDurchrechnungUebersicht()
    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean
    
    On Error GoTo CleanFail
    
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Worksheets(PID_UBERSICHT_SHEET)
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
    On Error GoTo CleanFail
    
    If PID_FinanzUebersichtBlockExists(ws) Then
        PID_ApplyFinanzUebersichtFormats ws, True
    End If
    
    If Not PID_DurchrechnungBlockExists(ws) Then
        MsgBox "Kein Durchrechnungsblock auf UEBERSICHT gefunden. Bitte zuerst PID_BuildDurchrechnungUebersicht ausfuehren.", _
               vbExclamation, "Durchrechnung"
        GoTo CleanProtect
    End If
    
    PID_UnlockDurchrechnungInputs ws
    PID_ApplyDurchrechnungFormats ws
    
CleanProtect:
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

Attribute VB_Name = "mod_BuildDurchrechnungUebersicht"
Option Explicit

Private Const PID_DR_START_ROW As Long = 28
Private Const PID_DR_END_ROW As Long = 38
Private Const PID_DR_FIRST_COL As Long = 2   ' B
Private Const PID_DR_LAST_COL As Long = 9    ' I

Private Const PID_DR_LOHN_CELL As String = "C30"
Private Const PID_DR_JAEN_VERF_CELL As String = "E30"
Private Const PID_DR_JAEN_MUST_CELL As String = "G30"


Public Sub PID_BuildDurchrechnungUebersicht()
    Dim ws As Worksheet
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim savedJaenVerf As Variant
    Dim savedJaenMust As Variant
    Dim savedLohn As Variant
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ws = ThisWorkbook.Worksheets("UBERSICHT")
    
    savedLohn = ws.Range(PID_DR_LOHN_CELL).Value2
    savedJaenVerf = ws.Range(PID_DR_JAEN_VERF_CELL).Value2
    savedJaenMust = ws.Range(PID_DR_JAEN_MUST_CELL).Value2
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    PID_ClearDurchrechnungBlock ws
    PID_WriteDurchrechnungBlock ws
    
    If Len(Trim$(CStr(savedLohn))) > 0 Then ws.Range(PID_DR_LOHN_CELL).Value2 = savedLohn
    If Len(Trim$(CStr(savedJaenVerf))) > 0 Then ws.Range(PID_DR_JAEN_VERF_CELL).Value2 = savedJaenVerf
    If Len(Trim$(CStr(savedJaenMust))) > 0 Then ws.Range(PID_DR_JAEN_MUST_CELL).Value2 = savedJaenMust
    
    PID_UnlockDurchrechnungInputs ws
    PID_ApplyDurchrechnungFormats ws
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    
    MsgBox "Durchrechnungsblock auf UEBERSICHT wurde erstellt." & vbCrLf & vbCrLf & _
           "Gelbe Felder:" & vbCrLf & _
           "- Stundenlohn (C30)" & vbCrLf & _
           "- Jaenner Verfuegbar Plan (E30)" & vbCrLf & _
           "- Jaenner Muster Plan (G30)", _
           vbInformation, "Durchrechnung"
    
    GoTo CleanExit

CleanFail:
    MsgBox "Fehler bei PID_BuildDurchrechnungUebersicht:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Durchrechnung"

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Private Sub PID_ClearDurchrechnungBlock(ByVal ws As Worksheet)
    Dim target As Range
    
    Set target = ws.Range(ws.Cells(PID_DR_START_ROW, PID_DR_FIRST_COL), _
                          ws.Cells(PID_DR_END_ROW, PID_DR_LAST_COL))
    
    target.ClearContents
    target.Interior.Pattern = xlNone
    target.Font.Bold = False
    target.Font.Color = vbBlack
    target.HorizontalAlignment = xlGeneral
    target.VerticalAlignment = xlCenter
    target.WrapText = False
    
    target.FormatConditions.Delete
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
        .Range("B" & titleRow & ":I" & titleRow).Merge
        .Cells(titleRow, 2).Formula = "= ""DURCHRECHNUNGSSTUNDEN - "" & EINSTELLUNG!C35"
        .Cells(titleRow, 2).Font.Bold = True
        .Cells(titleRow, 2).Font.Size = 12
        
        .Range("B" & hintRow & ":I" & hintRow).Merge
        .Cells(hintRow, 2).Value = "Schlussmonat ist entscheidend. Jaenner-Plan = naechstes Kalenderjahr (EINSTELLUNG C35 + 1). Feiertags-Ausgleich wird nicht extra berechnet."
        .Cells(hintRow, 2).Font.Size = 9
        .Cells(hintRow, 2).WrapText = True
        
        .Cells(inputRow, 2).Value = "Stundenlohn (EUR/h):"
        .Cells(inputRow, 4).Value = "Jaenner Verfuegbar Plan:"
        .Cells(inputRow, 6).Value = "Jaenner Muster Plan:"
        
        .Cells(headerRow, 2).Value = "Zeitraum"
        .Cells(headerRow, 3).Value = "Schlussmonat"
        .Cells(headerRow, 4).Value = "Verfuegbar Summe"
        .Cells(headerRow, 5).Value = "Muster Summe"
        .Cells(headerRow, 6).Value = "Differenz"
        .Cells(headerRow, 7).Value = "Ueberstunden Std"
        .Cells(headerRow, 8).Value = "Ueberstunden EUR"
        .Cells(headerRow, 9).Value = "Status"
        .Rows(headerRow).Font.Bold = True
        
        dataRow = headerRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, "Feb-Maer-Apr", "April", _
            "=Februar!Q12+Marz!Q13+April!Q13", _
            "=EINSTELLUNG!L7+EINSTELLUNG!L8+EINSTELLUNG!L9", _
            "=April!Q15"
        
        dataRow = dataRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, "Mai-Jun-Jul", "Juli", _
            "=Mai!Q12+Juni!Q13+Juli!Q13", _
            "=EINSTELLUNG!L10+EINSTELLUNG!L11+EINSTELLUNG!L12", _
            "=Juli!Q15"
        
        dataRow = dataRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, "Aug-Sep-Okt", "Oktober", _
            "=August!Q12+September!Q13+Oktober!Q13", _
            "=EINSTELLUNG!L13+EINSTELLUNG!L14+EINSTELLUNG!L15", _
            "=Oktober!Q15"
        
        dataRow = dataRow + 1
        PID_WriteDurchrechnungDataRow ws, dataRow, "Nov-Dez-Jaen", "Jaen (Plan)", _
            "=November!Q12+Dezember!Q13+" & PID_DR_JAEN_VERF_CELL, _
            "=EINSTELLUNG!L16+EINSTELLUNG!L17+" & PID_DR_JAEN_MUST_CELL, _
            "=D" & dataRow & "-E" & dataRow
        
        .Range("B" & noteRow & ":I" & noteRow).Merge
        .Cells(noteRow, 2).Value = "Ueberstunden EUR = Ueberstunden Std x Stundenlohn x 1,5. Nur bei negativem Schluss (Muster hoeher als Verfuegbar). Positive Differenz = Reserve, kein EUR-Risiko."
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
                                          ByVal diffFormula As String)
    With ws
        .Cells(dataRow, 2).Value = periodLabel
        .Cells(dataRow, 3).Value = closingMonthLabel
        .Cells(dataRow, 4).Formula = verfFormula
        .Cells(dataRow, 5).Formula = musterFormula
        .Cells(dataRow, 6).Formula = diffFormula
        .Cells(dataRow, 7).Formula = "=MAX(0,-F" & dataRow & ")"
        .Cells(dataRow, 8).Formula = "=IF($C$30="""","""",G" & dataRow & "*$C$30*1.5)"
        .Cells(dataRow, 9).Formula = "=IF(F" & dataRow & "<0,""ACHTUNG: Ueberstunden-Risiko"",IF(F" & dataRow & ">0,""Hinweis: Reserve / Minus-Stunden moeglich"",""OK""))"
    End With
End Sub


Private Sub PID_UnlockDurchrechnungInputs(ByVal ws As Worksheet)
    ws.Range(PID_DR_LOHN_CELL).Locked = False
    ws.Range(PID_DR_JAEN_VERF_CELL).Locked = False
    ws.Range(PID_DR_JAEN_MUST_CELL).Locked = False
End Sub


Private Sub PID_ApplyDurchrechnungFormats(ByVal ws As Worksheet)
    Dim headerRow As Long
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim inputRow As Long
    Dim diffRange As Range
    Dim yellowColor As Long
    
    headerRow = PID_DR_START_ROW + 3
    inputRow = PID_DR_START_ROW + 2
    dataStartRow = headerRow + 1
    dataEndRow = dataStartRow + 3
    yellowColor = RGB(255, 242, 204)
    
    ws.Rows(PID_DR_START_ROW).RowHeight = 18
    ws.Rows(PID_DR_START_ROW + 1).RowHeight = 30
    ws.Rows(inputRow).RowHeight = 18
    ws.Rows(headerRow).RowHeight = 18
    ws.Rows(PID_DR_END_ROW).RowHeight = 36
    
    ws.Range("B" & headerRow & ":I" & headerRow).HorizontalAlignment = xlCenter
    ws.Range("B" & dataStartRow & ":C" & dataEndRow).HorizontalAlignment = xlLeft
    ws.Range("D" & dataStartRow & ":H" & dataEndRow).HorizontalAlignment = xlRight
    ws.Range("I" & dataStartRow & ":I" & dataEndRow).HorizontalAlignment = xlLeft
    ws.Range("I" & dataStartRow & ":I" & dataEndRow).WrapText = True
    
    ws.Range(PID_DR_LOHN_CELL).Interior.Color = yellowColor
    ws.Range(PID_DR_JAEN_VERF_CELL).Interior.Color = yellowColor
    ws.Range(PID_DR_JAEN_MUST_CELL).Interior.Color = yellowColor
    
    ws.Range("D" & dataStartRow & ":H" & dataEndRow).NumberFormat = "0.00"
    ws.Range(PID_DR_LOHN_CELL).NumberFormat = "0.00"
    ws.Range(PID_DR_JAEN_VERF_CELL).NumberFormat = "0.00"
    ws.Range(PID_DR_JAEN_MUST_CELL).NumberFormat = "0.00"
    
    With ws.Range("B" & headerRow & ":I" & dataEndRow).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range("B" & headerRow & ":I" & dataEndRow).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range("B" & headerRow & ":I" & dataEndRow).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range("B" & headerRow & ":I" & dataEndRow).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range("B" & headerRow & ":I" & dataEndRow).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range("B" & headerRow & ":I" & dataEndRow).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Set diffRange = ws.Range("F" & dataStartRow & ":F" & dataEndRow)
    diffRange.FormatConditions.Delete
    
    diffRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    diffRange.FormatConditions(diffRange.FormatConditions.Count).Interior.Color = RGB(255, 199, 206)
    
    diffRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    diffRange.FormatConditions(diffRange.FormatConditions.Count).Interior.Color = RGB(255, 235, 156)
    
    diffRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
    diffRange.FormatConditions(diffRange.FormatConditions.Count).Interior.Color = RGB(198, 239, 206)
End Sub

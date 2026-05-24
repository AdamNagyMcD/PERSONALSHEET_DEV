Attribute VB_Name = "mod_BuildFluktuationAnalyse"
Option Explicit

Public Sub BuildFluktuationAnalyse()
    Dim dataWs As Worksheet
    Dim analyseWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long
    
    Dim totalExits As Long
    Dim neutralExits As Long
    Dim earlyExits As Long
    Dim experiencedLoss As Long
    Dim importantExits As Long
    Dim incompleteExits As Long
    
    Dim totalLoss As Double
    Dim avgLoss As Double
    Dim riskLevel As String
    Dim managerSummary As String
    Dim recommendationItems As Variant
    Dim criticalExits As Long
    Dim normalExits As Long
    Dim statusRow As Long
    Dim kpiLabelRow As Long
    Dim kpiValueRow As Long
    Dim alertsHeaderRow As Long
    Dim alertsEndRow As Long
    Dim recHeaderRow As Long
    Dim recEndRow As Long
    Dim chartRow As Long
    Dim currentYear As Long
    
    Dim monthNames As Variant
    Dim monthExit(1 To 12) As Long
    Dim monthNeutral(1 To 12) As Long
    Dim monthEarly(1 To 12) As Long
    Dim monthExperienced(1 To 12) As Long
    Dim monthImportant(1 To 12) As Long
    Dim monthIncomplete(1 To 12) As Long
    Dim monthLoss(1 To 12) As Double
    Dim monthPersonalEnde(1 To 12) As Long
    Dim monthFluctuation(1 To 12) As Double
    
    Dim showEarly As Boolean
    Dim showExperienced As Boolean
    Dim showImportant As Boolean
    Dim showNeutral As Boolean
    Dim showIncomplete As Boolean
    
    Dim monthName As String
    Dim monthIndex As Long
    Dim categoryText As String
    Dim lossValue As Double
    Dim avgMonthLoss As Double
    Dim reasonSummary As String
    Dim monthHint As String
    
    Dim currentCol As Long
    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim lastTableCol As Long
    Dim monthlyTitleRow As Long
    Dim explanationStartRow As Long
    Dim outputRow As Long
    
    Dim ytdFluctuation As Double
    Dim quarterFluctuation(1 To 4) As Double
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    monthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    Set dataWs = ThisWorkbook.Worksheets("FLUKTUATION_DATEN")
    Set analyseWs = ThisWorkbook.Worksheets(PID_FLUKTUATION_SHEET)
    currentYear = PID_GetWorkbookYear()
    
    On Error Resume Next
    analyseWs.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    With analyseWs
        .Cells.UnMerge
        .Cells.Clear
        
        .Cells.Font.Color = vbBlack
        .Cells.Interior.Pattern = xlNone
        .Cells.HorizontalAlignment = xlGeneral
        .Cells.VerticalAlignment = xlCenter
        .Cells.WrapText = False
        .Rows.RowHeight = 15
    End With
    
    ' --- Abschnitt 1: Daten aus FLUKTUATION_DATEN aggregieren ---
    lastRow = dataWs.Cells(dataWs.Rows.count, "A").End(xlUp).Row
    
    If lastRow >= 2 Then
        For r = 2 To lastRow
            totalExits = totalExits + 1
            
            monthName = Trim$(CStr(dataWs.Cells(r, "A").Value))
            monthIndex = PID_GetMonthIndexFromName(monthName)
            
            If IsNumeric(dataWs.Cells(r, "J").Value) Then
                lossValue = CDbl(dataWs.Cells(r, "J").Value)
            Else
                lossValue = 0
            End If
            
            categoryText = Trim$(CStr(dataWs.Cells(r, "K").Value))
            totalLoss = totalLoss + lossValue
            
            If monthIndex >= 1 And monthIndex <= 12 Then
                monthExit(monthIndex) = monthExit(monthIndex) + 1
                monthLoss(monthIndex) = monthLoss(monthIndex) + lossValue
            End If
            
            Select Case categoryText
                Case "Neutrale Bewegung"
                    neutralExits = neutralExits + 1
                    If monthIndex >= 1 And monthIndex <= 12 Then monthNeutral(monthIndex) = monthNeutral(monthIndex) + 1
                
                Case "Austritt in den ersten 90 Tagen"
                    earlyExits = earlyExits + 1
                    If monthIndex >= 1 And monthIndex <= 12 Then monthEarly(monthIndex) = monthEarly(monthIndex) + 1
                
                Case "Verlust erfahrener Mitarbeiter"
                    experiencedLoss = experiencedLoss + 1
                    If monthIndex >= 1 And monthIndex <= 12 Then monthExperienced(monthIndex) = monthExperienced(monthIndex) + 1
                
                Case "Wichtiger Austritt"
                    importantExits = importantExits + 1
                    If monthIndex >= 1 And monthIndex <= 12 Then monthImportant(monthIndex) = monthImportant(monthIndex) + 1
                
                Case "Austrittsgrund fehlt"
                    incompleteExits = incompleteExits + 1
                    If monthIndex >= 1 And monthIndex <= 12 Then monthIncomplete(monthIndex) = monthIncomplete(monthIndex) + 1
                
                Case "Austrittsgrund unbekannt"
                    incompleteExits = incompleteExits + 1
                    If monthIndex >= 1 And monthIndex <= 12 Then monthIncomplete(monthIndex) = monthIncomplete(monthIndex) + 1
                
                Case "Normaler Austritt"
                    normalExits = normalExits + 1
            End Select
        Next r
    End If
    
    showEarly = (earlyExits > 0)
    showExperienced = (experiencedLoss > 0)
    showImportant = (importantExits > 0)
    showNeutral = (neutralExits > 0)
    showIncomplete = (incompleteExits > 0)
    
    If totalExits > 0 Then
        avgLoss = totalLoss / totalExits
    Else
        avgLoss = 0
    End If
    
    ' --- Abschnitt 2: Monatliche Fluktuation und YTD berechnen ---
    PID_FillFluktuationRates monthExit, monthPersonalEnde, monthFluctuation, quarterFluctuation, ytdFluctuation, currentYear
    
    ' --- Abschnitt 3: Risikobewertung und Manager-Texte ---
    criticalExits = earlyExits + experiencedLoss + importantExits
    riskLevel = GetFluktuationRiskLevel(totalLoss, totalExits, earlyExits, experiencedLoss, importantExits, incompleteExits)
    managerSummary = GetFluktuationManagerSummary(riskLevel, currentYear, totalExits, incompleteExits, criticalExits, ytdFluctuation, totalLoss)
    recommendationItems = GetFluktuationRecommendationItems(totalExits, neutralExits, earlyExits, experiencedLoss, importantExits, incompleteExits, totalLoss, ytdFluctuation)
    
    ' --- Abschnitt 4: Ausgabe in Fluktuation-Blatt schreiben ---
    With analyseWs
        PID_ClearFluktuationCharts analyseWs
        
        .Range("A1").Value = "Fluktuation - Uebersicht fuer die Restaurantleitung"
        .Range("A2").Value = "Jahr"
        .Range("B2").Value = currentYear
        .Range("C2").Value = "Stand"
        .Range("D2").Value = Format(Date, "dd.mm.yyyy")
        
        statusRow = 5
        kpiLabelRow = 8
        kpiValueRow = 9
        
        .Range("A4").Value = "Auf einen Blick"
        .Range("A" & statusRow).Value = "Status"
        .Range("B" & statusRow).Value = riskLevel
        .Range("C" & statusRow).Value = managerSummary
        .Range("C" & statusRow & ":E" & statusRow + 1).Merge
        .Range("A" & statusRow & ":A" & statusRow + 1).Merge
        .Range("B" & statusRow & ":B" & statusRow + 1).Merge
        
        .Cells(kpiLabelRow, 1).Value = "Austritte gesamt"
        .Cells(kpiLabelRow, 2).Value = "Jahresfluktuation"
        .Cells(kpiLabelRow, 3).Value = "Verlust-Score"
        .Cells(kpiLabelRow, 4).Value = "Kritische Austritte"
        .Cells(kpiLabelRow, 5).Value = "Daten offen"
        
        .Cells(kpiValueRow, 1).Value = totalExits
        .Cells(kpiValueRow, 2).Value = ytdFluctuation
        .Cells(kpiValueRow, 3).Value = totalLoss
        .Cells(kpiValueRow, 4).Value = criticalExits
        .Cells(kpiValueRow, 5).Value = incompleteExits
        
        alertsHeaderRow = kpiValueRow + 2
        alertsEndRow = WriteFluktuationAlertsSection(analyseWs, dataWs, alertsHeaderRow)
        
        recHeaderRow = alertsEndRow + 2
        recEndRow = WriteFluktuationRecommendationsSection(analyseWs, recommendationItems, recHeaderRow)
        
        chartRow = recEndRow + 2
        
        ' --- Abschnitt 4b: Monatstabelle ---
        monthlyTitleRow = chartRow + 34
        headerRow = monthlyTitleRow + 2
        firstDataRow = headerRow + 1
        
        .Range("A" & monthlyTitleRow).Value = "Monatsuebersicht im Detail"
        
        currentCol = 1
        
        .Cells(headerRow, currentCol).Value = "Monat"
        currentCol = currentCol + 1
        
        .Cells(headerRow, currentCol).Value = "Austritte"
        currentCol = currentCol + 1
        
        .Cells(headerRow, currentCol).Value = "Fluktuation %"
        currentCol = currentCol + 1
        
        .Cells(headerRow, currentCol).Value = "Verlust-Score"
        currentCol = currentCol + 1
        
        .Cells(headerRow, currentCol).Value = "Durchschnitt"
        currentCol = currentCol + 1
        
        If showEarly Then
            .Cells(headerRow, currentCol).Value = "Erste 90 Tage"
            currentCol = currentCol + 1
        End If
        
        If showExperienced Then
            .Cells(headerRow, currentCol).Value = "Erfahrene MA"
            currentCol = currentCol + 1
        End If
        
        If showImportant Then
            .Cells(headerRow, currentCol).Value = "Wichtige Austritte"
            currentCol = currentCol + 1
        End If
        
        If showNeutral Then
            .Cells(headerRow, currentCol).Value = "Neutrale Bewegungen"
            currentCol = currentCol + 1
        End If
        
        If showIncomplete Then
            .Cells(headerRow, currentCol).Value = "Unvollstaendig"
            currentCol = currentCol + 1
        End If
        
        .Cells(headerRow, currentCol).Value = "Austrittsgruende"
        currentCol = currentCol + 1
        
        .Cells(headerRow, currentCol).Value = "Monats-Hinweis"
        lastTableCol = currentCol
        
        outputRow = firstDataRow
        
        For i = 1 To 12
            If monthExit(i) > 0 Then
                currentCol = 1
                
                .Cells(outputRow, currentCol).Value = monthNames(i - 1)
                currentCol = currentCol + 1
                
                .Cells(outputRow, currentCol).Value = monthExit(i)
                currentCol = currentCol + 1
                
                .Cells(outputRow, currentCol).Value = monthFluctuation(i)
                currentCol = currentCol + 1
                
                .Cells(outputRow, currentCol).Value = monthLoss(i)
                currentCol = currentCol + 1
                
                avgMonthLoss = monthLoss(i) / monthExit(i)
                
                .Cells(outputRow, currentCol).Value = avgMonthLoss
                currentCol = currentCol + 1
                
                If showEarly Then
                    .Cells(outputRow, currentCol).Value = monthEarly(i)
                    currentCol = currentCol + 1
                End If
                
                If showExperienced Then
                    .Cells(outputRow, currentCol).Value = monthExperienced(i)
                    currentCol = currentCol + 1
                End If
                
                If showImportant Then
                    .Cells(outputRow, currentCol).Value = monthImportant(i)
                    currentCol = currentCol + 1
                End If
                
                If showNeutral Then
                    .Cells(outputRow, currentCol).Value = monthNeutral(i)
                    currentCol = currentCol + 1
                End If
                
                If showIncomplete Then
                    .Cells(outputRow, currentCol).Value = monthIncomplete(i)
                    currentCol = currentCol + 1
                End If
                
                reasonSummary = GetExitReasonSummaryForMonth(dataWs, CStr(monthNames(i - 1)))
                .Cells(outputRow, currentCol).Value = reasonSummary
                currentCol = currentCol + 1
                
                monthHint = GetMonthHint(monthExit(i), monthEarly(i), monthExperienced(i), monthImportant(i), monthNeutral(i), monthIncomplete(i), monthLoss(i))
                .Cells(outputRow, currentCol).Value = monthHint
                
                .Range(.Cells(outputRow, lastTableCol), .Cells(outputRow, lastTableCol + 2)).Merge
                
                outputRow = outputRow + 1
            End If
        Next i
        
        If outputRow = firstDataRow Then
            .Cells(firstDataRow, 1).Value = "Keine Austritte erfasst."
            .Range(.Cells(firstDataRow, 1), .Cells(firstDataRow, lastTableCol + 2)).Merge
            outputRow = outputRow + 1
        End If
        
        explanationStartRow = outputRow + 3
        
        PID_WriteFluktuationExplanationRows analyseWs, explanationStartRow
        
        FormatFluktuationSheet analyseWs, statusRow, kpiLabelRow, kpiValueRow, alertsHeaderRow, alertsEndRow, recHeaderRow, recEndRow, chartRow, monthlyTitleRow, headerRow, firstDataRow, outputRow, lastTableCol, explanationStartRow, riskLevel
        
        BuildFluktuationCharts analyseWs, dataWs, chartRow, monthNames, monthExit, earlyExits, experiencedLoss, importantExits, neutralExits, normalExits, incompleteExits
        
        PID_SyncFluktuationToDisplaySheets monthFluctuation, quarterFluctuation, ytdFluctuation
        
        .Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
    End With

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not analyseWs Is Nothing Then
        analyseWs.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
    End If
    
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler in BuildFluktuationAnalyse:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Fluktuation Analyse"
End Sub


Public Sub FormatFluktuationSheet(ByVal ws As Worksheet, _
                                  ByVal statusRow As Long, _
                                  ByVal kpiLabelRow As Long, _
                                  ByVal kpiValueRow As Long, _
                                  ByVal alertsHeaderRow As Long, _
                                  ByVal alertsEndRow As Long, _
                                  ByVal recHeaderRow As Long, _
                                  ByVal recEndRow As Long, _
                                  ByVal chartRow As Long, _
                                  ByVal monthlyTitleRow As Long, _
                                  ByVal headerRow As Long, _
                                  ByVal firstDataRow As Long, _
                                  ByVal outputRow As Long, _
                                  ByVal lastTableCol As Long, _
                                  ByVal explanationStartRow As Long, _
                                  ByVal riskLevel As String)
    Dim tableLastCol As Long
    Dim dataRow As Long
    Dim r As Long
    
    tableLastCol = Application.WorksheetFunction.Max(5, lastTableCol + 2)
    
    With ws
        .Range("A1:E1").Merge
        .Rows(1).RowHeight = PID_STYLE_TITLE_ROW_HEIGHT
        PID_StyleApplyTitleBand .Range("A1:E1")
        
        .Rows(2).RowHeight = PID_STYLE_TITLE_SUB_ROW_HEIGHT
        PID_StyleApplyHeaderBand .Range("A2:E2")
        
        .Range("A4:E4").Merge
        .Rows(4).RowHeight = 22
        PID_StyleApplySubsectionTitle .Range("A4:E4"), True
        
        .Rows(statusRow).RowHeight = 24
        .Rows(statusRow + 1).RowHeight = 24
        PID_StyleApplyHeaderBand .Range("A" & statusRow & ":A" & statusRow + 1)
        .Range("A" & statusRow).HorizontalAlignment = xlCenter
        .Range("C" & statusRow).WrapText = True
        .Range("C" & statusRow).HorizontalAlignment = xlLeft
        .Range("B" & statusRow).Font.Size = 11
        ApplyRiskFormatting .Range("B" & statusRow), riskLevel
        PID_StyleApplyTableBorders .Range("A" & statusRow & ":E" & statusRow + 1)
        
        .Rows(kpiLabelRow).RowHeight = 24
        .Rows(kpiValueRow).RowHeight = PID_STYLE_COMPACT_DATA_ROW_HEIGHT
        PID_StyleApplyHeaderBand .Range("A" & kpiLabelRow & ":E" & kpiLabelRow)
        .Range("A" & kpiLabelRow & ":E" & kpiLabelRow).WrapText = True
        .Range("A" & kpiLabelRow & ":E" & kpiValueRow).HorizontalAlignment = xlCenter
        .Range("B" & kpiValueRow).NumberFormat = "0.00%"
        .Range("C" & kpiValueRow).NumberFormat = "0.00"
        
        If incompleteExitsCellNeedsHighlight(ws, kpiValueRow) Then
            PID_StyleApplyInputHighlight .Range("E" & kpiValueRow)
        End If
        PID_StyleApplyTableBorders .Range("A" & kpiLabelRow & ":E" & kpiValueRow)
        
        .Range("A" & alertsHeaderRow & ":E" & alertsHeaderRow).Merge
        .Rows(alertsHeaderRow).RowHeight = 22
        PID_StyleApplySubsectionTitle .Range("A" & alertsHeaderRow & ":E" & alertsHeaderRow), True
        If alertsEndRow > alertsHeaderRow + 2 Then
            PID_StyleApplyHeaderBand .Range(.Cells(alertsHeaderRow + 1, 1), .Cells(alertsHeaderRow + 1, 5))
            PID_StyleApplyTableBorders .Range(.Cells(alertsHeaderRow + 1, 1), .Cells(alertsEndRow - 1, 5))
        ElseIf alertsEndRow = alertsHeaderRow + 2 Then
            .Range(.Cells(alertsHeaderRow + 1, 1), .Cells(alertsHeaderRow + 1, 5)).WrapText = True
            PID_StyleApplyTableBorders .Range(.Cells(alertsHeaderRow + 1, 1), .Cells(alertsHeaderRow + 1, 5))
        End If
        
        .Range("A" & recHeaderRow & ":E" & recHeaderRow).Merge
        .Rows(recHeaderRow).RowHeight = 22
        PID_StyleApplySubsectionTitle .Range("A" & recHeaderRow & ":E" & recHeaderRow), True
        If recEndRow > recHeaderRow Then
            PID_StyleApplyTableBorders .Range(.Cells(recHeaderRow + 1, 1), .Cells(recEndRow, 5))
        End If
        
        .Range(.Cells(monthlyTitleRow, 1), .Cells(monthlyTitleRow, tableLastCol)).Merge
        .Rows(monthlyTitleRow).RowHeight = PID_STYLE_TITLE_ROW_HEIGHT
        PID_StyleApplyTitleBand .Range(.Cells(monthlyTitleRow, 1), .Cells(monthlyTitleRow, tableLastCol))
        
        .Rows(headerRow).RowHeight = PID_STYLE_COMPACT_HEADER_ROW_HEIGHT
        PID_StyleApplyHeaderBand .Range(.Cells(headerRow, 1), .Cells(headerRow, tableLastCol))
        
        PID_StyleApplyTableBorders .Range(.Cells(headerRow, 1), .Cells(outputRow - 1, tableLastCol))
        .Range(.Cells(headerRow, 1), .Cells(outputRow - 1, tableLastCol)).HorizontalAlignment = xlCenter
        .Range(.Cells(firstDataRow, 1), .Cells(outputRow - 1, tableLastCol)).WrapText = True
        
        If lastTableCol >= 2 Then
            .Range(.Cells(firstDataRow, lastTableCol - 1), .Cells(outputRow - 1, lastTableCol)).HorizontalAlignment = xlLeft
        End If
        
        .Range(.Cells(firstDataRow, 3), .Cells(outputRow - 1, 3)).NumberFormat = "0.00%"
        .Range(.Cells(firstDataRow, 4), .Cells(outputRow - 1, 5)).NumberFormat = "0.00"
        
        .Range("A" & explanationStartRow & ":E" & explanationStartRow).Merge
        .Rows(explanationStartRow).RowHeight = 22
        PID_StyleApplySubsectionTitle .Range("A" & explanationStartRow & ":E" & explanationStartRow), True
        .Range("B" & explanationStartRow + 1 & ":E" & explanationStartRow + 8).WrapText = True
        PID_StyleApplyTableBorders .Range("A" & explanationStartRow & ":E" & explanationStartRow + 8)
        
        PID_FlApplyReadOnlyZones ws, statusRow, kpiValueRow, alertsHeaderRow, alertsEndRow, recHeaderRow, recEndRow, firstDataRow, outputRow, tableLastCol, lastTableCol, explanationStartRow
        
        .Columns("A").ColumnWidth = 16
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 26
        .Columns("D").ColumnWidth = 26
        .Columns("E").ColumnWidth = 26
        .Columns("F").ColumnWidth = 18
        .Columns("G").ColumnWidth = 18
        .Columns("H").ColumnWidth = 22
        .Columns("I").ColumnWidth = 36
        .Columns("J").ColumnWidth = 36
        .Columns("K").ColumnWidth = 36
        .Columns("L").ColumnWidth = 36
        .Columns("P").ColumnWidth = 2
        .Columns("Q").ColumnWidth = 2
        .Columns("R").ColumnWidth = 2
        .Columns("S").ColumnWidth = 2
        
        .Rows(firstDataRow & ":" & outputRow - 1).RowHeight = 34
        .Rows(chartRow + 1 & ":" & chartRow + 32).RowHeight = 18
        
        PID_AutoFitFluktuationTextRows ws, statusRow, alertsHeaderRow, alertsEndRow, recHeaderRow, recEndRow, explanationStartRow
        
        .Range(.Cells(1, 1), .Cells(explanationStartRow + 8, tableLastCol)).VerticalAlignment = xlCenter
    End With
End Sub


Private Sub PID_FlApplyReadOnlyZones(ByVal ws As Worksheet, _
                                     ByVal statusRow As Long, _
                                     ByVal kpiValueRow As Long, _
                                     ByVal alertsHeaderRow As Long, _
                                     ByVal alertsEndRow As Long, _
                                     ByVal recHeaderRow As Long, _
                                     ByVal recEndRow As Long, _
                                     ByVal firstDataRow As Long, _
                                     ByVal outputRow As Long, _
                                     ByVal tableLastCol As Long, _
                                     ByVal lastTableCol As Long, _
                                     ByVal explanationStartRow As Long)
    Dim dataRow As Long
    Dim r As Long
    
    PID_StyleApplyReadOnlyGuideCell ws.Range("C" & statusRow & ":E" & statusRow + 1)
    ws.Range("C" & statusRow & ":E" & statusRow + 1).HorizontalAlignment = xlLeft
    ws.Range("C" & statusRow & ":E" & statusRow + 1).WrapText = True
    
    PID_StyleApplyReadOnlyGuideCell ws.Range("A" & kpiValueRow)
    PID_StyleApplyReadOnlyGuideCell ws.Range("C" & kpiValueRow & ":D" & kpiValueRow)
    If Not incompleteExitsCellNeedsHighlight(ws, kpiValueRow) Then
        PID_StyleApplyReadOnlyGuideCell ws.Range("E" & kpiValueRow)
    End If
    
    If alertsEndRow > alertsHeaderRow + 2 Then
        For r = alertsHeaderRow + 2 To alertsEndRow - 1
            PID_StyleApplyInputGuideLabel ws.Cells(r, 1)
            PID_StyleApplyReadOnlyGuideCell ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).HorizontalAlignment = xlLeft
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).WrapText = True
        Next r
    ElseIf alertsEndRow = alertsHeaderRow + 2 Then
        PID_StyleApplyReadOnlyGuideCell ws.Range(ws.Cells(alertsHeaderRow + 1, 1), ws.Cells(alertsHeaderRow + 1, 5))
        ws.Range(ws.Cells(alertsHeaderRow + 1, 1), ws.Cells(alertsHeaderRow + 1, 5)).HorizontalAlignment = xlLeft
        ws.Range(ws.Cells(alertsHeaderRow + 1, 1), ws.Cells(alertsHeaderRow + 1, 5)).WrapText = True
    End If
    
    If recEndRow > recHeaderRow Then
        For r = recHeaderRow + 1 To recEndRow
            PID_StyleApplyInputGuideLabel ws.Cells(r, 1)
            PID_StyleApplyReadOnlyGuideCell ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).HorizontalAlignment = xlLeft
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).WrapText = True
        Next r
    End If
    
    For dataRow = firstDataRow To outputRow - 1
        PID_StyleApplyInputGuideLabel ws.Cells(dataRow, 1)
        If tableLastCol >= 2 Then
            PID_StyleApplyReadOnlyGuideCell ws.Range(ws.Cells(dataRow, 2), ws.Cells(dataRow, tableLastCol))
            ws.Range(ws.Cells(dataRow, 2), ws.Cells(dataRow, tableLastCol)).Font.Color = vbBlack
            ws.Range(ws.Cells(dataRow, 2), ws.Cells(dataRow, tableLastCol)).Font.Bold = False
            If lastTableCol >= 2 Then
                ws.Range(ws.Cells(dataRow, lastTableCol - 1), ws.Cells(dataRow, tableLastCol)).HorizontalAlignment = xlLeft
            End If
        End If
    Next dataRow
    
    For r = explanationStartRow + 1 To explanationStartRow + 8
        PID_StyleApplyInputGuideLabel ws.Cells(r, 1)
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Color = PID_StyleColorNavy()
        PID_StyleApplyReadOnlyGuideCell ws.Range("B" & r & ":E" & r)
        ws.Range("B" & r & ":E" & r).HorizontalAlignment = xlLeft
        ws.Range("B" & r & ":E" & r).Font.Color = vbBlack
        ws.Range("B" & r & ":E" & r).Font.Bold = False
    Next r
End Sub


Private Sub PID_AutoFitFluktuationTextRows(ByVal ws As Worksheet, _
                                           ByVal statusRow As Long, _
                                           ByVal alertsHeaderRow As Long, _
                                           ByVal alertsEndRow As Long, _
                                           ByVal recHeaderRow As Long, _
                                           ByVal recEndRow As Long, _
                                           ByVal explanationStartRow As Long)
    Dim r As Long
    Dim maxHeight As Double
    Dim rowHeight As Double
    
    maxHeight = 180
    
    ws.Rows(statusRow & ":" & statusRow + 1).AutoFit
    For r = statusRow To statusRow + 1
        If ws.Rows(r).RowHeight < 24 Then ws.Rows(r).RowHeight = 24
    Next r
    
    If alertsEndRow > alertsHeaderRow + 2 Then
        ws.Rows(alertsHeaderRow + 1).RowHeight = 22
        For r = alertsHeaderRow + 2 To alertsEndRow - 1
            rowHeight = PID_EstimateWrappedRowHeightForCell(ws, r, 2, 2)
            rowHeight = Application.WorksheetFunction.Max(rowHeight, PID_EstimateWrappedRowHeightForCell(ws, r, 3, 3))
            rowHeight = Application.WorksheetFunction.Max(rowHeight, PID_EstimateWrappedRowHeightForCell(ws, r, 4, 5))
            If rowHeight > maxHeight Then rowHeight = maxHeight
            ws.Rows(r).RowHeight = rowHeight
        Next r
    ElseIf alertsEndRow = alertsHeaderRow + 2 Then
        rowHeight = PID_EstimateWrappedRowHeightForCell(ws, alertsHeaderRow + 1, 1, 5)
        If rowHeight < 28 Then rowHeight = 28
        ws.Rows(alertsHeaderRow + 1).RowHeight = rowHeight
    End If
    
    If recEndRow > recHeaderRow Then
        For r = recHeaderRow + 1 To recEndRow
            rowHeight = PID_EstimateWrappedRowHeightForCell(ws, r, 2, 5)
            If rowHeight > maxHeight Then rowHeight = maxHeight
            ws.Rows(r).RowHeight = rowHeight
        Next r
    End If
    
    For r = explanationStartRow + 1 To explanationStartRow + 8
        rowHeight = PID_EstimateWrappedRowHeightForCell(ws, r, 2, 5)
        If rowHeight < 28 Then rowHeight = 28
        If rowHeight > maxHeight Then rowHeight = maxHeight
        ws.Rows(r).RowHeight = rowHeight
    Next r
End Sub


Private Function PID_EstimateWrappedRowHeightForCell(ByVal ws As Worksheet, _
                                                     ByVal rowNum As Long, _
                                                     ByVal firstCol As Long, _
                                                     ByVal lastCol As Long) As Double
    Dim txt As String
    Dim totalWidth As Double
    Dim c As Long
    Dim charsPerLine As Long
    Dim lines As Long
    Dim lineLen As Long
    Dim words() As String
    Dim i As Long
    Dim word As String
    
    txt = Trim$(CStr(ws.Cells(rowNum, firstCol).Value))
    If txt = "" Then
        PID_EstimateWrappedRowHeightForCell = 24
        Exit Function
    End If
    
    totalWidth = 0
    For c = firstCol To lastCol
        totalWidth = totalWidth + ws.Columns(c).ColumnWidth
    Next c
    
    charsPerLine = CLng(totalWidth)
    If charsPerLine < 12 Then charsPerLine = 12
    
    words = Split(txt, " ")
    lines = 0
    lineLen = 0
    
    For i = LBound(words) To UBound(words)
        word = Trim$(CStr(words(i)))
        If word = "" Then GoTo NextWord
        
        If lineLen = 0 Then
            lineLen = Len(word)
            lines = lines + 1
        ElseIf lineLen + 1 + Len(word) <= charsPerLine Then
            lineLen = lineLen + 1 + Len(word)
        Else
            lines = lines + 1
            lineLen = Len(word)
        End If
NextWord:
    Next i
    
    If lines < 1 Then lines = 1
    
    PID_EstimateWrappedRowHeightForCell = (lines * 15.75) + 10
    If PID_EstimateWrappedRowHeightForCell < 28 Then PID_EstimateWrappedRowHeightForCell = 28
End Function


Private Function incompleteExitsCellNeedsHighlight(ByVal ws As Worksheet, ByVal kpiValueRow As Long) As Boolean
    If IsNumeric(ws.Cells(kpiValueRow, 5).Value) Then
        incompleteExitsCellNeedsHighlight = (CLng(ws.Cells(kpiValueRow, 5).Value) > 0)
    Else
        incompleteExitsCellNeedsHighlight = False
    End If
End Function


Public Function GetLastMonthWithExit(ByRef monthExit() As Long) As Long
    Dim i As Long
    
    GetLastMonthWithExit = 0
    
    For i = 12 To 1 Step -1
        If monthExit(i) > 0 Then
            GetLastMonthWithExit = i
            Exit Function
        End If
    Next i
End Function


Public Function GetPersonalEndeForMonth(ByVal monthSheetName As String, _
                                        ByVal currentYear As Long, _
                                        ByVal monthNumber As Long) As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim monthEndDate As Date
    Dim employeeID As Variant
    Dim employeeName As Variant
    Dim entryDate As Variant
    Dim exitDate As Variant
    Dim countEmployees As Long
    
    On Error GoTo SafeExit
    
    Set ws = ThisWorkbook.Worksheets(monthSheetName)
    monthEndDate = DateSerial(currentYear, monthNumber + 1, 0)
    
    countEmployees = 0
    
    For r = 3 To 82
        employeeID = ws.Cells(r, "B").Value
        employeeName = ws.Cells(r, "C").Value
        entryDate = ws.Cells(r, "D").Value
        exitDate = ws.Cells(r, "I").Value
        
        If Trim$(CStr(employeeID)) <> "" Or Trim$(CStr(employeeName)) <> "" Then
            If IsDate(entryDate) Then
                If CDate(entryDate) <= monthEndDate Then
                    If Not IsDate(exitDate) Then
                        countEmployees = countEmployees + 1
                    ElseIf CDate(exitDate) > monthEndDate Then
                        countEmployees = countEmployees + 1
                    End If
                End If
            Else
                If Not IsDate(exitDate) Then
                    countEmployees = countEmployees + 1
                ElseIf CDate(exitDate) > monthEndDate Then
                    countEmployees = countEmployees + 1
                End If
            End If
        End If
    Next r
    
    GetPersonalEndeForMonth = countEmployees
    Exit Function

SafeExit:
    GetPersonalEndeForMonth = 0
End Function


Public Function GetExitReasonSummaryForMonth(ByVal dataWs As Worksheet, ByVal monthName As String) As String
    Dim lastRow As Long
    Dim r As Long
    Dim reasonText As String
    Dim categoryText As String
    Dim checkMonth As String
    Dim resultText As String
    
    lastRow = dataWs.Cells(dataWs.Rows.count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        GetExitReasonSummaryForMonth = "-"
        Exit Function
    End If
    
    resultText = ""
    
    For r = 2 To lastRow
        checkMonth = Trim$(CStr(dataWs.Cells(r, "A").Value))
        
        If checkMonth = monthName Then
            reasonText = Trim$(CStr(dataWs.Cells(r, "F").Value))
            categoryText = Trim$(CStr(dataWs.Cells(r, "K").Value))
            
            If categoryText = "Austrittsgrund fehlt" Then
                reasonText = "Austrittsgrund fehlt"
            ElseIf categoryText = "Austrittsgrund unbekannt" Then
                reasonText = "Austrittsgrund unbekannt"
            End If
            
            If reasonText = "" Then reasonText = "Austrittsgrund fehlt"
            
            resultText = AddReasonToSummary(resultText, reasonText)
        End If
    Next r
    
    If resultText = "" Then
        GetExitReasonSummaryForMonth = "-"
    Else
        GetExitReasonSummaryForMonth = resultText
    End If
End Function


Public Function AddReasonToSummary(ByVal currentText As String, ByVal reasonText As String) As String
    Dim parts As Variant
    Dim i As Long
    Dim partText As String
    Dim baseReason As String
    Dim countText As String
    Dim countValue As Long
    Dim resultText As String
    
    If currentText = "" Then
        AddReasonToSummary = reasonText & " (1)"
        Exit Function
    End If
    
    parts = Split(currentText, ", ")
    resultText = ""
    
    For i = LBound(parts) To UBound(parts)
        partText = CStr(parts(i))
        
        If InStr(partText, " (") > 0 Then
            baseReason = Left$(partText, InStr(partText, " (") - 1)
            countText = Replace(partText, baseReason & " (", "")
            countText = Replace(countText, ")", "")
            
            If IsNumeric(countText) Then
                countValue = CLng(countText)
            Else
                countValue = 1
            End If
            
            If baseReason = reasonText Then
                countValue = countValue + 1
                partText = baseReason & " (" & countValue & ")"
            End If
        End If
        
        If resultText <> "" Then resultText = resultText & ", "
        resultText = resultText & partText
    Next i
    
    If InStr(currentText, reasonText & " (") = 0 Then
        resultText = resultText & ", " & reasonText & " (1)"
    End If
    
    AddReasonToSummary = resultText
End Function


Public Function GetMonthHint(ByVal exitsCount As Long, _
                             ByVal earlyCount As Long, _
                             ByVal experiencedCount As Long, _
                             ByVal importantCount As Long, _
                             ByVal neutralCount As Long, _
                             ByVal incompleteCount As Long, _
                             ByVal lossValue As Double) As String
    If exitsCount = 0 Then
        GetMonthHint = "Keine Austritte erfasst."
    ElseIf incompleteCount > 0 Then
        GetMonthHint = "Mindestens ein Austritt hat keinen gueltigen Austrittsgrund. Bitte Daten pruefen und ergaenzen."
    ElseIf neutralCount = exitsCount Then
        GetMonthHint = "Nur neutrale Bewegungen. Kein direkter Handlungsbedarf."
    ElseIf experiencedCount >= 1 Then
        GetMonthHint = "Erfahrener Mitarbeiter ausgeschieden. Mitarbeiterbindung und Teamstabilitaet pruefen."
    ElseIf earlyCount >= 1 Then
        GetMonthHint = "Austritt in den ersten 90 Tagen. Onboarding, Training und erste Dienstplaene pruefen."
    ElseIf importantCount >= 1 Then
        GetMonthHint = "Wichtiger Austritt. Arbeitsklima, Fuehrung und Kommunikation pruefen."
    ElseIf lossValue >= 1 Then
        GetMonthHint = "Verlust-Score erhoeht. Austrittsgrund genauer pruefen."
    Else
        GetMonthHint = "Austritt vorhanden, aktuell nicht kritisch. Entwicklung beobachten."
    End If
End Function


Public Function GetFluktuationRiskLevel(ByVal totalLoss As Double, _
                                        ByVal totalExits As Long, _
                                        ByVal earlyExits As Long, _
                                        ByVal experiencedLoss As Long, _
                                        ByVal importantExits As Long, _
                                        ByVal incompleteExits As Long) As String
    If incompleteExits > 0 Then
        GetFluktuationRiskLevel = "Daten pruefen"
    ElseIf totalExits = 0 Then
        GetFluktuationRiskLevel = "Niedrig"
    ElseIf totalLoss >= 5 Or experiencedLoss >= 3 Or importantExits >= 4 Then
        GetFluktuationRiskLevel = "Kritisch"
    ElseIf totalLoss >= 3 Or experiencedLoss >= 2 Or importantExits >= 2 Then
        GetFluktuationRiskLevel = "Hoch"
    ElseIf totalLoss >= 1.5 Or earlyExits >= 2 Then
        GetFluktuationRiskLevel = "Mittel"
    Else
        GetFluktuationRiskLevel = "Niedrig"
    End If
End Function


Public Sub ApplyRiskFormatting(ByVal targetCell As Range, ByVal riskLevel As String)
    With targetCell
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        Select Case riskLevel
            Case "Niedrig"
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)
            
            Case "Mittel"
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            
            Case "Hoch"
                .Interior.Color = RGB(244, 176, 132)
                .Font.Color = RGB(128, 64, 0)
            
            Case "Kritisch"
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(156, 0, 6)
            
            Case "Daten pruefen"
                .Interior.Color = RGB(221, 235, 247)
                .Font.Color = RGB(31, 78, 121)
            
            Case Else
                .Interior.Pattern = xlNone
                .Font.Color = vbBlack
        End Select
    End With
End Sub


Public Function GetFluktuationManagerSummary(ByVal riskLevel As String, _
                                              ByVal currentYear As Long, _
                                              ByVal totalExits As Long, _
                                              ByVal incompleteExits As Long, _
                                              ByVal criticalExits As Long, _
                                              ByVal ytdFluctuation As Double, _
                                              ByVal totalLoss As Double) As String
    If incompleteExits > 0 Then
        GetFluktuationManagerSummary = "Die Auswertung ist noch nicht vollstaendig. " & incompleteExits & " Austritt(e) haben keinen gueltigen Austrittsgrund. Bitte zuerst die markierten Faelle unten in den Monatsblaettern ergaenzen."
    ElseIf totalExits = 0 Then
        GetFluktuationManagerSummary = "Im Jahr " & currentYear & " sind bisher keine Austritte erfasst. Die Fluktuation ist aktuell unauffaellig."
    ElseIf riskLevel = "Kritisch" Then
        GetFluktuationManagerSummary = "Die Fluktuation ist kritisch. Es gibt " & totalExits & " Austritte, davon " & criticalExits & " mit erhoehtem Risiko. Verlust-Score: " & Format(totalLoss, "0.00") & ". Sofort Massnahmen pruefen."
    ElseIf riskLevel = "Hoch" Then
        GetFluktuationManagerSummary = "Die Fluktuation ist erhoeht (" & Format(ytdFluctuation, "0.0%") & " bisher im Jahr). " & criticalExits & " Austritt(e) sind besonders relevant. Ursachen und Massnahmen unten pruefen."
    ElseIf riskLevel = "Mittel" Then
        GetFluktuationManagerSummary = "Es gibt " & totalExits & " Austritte im Jahr " & currentYear & ". Die Lage ist beobachtungswuerdig, aber noch nicht kritisch. Empfehlungen und Monatsdetails unten beachten."
    Else
        GetFluktuationManagerSummary = "Es gibt " & totalExits & " Austritte, die Gesamtbewertung ist aktuell stabil. Entwicklung weiter beobachten und Daten aktuell halten."
    End If
End Function


Public Function GetFluktuationRecommendationItems(ByVal totalExits As Long, _
                                                   ByVal neutralExits As Long, _
                                                   ByVal earlyExits As Long, _
                                                   ByVal experiencedLoss As Long, _
                                                   ByVal importantExits As Long, _
                                                   ByVal incompleteExits As Long, _
                                                   ByVal totalLoss As Double, _
                                                   ByVal ytdFluctuation As Double) As Variant
    Dim items() As String
    Dim itemCount As Long
    
    itemCount = 0
    
    If incompleteExits > 0 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Daten vervollstaendigen: Bei jedem Austritt in Spalte I das Austrittsdatum und in Spalte N den Austrittsgrund aus der Liste eintragen. Ohne diese Angaben kann die Analyse nicht zuverlaessig bewerten."
    End If
    
    If experiencedLoss >= 1 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Erfahrene Mitarbeiter halten: Exit-Gespraeche auswerten, Entwicklungsgespraeche planen, Schichtplanung und Fuehrung im Team pruefen."
    End If
    
    If earlyExits >= 1 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Erste 90 Tage staerken: Recruiting, Einarbeitung, Buddy-System, Training und die ersten Dienstplaene gezielt verbessern."
    End If
    
    If importantExits >= 1 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Wichtige Austritte analysieren: Austrittsgrund, Arbeitsklima, Kommunikation und Fuehrung im betroffenen Bereich besprechen."
    End If
    
    If totalLoss >= 3 And incompleteExits = 0 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Verlust-Score senken: Haeufige Austrittsgruende sammeln und konkrete Massnahmen zur Mitarbeiterbindung ableiten."
    End If
    
    If ytdFluctuation >= 0.08 And incompleteExits = 0 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Fluktuation beobachten: Die bisherige Jahresfluktuation liegt ueber 8 Prozent. Personalplanung und Nachbesetzungen fruehzeitig abstimmen."
    End If
    
    If totalExits = 0 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Keine Massnahmen noetig: Keine Austritte erfasst. Monatsblaetter weiter aktuell pflegen."
    ElseIf neutralExits = totalExits Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Neutrale Bewegungen: Storetransfer, Befoerderung oder Karenz sind kein negatives Fluktuationssignal. Kein unmittelbarer Handlungsbedarf."
    ElseIf itemCount = 0 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = "Entwicklung beobachten: Die Fluktuation ist vorhanden, aber aktuell nicht kritisch. Monatliche Entwicklung und Austrittsgruende im Blick behalten."
    End If
    
    GetFluktuationRecommendationItems = items
End Function


Public Function WriteFluktuationAlertsSection(ByVal ws As Worksheet, _
                                              ByVal dataWs As Worksheet, _
                                              ByVal startRow As Long) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long
    Dim alertNum As Long
    Dim hasAlerts As Boolean
    
    ws.Range("A" & startRow).Value = "Sofort pruefen"
    outRow = startRow + 1
    
    lastRow = dataWs.Cells(dataWs.Rows.count, "A").End(xlUp).Row
    hasAlerts = FluktuationHasAlertRows(dataWs, lastRow)
    
    If hasAlerts Then
        ws.Cells(outRow, 1).Value = "Nr."
        ws.Cells(outRow, 2).Value = "Problem"
        ws.Cells(outRow, 3).Value = "Wo nachschauen"
        ws.Cells(outRow, 4).Value = "Naechster Schritt"
        ws.Range(ws.Cells(outRow, 4), ws.Cells(outRow, 5)).Merge
        outRow = outRow + 1
    End If
    
    alertNum = 0
    
    If lastRow >= 2 And hasAlerts Then
        For r = 2 To lastRow
            If alertNum >= 12 Then Exit For
            If WriteFluktuationAlertIfMatch(ws, dataWs, r, outRow, alertNum, True) Then
                outRow = outRow + 1
            End If
        Next r
        
        For r = 2 To lastRow
            If alertNum >= 12 Then Exit For
            If WriteFluktuationAlertIfMatch(ws, dataWs, r, outRow, alertNum, False) Then
                outRow = outRow + 1
            End If
        Next r
    End If
    
    If Not hasAlerts Then
        ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, 5)).Merge
        ws.Cells(outRow, 1).Value = "Keine offenen Datenprobleme oder kritischen Einzelfaelle. Die erfassten Austritte sind vollstaendig und aktuell bewertbar."
        ws.Cells(outRow, 1).Interior.Color = RGB(198, 239, 206)
        ws.Cells(outRow, 1).Font.Color = RGB(0, 97, 0)
        outRow = outRow + 1
    End If
    
    WriteFluktuationAlertsSection = outRow
End Function


Private Function FluktuationHasAlertRows(ByVal dataWs As Worksheet, ByVal lastRow As Long) As Boolean
    Dim r As Long
    Dim categoryText As String
    
    If lastRow < 2 Then Exit Function
    
    For r = 2 To lastRow
        categoryText = Trim$(CStr(dataWs.Cells(r, "K").Value))
        If categoryText = "Austrittsgrund fehlt" Or categoryText = "Austrittsgrund unbekannt" Or _
           categoryText = "Verlust erfahrener Mitarbeiter" Or _
           categoryText = "Austritt in den ersten 90 Tagen" Or _
           categoryText = "Wichtiger Austritt" Then
            FluktuationHasAlertRows = True
            Exit Function
        End If
    Next r
End Function


Private Function WriteFluktuationAlertIfMatch(ByVal ws As Worksheet, _
                                              ByVal dataWs As Worksheet, _
                                              ByVal sourceRow As Long, _
                                              ByVal targetRow As Long, _
                                              ByRef alertNum As Long, _
                                              ByVal incompleteOnly As Boolean) As Boolean
    Dim categoryText As String
    Dim monthName As String
    Dim employeeName As String
    Dim problemText As String
    Dim locationText As String
    Dim actionText As String
    Dim reasonText As String
    
    categoryText = Trim$(CStr(dataWs.Cells(sourceRow, "K").Value))
    
    If incompleteOnly Then
        If categoryText <> "Austrittsgrund fehlt" And categoryText <> "Austrittsgrund unbekannt" Then
            WriteFluktuationAlertIfMatch = False
            Exit Function
        End If
    Else
        If categoryText = "Austrittsgrund fehlt" Or categoryText = "Austrittsgrund unbekannt" Then
            WriteFluktuationAlertIfMatch = False
            Exit Function
        End If
        If categoryText <> "Verlust erfahrener Mitarbeiter" And _
           categoryText <> "Austritt in den ersten 90 Tagen" And _
           categoryText <> "Wichtiger Austritt" Then
            WriteFluktuationAlertIfMatch = False
            Exit Function
        End If
    End If
    
    monthName = Trim$(CStr(dataWs.Cells(sourceRow, "A").Value))
    employeeName = Trim$(CStr(dataWs.Cells(sourceRow, "C").Value))
    reasonText = Trim$(CStr(dataWs.Cells(sourceRow, "F").Value))
    
    Select Case categoryText
        Case "Austrittsgrund fehlt"
            problemText = "Austrittsgrund fehlt"
            actionText = "Monat oeffnen, Mitarbeiter finden, in Spalte N den Austrittsgrund aus der Liste waehlen."
        Case "Austrittsgrund unbekannt"
            problemText = "Austrittsgrund unbekannt"
            actionText = "Spalte N pruefen und einen gueltigen Grund aus EINSTELLUNG / Dropdown waehlen."
        Case "Verlust erfahrener Mitarbeiter"
            problemText = "Erfahrener Mitarbeiter ausgeschieden"
            actionText = "Exit-Gespraech auswerten, Wissensuebergabe und Nachbesetzung planen."
        Case "Austritt in den ersten 90 Tagen"
            problemText = "Austritt in den ersten 90 Tagen"
            actionText = "Onboarding, Training und erste Dienstplaene fuer neue Mitarbeiter pruefen."
        Case "Wichtiger Austritt"
            problemText = "Wichtiger Austritt"
            actionText = "Austrittsgrund und Teamsituation mit der Fuehrung besprechen."
        Case Else
            problemText = categoryText
            actionText = "Fall im Monatsblatt pruefen."
    End Select
    
    If employeeName = "" Then employeeName = "Unbekannter Name"
    locationText = "Monat " & monthName & ", Name: " & employeeName
    
    If reasonText <> "" And categoryText <> "Austrittsgrund fehlt" Then
        locationText = locationText & ", Grund: " & reasonText
    End If
    
    alertNum = alertNum + 1
    ws.Cells(targetRow, 1).Value = alertNum
    ws.Cells(targetRow, 2).Value = problemText
    ws.Cells(targetRow, 3).Value = locationText
    ws.Cells(targetRow, 4).Value = actionText
    ws.Range(ws.Cells(targetRow, 4), ws.Cells(targetRow, 5)).Merge
    
    WriteFluktuationAlertIfMatch = True
End Function


Public Function WriteFluktuationRecommendationsSection(ByVal ws As Worksheet, _
                                                       ByVal recommendationItems As Variant, _
                                                       ByVal startRow As Long) As Long
    Dim i As Long
    Dim outRow As Long
    Dim itemText As String
    
    ws.Range("A" & startRow).Value = "Empfehlungen"
    outRow = startRow + 1
    
    If IsArray(recommendationItems) Then
        For i = LBound(recommendationItems) To UBound(recommendationItems)
            itemText = Trim$(CStr(recommendationItems(i)))
            If itemText <> "" Then
                ws.Cells(outRow, 1).Value = i & "."
                ws.Cells(outRow, 2).Value = itemText
                ws.Range(ws.Cells(outRow, 2), ws.Cells(outRow, 5)).Merge
                outRow = outRow + 1
            End If
        Next i
    Else
        ws.Cells(outRow, 1).Value = "1."
        ws.Cells(outRow, 2).Value = "Entwicklung beobachten und Monatsblaetter aktuell halten."
        ws.Range(ws.Cells(outRow, 2), ws.Cells(outRow, 5)).Merge
        outRow = outRow + 1
    End If
    
    WriteFluktuationRecommendationsSection = outRow - 1
End Function


Public Sub PID_ClearFluktuationCharts(ByVal ws As Worksheet)
    Dim chartObj As ChartObject
    
    On Error Resume Next
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
    On Error GoTo 0
End Sub


Public Sub BuildFluktuationCharts(ByVal ws As Worksheet, _
                                  ByVal dataWs As Worksheet, _
                                  ByVal chartRow As Long, _
                                  ByVal monthNames As Variant, _
                                  ByRef monthExit() As Long, _
                                  ByVal earlyExits As Long, _
                                  ByVal experiencedLoss As Long, _
                                  ByVal importantExits As Long, _
                                  ByVal neutralExits As Long, _
                                  ByVal normalExits As Long, _
                                  ByVal incompleteExits As Long)
    Dim monthDataRow As Long
    Dim categoryDataRow As Long
    Dim reasonDataRow As Long
    Dim monthCount As Long
    Dim categoryCount As Long
    Dim reasonCount As Long
    Dim categoryLastRow As Long
    Dim monthLastRow As Long
    Dim reasonLastRow As Long
    Dim i As Long
    Dim chartLeft As Double
    Dim chartTop As Double
    Dim chartTopRow2 As Double
    Dim monthChart As ChartObject
    Dim categoryChart As ChartObject
    Dim reasonChart As ChartObject
    Dim valueRange As Range
    Dim labelRange As Range
    Dim chartWidth As Double
    Dim chartHeight As Double
    
    On Error GoTo ChartFail
    
    chartWidth = 300
    chartHeight = 210
    monthDataRow = 2
    categoryDataRow = 20
    reasonDataRow = 40
    
    ws.Range("P:Q").ClearContents
    ws.Range("R:S").ClearContents
    ws.Cells(1, 16).Value = "Monat"
    ws.Cells(1, 17).Value = "Austritte"
    
    monthCount = 0
    For i = 1 To 12
        If monthExit(i) > 0 Then
            monthCount = monthCount + 1
            ws.Cells(monthDataRow + monthCount - 1, 16).Value = monthNames(i - 1)
            ws.Cells(monthDataRow + monthCount - 1, 17).Value = monthExit(i)
        End If
    Next i
    
    ws.Cells(categoryDataRow, 16).Value = "Kategorie"
    ws.Cells(categoryDataRow, 17).Value = "Anzahl"
    
    categoryCount = 0
    If earlyExits > 0 Then
        categoryCount = categoryCount + 1
        ws.Cells(categoryDataRow + categoryCount, 16).Value = "Erste 90 Tage"
        ws.Cells(categoryDataRow + categoryCount, 17).Value = earlyExits
    End If
    If experiencedLoss > 0 Then
        categoryCount = categoryCount + 1
        ws.Cells(categoryDataRow + categoryCount, 16).Value = "Erfahrene MA"
        ws.Cells(categoryDataRow + categoryCount, 17).Value = experiencedLoss
    End If
    If importantExits > 0 Then
        categoryCount = categoryCount + 1
        ws.Cells(categoryDataRow + categoryCount, 16).Value = "Wichtige Austritte"
        ws.Cells(categoryDataRow + categoryCount, 17).Value = importantExits
    End If
    If normalExits > 0 Then
        categoryCount = categoryCount + 1
        ws.Cells(categoryDataRow + categoryCount, 16).Value = "Normal"
        ws.Cells(categoryDataRow + categoryCount, 17).Value = normalExits
    End If
    If neutralExits > 0 Then
        categoryCount = categoryCount + 1
        ws.Cells(categoryDataRow + categoryCount, 16).Value = "Neutral"
        ws.Cells(categoryDataRow + categoryCount, 17).Value = neutralExits
    End If
    If incompleteExits > 0 Then
        categoryCount = categoryCount + 1
        ws.Cells(categoryDataRow + categoryCount, 16).Value = "Daten offen"
        ws.Cells(categoryDataRow + categoryCount, 17).Value = incompleteExits
    End If
    
    reasonCount = WriteFluktuationReasonChartData(dataWs, ws, reasonDataRow)
    
    ws.Range("A" & chartRow).Value = "Diagramme"
    ws.Range("A" & chartRow).Font.Bold = True
    ws.Range("A" & chartRow).Font.Size = 13
    ws.Range("A" & chartRow & ":E" & chartRow).Merge
    
    chartLeft = ws.Range("A" & (chartRow + 1)).Left
    chartTop = ws.Range("A" & (chartRow + 1)).Top
    chartTopRow2 = chartTop + chartHeight + 12
    
    If monthCount > 0 Then
        monthLastRow = monthDataRow + monthCount - 1
        Set labelRange = ws.Range(ws.Cells(monthDataRow, 16), ws.Cells(monthLastRow, 16))
        Set valueRange = ws.Range(ws.Cells(monthDataRow, 17), ws.Cells(monthLastRow, 17))
        
        Set monthChart = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=chartWidth, Height:=chartHeight)
        With monthChart.Chart
            .ChartType = xlColumnClustered
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            With .SeriesCollection.NewSeries
                .Name = "Austritte"
                .Values = valueRange
                .XValues = labelRange
            End With
            .HasTitle = True
            .ChartTitle.Text = "Austritte pro Monat"
            On Error Resume Next
            .Legend.Delete
            On Error GoTo ChartFail
        End With
    End If
    
    If categoryCount > 0 Then
        categoryLastRow = categoryDataRow + categoryCount
        Set labelRange = ws.Range(ws.Cells(categoryDataRow + 1, 16), ws.Cells(categoryLastRow, 16))
        Set valueRange = ws.Range(ws.Cells(categoryDataRow + 1, 17), ws.Cells(categoryLastRow, 17))
        
        Set categoryChart = ws.ChartObjects.Add(Left:=chartLeft + chartWidth + 16, Top:=chartTop, Width:=chartWidth, Height:=chartHeight)
        With categoryChart.Chart
            .ChartType = xlPie
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            With .SeriesCollection.NewSeries
                .Name = "Kategorien"
                .Values = valueRange
                .XValues = labelRange
            End With
            .HasTitle = True
            .ChartTitle.Text = "Austritte nach Kategorie"
        End With
    End If
    
    If reasonCount > 0 Then
        reasonLastRow = reasonDataRow + reasonCount
        Set labelRange = ws.Range(ws.Cells(reasonDataRow + 1, 18), ws.Cells(reasonLastRow, 18))
        Set valueRange = ws.Range(ws.Cells(reasonDataRow + 1, 19), ws.Cells(reasonLastRow, 19))
        
        Set reasonChart = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTopRow2, Width:=(chartWidth * 2) + 16, Height:=chartHeight)
        With reasonChart.Chart
            .ChartType = xlBarClustered
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            With .SeriesCollection.NewSeries
                .Name = "Austritte"
                .Values = valueRange
                .XValues = labelRange
            End With
            .HasTitle = True
            .ChartTitle.Text = "Austritte nach Austrittsgrund"
            On Error Resume Next
            .Legend.Delete
            On Error GoTo ChartFail
        End With
    End If
    
    Exit Sub

ChartFail:
End Sub


Private Function WriteFluktuationReasonChartData(ByVal dataWs As Worksheet, _
                                                 ByVal ws As Worksheet, _
                                                 ByVal reasonDataRow As Long) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim reasonText As String
    Dim categoryText As String
    Dim summaryText As String
    Dim parts As Variant
    Dim i As Long
    Dim partText As String
    Dim baseReason As String
    Dim countText As String
    Dim countValue As Long
    Dim outRow As Long
    
    WriteFluktuationReasonChartData = 0
    summaryText = ""
    
    lastRow = dataWs.Cells(dataWs.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    For r = 2 To lastRow
        reasonText = Trim$(CStr(dataWs.Cells(r, "F").Value))
        categoryText = Trim$(CStr(dataWs.Cells(r, "K").Value))
        
        If categoryText = "Austrittsgrund fehlt" Then
            reasonText = "Grund fehlt"
        ElseIf categoryText = "Austrittsgrund unbekannt" Then
            reasonText = "Grund unbekannt"
        ElseIf reasonText = "" Then
            reasonText = "Grund fehlt"
        End If
        
        summaryText = AddReasonToSummary(summaryText, reasonText)
    Next r
    
    If summaryText = "" Then Exit Function
    
    ws.Cells(reasonDataRow, 18).Value = "Austrittsgrund"
    ws.Cells(reasonDataRow, 19).Value = "Anzahl"
    
    parts = Split(summaryText, ", ")
    outRow = reasonDataRow
    
    For i = LBound(parts) To UBound(parts)
        partText = Trim$(CStr(parts(i)))
        baseReason = partText
        countValue = 1
        
        If InStr(partText, " (") > 0 Then
            baseReason = Left$(partText, InStr(partText, " (") - 1)
            countText = Replace(partText, baseReason & " (", "")
            countText = Replace(countText, ")", "")
            If IsNumeric(countText) Then countValue = CLng(countText)
        End If
        
        outRow = outRow + 1
        ws.Cells(outRow, 18).Value = baseReason
        ws.Cells(outRow, 19).Value = countValue
    Next i
    
    WriteFluktuationReasonChartData = outRow - reasonDataRow
End Function


Private Sub PID_WriteFluktuationExplanationRows(ByVal ws As Worksheet, ByVal startRow As Long)
    Dim i As Long
    
    ws.Range("A" & startRow).Value = "Kurz erklaert"
    
    ws.Range("A" & startRow + 1).Value = "Aktuelle Jahresfluktuation"
    ws.Range("B" & startRow + 1).Value = "Zeigt die Fluktuation vom Jahresbeginn bis zum aktuellen Auswertungsmonat. Der Wert wird nicht auf das ganze Jahr hochgerechnet."
    
    ws.Range("A" & startRow + 2).Value = "Verlust-Score"
    ws.Range("B" & startRow + 2).Value = "Der Verlust-Score zeigt, wie schwer ein Austritt fuer das Restaurant bewertet wird. Er besteht aus Austrittsgrund und Dauer der Betriebszugehoerigkeit."
    
    ws.Range("A" & startRow + 3).Value = "Durchschnittlicher Verlust-Score"
    ws.Range("B" & startRow + 3).Value = "Durchschnittlicher Verlust-Score pro Austritt. Je hoeher der Wert, desto schwerer wiegen die Austritte im Durchschnitt."
    
    ws.Range("A" & startRow + 4).Value = "Austritt in den ersten 90 Tagen"
    ws.Range("B" & startRow + 4).Value = "Ein Austritt kurz nach Eintritt. Das kann auf Recruiting, Onboarding, Training oder erste Dienstplaene hinweisen."
    
    ws.Range("A" & startRow + 5).Value = "Verlust erfahrener Mitarbeiter"
    ws.Range("B" & startRow + 5).Value = "Ein erfahrener Mitarbeiter verlaesst das Restaurant. Das bedeutet meist Wissensverlust, Stabilitaetsverlust und hoeheren Nachbesetzungsaufwand."
    
    ws.Range("A" & startRow + 6).Value = "Wichtiger Austritt"
    ws.Range("B" & startRow + 6).Value = "Ein Austritt mit erhoehtem Verlust-Score. Die Bewertung entsteht automatisch aus Austrittsgrund und Betriebszugehoerigkeit."
    
    ws.Range("A" & startRow + 7).Value = "Neutrale Bewegung"
    ws.Range("B" & startRow + 7).Value = "Storetransfer, Befoerderung, Karenz oder Nicht eingetreten werden nicht negativ bewertet."
    
    ws.Range("A" & startRow + 8).Value = "Unvollstaendiger Austritt"
    ws.Range("B" & startRow + 8).Value = "Der Austrittsgrund fehlt oder ist unbekannt. Diese Austritte muessen geprueft und ergaenzt werden."
    
    For i = startRow + 1 To startRow + 8
        ws.Range("B" & i & ":E" & i).Merge
    Next i
End Sub

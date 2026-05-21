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
    
    Dim ytdMonthLimit As Long
    Dim ytdExits As Long
    Dim ytdPersonalSum As Long
    Dim ytdPersonalMonths As Long
    Dim ytdAveragePersonal As Double
    Dim ytdFluctuation As Double
    
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
        .Cells.VerticalAlignment = xlBottom
        .Cells.WrapText = False
        .Rows.RowHeight = 15
    End With
    
    ' --- Abschnitt 1: Daten aus FLUKTUATION_DATEN aggregieren ---
    lastRow = dataWs.Cells(dataWs.Rows.count, "A").End(xlUp).Row
    
    If lastRow >= 2 Then
        For r = 2 To lastRow
            totalExits = totalExits + 1
            
            monthName = Trim$(CStr(dataWs.Cells(r, "A").Value))
            monthIndex = GetMonthIndexFromName(monthName)
            
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
    For i = 1 To 12
        monthPersonalEnde(i) = GetPersonalEndeForMonth(CStr(monthNames(i - 1)), currentYear, i)
        
        If monthPersonalEnde(i) > 0 Then
            monthFluctuation(i) = monthExit(i) / monthPersonalEnde(i)
        Else
            monthFluctuation(i) = 0
        End If
    Next i
    
    If currentYear = Year(Date) Then
        ytdMonthLimit = Month(Date)
    ElseIf currentYear < Year(Date) Then
        ytdMonthLimit = 12
    Else
        ytdMonthLimit = GetLastMonthWithExit(monthExit)
    End If
    
    ytdExits = 0
    ytdPersonalSum = 0
    ytdPersonalMonths = 0
    
    If ytdMonthLimit > 0 Then
        For i = 1 To ytdMonthLimit
            ytdExits = ytdExits + monthExit(i)
            
            If monthPersonalEnde(i) > 0 Then
                ytdPersonalSum = ytdPersonalSum + monthPersonalEnde(i)
                ytdPersonalMonths = ytdPersonalMonths + 1
            End If
        Next i
    End If
    
    If ytdPersonalMonths > 0 Then
        ytdAveragePersonal = ytdPersonalSum / ytdPersonalMonths
    Else
        ytdAveragePersonal = 0
    End If
    
    If ytdAveragePersonal > 0 Then
        ytdFluctuation = ytdExits / ytdAveragePersonal
    Else
        ytdFluctuation = 0
    End If
    
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
        BuildFluktuationCharts analyseWs, chartRow, monthNames, monthExit, earlyExits, experiencedLoss, importantExits, neutralExits, normalExits, incompleteExits
        
        ' --- Abschnitt 4b: Monatstabelle ---
        monthlyTitleRow = chartRow + 14
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
    With ws
        .Range("A1:E1").Merge
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A2:D2").Font.Bold = True
        .Range("D2").HorizontalAlignment = xlLeft
        
        .Range("A4:E4").Merge
        .Range("A4").Font.Bold = True
        .Range("A4").Font.Size = 14
        
        .Range("A" & statusRow & ":E" & statusRow + 1).Borders.LineStyle = xlContinuous
        .Range("A" & statusRow & ":E" & statusRow + 1).Borders.Weight = xlMedium
        .Range("A" & statusRow).Font.Bold = True
        .Range("A" & statusRow).HorizontalAlignment = xlCenter
        .Range("B" & statusRow).Font.Size = 14
        ApplyRiskFormatting .Range("B" & statusRow), riskLevel
        .Range("C" & statusRow).WrapText = True
        .Range("C" & statusRow).VerticalAlignment = xlTop
        .Range("C" & statusRow).HorizontalAlignment = xlLeft
        
        .Range("A" & kpiLabelRow & ":E" & kpiValueRow).Font.Bold = True
        .Range("A" & kpiLabelRow & ":E" & kpiValueRow).Borders.LineStyle = xlContinuous
        .Range("A" & kpiLabelRow & ":E" & kpiValueRow).Borders.Weight = xlThin
        .Range("A" & kpiLabelRow & ":E" & kpiValueRow).HorizontalAlignment = xlCenter
        .Range("B" & kpiValueRow).NumberFormat = "0.00%"
        .Range("C" & kpiValueRow).NumberFormat = "0.00"
        
        If incompleteExitsCellNeedsHighlight(ws, kpiValueRow) Then
            .Range("E" & kpiValueRow).Interior.Color = RGB(221, 235, 247)
            .Range("E" & kpiValueRow).Font.Color = RGB(31, 78, 121)
        End If
        
        .Range("A" & alertsHeaderRow).Font.Bold = True
        .Range("A" & alertsHeaderRow).Font.Size = 13
        If alertsEndRow > alertsHeaderRow + 2 Then
            .Range(.Cells(alertsHeaderRow + 1, 1), .Cells(alertsHeaderRow + 1, 5)).Font.Bold = True
            .Range(.Cells(alertsHeaderRow + 1, 1), .Cells(alertsEndRow - 1, 5)).Borders.LineStyle = xlContinuous
            .Range(.Cells(alertsHeaderRow + 1, 1), .Cells(alertsEndRow - 1, 5)).Borders.Weight = xlThin
            .Range(.Cells(alertsHeaderRow + 2, 1), .Cells(alertsEndRow - 1, 1)).Font.Bold = True
            .Range(.Cells(alertsHeaderRow + 2, 1), .Cells(alertsEndRow - 1, 1)).HorizontalAlignment = xlCenter
            .Range(.Cells(alertsHeaderRow + 2, 2), .Cells(alertsEndRow - 1, 5)).WrapText = True
            .Range(.Cells(alertsHeaderRow + 2, 2), .Cells(alertsEndRow - 1, 5)).VerticalAlignment = xlTop
            .Range(.Cells(alertsHeaderRow + 2, 2), .Cells(alertsEndRow - 1, 5)).HorizontalAlignment = xlLeft
        ElseIf alertsEndRow = alertsHeaderRow + 2 Then
            .Range(.Cells(alertsHeaderRow + 1, 1), .Cells(alertsHeaderRow + 1, 5)).WrapText = True
        End If
        
        .Range("A" & recHeaderRow).Font.Bold = True
        .Range("A" & recHeaderRow).Font.Size = 13
        If recEndRow > recHeaderRow Then
            .Range(.Cells(recHeaderRow + 1, 1), .Cells(recEndRow, 5)).Borders.LineStyle = xlContinuous
            .Range(.Cells(recHeaderRow + 1, 1), .Cells(recEndRow, 5)).Borders.Weight = xlThin
            .Range(.Cells(recHeaderRow + 1, 2), .Cells(recEndRow, 5)).WrapText = True
            .Range(.Cells(recHeaderRow + 1, 2), .Cells(recEndRow, 5)).VerticalAlignment = xlTop
            .Range(.Cells(recHeaderRow + 1, 2), .Cells(recEndRow, 5)).HorizontalAlignment = xlLeft
            .Range(.Cells(recHeaderRow + 1, 1), .Cells(recEndRow, 1)).Font.Bold = True
            .Range(.Cells(recHeaderRow + 1, 1), .Cells(recEndRow, 1)).HorizontalAlignment = xlCenter
        End If
        
        .Range(.Cells(monthlyTitleRow, 1), .Cells(monthlyTitleRow, lastTableCol + 2)).Merge
        .Cells(monthlyTitleRow, 1).Font.Size = 16
        .Cells(monthlyTitleRow, 1).Font.Bold = True
        .Cells(monthlyTitleRow, 1).HorizontalAlignment = xlCenter
        
        .Range(.Cells(headerRow, 1), .Cells(headerRow, lastTableCol + 2)).Font.Bold = True
        .Range(.Cells(headerRow, 1), .Cells(outputRow - 1, lastTableCol + 2)).Borders.LineStyle = xlContinuous
        .Range(.Cells(headerRow, 1), .Cells(outputRow - 1, lastTableCol + 2)).Borders.Weight = xlThin
        
        .Range(.Cells(headerRow, 1), .Cells(outputRow - 1, lastTableCol + 2)).HorizontalAlignment = xlCenter
        .Range(.Cells(headerRow, 1), .Cells(outputRow - 1, lastTableCol + 2)).VerticalAlignment = xlCenter
        
        .Range(.Cells(firstDataRow, 1), .Cells(outputRow - 1, lastTableCol + 2)).WrapText = True
        
        If lastTableCol >= 2 Then
            .Range(.Cells(firstDataRow, lastTableCol - 1), .Cells(outputRow - 1, lastTableCol + 2)).HorizontalAlignment = xlLeft
        End If
        
        .Range(.Cells(firstDataRow, 3), .Cells(outputRow - 1, 3)).NumberFormat = "0.00%"
        .Range(.Cells(firstDataRow, 4), .Cells(outputRow - 1, 5)).NumberFormat = "0.00"
        
        .Range("A" & explanationStartRow & ":E" & explanationStartRow).Merge
        .Range("A" & explanationStartRow).Font.Bold = True
        .Range("A" & explanationStartRow).Font.Size = 13
        
        .Range("A" & explanationStartRow & ":E" & explanationStartRow + 8).Borders.LineStyle = xlContinuous
        .Range("A" & explanationStartRow & ":E" & explanationStartRow + 8).Borders.Weight = xlThin
        
        .Range("A" & explanationStartRow + 1 & ":A" & explanationStartRow + 8).Font.Bold = True
        .Range("B" & explanationStartRow + 1 & ":E" & explanationStartRow + 8).WrapText = True
        
        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 24
        .Columns("D").ColumnWidth = 24
        .Columns("E").ColumnWidth = 28
        .Columns("F").ColumnWidth = 18
        .Columns("G").ColumnWidth = 18
        .Columns("H").ColumnWidth = 22
        .Columns("I").ColumnWidth = 36
        .Columns("J").ColumnWidth = 36
        .Columns("K").ColumnWidth = 36
        .Columns("L").ColumnWidth = 36
        .Columns("P:Q").Hidden = True
        
        .Rows(statusRow & ":" & statusRow + 1).RowHeight = 36
        .Rows(kpiLabelRow & ":" & kpiValueRow).RowHeight = 22
        If alertsEndRow > alertsHeaderRow + 2 Then
            .Rows(alertsHeaderRow + 2 & ":" & alertsEndRow - 1).RowHeight = 34
        ElseIf alertsEndRow = alertsHeaderRow + 2 Then
            .Rows(alertsHeaderRow + 1).RowHeight = 34
        End If
        If recEndRow > recHeaderRow Then
            .Rows(recHeaderRow + 1 & ":" & recEndRow).RowHeight = 30
        End If
        .Rows(firstDataRow & ":" & outputRow - 1).RowHeight = 42
        .Rows(explanationStartRow + 1 & ":" & explanationStartRow + 8).RowHeight = 38
        
        .Range("A1:L" & explanationStartRow + 8).VerticalAlignment = xlCenter
        .Range("C" & statusRow).VerticalAlignment = xlTop
    End With
End Sub


Private Function incompleteExitsCellNeedsHighlight(ByVal ws As Worksheet, ByVal kpiValueRow As Long) As Boolean
    If IsNumeric(ws.Cells(kpiValueRow, 5).Value) Then
        incompleteExitsCellNeedsHighlight = (CLng(ws.Cells(kpiValueRow, 5).Value) > 0)
    Else
        incompleteExitsCellNeedsHighlight = False
    End If
End Function


Public Function GetMonthIndexFromName(ByVal monthName As String) As Long
    Select Case UCase$(Trim$(CStr(monthName)))
        Case "JANUAR"
            GetMonthIndexFromName = 1
        Case "FEBRUAR"
            GetMonthIndexFromName = 2
        Case "MARZ", "MÄRZ", "MAERZ"
            GetMonthIndexFromName = 3
        Case "APRIL"
            GetMonthIndexFromName = 4
        Case "MAI"
            GetMonthIndexFromName = 5
        Case "JUNI"
            GetMonthIndexFromName = 6
        Case "JULI"
            GetMonthIndexFromName = 7
        Case "AUGUST"
            GetMonthIndexFromName = 8
        Case "SEPTEMBER"
            GetMonthIndexFromName = 9
        Case "OKTOBER"
            GetMonthIndexFromName = 10
        Case "NOVEMBER"
            GetMonthIndexFromName = 11
        Case "DEZEMBER"
            GetMonthIndexFromName = 12
        Case Else
            GetMonthIndexFromName = 0
    End Select
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


Public Function GetFluktuationFocusText(ByVal totalExits As Long, _
                                        ByVal neutralExits As Long, _
                                        ByVal earlyExits As Long, _
                                        ByVal experiencedLoss As Long, _
                                        ByVal importantExits As Long, _
                                        ByVal incompleteExits As Long, _
                                        ByVal totalLoss As Double) As String
    If incompleteExits > 0 Then
        GetFluktuationFocusText = "Es gibt Austritte mit fehlendem oder unbekanntem Austrittsgrund. Bitte zuerst die Daten ergaenzen, damit die Analyse korrekt bewertet werden kann."
    ElseIf totalExits = 0 Then
        GetFluktuationFocusText = "Keine relevanten Austritte erfasst. Die Fluktuation ist aktuell unauffaellig."
    ElseIf neutralExits = totalExits Then
        GetFluktuationFocusText = "Die erfassten Bewegungen sind neutral oder positiv. Kein unmittelbarer Handlungsbedarf aus Fluktuationssicht."
    ElseIf experiencedLoss >= 2 Then
        GetFluktuationFocusText = "Mehrere erfahrene Mitarbeiter sind ausgeschieden. Fokus: Mitarbeiterbindung, Fuehrung, Entwicklungsgespraeche und Teamstabilitaet."
    ElseIf earlyExits >= 2 Then
        GetFluktuationFocusText = "Mehrere Austritte passieren in den ersten 90 Tagen. Fokus: Recruiting, Onboarding, Training und Qualitaet der ersten Dienstplaene."
    ElseIf importantExits >= 2 Then
        GetFluktuationFocusText = "Es gibt mehrere wichtige Austritte. Fokus: Arbeitsklima, Fuehrung, Kommunikation und Mitarbeitergespraeche."
    ElseIf totalLoss >= 3 Then
        GetFluktuationFocusText = "Der Verlust-Score ist erhoeht. Austrittsgruende pruefen und gezielte Massnahmen zur Mitarbeiterbindung ableiten."
    Else
        GetFluktuationFocusText = "Die Fluktuation ist vorhanden, aber aktuell nicht kritisch. Entwicklung weiter beobachten."
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
    Dim personalID As Variant
    Dim employeeName As String
    Dim exitDate As Variant
    Dim sheetRow As Long
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
    personalID = dataWs.Cells(sourceRow, "B").Value
    employeeName = Trim$(CStr(dataWs.Cells(sourceRow, "C").Value))
    exitDate = dataWs.Cells(sourceRow, "E").Value
    reasonText = Trim$(CStr(dataWs.Cells(sourceRow, "F").Value))
    sheetRow = FindEmployeeRowOnMonthSheet(monthName, personalID, employeeName, exitDate)
    
    Select Case categoryText
        Case "Austrittsgrund fehlt"
            problemText = "Austrittsgrund fehlt"
            actionText = "Monatsblatt oeffnen, Zeile finden, in Spalte N den Austrittsgrund aus der Liste waehlen."
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
    If sheetRow > 0 Then
        locationText = "Blatt """ & monthName & """, Zeile " & sheetRow & ", Name: " & employeeName
    Else
        locationText = "Blatt """ & monthName & """, Name: " & employeeName
    End If
    
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


Public Function FindEmployeeRowOnMonthSheet(ByVal monthSheetName As String, _
                                            ByVal personalID As Variant, _
                                            ByVal employeeName As String, _
                                            ByVal exitDate As Variant) As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim rowID As String
    Dim rowName As String
    
    On Error GoTo SafeExit
    
    Set ws = ThisWorkbook.Worksheets(monthSheetName)
    
    For r = 3 To 82
        rowID = Trim$(CStr(ws.Cells(r, "B").Value))
        rowName = Trim$(CStr(ws.Cells(r, "C").Value))
        
        If Trim$(CStr(personalID)) <> "" Then
            If rowID = Trim$(CStr(personalID)) Then
                If IsDate(exitDate) And IsDate(ws.Cells(r, "I").Value) Then
                    If CDate(ws.Cells(r, "I").Value) = CDate(exitDate) Then
                        FindEmployeeRowOnMonthSheet = r
                        Exit Function
                    End If
                Else
                    FindEmployeeRowOnMonthSheet = r
                    Exit Function
                End If
            End If
        End If
        
        If employeeName <> "" Then
            If UCase$(rowName) = UCase$(employeeName) Then
                If IsDate(exitDate) And IsDate(ws.Cells(r, "I").Value) Then
                    If CDate(ws.Cells(r, "I").Value) = CDate(exitDate) Then
                        FindEmployeeRowOnMonthSheet = r
                        Exit Function
                    End If
                ElseIf Not IsDate(exitDate) Then
                    FindEmployeeRowOnMonthSheet = r
                    Exit Function
                End If
            End If
        End If
    Next r

SafeExit:
    FindEmployeeRowOnMonthSheet = 0
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
    Dim monthCount As Long
    Dim categoryCount As Long
    Dim i As Long
    Dim chartLeft As Double
    Dim chartTop As Double
    Dim monthChart As ChartObject
    Dim categoryChart As ChartObject
    Dim monthRange As Range
    Dim categoryRange As Range
    
    monthDataRow = 2
    categoryDataRow = 20
    
    ws.Range("P:Q").ClearContents
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
    
    chartLeft = ws.Range("A" & chartRow).Left
    chartTop = ws.Range("A" & chartRow).Top
    
    ws.Range("A" & chartRow).Value = "Diagramme"
    ws.Range("A" & chartRow).Font.Bold = True
    ws.Range("A" & chartRow).Font.Size = 13
    
    If monthCount > 0 Then
        Set monthRange = ws.Range(ws.Cells(monthDataRow, 16), ws.Cells(monthDataRow + monthCount - 1, 17))
        Set monthChart = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop + 18, Width:=320, Height:=210)
        With monthChart.Chart
            .ChartType = xlColumnClustered
            .SetSourceData Source:=monthRange, PlotBy:=xlColumns
            .HasTitle = True
            .ChartTitle.Text = "Austritte pro Monat"
            .Legend.Delete
        End With
    End If
    
    If categoryCount > 0 Then
        Set categoryRange = ws.Range(ws.Cells(categoryDataRow + 1, 16), ws.Cells(categoryDataRow + categoryCount, 17))
        Set categoryChart = ws.ChartObjects.Add(Left:=chartLeft + 340, Top:=chartTop + 18, Width:=320, Height:=210)
        With categoryChart.Chart
            .ChartType = xlPie
            .SetSourceData Source:=categoryRange, PlotBy:=xlColumns
            .HasTitle = True
            .ChartTitle.Text = "Austritte nach Kategorie"
        End With
    End If
End Sub


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

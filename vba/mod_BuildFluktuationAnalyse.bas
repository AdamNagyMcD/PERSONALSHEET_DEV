Attribute VB_Name = "mod_BuildFluktuationAnalyse"
Option Explicit

Public Sub BuildFluktuationAnalyse()
    Dim dataWs As Worksheet
    Dim analyseWs As Worksheet
    Dim lohnWs As Worksheet
    
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
    Dim focusText As String
    Dim explanationText As String
    Dim yearValue As Variant
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
    Set analyseWs = ThisWorkbook.Worksheets("Fluktuation")
    Set lohnWs = ThisWorkbook.Worksheets("LOHNTABELLE")
    
    yearValue = lohnWs.Range("G3").Value
    
    If IsNumeric(yearValue) Then
        currentYear = CLng(yearValue)
    Else
        currentYear = Year(Date)
    End If
    
    On Error Resume Next
    analyseWs.Unprotect Password:="company"
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
    
    riskLevel = GetFluktuationRiskLevel(totalLoss, totalExits, earlyExits, experiencedLoss, importantExits, incompleteExits)
    focusText = GetFluktuationFocusText(totalExits, neutralExits, earlyExits, experiencedLoss, importantExits, incompleteExits, totalLoss)
    
    explanationText = "Diese Auswertung bewertet Austritte nicht nur nach Anzahl, sondern auch nach Art des Austritts und Dauer der Betriebszugehoerigkeit."
    explanationText = explanationText & " Ein Austritt eines langjaehrigen Mitarbeiters zaehlt staerker als ein Austritt in den ersten Wochen."
    explanationText = explanationText & " Neutrale Bewegungen wie Storetransfer, Befoerderung, Karenz oder Nicht eingetreten werden nicht negativ bewertet."
    
    With analyseWs
        .Range("A1").Value = "Fluktuation"
        .Range("A2").Value = "Jahresanalyse"
        .Range("B2").Value = yearValue
        
        .Range("A4").Value = "Was bedeutet diese Auswertung?"
        .Range("A5").Value = explanationText
        
        .Range("A8").Value = "Kennzahl"
        .Range("B8").Value = "Wert"
        .Range("C8").Value = "Einfache Erklaerung"
        
        .Range("A9").Value = "Austritte gesamt"
        .Range("B9").Value = totalExits
        .Range("C9").Value = "Alle erfassten Austritte im Jahr."
        
        .Range("A10").Value = "Aktuelle Jahresfluktuation"
        .Range("B10").Value = ytdFluctuation
        .Range("C10").Value = "Zeigt die bisherige Fluktuation im Jahr. Nicht auf das ganze Jahr hochgerechnet."
        
        .Range("A11").Value = "Verlust-Score"
        .Range("B11").Value = totalLoss
        .Range("C11").Value = "Bewertet, wie schwer die Austritte fuer das Restaurant sind."
        
        .Range("A12").Value = "Durchschnittlicher Verlust-Score"
        .Range("B12").Value = avgLoss
        .Range("C12").Value = "Durchschnittliche Schwere pro Austritt."
        
        .Range("A13").Value = "Austritte in den ersten 90 Tagen"
        .Range("B13").Value = earlyExits
        .Range("C13").Value = "Hinweis auf Recruiting, Onboarding oder Training."
        
        .Range("A14").Value = "Verlust erfahrener Mitarbeiter"
        .Range("B14").Value = experiencedLoss
        .Range("C14").Value = "Erfahrene Mitarbeiter verlassen das Restaurant."
        
        .Range("A15").Value = "Wichtige Austritte"
        .Range("B15").Value = importantExits
        .Range("C15").Value = "Austritte mit erhoehtem Risiko fuer den Betrieb."
        
        .Range("A16").Value = "Neutrale Bewegungen"
        .Range("B16").Value = neutralExits
        .Range("C16").Value = "Storetransfer, Befoerderung, Karenz oder Nicht eingetreten."
        
        .Range("A17").Value = "Unvollstaendige Austritte"
        .Range("B17").Value = incompleteExits
        .Range("C17").Value = "Austritte mit fehlendem oder unbekanntem Austrittsgrund."
        
        .Range("A19").Value = "Risiko-Einschaetzung"
        .Range("B19").Value = riskLevel
        
        .Range("A21").Value = "Empfehlung"
        .Range("B21").Value = focusText
        
        For i = 9 To 17
            .Range("C" & i & ":E" & i).Merge
        Next i
        
        .Range("A21:A22").Merge
        .Range("B21:E22").Merge
        
        monthlyTitleRow = 25
        headerRow = 27
        firstDataRow = 28
        
        .Range("A" & monthlyTitleRow).Value = "Monatsuebersicht"
        
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
                
                If monthExit(i) > 0 Then
                    avgMonthLoss = monthLoss(i) / monthExit(i)
                Else
                    avgMonthLoss = 0
                End If
                
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
        
        .Range("A" & explanationStartRow).Value = "Kurz erklaert"
        
        .Range("A" & explanationStartRow + 1).Value = "Aktuelle Jahresfluktuation"
        .Range("B" & explanationStartRow + 1).Value = "Zeigt die Fluktuation vom Jahresbeginn bis zum aktuellen Auswertungsmonat. Der Wert wird nicht auf das ganze Jahr hochgerechnet."
        
        .Range("A" & explanationStartRow + 2).Value = "Verlust-Score"
        .Range("B" & explanationStartRow + 2).Value = "Der Verlust-Score zeigt, wie schwer ein Austritt fuer das Restaurant bewertet wird. Er besteht aus Austrittsgrund und Dauer der Betriebszugehoerigkeit."
        
        .Range("A" & explanationStartRow + 3).Value = "Durchschnittlicher Verlust-Score"
        .Range("B" & explanationStartRow + 3).Value = "Durchschnittlicher Verlust-Score pro Austritt. Je hoeher der Wert, desto schwerer wiegen die Austritte im Durchschnitt."
        
        .Range("A" & explanationStartRow + 4).Value = "Austritt in den ersten 90 Tagen"
        .Range("B" & explanationStartRow + 4).Value = "Ein Austritt kurz nach Eintritt. Das kann auf Recruiting, Onboarding, Training oder erste Dienstplaene hinweisen."
        
        .Range("A" & explanationStartRow + 5).Value = "Verlust erfahrener Mitarbeiter"
        .Range("B" & explanationStartRow + 5).Value = "Ein erfahrener Mitarbeiter verlaesst das Restaurant. Das bedeutet meist Wissensverlust, Stabilitaetsverlust und hoeheren Nachbesetzungsaufwand."
        
        .Range("A" & explanationStartRow + 6).Value = "Wichtiger Austritt"
        .Range("B" & explanationStartRow + 6).Value = "Ein Austritt mit erhoehtem Verlust-Score. Die Bewertung entsteht automatisch aus Austrittsgrund und Betriebszugehoerigkeit."
        
        .Range("A" & explanationStartRow + 7).Value = "Neutrale Bewegung"
        .Range("B" & explanationStartRow + 7).Value = "Storetransfer, Befoerderung, Karenz oder Nicht eingetreten werden nicht negativ bewertet."
        
        .Range("A" & explanationStartRow + 8).Value = "Unvollstaendiger Austritt"
        .Range("B" & explanationStartRow + 8).Value = "Der Austrittsgrund fehlt oder ist unbekannt. Diese Austritte muessen geprueft und ergaenzt werden."
        
        For i = explanationStartRow + 1 To explanationStartRow + 8
            .Range("B" & i & ":E" & i).Merge
        Next i
        
        FormatFluktuationSheet analyseWs, monthlyTitleRow, headerRow, firstDataRow, outputRow, lastTableCol, explanationStartRow, riskLevel
        
        .Protect Password:="company", UserInterfaceOnly:=True
    End With

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not analyseWs Is Nothing Then
        analyseWs.Protect Password:="company", UserInterfaceOnly:=True
    End If
    
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler in BuildFluktuationAnalyse:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Fluktuation Analyse"
End Sub


Public Sub FormatFluktuationSheet(ByVal ws As Worksheet, _
                                  ByVal monthlyTitleRow As Long, _
                                  ByVal headerRow As Long, _
                                  ByVal firstDataRow As Long, _
                                  ByVal outputRow As Long, _
                                  ByVal lastTableCol As Long, _
                                  ByVal explanationStartRow As Long, _
                                  ByVal riskLevel As String)
    With ws
        .Range("A1:E1").Merge
        .Range("A1").Font.Size = 22
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A2:B2").Font.Bold = True
        
        .Range("A4:E4").Merge
        .Range("A4").Font.Bold = True
        .Range("A4").Font.Size = 13
        
        .Range("A5:E6").Merge
        .Range("A5").WrapText = True
        .Range("A5").VerticalAlignment = xlTop
        
        .Range("A8:E8").Font.Bold = True
        .Range("A8:E17").Borders.LineStyle = xlContinuous
        .Range("A8:E17").Borders.Weight = xlThin
        
        .Range("B9:B17").HorizontalAlignment = xlCenter
        .Range("B9:B17").VerticalAlignment = xlCenter
        
        .Range("C9:E17").HorizontalAlignment = xlLeft
        .Range("C9:E17").VerticalAlignment = xlCenter
        .Range("C9:E17").WrapText = True
        
        .Range("A19:B19").Font.Bold = True
        .Range("A19:B19").Font.Size = 13
        .Range("A19:B19").Borders.LineStyle = xlContinuous
        .Range("A19:B19").Borders.Weight = xlMedium
        .Range("A19:B19").HorizontalAlignment = xlCenter
        
        ApplyRiskFormatting .Range("B19"), riskLevel
        
        .Range("A21:E22").Borders.LineStyle = xlContinuous
        .Range("A21").Font.Bold = True
        .Range("B21:E22").WrapText = True
        .Range("B21:E22").VerticalAlignment = xlTop
        
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
        
        .Range("B10").NumberFormat = "0.00%"
        .Range("B11:B12").NumberFormat = "0.00"
        
        .Columns("A").ColumnWidth = 32
        .Columns("B").ColumnWidth = 16
        .Columns("C").ColumnWidth = 16
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 16
        .Columns("F").ColumnWidth = 18
        .Columns("G").ColumnWidth = 18
        .Columns("H").ColumnWidth = 22
        .Columns("I").ColumnWidth = 36
        .Columns("J").ColumnWidth = 36
        .Columns("K").ColumnWidth = 36
        .Columns("L").ColumnWidth = 36
        
        .Rows("5:6").RowHeight = 42
        .Rows("19:19").RowHeight = 24
        .Rows("21:22").RowHeight = 48
        .Rows(firstDataRow & ":" & outputRow - 1).RowHeight = 42
        .Rows(explanationStartRow + 1 & ":" & explanationStartRow + 8).RowHeight = 38
        
        .Range("A1:L" & explanationStartRow + 8).VerticalAlignment = xlCenter
    End With
End Sub


Public Function GetMonthIndexFromName(ByVal monthName As String) As Long
    Select Case Trim$(CStr(monthName))
        Case "Januar"
            GetMonthIndexFromName = 1
        Case "Februar"
            GetMonthIndexFromName = 2
        Case "Marz"
            GetMonthIndexFromName = 3
        Case "April"
            GetMonthIndexFromName = 4
        Case "Mai"
            GetMonthIndexFromName = 5
        Case "Juni"
            GetMonthIndexFromName = 6
        Case "Juli"
            GetMonthIndexFromName = 7
        Case "August"
            GetMonthIndexFromName = 8
        Case "September"
            GetMonthIndexFromName = 9
        Case "Oktober"
            GetMonthIndexFromName = 10
        Case "November"
            GetMonthIndexFromName = 11
        Case "Dezember"
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


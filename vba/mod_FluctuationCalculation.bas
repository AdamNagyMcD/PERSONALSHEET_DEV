Attribute VB_Name = "mod_FluctuationCalculation"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Private Const PID_FLUKTUATION_PERCENT_FORMAT As String = "0.00%"


Public Sub PID_CalculateFluctuation(ByVal ws As Worksheet)
    Dim monthNumber As Long
    Dim currentYear As Long
    Dim monthEndDate As Date
    
    Dim austritte As Long
    Dim fluctuation As Double
    Dim avgHeadcount As Double
    Dim monthStartDate As Date
    
    Dim r As Long
    Dim employeeID As Variant
    Dim employeeName As Variant
    Dim entryDate As Variant
    Dim exitDate As Variant
    
    Dim arrID As Variant
    Dim arrName As Variant
    Dim arrEntry As Variant
    Dim arrExit As Variant
    Dim arrReason As Variant
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Sub
    
    monthNumber = CLng(ws.Range("A1").Value)
    currentYear = PID_GetWorkbookYear()
    
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    monthEndDate = DateSerial(currentYear, monthNumber + 1, 0)
    
    arrID = ws.Range("B3:B82").Value
    arrName = ws.Range("C3:C82").Value
    arrEntry = ws.Range("D3:D82").Value
    arrExit = ws.Range("I3:I82").Value
    arrReason = ws.Range("N3:N82").Value
    
    austritte = 0
    
    For r = 1 To 80
        employeeID = arrID(r, 1)
        employeeName = arrName(r, 1)
        entryDate = arrEntry(r, 1)
        exitDate = arrExit(r, 1)
        
        If IsDate(exitDate) Then
            If Year(CDate(exitDate)) = currentYear Then
                If Month(CDate(exitDate)) = monthNumber Then
                    If PID_FluctuationRowHasEmployee(employeeID, employeeName) Then
                        ' FP-Flukt FIX 2: neutrale Austrittsgruende (Spalte N = Austrittsgrund)
                        ' zaehlen NICHT in die Live-Monatsfluktuation. Zentrale Wahrheit
                        ' PID_IsNeutralFluctuationExitReason - identisch zu PID_CountExitsInPeriod.
                        If Not PID_IsNeutralFluctuationExitReason(CStr(arrReason(r, 1))) Then
                            austritte = austritte + 1
                        End If
                    End If
                End If
            End If
        End If
        
    Next r
    
    ' FP-FLUKT: Nenner = durchschnittlicher Personalbestand (Anfang+Ende)/2 statt nur Monatsende.
    ' Zaehler bleibt der sofort vom Monatsblatt gezaehlte Austrittswert (Live-Update bei I-Aenderung,
    ' bevor FLUKTUATION_DATEN neu aufgebaut wird).
    monthStartDate = DateSerial(currentYear, monthNumber, 1)
    avgHeadcount = (PID_CountEmployeesAtDate(ws, monthStartDate) + PID_CountEmployeesAtDate(ws, monthEndDate)) / 2#
    
    If avgHeadcount > 0 Then
        fluctuation = austritte / avgHeadcount
    Else
        fluctuation = 0
    End If
    
    ws.Range("Q31").Value = fluctuation
    ws.Range("Q31").NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT

SafeExit:
End Sub


Public Sub PID_FillFluktuationRates(ByRef monthExit() As Long, _
                                    ByRef monthPersonalEnde() As Long, _
                                    ByRef monthFluctuation() As Double, _
                                    ByRef quarterFluctuation() As Double, _
                                    ByRef ytdFluctuation As Double, _
                                    ByVal currentYear As Long)
    ' FP-FLUKT: Einheitliche Berechnung fuer Monat, Quartal und YTD.
    ' Jeder Wert = Austritte im Zeitraum / durchschnittlicher Personalbestand.
    ' Durchschnittsbestand = (Bestand Periodenanfang + Bestand Periodenende) / 2.
    ' Quartal und YTD werden NICHT aus Monatsprozenten gemittelt, sondern direkt
    ' ueber den Gesamtzeitraum (PID_ComputeFluctuationForPeriod) berechnet.
    Dim i As Long
    Dim ytdMonthLimit As Long
    Dim monthNames As Variant
    
    monthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    For i = 1 To 12
        ' Monatsend-Bestand weiterhin als Kennzahl fuellen (Rueckwaertskompatibilitaet).
        monthPersonalEnde(i) = GetPersonalEndeForMonth(CStr(monthNames(i - 1)), currentYear, i)
        
        ' Monatsfluktuation nach einheitlichem Standard (Anfang+Ende)/2 als Nenner.
        monthFluctuation(i) = PID_ComputeFluctuationForPeriod(DateSerial(currentYear, i, 1), DateSerial(currentYear, i + 1, 0))
    Next i
    
    quarterFluctuation(1) = PID_ComputeFluctuationForPeriod(DateSerial(currentYear, 1, 1), DateSerial(currentYear, 4, 0))
    quarterFluctuation(2) = PID_ComputeFluctuationForPeriod(DateSerial(currentYear, 4, 1), DateSerial(currentYear, 7, 0))
    quarterFluctuation(3) = PID_ComputeFluctuationForPeriod(DateSerial(currentYear, 7, 1), DateSerial(currentYear, 10, 0))
    quarterFluctuation(4) = PID_ComputeFluctuationForPeriod(DateSerial(currentYear, 10, 1), DateSerial(currentYear, 13, 0))
    
    If currentYear = Year(Date) Then
        ytdMonthLimit = Month(Date)
    ElseIf currentYear < Year(Date) Then
        ytdMonthLimit = 12
    Else
        ytdMonthLimit = GetLastMonthWithExit(monthExit)
    End If
    
    If ytdMonthLimit >= 1 And ytdMonthLimit <= 12 Then
        ytdFluctuation = PID_ComputeFluctuationForPeriod(DateSerial(currentYear, 1, 1), DateSerial(currentYear, ytdMonthLimit + 1, 0))
    Else
        ytdFluctuation = 0
    End If
End Sub


Public Sub PID_SyncFluktuationToDisplaySheets(ByRef monthFluctuation() As Double, _
                                              ByRef quarterFluctuation() As Double, _
                                              ByVal ytdFluctuation As Double)
    PID_SyncFluktuationToUbersicht monthFluctuation, quarterFluctuation, ytdFluctuation
    PID_SyncFluktuationToMonthSheets monthFluctuation
End Sub


Public Sub PID_SyncFluktuationToUbersicht(ByRef monthFluctuation() As Double, _
                                          ByRef quarterFluctuation() As Double, _
                                          ByVal ytdFluctuation As Double)
    Dim ws As Worksheet
    Dim monthRows As Variant
    Dim quarterRows As Variant
    Dim i As Long
    Dim wasProtected As Boolean
    
    On Error GoTo SafeExit
    
    Set ws = ThisWorkbook.Worksheets("UBERSICHT")
    
    wasProtected = ws.ProtectContents
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect
    End If
    On Error GoTo SafeExit
    
    monthRows = Array(7, 8, 9, 11, 12, 13, 15, 16, 17, 19, 20, 21)
    quarterRows = Array(10, 14, 18, 22)
    
    For i = 1 To 12
        ws.Cells(CLng(monthRows(i - 1)), 17).ClearContents
        ws.Cells(CLng(monthRows(i - 1)), 17).Value2 = monthFluctuation(i)
        ws.Cells(CLng(monthRows(i - 1)), 17).NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT
    Next i
    
    For i = 1 To 4
        ws.Cells(CLng(quarterRows(i - 1)), 17).ClearContents
        ws.Cells(CLng(quarterRows(i - 1)), 17).Value2 = quarterFluctuation(i)
        ws.Cells(CLng(quarterRows(i - 1)), 17).NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT
    Next i
    
    ws.Cells(23, 17).ClearContents
    ws.Cells(23, 17).Value2 = ytdFluctuation
    ws.Cells(23, 17).NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT
    
    If wasProtected Then
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
    End If

SafeExit:
End Sub


Public Sub PID_SyncFluktuationToMonthSheets(ByRef monthFluctuation() As Double)
    ' FP-FLUKT: robust ueber den zentralen Helper. Q31 wird je Monatsblatt aus
    ' PID_ComputeFluctuationForPeriod geschrieben (identisch zu UBERSICHT/FLUKTUATION) und
    ' der Blattschutz je Blatt sauber aufgehoben/wiederhergestellt, damit das gesperrte Q31
    ' nicht lautlos schreibgeschuetzt bleibt (UserInterfaceOnly wird beim Speichern nicht
    ' persistiert). Parameter bleibt aus Kompatibilitaet erhalten (Werte werden neu berechnet).
    PID_SyncAllMonthSheetsFluctuationQ31
End Sub


Public Sub PID_SyncAllMonthSheetsFluctuationQ31()
    ' Synchronisiert O31-Label + Q31-Wert auf allen 12 Monatsblaettern.
    Dim ws As Worksheet
    Dim i As Long
    
    For i = 1 To 12
        Set ws = PID_ResolveMonthSheetForIndex(i)
        If Not ws Is Nothing Then
            PID_SyncMonthSheetFluctuationQ31 ws
        End If
    Next i
End Sub


Public Sub PID_SyncMonthSheetFluctuationQ31(ByVal ws As Worksheet)
    ' Schreibt O31-Label "Fluktuation:" (IMMER, nicht nur wenn leer) und Q31 = Monatsfluktuation
    ' aus der zentralen Logik (PID_ComputeFluctuationForPeriod) - damit ist Q31 identisch zu
    ' UBERSICHT/FLUKTUATION. O31/Q31 sind gesperrte Zellen; der Schutz wird kurz aufgehoben,
    ' geschrieben und anschliessend kanonisch (UserInterfaceOnly) wiederhergestellt.
    Dim monthNumber As Long
    Dim currentYear As Long
    Dim fluctuation As Double
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Sub
    If Not IsNumeric(ws.Range("A1").Value) Then Exit Sub
    
    monthNumber = CLng(ws.Range("A1").Value)
    If monthNumber < 1 Or monthNumber > 12 Then Exit Sub
    
    currentYear = PID_GetWorkbookYear()
    If currentYear <= 0 Then Exit Sub
    
    ' Schutz aufheben (mit Passwort, Fallback ohne), damit O31/Q31 beschreibbar sind.
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    If ws.ProtectContents Then ws.Unprotect
    On Error GoTo SafeExit
    
    ' O31-Label IMMER auf "Fluktuation:" setzen (korrigiert auch falsche/alte Werte).
    ' Direkt nach dem Unprotect und VOR der Berechnung, damit das Label auch bei einem
    ' Berechnungsfehler korrekt gesetzt bleibt.
    ws.Range("O31").Value = "Fluktuation:"
    
    ' Q31 = Monatsfluktuation aus der zentralen Logik (unveraendert).
    fluctuation = PID_ComputeFluctuationForPeriod(DateSerial(currentYear, monthNumber, 1), DateSerial(currentYear, monthNumber + 1, 0))
    ws.Range("Q31").Value2 = fluctuation
    ws.Range("Q31").NumberFormat = PID_FLUKTUATION_PERCENT_FORMAT
    
    ' Kanonischer Monatsblatt-Schutz (UserInterfaceOnly, AllowSorting:=False) wiederherstellen.
    PID_ProtectWorkerMonthSheet ws
    Exit Sub

SafeExit:
    ' Best-effort: Schutz im Fehlerfall nicht offen lassen.
    On Error Resume Next
    If Not ws Is Nothing Then
        If Not ws.ProtectContents Then PID_ProtectWorkerMonthSheet ws
    End If
End Sub


Private Function PID_FluctuationRowHasEmployee(ByVal employeeID As Variant, ByVal employeeName As Variant) As Boolean
    PID_FluctuationRowHasEmployee = (Len(Trim$(CStr(employeeID))) > 0 Or Len(Trim$(CStr(employeeName))) > 0)
End Function


'==============================================================================
' FP-FLUKT: Zentrale Helper fuer einheitliche Fluktuationsberechnung
' (HR-Controlling-/BDA-Standard). Brutto-Fluktuation:
'   Fluktuationsrate = Austritte im Zeitraum / durchschnittlicher Personalbestand
'   Durchschnittsbestand = (Bestand Periodenanfang + Bestand Periodenende) / 2
' Excel-2016-kompatibel (Windows), keine dynamischen Array-Funktionen.
'==============================================================================

Public Function PID_CountEmployeesAtDate(ByVal ws As Worksheet, ByVal checkDate As Date) As Long
    ' Zaehlt Mitarbeiter, die am Stichtag checkDate zum Bestand zaehlen:
    '   (Eintrittsdatum leer ODER Eintritt <= Stichtag)
    '   UND (Austrittsdatum leer ODER Austritt > Stichtag)
    ' Austritt genau am Stichtag zaehlt NICHT mehr zum Bestand.
    Dim r As Long
    Dim arrID As Variant
    Dim arrName As Variant
    Dim arrEntry As Variant
    Dim arrExit As Variant
    Dim employeeID As Variant
    Dim employeeName As Variant
    Dim entryDate As Variant
    Dim exitDate As Variant
    Dim countEmployees As Long
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    
    countEmployees = 0
    
    ' Vier Bereichslesevorgaenge statt 320 Einzelzugriffen: die Funktion laeuft bei
    ' jedem Monatsblatt-Tabwechsel zweimal (Monatsanfang und -ende) ueber Q31.
    ' Gleiches Muster wie in PID_CalculateFluctuation.
    arrID = ws.Range("B3:B82").Value
    arrName = ws.Range("C3:C82").Value
    arrEntry = ws.Range("D3:D82").Value
    arrExit = ws.Range("I3:I82").Value
    
    For r = 1 To 80
        employeeID = arrID(r, 1)
        employeeName = arrName(r, 1)
        entryDate = arrEntry(r, 1)
        exitDate = arrExit(r, 1)
        
        If Trim$(CStr(employeeID)) <> "" Or Trim$(CStr(employeeName)) <> "" Then
            If IsDate(entryDate) Then
                If CDate(entryDate) <= checkDate Then
                    If Not IsDate(exitDate) Then
                        countEmployees = countEmployees + 1
                    ElseIf CDate(exitDate) > checkDate Then
                        countEmployees = countEmployees + 1
                    End If
                End If
            Else
                If Not IsDate(exitDate) Then
                    countEmployees = countEmployees + 1
                ElseIf CDate(exitDate) > checkDate Then
                    countEmployees = countEmployees + 1
                End If
            End If
        End If
    Next r
    
    PID_CountEmployeesAtDate = countEmployees
    Exit Function

SafeExit:
    PID_CountEmployeesAtDate = countEmployees
End Function


Public Function PID_CountExitsInPeriod(ByVal startDate As Date, ByVal endDate As Date) As Long
    ' Zaehlt deduplizierte Austritte mit Austrittsdatum im Zeitraum [startDate; endDate]
    ' aus FLUKTUATION_DATEN. Gleiche Dedup-Logik (PID_FluctuationExitDedupKey) wie die
    ' Analyse-Aggregation, damit die KPI-Werte konsistent bleiben.
    Dim dataWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim exitDate As Variant
    Dim exitKey As String
    Dim seenKeys As Collection
    Dim cnt As Long
    
    On Error GoTo SafeExit
    
    On Error Resume Next
    Set dataWs = ThisWorkbook.Worksheets("FLUKTUATION_DATEN")
    On Error GoTo SafeExit
    If dataWs Is Nothing Then Exit Function
    
    lastRow = dataWs.Cells(dataWs.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    Set seenKeys = New Collection
    cnt = 0
    
    For r = 2 To lastRow
        exitDate = dataWs.Cells(r, "E").Value
        If IsDate(exitDate) Then
            If CDate(exitDate) >= startDate And CDate(exitDate) <= endDate Then
                ' FP-Flukt FIX 2: neutrale Austrittsgruende (Karenz, Store transfer,
                ' Nicht eingetreten, Befoerderung) zaehlen NICHT in die Fluktuationsrate.
                ' Zentrale Wahrheit = PID_IsNeutralFluctuationExitReason (Spalte F = Austrittsgrund).
                If Not PID_IsNeutralFluctuationExitReason(CStr(dataWs.Cells(r, "F").Value)) Then
                    exitKey = PID_FluctuationExitDedupKey(dataWs, r)
                    If Len(exitKey) > 0 Then
                        On Error Resume Next
                        seenKeys.Add exitKey, exitKey
                        If Err.Number = 0 Then cnt = cnt + 1
                        Err.Clear
                        On Error GoTo SafeExit
                    End If
                End If
            End If
        End If
    Next r
    
    PID_CountExitsInPeriod = cnt
    Exit Function

SafeExit:
    PID_CountExitsInPeriod = cnt
End Function


Public Function PID_GetAverageHeadcountForPeriod(ByVal startDate As Date, ByVal endDate As Date) As Double
    ' Durchschnittlicher Personalbestand = (Bestand am Periodenanfang + Bestand am Periodenende) / 2.
    ' Periodenanfang wird auf dem Start-Monatsblatt, Periodenende auf dem End-Monatsblatt gezaehlt.
    Dim startWs As Worksheet
    Dim endWs As Worksheet
    Dim startCount As Long
    Dim endCount As Long
    
    On Error GoTo SafeExit
    
    Set startWs = PID_ResolveMonthSheetForIndex(Month(startDate))
    Set endWs = PID_ResolveMonthSheetForIndex(Month(endDate))
    
    If startWs Is Nothing And endWs Is Nothing Then
        PID_GetAverageHeadcountForPeriod = 0
        Exit Function
    ElseIf startWs Is Nothing Then
        PID_GetAverageHeadcountForPeriod = PID_CountEmployeesAtDate(endWs, endDate)
        Exit Function
    ElseIf endWs Is Nothing Then
        PID_GetAverageHeadcountForPeriod = PID_CountEmployeesAtDate(startWs, startDate)
        Exit Function
    End If
    
    startCount = PID_CountEmployeesAtDate(startWs, startDate)
    endCount = PID_CountEmployeesAtDate(endWs, endDate)
    
    PID_GetAverageHeadcountForPeriod = (startCount + endCount) / 2#
    Exit Function

SafeExit:
    PID_GetAverageHeadcountForPeriod = 0
End Function


Public Function PID_ComputeFluctuationForPeriod(ByVal startDate As Date, ByVal endDate As Date) As Double
    ' Einheitliche Brutto-Fluktuationsrate fuer einen beliebigen Zeitraum.
    Dim exits As Long
    Dim avgHeadcount As Double
    
    exits = PID_CountExitsInPeriod(startDate, endDate)
    avgHeadcount = PID_GetAverageHeadcountForPeriod(startDate, endDate)
    
    If avgHeadcount > 0 Then
        PID_ComputeFluctuationForPeriod = exits / avgHeadcount
    Else
        PID_ComputeFluctuationForPeriod = 0
    End If
End Function


Public Function PID_GetFluctuationRating(ByVal fluctuationRate As Double) As String
    ' HR-Controlling-Einstufung der Brutto-Fluktuationsrate (technische Grenzwerte
    ' 0.2 / 0.35 / 0.5 / 0.7 / 1.0). Umlaute ASCII-sicher via ChrW (Excel 2016).
    If fluctuationRate < 0.2 Then
        PID_GetFluctuationRating = "Sehr gut / stabil"
    ElseIf fluctuationRate < 0.35 Then
        PID_GetFluctuationRating = "Gut / normal"
    ElseIf fluctuationRate < 0.5 Then
        PID_GetFluctuationRating = "Erh" & ChrW(246) & "ht / beobachten"
    ElseIf fluctuationRate < 0.7 Then
        PID_GetFluctuationRating = "Hoch / analysieren"
    ElseIf fluctuationRate < 1 Then
        PID_GetFluctuationRating = "Sehr hoch / kritisch"
    Else
        PID_GetFluctuationRating = "Extrem hoch / akuter Handlungsbedarf"
    End If
End Function


Public Function PID_GetFluctuationStatusShort(ByVal fluctuationRate As Double) As String
    ' Kurz-Ampel fuer den Statustext (B5), abgeleitet aus derselben Rate/Logik wie
    ' PID_GetFluctuationRating (gleiche Grenzwerte 0.2 / 0.35 / 0.5 / 0.7 / 1.0).
    ' Langform liefert PID_GetFluctuationRating; hier nur die Management-Kurzform.
    If fluctuationRate < 0.2 Then
        PID_GetFluctuationStatusShort = "Sehr gut"
    ElseIf fluctuationRate < 0.35 Then
        PID_GetFluctuationStatusShort = "Gut"
    ElseIf fluctuationRate < 0.5 Then
        PID_GetFluctuationStatusShort = "Erh" & ChrW(246) & "ht"
    ElseIf fluctuationRate < 0.7 Then
        PID_GetFluctuationStatusShort = "Hoch"
    ElseIf fluctuationRate < 1 Then
        PID_GetFluctuationStatusShort = "Kritisch"
    Else
        PID_GetFluctuationStatusShort = "Extrem"
    End If
End Function


Public Function PID_GetFluctuationActionHint(ByVal fluctuationRate As Double) As String
    ' Kurzer Handlungsbedarf-Hinweis fuer die Management-Kurzbewertung, abgeleitet aus
    ' derselben Rate/Logik (Grenzwerte 0.2 / 0.35 / 0.5 / 0.7 / 1.0).
    If fluctuationRate < 0.2 Then
        PID_GetFluctuationActionHint = "kein akuter Handlungsbedarf"
    ElseIf fluctuationRate < 0.35 Then
        PID_GetFluctuationActionHint = "Entwicklung beobachten"
    ElseIf fluctuationRate < 0.5 Then
        PID_GetFluctuationActionHint = "Ursachen pr" & ChrW(252) & "fen"
    ElseIf fluctuationRate < 0.7 Then
        PID_GetFluctuationActionHint = "Ma" & ChrW(223) & "nahmen einleiten"
    ElseIf fluctuationRate < 1 Then
        PID_GetFluctuationActionHint = "dringend gegensteuern"
    Else
        PID_GetFluctuationActionHint = "akuter Handlungsbedarf"
    End If
End Function


Private Function PID_ResolveMonthSheetForIndex(ByVal monthIndex As Long) As Worksheet
    ' Liefert das Monatsblatt zum Monatsindex (1-12) oder Nothing, falls nicht vorhanden.
    Dim monthNames As Variant
    
    On Error GoTo SafeExit
    
    If monthIndex < 1 Or monthIndex > 12 Then Exit Function
    
    monthNames = PID_MonthNames()
    
    On Error Resume Next
    Set PID_ResolveMonthSheetForIndex = ThisWorkbook.Worksheets(CStr(monthNames(LBound(monthNames) + monthIndex - 1)))
    On Error GoTo 0

SafeExit:
End Function

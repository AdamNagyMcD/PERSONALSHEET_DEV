Attribute VB_Name = "mod_BuildFluktuationDaten"
Option Explicit


' Anzeigetexte mit Umlauten (ChrW = ASCII-sichere Quelle, Win/Mac Excel 2016).
Private Function PID_FlDatTxtAe() As String
    PID_FlDatTxtAe = ChrW(228)
End Function


Private Function PID_FlDatTxtOe() As String
    PID_FlDatTxtOe = ChrW(246)
End Function


Private Function PID_FlDatTxtUe() As String
    PID_FlDatTxtUe = ChrW(252)
End Function


Private Function PID_FlDatTxtAufloesung() As String
    PID_FlDatTxtAufloesung = "Einvernehmliche Aufl" & PID_FlDatTxtOe() & "sung"
End Function


Private Function PID_FlDatTxtDienstgeberKuendigung() As String
    PID_FlDatTxtDienstgeberKuendigung = "Dienstgeber K" & PID_FlDatTxtUe() & "ndigung"
End Function


Private Function PID_FlDatTxtDienstnehmerKuendigung() As String
    PID_FlDatTxtDienstnehmerKuendigung = "Dienstnehmer K" & PID_FlDatTxtUe() & "ndigung"
End Function


Private Function PID_FlDatTxtBefoerderung() As String
    PID_FlDatTxtBefoerderung = "Bef" & PID_FlDatTxtOe() & "rderung"
End Function


Private Function PID_FlDisplayMonthName(ByVal monthIndex As Long) As String
    If monthIndex = 3 Then
        PID_FlDisplayMonthName = "M" & ChrW(228) & "rz"
    ElseIf monthIndex >= 1 And monthIndex <= 12 Then
        PID_FlDisplayMonthName = CStr(PID_MonthNames()(monthIndex - 1))
    Else
        PID_FlDisplayMonthName = ""
    End If
End Function


Public Sub BuildFluktuationDaten()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim outWs As Worksheet
    
    Dim i As Long
    Dim r As Long 
    Dim outRow As Long
    Dim maxRows As Long
    
    Dim monthName As String
    Dim monthNumber As Long
    Dim currentYear As Long
    
    Dim personalID As Variant
    Dim employeeName As Variant
    Dim entryDate As Variant
    Dim exitDate As Variant
    Dim exitReason As String
    
    Dim daysInCompany As Long
    Dim reasonWeight As Double
    Dim timeFactor As Double
    Dim lossValue As Double
    Dim categoryText As String
    
    Dim arrOut() As Variant
    Dim oldDataLastRow As Long
    
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
    
    monthNames = PID_MonthNames()
    
    Set outWs = ThisWorkbook.Worksheets("FLUKTUATION_DATEN")
    
    currentYear = PID_GetWorkbookYear()
    If currentYear <= 0 Then GoTo CleanExit
    
    maxRows = 12 * 80
    ReDim arrOut(1 To maxRows, 1 To 11)
    outRow = 0
    
    outWs.Visible = xlSheetVisible
    outWs.Unprotect Password:=PID_WORKBOOK_PASSWORD
    outWs.Cells.Clear
    
    For i = LBound(monthNames) To UBound(monthNames)
        monthName = CStr(monthNames(i))
        monthNumber = i + 1
        
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(monthName)
        On Error GoTo CleanFail
        
        If Not ws Is Nothing Then
            If IsNumeric(ws.Range("A1").Value) Then
                monthNumber = CLng(ws.Range("A1").Value)
            Else
                monthNumber = i + 1
            End If
            
            If monthNumber < 1 Or monthNumber > 12 Then GoTo NextMonthSheet
            
            For r = 3 To 82
                personalID = ws.Cells(r, "B").Value
                employeeName = ws.Cells(r, "C").Value
                entryDate = ws.Cells(r, "D").Value
                exitDate = ws.Cells(r, "I").Value
                exitReason = NormalizeExitReason(CStr(ws.Cells(r, "N").Value))
                
                If IsDate(exitDate) Then
                    If Year(CDate(exitDate)) = currentYear Then
                        If Month(CDate(exitDate)) = monthNumber Then
                            If Trim$(CStr(personalID)) <> "" Or Trim$(CStr(employeeName)) <> "" Then
                                
                                If IsDate(entryDate) Then
                                    daysInCompany = DateDiff("d", CDate(entryDate), CDate(exitDate)) + 1
                                Else
                                    daysInCompany = 0
                                End If
                                
                                reasonWeight = GetReasonWeight(exitReason)
                                timeFactor = GetTimeFactor(daysInCompany)
                                categoryText = GetFluctuationCategory(exitReason, daysInCompany, reasonWeight, timeFactor)
                                
                                If categoryText = "Austrittsgrund fehlt" Then
                                    lossValue = 0
                                ElseIf categoryText = "Austrittsgrund unbekannt" Then
                                    lossValue = 0
                                ElseIf reasonWeight < 0 Then
                                    lossValue = 0
                                Else
                                    lossValue = reasonWeight * timeFactor
                                End If
                                
                                outRow = outRow + 1
                                
                                arrOut(outRow, 1) = PID_FlDisplayMonthName(Month(CDate(exitDate)))
                                arrOut(outRow, 2) = personalID
                                arrOut(outRow, 3) = employeeName
                                arrOut(outRow, 4) = entryDate
                                arrOut(outRow, 5) = exitDate
                                arrOut(outRow, 6) = exitReason
                                arrOut(outRow, 7) = daysInCompany
                                arrOut(outRow, 8) = reasonWeight
                                arrOut(outRow, 9) = timeFactor
                                arrOut(outRow, 10) = lossValue
                                arrOut(outRow, 11) = categoryText
                            End If
                        End If
                    End If
                End If
            Next r
        End If
        
NextMonthSheet:
    Next i
    
    With outWs
        .Range("A1").Value = "Monat"
        .Range("B1").Value = "Personal ID"
        .Range("C1").Value = "Name"
        .Range("D1").Value = "Eintrittsdatum"
        .Range("E1").Value = "Austrittsdatum"
        .Range("F1").Value = "Austrittsgrund"
        .Range("G1").Value = "Tage im Unternehmen"
        .Range("H1").Value = "Grund-Gewicht"
        .Range("I1").Value = "Zeit-Faktor"
        .Range("J1").Value = "Verlust-Score"
        .Range("K1").Value = "Kategorie"
        
        If outRow > 0 Then
            oldDataLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            
            .Range("A2").Resize(outRow, 11).Value = PID_FluctuationCopyOutputRows(arrOut, outRow)
            
            If oldDataLastRow > outRow + 1 Then
                .Rows((outRow + 2) & ":" & oldDataLastRow).Clear
            End If
        End If
        
        With .Range("A1:K1")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        .Columns("A:K").AutoFit
        .Range("D:E").NumberFormat = "dd.mm.yyyy"
        .Range("H:J").NumberFormat = "0.00"
        
        .Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
        .Visible = xlSheetVeryHidden
    End With

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not outWs Is Nothing Then
        outWs.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
        outWs.Visible = xlSheetVeryHidden
    End If
    
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler in BuildFluktuationDaten:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "Fluktuation Daten"
End Sub


Public Function NormalizeExitReason(ByVal exitReason As String) As String
    Dim s As String
    Dim k As String
    
    s = Trim$(CStr(exitReason))
    
    If s = "" Then
        NormalizeExitReason = ""
        Exit Function
    End If
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    k = ReasonKey(s)
    
    Select Case k
        Case "einvernehmliche aufloesung"
            NormalizeExitReason = PID_FlDatTxtAufloesung()
        
        Case "dienstgeber kuendigung"
            NormalizeExitReason = PID_FlDatTxtDienstgeberKuendigung()
        
        Case "dienstnehmer kuendigung"
            NormalizeExitReason = PID_FlDatTxtDienstnehmerKuendigung()
        
        Case "unberechtigter vorzeitiger austritt"
            NormalizeExitReason = "Unberechtigter vorzeitiger Austritt"
        
        Case "unberechtiger vorzeitiger austritt"
            NormalizeExitReason = "Unberechtigter vorzeitiger Austritt"
        
        Case "unberechtigte vorzeitige austritt"
            NormalizeExitReason = "Unberechtigter vorzeitiger Austritt"
        
        Case "unberechtigter vorzeitige austritt"
            NormalizeExitReason = "Unberechtigter vorzeitiger Austritt"
        
        Case "unberechtige vorzeitige austritt"
            NormalizeExitReason = "Unberechtigter vorzeitiger Austritt"
        
        Case "berechtigter vorzeitiger austritt"
            NormalizeExitReason = "Berechtigter vorzeitiger Austritt"
        
        Case "berechtigte vorzeitige austritt"
            NormalizeExitReason = "Berechtigter vorzeitiger Austritt"
        
        Case "berechtigter vorzeitige austritt"
            NormalizeExitReason = "Berechtigter vorzeitiger Austritt"
        
        Case "probezeit beendet"
            NormalizeExitReason = "Probezeit beendet"
        
        Case "befristung ende"
            NormalizeExitReason = "Befristung Ende"
        
        Case "karenz"
            NormalizeExitReason = "Karenz"
        
        Case "storetransfer"
            NormalizeExitReason = "Storetransfer"
        
        Case "befoerderung"
            NormalizeExitReason = PID_FlDatTxtBefoerderung()
        
        Case "nicht eingetreten"
            NormalizeExitReason = "Nicht eingetreten"
        
        Case "sonstiges"
            NormalizeExitReason = "Sonstiges"
        
        Case Else
            NormalizeExitReason = s
    End Select
End Function


Public Function ReasonKey(ByVal textValue As String) As String
    Dim k As String
    
    k = LCase$(Trim$(CStr(textValue)))
    
    k = Replace(k, ChrW(228), "ae")
    k = Replace(k, ChrW(246), "oe")
    k = Replace(k, ChrW(252), "ue")
    k = Replace(k, ChrW(196), "ae")
    k = Replace(k, ChrW(214), "oe")
    k = Replace(k, ChrW(220), "ue")
    k = Replace(k, ChrW(223), "ss")
    
    Do While InStr(k, "  ") > 0
        k = Replace(k, "  ", " ")
    Loop
    
    ReasonKey = k
End Function


Public Function GetReasonWeight(ByVal exitReason As String) As Double
    Dim wsConfig As Worksheet
    Dim normalizedReason As String
    Dim rowReason As String
    Dim r As Long
    
    normalizedReason = NormalizeExitReason(exitReason)
    
    If normalizedReason = "" Then
        GetReasonWeight = 0
        Exit Function
    End If
    
    On Error GoTo UnknownReason
    
    Set wsConfig = ThisWorkbook.Worksheets(PID_EINSTELLUNG_SHEET)
    
    For r = PID_FLUKTUATION_REASON_FIRST_ROW To PID_FLUKTUATION_REASON_LAST_ROW
        rowReason = NormalizeExitReason(CStr(wsConfig.Cells(r, "B").Value))
        
        If rowReason = normalizedReason Then
            GetReasonWeight = GetSafeDouble(wsConfig.Cells(r, "C").Value, -1)
            Exit Function
        End If
    Next r

UnknownReason:
    GetReasonWeight = -1
End Function


Public Function GetTimeFactor(ByVal daysInCompany As Long) As Double
    Dim wsConfig As Worksheet
    
    If daysInCompany <= 0 Then
        GetTimeFactor = 0
        Exit Function
    End If
    
    On Error GoTo UseDefaults
    
    Set wsConfig = ThisWorkbook.Worksheets(PID_EINSTELLUNG_SHEET)
    
    If daysInCompany <= 30 Then
        GetTimeFactor = GetSafeDouble(wsConfig.Range("C53").Value, 0.3)
    ElseIf daysInCompany <= 90 Then
        GetTimeFactor = GetSafeDouble(wsConfig.Range("C54").Value, 0.5)
    ElseIf daysInCompany <= 180 Then
        GetTimeFactor = GetSafeDouble(wsConfig.Range("C55").Value, 0.75)
    ElseIf daysInCompany <= 365 Then
        GetTimeFactor = GetSafeDouble(wsConfig.Range("C56").Value, 1)
    ElseIf daysInCompany <= 730 Then
        GetTimeFactor = GetSafeDouble(wsConfig.Range("C57").Value, 1.25)
    ElseIf daysInCompany <= 1825 Then
        GetTimeFactor = GetSafeDouble(wsConfig.Range("C58").Value, 1.5)
    Else
        GetTimeFactor = GetSafeDouble(wsConfig.Range("C59").Value, 2)
    End If
    
    Exit Function

UseDefaults:
    If daysInCompany <= 30 Then
        GetTimeFactor = 0.3
    ElseIf daysInCompany <= 90 Then
        GetTimeFactor = 0.5
    ElseIf daysInCompany <= 180 Then
        GetTimeFactor = 0.75
    ElseIf daysInCompany <= 365 Then
        GetTimeFactor = 1
    ElseIf daysInCompany <= 730 Then
        GetTimeFactor = 1.25
    ElseIf daysInCompany <= 1825 Then
        GetTimeFactor = 1.5
    Else
        GetTimeFactor = 2
    End If
End Function


Public Function GetSafeDouble(ByVal valueToCheck As Variant, ByVal defaultValue As Double) As Double
    If IsNumeric(valueToCheck) Then
        GetSafeDouble = CDbl(valueToCheck)
    Else
        GetSafeDouble = defaultValue
    End If
End Function


Public Function GetFluctuationCategory(ByVal exitReason As String, ByVal daysInCompany As Long, ByVal reasonWeight As Double, ByVal timeFactor As Double) As String
    Dim normalizedReason As String
    
    normalizedReason = NormalizeExitReason(exitReason)
    
    If normalizedReason = "" Then
        GetFluctuationCategory = "Austrittsgrund fehlt"
        Exit Function
    End If
    
    Select Case normalizedReason
        Case "Storetransfer"
            GetFluctuationCategory = "Neutrale Bewegung"
            Exit Function
        
        Case PID_FlDatTxtBefoerderung()
            GetFluctuationCategory = "Neutrale Bewegung"
            Exit Function
        
        Case "Karenz"
            GetFluctuationCategory = "Neutrale Bewegung"
            Exit Function
        
        Case "Nicht eingetreten"
            GetFluctuationCategory = "Neutrale Bewegung"
            Exit Function
    End Select
    
    If reasonWeight < 0 Then
        GetFluctuationCategory = "Austrittsgrund unbekannt"
        Exit Function
    End If
    
    If daysInCompany <= 90 Then
        GetFluctuationCategory = "Austritt in den ersten 90 Tagen"
    ElseIf daysInCompany > 730 And (reasonWeight * timeFactor) >= 1 Then
        GetFluctuationCategory = "Verlust erfahrener Mitarbeiter"
    ElseIf (reasonWeight * timeFactor) >= 1 Then
        GetFluctuationCategory = "Wichtiger Austritt"
    Else
        GetFluctuationCategory = "Normaler Austritt"
    End If
End Function


Public Function PID_FluctuationDataRowIsExit(ByVal ws As Worksheet, _
                                            ByVal rowNumber As Long, _
                                            Optional ByVal workbookYear As Long = 0) As Boolean
    Dim exitDate As Variant
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If rowNumber < 2 Then Exit Function
    
    If Trim$(CStr(ws.Cells(rowNumber, "B").Value)) = "" _
       And Trim$(CStr(ws.Cells(rowNumber, "C").Value)) = "" Then Exit Function
    
    exitDate = ws.Cells(rowNumber, "E").Value
    If Not IsDate(exitDate) Then Exit Function
    
    If workbookYear > 0 Then
        If Year(CDate(exitDate)) <> workbookYear Then Exit Function
    End If
    
    PID_FluctuationDataRowIsExit = True

SafeExit:
End Function


Public Function PID_FluctuationExitDedupKey(ByVal ws As Worksheet, ByVal rowNumber As Long) As String
    Dim personalID As String
    Dim employeeName As String
    Dim exitDate As Variant
    
    On Error GoTo SafeExit
    
    If ws Is Nothing Then Exit Function
    If rowNumber < 2 Then Exit Function
    
    personalID = Trim$(CStr(ws.Cells(rowNumber, "B").Value))
    employeeName = Trim$(CStr(ws.Cells(rowNumber, "C").Value))
    exitDate = ws.Cells(rowNumber, "E").Value
    
    If personalID <> "" Then
        PID_FluctuationExitDedupKey = "ID|" & personalID & "|" & Format$(CDate(exitDate), "yyyymmdd")
    Else
        PID_FluctuationExitDedupKey = "NAME|" & employeeName & "|" & Format$(CDate(exitDate), "yyyymmdd")
    End If

SafeExit:
End Function


Public Function PID_FluctuationCopyOutputRows(ByVal sourceArr As Variant, ByVal rowCount As Long) As Variant
    Dim dest() As Variant
    Dim r As Long
    Dim c As Long
    
    If rowCount <= 0 Then Exit Function
    If Not IsArray(sourceArr) Then Exit Function
    
    ReDim dest(1 To rowCount, 1 To 11)
    
    For r = 1 To rowCount
        For c = 1 To 11
            dest(r, c) = sourceArr(r, c)
        Next c
    Next r
    
    PID_FluctuationCopyOutputRows = dest
End Function


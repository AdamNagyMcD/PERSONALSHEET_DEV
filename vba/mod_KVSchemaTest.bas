Attribute VB_Name = "mod_KVSchemaTest"
Option Explicit

Public Sub BuildLohntabelleTest()
    Dim ws As Worksheet
    
    Dim firstStartYearInput As Variant
    Dim firstSchemaCountInput As Variant
    Dim secondSchemaCountInput As Variant
    
    Dim firstStartYear As Long
    Dim firstSchemaCount As Long
    Dim secondSchemaCount As Long
    Dim nextRow As Long
    
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    On Error GoTo CleanFail
    
    firstStartYearInput = Application.InputBox( _
        Prompt:="Startjahr der ersten KV-Periode eingeben." & vbCrLf & _
                "Beispiel: 2024 ergibt KV 2024/2025.", _
        Title:="LOHNTABELLE_TEST erstellen", _
        Default:=Year(Date) - 1, _
        Type:=1)
    
    If firstStartYearInput = False Then Exit Sub
    
    firstSchemaCountInput = Application.InputBox( _
        Prompt:="Wie viele Monatsstunden-Zeilen soll jede KV-Code-Gruppe in der ersten Periode haben?" & vbCrLf & _
                "Beispiel: 13", _
        Title:="Anzahl Schemata erste Periode", _
        Default:=13, _
        Type:=1)
    
    If firstSchemaCountInput = False Then Exit Sub
    
    secondSchemaCountInput = Application.InputBox( _
        Prompt:="Wie viele Monatsstunden-Zeilen soll jede KV-Code-Gruppe in der zweiten Periode haben?" & vbCrLf & _
                "Beispiel: 15", _
        Title:="Anzahl Schemata zweite Periode", _
        Default:=15, _
        Type:=1)
    
    If secondSchemaCountInput = False Then Exit Sub
    
    firstStartYear = CLng(firstStartYearInput)
    firstSchemaCount = CLng(firstSchemaCountInput)
    secondSchemaCount = CLng(secondSchemaCountInput)
    
    If firstStartYear < 2000 Or firstStartYear > 2100 Then
        MsgBox "Bitte ein gueltiges Startjahr zwischen 2000 und 2100 eingeben.", _
               vbExclamation, "LOHNTABELLE_TEST"
        Exit Sub
    End If
    
    If firstSchemaCount < 1 Or firstSchemaCount > 50 Then
        MsgBox "Bitte fuer die erste Periode eine Anzahl zwischen 1 und 50 eingeben.", _
               vbExclamation, "LOHNTABELLE_TEST"
        Exit Sub
    End If
    
    If secondSchemaCount < 1 Or secondSchemaCount > 50 Then
        MsgBox "Bitte fuer die zweite Periode eine Anzahl zwischen 1 und 50 eingeben.", _
               vbExclamation, "LOHNTABELLE_TEST"
        Exit Sub
    End If
    
    If MsgBox("ACHTUNG:" & vbCrLf & vbCrLf & _
              "LOHNTABELLE_TEST wird komplett neu aufgebaut." & vbCrLf & _
              "Bestehende KV-Daten auf diesem Blatt werden geloescht." & vbCrLf & vbCrLf & _
              "Fortfahren?", _
              vbQuestion + vbYesNo, "LOHNTABELLE_TEST neu erstellen") <> vbYes Then
        Exit Sub
    End If
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    Set ws = GetOrCreateKVTestSheet("LOHNTABELLE_TEST")
    
    On Error Resume Next
    ws.Unprotect Password:=PID_WORKBOOK_PASSWORD
    On Error GoTo CleanFail
    
    ws.Cells.Clear
    
    BuildKVTestHeader ws
    BuildKVTableHeader ws
    
    nextRow = 4
    
    ' Neue Periode oben, alte Periode darunter.
    nextRow = AddKVPeriodRows(ws, nextRow, firstStartYear + 1, secondSchemaCount)
    nextRow = AddKVPeriodRows(ws, nextRow, firstStartYear, firstSchemaCount)
    
    FormatLohntabelleTest ws, nextRow - 1
    
    ws.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    
    MarkKVDropdownsDirty
    
    MsgBox "LOHNTABELLE_TEST wurde neu erstellt." & vbCrLf & vbCrLf & _
           "Bitte danach Monatsstunden und Monatslohn kontrollieren.", _
           vbInformation, "LOHNTABELLE_TEST"

CleanExit:
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    On Error Resume Next
    
    If Not ws Is Nothing Then
        ws.Protect Password:=PID_WORKBOOK_PASSWORD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    
    MsgBox "Fehler beim Erstellen von LOHNTABELLE_TEST:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbCritical, "LOHNTABELLE_TEST"
End Sub


Private Function GetOrCreateKVTestSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    End If
    
    Set GetOrCreateKVTestSheet = ws
End Function


Private Sub BuildKVTestHeader(ByVal ws As Worksheet)
    ws.Range("A1").Value = "KV-SCHEMA PFLEGE"
    
    ws.Range("A2").Value = _
        "Wichtig: Alte KV-Perioden niemals loeschen oder ueberschreiben. " & _
        "Wenn ab Mai neue Werte gueltig sind, immer eine neue KV-Periode hinzufuegen. " & _
        "In den Monatsblaettern wird spaeter nur der KV-Code ausgewaehlt. " & _
        "Nur Zeilen mit Status OK verwenden."
End Sub


Private Sub BuildKVTableHeader(ByVal ws As Worksheet)
    ws.Range("A3").Value = "KV-Periode"
    ws.Range("B3").Value = "Gueltig ab"
    ws.Range("C3").Value = "Gueltig bis"
    ws.Range("D3").Value = "KV-Code"
    ws.Range("E3").Value = "KV-Gruppe"
    ws.Range("F3").Value = "Beschaeftigungsdauer"
    ws.Range("G3").Value = "Monatsstunden"
    ws.Range("H3").Value = "Monatslohn"
    ws.Range("I3").Value = "Status"
    ws.Range("J3").Value = "Pruefung"
End Sub


Private Function AddKVPeriodRows(ByVal ws As Worksheet, _
                                 ByVal startRow As Long, _
                                 ByVal startYear As Long, _
                                 ByVal schemaCount As Long) As Long
    Dim codeIndex As Long
    Dim schemaIndex As Long
    Dim r As Long
    
    Dim periodName As String
    Dim validFrom As Date
    Dim validTo As Date
    
    Dim kvCode As String
    Dim kvGroup As String
    Dim durationText As String
    
    periodName = "KV " & CStr(startYear) & "/" & CStr(startYear + 1)
    validFrom = DateSerial(startYear, 5, 1)
    validTo = DateSerial(startYear + 1, 4, 30)
    
    r = startRow
    
    r = AddKVPeriodTitleRow(ws, r, periodName, validFrom, validTo)
    
    For codeIndex = 1 To 12
        kvCode = GetKVCodeByIndex(codeIndex)
        kvGroup = GetKVGroupByIndex(codeIndex)
        durationText = GetKVDurationByIndex(codeIndex)
        
        For schemaIndex = 1 To schemaCount
            WriteKVDataRow ws, r, periodName, validFrom, validTo, kvCode, kvGroup, durationText, "", ""
            r = r + 1
        Next schemaIndex
        
        With ws.Range("A" & r - 1 & ":J" & r - 1).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Next codeIndex
    
    r = r + 1
    
    AddKVPeriodRows = r
End Function


Private Function AddKVPeriodTitleRow(ByVal ws As Worksheet, _
                                     ByVal rowNumber As Long, _
                                     ByVal periodName As String, _
                                     ByVal validFrom As Date, _
                                     ByVal validTo As Date) As Long
    ws.Range("A" & rowNumber & ":J" & rowNumber).UnMerge
    ws.Range("A" & rowNumber & ":J" & rowNumber).Clear
    
    ws.Range("A" & rowNumber & ":J" & rowNumber).Merge
    ws.Range("A" & rowNumber).Value = periodName & _
                                      "   |   gueltig von " & Format$(validFrom, "dd.mm.yyyy") & _
                                      " bis " & Format$(validTo, "dd.mm.yyyy")
    
    With ws.Range("A" & rowNumber & ":J" & rowNumber)
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
    AddKVPeriodTitleRow = rowNumber + 1
End Function


Private Sub WriteKVDataRow(ByVal ws As Worksheet, _
                           ByVal rowNumber As Long, _
                           ByVal periodName As String, _
                           ByVal validFrom As Date, _
                           ByVal validTo As Date, _
                           ByVal kvCode As String, _
                           ByVal kvGroup As String, _
                           ByVal durationText As String, _
                           ByVal monatsstundenValue As Variant, _
                           ByVal monatslohnValue As Variant)
    ws.Range("A" & rowNumber & ":J" & rowNumber).UnMerge
    ws.Range("A" & rowNumber & ":J" & rowNumber).ClearContents
    
    ws.Cells(rowNumber, "A").Value = periodName
    ws.Cells(rowNumber, "B").Value = validFrom
    ws.Cells(rowNumber, "C").Value = validTo
    ws.Cells(rowNumber, "D").Value = kvCode
    ws.Cells(rowNumber, "E").Value = kvGroup
    ws.Cells(rowNumber, "F").Value = durationText
    ws.Cells(rowNumber, "G").Value = monatsstundenValue
    ws.Cells(rowNumber, "H").Value = monatslohnValue
    
    ws.Cells(rowNumber, "I").FormulaR1C1 = _
        "=IF(RC1="""","""",IF(OR(RC2="""",RC3="""",RC4="""",RC5="""",RC6=""""),""Stammdaten fehlen"",IF(AND(RC7="""",RC8=""""),""Werte fehlen"",IF(RC7="""",""Monatsstunden fehlen"",IF(RC8="""",""Monatslohn fehlt"",""OK"")))))"
    
    ws.Cells(rowNumber, "J").FormulaR1C1 = _
        "=IF(RC7="""","""",IF(COUNTIFS(C1,RC1,C4,RC4,C7,RC7)>1,""Doppelte Monatsstunden"",""""))"
    
    With ws.Range("A" & rowNumber & ":J" & rowNumber)
        .Font.Bold = False
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
End Sub


Private Sub FormatLohntabelleTest(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim dataRange As Range
    Dim inputRange As Range
    Dim statusRange As Range
    Dim checkRange As Range
    
    If lastRow < 4 Then Exit Sub
    
    Set dataRange = ws.Range("A3:J" & lastRow)
    Set inputRange = ws.Range("G4:H" & lastRow)
    Set statusRange = ws.Range("I4:I" & lastRow)
    Set checkRange = ws.Range("J4:J" & lastRow)
    
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Cells.RowHeight = 18
    
    ws.Range("K:XFD").Clear
    
    ws.Range("A1:J1").UnMerge
    ws.Range("A1:J1").Merge
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 20
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    
    ws.Range("A2:J2").UnMerge
    ws.Range("A2:J2").Merge
    ws.Range("A2").WrapText = True
    ws.Range("A2").VerticalAlignment = xlCenter
    ws.Range("A2").HorizontalAlignment = xlLeft
    
    ws.Range("A3:J3").Font.Bold = True
    ws.Range("A3:J3").HorizontalAlignment = xlCenter
    ws.Range("A3:J3").VerticalAlignment = xlCenter
    
    dataRange.Borders.LineStyle = xlContinuous
    dataRange.Borders.Weight = xlThin
    
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 24
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 22
    ws.Columns("J").ColumnWidth = 24
    ws.Columns("K").ColumnWidth = 3
    
    ws.Range("B4:C" & lastRow).NumberFormat = "dd.mm.yyyy"
    ws.Range("G4:G" & lastRow).NumberFormatLocal = "0,00"
    PID_ApplyEuroNumberFormat ws.Range("H4:H" & lastRow)
    
    inputRange.Interior.Color = RGB(255, 242, 204)
    
    statusRange.FormatConditions.Delete
    
    statusRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OK"""
    statusRange.FormatConditions(1).Interior.Color = RGB(198, 239, 206)
    statusRange.FormatConditions(1).Font.Color = RGB(0, 97, 0)
    
    statusRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=""OK"""
    statusRange.FormatConditions(2).Interior.Color = RGB(255, 235, 156)
    statusRange.FormatConditions(2).Font.Color = RGB(156, 101, 0)
    
    checkRange.FormatConditions.Delete
    
    checkRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="="""""
    checkRange.FormatConditions(1).Interior.Color = RGB(255, 199, 206)
    checkRange.FormatConditions(1).Font.Color = RGB(156, 0, 6)
    
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
    
    dataRange.AutoFilter
    
    ws.Rows("1").RowHeight = 28
    ws.Rows("2").RowHeight = 42
    ws.Rows("3").RowHeight = 24
    ws.Rows("4:" & lastRow).RowHeight = 18
    
    ApplyKVPeriodTitleRowHeights ws, lastRow
    
    On Error Resume Next
    ws.Activate
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub


Private Sub ApplyKVPeriodTitleRowHeights(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    
    For r = 4 To lastRow
        If ws.Cells(r, "A").MergeCells Then
            If Left$(CStr(ws.Cells(r, "A").Value), 2) = "KV" Then
                ws.Rows(r).RowHeight = 24
            End If
        End If
    Next r
End Sub


Private Function GetKVCodeByIndex(ByVal indexNumber As Long) As String
    Select Case indexNumber
        Case 1
            GetKVCodeByIndex = "BG1_Basis"
        Case 2
            GetKVCodeByIndex = "BG1_5"
        Case 3
            GetKVCodeByIndex = "BG1_10"
        Case 4
            GetKVCodeByIndex = "BG1_15"
        Case 5
            GetKVCodeByIndex = "BG2_Basis"
        Case 6
            GetKVCodeByIndex = "BG2_5"
        Case 7
            GetKVCodeByIndex = "BG2_10"
        Case 8
            GetKVCodeByIndex = "BG2_15"
        Case 9
            GetKVCodeByIndex = "BG3_Basis"
        Case 10
            GetKVCodeByIndex = "BG3_5"
        Case 11
            GetKVCodeByIndex = "BG3_10"
        Case 12
            GetKVCodeByIndex = "BG3_15"
        Case Else
            GetKVCodeByIndex = ""
    End Select
End Function


Private Function GetKVGroupByIndex(ByVal indexNumber As Long) As String
    Select Case indexNumber
        Case 1 To 4
            GetKVGroupByIndex = "BG1"
        Case 5 To 8
            GetKVGroupByIndex = "BG2"
        Case 9 To 12
            GetKVGroupByIndex = "BG3"
        Case Else
            GetKVGroupByIndex = ""
    End Select
End Function


Private Function GetKVDurationByIndex(ByVal indexNumber As Long) As String
    Select Case indexNumber
        Case 1, 5, 9
            GetKVDurationByIndex = "Basis / bis 5 Jahre"
        Case 2, 6, 10
            GetKVDurationByIndex = "5 bis 10 Jahre"
        Case 3, 7, 11
            GetKVDurationByIndex = "10 bis 15 Jahre"
        Case 4, 8, 12
            GetKVDurationByIndex = "ueber 15 Jahre"
        Case Else
            GetKVDurationByIndex = ""
    End Select
End Function


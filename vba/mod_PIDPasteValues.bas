Attribute VB_Name = "mod_PIDPasteValues"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' TR-02 - Einfuegen immer nur als Wert, in der ganzen Mappe.
'
' Zwei Ebenen:
'
' 1. VORBEUGEND (dieses Modul): Strg+V, Strg+Umschalt+V und Umschalt+Einfg werden
'    abgefangen und durch ein reines Werte-Einfuegen ersetzt. Damit kommt fremde
'    Formatierung gar nicht erst ins Blatt - und die Gueltigkeitspruefung (Dropdowns)
'    der Zielzellen bleibt erhalten, die ein normales Einfuegen ueberschreiben wuerde.
'
' 2. NETZ (EnforcePasteValuesOnly in DieseArbeitsmappe): raeumt nachtraeglich auf, was
'    ueber Menueband, Rechtsklick oder Ziehen eingefuegt wurde.
'
' Wichtig: Application.OnKey gilt fuer die ganze Excel-Anwendung, nicht nur fuer diese
' Mappe. Deshalb wird die Belegung beim Aktivieren gesetzt und beim Deaktivieren bzw.
' Schliessen wieder entfernt.
'
' Ebenso wichtig: die Blaetter sind mit UserInterfaceOnly:=True geschuetzt. VBA darf
' damit auch in gesperrte Zellen schreiben, was ein normales Strg+V nie koennte.
' Vor jedem Einfuegen wird deshalb geprueft, ob der Zielbereich gesperrte Zellen
' (z. B. die Formelspalten G, H, K, L) treffen wuerde.

Private Const PID_PV_KEY_PASTE As String = "^v"
Private Const PID_PV_KEY_PASTE_SPECIAL As String = "+^v"
Private Const PID_PV_KEY_PASTE_INSERT As String = "+{INSERT}"

Private Const PID_PV_MAX_CELLS As Long = 20000

' Zaehler statt Boolean: verschachtelte Aufrufe duerfen die Markierung nicht vorzeitig loeschen.
Private mManagedPasteDepth As Long


Public Sub PID_InstallPasteHooks()
    Dim macroName As String

    On Error Resume Next

    macroName = "'" & ThisWorkbook.Name & "'!PID_PasteValuesOnly"

    Application.OnKey PID_PV_KEY_PASTE, macroName
    Application.OnKey PID_PV_KEY_PASTE_SPECIAL, macroName
    Application.OnKey PID_PV_KEY_PASTE_INSERT, macroName
End Sub


Public Sub PID_RemovePasteHooks()
    On Error Resume Next

    ' Ohne zweites Argument stellt OnKey die Excel-Standardbelegung wieder her.
    Application.OnKey PID_PV_KEY_PASTE
    Application.OnKey PID_PV_KEY_PASTE_SPECIAL
    Application.OnKey PID_PV_KEY_PASTE_INSERT
End Sub


Public Sub PID_PasteValuesOnly()
    Dim selRange As Range
    Dim destRange As Range
    Dim clipRows As Long
    Dim clipCols As Long

    On Error GoTo PasteFailed

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    Set selRange = Selection

    If selRange.Areas.Count > 1 Then
        PID_PVInfo "Bitte nur einen zusammenh" & PID_UTxtAe() & "ngenden Bereich ausw" & PID_UTxtAe() & "hlen."
        Exit Sub
    End If

    ' Ausschneiden + Einfuegen verschiebt Formeln und Formate und kann Bezuege zerstoeren.
    If Application.CutCopyMode = xlCut Then
        Application.CutCopyMode = False
        PID_PVInfo "Ausschneiden und Einf" & PID_UTxtUe() & "gen ist hier nicht vorgesehen." & vbCrLf & vbCrLf & _
                   "Bitte mit Strg+C kopieren und dann einf" & PID_UTxtUe() & "gen."
        Exit Sub
    End If

    PID_PVGetClipboardExtent clipRows, clipCols

    If clipRows = 0 And Application.CutCopyMode = False Then
        PID_PVInfo "Die Zwischenablage enth" & PID_UTxtAe() & "lt keinen Text zum Einf" & PID_UTxtUe() & "gen."
        Exit Sub
    End If

    Set destRange = PID_PVBuildDestinationRange(selRange, clipRows, clipCols)
    If destRange Is Nothing Then Exit Sub

    If destRange.CountLarge > PID_PV_MAX_CELLS Then
        PID_PVInfo "Der Bereich ist zu gro" & PID_UTxtSs() & " zum Einf" & PID_UTxtUe() & "gen." & vbCrLf & vbCrLf & _
                   "Bitte in kleineren Bl" & PID_UTxtOe() & "cken einf" & PID_UTxtUe() & "gen."
        Exit Sub
    End If

    If PID_PVHasLockedCells(destRange) Then
        PID_PVInfo "Der eingef" & PID_UTxtUe() & "gte Bereich w" & PID_UTxtUe() & "rde " & PID_UTxtGeschuetzt() & "e Zellen " & _
                   "(z. B. die Formelspalten) " & PID_UTxtUe() & "berschreiben." & vbCrLf & vbCrLf & _
                   "Bitte weniger Spalten oder Zeilen einf" & PID_UTxtUe() & "gen."
        Exit Sub
    End If

    PID_PVBeginManagedPaste

    If Application.CutCopyMode = xlCopy Then
        ' Quelle liegt in dieser Excel-Instanz: Excel uebernimmt nur die Werte.
        selRange.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    Else
        ' Quelle ausserhalb (Word, Browser, andere Excel-Instanz): Text der Zwischenablage
        ' selbst Zelle fuer Zelle schreiben. Das ist unabhaengig von der Sprachversion und
        ' kann keine Formatierung mitbringen.
        PID_PVWriteClipboardText selRange.Cells(1, 1)
    End If

    PID_PVEndManagedPaste
    Exit Sub

PasteFailed:
    PID_PVEndManagedPaste

    On Error Resume Next
    Application.CutCopyMode = False

    PID_PVInfo "Der Inhalt konnte hier nicht eingef" & PID_UTxtUe() & "gt werden." & vbCrLf & vbCrLf & _
               "M" & PID_UTxtOe() & "gliche Gr" & PID_UTxtUe() & "nde: die Zelle ist " & PID_UTxtGeschuetzt() & ", " & _
               "der kopierte Bereich passt nicht zur Auswahl, oder die Zwischenablage " & _
               "enth" & PID_UTxtAe() & "lt keinen Text."
End Sub


' Zielbereich = Auswahl plus der Block, den die Zwischenablage ab der linken oberen
' Zelle belegen wuerde. Bewusst grosszuegig: die Sperr-Pruefung soll eher zu viel als
' zu wenig abdecken.
Private Function PID_PVBuildDestinationRange(ByVal selRange As Range, _
                                             ByVal clipRows As Long, _
                                             ByVal clipCols As Long) As Range
    Dim topLeft As Range
    Dim clipBlock As Range

    On Error GoTo SafeExit

    Set topLeft = selRange.Cells(1, 1)

    If clipRows <= 0 Or clipCols <= 0 Then
        Set PID_PVBuildDestinationRange = selRange
        Exit Function
    End If

    If topLeft.Row + clipRows - 1 > topLeft.Worksheet.Rows.Count Then GoTo SafeExit
    If topLeft.Column + clipCols - 1 > topLeft.Worksheet.Columns.Count Then GoTo SafeExit

    Set clipBlock = topLeft.Resize(clipRows, clipCols)
    Set PID_PVBuildDestinationRange = Union(selRange, clipBlock)
    Exit Function

SafeExit:
    Set PID_PVBuildDestinationRange = selRange
End Function


Private Function PID_PVHasLockedCells(ByVal rng As Range) As Boolean
    Dim lockedState As Variant

    On Error Resume Next

    If rng Is Nothing Then Exit Function
    If Not rng.Worksheet.ProtectContents Then Exit Function

    lockedState = rng.Locked

    ' Null = gemischt, also mindestens eine gesperrte Zelle im Bereich.
    If IsNull(lockedState) Then
        PID_PVHasLockedCells = True
    Else
        PID_PVHasLockedCells = CBool(lockedState)
    End If
End Function


Private Sub PID_PVWriteClipboardText(ByVal topLeft As Range)
    Dim lines() As String
    Dim cols() As String
    Dim lineTotal As Long
    Dim r As Long
    Dim c As Long
    Dim targetCell As Range

    lines = PID_PVGetClipboardLines()
    lineTotal = PID_PVLineCount(lines)
    If lineTotal = 0 Then Exit Sub

    For r = 0 To lineTotal - 1
        cols = Split(lines(r), vbTab)

        For c = 0 To UBound(cols)
            Set targetCell = topLeft.Offset(r, c)

            If targetCell.MergeCells Then
                ' Nur in die fuehrende Zelle des verbundenen Bereichs schreiben.
                If targetCell.Address = targetCell.MergeArea.Cells(1, 1).Address Then
                    targetCell.Value = PID_PVCleanCellText(cols(c))
                End If
            Else
                targetCell.Value = PID_PVCleanCellText(cols(c))
            End If
        Next c
    Next r
End Sub


' Anfuehrungszeichen um mehrzeilige Zellen sowie harte Zeilenumbrueche entfernen -
' im Blatt sollen einzeilige, saubere Werte stehen.
Private Function PID_PVCleanCellText(ByVal rawText As String) As String
    Dim result As String

    result = rawText

    If Len(result) >= 2 Then
        If Left$(result, 1) = """" And Right$(result, 1) = """" Then
            result = Mid$(result, 2, Len(result) - 2)
            result = Replace(result, """""", """")
        End If
    End If

    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")

    PID_PVCleanCellText = Trim$(result)
End Function


Private Sub PID_PVGetClipboardExtent(ByRef rowCount As Long, ByRef colCount As Long)
    Dim lines() As String
    Dim i As Long
    Dim cols As Long

    rowCount = 0
    colCount = 0

    lines = PID_PVGetClipboardLines()

    rowCount = PID_PVLineCount(lines)
    If rowCount = 0 Then Exit Sub

    For i = 0 To rowCount - 1
        cols = UBound(Split(lines(i), vbTab)) + 1
        If cols > colCount Then colCount = cols
    Next i
End Sub


Private Function PID_PVGetClipboardLines() As String()
    Dim clipText As String
    Dim emptyResult() As String

    clipText = PID_PVGetClipboardText()

    If Len(clipText) = 0 Then
        PID_PVGetClipboardLines = emptyResult
        Exit Function
    End If

    clipText = Replace(clipText, vbCrLf, vbLf)
    clipText = Replace(clipText, vbCr, vbLf)

    Do While Len(clipText) > 0
        If Right$(clipText, 1) <> vbLf Then Exit Do
        clipText = Left$(clipText, Len(clipText) - 1)
    Loop

    If Len(clipText) = 0 Then
        PID_PVGetClipboardLines = emptyResult
        Exit Function
    End If

    PID_PVGetClipboardLines = Split(clipText, vbLf)
End Function


' Ein nicht initialisiertes String-Array laesst UBound scheitern - das ist der Normalfall
' bei leerer Zwischenablage und bedeutet schlicht "keine Zeilen".
Private Function PID_PVLineCount(ByRef lines() As String) As Long
    On Error Resume Next
    PID_PVLineCount = UBound(lines) - LBound(lines) + 1
    If Err.Number <> 0 Then PID_PVLineCount = 0
    Err.Clear
End Function


Private Function PID_PVGetClipboardText() As String
    Dim dataObj As Object

    On Error Resume Next

    ' Spaete Bindung an MSForms.DataObject - so wird kein zusaetzlicher Verweis gebraucht.
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If dataObj Is Nothing Then Exit Function

    dataObj.GetFromClipboard
    PID_PVGetClipboardText = dataObj.GetText
End Function


' Waehrend eines selbst gesteuerten Einfuegens muss das Netz in EnforcePasteValuesOnly
' stillhalten: die Werte sind bereits sauber, und ein Application.Undo an dieser Stelle
' wuerde nur die Historie durcheinanderbringen.
Public Sub PID_PVBeginManagedPaste()
    mManagedPasteDepth = mManagedPasteDepth + 1
End Sub


Public Sub PID_PVEndManagedPaste()
    mManagedPasteDepth = mManagedPasteDepth - 1
    If mManagedPasteDepth < 0 Then mManagedPasteDepth = 0
End Sub


Public Function PID_IsManagedPasteRunning() As Boolean
    PID_IsManagedPasteRunning = (mManagedPasteDepth > 0)
End Function


Public Sub PID_ResetManagedPasteState()
    mManagedPasteDepth = 0
End Sub


Private Sub PID_PVInfo(ByVal message As String)
    MsgBox message, vbExclamation, "Einf" & PID_UTxtUe() & "gen"
End Sub

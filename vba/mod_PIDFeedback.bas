Attribute VB_Name = "mod_PIDFeedback"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Fehlermeldung aus der Datei heraus.
'
' Waehrend der Testphase melden Kollegen Probleme meist als "es funktioniert nicht".
' Dieses Makro sammelt in einem Schritt alles, was zur Nachstellung noetig ist:
' Version, Datei, Excel-Umgebung, aktives Blatt, Auswahl, Jahr, Rechenmodus und die
' letzten Aktionen aus dem Aktionsprotokoll - dazu zwei kurze Antworten des Benutzers.
'
' Ergebnis: Textdatei im Ordner "Feedback" neben der Mappe, zusaetzlich in der
' Zwischenablage (direkt in eine E-Mail einfuegbar).
'
' Version hier pflegen - einzige Stelle im VBA-Code.

Private Const PID_FB_VERSION As String = "v0.9.0-test"
Private Const PID_FB_FOLDER As String = "Feedback"
Private Const PID_FB_MAX_LOG_LINES As Long = 15
' Leer lassen = kein E-Mail-Schritt. Sonst Zieladresse fuer die Fehlermeldung.
Private Const PID_FB_MAIL_TO As String = "adam.nagy@at.mcd.com"
Private Const PID_FB_MAIL_MAX_BODY As Long = 1500


Public Sub FehlerMelden()
    PID_FehlerMelden
End Sub


Public Sub PID_AdminFehlerMelden()
    PID_FehlerMelden
End Sub


Public Function PID_GetAppVersion() As String
    PID_GetAppVersion = PID_FB_VERSION
End Function


Public Sub PID_FehlerMelden()
    Dim contextBlock As String
    Dim whatWanted As String
    Dim whatHappened As String
    Dim reportText As String
    Dim savedPath As String
    Dim clipboardOk As Boolean
    Dim resultMessage As String

    On Error GoTo CleanFail

    ' Kontext VOR den Dialogen einsammeln: danach kann sich die Auswahl geaendert haben.
    contextBlock = PID_FBBuildContextBlock()

    If Not PID_FBAskUser(whatWanted, whatHappened) Then Exit Sub

    reportText = PID_FBBuildReport(contextBlock, whatWanted, whatHappened)

    savedPath = PID_FBSaveReport(reportText)
    clipboardOk = PID_FBCopyToClipboard(reportText)

    PID_TrackAction "FehlerMelden", Left$(whatHappened, 80)

    If savedPath <> "" Then
        resultMessage = "Danke! Die Meldung wurde gespeichert:" & vbCrLf & savedPath
    Else
        resultMessage = "Die Meldung konnte nicht als Datei gespeichert werden."
    End If

    If clipboardOk Then
        resultMessage = resultMessage & vbCrLf & vbCrLf & _
                        "Der Text liegt auch in der Zwischenablage - du kannst ihn mit Strg+V " & _
                        "direkt in eine E-Mail einf" & PID_UTxtUe() & "gen."
    End If

    MsgBox resultMessage, vbInformation, PID_FBTitle()

    PID_FBOfferMail reportText, savedPath

    Exit Sub

CleanFail:
    MsgBox "Die Fehlermeldung konnte nicht erstellt werden:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, PID_FBTitle()
End Sub


Private Function PID_FBAskUser(ByRef outWanted As String, ByRef outHappened As String) As Boolean
    Dim answer As Variant

    On Error GoTo SafeExit

    answer = Application.InputBox( _
        "Was wolltest du gerade tun?" & vbCrLf & vbCrLf & _
        "Kurz und in eigenen Worten, z. B.: " & Chr$(34) & _
        "Mitarbeiter im Juli eingetragen und dann aktualisiert" & Chr$(34) & ".", _
        PID_FBTitle(), "", , , , , 2)

    If VarType(answer) = vbBoolean Then Exit Function
    outWanted = Trim$(CStr(answer))

    answer = Application.InputBox( _
        "Was ist passiert?" & vbCrLf & vbCrLf & _
        "Was war anders als erwartet? Wenn eine Fehlermeldung kam, bitte den Text abtippen.", _
        PID_FBTitle(), "", , , , , 2)

    If VarType(answer) = vbBoolean Then Exit Function
    outHappened = Trim$(CStr(answer))

    If outWanted = "" And outHappened = "" Then
        MsgBox "Es wurde nichts eingetragen - die Meldung wurde nicht erstellt.", _
               vbInformation, PID_FBTitle()
        Exit Function
    End If

    PID_FBAskUser = True

SafeExit:
End Function


Private Function PID_FBBuildContextBlock() As String
    Dim lines As String

    On Error Resume Next

    lines = "Zeitpunkt:   " & Format$(Now, "dd.mm.yyyy hh:nn:ss") & vbCrLf
    lines = lines & "Version:     " & PID_FB_VERSION & vbCrLf
    lines = lines & "Datei:       " & ThisWorkbook.Name & vbCrLf
    lines = lines & "Ordner:      " & ThisWorkbook.Path & vbCrLf
    lines = lines & "Excel:       " & Application.Version & " (Build " & Application.Build & ")" & vbCrLf
    lines = lines & "Benutzer:    " & Application.UserName & vbCrLf
    lines = lines & "Jahr:        " & PID_GetWorkbookYear() & vbCrLf
    lines = lines & "Blatt:       " & PID_FBActiveSheetName() & vbCrLf
    lines = lines & "Auswahl:     " & PID_FBSelectionAddress() & vbCrLf
    lines = lines & "Berechnung:  " & PID_FBCalculationText() & vbCrLf

    On Error GoTo 0

    PID_FBBuildContextBlock = lines
End Function


Private Function PID_FBBuildReport(ByVal contextBlock As String, _
                                   ByVal whatWanted As String, _
                                   ByVal whatHappened As String) As String
    Dim reportText As String
    Dim logText As String

    reportText = "PERSONALSHEET - Fehlermeldung" & vbCrLf
    reportText = reportText & "=============================" & vbCrLf & vbCrLf
    reportText = reportText & contextBlock & vbCrLf
    reportText = reportText & "Was wollte ich tun:" & vbCrLf & whatWanted & vbCrLf & vbCrLf
    reportText = reportText & "Was ist passiert:" & vbCrLf & whatHappened & vbCrLf & vbCrLf

    logText = PID_GetLastActionsText(PID_FB_MAX_LOG_LINES)

    If logText <> "" Then
        reportText = reportText & "Letzte Aktionen:" & vbCrLf & logText
    Else
        reportText = reportText & "Letzte Aktionen: (keine aufgezeichnet)" & vbCrLf
    End If

    PID_FBBuildReport = reportText
End Function


Private Function PID_FBSaveReport(ByVal reportText As String) As String
    Dim folderPath As String
    Dim filePath As String
    Dim fileNumber As Integer

    On Error GoTo SafeExit

    folderPath = PID_FBGetFeedbackFolder()
    If folderPath = "" Then Exit Function

    filePath = folderPath & "\Fehlermeldung_" & Format$(Now, "yyyy-mm-dd_hhnnss") & ".txt"

    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, reportText
    Close #fileNumber

    PID_FBSaveReport = filePath
    Exit Function

SafeExit:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
End Function


' Neben der Mappe ablegen. Bei OneDrive/SharePoint ist ThisWorkbook.Path eine URL -
' dort funktioniert MkDir nicht, deshalb der TEMP-Fallback.
Private Function PID_FBGetFeedbackFolder() As String
    Dim basePath As String
    Dim folderPath As String

    On Error GoTo UseTemp

    basePath = ThisWorkbook.Path

    If basePath = "" Or InStr(1, basePath, "http", vbTextCompare) = 1 Then GoTo UseTemp

    folderPath = basePath & "\" & PID_FB_FOLDER
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    PID_FBGetFeedbackFolder = folderPath
    Exit Function

UseTemp:
    On Error GoTo SafeExit

    folderPath = Environ$("TEMP") & "\" & PID_FB_FOLDER
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    PID_FBGetFeedbackFolder = folderPath

SafeExit:
End Function


' Zwischenablage ohne Verweis auf die Forms-Bibliothek (Windows-only Projekt).
Private Function PID_FBCopyToClipboard(ByVal textValue As String) As Boolean
    Dim dataObj As Object

    On Error GoTo SafeExit

    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText textValue
    dataObj.PutInClipboard

    PID_FBCopyToClipboard = True

SafeExit:
End Function


Private Sub PID_FBOfferMail(ByVal reportText As String, ByVal attachmentPath As String)
    On Error GoTo SafeExit

    If PID_FB_MAIL_TO = "" Then Exit Sub

    If MsgBox("Soll jetzt eine E-Mail an " & PID_FB_MAIL_TO & " ge" & PID_UTxtOe() & "ffnet werden?" & vbCrLf & vbCrLf & _
              "Die Mail wird nur vorbereitet - abgeschickt wird sie erst von dir.", _
              vbQuestion + vbYesNo, PID_FBTitle()) <> vbYes Then Exit Sub

    ' Outlook bevorzugt: dort landet der ganze Bericht im Text und die Datei als Anhang.
    If PID_FBOpenOutlookMail(reportText, attachmentPath) Then Exit Sub

    ' Kein Outlook verfuegbar: Standard-Mailprogramm ueber mailto (Text gekuerzt).
    If PID_FBOpenMailtoLink(reportText) Then Exit Sub

    MsgBox "Es konnte kein E-Mail-Programm ge" & PID_UTxtOe() & "ffnet werden." & vbCrLf & vbCrLf & _
           "Der Text liegt in der Zwischenablage - bitte manuell an " & PID_FB_MAIL_TO & " senden.", _
           vbExclamation, PID_FBTitle()

SafeExit:
End Sub


Private Function PID_FBOpenOutlookMail(ByVal reportText As String, _
                                       ByVal attachmentPath As String) As Boolean
    Dim outlookApp As Object
    Dim outlookMail As Object

    On Error GoTo SafeExit

    Set outlookApp = CreateObject("Outlook.Application")
    If outlookApp Is Nothing Then Exit Function

    ' 0 = olMailItem (spaete Bindung, kein Verweis auf die Outlook-Bibliothek noetig).
    Set outlookMail = outlookApp.CreateItem(0)
    If outlookMail Is Nothing Then Exit Function

    outlookMail.To = PID_FB_MAIL_TO
    outlookMail.Subject = "Personalsheet Fehlermeldung " & PID_FB_VERSION
    outlookMail.Body = reportText

    If attachmentPath <> "" Then
        On Error Resume Next
        outlookMail.Attachments.Add attachmentPath
        On Error GoTo SafeExit
    End If

    ' Display statt Send: der Benutzer sieht die Mail und schickt sie selbst ab.
    outlookMail.Display

    PID_FBOpenOutlookMail = True

SafeExit:
End Function


Private Function PID_FBOpenMailtoLink(ByVal reportText As String) As Boolean
    Dim mailUrl As String

    On Error GoTo SafeExit

    mailUrl = "mailto:" & PID_FB_MAIL_TO & _
              "?subject=" & PID_FBUrlEncode("Personalsheet Fehlermeldung " & PID_FB_VERSION) & _
              "&body=" & PID_FBUrlEncode(Left$(reportText, PID_FB_MAIL_MAX_BODY))

    ThisWorkbook.FollowHyperlink mailUrl

    PID_FBOpenMailtoLink = True

SafeExit:
End Function


Private Function PID_FBUrlEncode(ByVal textValue As String) As String
    Dim i As Long
    Dim charCode As Long
    Dim singleChar As String
    Dim resultText As String

    For i = 1 To Len(textValue)
        singleChar = Mid$(textValue, i, 1)
        charCode = Asc(singleChar)

        Select Case charCode
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                resultText = resultText & singleChar
            Case Else
                resultText = resultText & "%" & Right$("0" & Hex$(charCode And &HFF), 2)
        End Select
    Next i

    PID_FBUrlEncode = resultText
End Function


Private Function PID_FBActiveSheetName() As String
    On Error GoTo SafeExit

    If TypeName(ActiveSheet) = "Worksheet" Then PID_FBActiveSheetName = ActiveSheet.Name

SafeExit:
End Function


Private Function PID_FBSelectionAddress() As String
    On Error GoTo SafeExit

    If TypeName(Selection) = "Range" Then PID_FBSelectionAddress = Selection.Address(False, False)

SafeExit:
End Function


Private Function PID_FBCalculationText() As String
    Select Case Application.Calculation
        Case xlCalculationAutomatic
            PID_FBCalculationText = "Automatisch"
        Case xlCalculationManual
            PID_FBCalculationText = "Manuell"
        Case xlCalculationSemiautomatic
            PID_FBCalculationText = "Halbautomatisch"
        Case Else
            PID_FBCalculationText = "Unbekannt"
    End Select
End Function


Private Function PID_FBTitle() As String
    PID_FBTitle = "Fehler melden"
End Function

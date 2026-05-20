Attribute VB_Name = "mod_ResetAndImportVBAFiles"
Public Sub ResetAndImportVBAFiles()

    Dim vbProj As Object
    Dim vbComp As Object
    Dim fileName As String
    Dim vbaFolder As String
    Dim pathSeparator As String
    Dim i As Long
    Dim imported As Long
    Dim deleted As Long

    ' --- Betriebssystem erkennen ---
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        pathSeparator = "/"
    Else
        pathSeparator = "\"
    End If

    ' --- VBA-Ordner relativ zur Arbeitsmappe ---
    If ThisWorkbook.Path = "" Then
        MsgBox "Die Arbeitsmappe ist noch nicht gespeichert." & vbCrLf & vbCrLf & _
               "Bitte zuerst speichern, damit der Ordner ""vba"" neben der .xlsm gefunden werden kann.", _
               vbExclamation, "VBA Import"
        Exit Sub
    End If

    vbaFolder = ThisWorkbook.Path & pathSeparator & "vba" & pathSeparator

    If Dir(vbaFolder, vbDirectory) = "" Then
        MsgBox "VBA-Ordner nicht gefunden:" & vbCrLf & vbaFolder, vbExclamation, "VBA Import"
        Exit Sub
    End If

    On Error GoTo VBProjectBlocked
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0

    ' --- Standard-, Klassen- und UserForm-Module loeschen ---
    ' WICHTIG: Dieses Import-Modul bleibt erhalten
    For i = vbProj.VBComponents.Count To 1 Step -1

        Set vbComp = vbProj.VBComponents(i)

        Select Case vbComp.Type

            Case 1, 2, 3
                If vbComp.Name <> "mod_ResetAndImportVBAFiles" Then
                    vbProj.VBComponents.Remove vbComp
                    deleted = deleted + 1
                End If

            Case 100
                ' DieseArbeitsmappe / Tabellenmodule bleiben erhalten

        End Select

    Next i

    ' --- .bas Dateien importieren ---
    fileName = Dir(vbaFolder & "*.bas")

    Do While fileName <> ""

        If LCase(fileName) <> LCase("mod_ResetAndImportVBAFiles.bas") Then
            vbProj.VBComponents.Import vbaFolder & fileName
            imported = imported + 1
        End If

        fileName = Dir

    Loop

    ' --- .cls Dateien importieren ---
    fileName = Dir(vbaFolder & "*.cls")

    Do While fileName <> ""

        If LCase(fileName) <> LCase("DieseArbeitsmappe.cls") _
           And LCase(fileName) <> LCase("ThisWorkbook.cls") _
           And Left(LCase(fileName), 7) <> "tabelle" _
           And Left(LCase(fileName), 5) <> "sheet" Then

            vbProj.VBComponents.Import vbaFolder & fileName
            imported = imported + 1
        End If

        fileName = Dir

    Loop

    MsgBox deleted & " Module geloescht." & vbCrLf & _
           imported & " VBA-Dateien importiert.", vbInformation
    Exit Sub

VBProjectBlocked:
    MsgBox "Zugriff auf VBProject ist blockiert (Fehler " & Err.Number & ")." & vbCrLf & vbCrLf & _
           "Unter Windows muss in Excel aktiviert werden:" & vbCrLf & _
           "Datei > Optionen > Trust Center > Trust Center-Einstellungen > " & _
           "Makroeinstellungen > ""Zugriff auf das VBA-Projektobjektmodell vertrauen""" & vbCrLf & vbCrLf & _
           "Excel danach neu starten und dieses Makro erneut ausfuehren.", _
           vbExclamation, "VBA Import"

End Sub


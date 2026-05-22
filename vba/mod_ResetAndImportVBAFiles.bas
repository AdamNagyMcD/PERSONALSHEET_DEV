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
    Dim syncOk As Boolean
    Dim syncDetails As String

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

    PID_SyncDieseArbeitsmappeFromExport syncOk, syncDetails
    
    If syncOk Then
        MsgBox deleted & " Module geloescht." & vbCrLf & _
               imported & " VBA-Dateien importiert." & vbCrLf & _
               "DieseArbeitsmappe aus vba/DieseArbeitsmappe.cls synchronisiert.", vbInformation
    Else
        MsgBox deleted & " Module geloescht." & vbCrLf & _
               imported & " VBA-Dateien importiert." & vbCrLf & vbCrLf & _
               "Hinweis: DieseArbeitsmappe konnte nicht automatisch synchronisiert werden." & vbCrLf & _
               syncDetails, vbInformation
    End If
    Exit Sub

VBProjectBlocked:
    MsgBox "Zugriff auf VBProject ist blockiert (Fehler " & Err.Number & ")." & vbCrLf & vbCrLf & _
           "Unter Windows muss in Excel aktiviert werden:" & vbCrLf & _
           "Datei > Optionen > Trust Center > Trust Center-Einstellungen > " & _
           "Makroeinstellungen > ""Zugriff auf das VBA-Projektobjektmodell vertrauen""" & vbCrLf & vbCrLf & _
           "Excel danach neu starten und dieses Makro erneut ausfuehren.", _
           vbExclamation, "VBA Import"

End Sub


Public Sub PID_SyncDieseArbeitsmappeFromExport(Optional ByRef syncOk As Boolean = False, Optional ByRef syncDetails As String = "")
    Dim vbaFolder As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim foundCode As Boolean
    Dim codeText As String
    Dim codeModule As Object
    
    syncOk = False
    syncDetails = ""
    
    On Error GoTo SyncFail
    
    If ThisWorkbook.Path = "" Then
        syncDetails = "Arbeitsmappe ist noch nicht gespeichert."
        Exit Sub
    End If
    
    vbaFolder = ThisWorkbook.Path & Application.PathSeparator & "vba" & Application.PathSeparator
    filePath = vbaFolder & "DieseArbeitsmappe.cls"
    
    If Dir(filePath) = "" Then
        syncDetails = "Datei nicht gefunden: " & filePath
        Exit Sub
    End If
    
    foundCode = False
    codeText = ""
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        
        If Not foundCode Then
            If StrComp(Trim$(lineText), "Option Explicit", vbTextCompare) = 0 Then
                foundCode = True
                codeText = lineText
            End If
        Else
            codeText = codeText & vbCrLf & lineText
        End If
    Loop
    
    Close #fileNum
    
    If Not foundCode Then
        syncDetails = "Option Explicit nicht gefunden in DieseArbeitsmappe.cls."
        Exit Sub
    End If
    
    If InStr(1, codeText, "VERSION 1.0", vbTextCompare) > 0 _
       Or InStr(1, codeText, "Attribute VB_", vbTextCompare) > 0 Then
        syncDetails = "Export-Header erkannt. Bitte nur ab Option Explicit synchronisieren."
        Exit Sub
    End If
    
    Set codeModule = ThisWorkbook.VBProject.VBComponents("DieseArbeitsmappe").CodeModule
    
    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If
    
    codeModule.AddFromString codeText
    
    syncOk = True
    Exit Sub

SyncFail:
    syncDetails = Err.Number & " - " & Err.Description
End Sub


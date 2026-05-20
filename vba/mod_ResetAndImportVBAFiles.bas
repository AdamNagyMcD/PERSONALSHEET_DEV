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
    vbaFolder = ThisWorkbook.Path & pathSeparator & "vba" & pathSeparator

    Set vbProj = ThisWorkbook.VBProject

    ' --- Standard-, Klassen- und UserForm-Module lšschen ---
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

    MsgBox deleted & " Module gelšscht." & vbCrLf & _
           imported & " VBA-Dateien importiert.", vbInformation

End Sub


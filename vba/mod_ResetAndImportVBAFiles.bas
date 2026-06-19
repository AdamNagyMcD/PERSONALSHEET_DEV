Attribute VB_Name = "mod_ResetAndImportVBAFiles"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Public Sub ResetAndImportVBAFiles()

    Dim vbProj As Object
    Dim vbComp As Object
    Dim fileName As String
    Dim vbaFolder As String
    Dim pathSeparator As String
    Dim i As Long
    Dim imported As Long
    Dim deleted As Long
    Dim updatedCodeModules As Long
    Dim skipped As Long
    Dim compName As String

    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        pathSeparator = "/"
    Else
        pathSeparator = "\"
    End If

    If ThisWorkbook.Path = "" Then
        MsgBox "Die Arbeitsmappe ist noch nicht gespeichert." & vbCrLf & vbCrLf & _
               "Bitte zuerst speichern, damit der Ordner ""vba"" neben der .xlsm gefunden werden kann.", _
               vbExclamation, "VBA Import"
        Exit Sub
    End If

    vbaFolder = ThisWorkbook.Path & pathSeparator & "vba" & pathSeparator

    If Not FolderExistsVBA(vbaFolder) Then
        MsgBox "VBA-Ordner nicht gefunden:" & vbCrLf & vbaFolder, vbExclamation, "VBA Import"
        Exit Sub
    End If

    If Not PID_ConfirmAdminAction( _
        "Alle Standard-VBA-Module werden " & PID_UTxtGeloescht() & " und aus dem Ordner ""vba"" neu importiert.", _
        "VBA Import") Then
        Exit Sub
    End If

    On Error GoTo VBProjectBlocked
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo ImportError

    PID_FixLegacyModul11Name vbProj

    ' Standard-, Klassen- und UserForm-Module loeschen.
    ' mod_ResetAndImportVBAFiles bleibt erhalten.
    For i = vbProj.VBComponents.Count To 1 Step -1

        Set vbComp = vbProj.VBComponents(i)

        Select Case vbComp.Type

            Case 1, 2, 3
                If vbComp.Name <> "mod_ResetAndImportVBAFiles" Then
                    vbProj.VBComponents.Remove vbComp
                    deleted = deleted + 1
                End If

            Case 100
                ' DieseArbeitsmappe / Tabellenmodule bleiben erhalten.
                ' Der Code wird spaeter aus den passenden .cls Dateien aktualisiert.

        End Select

    Next i

    fileName = Dir(vbaFolder & "*.bas")

    Do While fileName <> ""

        If LCase$(fileName) <> LCase$("mod_ResetAndImportVBAFiles.bas") Then
            compName = GetVBNameFromBasFile(vbaFolder & fileName)
            If Len(compName) > 0 Then
                PID_RemoveVBComponentAndNumberedCopies vbProj, compName
            End If
            vbProj.VBComponents.Import vbaFolder & fileName
            imported = imported + 1
        Else
            skipped = skipped + 1
        End If

        fileName = Dir

    Loop

    fileName = Dir(vbaFolder & "*.cls")

    Do While fileName <> ""

        If ShouldSkipClsImportFile(fileName) Then
            skipped = skipped + 1
        Else
            compName = FileNameWithoutExtension(fileName)

            If ComponentExists(vbProj, compName) Then

                Set vbComp = vbProj.VBComponents(compName)

                If vbComp.Type = 100 Then
                    If UpdateCodeModuleFromFile(vbComp, vbaFolder & fileName, True) Then
                        updatedCodeModules = updatedCodeModules + 1
                    End If
                Else
                    skipped = skipped + 1
                End If

            Else
                vbProj.VBComponents.Import vbaFolder & fileName
                imported = imported + 1
            End If
        End If

        fileName = Dir

    Loop

    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    Application.StatusBar = False
    On Error GoTo ImportError

    MsgBox deleted & " Module " & PID_UTxtGeloescht() & ", " & imported & " importiert, " & _
           updatedCodeModules & " aktualisiert." & vbCrLf & vbCrLf & _
           "Danach: Kompilieren > Speichern > Excel neu starten > FullSystemRefresh", _
           vbInformation, "VBA Import"
    
    On Error Resume Next
    PID_RestoreLetztesGehaltFormulasSilent
    Err.Clear

    Exit Sub

VBProjectBlocked:
    MsgBox "Zugriff auf VBProject ist blockiert (Fehler " & Err.Number & ")." & vbCrLf & vbCrLf & _
           "Unter Windows muss in Excel aktiviert werden:" & vbCrLf & _
           "Datei > Optionen > Trust Center > Trust Center-Einstellungen > " & _
           "Makroeinstellungen > ""Zugriff auf das VBA-Projektobjektmodell vertrauen""" & vbCrLf & vbCrLf & _
           "Excel danach neu starten und dieses Makro erneut " & PID_UTxtAusfuehren() & ".", _
           vbExclamation, "VBA Import"
    Exit Sub

ImportError:
    MsgBox "Fehler beim VBA-Import:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbExclamation, "VBA Import"

End Sub


Public Sub PID_RepairWorkbookAfterVBAImport()
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    Dim wbName As String
    
    On Error GoTo CleanFail
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    wbName = ThisWorkbook.Name
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Application.Run "'" & wbName & "'!PID_RestoreMonatslohnFormulasSilent"
    Application.Run "'" & wbName & "'!PID_RestoreAktuelleStundenFormulasSilent"
    Application.Run "'" & wbName & "'!PID_RestoreUrlaubGeldFormulasSilent"
    Application.Run "'" & wbName & "'!PID_RestoreLetztesGehaltFormulasSilent"
    Application.Run "'" & wbName & "'!PID_RestoreMonthSheetDropdownsAfterFormatSilent"
    Application.CalculateFull
    Err.Clear
    On Error GoTo CleanFail
    
    GoTo CleanExit

CleanFail:
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
End Sub


Public Sub PID_SyncDieseArbeitsmappeFromExport(Optional ByRef syncOk As Boolean = False, Optional ByRef syncDetails As String = "")
    Dim vbComp As Object

    syncOk = False
    syncDetails = ""

    On Error GoTo SyncFail

    If ThisWorkbook.Path = "" Then
        syncDetails = "Arbeitsmappe ist noch nicht gespeichert."
        Exit Sub
    End If

    Set vbComp = ThisWorkbook.VBProject.VBComponents("DieseArbeitsmappe")

    syncOk = UpdateCodeModuleFromFile(vbComp, _
        ThisWorkbook.Path & Application.PathSeparator & "vba" & Application.PathSeparator & "DieseArbeitsmappe.cls")

    If Not syncOk Then
        syncDetails = "Keine Aenderung oder Datei nicht gefunden."
    End If

    Exit Sub

SyncFail:
    syncDetails = Err.Number & " - " & Err.Description
End Sub


Public Sub SyncDieseArbeitsmappeFromExport()
    Dim syncOk As Boolean
    Dim syncDetails As String
    
    PID_SyncDieseArbeitsmappeFromExport syncOk, syncDetails
    
    If syncOk Then
        MsgBox "DieseArbeitsmappe wurde aus vba/DieseArbeitsmappe.cls synchronisiert.", vbInformation, "VBA Sync"
    Else
        MsgBox "Synchronisation fehlgeschlagen:" & vbCrLf & syncDetails, vbExclamation, "VBA Sync"
    End If
End Sub


Private Sub PID_RemoveVBComponentAndNumberedCopies(ByVal vbProj As Object, ByVal baseName As String)
    Dim i As Long
    Dim compName As String
    Dim suffix As String
    
    If Len(Trim$(baseName)) = 0 Then Exit Sub
    
    For i = vbProj.VBComponents.Count To 1 Step -1
        compName = vbProj.VBComponents(i).Name
        
        If StrComp(compName, baseName, vbTextCompare) = 0 Then
            vbProj.VBComponents.Remove vbProj.VBComponents(i)
        ElseIf Len(compName) > Len(baseName) Then
            If Left$(compName, Len(baseName)) = baseName Then
                suffix = Mid$(compName, Len(baseName) + 1)
                If Len(suffix) > 0 And IsNumeric(suffix) Then
                    vbProj.VBComponents.Remove vbProj.VBComponents(i)
                End If
            End If
        End If
    Next i
End Sub


Private Function GetVBNameFromBasFile(ByVal fullPath As String) As String
    Dim f As Integer
    Dim lineText As String
    Dim trimmedLine As String
    Dim eqPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long
    
    GetVBNameFromBasFile = ""
    If Dir(fullPath) = "" Then Exit Function
    
    f = FreeFile
    Open fullPath For Input As #f
    
    Do Until EOF(f)
        Line Input #f, lineText
        trimmedLine = Trim$(Replace(lineText, ChrW(65279), ""))
        
        If Left$(trimmedLine, 16) = "Attribute VB_Name" Then
            eqPos = InStr(trimmedLine, "=")
            quoteStart = InStr(eqPos, trimmedLine, """")
            quoteEnd = InStr(quoteStart + 1, trimmedLine, """")
            If quoteStart > 0 And quoteEnd > quoteStart Then
                GetVBNameFromBasFile = Mid$(trimmedLine, quoteStart + 1, quoteEnd - quoteStart - 1)
            End If
            Exit Do
        End If
    Loop
    
    Close #f
End Function


Private Function PID_FixLegacyModul11Name(ByVal vbProj As Object) As Boolean
    If Not ComponentExists(vbProj, "Modul11") Then Exit Function
    If ComponentExists(vbProj, "mod_BuildDurchrechnung") Then Exit Function
    
    vbProj.VBComponents("Modul11").Name = "mod_BuildDurchrechnung"
    PID_FixLegacyModul11Name = True
End Function


Private Function ShouldSkipClsImportFile(ByVal fileName As String) As Boolean
    Dim lowerName As String

    lowerName = LCase$(fileName)

    If lowerName = "thisworkbook.cls" Then
        ShouldSkipClsImportFile = True
        Exit Function
    End If

    If Left$(lowerName, 7) = "tabelle" Then
        ShouldSkipClsImportFile = True
        Exit Function
    End If

    If Left$(lowerName, 5) = "sheet" Then
        ShouldSkipClsImportFile = True
        Exit Function
    End If

    ShouldSkipClsImportFile = False
End Function


Private Function FolderExistsVBA(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExistsVBA = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function


Private Function ComponentExists(ByVal vbProj As Object, ByVal componentName As String) As Boolean
    Dim tempComp As Object

    On Error Resume Next
    Set tempComp = vbProj.VBComponents(componentName)
    ComponentExists = Not tempComp Is Nothing
    On Error GoTo 0
End Function


Private Function FileNameWithoutExtension(ByVal fileName As String) As String
    Dim p As Long

    p = InStrRev(fileName, ".")
    If p > 0 Then
        FileNameWithoutExtension = Left$(fileName, p - 1)
    Else
        FileNameWithoutExtension = fileName
    End If
End Function


Private Function UpdateCodeModuleFromFile(ByVal vbComp As Object, ByVal fullPath As String, _
                                          Optional ByVal forceUpdate As Boolean = False) As Boolean
    Dim newCode As String
    Dim oldCode As String
    Dim cm As Object

    If Dir(fullPath) = "" Then Exit Function

    newCode = ReadVBAFileWithoutAttributes(fullPath)
    If Len(Trim$(newCode)) = 0 Then Exit Function

    Set cm = vbComp.CodeModule

    If cm.CountOfLines > 0 Then
        oldCode = cm.Lines(1, cm.CountOfLines)
    Else
        oldCode = ""
    End If

    If Not forceUpdate Then
        If NormalizeCodeText(oldCode) = NormalizeCodeText(newCode) Then
            UpdateCodeModuleFromFile = False
            Exit Function
        End If
    End If

    If cm.CountOfLines > 0 Then
        cm.DeleteLines 1, cm.CountOfLines
    End If

    cm.InsertLines 1, newCode

    UpdateCodeModuleFromFile = True
End Function


Private Function ReadVBAFileWithoutAttributes(ByVal fullPath As String) As String
    Dim f As Integer
    Dim lineText As String
    Dim result As String
    Dim foundCode As Boolean
    Dim trimmedLine As String

    f = FreeFile
    Open fullPath For Input As #f

    Do Until EOF(f)

        Line Input #f, lineText

        lineText = Replace(lineText, ChrW(65279), "")
        trimmedLine = Trim$(lineText)

        If Not foundCode Then
            ' Export-Header (VERSION/BEGIN/Attribute) wird uebersprungen.
            ' Erst ab Option Explicit gehoert Code ins Modul.
            If StrComp(trimmedLine, "Option Explicit", vbTextCompare) = 0 Then
                foundCode = True
                result = lineText & vbCrLf
            End If
        Else
            result = result & lineText & vbCrLf
        End If

    Loop

    Close #f

    ReadVBAFileWithoutAttributes = result
End Function


Private Function NormalizeCodeText(ByVal txt As String) As String
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)

    Do While Len(txt) > 0
        If Right$(txt, 1) = vbLf Or Right$(txt, 1) = " " Then
            txt = Left$(txt, Len(txt) - 1)
        Else
            Exit Do
        End If
    Loop

    NormalizeCodeText = txt
End Function

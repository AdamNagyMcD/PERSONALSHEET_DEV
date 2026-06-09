VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range

    On Error GoTo SafeExit

    ' Fuer Monatsblaetter uebernimmt DieseArbeitsmappe.Workbook_SheetChange den Einfuege-Schutz.
    If PID_IsWorkerMonthSheet(Me) Then Exit Sub

    If Application.CutCopyMode <> False Then
        Application.EnableEvents = False

        Set rng = Target

        ' Einfuegen rueckgaengig machen, damit das Zellenformat erhalten bleibt
        Application.Undo

        ' Nur Wert an gleicher Stelle einfuegen
        rng.PasteSpecial Paste:=xlPasteValues
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

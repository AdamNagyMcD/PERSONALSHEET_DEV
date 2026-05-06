VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit

    If Application.CutCopyMode <> False Then
        Application.EnableEvents = False

        Dim rng As Range
        Set rng = Target

        ' Beilleszt_s visszavonösa, hogy a cella formötuma megmaradjon
        Application.Undo

        ' Csak _rt_k vissza ugyanoda
        rng.PasteSpecial Paste:=xlPasteValues
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

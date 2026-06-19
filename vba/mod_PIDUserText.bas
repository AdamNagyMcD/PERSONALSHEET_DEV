Attribute VB_Name = "mod_PIDUserText"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' Zentrale Benutzertexte mit Umlauten (ChrW = ASCII-sichere Quelle, Win/Mac Excel 2016).
' Nur fuer sichtbare Dialoge/Meldungen — keine Literal-Umlaute in .bas-Dateien.


Public Function PID_UTxtAe() As String
    PID_UTxtAe = ChrW(228)
End Function


Public Function PID_UTxtOe() As String
    PID_UTxtOe = ChrW(246)
End Function


Public Function PID_UTxtUe() As String
    PID_UTxtUe = ChrW(252)
End Function


Public Function PID_UTxtSs() As String
    PID_UTxtSs = ChrW(223)
End Function


Public Function PID_UTxtLoeschen() As String
    PID_UTxtLoeschen = "l" & PID_UTxtOe() & "schen"
End Function


Public Function PID_UTxtGeloescht() As String
    PID_UTxtGeloescht = "gel" & PID_UTxtOe() & "scht"
End Function


Public Function PID_UTxtGeloeschtWerdenLabel() As String
    PID_UTxtGeloeschtWerdenLabel = "Gel" & PID_UTxtOe() & "scht werden:"
End Function


Public Function PID_UTxtGueltig() As String
    PID_UTxtGueltig = "g" & PID_UTxtUe() & "ltig"
End Function


Public Function PID_UTxtGueltigen() As String
    PID_UTxtGueltigen = PID_UTxtGueltig() & "en"
End Function


Public Function PID_UTxtAuswaehlen() As String
    PID_UTxtAuswaehlen = "ausw" & PID_UTxtAe() & "hlen"
End Function


Public Function PID_UTxtAusgewaehlte() As String
    PID_UTxtAusgewaehlte = "Ausgew" & PID_UTxtAe() & "hlte"
End Function


Public Function PID_UTxtDatenLoeschen() As String
    PID_UTxtDatenLoeschen = "Daten " & PID_UTxtLoeschen()
End Function


Public Function PID_UTxtZeilenLoeschen() As String
    PID_UTxtZeilenLoeschen = "Zeilen " & PID_UTxtLoeschen()
End Function


Public Function PID_UTxtMonatsdatenLoeschen() As String
    PID_UTxtMonatsdatenLoeschen = "Monatsdaten " & PID_UTxtLoeschen()
End Function


Public Function PID_UTxtAusgewaehlteZeilenLoeschen() As String
    PID_UTxtAusgewaehlteZeilenLoeschen = PID_UTxtAusgewaehlte() & " Zeilen " & PID_UTxtLoeschen()
End Function


Public Function PID_UTxtEintraege() As String
    PID_UTxtEintraege = "Eintr" & PID_UTxtAe() & "ge"
End Function


Public Function PID_UTxtBlaetter() As String
    PID_UTxtBlaetter = "Bl" & PID_UTxtAe() & "tter"
End Function


Public Function PID_UTxtMonatsblaetter() As String
    PID_UTxtMonatsblaetter = "Monatsbl" & PID_UTxtAe() & "tter"
End Function


Public Function PID_UTxtMonatsblaettern() As String
    PID_UTxtMonatsblaettern = PID_UTxtMonatsblaetter() & "n"
End Function


Public Function PID_UTxtZurueckgesetzt() As String
    PID_UTxtZurueckgesetzt = "zur" & PID_UTxtUe() & "ckgesetzt"
End Function


Public Function PID_UTxtPruefung() As String
    PID_UTxtPruefung = "Pr" & PID_UTxtUe() & "fung"
End Function


Public Function PID_UTxtAenderung() As String
    PID_UTxtAenderung = ChrW(196) & "nderung"
End Function


Public Function PID_UTxtFuer() As String
    PID_UTxtFuer = "f" & PID_UTxtUe() & "r"
End Function


Public Function PID_UTxtVollstaendig() As String
    PID_UTxtVollstaendig = "vollst" & PID_UTxtAe() & "ndig"
End Function


Public Function PID_UTxtAusfuehren() As String
    PID_UTxtAusfuehren = "ausf" & PID_UTxtUe() & "hren"
End Function


Public Function PID_UTxtGeschuetzt() As String
    PID_UTxtGeschuetzt = "gesch" & PID_UTxtUe() & "tzt"
End Function

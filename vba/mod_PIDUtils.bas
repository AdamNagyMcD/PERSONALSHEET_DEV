Attribute VB_Name = "mod_PIDUtils"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

' True waehrend PID_CopyDataToFollowingMonths laeuft.
' Schuetzt den Hour-Override-Log davor, dass Mac Excel Worksheet_Change-Events
' auch bei Application.EnableEvents=False feuert und dabei spaeteren Log-Eintraege loescht.
' Windows: EnableEvents=False unterdrueckt Events zuverlaessig → Guard wird nie aktiv.
Public gCopyDataRunning As Boolean


Public Function PID_CollectionHasKey(ByVal col As Collection, ByVal key As String) As Boolean
    Dim tmp As Variant
    
    On Error GoTo NotFound
    
    tmp = col.item(key)
    
    PID_CollectionHasKey = True
    Exit Function

NotFound:
    PID_CollectionHasKey = False
End Function


Public Function PID_MonthNames() As Variant
    PID_MonthNames = Array("Januar", "Februar", "Marz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
End Function


Public Function PID_GetMonthIndexFromName(ByVal monthName As String) As Long
    Select Case UCase$(Trim$(CStr(monthName)))
        Case "JANUAR"
            PID_GetMonthIndexFromName = 1
        Case "FEBRUAR"
            PID_GetMonthIndexFromName = 2
        Case "MARZ", "M" & ChrW(228) & "RZ", "MAERZ"
            PID_GetMonthIndexFromName = 3
        Case "APRIL"
            PID_GetMonthIndexFromName = 4
        Case "MAI"
            PID_GetMonthIndexFromName = 5
        Case "JUNI"
            PID_GetMonthIndexFromName = 6
        Case "JULI"
            PID_GetMonthIndexFromName = 7
        Case "AUGUST"
            PID_GetMonthIndexFromName = 8
        Case "SEPTEMBER"
            PID_GetMonthIndexFromName = 9
        Case "OKTOBER"
            PID_GetMonthIndexFromName = 10
        Case "NOVEMBER"
            PID_GetMonthIndexFromName = 11
        Case "DEZEMBER"
            PID_GetMonthIndexFromName = 12
        Case Else
            PID_GetMonthIndexFromName = 0
    End Select
End Function


Public Function PID_GetMonthIndexFromSheetName(ByVal sheetName As String) As Long
    PID_GetMonthIndexFromSheetName = PID_GetMonthIndexFromName(sheetName)
End Function

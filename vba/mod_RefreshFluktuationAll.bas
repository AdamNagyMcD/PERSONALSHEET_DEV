Attribute VB_Name = "mod_RefreshFluktuationAll"
'==============================================================================
' Personalsheet – VBA
' Copyright (c) Adam Nagy / McOpCo. All rights reserved.
' Unauthorized copying, modification or distribution prohibited.
'==============================================================================

Option Explicit

Public gFluktuationDirty As Boolean
Public gFluktuationDataDirty As Boolean
Public gFluktuationAnalyseDirty As Boolean

Private mFluktuationDirtyScopeAll As Boolean
Private mFluktuationDirtyMonths As Collection


Public Sub MarkFluktuationDirty()
    gFluktuationDataDirty = True
    gFluktuationAnalyseDirty = True
    gFluktuationDirty = True
    mFluktuationDirtyScopeAll = True
    Set mFluktuationDirtyMonths = New Collection
End Sub


Public Sub MarkFluktuationDirtyForMonth(ByVal monthNumber As Long)
    If monthNumber < 1 Or monthNumber > 12 Then
        MarkFluktuationDirty
        Exit Sub
    End If
    
    gFluktuationDataDirty = True
    gFluktuationAnalyseDirty = True
    gFluktuationDirty = True
    mFluktuationDirtyScopeAll = False
    
    If mFluktuationDirtyMonths Is Nothing Then Set mFluktuationDirtyMonths = New Collection
    
    If Not PID_CollectionHasKey(mFluktuationDirtyMonths, CStr(monthNumber)) Then
        mFluktuationDirtyMonths.Add monthNumber, CStr(monthNumber)
    End If
End Sub


Public Sub MarkFluktuationDirtyForMonthSheet(ByVal wsMonth As Worksheet)
    Dim monthNumber As Long
    
    If wsMonth Is Nothing Then
        MarkFluktuationDirty
        Exit Sub
    End If
    
    If Not PID_IsWorkerMonthSheet(wsMonth) Then
        MarkFluktuationDirty
        Exit Sub
    End If
    
    If IsNumeric(wsMonth.Range("A1").Value) Then
        monthNumber = CLng(wsMonth.Range("A1").Value)
    Else
        monthNumber = PID_GetMonthIndexFromName(wsMonth.Name)
    End If
    
    If monthNumber < 1 Or monthNumber > 12 Then
        MarkFluktuationDirty
        Exit Sub
    End If
    
    MarkFluktuationDirtyForMonth monthNumber
End Sub


Public Sub MarkFluktuationClean()
    gFluktuationDataDirty = False
    gFluktuationAnalyseDirty = False
    gFluktuationDirty = False
    mFluktuationDirtyScopeAll = False
    Set mFluktuationDirtyMonths = New Collection
End Sub


Public Function IsFluktuationDirty() As Boolean
    IsFluktuationDirty = gFluktuationDataDirty Or gFluktuationAnalyseDirty
End Function


Public Function PID_FluktuationUseIncrementalDataRebuild() As Boolean
    PID_FluktuationUseIncrementalDataRebuild = (Not mFluktuationDirtyScopeAll) _
        And (Not mFluktuationDirtyMonths Is Nothing) _
        And (mFluktuationDirtyMonths.Count > 0)
End Function


Public Function PID_FluktuationMonthDisplayNameIsDirty(ByVal monthDisplayName As String) As Boolean
    Dim i As Long
    Dim monthNumber As Long
    
    monthDisplayName = Trim$(CStr(monthDisplayName))
    If monthDisplayName = "" Then Exit Function
    
    If mFluktuationDirtyScopeAll Then
        PID_FluktuationMonthDisplayNameIsDirty = True
        Exit Function
    End If
    
    If mFluktuationDirtyMonths Is Nothing Then Exit Function
    
    For i = 1 To mFluktuationDirtyMonths.Count
        monthNumber = CLng(mFluktuationDirtyMonths(i))
        If StrComp(PID_FlDisplayMonthName(monthNumber), monthDisplayName, vbTextCompare) = 0 Then
            PID_FluktuationMonthDisplayNameIsDirty = True
            Exit Function
        End If
    Next i
End Function


Public Function PID_FluktuationDirtyMonthsClone() As Collection
    Dim cloneMonths As Collection
    Dim i As Long
    
    Set cloneMonths = New Collection
    
    If mFluktuationDirtyMonths Is Nothing Then
        Set PID_FluktuationDirtyMonthsClone = cloneMonths
        Exit Function
    End If
    
    For i = 1 To mFluktuationDirtyMonths.Count
        cloneMonths.Add mFluktuationDirtyMonths(i)
    Next i
    
    Set PID_FluktuationDirtyMonthsClone = cloneMonths
End Function


Public Sub PID_ClearFluktuationDirtyMonthsAfterDataRebuild()
    mFluktuationDirtyScopeAll = False
    Set mFluktuationDirtyMonths = New Collection
End Sub


Public Sub RefreshFluktuationDataIfDirty()
    On Error GoTo CleanFail
    
    If Not gFluktuationDataDirty Then Exit Sub
    
    BuildFluktuationDaten
    PID_ClearFluktuationDirtyMonthsAfterDataRebuild
    gFluktuationDataDirty = False
    
    If Not gFluktuationAnalyseDirty Then
        gFluktuationDirty = False
    End If
    
    On Error Resume Next
    ThisWorkbook.Worksheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    Exit Sub

CleanFail:
    gFluktuationDataDirty = True
    gFluktuationDirty = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub RefreshFluktuationIfDirty()
    If gFluktuationDataDirty Or gFluktuationAnalyseDirty Then
        RefreshFluktuationAllPartial gFluktuationDataDirty, gFluktuationAnalyseDirty
    End If
End Sub


Public Sub RefreshFluktuationNow()
    RefreshFluktuationAll
End Sub


Public Sub RefreshFluktuationAll()
    RefreshFluktuationAllPartial True, True
End Sub


Private Sub RefreshFluktuationAllPartial(ByVal refreshData As Boolean, ByVal refreshAnalyse As Boolean)
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    
    On Error GoTo CleanFail
    
    If Not refreshData And Not refreshAnalyse Then Exit Sub
    
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Fluktuation wird aktualisiert..."
    
    If refreshData Then
        BuildFluktuationDaten
        PID_ClearFluktuationDirtyMonthsAfterDataRebuild
        gFluktuationDataDirty = False
    End If
    
    If refreshAnalyse Then
        BuildFluktuationAnalyse
        gFluktuationAnalyseDirty = False
    End If
    
    On Error Resume Next
    ThisWorkbook.Worksheets("FLUKTUATION_DATEN").Visible = xlSheetVeryHidden
    On Error GoTo CleanFail
    
    gFluktuationDirty = (gFluktuationDataDirty Or gFluktuationAnalyseDirty)
    If Not gFluktuationDirty Then MarkFluktuationClean

CleanExit:
    Application.StatusBar = oldStatusBar
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

CleanFail:
    gFluktuationDataDirty = refreshData Or gFluktuationDataDirty
    gFluktuationAnalyseDirty = refreshAnalyse Or gFluktuationAnalyseDirty
    gFluktuationDirty = True
    
    MsgBox "Die Fluktuation konnte nicht vollstaendig aktualisiert werden." & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Fluktuation aktualisieren"
    
    Resume CleanExit
End Sub

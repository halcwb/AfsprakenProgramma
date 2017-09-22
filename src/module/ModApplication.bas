Attribute VB_Name = "ModApplication"
Option Explicit

Private blnDontClose As Boolean
Private blnCloseHaseRun As Boolean

Private Const constVersie As String = "Var_Glob_AppVersie"
Private Const constDate As String = "Var_AfspraakDatum"

Private Const constBarDel As String = " | "

Public Enum EnumAppLanguage
    Dutch = 1043
    English = 0
End Enum

Public Sub SetDontClose(ByVal blnClose As Boolean)
    
    blnDontClose = blnClose

End Sub

Public Function Application_GetVersion() As String

    Application_GetVersion = ModRange.GetRangeValue(constVersie, vbNullString)

End Function

Public Sub SetToDevelopmentMode()

    Dim blnDevelop As Boolean
    Dim objWindow As Window
    
    blnDevelop = Not ModSetting.IsDevelopmentMode()
    
    shtGlobGuiFront.Select
    
    If blnDevelop Then
        ModProgress.StartProgress "Zet in Ontwikkel Modus"
        Application.ScreenUpdating = False
        
        ModSheet.UnprotectUserInterfaceSheets True
        ModSheet.UnhideNonUserInterfaceSheets True
                
        ModSetting.SetDevelopmentMode True
        For Each objWindow In WbkAfspraken.Windows
            SetWindowToCloseApp objWindow
        Next
        
        Application.ScreenUpdating = True
        ModProgress.FinishProgress
        
        Application.DisplayFormulaBar = True
    Else
        ModMessage.ShowMsgBoxInfo "Weer terug zetten in Gebruikers Modus"
        ModSetting.SetDevelopmentMode False
        
        InitializeAfspraken
    End If

End Sub

Public Sub CloseAfspraken()

    Dim strAction As String
    Dim strParams() As Variant
    
    Dim intN As Integer
    Dim intC As Integer
    
    Dim objWindow As Window
    
    If blnCloseHaseRun Then ' Second CloseAfspraken triggert by WbkAfspraken.Workbook_BeforeClose
        Exit Sub
    End If
    
    If Application.Workbooks.Count > 1 Then
        ModMessage.ShowMsgBoxExclam "Er zijn nog andere Excel bestanden geopend, sla deze eerst op anders worden deze niet opgeslagen!"
        WbkAfspraken.blnCancelAfsprakenClose = True
        Exit Sub
    End If
    
    shtGlobGuiFront.Select
    strAction = "ModApplication.CloseAfspraken"
    
    ModLog.LogActionStart strAction, strParams
    
    ModProgress.StartProgress "Afspraken Programma Afsluiten"
    
    intN = 1
    intC = WbkAfspraken.Windows.Count
    For Each objWindow In WbkAfspraken.Windows
        SetWindowToCloseApp objWindow
        ModProgress.SetJobPercentage "Windows Terugzetten", intC, intN
        intN = intN + 1
    Next
 
    Toolbars("Afspraken").Visible = False
    
    With Application
         .Caption = vbNullString
         .DisplayFormulaBar = True
         .Cursor = xlDefault
    End With
    
    ModSetting.SetDevelopmentMode False
        
    ModProgress.FinishProgress
    ModLog.LogActionEnd strAction
    
    blnCloseHaseRun = True
            
    If Not blnDontClose Then
        Application.StatusBar = vbNullString
        Application.DisplayAlerts = False
        WbkAfspraken.blnCancelAfsprakenClose = False
        Application.Quit
    End If

End Sub

Private Sub TestCloseAfspraken()
    blnDontClose = True
    CloseAfspraken
    MsgBox Application.DisplayAlerts
End Sub

Private Sub SetWindow(objWindow As Window, ByVal blnDisplay As Boolean)

    With objWindow
        .DisplayWorkbookTabs = blnDisplay
        .DisplayGridlines = blnDisplay
        .DisplayHeadings = blnDisplay
        .DisplayOutline = blnDisplay
        .DisplayZeros = blnDisplay
        .DisplayVerticalScrollBar = True
        .DisplayHorizontalScrollBar = blnDisplay
        .WindowState = xlMaximized
    End With
    
    Application.DisplayFormulaBar = blnDisplay
    
End Sub

Public Sub SetWindowToCloseApp(objWindow As Window)
    
    SetWindow objWindow, True

End Sub

Public Sub SetWindowToOpenApp(objWindow As Window)
    
    SetWindow objWindow, False

End Sub

Public Sub InitializeAfspraken()

    Dim strError As String
    Dim strAction As String
    Dim strBed As String
    Dim strParams() As Variant
    Dim objWind As Window
    
    On Error GoTo InitializeError
    
    strAction = "ModApplication.InitializeAfspraken"
    
    ModLog.LogActionStart strAction, strParams
    
    shtGlobGuiFront.Select
    DoEvents                           ' Make sure sheet is shown before proceding
        
    SetCaptionAndHideBars              ' Setup Excel Application
    
    For Each objWind In WbkAfspraken.Windows
        SetWindowToOpenApp objWind
    Next
        
    Application.Visible = True
    ModProgress.StartProgress "Start Afspraken Programma"
            
    Application.ScreenUpdating = False ' Prevent cycling through all windows when sheets are processed
    
    ' Setup sheets
    ModSheet.ProtectUserInterfaceSheets True
    ModSheet.HideAndUnProtectNonUserInterfaceSheets True
    
    Application.ScreenUpdating = True
    
    ' Clean everything
    ModRange.SetRangeValue constVersie, vbNullString
    ModSetting.SetDevelopmentMode False    ' Default development mode is false
            
    strBed = ModMetaVision.MetaVision_GetCurrentBedName()
    If strBed <> vbNullString Then
        ModBed.SetBed strBed
        ModBed.OpenBedAsk False, True
    Else
        ModPatient.PatientClearAll False, True ' Default start with no patient
    End If
    
    ModSheet.SelectPedOrNeoStartSheet False  ' Select the first GUI sheet
    
    ModProgress.FinishProgress
    ModLog.LogActionEnd strAction
            
    Exit Sub
    
InitializeError:
    
    ModProgress.FinishProgress
    Application.Visible = True

    strError = "Kan de applicatie niet opstarten"
    ModMessage.ShowMsgBoxError strError
    
    strError = strError & vbNewLine & strAction
    ModLog.LogError strError
    
End Sub

Public Sub UpdateStatusBar(ByVal strItem As String, ByVal strMessage As String)

    Dim varStatus() As String
    Dim varItem() As String
    Dim intN As Integer
    Dim intC As Integer
    Dim blnItemSet As Boolean
    
    varStatus = Split(Application.StatusBar, constBarDel)
    intC = UBound(varStatus)
    blnItemSet = False
    
    For intN = 0 To intC
        varItem = Split(varStatus(intN), ":")
        If UBound(varItem) > 0 Then
            If Trim(varItem(0)) = Trim(strItem) Then
                varStatus(intN) = strItem & ": " & strMessage
                blnItemSet = True
                Exit For
            End If
        End If
    Next
    
    Application.StatusBar = Join(varStatus, constBarDel)
    
    If Not blnItemSet Then Application.StatusBar = Application.StatusBar & " " & constBarDel & " " & strItem & ": " & strMessage

End Sub

Public Sub TestUpdateStatusBar()

    Application.StatusBar = " "
    UpdateStatusBar "Setting", "Test2"
    
End Sub

Public Sub SetDateToDayFormula()

    ModRange.SetRangeFormula constDate, GetToDayFormula()

End Sub

Private Sub SetCaptionAndHideBars()

    Dim blnIsDevelop As Boolean
    
    blnIsDevelop = ModSetting.IsDevelopmentMode()
    
    SetApplicationTitle
    
    With Application
        .DisplayFormulaBar = blnIsDevelop
        .DisplayStatusBar = True
        .DisplayFullScreen = False
        .DisplayScrollBars = True
        .WindowState = xlMaximized
    End With
    
    Application.StatusBar = ModConst.CONST_APPLICATION_NAME
    UpdateStatusBar "Versie", ModRange.GetRangeValue("Var_Glob_AppVersie", vbNullString)
    UpdateStatusBar "Omgeving", GetEnvironment()
    UpdateStatusBar "Afdeling", IIf(IsPedDir, "Pediatrie", "Neonatologie")
    
End Sub

Private Function GetEnvironment() As String

    Dim strEnv As String
    Dim strPath As String
    
    strPath = WbkAfspraken.Path
    strEnv = IIf(ModString.ContainsCaseInsensitive(strPath, "Test"), "Test", "")
    strEnv = IIf(ModString.ContainsCaseInsensitive(strPath, "Training"), "Training", strEnv)
    strEnv = IIf(ModString.ContainsCaseInsensitive(strPath, "Productie"), "Productie", strEnv)
    
    GetEnvironment = strEnv

End Function

Public Sub SetApplicationTitle()

    Dim strTitle As String
    Dim strBed As String
    Dim strVN As String
    Dim strAN As String
    
    strTitle = ModConst.CONST_APPLICATION_NAME
    strBed = ModBed.GetBed()
    strVN = ModPatient.PatientVoorNaam()
    strAN = ModPatient.PatientAchterNaam()
    
    If Not strBed = "0" Then
        strTitle = strTitle & " Patient: " & strAN & " " & strVN & ", Bed: " & strBed
    End If
    
    Application.Caption = strTitle

End Sub

Public Function GetLanguage() As EnumAppLanguage
    
    Dim enmLanguage As EnumAppLanguage
    
    Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    Case EnumAppLanguage.Dutch: enmLanguage = Dutch
    Case Else: enmLanguage = EnumAppLanguage.English
    End Select
    
    GetLanguage = enmLanguage

End Function

Private Sub TestGetLanguage()

    MsgBox GetLanguage()

End Sub

Private Function GetToDayFormula() As String
    
'    Dim strToDay As String

'    -- Probably not necessary with Formula instead of FormulaLocal
'    Select Case GetLanguage()
'    Case EnumAppLanguage.Dutch: strToDay = "= Vandaag()"
'    Case Else: strToDay = "= ToDay()"
'    End Select
    
    GetToDayFormula = "= ToDay()"

End Function

Private Function HasInPath(ByVal strDir As String) As Boolean

    Dim strPath As String

    strPath = WbkAfspraken.Path
    
    HasInPath = ModString.ContainsCaseInsensitive(strPath, strDir)

End Function

Public Function IsPedDir() As Boolean

    IsPedDir = ModSetting.IsPed()
    
End Function

Public Function IsNeoDir() As Boolean

    IsNeoDir = ModSetting.IsNeo()

End Function


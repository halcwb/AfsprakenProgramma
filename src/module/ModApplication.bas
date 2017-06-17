Attribute VB_Name = "ModApplication"
Option Explicit

Private blnDontClose As Boolean
Private blnCloseHaseRun As Boolean

Private Const constVersie As String = "Var_Glob_Versie"
Private Const constDate As String = "Var_AfspraakDatum"

Private Const constBarDel As String = " | "

Public Enum EnumAppLanguage
    Dutch = 1043
    English = 0
End Enum

Public Sub SetDontClose(ByVal blnClose As Boolean)
    
    blnDontClose = blnClose

End Sub

Public Sub SetToDevelopmentMode()

    Dim blnDevelop As Boolean
    
    blnDevelop = Not ModSetting.GetDevelopmentMode()
    
    If blnDevelop Then
        ModProgress.StartProgress "Zet in Ontwikkel Modus"
        Application.ScreenUpdating = False
        
        ModSheet.UnprotectUserInterfaceSheets True
        ModSheet.UnhideNonUserInterfaceSheets True
                
        ModSetting.SetDevelopmentMode True
        SetWindowToCloseApp WbkAfspraken.Windows(1)
        
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
    
    strAction = "ModApplication.CloseAfspraken"
    strParams = Array()
    
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
        Application.Quit
    End If

End Sub

Private Sub TestCloseAfspraken()
    blnDontClose = True
    CloseAfspraken
    MsgBox Application.DisplayAlerts
End Sub

Private Sub SetWindow(ByRef objWindow As Window, ByRef blnDisplay As Boolean)

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

Public Sub SetWindowToCloseApp(ByRef objWindow As Window)
    
    SetWindow objWindow, True

End Sub

Public Sub SetWindowToOpenApp(ByRef objWindow As Window)
    
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
    strParams = Array()
    
    ModLog.LogActionStart strAction, strParams
    
    SetCaptionAndHideBars              ' Setup Excel Application
    shtGlobGuiFront.Select
    
    For Each objWind In WbkAfspraken.Windows
        SetWindowToOpenApp objWind
    Next
    
    DoEvents                           ' Make sure sheet is shown before proceding
        
    ModProgress.StartProgress "Start Afspraken Programma"
            
    Application.ScreenUpdating = False ' Prevent cycling through all windows when sheets are processed
    
    ' Setup sheets
    ModSheet.ProtectUserInterfaceSheets True
    ModSheet.HideAndUnProtectNonUserInterfaceSheets True
    
    Application.ScreenUpdating = True
    
    ' Clean everything
    ModRange.SetRangeValue constVersie, vbNullString
    ModPatient.PatientClearAll False, True ' Default start with no patient
    ModSetting.SetDevelopmentMode False    ' Default development mode is false
    
    ModProgress.FinishProgress
    
    ModSheet.SelectPedOrNeoStartSheet  ' Select the first GUI sheet
    
    strBed = ModMetaVision.MetaVision_GetCurrentBedName()
    If strBed <> vbNullString Then
        ModBed.SetBed strBed
        ModBed.OpenBedAsk False, True
    End If
    
    ModLog.LogActionEnd strAction
        
    Exit Sub
    
InitializeError:
    
    ModProgress.FinishProgress

    strError = "Kan de applicatie niet opstarten"
    ModMessage.ShowMsgBoxError strError
    
    strError = strError & vbNewLine & strAction
    ModLog.LogError strError
    
End Sub

Public Sub UpdateStatusBar(ByVal strItem, ByVal strMessage)

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
    
    blnIsDevelop = ModSetting.GetDevelopmentMode()
    
    SetApplicationTitle
    
    With Application
        .DisplayFormulaBar = blnIsDevelop
        .DisplayStatusBar = blnIsDevelop
        .DisplayFullScreen = False
        .DisplayScrollBars = True
        .WindowState = xlMaximized
    End With
    
    Application.StatusBar = ModConst.CONST_APPLICATION_NAME
    UpdateStatusBar "Afdeling", IIf(IsPedDir, "Pediatrie", "Neonatologie")
    
End Sub

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

    IsPedDir = HasInPath(ModSetting.GetPedDir())
    
End Function

Public Function IsNeoDir() As Boolean

    IsNeoDir = HasInPath(ModSetting.GetNeoDir())

End Function


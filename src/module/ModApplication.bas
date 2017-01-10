Attribute VB_Name = "ModApplication"
Option Explicit

Private blnDontClose As Boolean
Private blnCloseHaseRun As Boolean

Private Const constVersie As String = "Var_Glob_Versie"
Private Const constDate As String = "Var_AfspraakDatum"

Private Const constDateFormatDutch As String = "dd-mmm-jj"
Private Const constDateFormatEnglish As String = "dd-mmm-yy"
Private Const constReplDate As String = "{DATEFORMAT}"
Private Const constReplSpace As String = "{SPACE}"
Private Const constReplEmpty As String = "{EMPTYSTRING}"
Private Const constDateFormula As String = "=IF(_Pat_OpnDatum>EmptyDate(),TEXT(_Pat_OpnDatum,{DATEFORMAT})&{SPACE}&B20,{EMPTYSTRING})"
Private Const constOpnameDate As String = "Var_Pat_OpnameDat"

Public Enum EnumAppLanguage
    Dutch = 1043
    English = 0
End Enum

Public Sub SetDontClose(ByVal blnClose As Boolean)
    
    blnDontClose = blnClose

End Sub

Public Sub SetToDevelopmentMode()

    ModProgress.StartProgress "Zet in Ontwikkel Modus"
    
    ModSheet.UnprotectUserInterfaceSheets True
    ModSheet.UnhideNonUserInterfaceSheets True
            
    ModSetting.SetDevelopmentMode True
    
    ModProgress.FinishProgress
    
    Application.DisplayFormulaBar = True

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
    intC = Application.Windows.Count
    For Each objWindow In Application.Windows
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
        
    ModProgress.FinishProgress
    ModLog.LogActionEnd strAction
    
    blnCloseHaseRun = True
            
    If Not blnDontClose Then
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

    blnDisplay = blnDisplay Or ModSetting.IsDevelopmentMode()

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
    Dim strParams() As Variant
    
    On Error GoTo InitializeError
    
    strAction = "ModApplication.InitializeAfspraken"
    strParams = Array()
    
    ModLog.LogActionStart strAction, strParams
        
    ModProgress.StartProgress "Start Afspraken Programma"
        
    SetCaptionAndHideBars              ' Setup Excel Application
    
    ModSheet.SelectPedOrNeoStartSheet  ' Select the first GUI sheet
    DoEvents                           ' Make sure sheet is shown before proceding
    ModApplication.SetWindowToOpenApp WbkAfspraken.Windows(1)
    DoEvents
    
    Application.ScreenUpdating = False ' Prevent cycling through all windows when sheets are processed
    
    ' Setup sheets
    ModSheet.ProtectUserInterfaceSheets True
    ModSheet.HideAndUnProtectNonUserInterfaceSheets True
    
    Application.ScreenUpdating = True
    
    ' Localization of formula's
    SetOpnameDateFormula
    
    ' Clean everything
    ModRange.SetRangeValue constVersie, vbNullString
    ModPatient.PatientClearAll False, True ' Default start with no patient
    ModSetting.SetDevelopmentMode False    ' Default development mode is false
    
    ModProgress.FinishProgress
    
    ModLog.LogActionEnd strAction
        
    Exit Sub
    
InitializeError:
    
    ModProgress.FinishProgress

    strError = ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & " Kan de applicatie niet opstarten"
    ModMessage.ShowMsgBoxError strError
    
    strError = strError & vbNewLine & strAction
    ModLog.LogError strError
    
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
        .DisplayStatusBar = blnIsDevelop
        .DisplayFullScreen = False
        .DisplayScrollBars = True
        .WindowState = xlMaximized
    End With
    
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

Private Sub SetOpnameDateFormula()

    Dim strFormula As String
    Dim strError As String
    
    Select Case GetLanguage()
    
        Case Dutch
            strFormula = Strings.Replace(constDateFormula, constReplDate, constDateFormatDutch)
            
        Case English
            strFormula = Strings.Replace(constDateFormula, constReplDate, constDateFormatEnglish)
        Case Else
            GoTo SetOpnameDateFormulaError
    
    End Select
    
    strFormula = Strings.Replace(strFormula, constReplSpace, Chr(34) & " " & Chr(34))
    strFormula = Strings.Replace(strFormula, constReplEmpty, Chr(34) & Chr(34))
    
    ModRange.SetRangeFormula constOpnameDate, strFormula
    
    Exit Sub
    
SetOpnameDateFormulaError:

    strError = "Language setting is not supported. Only English and Dutch"
    ModMessage.ShowMsgBoxError "Language setting is not supported. Only English and Dutch"
    ModLog.LogError strError

End Sub


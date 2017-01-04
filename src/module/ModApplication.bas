Attribute VB_Name = "ModApplication"
Option Explicit

Private blnDontClose As Boolean
Private blnCloseHaseRun As Boolean

Public Enum EnumAppLanguage
    Dutch = 1043
    English = 0
End Enum

Public Sub SetDontClose(blnClose As Boolean)
    
    blnDontClose = blnClose

End Sub

Public Sub SetToDevelopmentMode()

    Dim objSheet As Worksheet
    
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

Private Sub SetWindow(objWindow As Window, blnDisplay As Boolean)

    blnDisplay = blnDisplay Or ModSetting.IsDevelopmentMode()

    With objWindow
        .DisplayWorkbookTabs = blnDisplay
        .DisplayGridlines = blnDisplay
        .DisplayHeadings = blnDisplay
        .DisplayOutline = blnDisplay
        .DisplayZeros = blnDisplay
        .WindowState = xlMaximized
    End With

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
    Dim strParams() As Variant
    Dim objWindow As Window
    
    On Error GoTo InitializeError
    
    strAction = "ModApplication.InitializeAfspraken"
    strParams = Array()
    
    ModLog.LogActionStart strAction, strParams
    
    Application.WindowState = xlMaximized
    
    ModProgress.StartProgress "Start Afspraken Programma ... "
    
    ModSheet.SelectPedOrNeoStartSheet ' Select the first GUI sheet
    DoEvents                          ' Make sure sheet is shown before proceding
    
    Application.ScreenUpdating = False ' Prevent cycling through all windows when sheets are processed
    
    ' Setup sheets
    ModSheet.ProtectUserInterfaceSheets True
    ModSheet.HideAndUnProtectNonUserInterfaceSheets True
    ModApplication.SetWindowToOpenApp WbkAfspraken.Windows(1)

    ' Setup Excel Application
    SetCaptionAndHideBars
    
    ' Clean everything
    ModRange.SetRangeValue ModConst.CONST_RANGE_VERSIE, vbNullString
    SetDateToDayFormula
    ModPatient.PatientClearAll False, True ' Default start with no patient
    ModSetting.SetDevelopmentMode False ' Default development mode is false
    
    ModProgress.FinishProgress
    
    Application.ScreenUpdating = True
    
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

    ModRange.SetRangeFormula ModConst.CONST_RANGE_DATE, GetToDayFormula()

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
    End With
    
End Sub

Public Sub SetApplicationTitle()

    Dim strTitle As String
    Dim strBed As String
    Dim strVN As String
    Dim strAN As String
    
    strTitle = ModConst.CONST_APPLICATION_NAME
    strBed = ModRange.GetRangeValue(ModConst.CONST_RANGE_BED, "")
    strVN = ModRange.GetRangeValue(ModConst.CONST_RANGE_VN, "")
    strAN = ModRange.GetRangeValue(ModConst.CONST_RANGE_AN, "")
    
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

Private Function HasInPath(strDir As String) As Boolean

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


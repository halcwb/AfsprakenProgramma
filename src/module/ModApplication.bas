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
    
    shtGlobGuiFront.Select
    Application.ScreenUpdating = False ' Prevent cycling through all windows when sheets are processed
    
    strAction = "ModApplication.InitializeAfspraken"
    
    ModLog.LogActionStart strAction, strParams
          
    SetCaptionAndHideBars              ' Setup Excel Application
    
    For Each objWind In WbkAfspraken.Windows
        SetWindowToOpenApp objWind
    Next
        
    ModProgress.StartProgress "Start Afspraken Programma"
                
    ' Setup sheets
    ModSheet.ProtectUserInterfaceSheets True
    ModSheet.HideAndUnProtectNonUserInterfaceSheets True
    
    ' Load config tables
    LoadConfigTables
        
    ' Clean everything
    ModRange.SetRangeValue "Var_Glob_Versie", vbNullString
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
    
    Application.ScreenUpdating = True
            
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
    Dim strStatus As String
    Dim varItem() As String
    Dim intN As Integer
    Dim intC As Integer
    Dim blnItemSet As Boolean
    
    On Error Resume Next
    
    strStatus = CStr(Application.StatusBar)
    varStatus = Split(strStatus, constBarDel)
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
    
    On Error GoTo SetCaptionAndHideBarsError
    
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
    
    Exit Sub
    
SetCaptionAndHideBarsError:

    ModLog.LogError "SetCaptionAndHideBarsError"
    
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
    
    strTitle = IIf(strAN = vbNullString, strTitle, strTitle & " Patient: " & strAN)
    strTitle = IIf(strVN = vbNullString, strTitle, strTitle & ", " & strVN)
    strTitle = IIf(strBed = vbNullString, strTitle, strTitle & " Bed: " & strBed)

    
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

Private Function IsPedDir() As Boolean

    IsPedDir = MetaVision_IsPediatrie()
    
End Function

Private Function IsNeoDir() As Boolean

    IsNeoDir = Not IsPedDir()

End Function

Private Sub LoadConfigTables()

    Dim strFile As String
    Dim strTable As String
    Dim strSrc As String
    
    strTable = "Tbl_Admin_NeoMedCont"
    strSrc = "A2:S24"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    LoadConfigTable strFile, strTable, strSrc
    
    strTable = "Var_Neo_MedCont_VerdunningTekst"
    strSrc = "A1"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    LoadConfigTable strFile, strTable, strSrc
    
    strTable = "Tbl_Admin_PedMedCont"
    strSrc = "A4:Q34"
    strFile = WbkAfspraken.Path & "\db\PedMedCont.xlsx"
    
    LoadConfigTable strFile, strTable, strSrc

    strTable = "Tbl_Admin_ParEnt"
    strSrc = "A5:N45"
    strFile = WbkAfspraken.Path & "\db\GlobParEnt.xlsx"
    
    LoadConfigTable strFile, strTable, strSrc

End Sub

Public Sub Application_SaveNeoMedContConfig()

    Dim strFile As String
    Dim strTable As String
    Dim strDst As String
    
    Application.ScreenUpdating = False
    ModProgress.StartProgress "Neo Continue Medicatie Configuratie Opslaan"
    
    strTable = "Tbl_Admin_NeoMedCont"
    strDst = "A2:S24"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    SaveConfigTable strFile, strTable, strDst
    
    strTable = "Var_Neo_MedCont_VerdunningTekst"
    strDst = "A1"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    SaveConfigTable strFile, strTable, strDst
    
    ModProgress.FinishProgress
    Application.ScreenUpdating = True

End Sub

Public Sub Application_SaveParEntConfig()

    Dim strFile As String
    Dim strTable As String
    Dim strDst As String
    
    Application.ScreenUpdating = False
    ModProgress.StartProgress "Parenteralia Configuratie Opslaan"
    
    strTable = "Tbl_Admin_ParEnt"
    strDst = "A5:N45"
    strFile = WbkAfspraken.Path & "\db\GlobParEnt.xlsx"
    
    SaveConfigTable strFile, strTable, strDst
        
    ModProgress.FinishProgress
    Application.ScreenUpdating = True

End Sub

Private Sub LoadConfigTable(ByVal strFile As String, ByVal strTable As String, ByVal strConfig As String)
    
    Dim objConfigWbk As Workbook
    Dim objSrc As Range
    Dim objDst As Range
    
    Dim intR As Integer
    Dim intC As Integer
    
    Dim intN As Integer
    Dim intJ As Integer
    Dim intT As Integer
    
    Dim strMsg As String
    
    On Error GoTo LoadConfigTableError
       
    Application.DisplayAlerts = False
        
    Set objConfigWbk = Workbooks.Open(strFile, True, True)
    Set objSrc = objConfigWbk.Sheets(strTable).Range(strConfig)
    Set objDst = ModRange.GetRange(strTable)
        
    intR = objSrc.Rows.Count
    intC = objSrc.Columns.Count
    
    If Not intR = objDst.Rows.Count Or Not intC = objDst.Columns.Count Then Err.Raise ModConst.CONST_APP_ERROR, , ModConst.CONST_DEFAULTERROR_MSG
    
    intT = intR
    For intN = 1 To intR
        strMsg = strTable & " " & objDst.Cells(intN, 1).Value2
        For intJ = 1 To intC
            objDst.Cells(intN, intJ).Formula = objSrc.Cells(intN, intJ).Formula
        Next
        ModProgress.SetJobPercentage strMsg, intT, intN
    Next
    
    objConfigWbk.Close False
    
    Set objConfigWbk = Nothing
    
    Application.DisplayAlerts = True
    
    Exit Sub
    
LoadConfigTableError:

    ModLog.LogError "Kan config table " & strTable & " niet laden"
    
    On Error Resume Next
    
    objConfigWbk.Close False
    
    Set objDst = Nothing
    Set objSrc = Nothing
    Set objConfigWbk = Nothing
    
    Application.DisplayAlerts = True
End Sub

Private Sub SaveConfigTable(ByVal strFile As String, ByVal strTable As String, ByVal strConfig As String)
    
    Dim objConfigWbk As Workbook
    Dim objSrc As Range
    Dim objDst As Range
    
    Dim intR As Integer
    Dim intC As Integer
    
    Dim intN As Integer
    Dim intJ As Integer
    Dim intT As Integer
    
    Dim strMsg As String
    
    On Error GoTo SaveConfigTableError
    
    Application.DisplayAlerts = False
            
    Set objConfigWbk = Workbooks.Open(strFile, True, False)
    Set objDst = objConfigWbk.Sheets(strTable).Range(strConfig)
    Set objSrc = ModRange.GetRange(strTable)
        
    intR = objSrc.Rows.Count
    intC = objSrc.Columns.Count
    
    If Not intR = objDst.Rows.Count Or Not intC = objDst.Columns.Count Then Err.Raise ModConst.CONST_APP_ERROR, , ModConst.CONST_DEFAULTERROR_MSG
    
    intT = intR
    For intN = 1 To intR
        strMsg = objSrc.Cells(intN, 1).Value2
        For intJ = 1 To intC
            objDst.Cells(intN, intJ).Formula = objSrc.Cells(intN, intJ).Formula
        Next
        ModProgress.SetJobPercentage strMsg, intT, intN
    Next
    
    objConfigWbk.Save
    objConfigWbk.Close True
    
    Set objConfigWbk = Nothing
    
    Application.DisplayAlerts = True
    Exit Sub
    
SaveConfigTableError:

    ModLog.LogError "Kan config table " & strTable & " niet opslaan"
    
    On Error Resume Next
    
    objConfigWbk.Close False
    
    Set objDst = Nothing
    Set objSrc = Nothing
    Set objConfigWbk = Nothing
    
    Application.DisplayAlerts = True
End Sub



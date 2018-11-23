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

Public Sub ToggleDevelopmentMode()

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
        
        Application_Initialize
    End If

End Sub

Public Sub Application_CloseApplication()

    Dim strAction As String
    Dim strParams() As Variant
    
    Dim intN As Integer
    Dim intC As Integer
    
    Dim objWindow As Window
    
    On Error GoTo ErrorHandler
    
    If blnCloseHaseRun Then ' Second Application_CloseApplication triggert by WbkAfspraken.Workbook_BeforeClose
        Exit Sub
    End If
    
    If Application.Workbooks.Count > 1 Then
        ModMessage.ShowMsgBoxExclam "Er zijn nog andere Excel bestanden geopend, sla deze eerst op anders worden deze niet opgeslagen!"
        WbkAfspraken.blnCancelAfsprakenClose = True
        Exit Sub
    End If
    
    shtGlobGuiFront.Select
    strAction = "Application_CloseApplication"
    
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
        Database_LogAction "Close application"
        Application.Quit
    End If
        
    Exit Sub
    
ErrorHandler:

    ModLog.LogError Err, "Application_CloseApplication"

End Sub

Private Sub TestCloseAfspraken()
    blnDontClose = True
    Application_CloseApplication
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

Public Sub Application_Initialize()

    Dim strError As String
    Dim strAction As String
    Dim strBed As String
    Dim strPw As String
    Dim strParams() As Variant
    Dim objWind As Window
    
    On Error GoTo ErrorHandler
    
    shtGlobGuiFront.Select
    Application.ScreenUpdating = False ' Prevent cycling through all windows when sheets are processed
    
    strAction = "ModApplication.Application_Initialize"
    
    ModLog.LogActionStart strAction, strParams
          
    ClearLogin
    MetaVision_SetUser
    If ModRange.GetRangeValue("_User_Login", vbNullString) = vbNullString Then
        strPw = ModMessage.ShowPasswordBox("Geef wachtwoord op om in te loggen, of start de applicatie op vanuit MetaVision")
        If Not strPw = ModConst.CONST_PASSWORD Then
            ModMessage.ShowMsgBoxExclam "Kan niet inloggen met dit wachtwoord"
            Application.Quit
        Else
            ModRange.SetRangeValue "_User_Login", "Develop"
            ModRange.SetRangeValue "_User_FirstName", "Ontwikkelaar"
            ModRange.SetRangeValue "_User_LastName", "AfsprakenProgramma"
            ModRange.SetRangeValue "_User_Type", "Beheerders"
        End If
    End If
    
    SetCaptionAndHideBars              ' Setup Excel Application
    
    For Each objWind In WbkAfspraken.Windows
        SetWindowToOpenApp objWind
    Next
        
    ModProgress.StartProgress "Start Afspraken Programma"
                
    ' Setup sheets
    ModSheet.ProtectUserInterfaceSheets True
    ModSheet.HideAndUnProtectNonUserInterfaceSheets True
    
    ' Load config tables
    If Setting_UseDatabase Then
        If Not LoadConfigTablesFromDatabase() Then Err.Raise ModConst.CONST_APP_ERROR, "Application_Initialize", "Kan config tabellen niet laden"
    Else
        If Not LoadConfigTables() Then Err.Raise ModConst.CONST_APP_ERROR, "Application_Initialize", "Kan config tabellen niet laden"
    End If
    
    ' Clean everything
    ModRange.SetRangeValue "Var_Glob_Versie", vbNullString
    ModSetting.SetDevelopmentMode False    ' Default development mode is false
            
    strBed = ModMetaVision.MetaVision_GetCurrentBedName()
    If strBed <> vbNullString Then
        If Setting_UseDatabase Then
            ModBed.SetBed strBed
            Patient_OpenPatient
        Else
            ModBed.SetBed strBed
            ModBed.OpenBedAsk False, True
        End If
    Else
        Patient_ClearAll False, True ' Default start with no patient
    End If
    
    ModSheet.SelectPedOrNeoStartSheet False  ' Select the first GUI sheet
    
    ModProgress.FinishProgress
    
    ModLog.LogActionEnd strAction
    Database_LogAction "Initialize Application"
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    
    ModProgress.FinishProgress
    Application.Visible = True

    strError = "De applicatie is mogelijk niet goed opgestart." & vbNewLine
    strError = strError & "Sluit evt. andere Excel bestanden en probeer het opnieuw." & vbNewLine
    ModMessage.ShowMsgBoxError strError
    
    strError = strError & vbNewLine & strAction
    ModLog.LogError Err, strError
    
    Application_CloseApplication
    
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
    UpdateStatusBar "Login", ModRange.GetRangeValue("_User_Login", vbNullString)
    
    Exit Sub
    
SetCaptionAndHideBarsError:

    ModLog.LogError Err, "SetCaptionAndHideBarsError"
    
End Sub

Private Function GetEnvironment() As String

    Dim strEnv As String
    Dim strPath As String
    
    strPath = WbkAfspraken.Path
    strEnv = IIf(ModString.ContainsCaseInsensitive(strPath, "Test"), "Test", vbNullString)
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
    strVN = ModPatient.Patient_GetFirstName()
    strAN = ModPatient.Patient_GetLastName()
    
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

Private Function LoadConfigTables() As Boolean

    Dim strFile As String
    Dim strTable As String
    Dim strSrc As String
    Dim blnLoaded As Boolean
    
    blnLoaded = True
    
    strTable = "Tbl_Admin_NeoMedCont"
    strSrc = "A2:S24"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    blnLoaded = blnLoaded And LoadConfigTable(strFile, strTable, strSrc)
    
    strTable = "Var_Neo_MedCont_VerdunningTekst"
    strSrc = "A1"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    blnLoaded = blnLoaded And LoadConfigTable(strFile, strTable, strSrc)
    
    strTable = "Tbl_Admin_PedMedCont"
    strSrc = "A4:R34"
    strFile = WbkAfspraken.Path & "\db\PedMedCont.xlsx"
    
    blnLoaded = blnLoaded And LoadConfigTable(strFile, strTable, strSrc)

    strTable = "Tbl_Admin_ParEnt"
    strSrc = "A6:N47"
    strFile = WbkAfspraken.Path & "\db\GlobParEnt.xlsx"
    
    blnLoaded = blnLoaded And LoadConfigTable(strFile, strTable, strSrc)
    
    LoadConfigTables = blnLoaded

End Function

Private Sub Test_LoadConfigTables()

    MsgBox LoadConfigTables()

End Sub

Private Function LoadConfigTablesFromDatabase() As Boolean

    On Error GoTo ErrorHandler

    ModDatabase.Database_LoadNeoConfigMedCont
    ModDatabase.Database_LoadPedConfigMedCont
    ModDatabase.Database_LoadConfigParEnt
    
    LoadConfigTablesFromDatabase = True
    
    Exit Function
    
ErrorHandler:

    ModLog.LogError Err, "LoadConfigTablesFromDatabase"

End Function

Public Sub Application_SaveNeoMedContConfig()

    Dim strFile As String
    Dim strTable As String
    Dim strDst As String
    
    Application.ScreenUpdating = False
    ModProgress.StartProgress "Neo Continue Medicatie Configuratie Opslaan"
    
    strTable = "Tbl_Admin_NeoMedCont"
    strDst = "A2:R24"
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
    strDst = "A5:N47"
    strFile = WbkAfspraken.Path & "\db\GlobParEnt.xlsx"
    
    SaveConfigTable strFile, strTable, strDst
        
    ModProgress.FinishProgress
    Application.ScreenUpdating = True

End Sub


Private Function LoadConfigTable(ByVal strFile As String, ByVal strTable As String, ByVal strConfig As String) As Boolean
    
    Dim objConfigWbk As Workbook
    Dim objSrc As Range
    Dim objDst As Range
    Dim lngErr As Long
    
    Dim strMsg As String
    
    On Error GoTo LoadConfigTableError
       
    Application.DisplayAlerts = False
        
    Set objConfigWbk = Workbooks.Open(strFile, True, True)
    Set objSrc = objConfigWbk.Sheets(strTable).Range(strConfig)
    Set objDst = ModRange.GetRange(strTable)
        
    Sheet_CopyRangeFormulaToDst objSrc, objDst
    
    objConfigWbk.Close False
    
    Set objConfigWbk = Nothing
    
    Application.DisplayAlerts = True
    LoadConfigTable = True
    
    Exit Function
    
LoadConfigTableError:

    lngErr = Err.Number
    ModLog.LogError Err, "Kan config table " & strTable & " niet laden"
    
    On Error Resume Next
    
    If Not objConfigWbk Is Nothing Then objConfigWbk.Close False
    
    Set objDst = Nothing
    Set objSrc = Nothing
    Set objConfigWbk = Nothing
    
    Application.DisplayAlerts = True
    
    LoadConfigTable = False
    
End Function


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

    ModLog.LogError Err, "Kan config table " & strTable & " niet opslaan"
    
    On Error Resume Next
    
    objConfigWbk.Close False
    
    Set objDst = Nothing
    Set objSrc = Nothing
    Set objConfigWbk = Nothing
    
    Application.DisplayAlerts = True
End Sub

Private Sub PrintScreen()

    Application.SendKeys "(%{1068})"
    DoEvents

End Sub

Public Sub Application_SendPrintScreen()

'     Dim objClip
    PrintScreen
    
End Sub

Private Sub ClearLogin()

    ModRange.SetRangeValue "_User_Login", vbNullString
    ModRange.SetRangeValue "_User_FirstName", vbNullString
    ModRange.SetRangeValue "_User_LastName", vbNullString
    ModRange.SetRangeValue "_User_Type", vbNullString
    
End Sub

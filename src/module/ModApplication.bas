Attribute VB_Name = "ModApplication"
Option Explicit

Private blnDontClose As Boolean
Private blnCloseHaseRun As Boolean

Public Const CONST_PRESCRIPTIONS_DATE As String = "Var_AfspraakDatum"

Private Const constBarDel As String = " | "

Public Enum EnumAppLanguage
    Dutch = 1043
    English = 0
End Enum

Public Sub App_SetDontClose(ByVal blnClose As Boolean)
    
    blnDontClose = blnClose

End Sub

Public Function App_GetApplicationVersion() As String

    App_GetApplicationVersion = ModRange.GetRangeValue(CONST_APP_VERSION, vbNullString)

End Function

Public Sub App_ToggleDevelopmentMode()

    Dim blnDevelop As Boolean
    Dim objWindow As Window
    
    blnDevelop = Not ModSetting.IsDevelopmentMode()
    
    shtGlobGuiFront.Select
    
    If blnDevelop Then
        ModProgress.StartProgress "Zet in Ontwikkel Modus"
        ImprovePerf True
        
        ModSheet.UnprotectUserInterfaceSheets True
        ModSheet.UnhideNonUserInterfaceSheets True
                
        ModSetting.SetDevelopmentMode True
        For Each objWindow In WbkAfspraken.Windows
            App_SetWindowToClose objWindow
        Next
        
        ImprovePerf False
        ModProgress.FinishProgress
        
        Application.DisplayFormulaBar = True
    Else
        ModMessage.ShowMsgBoxInfo "Weer terug zetten in Gebruikers Modus"
        ModSetting.SetDevelopmentMode False
        
        App_Initialize
    End If

End Sub

Public Sub App_CloseApplication()

    Dim strAction As String
    Dim strParams() As Variant
    
    Dim intN As Integer
    Dim intC As Integer
    
    Dim objWindow As Window
    
    On Error GoTo ErrorHandler
    
    If blnCloseHaseRun Then ' Second App_CloseApplication triggert by WbkAfspraken.Workbook_BeforeClose
        Exit Sub
    End If
    
    If Application.Workbooks.Count > 1 Then
        ModMessage.ShowMsgBoxExclam "Er zijn nog andere Excel bestanden geopend, sla deze eerst op anders worden deze niet opgeslagen!"
        WbkAfspraken.blnCancelAfsprakenClose = True
        Exit Sub
    End If
    
    shtGlobGuiFront.Select
    strAction = "App_CloseApplication"
    
    ModLog.LogActionStart strAction, strParams
    
    ModProgress.StartProgress "Afspraken Programma Afsluiten"
    ImprovePerf True
    
    intN = 1
    intC = WbkAfspraken.Windows.Count
    For Each objWindow In WbkAfspraken.Windows
        App_SetWindowToClose objWindow
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
    ImprovePerf False
    
    If Not blnDontClose Then
        Application.StatusBar = vbNullString
        Application.DisplayAlerts = False
        Application.DisplayFormulaBar = True
        
        WbkAfspraken.blnCancelAfsprakenClose = False
        Database_LogAction "Close application", ModUser.User_GetCurrent().Login, ModPatient.Patient_GetHospitalNumber()
        Application.Quit
    End If
        
    Exit Sub
    
ErrorHandler:

    ImprovePerf False
    ModLog.LogError Err, "App_CloseApplication"

End Sub

Private Sub Test_CloseAfspraken()
    blnDontClose = True
    App_CloseApplication
    MsgBox Application.DisplayAlerts
End Sub

Private Sub Util_SetWindow(objWindow As Window, ByVal blnDisplay As Boolean)

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

Public Sub App_SetWindowToClose(objWindow As Window)
    
    Util_SetWindow objWindow, True

End Sub

Public Sub App_SetWindowToOpen(objWindow As Window)
    
    Util_SetWindow objWindow, False

End Sub

Public Sub App_Initialize()

    Dim strError As String
    Dim strAction As String
    Dim strBed As String
    Dim strPw As String
    Dim strParams() As Variant
    Dim objWind As Window
    Dim objUser As ClassUser
    
    On Error GoTo ErrorHandler
    
    shtGlobGuiFront.Select
    ImprovePerf True
    
    strAction = "ModApplication.App_Initialize"
    
    ModLog.LogActionStart strAction, strParams
          
    Util_ClearLogin
    MetaVision_SetUser
    Set objUser = User_GetCurrent()
    If Not objUser.IsValid Then
        strPw = ModMessage.ShowPasswordBox("Geef wachtwoord op om in te loggen, of start de applicatie op vanuit MetaVision")
        If Not strPw = ModConst.CONST_PASSWORD Then
            ModMessage.ShowMsgBoxExclam "Kan niet inloggen met dit wachtwoord"
            Application.Quit
        Else
            With objUser
                .Login = "system"
                .LastName = "User"
                .FirstName = "System"
                .Role = "Beheerders"
            End With
            ModUser.User_SetUser objUser
        End If
    End If
    
    Util_SetCaptionAndHideBars              ' Setup Excel Application
    
    For Each objWind In WbkAfspraken.Windows
        App_SetWindowToOpen objWind
    Next
        
    ModProgress.StartProgress "Start Afspraken Programma"
                
    ' Setup sheets
    ModSheet.ProtectUserInterfaceSheets True
    ModSheet.HideAndUnProtectNonUserInterfaceSheets True
    
    ' Load config tables
    If Setting_UseDatabase Then
        If Not App_LoadConfigTablesFromDatabase() Then Err.Raise ModConst.CONST_APP_ERROR, "App_Initialize", "Kan config tabellen niet laden"
    Else
        If Not Util_LoadConfigTables() Then Err.Raise ModConst.CONST_APP_ERROR, "App_Initialize", "Kan config tabellen niet laden"
    End If
    
    ' Clean everything
    ModRange.SetRangeValue "Var_Glob_Versie", vbNullString
    ModSetting.SetDevelopmentMode False    ' Default development mode is false
            
    strBed = ModMetaVision.MetaVision_GetCurrentBedName()
    If Not Setting_UseDatabase And strBed <> vbNullString Then
        Bed_SetBed strBed
        Bed_OpenBedAndAsk False, True
        MetaVision_SyncLab
    ElseIf Setting_UseDatabase Then
        ModPatient.Patient_SetHospitalNumber vbNullString
        Patient_OpenPatient
        MetaVision_SyncLab
    Else
        Patient_ClearAll False, True ' Default start with no patient
    End If
    MetaVision_SetUser ' Have to reload current user login
    
    ModSheet.SelectPedOrNeoStartSheet False  ' Select the first GUI sheet
    
    ModProgress.FinishProgress
    
    ModLog.LogActionEnd strAction
    Database_LogAction "Initialize Application", objUser.Login, ModPatient.Patient_GetHospitalNumber()
    
    ImprovePerf False
    
    If Util_CheckMedicationValidation() Then App_CloseApplication
    
    Exit Sub
    
ErrorHandler:
    
    ImprovePerf False
    ModProgress.FinishProgress
    Application.Visible = True

    strError = "De applicatie is mogelijk niet goed opgestart." & vbNewLine
    strError = strError & "Sluit evt. andere Excel bestanden en probeer het opnieuw." & vbNewLine
    ModMessage.ShowMsgBoxError strError
    
    strError = strError & vbNewLine & strAction
    ModLog.LogError Err, strError
    
    If ModSetting.IsProductionDir() Then App_CloseApplication
    
End Sub

Public Sub App_UpdateStatusBar(ByVal strItem As String, ByVal strMessage As String)

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

Private Sub Test_UpdateStatusBar()

    Application.StatusBar = " "
    App_UpdateStatusBar "Setting", "Test2"
    
End Sub

Public Sub App_SetPrescriptionsDate()

    ModRange.SetRangeFormula CONST_PRESCRIPTIONS_DATE, Util_GetToDayFormula()

End Sub

Private Sub Util_SetCaptionAndHideBars()

    Dim blnIsDevelop As Boolean
    
    On Error GoTo SetCaptionAndHideBarsError
    
    blnIsDevelop = ModSetting.IsDevelopmentMode()
    
    App_SetApplicationTitle
    
    With Application
        .DisplayFormulaBar = blnIsDevelop
        .DisplayStatusBar = True
        .DisplayFullScreen = False
        .DisplayScrollBars = True
        .WindowState = xlMaximized
    End With
    
    Application.StatusBar = ModConst.CONST_APPLICATION_NAME
    App_UpdateStatusBar "Versie", ModRange.GetRangeValue("Var_Glob_AppVersie", vbNullString)
    App_UpdateStatusBar "Omgeving", Util_GetEnvironment()
    App_UpdateStatusBar "Afdeling", IIf(Util_IsPedDir, "PICU", "NICU")
    App_UpdateStatusBar "Login", ModRange.GetRangeValue("_User_Login", vbNullString)
    App_UpdateStatusBar "Database", ModSetting.Setting_GetDatabase()
    
    Exit Sub
    
SetCaptionAndHideBarsError:

    ModLog.LogError Err, "SetCaptionAndHideBarsError"
    
End Sub

Private Function Util_GetEnvironment() As String

    Dim strEnv As String
    Dim strPath As String
    
    strPath = WbkAfspraken.Path
    strEnv = IIf(ModString.ContainsCaseInsensitive(strPath, "Test"), "Test", vbNullString)
    strEnv = IIf(ModString.ContainsCaseInsensitive(strPath, "Training"), "Training", strEnv)
    strEnv = IIf(ModString.ContainsCaseInsensitive(strPath, "Productie"), "Productie", strEnv)
    
    Util_GetEnvironment = strEnv

End Function

Public Sub App_SetApplicationTitle()

    Dim strTitle As String
    Dim strBed As String
    Dim strVN As String
    Dim strAN As String
    
    strTitle = ModConst.CONST_APPLICATION_NAME
    strBed = Bed_GetBedName()
    strVN = ModPatient.Patient_GetFirstName()
    strAN = ModPatient.Patient_GetLastName()
    
    strTitle = IIf(strAN = vbNullString, strTitle, strTitle & " Patient: " & strAN)
    strTitle = IIf(strVN = vbNullString, strTitle, strTitle & ", " & strVN)
    strTitle = IIf(strBed = vbNullString, strTitle, strTitle & " Bed: " & strBed)

    
    Application.Caption = strTitle

End Sub

Private Function Util_GetLanguage() As EnumAppLanguage
    
    Dim enmLanguage As EnumAppLanguage
    
    Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    Case EnumAppLanguage.Dutch: enmLanguage = Dutch
    Case Else: enmLanguage = EnumAppLanguage.English
    End Select
    
    Util_GetLanguage = enmLanguage

End Function

Private Sub Test_GetLanguage()

    MsgBox Util_GetLanguage()

End Sub

Private Function Util_GetToDayFormula() As String
    
    Util_GetToDayFormula = "= ToDay()"

End Function

Private Function Util_HasInPath(ByVal strDir As String) As Boolean

    Dim strPath As String

    strPath = WbkAfspraken.Path
    
    Util_HasInPath = ModString.ContainsCaseInsensitive(strPath, strDir)

End Function

Private Function Util_IsPedDir() As Boolean

    Util_IsPedDir = MetaVision_IsPICU()
    
End Function

Private Function Util_IsNeoDir() As Boolean

    Util_IsNeoDir = Not Util_IsPedDir()

End Function

Private Function Util_LoadConfigTables() As Boolean

    Dim strFile As String
    Dim strTable As String
    Dim strSrc As String
    Dim blnLoaded As Boolean
    
    blnLoaded = True
    
    strTable = "Tbl_Admin_NeoMedCont"
    strSrc = "A2:S24"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    blnLoaded = blnLoaded And Util_LoadConfigTable(strFile, strTable, strSrc)
    
    strTable = "Var_Neo_MedCont_VerdunningTekst"
    strSrc = "A1"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    blnLoaded = blnLoaded And Util_LoadConfigTable(strFile, strTable, strSrc)
    
    strTable = "Tbl_Admin_PedMedCont"
    strSrc = "A4:R34"
    strFile = WbkAfspraken.Path & "\db\PedMedCont.xlsx"
    
    blnLoaded = blnLoaded And Util_LoadConfigTable(strFile, strTable, strSrc)

    strTable = "Tbl_Admin_ParEnt"
    strSrc = "A6:N47"
    strFile = WbkAfspraken.Path & "\db\GlobParEnt.xlsx"
    
    blnLoaded = blnLoaded And Util_LoadConfigTable(strFile, strTable, strSrc)
    
    Util_LoadConfigTables = blnLoaded

End Function

Private Sub Test_Util_LoadConfigTables()

    MsgBox Util_LoadConfigTables()

End Sub

Public Function App_LoadConfigTablesFromDatabase() As Boolean

    On Error GoTo ErrorHandler

    ModDatabase.Database_LoadNeoConfigMedCont
    ModDatabase.Database_LoadPedConfigMedCont
    ModDatabase.Database_LoadConfigParEnt
    
    App_LoadConfigTablesFromDatabase = True
    
    Exit Function
    
ErrorHandler:

    ModLog.LogError Err, "Util_LoadConfigTablesFromDatabase"

End Function

' ToDo check if this is not death code
Public Sub App_SaveNeoMedContConfig()

    Dim strFile As String
    Dim strTable As String
    Dim strDst As String
    
    ImprovePerf True
    ModProgress.StartProgress "Neo Continue Medicatie Configuratie Opslaan"
    
    strTable = "Tbl_Admin_NeoMedCont"
    strDst = "A2:R24"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    Util_SaveConfigTable strFile, strTable, strDst
    
    strTable = "Var_Neo_MedCont_VerdunningTekst"
    strDst = "A1"
    strFile = WbkAfspraken.Path & "\db\NeoMedCont.xlsx"
    
    Util_SaveConfigTable strFile, strTable, strDst
    
    ModProgress.FinishProgress
    ImprovePerf False

End Sub

Public Sub App_SaveParEntConfig()

    Dim strFile As String
    Dim strTable As String
    Dim strDst As String
    
    ImprovePerf True
    ModProgress.StartProgress "Parenteralia Configuratie Opslaan"
    
    strTable = "Tbl_Admin_ParEnt"
    strDst = "A5:N47"
    strFile = WbkAfspraken.Path & "\db\GlobParEnt.xlsx"
    
    Util_SaveConfigTable strFile, strTable, strDst
        
    ModProgress.FinishProgress
    ImprovePerf False

End Sub

Private Function Util_LoadConfigTable(ByVal strFile As String, ByVal strTable As String, ByVal strConfig As String) As Boolean
    
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
    Util_LoadConfigTable = True
    
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
    
    Util_LoadConfigTable = False
    
End Function

Private Sub Util_SaveConfigTable(ByVal strFile As String, ByVal strTable As String, ByVal strConfig As String)
    
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

' Check whether this is used and if not should be used
Public Sub App_SendPrintScreen()

    Application.SendKeys "(%{1068})"
    DoEvents
    
End Sub

Private Sub Util_ClearLogin()

    ModRange.SetRangeValue "_User_Login", vbNullString
    ModRange.SetRangeValue "_User_FirstName", vbNullString
    ModRange.SetRangeValue "_User_LastName", vbNullString
    ModRange.SetRangeValue "_User_Type", vbNullString
    
End Sub

' Not sure if this function should be here
Private Function Util_CheckMedicationValidation() As Boolean

    Dim blnCheck As Boolean
    Dim strRegPath As String
    Dim strKey As String
    
    strRegPath = "HKCU\SOFTWARE\UMCU\MV"
    strKey = "MedicatieValidatie"
    If ModRegistry.RegistryKeyExists(strRegPath, strKey) Then
        blnCheck = ModRegistry.ReadRegistryKey(strRegPath, strKey) = 1
        If blnCheck Then ModMedDisc.SendApotheekMedDiscValidation
    Else
        blnCheck = False
    End If
    
    Util_CheckMedicationValidation = blnCheck

End Function

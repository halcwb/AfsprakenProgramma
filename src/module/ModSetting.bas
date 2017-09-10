Attribute VB_Name = "ModSetting"
Option Explicit

Public Const CONST_DATA_SHEET As String = "Data"
Public Const CONST_PATIENTS_SHEET As String = "Patienten"

Public Const CONST_PICU_BEDS As String = "PICU.xlsx"
Public Const CONST_NICU_BEDS As String = "NICU.xlsx"
Public Const CONST_BEDS_SHEET As String = "Beds"

Private Const constPatientsFile As String = "Patienten.xlsx"

Private Const constExt As String = ".xlsx"
Private Const constDevMode As String = "SettingDevMode"
Private Const constLogging As String = "SettingLogging"
Private Const constNeoDir As String = "SettingNeoDir"
Private Const constPedDir As String = "SettingPedDir"
Private Const constDevDir As String = "SettingDevDir"
Private Const constTestLogDir As String = "SettingTestLogDir"
Private Const constLogDir As String = "SettingLogDir"
Private Const constDataDir As String = "SettingDataDir"
Private Const constDbDir As String = "SettingDbDir"
Private Const constPreData As String = vbNullString
Private Const constPostData As String = "_Data"
Private Const constPreText As String = vbNullString
Private Const constPostText As String = "_Text"
Private Const constPedBeds As String = "Tbl_Ped_Beds"
Private Const constNeoBeds As String = "Tbl_Neo_Beds"

Private Const constUMCU_LogPath As String = "\\ds.umcutrecht.nl\Algemeen\Apps\Metavision-Appsdata\MetavisionLOG"

Private Function GetSetting(ByVal strSetting As String) As Variant

    Dim strMsg As String

    On Error GoTo GetSettingError:

    GetSetting = shtGlobSettings.Range(strSetting).Value2 'ModRange.GetRangeValue(strSetting, varDefault)
    
    Exit Function
    
GetSettingError:

    strMsg = "Kan setting: " & strSetting & " niet ophalen"
    ModMessage.ShowMsgBoxError strMsg

End Function

Private Sub SetSetting(ByVal strSetting As String, ByVal varValue As Variant)

    Dim strMsg As String

    On Error GoTo SetSettingError:

    shtGlobSettings.Range(strSetting).Value2 = varValue ' ModRange.SetRangeValue strSetting, varValue
    
    Exit Sub

SetSettingError:

    strMsg = "Kan setting: " & strSetting & " niet opslaan"
    ModMessage.ShowMsgBoxError strMsg
    
End Sub

Public Function GetDevelopmentMode() As Boolean

    GetDevelopmentMode = CBool(GetSetting(constDevMode))

End Function

Public Function IsDevelopmentDir() As Boolean

    Dim blnDevDir As Boolean
    Dim strActDir As String
    
    strActDir = WbkAfspraken.Path
    blnDevDir = ModString.ContainsCaseInsensitive(strActDir, GetDevelopmentDir)
    
    IsDevelopmentDir = blnDevDir

End Function

Private Sub Test_IsDevelopmentDir()

    MsgBox IsDevelopmentDir()

End Sub

Public Sub SetDevelopmentMode(ByVal blnMode As Boolean)

    SetSetting constDevMode, blnMode
    ModApplication.UpdateStatusBar "DevelopmentMode", IIf(blnMode, "Aan", "Uit")

End Sub

Public Function GetEnableLogging() As Boolean

    GetEnableLogging = CBool(GetSetting(constLogging))

End Function

Public Sub SetEnableLogging(ByVal blnMode As Boolean)

    SetSetting constLogging, blnMode
    ModApplication.UpdateStatusBar "Logging", IIf(blnMode, "Aan", "Uit")

End Sub

Public Sub ToggleLogging()

    Dim blnLog As Boolean
    
    blnLog = Not GetEnableLogging()
    SetEnableLogging blnLog
    
    If blnLog Then
        ModMessage.ShowMsgBoxInfo "Logging staat nu aan"
    Else
        ModMessage.ShowMsgBoxInfo "Logging staat nu uit"
    End If

End Sub

Private Function GetNeoDir() As String

    GetNeoDir = CStr(GetSetting(constNeoDir))

End Function

Public Sub SetNeoDir(ByVal strDir As String)

    SetSetting constNeoDir, strDir

End Sub

Private Function GetPedDir() As String

    GetPedDir = CStr(GetSetting(constPedDir))

End Function

Public Sub SetPedDir(ByVal strDir As String)

    SetSetting constPedDir, strDir

End Sub

Public Function GetDevelopmentDir() As String

    GetDevelopmentDir = CStr(GetSetting(constDevDir))

End Function

Public Sub SetDevelopmentDir(ByVal strDir As String)

    SetSetting constDevDir, strDir

End Sub

Public Function GetTestLogDir() As String

    GetTestLogDir = CStr(GetSetting(constTestLogDir))

End Function

Public Sub SetTestLogDir(ByVal strDir As String)

    SetSetting constTestLogDir, strDir

End Sub

Public Function GetTestLogPath() As String

    GetTestLogPath = WbkAfspraken.Path & "\" & GetTestLogDir()

End Function

Public Function GetLogDir() As String

    GetLogDir = CStr(GetSetting(constLogDir))

End Function

Public Sub SetLogDir(ByVal strDir As String)

    SetSetting constLogDir, strDir

End Sub

Public Function GetLogPath() As String

    Dim strPath As String
    Dim strDom As String
    Dim strUser As String
    Dim strCmpN As String
    
    strCmpN = Environ$("COMPUTERNAME")
    strUser = Environ$("USERNAME")
    strDom = Environ$("USERDOMAIN")
    
    If strDom = "DS" Then
        strPath = constUMCU_LogPath & "\" & strUser & "_" & strCmpN & "_" & "Afspraken2017.log"
    Else
        strPath = WbkAfspraken.Path & "\" & GetLogDir()
    End If
    
    GetLogPath = strPath

End Function

Private Sub Test_GetLogPath()

    MsgBox GetLogPath()

End Sub

Public Function GetDataDir() As String

    GetDataDir = CStr(GetSetting(constDataDir))

End Function

Public Sub SetDataDir(ByVal strDir As String)

    SetSetting constDataDir, strDir

End Sub

Public Function GetFormDbDir() As String

    GetFormDbDir = CStr(GetSetting(constDbDir))

End Function

Private Sub Test_GetFormDbDir()

    MsgBox GetFormDbDir()

End Sub

Public Sub SetFormDbDir(ByVal strDir As String)

    SetSetting constDbDir, strDir

End Sub

Public Function GetPatientDataPath() As String

    Dim strDir As String
    
    strDir = ModSetting.GetDataDir()
    GetPatientDataPath = GetAbsolutePath(strDir)

End Function

Private Function GetAbsolutePath(ByVal strPath As String) As String

    GetAbsolutePath = WbkAfspraken.Path & strPath

End Function

Public Function GetPatientTextWorkBookName(ByVal strBed As String) As String

    GetPatientTextWorkBookName = constPreText & strBed & constPostText & constExt

End Function

Public Function GetPatientDataWorkBookName(ByVal strBed As String) As String

    GetPatientDataWorkBookName = constPreData & strBed & constPostData + constExt

End Function

Public Function GetPatientDataFile(ByVal strBed As String) As String

    GetPatientDataFile = GetPatientDataPath() & GetPatientDataWorkBookName(strBed)

End Function

Public Function GetPatientTextFile(ByVal strBed As String) As String

    GetPatientTextFile = GetPatientDataPath() & GetPatientTextWorkBookName(strBed)

End Function

Private Function GetBeds(ByVal strRange As String) As Variant()

    Dim arrBeds() As Variant
    Dim objBeds As Range
    Dim intC As Integer
    Dim intN As Integer
    
    Set objBeds = shtGlobSettings.Range(strRange)
    intC = objBeds.Rows.Count
    For intN = 1 To intC
        ModArray.AddItemToVariantArray arrBeds, objBeds.Cells(intN, 1).Value2
    Next intN
    
    GetBeds = arrBeds

End Function

Public Function GetPedBeds() As Variant()

    GetPedBeds = GetBeds(constPedBeds)

End Function

Public Function GetNeoBeds() As Variant()

    GetNeoBeds = GetBeds(constNeoBeds)

End Function

Public Function GetPatientsFileName() As String
    
    Dim blnNeo As Boolean
    
    blnNeo = False
    If IsDevelopmentDir() Then blnNeo = ModMessage.ShowMsgBoxYesNo("Ja(Yes) voor Neo anders Pediatrie") = vbYes
    
    GetPatientsFileName = IIf(IsNeoDir() Or IsNeo() Or blnNeo, CONST_NICU_BEDS, CONST_PICU_BEDS)

End Function

Private Sub Test_GetPatientFileName()

    MsgBox GetPatientsFileName()

End Sub

Public Function GetPatientsFilePath(ByVal strFileName As String) As String

    GetPatientsFilePath = GetPatientDataPath() & strFileName

End Function

Private Sub test()

    MsgBox GetPatientsFilePath("Test")
    
End Sub

Public Function IsNeo() As Boolean

    Dim strPath As String
    Dim strDir As String
    Dim blnIsNeo As Boolean

    strPath = Application.ActiveWorkbook.Path
    strDir = ModSetting.GetNeoDir()
    
    blnIsNeo = ModString.ContainsCaseInsensitive(strPath, strDir)
    blnIsNeo = IIf(ModMetaVision.MetaVision_GetDepartment = vbNullString, blnIsNeo, ModMetaVision.MetaVision_GetDepartment = "Neonatologie")
    
    IsNeo = blnIsNeo

End Function

Private Sub Test_IsNeo()

    MsgBox IsNeo()

End Sub

Public Function IsPed() As Boolean

    Dim strPath As String
    Dim strDir As String
    Dim blnIsPed As Boolean

    blnIsPed = Not IsNeo()
    
    IsPed = blnIsPed

End Function

Private Sub Test_IsPed()

    MsgBox IsPed()

End Sub


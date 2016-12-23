Attribute VB_Name = "ModSettings"
Option Explicit

Private Const CONST_DEVMODE = "SettingDevMode"
Private Const CONST_LOGGING = "SettingLogging"
Private Const CONST_NEODIR = "SettingNeoDir"
Private Const CONST_PEDDIR = "SettingPedDir"
Private Const CONST_DEVDIR = "SettingDevDir"
Private Const CONST_TESTLOGDIR = "SettingTestLogDir"
Private Const CONST_LOGDIR = "SettingLogDir"
Private Const CONST_DATADIR = "SettingDataDir"
Private Const CONST_DBDIR = "SettingDbDir"

Private Function GetSetting(strSetting As String) As Variant

    GetSetting = Range(strSetting).Value2

End Function

Private Sub SetSetting(strSetting As String, varValue As Variant)

    Range(strSetting).Value2 = varValue
    
End Sub

Public Function GetDevelopmentMode() As Boolean

    GetDevelopmentMode = CBool(GetSetting(CONST_DEVMODE))

End Function

Public Function IsDevelopmentMode() As Boolean

    Dim blnDevDir As Boolean
    Dim strActDir As String
    
    strActDir = ModConst.GetAfsprakenProgramFilePath()
    blnDevDir = ModString.StringContainsCaseInsensitive(strActDir, GetDevelopmentDir)
    
    IsDevelopmentMode = GetDevelopmentMode() Or blnDevDir

End Function

Public Sub SetDevelopmentMode(blnMode As Boolean)

    SetSetting CONST_DEVMODE, blnMode

End Sub

Public Function GetEnableLogging() As Boolean

    GetEnableLogging = CBool(GetSetting(CONST_LOGGING))

End Function

Public Sub SetEnableLogging(blnMode As Boolean)

    SetSetting CONST_LOGGING, blnMode

End Sub

Public Sub ToggleLogging()

    SetEnableLogging Not GetEnableLogging()

End Sub

Public Function GetNeoDir() As String

    GetNeoDir = CStr(GetSetting(CONST_NEODIR))

End Function

Public Sub SetNeoDir(strDir As String)

    SetSetting CONST_NEODIR, strDir

End Sub

Public Function GetPedDir() As String

    GetPedDir = CStr(GetSetting(CONST_PEDDIR))

End Function

Public Sub SetPedDir(strDir As String)

    SetSetting CONST_PEDDIR, strDir

End Sub

Public Function GetDevelopmentDir() As String

    GetDevelopmentDir = CStr(GetSetting(CONST_DEVDIR))

End Function

Public Sub SetDevelopmentDir(strDir As String)

    SetSetting CONST_DEVDIR, strDir

End Sub

Public Function GetTestLogDir() As String

    GetTestLogDir = CStr(GetSetting(CONST_TESTLOGDIR))

End Function

Public Sub SetTestLogDir(strDir As String)

    SetSetting CONST_TESTLOGDIR, strDir

End Sub

Public Function GetLogDir() As String

    GetLogDir = CStr(GetSetting(CONST_LOGDIR))

End Function

Public Sub SetLogDir(strDir As String)

    SetSetting CONST_LOGDIR, strDir

End Sub

Public Function GetDataDir() As String

    GetDataDir = CStr(GetSetting(CONST_DATADIR))

End Function

Public Sub SetDataDir(strDir As String)

    SetSetting CONST_DATADIR, strDir

End Sub

Public Function GetFormDbDir() As String

    GetFormDbDir = CStr(GetSetting(CONST_DBDIR))

End Function

Public Sub SetFormDbDir(strDir As String)

    SetSetting CONST_DBDIR, strDir

End Sub

Private Sub Test()

    MsgBox GetDevelopmentMode()
    SetDevelopmentMode False
    MsgBox GetDevelopmentMode()
    MsgBox IsDevelopmentMode()

End Sub

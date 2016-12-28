Attribute VB_Name = "ModSetting"
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

Private Function GetSetting(strSetting As String, varDefault As Variant) As Variant

    Dim strMsg As String

    On Error GoTo GetSettingError:

    GetSetting = Range(strSetting).Value2 'ModRange.GetRangeValue(strSetting, varDefault)
    
    Exit Function
    
GetSettingError:

    strMsg = ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & "Kan setting: " & strSetting & " niet ophalen"
    ModMessage.ShowMsgBoxError strMsg

End Function

Private Sub SetSetting(strSetting As String, varValue As Variant)

    Dim strMsg As String

    On Error GoTo SetSettingError:

    Range(strSetting).Value2 = varValue ' ModRange.SetRangeValue strSetting, varValue
    
    Exit Sub

SetSettingError:

    strMsg = ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & "Kan setting: " & strSetting & " niet opslaan"
    ModMessage.ShowMsgBoxError strMsg
    
End Sub

Public Function GetDevelopmentMode() As Boolean

    GetDevelopmentMode = CBool(GetSetting(CONST_DEVMODE, False))

End Function

Public Function IsDevelopmentMode() As Boolean

    Dim blnDevDir As Boolean
    Dim strActDir As String
    
    strActDir = WbkAfspraken.Path
    blnDevDir = ModString.ContainsCaseInsensitive(strActDir, GetDevelopmentDir)
    
    IsDevelopmentMode = GetDevelopmentMode() Or blnDevDir

End Function

Public Sub SetDevelopmentMode(blnMode As Boolean)

    SetSetting CONST_DEVMODE, blnMode

End Sub

Public Function GetEnableLogging() As Boolean

    GetEnableLogging = CBool(GetSetting(CONST_LOGGING, False))

End Function

Public Sub SetEnableLogging(blnMode As Boolean)

    SetSetting CONST_LOGGING, blnMode

End Sub

Public Sub ToggleLogging()

    SetEnableLogging Not GetEnableLogging()

End Sub

Public Function GetNeoDir() As String

    GetNeoDir = CStr(GetSetting(CONST_NEODIR, ""))

End Function

Public Sub SetNeoDir(strDir As String)

    SetSetting CONST_NEODIR, strDir

End Sub

Public Function GetPedDir() As String

    GetPedDir = CStr(GetSetting(CONST_PEDDIR, ""))

End Function

Public Sub SetPedDir(strDir As String)

    SetSetting CONST_PEDDIR, strDir

End Sub

Public Function GetDevelopmentDir() As String

    GetDevelopmentDir = CStr(GetSetting(CONST_DEVDIR, ""))

End Function

Public Sub SetDevelopmentDir(strDir As String)

    SetSetting CONST_DEVDIR, strDir

End Sub

Public Function GetTestLogDir() As String

    GetTestLogDir = CStr(GetSetting(CONST_TESTLOGDIR, ""))

End Function

Public Sub SetTestLogDir(strDir As String)

    SetSetting CONST_TESTLOGDIR, strDir

End Sub

Public Function GetTestLogPath() As String

    GetTestLogPath = WbkAfspraken.Path & "\" & GetTestLogDir()

End Function

Public Function GetLogDir() As String

    GetLogDir = CStr(GetSetting(CONST_LOGDIR, ""))

End Function

Public Sub SetLogDir(strDir As String)

    SetSetting CONST_LOGDIR, strDir

End Sub

Public Function GetLogPath() As String

    GetLogPath = WbkAfspraken.Path & "\" & GetLogDir()

End Function

Public Function GetDataDir() As String

    GetDataDir = CStr(GetSetting(CONST_DATADIR, ""))

End Function

Public Sub SetDataDir(strDir As String)

    SetSetting CONST_DATADIR, strDir

End Sub

Public Function GetFormDbDir() As String

    GetFormDbDir = CStr(GetSetting(CONST_DBDIR, ""))

End Function

Public Sub SetFormDbDir(strDir As String)

    SetSetting CONST_DBDIR, strDir

End Sub

Private Sub Test()

    MsgBox GetEnableLogging()
    SetEnableLogging True
    MsgBox GetEnableLogging()
    
End Sub

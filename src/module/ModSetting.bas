Attribute VB_Name = "ModSetting"
Option Explicit

Private Const constDevMode = "SettingDevMode"
Private Const constLogging = "SettingLogging"
Private Const constNeoDir = "SettingNeoDir"
Private Const constPedDir = "SettingPedDir"
Private Const constDevDir = "SettingDevDir"
Private Const constTestLogDir = "SettingTestLogDir"
Private Const constLogDir = "SettingLogDir"
Private Const constDataDir = "SettingDataDir"
Private Const constDbDir = "SettingDbDir"
Private Const constPrePend = "Patient"
Private Const constPostData = ""
Private Const constPostText = "_AfsprakenTekst"

Private Function GetSetting(strSetting As String, varDefault As Variant) As Variant

    Dim strMsg As String

    On Error GoTo GetSettingError:

    GetSetting = shtGlobSettings.Range(strSetting).Value2 'ModRange.GetRangeValue(strSetting, varDefault)
    
    Exit Function
    
GetSettingError:

    strMsg = ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & "Kan setting: " & strSetting & " niet ophalen"
    ModMessage.ShowMsgBoxError strMsg

End Function

Private Sub SetSetting(strSetting As String, varValue As Variant)

    Dim strMsg As String

    On Error GoTo SetSettingError:

    shtGlobSettings.Range(strSetting).Value2 = varValue ' ModRange.SetRangeValue strSetting, varValue
    
    Exit Sub

SetSettingError:

    strMsg = ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & "Kan setting: " & strSetting & " niet opslaan"
    ModMessage.ShowMsgBoxError strMsg
    
End Sub

Public Function GetDevelopmentMode() As Boolean

    GetDevelopmentMode = CBool(GetSetting(constDevMode, False))

End Function

Public Function IsDevelopmentMode() As Boolean

    Dim blnDevDir As Boolean
    Dim strActDir As String
    
    strActDir = WbkAfspraken.Path
    blnDevDir = ModString.ContainsCaseInsensitive(strActDir, GetDevelopmentDir)
    
    IsDevelopmentMode = GetDevelopmentMode() Or blnDevDir

End Function

Public Sub SetDevelopmentMode(blnMode As Boolean)

    SetSetting constDevMode, blnMode

End Sub

Public Function GetEnableLogging() As Boolean

    GetEnableLogging = CBool(GetSetting(constLogging, False))

End Function

Public Sub SetEnableLogging(blnMode As Boolean)

    SetSetting constLogging, blnMode

End Sub

Public Sub ToggleLogging()

    SetEnableLogging Not GetEnableLogging()

End Sub

Public Function GetNeoDir() As String

    GetNeoDir = CStr(GetSetting(constNeoDir, ""))

End Function

Public Sub SetNeoDir(strDir As String)

    SetSetting constNeoDir, strDir

End Sub

Public Function GetPedDir() As String

    GetPedDir = CStr(GetSetting(constPedDir, ""))

End Function

Public Sub SetPedDir(strDir As String)

    SetSetting constPedDir, strDir

End Sub

Public Function GetDevelopmentDir() As String

    GetDevelopmentDir = CStr(GetSetting(constDevDir, ""))

End Function

Public Sub SetDevelopmentDir(strDir As String)

    SetSetting constDevDir, strDir

End Sub

Public Function GetTestLogDir() As String

    GetTestLogDir = CStr(GetSetting(constTestLogDir, ""))

End Function

Public Sub SetTestLogDir(strDir As String)

    SetSetting constTestLogDir, strDir

End Sub

Public Function GetTestLogPath() As String

    GetTestLogPath = WbkAfspraken.Path & "\" & GetTestLogDir()

End Function

Public Function GetLogDir() As String

    GetLogDir = CStr(GetSetting(constLogDir, ""))

End Function

Public Sub SetLogDir(strDir As String)

    SetSetting constLogDir, strDir

End Sub

Public Function GetLogPath() As String

    GetLogPath = WbkAfspraken.Path & "\" & GetLogDir()

End Function

Public Function GetDataDir() As String

    GetDataDir = CStr(GetSetting(constDataDir, ""))

End Function

Public Sub SetDataDir(strDir As String)

    SetSetting constDataDir, strDir

End Sub

Public Function GetFormDbDir() As String

    GetFormDbDir = CStr(GetSetting(constDbDir, ""))

End Function

Public Sub SetFormDbDir(strDir As String)

    SetSetting constDbDir, strDir

End Sub

Public Function GetPatientDataPath() As String

    Dim strDir As String
    
    strDir = ModSetting.GetDataDir()
    GetPatientDataPath = GetAbsolutePath(strDir)

End Function

Private Function GetAbsolutePath(strPath As String) As String

    GetAbsolutePath = ActiveWorkbook.Path & strPath

End Function

Public Function GetPatientTextWorkBookName(strBed As String) As String

    GetPatientTextWorkBookName = constPrePend & strBed & constPostText & ".xls"

End Function

Public Function GetPatientDataWorkBookName(strBed As String) As String

    GetPatientDataWorkBookName = constPrePend & strBed & constPostData + ".xls"

End Function

Public Function GetPatientDataFile(strBed As String) As String

    GetPatientDataFile = GetPatientDataPath + GetPatientDataWorkBookName(strBed)

End Function

Public Function GetPatientTextFile(strBed As String) As String

    GetPatientTextFile = GetPatientDataPath + GetPatientTextWorkBookName(strBed)

End Function

Private Sub Test()

    MsgBox GetEnableLogging()
    SetEnableLogging True
    MsgBox GetEnableLogging()
    
End Sub

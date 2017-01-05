Attribute VB_Name = "ModSetting"
Option Explicit

Public Const CONST_DATA_SHEET As String = "Data"
Public Const CONST_PATIENTS_SHEET As String = "Patienten"

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
Private Const constPedBeds As String = "tbl_Ped_Beds"
Private Const constNeoBeds As String = "tbl_Neo_Beds"

Private Function GetSetting(ByVal strSetting As String) As Variant

    Dim strMsg As String

    On Error GoTo GetSettingError:

    GetSetting = shtGlobSettings.Range(strSetting).Value2 'ModRange.GetRangeValue(strSetting, varDefault)
    
    Exit Function
    
GetSettingError:

    strMsg = ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & "Kan setting: " & strSetting & " niet ophalen"
    ModMessage.ShowMsgBoxError strMsg

End Function

Private Sub SetSetting(ByVal strSetting As String, ByVal varValue As Variant)

    Dim strMsg As String

    On Error GoTo SetSettingError:

    shtGlobSettings.Range(strSetting).Value2 = varValue ' ModRange.SetRangeValue strSetting, varValue
    
    Exit Sub

SetSettingError:

    strMsg = ModConst.CONST_DEFAULTERROR_MSG & vbNewLine & "Kan setting: " & strSetting & " niet opslaan"
    ModMessage.ShowMsgBoxError strMsg
    
End Sub

Public Function GetDevelopmentMode() As Boolean

    GetDevelopmentMode = CBool(GetSetting(constDevMode))

End Function

Public Function IsDevelopmentMode() As Boolean

    Dim blnDevDir As Boolean
    Dim strActDir As String
    
    strActDir = WbkAfspraken.Path
    blnDevDir = ModString.ContainsCaseInsensitive(strActDir, GetDevelopmentDir)
    
    IsDevelopmentMode = GetDevelopmentMode() Or blnDevDir

End Function

Public Sub SetDevelopmentMode(ByVal blnMode As Boolean)

    SetSetting constDevMode, blnMode

End Sub

Public Function GetEnableLogging() As Boolean

    GetEnableLogging = CBool(GetSetting(constLogging))

End Function

Public Sub SetEnableLogging(ByVal blnMode As Boolean)

    SetSetting constLogging, blnMode

End Sub

Public Sub ToggleLogging()

    SetEnableLogging Not GetEnableLogging()

End Sub

Public Function GetNeoDir() As String

    GetNeoDir = CStr(GetSetting(constNeoDir))

End Function

Public Sub SetNeoDir(ByVal strDir As String)

    SetSetting constNeoDir, strDir

End Sub

Public Function GetPedDir() As String

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

    GetLogPath = WbkAfspraken.Path & "\" & GetLogDir()

End Function

Public Function GetDataDir() As String

    GetDataDir = CStr(GetSetting(constDataDir))

End Function

Public Sub SetDataDir(ByVal strDir As String)

    SetSetting constDataDir, strDir

End Sub

Public Function GetFormDbDir() As String

    GetFormDbDir = CStr(GetSetting(constDbDir))

End Function

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
    arrBeds = Array() ' Assign but keep empty
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

    GetPatientsFileName = constPatientsFile

End Function

Public Function GetPatientsFilePath() As String

    GetPatientsFilePath = GetPatientDataPath() & "\" & constPatientsFile

End Function

Private Sub test()

    MsgBox GetPatientsFilePath
    
End Sub

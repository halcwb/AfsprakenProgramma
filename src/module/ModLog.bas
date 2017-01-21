Attribute VB_Name = "ModLog"
Option Explicit

Public Enum LogLevel
    Error = 0
    Warning = 1
    Info = 2
End Enum

Public Sub ToggleLogging()

    Dim strPw As Variant
    
    strPw = CStr(InputBox("Voer wachtwoord in"))
    If strPw = CONST_PASSWORD Then ModSetting.ToggleLogging

End Sub

Public Sub EnableLogging()

    ModSetting.SetEnableLogging True

End Sub

Public Sub DisableLogging()
    
    ModSetting.SetEnableLogging False

End Sub

Public Function LogLevelToString(ByVal enmLevel As LogLevel) As String

    Select Case enmLevel
        Case 0
            LogLevelToString = "Error"
        Case 1
            LogLevelToString = "Warning"
        Case 2
            LogLevelToString = "Info"
    End Select
    
End Function

Public Sub LogError(ByVal strError As String)

    Dim blnLog As Boolean
    
    strError = " Number: " & err.Number & " Source: " & err.source & " Description: " & strError
    blnLog = ModSetting.GetEnableLogging()

    EnableLogging
    LogToFile ModSetting.GetLogPath(), Error, strError
    If Not blnLog Then ModLog.DisableLogging

End Sub

Public Sub LogInfo(ByVal strInfo As String)

    LogToFile ModSetting.GetLogPath(), Info, strInfo

End Sub

Public Sub LogTest(ByVal enmLevel As LogLevel, ByVal strMsg As String)
    Dim strFile As String

    strFile = WbkAfspraken.Path + ModSetting.GetTestLogDir()
    LogToFile strFile, enmLevel, strMsg
    
End Sub

Public Sub LogActionStart(ByVal strAction As String, ByRef strParams() As Variant)

    Dim strFile As String
    Dim strMsg As String

    strMsg = "Begin " + strAction + ": " + Join(strParams, ", ")

    strFile = WbkAfspraken.Path + ModSetting.GetLogDir()
    LogToFile strFile, Info, strMsg
    
End Sub

Public Sub LogActionEnd(ByVal strAction As String)

    Dim strFile As String
    Dim strMsg As String

    strMsg = "End " + strAction

    strFile = WbkAfspraken.Path + ModSetting.GetLogDir()
    LogToFile strFile, Info, strMsg
    
End Sub

Public Sub LogToFile(ByVal strFile As String, ByVal enmLevel As LogLevel, ByVal strMsg As String)
    
    If Not ModSetting.GetEnableLogging() Then Exit Sub

    strMsg = Replace(strMsg, vbNewLine, ". ")
    AppendToFile strFile, Strings.Format(DateTime.Now, vbNullString) + ": " + LogLevelToString(enmLevel) + ": " + strMsg
    
End Sub



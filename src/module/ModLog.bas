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

Public Function LogLevelToString(enmLevel As LogLevel) As String

    Select Case enmLevel
        Case 0
            LogLevelToString = "Error"
        Case 1
            LogLevelToString = "Warning"
        Case 2
            LogLevelToString = "Info"
    End Select
    
End Function

Public Sub LogError(strError As String)

    Dim blnLog
    
    strError = " Number: " & Err.Number & " Source: " & Err.Source & " Description: " & strError
    blnLog = ModSetting.GetEnableLogging()

    EnableLogging
    LogToFile ModSetting.GetLogPath(), Error, strError
    If Not blnLog Then ModLog.DisableLogging

End Sub

Public Sub LogTest(enmLevel As LogLevel, strMsg As String)
    Dim strFile As String

    strFile = WbkAfspraken.Path + ModSetting.GetTestLogDir()
    LogToFile strFile, enmLevel, strMsg
    
End Sub

Public Sub LogActionStart(strAction As String, strParams() As Variant)

    Dim strFile As String, strMsg As String

    strMsg = "Begin " + strAction + ": " + Join(strParams, ", ")

    strFile = WbkAfspraken.Path + ModSetting.GetLogDir()
    LogToFile strFile, Info, strMsg
    
End Sub

Public Sub LogActionEnd(strAction As String)

    Dim strFile As String, strMsg As String

    strMsg = "End " + strAction

    strFile = WbkAfspraken.Path + ModSetting.GetLogDir()
    LogToFile strFile, Info, strMsg
    
End Sub

Public Sub LogToFile(strFile As String, enmLevel As LogLevel, strMsg As String)
    
    If Not ModSetting.GetEnableLogging() Then Exit Sub

    strMsg = Replace(strMsg, vbNewLine, ". ")
    AppendToFile strFile, Format(DateTime.Now, vbNullString) + ": " + LogLevelToString(enmLevel) + ": " + strMsg
    
End Sub



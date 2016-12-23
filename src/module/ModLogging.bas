Attribute VB_Name = "ModLogging"
Option Explicit

Public Enum LogLevel
    Error = 0
    Warning = 1
    Info = 2
End Enum

Public Sub ToggleLogging()

    Dim strPw As Variant
    
    strPw = CStr(InputBox("Voer wachtwoord in"))
    If strPw = CONST_PASSWORD Then ModSettings.ToggleLogging

End Sub

Public Sub EnableLogging()

    ModSettings.SetEnableLogging True

End Sub

Public Sub DisableLogging()
    
    ModSettings.SetEnableLogging False

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

Public Sub LogTest(enmLevel As LogLevel, strMsg As String)
    Dim strFile As String

    strFile = ModConst.GetAfsprakenProgramFilePath() + ModSettings.GetTestLogDir()
    LogToFile strFile, enmLevel, strMsg
    
End Sub

Public Sub LogActionStart(strAction As String, strParams() As Variant)

    Dim strFile As String, strMsg As String

    strMsg = "Begin " + strAction + ": " + Join(strParams, ", ")

    strFile = ModConst.GetAfsprakenProgramFilePath() + ModSettings.GetLogDir()
    LogToFile strFile, Info, strMsg
    
End Sub

Public Sub LogActionEnd(strAction As String)

    Dim strFile As String, strMsg As String

    strMsg = "End " + strAction

    strFile = ModConst.GetAfsprakenProgramFilePath() + ModSettings.GetLogDir()
    LogToFile strFile, Info, strMsg
    
End Sub

Public Sub LogToFile(strFile As String, enmLevel As LogLevel, strMsg As String)
    
    If Not ModSettings.GetEnableLogging() Then Exit Sub

    strMsg = Replace(strMsg, vbNewLine, ". ")
    AppendToFile strFile, Format(DateTime.Now, vbNullString) + ": " + LogLevelToString(enmLevel) + ": " + strMsg
    
End Sub



Attribute VB_Name = "ModLogging"
Option Explicit

Public Enum LogLevel
    Error = 0
    Warning = 1
    Info = 2
End Enum

Public Sub ToggleLogging()

    Dim pw As String
    pw = InputBox("Voer wachtwoord in")
    If pw = CONST_PASSWORD Then BlnEnableLogging = Not BlnEnableLogging

End Sub

Public Sub EnableLogging()

    BlnEnableLogging = True

End Sub

Public Sub DisableLogging()
    
    BlnEnableLogging = False

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

    strFile = ModGlobal.GetAfsprakenProgramFilePath() + ModGlobal.CONST_TEST_CONST_LOGPATH
    ModLogging.LogToFile strFile, enmLevel, strMsg
    
End Sub

Public Sub LogActionStart(strAction As String, strParams() As Variant)

    Dim strFile As String, strMsg As String

    strMsg = "Begin " + strAction + ": " + Join(strParams, ", ")

    strFile = ModGlobal.GetAfsprakenProgramFilePath() + ModGlobal.CONST_LOGPATH
    LogToFile strFile, Info, strMsg
    
End Sub

Public Sub LogActionEnd(strAction As String)

    Dim strFile As String, strMsg As String

    strMsg = "End " + strAction

    strFile = ModGlobal.GetAfsprakenProgramFilePath() + ModGlobal.CONST_LOGPATH
    LogToFile strFile, Info, strMsg
    
End Sub


Public Sub LogToFile(strFile As String, enmLevel As LogLevel, strMsg As String)
    
    If Not BlnEnableLogging Then Exit Sub

    strMsg = Replace(strMsg, vbNewLine, ". ")
    AppendToFile strFile, Format(DateTime.Now, vbNullString) + ": " + LogLevelToString(enmLevel) + ": " + strMsg
    
End Sub



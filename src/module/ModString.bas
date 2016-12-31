Attribute VB_Name = "ModString"
Option Explicit

' Checks whether strString contains strValue case insensitive, ignores spaces
Public Function ContainsCaseInsensitive(ByVal strString As String, ByVal strValue As String) As Boolean
    strString = Trim(strString)
    strValue = Trim(strValue)
    
    If InStr(1, LCase(strString), LCase(strValue)) > 0 Then
        ContainsCaseInsensitive = True
    Else
        ContainsCaseInsensitive = False
    End If
    
End Function

' Checks whether strString contains strValue case sensitive, ignores spaces
Public Function ContainsCaseSensitive(ByVal strString As String, ByVal strValue As String) As Boolean
    strString = Trim(strString)
    strValue = Trim(strValue)

    If InStr(1, strString, strValue) > 0 Then
        ContainsCaseSensitive = True
    Else
        ContainsCaseSensitive = False
    End If
    
End Function

Public Function StartsWith(ByVal strString, strValue As String) As Boolean

    StartsWith = InStr(1, strString, strValue) = 1

End Function

Private Sub Test()
    Dim strString As String
    Dim strStart As String
    
    strString = "__4_GebDatum"
    strStart = "_"

    MsgBox StartsWith(strString, strStart)

End Sub

Public Function StringToDate(ByVal strValue As String) As Date

    Dim dtmDate As Date
    
    On Error GoTo StringToDateError
    
    dtmDate = CDate(strValue)
    StringToDate = dtmDate
    
    Exit Function
    
StringToDateError:

    ModLog.LogError "Cannot convert " & strValue & " to a date time"

End Function





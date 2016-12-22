Attribute VB_Name = "ModString"
Option Explicit

' Checks whether strString contains strValue case insensitive, ignores spaces
Public Function StringContainsCaseInsensitive(ByVal strString As String, ByVal strValue As String) As Boolean
    strString = Trim(strString)
    strValue = Trim(strValue)
    
    If InStr(1, LCase(strString), LCase(strValue)) > 0 Then
        StringContainsCaseInsensitive = True
    Else
        StringContainsCaseInsensitive = False
    End If
    
End Function

' Checks whether strString contains strValue case sensitive, ignores spaces
Public Function StringContainsCaseSensitive(ByVal strString As String, ByVal strValue As String) As Boolean
    strString = Trim(strString)
    strValue = Trim(strValue)

    If InStr(1, strString, strValue) > 0 Then
        StringContainsCaseSensitive = True
    Else
        StringContainsCaseSensitive = False
    End If
    
End Function


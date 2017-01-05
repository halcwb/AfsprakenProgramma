Attribute VB_Name = "ModString"
Option Explicit

'Public Function StartsWith(ByVal strString, ByVal strValue As String) As Boolean
'
'    If Len(strString) >= Len(strValue) Then
'        If Mid(strString, 0, Len(strValue)) = strValue Then
'            StartsWith = True
'            Exit Function
'        End If
'    End If
'
'    StartsWith = False
'
'End Function

' Checks whether strString contains strValue case insensitive, ignores spaces
Public Function ContainsCaseInsensitive(ByVal strString As String, ByVal strValue As String) As Boolean
    strString = Strings.Trim(strString)
    strValue = Strings.Trim(strValue)
    
    If Strings.InStr(1, Strings.LCase(strString), Strings.LCase(strValue)) > 0 Then
        ContainsCaseInsensitive = True
    Else
        ContainsCaseInsensitive = False
    End If
    
End Function

' Checks whether strString contains strValue case sensitive, ignores spaces
Public Function ContainsCaseSensitive(ByVal strString As String, ByVal strValue As String) As Boolean
    strString = Strings.Trim(strString)
    strValue = Strings.Trim(strValue)

    If Strings.InStr(1, strString, strValue) > 0 Then
        ContainsCaseSensitive = True
    Else
        ContainsCaseSensitive = False
    End If
    
End Function

Public Function StartsWith(ByVal strString As String, ByVal strValue As String) As Boolean

    StartsWith = Strings.InStr(1, strString, strValue) = 1

End Function

Private Sub test()
    Dim strString As String
    Dim strStart As String
    
    strString = "__4_GebDatum"
    strStart = "_"

    MsgBox StartsWith(strString, strStart)

End Sub

Public Function DateToString(ByVal dtmDate As Date) As String

    DateToString = Strings.Format(dtmDate, "dd-mmm-yyyy")

End Function

Private Sub TestDateToString()

    MsgBox DateToString(DateTime.Date)

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

Private Sub TestStringToDate()

    MsgBox DateToString(StringToDate("01-02-2017"))

End Sub





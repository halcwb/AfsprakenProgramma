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

Public Function DateToString(ByVal dtmDate As Date) As String

    DateToString = Format(dtmDate, "dd-mmm-yyyy")

End Function

Private Sub TestDateToString()

    MsgBox DateToString(Date)

End Sub

Public Function StringToDate(ByVal strValue As String) As Date

    Dim dtmDate As Date
    Dim intLocale As Integer
    
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





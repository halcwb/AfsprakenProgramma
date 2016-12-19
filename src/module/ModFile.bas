Attribute VB_Name = "ModFile"
Option Explicit

' Write strText to a file
' Creates the file if it doesn't exist
' Appends the text to the file if it does exist
Public Sub AppendToFile(strFile As String, strText As String)

    Open strFile For Append As #1
    Write #1, strText
    Close #1

End Sub

' Write strText to a file
' Creates the file if it doesn't exist
' Overwrites the file if it does exist
Public Sub WriteToFile(strFile As String, ByVal strText As String)

    Open strFile For Output As #1
    Write #1, strText
    Close #1

End Sub


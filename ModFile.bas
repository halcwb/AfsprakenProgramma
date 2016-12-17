Attribute VB_Name = "ModFile"
Option Explicit

Public Sub AppendToFile(strFilePath As String, strText As String)

    Open strFilePath For Append As #1
    Write #1, strText
    Close #1

End Sub


Attribute VB_Name = "ModFile"
Option Explicit

' Write strText to a file
' Creates the file if it doesn't exist
' Appends the text to the file if it does exist
Public Sub AppendToFile(ByVal strFile As String, ByVal strText As String)

    Dim strError As String

    On Error GoTo AppenToFileError

    Open strFile For Append As #1
    Write #1, strText
    Close #1
    
    Exit Sub
    
AppenToFileError:

    strError = "AppenToFileError" & vbNewLine
    strError = strError & "Kan '" & strText & "' niet wegschrijven naar '" & strFile & "'"
    strError = strError & vbNewLine & ModConst.CONST_DEFAULTERROR_MSG
    ModMessage.ShowMsgBoxError strError

End Sub

' Write strText to a file
' Creates the file if it doesn't exist
' Overwrites the file if it does exist
Public Sub WriteToFile(strFile As String, ByVal strText As String)

    Dim strError As String

    On Error GoTo WriteToFileError

    Open strFile For Output As #1
    Write #1, strText
    Close #1
    
    Exit Sub

WriteToFileError:

    strError = "WriteToFileError" & vbNewLine
    strError = strError & "Kan '" & strText & "' niet wegschrijven naar '" & strFile & "'"
    strError = strError & vbNewLine & ModConst.CONST_DEFAULTERROR_MSG
    ModMessage.ShowMsgBoxError strError

End Sub

Public Function FileExists(strFile As String) As Boolean

    Dim objFs As New FileSystemObject
    Dim blnExists As Boolean
    
    blnExists = objFs.FileExists(strFile)
    Set objFs = Nothing
    FileExists = blnExists

End Function

Public Sub FileDelete(strFile As String)
    
    Dim objFs As New FileSystemObject
    
    If FileExists(strFile) Then
        objFs.DeleteFile strFile
    End If

End Sub

Public Function GetFiles(strDir As String) As String()

    Dim arrFiles() As String
    Dim strFile As String
    Dim intN As Integer
    
    strFile = Dir(strDir & "*.*")
    intN = -1
    
    Do While Len(strFile) > 0
        intN = intN + 1
        ReDim Preserve arrFiles(intN)
        arrFiles(intN) = strFile
        strFile = Dir
    Loop
    
    GetFiles = arrFiles

End Function

Public Function StringArrayNotEmpty(arrArray() As String) As Boolean

    StringArrayNotEmpty = Len(Join(arrArray)) > 0

End Function

Public Sub DeleteAllFilesInDir(strDir As String)

    Dim varFile As Variant
    Dim arrFiles() As String
    
    arrFiles = GetFiles(strDir)
    
    If StringArrayNotEmpty(arrFiles) Then
    
        For Each varFile In arrFiles
            FileDelete (strDir & varFile)
        Next varFile
    
    End If

End Sub

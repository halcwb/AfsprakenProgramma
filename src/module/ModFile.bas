Attribute VB_Name = "ModFile"
Option Explicit

Private blnShowOnce As Boolean

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
    
    If Not blnShowOnce Then
        strError = "AppenToFileError" & vbNewLine
        strError = strError & "Kan '" & strText & "' niet wegschrijven naar '" & strFile & "'"
        ModMessage.ShowMsgBoxError strError
        
        blnShowOnce = True
    End If

End Sub

' Write strText to a file
' Creates the file if it doesn't exist
' Overwrites the file if it does exist
Public Sub WriteToFile(ByVal strFile As String, ByVal strText As String)

    Dim strError As String

    On Error GoTo WriteToFileError

    Open strFile For Output As #1
    Write #1, strText
    Close #1
    
    Exit Sub

WriteToFileError:

    strError = "WriteToFileError" & vbNewLine
    strError = strError & "Kan '" & strText & "' niet wegschrijven naar '" & strFile & "'"
    ModMessage.ShowMsgBoxError strError

End Sub

Public Function FileExists(ByVal strFile As String) As Boolean

    Dim objFs As FileSystemObject
    Dim blnExists As Boolean
    
    Set objFs = New FileSystemObject
    
    blnExists = objFs.FileExists(strFile)
    Set objFs = Nothing
    FileExists = blnExists

End Function

Public Function ReadFile(ByVal strFile As String) As String

    Dim objFs As FileSystemObject
    Dim objFile As File
    Dim objStream As TextStream
    Dim strLines As String
    
    Set objFs = New FileSystemObject
    
    If FileExists(strFile) Then
        Set objFile = objFs.GetFile(strFile)
        Set objStream = objFile.OpenAsTextStream(ForReading)
        strLines = objStream.ReadAll
    End If
    
    ReadFile = strLines
    
    Set objFs = Nothing
    Set objFile = Nothing
    Set objStream = Nothing

End Function

Private Sub Test_ReadFile()

    MsgBox ReadFile(WbkAfspraken.Path & "\" & "secret")

End Sub


Public Sub FileDelete(ByVal strFile As String)
    
    Dim objFs As FileSystemObject
    
    Set objFs = New FileSystemObject
    
    If FileExists(strFile) Then
        objFs.DeleteFile strFile
    End If

End Sub

Public Function GetFilesByPattern(ByVal strDir As String, ByVal strPattern As String) As String()

    Dim arrFiles() As String
    Dim strFile As String
    Dim intN As Integer
    
    strFile = Dir(strDir & "\" & strPattern)
    intN = -1
    
    Do While Len(strFile) > 0
        intN = intN + 1
        ReDim Preserve arrFiles(intN)
        arrFiles(intN) = strFile
        strFile = Dir
    Loop
    
    GetFilesByPattern = arrFiles

End Function

Public Function GetFiles(ByVal strDir As String) As String()

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

Public Sub DeleteAllFilesInDir(ByVal strDir As String)

    Dim varFile As Variant
    Dim arrFiles() As String
    
    arrFiles = GetFiles(strDir)
    
    If StringArrayNotEmpty(arrFiles) Then
    
        For Each varFile In arrFiles
            FileDelete (strDir & varFile)
        Next varFile
    
    End If

End Sub

Public Function File_GetTestFile() As String

    Dim dlgFile As FileDialog
    Dim varFile As Variant

    Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
    With dlgFile
        .AllowMultiSelect = False
        .InitialFileName = IIf(ModSetting.IsDevelopmentDir, WbkAfspraken.Path & "\tests\", WbkAfspraken.Path & "\..\Tests\")
        If .Show Then
            For Each varFile In .SelectedItems
                If Not varFile = vbNullString Then Exit For
                
            Next
        End If
    End With
    Set dlgFile = Nothing
    
    File_GetTestFile = IIf(IsEmpty(varFile), vbNullString, CStr(varFile))

End Function

Public Function File_GetFolderPath(ByVal strFile As String) As String

    File_GetFolderPath = Left(strFile, InStrRev(strFile, "\"))

End Function

Private Sub Test_File_GetFolderPath()

    MsgBox WbkAfspraken.Path
    
End Sub
    
Public Function GetFileWithDialog() As String
    
    Dim dlgFile As FileDialog
    Dim varFile As Variant
    
    Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
    With dlgFile
        .AllowMultiSelect = False
        .InitialFileName = IIf(ModSetting.IsDevelopmentDir, WbkAfspraken.Path & "\tests\", WbkAfspraken.Path & "\..\Tests\")
        If .Show Then
            For Each varFile In .SelectedItems
                If Not varFile = vbNullString Then Exit For
                
            Next
        End If
    End With
    
    GetFileWithDialog = CStr(varFile)

End Function

Private Sub Test_GetFileWithDialog()

    ModMessage.ShowMsgBoxInfo "File choosen: " & GetFileWithDialog()

End Sub
    
Public Function GetFolderWithDialog() As String
    
    Dim dlbFolder As FileDialog
    Dim varFile As Variant
    
    Set dlbFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With dlbFolder
        .AllowMultiSelect = False
        .InitialFileName = IIf(ModSetting.IsDevelopmentDir, WbkAfspraken.Path & "\tests\", WbkAfspraken.Path & "\..\Tests\")
        If .Show Then
            For Each varFile In .SelectedItems
                If Not varFile = vbNullString Then Exit For
                
            Next
        End If
    End With
    
    GetFolderWithDialog = CStr(varFile)

End Function

Private Sub Test_GetFolderWithDialog()

    ModMessage.ShowMsgBoxInfo "Folder choosen: " & GetFolderWithDialog()

End Sub

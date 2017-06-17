Attribute VB_Name = "ModWord"
Option Explicit

Public Function ReadWordFile(ByVal strFile As String, ByVal strPw As String) As String

    Dim objWord As Object
    Dim objDoc As Object
    Dim objRange As Object
    Dim strText As String
    
    On Error GoTo ErrorHandler
    
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
    objWord.DisplayAlerts = False
    
    objWord.Documents.Open strFile, ReadOnly:=True, PasswordDocument:=strPw
    Set objDoc = objWord.ActiveDocument
    Set objRange = objDoc.Content
    strText = objWord.CleanString(objRange.Text)
    
    objWord.Quit
    
    Set objWord = Nothing
    Set objDoc = Nothing
    Set objRange = Nothing
    
    ReadWordFile = strText
    
    Exit Function

ErrorHandler:

    MsgBox "Could not open " & strFile
    
    objWord.Quit
    
    Set objWord = Nothing
    Set objDoc = Nothing
    Set objRange = Nothing

End Function

Private Sub Test_ReadWordFile()

    MsgBox ReadWordFile(WbkAfspraken.Path & "\" & "secret.docx", "hlab27")

End Sub

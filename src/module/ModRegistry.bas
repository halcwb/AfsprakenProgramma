Attribute VB_Name = "ModRegistry"
Option Explicit

Private Function CreateKey(ByVal strKeyPath As String, ByVal strValueName As String) As String

    Dim strSlash As String
    
    strSlash = IIf(Mid(strKeyPath, Len(strKeyPath) - 1, Len(strKeyPath)) = "\", vbNullString, "\")
    CreateKey = strKeyPath & strSlash & strValueName

End Function

Public Function ReadRegistryKey(ByVal strKeyPath As String, ByVal strValueName As String) As String

    Dim objShell As Object
    
    ModLog.LogInfo "Read: " & strKeyPath & vbNewLine & "Value: " & strValueName
    
    Set objShell = CreateObject("WScript.Shell")
    
    If RegistryKeyExists(strKeyPath, strValueName) Then
        ReadRegistryKey = objShell.RegRead(CreateKey(strKeyPath, strValueName))
    Else
        ReadRegistryKey = vbNullString
    End If
    
End Function

Public Function RegistryKeyExists(ByVal strKeyPath As String, ByVal strValueName As String) As Boolean
    Dim objShell As Object
    
    On Error GoTo ErrorHandler
    'access Windows scripting
    Set objShell = CreateObject("WScript.Shell")
    'try to read the registry key
    objShell.RegRead CreateKey(strKeyPath, strValueName)
    'key was found
    RegistryKeyExists = True
    Exit Function
  
ErrorHandler:
  'key was not found
  RegistryKeyExists = False
End Function

Private Sub Test_ReadRegistryKey()

    MsgBox ReadRegistryKey("HKEY_CURRENT_USER\Software\UMCU\MV", "UserLogin")

End Sub

Private Sub Test_Read_UMCU_Registry()

    Dim objShell As Object
    Dim strKeyPath As String
    
    Set objShell = CreateObject("WScript.Shell")
    
    strKeyPath = "HKCU\SOFTWARE\UMCU\MV\Afdeling"
    MsgBox CreateKey(strKeyPath, vbNullString)
    MsgBox objShell.RegRead(strKeyPath)
    
    Set objShell = Nothing

End Sub


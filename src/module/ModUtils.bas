Attribute VB_Name = "ModUtils"
Option Explicit

' Exports all vba source code and
' formulas in sheets to source tree
' to facilitate source control
Public Sub ExportForSourceControl()

    Dim strPath As String
    
    strPath = WbkAfspraken.Path & "\src\"
    
    DeleteSourceFiles

    ExportFormulas
    ExportVbaCode
    ExportNames
    
    ModMessage.ShowMsgBoxInfo "All code and formulas has been exported to: " & strPath

End Sub

Public Sub DeleteSourceFiles()

    Dim strPath As String
    
    strPath = WbkAfspraken.Path & "\src\"
    
    ModFile.DeleteAllFilesInDir strPath & "sheet\"
    ModFile.DeleteAllFilesInDir strPath & "class\"
    ModFile.DeleteAllFilesInDir strPath & "document\"
    ModFile.DeleteAllFilesInDir strPath & "form\"
    ModFile.DeleteAllFilesInDir strPath & "module\"
    ModFile.DeleteAllFilesInDir strPath & "name\"

End Sub

Public Sub ExportFormulas()

    Dim shtSheet As Worksheet
    Dim objCell As Range
    Dim strText, strPath As String
    
    strPath = WbkAfspraken.Path & "\src\sheet\"
    
    For Each shtSheet In ActiveWorkbook.Sheets
    
        strText = ""
    
        If ModSheet.IsUserInterface(shtSheet) Then
            shtSheet.Unprotect ModConst.CONST_PASSWORD
        End If
    
        For Each objCell In shtSheet.Range("A1:AX200")
            
            If objCell.HasFormula Then
                strText = strText & objCell.AddressLocal & ": " & vbTab & objCell.Formula & vbNewLine
            End If
        
        Next
        
        If strText <> vbNullString Then ModFile.WriteToFile strPath & shtSheet.Name & ".txt", strText
        
        If ModSheet.IsUserInterface(shtSheet) Then
            shtSheet.Protect ModConst.CONST_PASSWORD
        End If
        
    Next shtSheet

End Sub

Public Sub ExportNames()

    Dim objName As Name
    Dim strText As String
    Dim strPath As String
    
    strPath = WbkAfspraken.Path & "\src\name\names.txt"
    
    For Each objName In WbkAfspraken.Names
        strText = strText & objName.NameLocal & ":" & vbTab & objName.RefersTo & vbNewLine
    Next objName
    
    ModFile.WriteToFile strPath, strText

End Sub

Public Sub ExportVbaCode()

    Dim vbcItem As VBComponent
    Dim strFile As String
    Dim strPath As String
    
    strPath = WbkAfspraken.Path

    For Each vbcItem In ActiveWorkbook.VBProject.VBComponents
        strFile = GetComponentFileName(vbcItem)
        vbcItem.Export (strPath & "\src\" & strFile)
    Next

End Sub

Public Function GetComponentFileName(vbcComp As VBComponent) As String

        Dim strExt As String
        Dim strPath As String

        Select Case vbcComp.Type
        
        Case vbext_ComponentType.vbext_ct_ClassModule
            strPath = "class"
            strExt = ".cls"
        Case vbext_ComponentType.vbext_ct_Document
            strPath = "document"
            strExt = ".doccls"
        Case vbext_ComponentType.vbext_ct_MSForm
            strPath = "form"
            strExt = ".frm"
        Case vbext_ComponentType.vbext_ct_StdModule
            strPath = "module"
            strExt = ".bas"
        Case Else
            Err.Raise 17, "GetComponentFileName", "ComponentType not supported: " & vbext_ComponentType.vbext_ct_ActiveXDesigner
        End Select
        
        GetComponentFileName = strPath & "\" & vbcComp.Name & strExt

End Function

Public Sub RunShell(strCmd As String, strArgs() As String)

    Dim strPath As String
    Dim dblExit As Double
    Dim strArg As Variant
    
    strPath = WbkAfspraken.Path
    
    For Each strArg In strArgs
        strCmd = strCmd & " " & strArg
    Next strArg
    
    Let dblExit = Shell(strPath & "\" & strCmd, vbNormalFocus)
    
    If dblExit > 0 Then
        MsgBox "Succesfully ran: " & strCmd
    Else
        MsgBox strCmd & " did not end successfully"
    End If

End Sub

' Check key ascii and make sure it is a valid number character
' Use both dot and comma as decimal separators.
Public Function CorrectNumberAscii(ByVal intKey As Integer) As Integer

    Dim intDot As Integer
    Dim intComma As Integer
    
    intDot = 46
    intComma = 44
    
    If intKey = intDot Or intKey = intComma Then
        ' Use both dot as comma as decimal separator
        intKey = Asc(Application.DecimalSeparator)
    Else
        If intKey >= 48 And intKey <= 57 Then
            ' Key ascii is OK
            ' Key ascii remains the same
        Else
            ' Any other value ignore key and beep
            intKey = 0
            Beep
        End If
    End If
    
    CorrectNumberAscii = intKey

End Function


Public Sub RunTestCmd()
    Dim strArgs(1) As String
    
    strArgs(0) = "git"
    strArgs(1) = "status"

    RunShell "run.cmd", strArgs

End Sub




Attribute VB_Name = "ModUtils"
Option Explicit


Public Sub ExportForSourceControl()

    Dim strPath As String
    
    strPath = ModGlobal.GetAfsprakenProgramFilePath & "\src\"

    ExportFormulas
    ExportVbaCode
    
    MsgBox "All code and formulas has been exported to: " & strPath

End Sub

Public Sub ExportFormulas()

    Dim shtSheet As Worksheet
    Dim objCell As Range
    Dim strText, strPath As String
    
    strPath = ModGlobal.GetAfsprakenProgramFilePath() & "\src\sheet\"
    
    For Each shtSheet In ActiveWorkbook.Sheets
    
        strText = ""
    
        shtSheet.Unprotect ModGlobal.CONST_PASSWORD
    
        For Each objCell In shtSheet.Range("A1:AX200")
            
            If objCell.HasFormula Then
                strText = strText & objCell.AddressLocal & ": " & vbTab & objCell.Formula & vbNewLine
            End If
        
        Next
        
        If strText <> vbNullString Then ModFile.WriteToFile strPath & shtSheet.Name & ".txt", strText
        
        shtSheet.Protect ModGlobal.CONST_PASSWORD
        
    Next shtSheet

End Sub

Public Sub ExportVbaCode()

    Dim vbcItem As VBComponent
    Dim strFile As String
    Dim strPath As String
    
    strPath = ModGlobal.GetAfsprakenProgramFilePath()

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
    
    strPath = ModGlobal.GetAfsprakenProgramFilePath()
    
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

Public Sub RunTestCmd()
    Dim strArgs(1) As String
    
    strArgs(0) = "git"
    strArgs(1) = "status"

    RunShell "run.cmd", strArgs

End Sub


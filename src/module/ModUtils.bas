Attribute VB_Name = "ModUtils"
Option Explicit

' Exports all vba source code and
' formulas in sheets to source tree
' to facilitate source control
Public Sub ExportForSourceControl()

    Dim strPath As String
    
    On Error GoTo ExportForSourceControlError
    
    strPath = WbkAfspraken.Path & "\src\"
    
    ModProgress.StartProgress "Exporting Source Files"
    
    DeleteSourceFiles

    ExportFormulas True
    ExportVbaCode True
    ExportNames True
    
    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxInfo "All code and formulas has been exported to: " & strPath
    
    Exit Sub

ExportForSourceControlError:

    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxExclam "Failed to export for source control"

End Sub

Public Sub CopyToClipboard(ByVal strText As String)

    Dim objClip As MSForms.DataObject
    
    Set objClip = New MSForms.DataObject
    If Not strText = vbNullString Then
        objClip.SetText strText
        objClip.PutInClipboard
    End If
    
End Sub

Public Function GetField(objRs As Recordset, ByVal strField As String) As Variant

    If Not IsNull(objRs.Fields(strField)) Then
        GetField = objRs.Fields(strField)
    Else
        GetField = vbNullString
    End If

End Function

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

Public Sub ExportFormulas(ByVal blnShowProgress As Boolean)

    Dim shtSheet As Worksheet
    Dim objCell As Range
    Dim intN As Integer
    Dim intC As Integer
    Dim strText As String
    Dim strPath As String
    Dim blnProtected As Boolean
    
    strPath = WbkAfspraken.Path & "\src\sheet\"
    
    intN = 1
    intC = WbkAfspraken.Sheets.Count
    For Each shtSheet In WbkAfspraken.Sheets
    
        blnProtected = False
        strText = vbNullString
    
        If shtSheet.ProtectContents Then
            shtSheet.Unprotect ModConst.CONST_PASSWORD
            blnProtected = True
        End If
    
        For Each objCell In shtSheet.Range("A1:AX200")
            
            If objCell.HasFormula Then
                strText = strText & objCell.AddressLocal & ": " & vbTab & objCell.Formula & vbNewLine
            End If
        
        Next
        
        If strText <> vbNullString Then ModFile.WriteToFile strPath & shtSheet.Name & ".txt", strText
        
        If blnProtected Then
            shtSheet.Protect ModConst.CONST_PASSWORD
        End If
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Export Formulas", intC, intN
        intN = intN + 1
        
    Next shtSheet

End Sub

Public Sub ExportNames(ByVal blnShowProgress As Boolean)

    Dim objName As Name
    Dim strText As String
    Dim strPath As String
    Dim intN As Integer
    Dim intC As Integer
    
    strPath = WbkAfspraken.Path & "\src\name\names.txt"
    
    intN = 1
    intC = WbkAfspraken.Names.Count
    For Each objName In WbkAfspraken.Names
        strText = strText & objName.NameLocal & ":" & vbTab & objName.RefersTo & vbNewLine
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Export Names", intC, intN
        intN = intN + 1
    Next objName
    
    ModFile.WriteToFile strPath, strText

End Sub

Public Sub ExportVbaCode(ByVal blnShowProgress As Boolean)

    Dim vbcItem As VBComponent
    Dim strError As String
    Dim strFile As String
    Dim strPath As String
    Dim intN As Integer
    Dim intC As Integer
    
    On Error GoTo ExportVbaCodeError:
    
    strPath = WbkAfspraken.Path
    
    intN = 1
    intC = WbkAfspraken.VBProject.VBComponents.Count
    For Each vbcItem In WbkAfspraken.VBProject.VBComponents
        strFile = GetComponentFileName(vbcItem)
        vbcItem.Export (strPath & "\src\" & strFile)
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Export VBA Code", intC, intN
        intN = intN + 1
    Next
    
    Exit Sub
    
ExportVbaCodeError:
    
    strError = "Kan VBA bestanden niet exporteren." & vbNewLine
    strError = strError & "Waarschijnlijk is het Afspraken project niet geopend (beveiligd met een passwoord)." & vbNewLine
    strError = strError & "Ook moet Tools|macro|security|trusted publishers tab|check trust access to visual basic project." & vbNewLine
    strError = strError & "Open dit eerst en probeer het opnieuw"

    ModMessage.ShowMsgBoxError strError
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

Public Sub RunShell(ByVal strCmd As String, strArgs() As String)

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

Public Function OnlyNumericAscii(ByVal intKey As Integer) As Integer

    If intKey >= 48 And intKey <= 57 Then
        ' Key ascii is OK
    Else
        ' Not a numeric ascii
        intKey = 0
        Beep
    End If
    
    OnlyNumericAscii = intKey

End Function


Public Sub RunTestCmd()
    Dim strArgs(1) As String
    
    strArgs(0) = "git"
    strArgs(1) = "status"

    RunShell "run.cmd", strArgs

End Sub

Public Sub EMailMessageToBeheer(ByVal strMsg As String)

    Dim objMsg As Object
    Dim strTo As String
    Dim strFrom As String
    Dim strSubject As String
    Dim strHTML As String
    
    On Error Resume Next
    
    Err.Clear
    Set objMsg = CreateObject("CDO.Message")
    
    strTo = "c.w.bollen@umcutrecht.nl"
    strFrom = "FunctioneelBeheerMetavision@umcutrecht.nl"
    strSubject = "AfsprakenProgramma fout"
    strHTML = strMsg
     
    With objMsg
         .To = CStr(strTo)
         .From = CStr(strFrom)
         .Subject = CStr(strSubject)
         .HTMLBody = CStr(strHTML)
         .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPickup=1, cdoSendUsingPort=2, cdoSendUsingExchange=3
         .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.umcutrecht.nl"
         .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
         .Configuration.Fields.Update
         .Send
    End With
    
    Set objMsg = Nothing
    
End Sub



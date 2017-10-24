Attribute VB_Name = "ModMessage"
Option Explicit

Private Function ShowMsgBox(ByVal strText As String, ByVal enmStyle As VbMsgBoxStyle) As VbMsgBoxResult

    ShowMsgBox = MsgBox(strText, enmStyle, ModConst.CONST_APPLICATION_NAME)

End Function

Public Function ShowMsgBoxOK(ByVal strText As String) As VbMsgBoxResult

    ShowMsgBoxOK = ShowMsgBox(strText, vbOKOnly)
    
End Function

Public Function ShowMsgBoxOKCancel(ByVal strText As String) As VbMsgBoxResult

    ShowMsgBoxOKCancel = ShowMsgBox(strText, vbOKCancel)
    
End Function

Public Function ShowMsgBoxInfo(ByVal strText As String) As VbMsgBoxResult

    ShowMsgBoxInfo = ShowMsgBox(strText, vbInformation)
    
End Function

Public Function ShowMsgBoxError(ByVal strText As String) As VbMsgBoxResult

    strText = strText & vbNewLine & ModConst.CONST_DEFAULTERROR_MSG
    ShowMsgBoxError = ShowMsgBox(strText, vbCritical)
    
End Function

Public Function ShowMsgBoxExclam(ByVal strText As String) As VbMsgBoxResult

    ShowMsgBoxExclam = ShowMsgBox(strText, vbExclamation)
    
End Function

Public Function ShowMsgBoxYesNo(ByVal strText As String) As VbMsgBoxResult

    ShowMsgBoxYesNo = ShowMsgBox(strText, vbYesNo)
    
End Function

Public Function ShowPasswordBox(ByVal strText As String) As String

    Dim objPw As FormPassword
    Dim strPw As String
    
    Set objPw = New FormPassword
    objPw.Caption = strText
    objPw.Show
    
    If objPw.lblValid.Caption = vbNullString Then strPw = objPw.txtPassword.Value
    
    ShowPasswordBox = strPw

End Function

Attribute VB_Name = "ModMessage"
Option Explicit

Private Function ShowMsgBox(strText As String, enmStyle As VbMsgBoxStyle) As VbMsgBoxResult

    ShowMsgBox = MsgBox(strText, enmStyle, ModConst.CONST_APPLICATION_NAME)

End Function

Public Function ShowMsgBoxOK(strText As String) As VbMsgBoxResult

    ShowMsgBoxOK = ShowMsgBox(strText, vbOKOnly)
    
End Function

Public Function ShowMsgBoxOKCancel(strText As String) As VbMsgBoxResult

    ShowMsgBoxOKCancel = ShowMsgBox(strText, vbOKCancel)
    
End Function

Public Function ShowMsgBoxInfo(strText As String) As VbMsgBoxResult

    ShowMsgBoxInfo = ShowMsgBox(strText, vbInformation)
    
End Function

Public Function ShowMsgBoxError(strText As String) As VbMsgBoxResult

    ShowMsgBoxError = ShowMsgBox(strText, vbCritical)
    
End Function

Public Function ShowMsgBoxExclam(strText As String) As VbMsgBoxResult

    ShowMsgBoxExclam = ShowMsgBox(strText, vbExclamation)
    
End Function

Public Function ShowMsgBoxYesNo(strText As String) As VbMsgBoxResult

    ShowMsgBoxYesNo = ShowMsgBox(strText, vbYesNo)
    
End Function


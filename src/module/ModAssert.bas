Attribute VB_Name = "ModAssert"
Option Explicit

Public Sub TestAssert()

    AssertEqual 1, 2, "TestAssert, One is not two", False
    AssertNotEqual 1, 1, "TestAssert, One is not not blnEqual to one", False

End Sub

Public Sub AssertTrue(ByVal blnIsTrue As Boolean, ByVal strMsg As String, ByVal blnDisplay As Boolean)

    If Not blnIsTrue Then ProcessAssert True, False, strMsg, True, blnDisplay

End Sub

Public Sub AssertEqual(ByVal varV1 As Variant, ByVal varV2 As Variant, ByVal strMsg As String, ByVal blnDisplay As Boolean)

    If varV1 <> varV2 Then ProcessAssert varV1, varV2, strMsg, True, blnDisplay

End Sub

Public Sub AssertNotEqual(ByVal varV1 As Variant, ByVal varV2 As Variant, ByVal strMsg As String, ByVal blnDisplay As Boolean)

    If varV1 = varV2 Then ProcessAssert varV1, varV2, strMsg, False, blnDisplay

End Sub

Public Sub DisplayAssert(ByVal strMsg As String)

    ModMessage.ShowMsgBoxExclam strMsg

End Sub

Private Sub ProcessAssert(ByVal varV1 As Variant, ByVal varV2 As Variant, ByVal strMsg As String, ByVal blnEqual As Boolean, ByVal blnDisplay As Boolean)
        
        If blnEqual Then
            strMsg = strMsg + vbNewLine + "Value " + CStr(varV1) + " is not equal to " + CStr(varV2)
        Else
            strMsg = strMsg + vbNewLine + "Value " + CStr(varV1) + " is not not equal to " + CStr(varV2)
        End If
        
        If blnDisplay Then DisplayAssert strMsg
        LogTest Warning, strMsg
    
End Sub


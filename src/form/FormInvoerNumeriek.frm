VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormInvoerNumeriek 
   ClientHeight    =   1183
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   3199
   OleObjectBlob   =   "FormInvoerNumeriek.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormInvoerNumeriek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Range As String

Private Sub cmdCancel_Click()
    
    Me.txtWaarde = vbNullString
    Me.Hide

End Sub

Private Sub cmdOk_Click()
    
    If Not m_Range = vbNullString Then ModRange.SetRangeValue m_Range, Val(txtWaarde.Value)
    Me.Hide

End Sub

Private Sub txtWaarde_KeyPress(ByVal intKeyAscii As MSForms.ReturnInteger)

    intKeyAscii = ModUtils.CorrectNumberAscii(intKeyAscii)

End Sub

Private Sub UserForm_Activate()

    Me.txtWaarde.SetFocus
    Me.Caption = ModConst.CONST_APPLICATION_NAME

End Sub

Public Sub SetValue(ByVal strRange As String, ByVal strItem As String, ByVal varValue As Variant, ByVal strUnit As String)
    
    Dim strError As String

    If varValue = vbNullString Then varValue = 0
    If Not IsNumeric(varValue) Then GoTo SetValueError
    
    m_Range = strRange
    
    lblParameter.Caption = strItem
    txtWaarde.Value = Val(varValue)
    lblEenheid.Caption = strUnit
    
    Exit Sub
    
SetValueError:

    strError = varValue & " is geen numerieke waarde" & vbNewLine
    strError = strError & ModConst.CONST_DEFAULTERROR_MSG
    ModMessage.ShowMsgBoxError strError
    
    Me.Hide

End Sub

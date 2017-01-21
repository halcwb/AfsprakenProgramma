VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormInvoerNumeriek 
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
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
Private m_Validate As String

Private Function Validate() As Boolean

    Dim strMsg As String
    
    Select Case m_Validate
        Case "Gewicht"
            strMsg = IIf(Not ModPatient.ValidWeightKg(val(txtWaarde.value)), "Geen geldig gewicht", vbNullString)
        Case "Lengte"
            strMsg = IIf(Not ModPatient.ValidLengthCm(val(txtWaarde.value)), "Geen geldige lengte", vbNullString)
        Case Else
            strMsg = vbNullString
    End Select
    
    lblValid.Caption = strMsg
    Validate = strMsg = vbNullString

End Function

Private Sub cmdCancel_Click()
    
    Me.txtWaarde = vbNullString
    Me.Hide

End Sub

Private Sub cmdClear_Click()

    txtWaarde.value = vbNullString

End Sub

Private Sub cmdOK_Click()
    
    If Not m_Range = vbNullString Then ModRange.SetRangeValue m_Range, val(txtWaarde.value)
    Me.Hide

End Sub

Private Sub txtWaarde_Change()

    cmdOK.Enabled = Validate
    
End Sub

Private Sub txtWaarde_KeyPress(ByVal intKeyAscii As MSForms.ReturnInteger)

    intKeyAscii = ModUtils.CorrectNumberAscii(intKeyAscii)

End Sub

Private Sub UserForm_Activate()

    Me.Caption = ModConst.CONST_APPLICATION_NAME
    Me.txtWaarde.SetFocus
    Me.txtWaarde.SelStart = 0
    Me.txtWaarde.SelLength = Len(Me.txtWaarde.value)

End Sub

Public Sub SetValue(ByVal strRange As String, ByVal strItem As String, ByVal varValue As Variant, ByVal strUnit As String, ByVal strValidate As String)
    
    Dim strError As String

    If varValue = vbNullString Then varValue = 0
    If Not IsNumeric(varValue) Then GoTo SetValueError
    
    m_Range = strRange
    m_Validate = strValidate
    
    lblParameter.Caption = strItem
    txtWaarde.value = val(varValue)
    lblEenheid.Caption = strUnit
    
    Exit Sub
    
SetValueError:

    strError = varValue & " is geen numerieke waarde" & vbNewLine
    ModMessage.ShowMsgBoxError strError
    
    Me.Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    cmdCancel_Click
    
End Sub

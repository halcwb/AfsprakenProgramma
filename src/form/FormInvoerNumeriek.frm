VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormInvoerNumeriek 
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
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
Private m_IsSST1 As Boolean
Private m_Extra As Double

Private Const constSST1Vol As String = "_Ped_TPN_SST1Vol"

Public Sub SetIsSST1()

    m_IsSST1 = True

End Sub

Private Function Validate() As Boolean

    Dim strMsg As String
    
    Select Case m_Validate
        Case "Gewicht"
            strMsg = IIf(Not ModPatient.ValidWeightKg(StringToDouble(txtWaarde.Value)), "Geen geldig gewicht", vbNullString)
        Case "Lengte"
            strMsg = IIf(Not ModPatient.ValidLengthCm(StringToDouble(txtWaarde.Value)), "Geen geldige lengte", vbNullString)
        Case Else
            strMsg = vbNullString
    End Select
    
    If StringToDouble(txtWaarde.Value) < 0 Then strMsg = "Kan geen negatieve waarde invoeren"
    
    lblValid.Caption = strMsg
    Validate = strMsg = vbNullString

End Function

Private Sub cmdCancel_Click()
    
    Me.txtWaarde = vbNullString
    Me.Hide

End Sub

Private Sub cmdClear_Click()

    txtWaarde.Value = vbNullString

End Sub

Private Sub cmdOK_Click()

    Dim dblValue As Double
    
    dblValue = StringToDouble(txtWaarde.Value) - m_Extra
    dblValue = IIf(dblValue < 0, 0, dblValue)
    If Not m_Range = vbNullString Then
        If m_IsSST1 Then
            ModRange.SetRangeValue constSST1Vol, dblValue
        Else
            ModRange.SetRangeValue m_Range, dblValue
        End If
    End If
    Me.Hide

End Sub

Private Sub txtWaarde_Change()

    cmdOK.Enabled = Validate
    
End Sub

Private Sub txtWaarde_KeyPress(ByVal intKeyAscii As MSForms.ReturnInteger)

    intKeyAscii = ModUtils.CorrectNumberAscii(intKeyAscii)

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm
    
    Me.Caption = ModConst.CONST_APPLICATION_NAME
    Me.txtWaarde.SetFocus
    Me.txtWaarde.SelStart = 0
    Me.txtWaarde.SelLength = Len(Me.txtWaarde.Value)

End Sub

Public Sub SetValue(ByVal strRange As String, ByVal strItem As String, ByVal varValue As Variant, ByVal strUnit As String, ByVal strValidate As String, Optional ByVal dblExtra As Double = 0)
    
    Dim strError As String
    
    m_Extra = dblExtra

    If varValue = vbNullString Then varValue = 0
    If Not IsNumeric(varValue) Then GoTo SetValueError
    
    m_Range = strRange
    m_Validate = strValidate
    
    lblParameter.Caption = strItem
    txtWaarde.Value = StringToDouble(varValue)
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

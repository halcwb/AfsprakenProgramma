VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPRN 
   Caption         =   "PRN"
   ClientHeight    =   1740
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   6435
   OleObjectBlob   =   "FormPRN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_intNo As Integer
Private Const constPRNCheck As String = "_Glob_MedDisc_PRN_"
Private Const constPRNText As String = "_Glob_MedDisc_PRNText_"

Private Sub Validate()

    Dim strValid As String
    
    txtPRN.Visible = chkPRN.Value
    strValid = IIf(txtPRN.Visible And txtPRN.Value = vbNullString, "Vul tekst in", vbNullString)
    
    lblValid.Caption = strValid
    cmdOK.Enabled = strValid = vbNullString

End Sub

Public Sub SetMedicamentNo(ByVal intN As Integer)

    m_intNo = intN

End Sub

Private Function RangeName(ByVal strRange As String) As String

    RangeName = strRange & IIf(m_intNo < 10, "0" & m_intNo, m_intNo)

End Function

Private Sub chkPRN_Click()

    Validate

End Sub

Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdClear_Click()

    ModRange.SetRangeValue RangeName(constPRNCheck), False
    ModRange.SetRangeValue RangeName(constPRNText), vbNullString
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    ModRange.SetRangeValue RangeName(constPRNCheck), chkPRN.Value
    ModRange.SetRangeValue RangeName(constPRNText), IIf(txtPRN.Visible, txtPRN.Text, vbNullString)
    Me.Hide

End Sub

Private Sub txtPRN_Change()

    Validate

End Sub

Private Sub CenterForm()

    StartUpPosition = 0
    Left = Application.Left + (0.5 * Application.Width) - (0.5 * Width)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)

End Sub

Private Sub UserForm_Activate()

    Dim blnPrn As Boolean
    
    CenterForm
    
    blnPrn = ModRange.GetRangeValue(RangeName(constPRNCheck), False)
    chkPRN.Value = blnPrn
    txtPRN.Text = ModRange.GetRangeValue(RangeName(constPRNText), vbNullString)

End Sub

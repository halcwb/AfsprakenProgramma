VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSpecialeVoeding 
   Caption         =   "Ingredienten Speciale Voeding ..."
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   OleObjectBlob   =   "FormSpecialeVoeding.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSpecialeVoeding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const constSpecVoed = "_Glob_SpecialeVoeding_"

'energy
'eiwit
'KH
'vet
'Na
'K
'Ca
'P
'Mg
'Fe
'VitD
'Cl


Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOK_Click()

    ModRange.SetRangeValue constSpecVoed & "01", txtCalorieen.Text
    ModRange.SetRangeValue constSpecVoed & "02", txtEiwit.Text
    ModRange.SetRangeValue constSpecVoed & "03", txtKoolHydraten.Text
    ModRange.SetRangeValue constSpecVoed & "04", txtVet.Text
    ModRange.SetRangeValue constSpecVoed & "05", txtNatrium.Text
    ModRange.SetRangeValue constSpecVoed & "06", txtKalium.Text
    ModRange.SetRangeValue constSpecVoed & "07", txtCalcium.Text
    ModRange.SetRangeValue constSpecVoed & "08", txtPhosfaat.Text
    ModRange.SetRangeValue constSpecVoed & "09", txtMagnesium.Text

    Me.Hide

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm

    txtCalorieen.Text = ModRange.GetRangeValue(constSpecVoed & "01", vbNullString)
    txtEiwit.Text = ModRange.GetRangeValue(constSpecVoed & "02", vbNullString)
    txtKoolHydraten.Text = ModRange.GetRangeValue(constSpecVoed & "03", vbNullString)
    txtVet.Text = ModRange.GetRangeValue(constSpecVoed & "04", vbNullString)
    txtNatrium.Text = ModRange.GetRangeValue(constSpecVoed & "05", vbNullString)
    txtKalium.Text = ModRange.GetRangeValue(constSpecVoed & "06", vbNullString)
    txtCalcium.Text = ModRange.GetRangeValue(constSpecVoed & "07", vbNullString)
    txtPhosfaat.Text = ModRange.GetRangeValue(constSpecVoed & "08", vbNullString)
    txtMagnesium.Text = ModRange.GetRangeValue(constSpecVoed & "09", vbNullString)

End Sub


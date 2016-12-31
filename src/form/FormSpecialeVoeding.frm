VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSpecialeVoeding 
   Caption         =   "Ingredienten Speciale Voeding ..."
   ClientHeight    =   5222
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   3605
   OleObjectBlob   =   "FormSpecialeVoeding.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSpecialeVoeding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOk_Click()

    ModRange.SetRangeValue "SpecVoed_1", txtCalorieen.Text
    ModRange.SetRangeValue "SpecVoed_2", txtEiwit.Text
    ModRange.SetRangeValue "SpecVoed_3", txtKoolHydraten.Text
    ModRange.SetRangeValue "SpecVoed_4", txtVet.Text
    ModRange.SetRangeValue "SpecVoed_5", txtNatrium.Text
    ModRange.SetRangeValue "SpecVoed_6", txtKalium.Text
    ModRange.SetRangeValue "SpecVoed_7", txtCalcium.Text
    ModRange.SetRangeValue "SpecVoed_8", txtPhosfaat.Text
    ModRange.SetRangeValue "SpecVoed_9", txtMagnesium.Text

    Me.Hide

End Sub

Private Sub UserForm_Activate()

    txtCalorieen.Text = ModRange.GetRangeValue("SpecVoed_1", vbNullString)
    txtEiwit.Text = ModRange.GetRangeValue("SpecVoed_2", vbNullString)
    txtKoolHydraten.Text = ModRange.GetRangeValue("SpecVoed_3", vbNullString)
    txtVet.Text = ModRange.GetRangeValue("SpecVoed_4", vbNullString)
    txtNatrium.Text = ModRange.GetRangeValue("SpecVoed_5", vbNullString)
    txtKalium.Text = ModRange.GetRangeValue("SpecVoed_6", vbNullString)
    txtCalcium.Text = ModRange.GetRangeValue("SpecVoed_7", vbNullString)
    txtPhosfaat.Text = ModRange.GetRangeValue("SpecVoed_8", vbNullString)
    txtMagnesium.Text = ModRange.GetRangeValue("SpecVoed_9", vbNullString)

End Sub

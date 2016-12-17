VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpecialeVoeding 
   Caption         =   "Ingredienten Speciale Voeding ..."
   ClientHeight    =   5222
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   3605
   OleObjectBlob   =   "frmSpecialeVoeding.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpecialeVoeding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













Option Explicit


Private Sub cmdCancel_Click()

    frmSpecialeVoeding.Hide

End Sub

Private Sub cmdOk_Click()

    Range("SpecVoed_1").Formula = txtCalorieen.Text
    Range("SpecVoed_2").Formula = txtEiwit.Text
    Range("SpecVoed_3").Formula = txtKoolHydraten.Text
    Range("SpecVoed_4").Formula = txtVet.Text
    Range("SpecVoed_5").Formula = txtNatrium.Text
    Range("SpecVoed_6").Formula = txtKalium.Text
    Range("SpecVoed_7").Formula = txtCalcium.Text
    Range("SpecVoed_8").Formula = txtPhosfaat.Text
    Range("SpecVoed_9").Formula = txtMagnesium.Text

    frmSpecialeVoeding.Hide

End Sub

Private Sub UserForm_Activate()

    txtCalorieen.Text = Range("SpecVoed_1").Formula
    txtEiwit.Text = Range("SpecVoed_2").Formula
    txtKoolHydraten.Text = Range("SpecVoed_3").Formula
    txtVet.Text = Range("SpecVoed_4").Formula
    txtNatrium.Text = Range("SpecVoed_5").Formula
    txtKalium.Text = Range("SpecVoed_6").Formula
    txtCalcium.Text = Range("SpecVoed_7").Formula
    txtPhosfaat.Text = Range("SpecVoed_8").Formula
    txtMagnesium.Text = Range("SpecVoed_9").Formula

End Sub


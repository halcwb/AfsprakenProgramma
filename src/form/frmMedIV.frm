VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMedIV 
   Caption         =   "Geef de naam van het medicament ..."
   ClientHeight    =   1624
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   5824
   OleObjectBlob   =   "frmMedIV.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMedIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Option Explicit

Private Sub cmdCancel_Click()
    txtMedicament.Text = vbNullString
    txtSterkte.Text = vbNullString
    frmMedIV.Hide
End Sub

Private Sub cmdOk_Click()
    frmMedIV.Hide
End Sub

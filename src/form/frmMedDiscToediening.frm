VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMedDiscToediening 
   Caption         =   "Edit toediening ..."
   ClientHeight    =   1218
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   2975
   OleObjectBlob   =   "frmMedDiscToediening.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMedDiscToediening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Private Sub cmdCancel_Click()
txtMedDiscToediening.Text = vbNullString
frmMedDiscToediening.Hide
End Sub

Private Sub cmdOk_Click()
frmMedDiscToediening.Hide
End Sub


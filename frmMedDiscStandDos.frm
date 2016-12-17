VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMedDiscStandDos 
   Caption         =   "Edit standaard dosis"
   ClientHeight    =   889
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   2856
   OleObjectBlob   =   "frmMedDiscStandDos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMedDiscStandDos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Private Sub cmdCancel_Click()
txtStandDos.Text = vbNullString
frmMedDiscStandDos.Hide
End Sub

Private Sub cmdOk_Click()
frmMedDiscStandDos.Hide
End Sub


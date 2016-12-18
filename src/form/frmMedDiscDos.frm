VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMedDiscDos 
   Caption         =   "Edit doserings advies ..."
   ClientHeight    =   3227
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   7861
   OleObjectBlob   =   "frmMedDiscDos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMedDiscDos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Private Sub cmdCancel_Click()
txtMedDiscDos.Text = vbNullString
frmMedDiscDos.Hide
End Sub

Private Sub cmdOk_Click()
frmMedDiscDos.Hide
End Sub

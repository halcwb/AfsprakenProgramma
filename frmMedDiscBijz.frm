VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMedDiscBijz 
   Caption         =   "Edit bijzonderheden ..."
   ClientHeight    =   3374
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   7980
   OleObjectBlob   =   "frmMedDiscBijz.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMedDiscBijz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Private Sub cmdCancel_Click()
txtMedDiscBijz.Text = vbNullString
frmMedDiscBijz.Hide
End Sub

Private Sub cmdOk_Click()
frmMedDiscBijz.Hide
End Sub

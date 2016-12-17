VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpmerking 
   Caption         =   "Opmerkingen ..."
   ClientHeight    =   1904
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   8211.001
   OleObjectBlob   =   "frmOpmerking.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOpmerking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















Private Sub cmdCancel_Click()

txtOpmerking.Text = "Cancel"
frmOpmerking.Hide

End Sub

Private Sub cmdOk_Click()

frmOpmerking.Hide

End Sub

Private Sub UserForm_Activate()

txtOpmerking.SetFocus

End Sub


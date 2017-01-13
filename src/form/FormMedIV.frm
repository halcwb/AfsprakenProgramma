VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedIV 
   Caption         =   "Geef de naam van het medicament ..."
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   OleObjectBlob   =   "FormMedIV.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMedIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    
    txtMedicament.Text = vbNullString
    txtSterkte.Text = vbNullString
    Me.Hide

End Sub

Private Sub cmdOK_Click()
    
    Me.Hide

End Sub


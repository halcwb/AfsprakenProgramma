VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPassword 
   Caption         =   "Voer paswoord in"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5085
   OleObjectBlob   =   "FormPassword.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Validate()

    Dim strValidate As String
    
    strValidate = IIf(txtPassword.value = vbNullString, "Voer paswoord in", vbNullString)
    
    cmdOK.Enabled = strValidate = vbNullString
    lblValid.Caption = strValidate

End Sub

Private Sub cmdCancel_Click()

    lblValid.Caption = "Cancel"
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Hide

End Sub

Private Sub txtPassword_Change()

    Validate

End Sub

Private Sub UserForm_Initialize()
    
    Validate

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    cmdCancel_Click

End Sub

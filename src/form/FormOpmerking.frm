VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOpmerking 
   Caption         =   "Opmerkingen ..."
   ClientHeight    =   1904
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   8211
   OleObjectBlob   =   "FormOpmerking.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOpmerking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()

    txtOpmerking.Text = "Cancel"
    Me.Hide

End Sub

Private Sub cmdOk_Click()

    Me.Hide

End Sub

Private Sub UserForm_Activate()

    txtOpmerking.SetFocus

End Sub

Public Sub SetText(strText As String)

    txtOpmerking.Text = IIf(strText = "0", vbNullString, strText)

End Sub


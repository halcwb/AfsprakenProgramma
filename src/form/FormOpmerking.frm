VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOpmerking 
   Caption         =   "Opmerkingen ..."
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   OleObjectBlob   =   "FormOpmerking.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOpmerking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CloseForm()

    Me.Hide

End Sub

Private Sub cmdCancel_Click()

    txtOpmerking.Text = "Cancel"
    CloseForm

End Sub

Private Sub cmdClear_Click()

    txtOpmerking.Value = vbNullString

End Sub

Private Sub cmdOK_Click()

    CloseForm

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm
    
    txtOpmerking.SetFocus

End Sub

Public Sub SetText(ByVal strText As String)

    txtOpmerking.Text = IIf(Trim(strText) = "0", vbNullString, strText)

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    CloseForm

End Sub

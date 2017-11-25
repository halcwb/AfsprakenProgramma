VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormTekstInvoer 
   Caption         =   "UserForm1"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285.001
   OleObjectBlob   =   "FormTekstInvoer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormTekstInvoer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgIsOK As Boolean

Public Property Get IsOK() As Boolean

    IsOK = mflgIsOK

End Property

Public Property Get Tekst() As String

    Tekst = txtTekst

End Property

Public Property Let Tekst(ByVal strText As String)

    txtTekst = VBA.Trim$(strText)

End Property

Private Sub cmdCancel_Click()

    mflgIsOK = False
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    mflgIsOK = True
    Me.Hide

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    Me.Hide
    
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormTekstInvoer 
   Caption         =   "UserForm1"
   ClientHeight    =   1561
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   8463
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

Public Property Let Tekst(TekstString As String)

    txtTekst = VBA.Trim$(TekstString)

End Property

Private Sub cmdCancel_Click()

    mflgIsOK = False
    Me.Hide

End Sub

Private Sub cmdOk_Click()

    mflgIsOK = True
    Me.Hide

End Sub


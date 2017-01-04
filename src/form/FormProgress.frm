VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProgress 
   ClientHeight    =   1099
   ClientLeft      =   21
   ClientTop       =   322
   ClientWidth     =   6384
   OleObjectBlob   =   "FormProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intWidth As Integer
Private intFormW As Integer
Private intFormH As Integer

Private Sub UserForm_Deactivate()

    Me.lblProgress.Width = intWidth

End Sub

Private Sub UserForm_Initialize()

    intWidth = Me.lblProgress.Width
    intFormW = Me.Width
    intFormH = Me.Height

End Sub

Private Sub UserForm_QueryClose(blnCancel As Integer, CloseMode As Integer)

    blnCancel = True

End Sub

Public Sub SetCaption(strTitle As String)

    Me.Caption = strTitle

End Sub

Public Sub SetJobPercentage(strJob As String, intPerc As Integer)
    
    Me.frmProgress.Caption = strJob & "..." & intPerc & "%"
    Me.lblProgress.Width = Int((CDbl(intWidth) / 100) * intPerc)
    Me.Repaint

End Sub

Private Sub UserForm_Resize()

    Me.Width = intFormW
    Me.Height = intFormH

End Sub
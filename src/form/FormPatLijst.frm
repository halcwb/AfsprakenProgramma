VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPatLijst 
   Caption         =   "Kies een patient ..."
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   OleObjectBlob   =   "FormPatLijst.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPatLijst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_pats As Collection

Private Sub lstPatienten_DblClick(ByVal blnCancel As MSForms.ReturnBoolean)
    
    Me.SetSelectedBed
    Me.Hide

End Sub

Public Sub LoadPatients(ByVal colPats As Collection)

    Dim objPat As ClassPatientInfo
    Set m_pats = colPats
    
    For Each objPat In colPats
        Me.lstPatienten.AddItem objPat.toString
    Next objPat

End Sub

Public Sub SetSelectedBed()
    
    Dim objPat As ClassPatientInfo
    Dim strBed As String

    If Me.lstPatienten.ListIndex > -1 Then
        Set objPat = m_pats(Me.lstPatienten.ListIndex + 1)
        strBed = objPat.Bed
    End If
    
    ModBed.SetBed strBed

End Sub

Private Sub CenterForm()

    StartUpPosition = 0
    Left = Application.Left + (0.5 * Application.Width) - (0.5 * Width)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm

End Sub

Private Sub UserForm_Terminate()
    
    Me.SetSelectedBed

End Sub

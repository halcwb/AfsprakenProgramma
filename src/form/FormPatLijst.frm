VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPatLijst 
   Caption         =   "Kies een patient ..."
   ClientHeight    =   4669
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   4025
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

Private Sub lstPatienten_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Me.Hide

End Sub

Public Sub LoadPatients(ByVal colPats As Collection)

    Dim objPat As ClassPatientInfo
    Set m_pats = colPats
    
    For Each objPat In colPats
        Me.lstPatienten.AddItem objPat.ToString
    Next objPat

End Sub

Public Function GetSelectedBed() As String
    
    Dim objPat As ClassPatientInfo
    Dim strBed As String

    If Me.lstPatienten.ListIndex > -1 Then
        Set objPat = m_pats(Me.lstPatienten.ListIndex)
        strBed = objPat.Bed
    End If
    
    GetSelectedBed = strBed

End Function

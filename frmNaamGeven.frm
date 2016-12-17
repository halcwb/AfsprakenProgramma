VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNaamGeven 
   Caption         =   "Naam geven"
   ClientHeight    =   1799
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   4326
   OleObjectBlob   =   "frmNaamGeven.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNaamGeven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    frmNaamGeven.Hide
    
    txtNaam.Text = vbNullString
    txtStart.Text = vbNullString
    RefNaam.Text = vbNullString

End Sub

Private Sub cmdOk_Click()
    Dim sRef As String, sNaam As String, intAant As Integer, i As Integer, Start As Integer
    
    RefNaam.SetFocus
    Range(RefNaam.Text).Select
    
    sNaam = txtNaam.Text
    Start = txtStart.Text
    
    With Selection
        intAant = .Rows.Count
        For i = 1 To intAant
            .Cells(i, 1).Name = sNaam & "_" & i + Start - 1
        Next i
    End With
    
    txtNaam.Text = vbNullString
    txtStart.Text = vbNullString
    RefNaam.Text = vbNullString
    
End Sub

Private Sub UserForm_Activate()
    
    txtNaam.Text = vbNullString
    txtStart.Text = vbNullString
    RefNaam.Text = vbNullString

End Sub


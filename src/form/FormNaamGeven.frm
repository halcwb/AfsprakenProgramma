VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormNaamGeven 
   Caption         =   "Naam geven"
   ClientHeight    =   1799
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   6972
   OleObjectBlob   =   "FormNaamGeven.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormNaamGeven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    txtNaam.Text = vbNullString
    txtStart.Text = vbNullString
    RefNaam.Text = vbNullString
    
    Me.Hide

End Sub

Private Sub cmdOk_Click()

    Dim strRes As String
    Dim strName As String
    Dim strGroup As String
    Dim intRows As Integer
    Dim intMax As Integer
    Dim intN As Integer
    Dim intStart As Integer
    
    RefNaam.SetFocus
    Range(RefNaam.Text).Select
    
    strName = txtNaam.Text
    strGroup = txtGroup.Text
    intStart = CInt(txtStart.Text)
    
    With Selection
        intRows = .Rows.Count
        intMax = intStart + intRows - 1
        For intN = 1 To intRows
            strRes = ModRange.CreateName(strName, strGroup, intN + intStart - 1, intMax)
            ModRange.SetNameToRange strRes, .Cells(intN, 1)
        Next intN
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


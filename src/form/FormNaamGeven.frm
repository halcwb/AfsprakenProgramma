VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormNaamGeven 
   Caption         =   "Naam geven"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
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
    txtGroup.Text = vbNullString
    RefNaam.Text = vbNullString
    
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Dim strRes As String
    Dim strName As String
    Dim strGroup As String
    Dim intRows As Integer
    Dim intMax As Integer
    Dim intN As Integer
    Dim intStart As Integer
    
    RefNaam.SetFocus
    
    If RefNaam.Text = vbNullString Then Exit Sub
    
    Range(RefNaam.Text).Select
    
    strName = txtNaam.Text
    strGroup = txtGroup.Text
    
    If txtStart.Text = vbNullString Then
        intStart = 0
    Else
        intStart = CInt(txtStart.Text)
    End If
    
    If strName = vbNullString Or strGroup = vbNullString Then Exit Sub
    
    With Selection
        intRows = .Rows.count
        If intRows = 1 Then
            strRes = IIf(chkIsData.value, "_" & strGroup & "_" & strName, strGroup & "_" & strName)
            ModRange.SetNameToRange strRes, .Cells(1, 1)
        Else
            intMax = intStart + intRows - 1
            For intN = 1 To intRows
                strRes = ModRange.CreateName(strName, strGroup, intN + intStart - 1, intMax, chkIsData.value)
                ModRange.SetNameToRange strRes, .Cells(intN, 1)
            Next intN
        End If
    End With
    
    txtNaam.Text = vbNullString
    txtStart.Text = vbNullString
    RefNaam.Text = vbNullString
    
    Me.Hide
    
End Sub

Private Sub UserForm_Activate()
    
    txtNaam.Text = vbNullString
    txtStart.Text = vbNullString
    RefNaam.Text = vbNullString

End Sub


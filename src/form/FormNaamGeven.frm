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
    Dim intCells As Integer
    Dim intMax As Integer
    Dim intN As Integer
    Dim intStart As Integer
    Dim strSelect As String
    Dim varCell As Variant
    Dim objCell As Range
    
    RefNaam.SetFocus
    
    If RefNaam.Text = vbNullString Then Exit Sub
    
    strSelect = RefNaam.Text
    strSelect = Replace(strSelect, ";", ",")
    Range(strSelect).Select
    
    strName = txtNaam.Text
    strGroup = txtGroup.Text
    
    If txtStart.Text = vbNullString Then
        intStart = 0
    Else
        intStart = CInt(txtStart.Text)
    End If
    
    If strName = vbNullString Or strGroup = vbNullString Then Exit Sub
    
    intCells = Selection.Cells.Count
    
    If intCells = 1 Then
        strRes = IIf(chkIsData.Value, "_" & strGroup & "_" & strName, strGroup & "_" & strName)
        ModRange.SetNameToRange strRes, Selection.Cells(1, 1)
    Else
        intMax = intStart + intCells - 1
        intN = 1
        For Each varCell In Selection.Cells
            strRes = ModRange.CreateName(strName, strGroup, intN + intStart - 1, intMax, chkIsData.Value)
            Set objCell = varCell
            ModRange.SetNameToRange strRes, objCell
            intN = intN + 1
        Next varCell
    End If
    
'    With Selection
'        intCells = .Cells.Count
'        If intCells = 1 Then
'            strRes = IIf(chkIsData.Value, "_" & strGroup & "_" & strName, strGroup & "_" & strName)
'            ModRange.SetNameToRange strRes, .Cells(1, 1)
'        Else
'            intMax = intStart + intCells - 1
'            For intN = 1 To intCells
'                strRes = ModRange.CreateName(strName, strGroup, intN + intStart - 1, intMax, chkIsData.Value)
'                ModRange.SetNameToRange strRes, .Cells(intN, 1)
'            Next intN
'        End If
'    End With
    
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


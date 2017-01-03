Attribute VB_Name = "ModNeoLab"
Option Explicit

Private Const constNeoLab = "_Neo_Lab_Opm"

Public Sub NeoLab_EnterText()
    
    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(constNeoLab, vbNullString)
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constNeoLab, frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing
    
End Sub

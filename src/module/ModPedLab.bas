Attribute VB_Name = "ModPedLab"
Option Explicit

Private Const constLabOpm = "_Ped_Lab_Opm"

Private Sub EnterText(strRange As String)

    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(strRange, vbNullString)
    
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue strRange, frmOpmerking.txtOpmerking.Text
    End If
    
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub PedLab_EnterText()
    
    EnterText constLabOpm
    
End Sub

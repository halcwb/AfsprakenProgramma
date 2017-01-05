Attribute VB_Name = "ModPedLab"
Option Explicit

Private Const constPedLab As String = "_Ped_Lab_"
Private Const constPedLabOpm As String = "_Ped_Lab_Opm"

Public Sub PedLab_Clear()
    
    ModProgress.StartProgress "Verwijder Ped Lab"
    ModPatient.ClearPatientData constPedLab, False, True
    ModProgress.FinishProgress

End Sub

Private Sub EnterText(ByVal strRange As String)

    Dim frmOpmerking As FormOpmerking
    
    Set frmOpmerking = New FormOpmerking
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(strRange, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue strRange, frmOpmerking.txtOpmerking.Text
    End If
    
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub PedLab_EnterText()
    
    EnterText constPedLabOpm
    
End Sub

Attribute VB_Name = "ModNeoLab"
Option Explicit

Private Const constNeoLab = "_Neo_Lab_"
Private Const constNeoLabOpm = "_Neo_Lab_Opm"

Public Sub NeoLab_Clear()

    Dim intN, intC As Integer
    Dim strRange As String
    
    ModProgress.StartProgress "Verwijder Neo Lab"
    ModPatient.ClearPatientData constNeoLab, False, True
    ModProgress.FinishProgress

End Sub

Public Sub NeoLab_EnterText()
    
    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(constNeoLabOpm, vbNullString)
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constNeoLabOpm, frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing
    
End Sub

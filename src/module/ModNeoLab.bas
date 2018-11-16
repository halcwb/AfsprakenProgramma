Attribute VB_Name = "ModNeoLab"
Option Explicit

Private Const constNeoLab As String = "_Neo_Lab_"
Private Const constNeoLabOpm As String = "_Neo_Lab_Opm"

Public Sub NeoLab_Clear()

    ModProgress.StartProgress "Verwijder Neo Lab"
    ModPatient.Patient_ClearData constNeoLab, False, True
    ModProgress.FinishProgress

End Sub

Public Sub NeoLab_EnterText()
    
    Dim frmOpmerking As FormOpmerking
    
    Set frmOpmerking = New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(constNeoLabOpm, vbNullString)
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constNeoLabOpm, frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub

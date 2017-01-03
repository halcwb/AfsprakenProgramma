Attribute VB_Name = "ModPedLijnPM"
Option Explicit

Private Const constOpm = "_Ped_IVLijn_Opm"
Private Const constPMTbl = "tbl_Ped_PMStandaard"
Private Const constPMSet = "tbl_Ped_PMInstelling"

Private Sub EnterOpm()

    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(constOpm, vbNullString)
    
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constOpm, frmOpmerking.txtOpmerking.Text
    End If
    
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub PedLijnPM_EnterText()
    
    EnterOpm
    
End Sub

Public Sub PedLijnPM_PaceMaker()

    shtPedBerIVenPM.Range(constPMTbl).Copy
    shtPedBerIVenPM.Range(constPMSet).PasteSpecial xlPasteValues

End Sub


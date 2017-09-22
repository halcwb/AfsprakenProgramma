Attribute VB_Name = "ModPedLab"
Option Explicit

Private Const constPedLab As String = "_Ped_Lab_"
Private Const constPedLabOpm As String = "_Ped_Lab_Opm"

Private Const constLabOpnVerw As String = "C3"
Private Const constLab14Verw As String = "E3"
Private Const constLab19Verw As String = "G3"
Private Const constLab24Verw As String = "I3"
Private Const constLabDag1Verw As String = "L3"

Public Enum PedLabRondes
    PedLabRondeOpn = 1
    PedLabRonde14 = 14
    PedLabRonde19 = 19
    PedLabRonde24 = 24
    PedLabRondeDag1 = 6
End Enum

Private Sub SetLabRonde(ByVal strRonde As String, ByVal intC As Integer, ByVal blnValue As Boolean)

    Dim intN As Integer
    
    For intN = 1 To intC
        ModRange.SetRangeValue strRonde & "_" & IntNToStrN(intN), blnValue
    Next

End Sub

Private Sub ToggleLabRondes(ByVal enmRonde As PedLabRondes, ByVal blnValue As Boolean)

    Dim strRonde As String
    
    strRonde = IIf(enmRonde = PedLabRondeDag1, "Dag1", IIf(enmRonde = PedLabRondeOpn, "Opn", enmRonde))
    strRonde = constPedLab & strRonde
    
    Select Case enmRonde
        Case PedLabRondes.PedLabRondeOpn
            SetLabRonde strRonde, 32, blnValue
        Case PedLabRondes.PedLabRonde14
            SetLabRonde strRonde, 31, blnValue
        Case PedLabRondes.PedLabRonde19
            SetLabRonde strRonde, 31, blnValue
        Case PedLabRondes.PedLabRonde24
            SetLabRonde strRonde, 31, blnValue
        Case PedLabRondes.PedLabRondeDag1
            SetLabRonde strRonde, 31, blnValue
    End Select
    
End Sub

Public Sub PedLab_Toggle_Opn()

    ToggleLabRondes PedLabRondeOpn, shtPedBerLab.Range(constLabOpnVerw).Value = vbNullString

End Sub

Public Sub PedLab_Toggle_14()

    ToggleLabRondes PedLabRonde14, shtPedBerLab.Range(constLab14Verw).Value = vbNullString

End Sub

Public Sub PedLab_Toggle_19()

    ToggleLabRondes PedLabRonde19, shtPedBerLab.Range(constLab19Verw).Value = vbNullString

End Sub

Public Sub PedLab_Toggle_24()

    ToggleLabRondes PedLabRonde24, shtPedBerLab.Range(constLab24Verw).Value = vbNullString

End Sub

Public Sub PedLab_Toggle_Dag1()

    ToggleLabRondes PedLabRondeDag1, shtPedBerLab.Range(constLabDag1Verw).Value = vbNullString

End Sub

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

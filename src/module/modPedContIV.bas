Attribute VB_Name = "modPedContIV"
Option Explicit

Private Const constTblMed = "tblMedicationContIV"
Private Const constMedIVKeuze = "_Ped_MedIV_Keuze_"
Private Const constMedIVSterkte = "_Ped_MedIV_Sterkte_"
Private Const constMedIVOpm = "_Ped_MedIV_Opm"

Private Sub SetToStandard(intRegel As Integer)

    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim varOplossing As Variant
    Dim strStand As String
    Dim strRegel As String
    
    On Error GoTo SetToStandardError

    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    strMedicament = "_Ped_MedIV_Keuze_" & strRegel
    strMedSterkte = "_Ped_MedIV_Sterkte_" & strRegel
    strOplHoev = "_Ped_MedIV_OplVol_" & strRegel
    strOplossing = "_Ped_MedIV_OplVlst_" & strRegel
    strStand = "_Ped_MedIV_Stand_" & strRegel
    
    ModRange.SetRangeValue strMedSterkte, 0
    ModRange.SetRangeValue strOplHoev, 0
    ModRange.SetRangeValue strStand, 0
    
    If ModRange.GetRangeValue(strMedicament, 0) = 1 Then
        ModRange.SetRangeValue strOplossing, 1
    Else
        varOplossing = Application.VLookup(Range(constTblMed).Cells(Range(strMedicament).Value, 1), Range(constTblMed), 22, False)
        ModRange.SetRangeValue strOplossing, varOplossing
    End If
    
    Exit Sub
    
SetToStandardError:

    ModLog.LogError "SetMedContIVToStandard: " & " Error for regel " & strRegel

End Sub

Public Sub PedContIV_SetStandard_01()
    
    SetToStandard 1

End Sub

Public Sub PedContIV_SetStandard_02()
    
    SetToStandard 2

End Sub

Public Sub PedContIV_SetStandard_03()
    
    SetToStandard 3

End Sub

Public Sub PedContIV_SetStandard_04()
        
    SetToStandard 4

End Sub

Public Sub PedContIV_SetStandard_05()
    
    SetToStandard 5

End Sub

Public Sub PedContIV_SetStandard_06()
    
    SetToStandard 6

End Sub

Public Sub PedContIV_SetStandard_07()

    SetToStandard 7

End Sub

Public Sub PedContIV_SetStandard_08()

    SetToStandard 8

End Sub

Public Sub PedContIV_SetStandard_09()
    
    SetToStandard 9

End Sub

Public Sub PedContIV_SetStandard_10()
    
    SetToStandard 10

End Sub

Private Sub EnterNumeric(intRegel As Integer, strRange As String, strUnit As String, intColumn As Integer)

    Dim frmInvoer As New FormInvoerNumeriek
    Dim varKeuze As Variant
    Dim strRegel As String
    
    On Error GoTo OpenInvoerNumeriekError
    
    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    varKeuze = ModRange.GetRangeValue("_Ped_MedIV_Keuze_" & strRegel, vbNullString)
    
    With frmInvoer
        .Caption = "Medicament " & intRegel
        .lblParameter = "Oplossing"
        .lblEenheid = strUnit
        If ModRange.GetRangeValue("_Ped_MedIV_OplVol_" & strRegel, 0) = 0 Then
            .txtWaarde = Application.WorksheetFunction.Index(Range(constTblMed), varKeuze, 12)
        Else
            .txtWaarde = ModRange.GetRangeValue(strRange & strRegel, vbNullString)
        End If
        .Show
        If IsNumeric(.txtWaarde) Then
            If CDbl(.txtWaarde) = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), varKeuze, 12) Then
                ModRange.SetRangeValue strRange & strRegel, 0
            Else
                ModRange.SetRangeValue strRange & strRegel, .txtWaarde
            End If
        End If
    End With
    
    Set frmInvoer = Nothing
    
    Exit Sub
    
OpenInvoerNumeriekError:

    ModLog.LogError "EnterNumeric(" & Join(Array(strRegel, strRange, strUnit, intColumn), ", ") & ")"
    Set frmInvoer = Nothing

End Sub

Private Sub SetMedConc(intRegel As Integer)

    Dim strUnit As String
    Dim strRegel As String
    
    On Error GoTo SetMedConcError

    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    strUnit = Application.WorksheetFunction.Index(Range(constTblMed), Range("_Ped_MedIV_Keuze_" & strRegel), 4)
    EnterNumeric intRegel, "MedIVSterkte_", strUnit, 11
    
    Exit Sub
    
SetMedConcError:

    ModLog.LogError "SetMedConc(" & intRegel & ")"

End Sub

Public Sub PedContIV_MedConc_01()
    
    SetMedConc 1

End Sub

Public Sub PedContIV_MedConc_02()
    
    SetMedConc 2

End Sub

Public Sub PedContIV_MedConc_03()
    
    SetMedConc 3

End Sub

Public Sub PedContIV_MedConc_04()
    
    SetMedConc 4

End Sub

Public Sub PedContIV_MedConc_05()
    
    SetMedConc 5

End Sub

Public Sub PedContIV_MedConc_06()
    
    SetMedConc 6

End Sub

Public Sub PedContIV_MedConc_07()
    
    SetMedConc 7

End Sub

Public Sub PedContIV_MedConc_08()
    
    SetMedConc 8

End Sub

Public Sub PedContIV_MedConc_09()
    
    SetMedConc 9

End Sub

Public Sub PedContIV_MedConc_10()
    
    SetMedConc 10

End Sub

Private Sub SetSolution(intRegel As Integer)

    EnterNumeric intRegel, "_Ped_MedIV_OplVol_", "mL", 12

End Sub

Public Sub PedContIV_SetSolution_01()
    
    SetSolution 1

End Sub

Public Sub PedContIV_SetSolution_02()
    
    SetSolution 2

End Sub

Public Sub PedContIV_SetSolution_03()
    
    SetSolution 3

End Sub

Public Sub PedContIV_SetSolution_04()
    
    SetSolution 4

End Sub

Public Sub PedContIV_SetSolution_05()
    
    SetSolution 5

End Sub

Public Sub PedContIV_SetSolution_06()
    
    SetSolution 6

End Sub

Public Sub PedContIV_SetSolution_07()
    
    SetSolution 7

End Sub

Public Sub PedContIV_SetSolution_08()
    
    SetSolution 8

End Sub

Public Sub PedContIV_SetSolution_09()
    
    SetSolution 9

End Sub

Public Sub PedContIV_SetSolution_10()
    
    SetSolution 10

End Sub

Private Sub EnterMed(intN As Integer)

    Dim strMed As String, strSterkte As String
    Dim frmMedIV As New FormMedIV
    
    frmMedIV.Show
    
    strMed = frmMedIV.txtMedicament.Text
    strSterkte = frmMedIV.txtSterkte.Text
    ModRange.SetRangeValue constMedIVKeuze & intN, strMed
    ModRange.SetRangeValue constMedIVSterkte & intN, strSterkte
    
    Set frmMedIV = Nothing
        
End Sub

Public Sub PedContIV_EnterMed_11()

    EnterMed 11
        
End Sub

Public Sub PedContIV_EnterMed_12()
    
    EnterMed 12

End Sub

Public Sub PedContIV_EnterMed_13()
    
    EnterMed 13

End Sub

Public Sub PedContIV_EnterMed_14()
    
    EnterMed 14

End Sub

Public Sub PedContIV_EnterMed_15()
    
    EnterMed 15

End Sub

Public Sub PedContIV_Text()

    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.SetText ModRange.GetRangeValue(constMedIVOpm, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constMedIVOpm, frmOpmerking.txtOpmerking.Text
    End If
    
    Set frmOpmerking = Nothing

End Sub


Attribute VB_Name = "modPedContIV"
Option Explicit

Private Const constTblMed As String = "tblMedicationContIV"
Private Const constMedIVKeuze As String = "_Ped_MedIV_Keuze_"
Private Const constMedIVSterkte As String = "_Ped_MedIV_Sterkte_"
Private Const constMedIVOpm As String = "_Ped_MedIV_Opm"
Private Const constMedIVOplVol As String = "_Ped_MedIV_OplVol_"
Private Const constMedIVOplVlst As String = "_Ped_MedIV_OplVlst_"
Private Const constMedIVStand As String = "_Ped_MedIV_Stand_"

Private Sub Clear(ByVal intN As Integer)

    Dim strN As String
    
    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim strStand As String
    
    strN = IIf(intN < 10, "0" + intN, intN)
    
    strN = IIf(intN < 10, "0" & intN, intN)
    strMedicament = constMedIVKeuze & strN
    strMedSterkte = constMedIVSterkte & strN
    strOplHoev = constMedIVOplVol & strN
    strOplossing = constMedIVOplVlst & strN
    strStand = constMedIVStand & strN
    
    If intN < 16 Then
        ModRange.SetRangeValue strMedicament, 1
    Else
        ModRange.SetRangeValue strMedicament, vbNullString
    End If
    
    ModRange.SetRangeValue strMedSterkte, 0
    ModRange.SetRangeValue strOplHoev, 0
    ModRange.SetRangeValue strOplossing, 1
    ModRange.SetRangeValue strStand, 0

End Sub

Public Sub PedContIV_Clear_01()

    Clear 1

End Sub

Public Sub PedContIV_Clear_02()

    Clear 2

End Sub

Public Sub PedContIV_Clear_03()

    Clear 3

End Sub

Public Sub PedContIV_Clear_04()

    Clear 4

End Sub

Public Sub PedContIV_Clear_05()

    Clear 5

End Sub

Public Sub PedContIV_Clear_06()

    Clear 6

End Sub

Public Sub PedContIV_Clear_07()

    Clear 7

End Sub

Public Sub PedContIV_Clear_08()

    Clear 8

End Sub

Public Sub PedContIV_Clear_09()

    Clear 1

End Sub

Public Sub PedContIV_Clear_10()

    Clear 10

End Sub

Public Sub PedContIV_Clear_11()

    Clear 11

End Sub

Public Sub PedContIV_Clear_12()

    Clear 12

End Sub

Public Sub PedContIV_Clear_13()

    Clear 13

End Sub

Public Sub PedContIV_Clear_14()

    Clear 14

End Sub

Public Sub PedContIV_Clear_15()

    Clear 15

End Sub

Public Sub PedContIV_Clear_16()

    Clear 16

End Sub

Public Sub PedContIV_Clear_17()

    Clear 17

End Sub

Public Sub PedContIV_Clear_18()

    Clear 18

End Sub

Public Sub PedContIV_Clear_19()

    Clear 19

End Sub

Public Sub PedContIV_Clear_20()

    Clear 20

End Sub

Private Sub SetToStandard(ByVal intN As Integer)

    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim varOplossing As Variant
    Dim strStand As String
    Dim strN As String
    Dim intKeuze As Integer
    
    On Error GoTo SetToStandardError

    strN = IIf(intN < 10, "0" & intN, intN)
    strMedicament = constMedIVKeuze & strN
    strMedSterkte = constMedIVSterkte & strN
    strOplHoev = constMedIVOplVol & strN
    strOplossing = constMedIVOplVlst & strN
    strStand = constMedIVStand & strN
    
    ModRange.SetRangeValue strMedSterkte, 0
    ModRange.SetRangeValue strOplHoev, 0
    ModRange.SetRangeValue strStand, 0
    
    intKeuze = ModRange.GetRangeValue(strMedicament, 0)
    If intKeuze = 0 Then GoTo SetToStandardError    ' Something is wrong, 0 is no valid value
    
    If intKeuze = 1 Then                            ' No medicament was selected so clear the line
        Clear intN
    Else                                            ' Else find the right standard concentration
        varOplossing = Application.VLookup(Range(constTblMed).Cells(intKeuze, 1), Range(constTblMed), 22, False)
        ModRange.SetRangeValue strOplossing, varOplossing
    End If
    
    Exit Sub
    
SetToStandardError:

    ModLog.LogError "SetMedContIVToStandard: " & " Error for regel " & strN

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

Public Sub PedContIV_SetStandard_11()
    
    SetToStandard 11

End Sub

Public Sub PedContIV_SetStandard_12()
    
    SetToStandard 12

End Sub

Public Sub PedContIV_SetStandard_13()
    
    SetToStandard 13

End Sub

Public Sub PedContIV_SetStandard_14()
    
    SetToStandard 14

End Sub

Public Sub PedContIV_SetStandard_15()
    
    SetToStandard 15

End Sub

Private Sub EnterNumeric(ByVal intRegel As Integer, ByVal strRange As String, ByVal strUnit As String, ByVal intColumn As Integer)

    Dim frmInvoer As FormInvoerNumeriek
    Dim varKeuze As Variant
    Dim strRegel As String
    
    On Error GoTo OpenInvoerNumeriekError
    
    Set frmInvoer = New FormInvoerNumeriek
    
    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    varKeuze = ModRange.GetRangeValue(constMedIVKeuze & strRegel, vbNullString)
    
    With frmInvoer
        .Caption = "Medicament " & intRegel
        .lblParameter = "Oplossing"
        .lblEenheid = strUnit
        If ModRange.GetRangeValue(constMedIVOplVol & strRegel, 0) = 0 Then
            .txtWaarde = Application.WorksheetFunction.Index(Range(constTblMed), varKeuze, 12)
        Else
            .txtWaarde = ModRange.GetRangeValue(strRange & strRegel, vbNullString)
        End If
        .Show
        If IsNumeric(.txtWaarde) Then
            If CDbl(.txtWaarde) = Application.WorksheetFunction.Index(Range(constTblMed), varKeuze, 12) Then
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

Private Sub SetMedConc(ByVal intRegel As Integer)

    Dim strUnit As String
    Dim strRegel As String
    
    On Error GoTo SetMedConcError

    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    strUnit = Application.WorksheetFunction.Index(Range(constTblMed), Range(constMedIVKeuze & strRegel), 4)
    EnterNumeric intRegel, constMedIVSterkte, strUnit, 11
    
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

Public Sub PedContIV_MedConc_11()
    
    SetMedConc 11

End Sub

Public Sub PedContIV_MedConc_12()
    
    SetMedConc 12

End Sub

Public Sub PedContIV_MedConc_13()
    
    SetMedConc 13

End Sub

Public Sub PedContIV_MedConc_14()
    
    SetMedConc 14

End Sub

Public Sub PedContIV_MedConc_15()
    
    SetMedConc 15

End Sub

Private Sub SetSolution(ByVal intRegel As Integer)

    EnterNumeric intRegel, constMedIVOplVol, "mL", 12

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

Public Sub PedContIV_SetSolution_11()
    
    SetSolution 11

End Sub

Public Sub PedContIV_SetSolution_12()
    
    SetSolution 12

End Sub

Public Sub PedContIV_SetSolution_13()
    
    SetSolution 13

End Sub

Public Sub PedContIV_SetSolution_14()
    
    SetSolution 14

End Sub

Public Sub PedContIV_SetSolution_15()
    
    SetSolution 15

End Sub

Private Sub EnterMed(ByVal intN As Integer)

    Dim strMed As String
    Dim strSterkte As String
    Dim frmMedIV As FormMedIV
    
    Set frmMedIV = New FormMedIV
    frmMedIV.Show
    
    strMed = frmMedIV.txtMedicament.Text
    strSterkte = frmMedIV.txtSterkte.Text
    ModRange.SetRangeValue constMedIVKeuze & intN, strMed
    ModRange.SetRangeValue constMedIVSterkte & intN, strSterkte
    
    Set frmMedIV = Nothing
        
End Sub

Public Sub PedContIV_EnterMed_16()

    EnterMed 16
        
End Sub

Public Sub PedContIV_EnterMed_17()
    
    EnterMed 17

End Sub

Public Sub PedContIV_EnterMed_18()
    
    EnterMed 18

End Sub

Public Sub PedContIV_EnterMed_19()
    
    EnterMed 19

End Sub

Public Sub PedContIV_EnterMed_20()
    
    EnterMed 20

End Sub

Public Sub PedContIV_Text()

    Dim frmOpmerking As FormOpmerking
    
    Set frmOpmerking = New FormOpmerking
    
    frmOpmerking.SetText ModRange.GetRangeValue(constMedIVOpm, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constMedIVOpm, frmOpmerking.txtOpmerking.Text
    End If
    
    Set frmOpmerking = Nothing

End Sub


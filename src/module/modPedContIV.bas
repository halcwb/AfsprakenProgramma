Attribute VB_Name = "modPedContIV"
Option Explicit

Private Const constTblMed = "tblMedicationContIV"
Private Const constMedIVKeuze = "_Ped_MedIV_Keuze_"
Private Const constMedIVSterkte = "_Ped_MedIV_Sterkte_"
Private Const constMedIVOpm = "_Ped_MedIV_Opm"

Private Sub SetMedContIVToStandardItem(intRegel As Integer)

    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim varOplossing As Variant
    Dim strStand As String
    Dim strRegel As String
    
    On Error GoTo SetMedContIVToStandardError

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
    
SetMedContIVToStandardError:

    ModLog.LogError "SetMedContIVToStandard: " & " Error for regel " & strRegel

End Sub

Public Sub SetMedContIVToStand_01()
    
    SetMedContIVToStandardItem 1

End Sub

Public Sub SetMedContIVToStand_02()
    
    SetMedContIVToStandardItem 2

End Sub

Public Sub SetMedContIVToStand_03()
    
    SetMedContIVToStandardItem 3

End Sub

Public Sub SetMedContIVToStand_04()
        
    SetMedContIVToStandardItem 4

End Sub

Public Sub SetMedContIVToStand_05()
    
    SetMedContIVToStandardItem 5

End Sub

Public Sub SetMedContIVToStand_06()
    
    SetMedContIVToStandardItem 6

End Sub

Public Sub SetMedContIVToStand_07()

    SetMedContIVToStandardItem 7

End Sub

Public Sub SetMedContIVToStand_08()

    SetMedContIVToStandardItem 8

End Sub

Public Sub SetMedContIVToStand_09()
    
    SetMedContIVToStandardItem 9

End Sub

Public Sub SetMedContIVToStand_10()
    
    SetMedContIVToStandardItem 10

End Sub

Private Sub OpenInvoerNumeriek(intRegel As Integer, strRange As String, strUnit As String, intColumn As Integer)

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

    ModLog.LogError "OpenInvoerNumeriek(" & Join(Array(strRegel, strRange, strUnit, intColumn), ", ") & ")"
    Set frmInvoer = Nothing

End Sub

Private Sub SetMedContIVSterkteItem(intRegel As Integer)

    Dim strUnit As String
    Dim strRegel As String
    
    On Error GoTo SetMedContIVSterkteError

    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    strUnit = Application.WorksheetFunction.Index(Range(constTblMed), Range("_Ped_MedIV_Keuze_" & strRegel), 4)
    OpenInvoerNumeriek intRegel, "MedIVSterkte_", strUnit, 11
    
    Exit Sub
    
SetMedContIVSterkteError:

    ModLog.LogError "SetMedContIVSterkteItem(" & intRegel & ")"

End Sub

Public Sub SetMedContIVSterkte_01()
    
    SetMedContIVSterkteItem 1

End Sub

Public Sub SetMedContIVSterkte_02()
    
    SetMedContIVSterkteItem 2

End Sub

Public Sub SetMedContIVSterkte_03()
    
    SetMedContIVSterkteItem 3

End Sub

Public Sub SetMedContIVSterkte_04()
    
    SetMedContIVSterkteItem 4

End Sub

Public Sub SetMedContIVSterkte_05()
    
    SetMedContIVSterkteItem 5

End Sub

Public Sub SetMedContIVSterkte_06()
    
    SetMedContIVSterkteItem 6

End Sub

Public Sub SetMedContIVSterkte_07()
    
    SetMedContIVSterkteItem 7

End Sub

Public Sub SetMedContIVSterkte_08()
    
    SetMedContIVSterkteItem 8

End Sub

Public Sub SetMedContIVSterkte_09()
    
    SetMedContIVSterkteItem 9

End Sub

Public Sub SetMedContIVSterkte_10()
    
    SetMedContIVSterkteItem 10

End Sub

Private Sub SetMedContIVOplossingItem(intRegel As Integer)

    OpenInvoerNumeriek intRegel, "_Ped_MedIV_OplVol_", "mL", 12

End Sub

Public Sub SetMedContIVOplossing_01()
    
    SetMedContIVOplossingItem 1

End Sub

Public Sub SetMedContIVOplossing_02()
    
    SetMedContIVOplossingItem 2

End Sub

Public Sub SetMedContIVOplossing_03()
    
    SetMedContIVOplossingItem 3

End Sub

Public Sub SetMedContIVOplossing_04()
    
    SetMedContIVOplossingItem 4

End Sub

Public Sub SetMedContIVOplossing_05()
    
    SetMedContIVOplossingItem 5

End Sub

Public Sub SetMedContIVOplossing_06()
    
    SetMedContIVOplossingItem 6

End Sub

Public Sub SetMedContIVOplossing_07()
    
    SetMedContIVOplossingItem 7

End Sub

Public Sub SetMedContIVOplossing_08()
    
    SetMedContIVOplossingItem 8

End Sub

Public Sub SetMedContIVOplossing_09()
    
    SetMedContIVOplossingItem 9

End Sub

Public Sub SetMedContIVOplossing_10()
    
    SetMedContIVOplossingItem 10

End Sub

Private Sub MedIVInvoer(intN As Integer)

    Dim strMed As String, strSterkte As String
    Dim frmMedIV As New FormMedIV
    
    frmMedIV.Show
    
    strMed = frmMedIV.txtMedicament.Text
    strSterkte = frmMedIV.txtSterkte.Text
    ModRange.SetRangeValue constMedIVKeuze & intN, strMed
    ModRange.SetRangeValue constMedIVSterkte & intN, strSterkte
    
    Set frmMedIV = Nothing
        
End Sub

Public Sub MedIV_11()

    MedIVInvoer 11
        
End Sub

Public Sub MedIV_12()
    
    MedIVInvoer 12

End Sub

Public Sub MedIV_13()
    
    MedIVInvoer 13

End Sub

Public Sub MedIV_14()
    
    MedIVInvoer 14

End Sub

Public Sub MedIV_15()
    
    MedIVInvoer 15

End Sub

Public Sub PedContIV_Opm()

    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.SetText ModRange.GetRangeValue(constMedIVOpm, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constMedIVOpm, frmOpmerking.txtOpmerking.Text
    End If
    
    Set frmOpmerking = Nothing

End Sub


Attribute VB_Name = "modPedContIV"
Option Explicit

Private Const constTblMed As String = "Tbl_Ped_MedContIV"
Private Const constMedIVKeuze As String = "_Ped_MedIV_Keuze_"
Private Const constMedIVSterkte As String = "_Ped_MedIV_Sterkte_"
Private Const constMedIVOpm As String = "_Ped_MedIV_Opm"
Private Const constMedIVOplVol As String = "_Ped_MedIV_OplVol_"
Private Const constMedIVOplVlst As String = "_Ped_MedIV_OplVlst_"
Private Const constMedIVStand As String = "_Ped_MedIV_Stand_"
Private Const constMedIVCount As Integer = 15

Private Const constStandOplKeuze As Integer = 2
Private Const constStandOplVlst As Integer = 17
Private Const constStandHoevIndx As Integer = 20
Private Const constStandVolIndx As Integer = 21
Private Const constUnitIndx As Integer = 2
Private Const constFactorIndex As Integer = 23


' Copy paste function cannot be reused because of private clear method
Private Sub ShowPickList(ByVal strTbl As String, ByVal strRange As String, ByVal intStart As Integer, ByVal intMax As Integer)

    Dim frmPickList As FormPedMedIVPickList
    Dim colTbl As Collection
    Dim intN As Integer
    Dim strN As String
    Dim intKeuze As Integer
    
    Set colTbl = ModRange.CollectionFromRange(strTbl, intStart)
    
    Set frmPickList = New FormPedMedIVPickList
    frmPickList.LoadMedicamenten colTbl
    
    For intN = 1 To intMax
        strN = IIf(intMax > 9, IntNToStrN(intN), intN)
        intKeuze = ModRange.GetRangeValue(strRange & strN, 1)
        If intKeuze > 1 Then frmPickList.SelectMedicament intKeuze
    Next intN
    
    frmPickList.Show
    
    If frmPickList.GetAction = vbNullString Then
    
        For intN = 1 To intMax                 ' First remove nonselected items
            strN = IntNToStrN(intN)
            intKeuze = ModRange.GetRangeValue(strRange & strN, 1)
            If intKeuze > 1 Then
                If frmPickList.IsMedicamentSelected(intKeuze) Then
                    frmPickList.UnselectMedicament (intKeuze)
                Else
                    Clear intN ' Remove is specific to PedContIV replace with appropriate sub when copy paste
                End If
            End If
        Next intN
        
        Do While frmPickList.HasSelectedMedicamenten()  ' Then add selected items
            For intN = 1 To intMax
                strN = IntNToStrN(intN)
                intKeuze = ModRange.GetRangeValue(strRange & strN, 1)
                If intKeuze <= 1 Then
                    intKeuze = frmPickList.GetFirstSelectedMedicament(True)
                    ModRange.SetRangeValue strRange & strN, intKeuze
                    SetToStandard intN
                    Exit For
                End If
            Next intN
        Loop
    
    End If
    
End Sub

Public Sub PedContIV_ShowPickList()

    ShowPickList constTblMed, constMedIVKeuze, 2, constMedIVCount
    
End Sub

Private Sub Clear(ByVal intN As Integer)

    Dim strN As String
    
    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim strStand As String
    
    strN = IIf(intN < 10, "0" + intN, intN)
    
    strN = IntNToStrN(intN)
    strMedicament = constMedIVKeuze & strN
    strMedSterkte = constMedIVSterkte & strN
    strOplHoev = constMedIVOplVol & strN
    strOplossing = constMedIVOplVlst & strN
    strStand = constMedIVStand & strN
    
    If intN < 16 Then
        ModRange.SetRangeValue strMedicament, 1
    Else
        ModRange.SetRangeValue strMedicament, vbNullString
        ModRange.SetRangeValue "_Ped_MedIV_DoseUnit_" & strN, ""
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

    Clear 9

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

' ToDo calculate drip
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

    strN = IntNToStrN(intN)
    strMedicament = constMedIVKeuze & strN
    strMedSterkte = constMedIVSterkte & strN
    strOplHoev = constMedIVOplVol & strN
    strOplossing = constMedIVOplVlst & strN
    strStand = constMedIVStand & strN
    
    ModRange.SetRangeValue strMedSterkte, 0
    ModRange.SetRangeValue strOplHoev, 0
    ModRange.SetRangeValue strStand, 0
    
    intKeuze = ModRange.GetRangeValue(strMedicament, 0)
    If intKeuze = 0 Then GoTo SetToStandardError ' Something is wrong, 0 is no valid value
    
    If intKeuze = 1 Then                         ' No medicament was selected so clear the line
        Clear intN
    Else                                         ' Else find the right standard solution
        varOplossing = ModExcel.Excel_VLookup(shtPedTblMedIV.Range(constTblMed).Cells(intKeuze, 1), constTblMed, constStandOplVlst)
        varOplossing = IIf(varOplossing = 0, constStandOplKeuze, varOplossing) ' Use NaCl 0.9% as stand solution if not specified otherwise
        ModRange.SetRangeValue strOplossing, varOplossing
    End If
    
    Exit Sub
    
SetToStandardError:

    ModLog.LogError Err, "SetMedContIVToStandard: " & " Error for regel " & strN

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
    Dim strMed As String
    
    On Error GoTo OpenInvoerNumeriekError
    
    Set frmInvoer = New FormInvoerNumeriek
    
    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    varKeuze = ModRange.GetRangeValue(constMedIVKeuze & strRegel, vbNullString)
    strMed = ModExcel.Excel_Index(constTblMed, varKeuze, 1)
    
    With frmInvoer
        .Caption = "Medicament " & intRegel
        .lblText.Caption = "Voer hoeveelheid in voor " & strMed
        .lblParameter = "Oplossing"
        .lblEenheid = strUnit
        If ModRange.GetRangeValue(constMedIVOplVol & strRegel, 0) = 0 Then
            .txtWaarde = ModExcel.Excel_Index(constTblMed, varKeuze, intColumn)
        Else
            .txtWaarde = ModRange.GetRangeValue(strRange & strRegel, vbNullString)
        End If
        .Show
        If IsNumeric(.txtWaarde) Then
            If CDbl(.txtWaarde) = ModExcel.Excel_Index(constTblMed, varKeuze, intColumn) Then
                ModRange.SetRangeValue strRange & strRegel, 0
            Else
                ModRange.SetRangeValue strRange & strRegel, .txtWaarde
            End If
        End If
    End With
    
    Exit Sub
    
OpenInvoerNumeriekError:

    ModLog.LogError Err, "EnterNumeric(" & Join(Array(strRegel, strRange, strUnit, intColumn), ", ") & ")"
    Set frmInvoer = Nothing

End Sub

Private Sub SetMedConc(ByVal intRegel As Integer)

    Dim strUnit As String
    Dim strRegel As String
    
    On Error GoTo SetMedConcError

    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    strUnit = Application.WorksheetFunction.Index(shtPedTblMedIV.Range(constTblMed), ModRange.GetRange(constMedIVKeuze & strRegel), constUnitIndx)
    EnterNumeric intRegel, constMedIVSterkte, strUnit, constStandHoevIndx
    
    Exit Sub
    
SetMedConcError:

    ModLog.LogError Err, "SetMedConc(" & intRegel & ")"

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

    EnterNumeric intRegel, constMedIVOplVol, "mL", constStandVolIndx

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
    Dim strDoseUnit As String
    Dim dblOplVol As Double
    
    Dim arrSterkte() As String
    Dim frmMedIV As FormMedIV
    
    Set frmMedIV = New FormMedIV
    
    strMed = ModRange.GetRangeValue(constMedIVKeuze & intN, vbNullString)
    strSterkte = ModRange.GetRangeValue(constMedIVSterkte & intN, vbNullString)
    strDoseUnit = ModRange.GetRangeValue("_Ped_MedIV_DoseUnit_" & intN, vbNullString)
    dblOplVol = ModRange.GetRangeValue("_Ped_MedIV_OplVol_" & intN, vbNullString)
    
    arrSterkte = Split(strSterkte, " ")
    
    frmMedIV.txtMedicament.Text = strMed
    frmMedIV.txtSterkte.Text = ModArray.StringArrayItem(arrSterkte, 0)
    frmMedIV.cboUnit.Value = ModArray.StringArrayItem(arrSterkte, 1)
    frmMedIV.txtSolVol = dblOplVol
    frmMedIV.cboDoseUnit.Value = strDoseUnit
    
    frmMedIV.Show
    
    If frmMedIV.lblValid.Caption = vbNullString Then
    
        strMed = frmMedIV.txtMedicament.Text
        strSterkte = frmMedIV.txtSterkte.Text & " " & Trim(frmMedIV.cboUnit.Text)
        ModRange.SetRangeValue constMedIVKeuze & intN, strMed
        ModRange.SetRangeValue constMedIVSterkte & intN, strSterkte
        ModRange.SetRangeValue "_Ped_MedIV_DoseUnit_" & intN, frmMedIV.cboDoseUnit.Text
        ModRange.SetRangeValue "_Ped_MedIV_OplVol_" & intN, frmMedIV.txtSolVol.Text
    
    End If
    
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

Public Sub PedContIV_TextClear()

    ModRange.SetRangeValue constMedIVOpm, vbNullString

End Sub

Public Sub PedContIV_Text()

    Dim frmOpmerking As FormOpmerking
    
    Set frmOpmerking = New FormOpmerking
    
    frmOpmerking.SetText ModRange.GetRangeValue(constMedIVOpm, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constMedIVOpm, frmOpmerking.txtOpmerking.Text
    End If
    
End Sub


Private Sub ResetOplVlst(ByVal strOpl As String, ByVal intOpl As Integer)

    ModMessage.ShowMsgBoxInfo "Ongeldige oplossing vloeistof voor dit medicament"
    ModRange.SetRangeValue strOpl, intOpl

End Sub

Private Sub CheckOplVlst(ByVal intN As Integer)
    
    Dim strN As String
    Dim intMed As Integer
    Dim intOplVlst As Integer
    Dim intAdvVlst As Integer
    
    strN = ModString.IntNToStrN(intN)
    intMed = ModRange.GetRangeValue(constMedIVKeuze & strN, 0)
    If intMed > 0 Then
        intAdvVlst = ModExcel.Excel_VLookup(shtPedTblMedIV.Range(constTblMed).Cells(intMed, 1), constTblMed, constStandOplVlst)
        'intAdvVlst = GetMedicamentItemWithIndex(intMed, constAdvOplIndex)
        intOplVlst = ModRange.GetRangeValue(constMedIVOplVlst & strN, 0)
        'Geen oplossing vloeistof
        If intAdvVlst = 1 And Not intOplVlst = 1 Then
            ResetOplVlst constMedIVOplVlst & strN, intAdvVlst
        End If
        'Oplossing vloeistof is NaCl
        If intAdvVlst = 2 And Not intOplVlst = 2 Then
            ResetOplVlst constMedIVOplVlst & strN, intAdvVlst
        End If
        'Oplossing vloeistof is glucose
        If intAdvVlst > 2 And intOplVlst <= 2 Then
            ResetOplVlst constMedIVOplVlst & strN, intAdvVlst
        End If
                
    End If
    
End Sub

Public Sub PedContIV_CheckOplVlst_01()

    CheckOplVlst 1

End Sub

Public Sub PedContIV_CheckOplVlst_02()

    CheckOplVlst 2

End Sub

Public Sub PedContIV_CheckOplVlst_03()

    CheckOplVlst 3

End Sub

Public Sub PedContIV_CheckOplVlst_04()

    CheckOplVlst 4

End Sub

Public Sub PedContIV_CheckOplVlst_05()

    CheckOplVlst 5

End Sub

Public Sub PedContIV_CheckOplVlst_06()

    CheckOplVlst 6

End Sub

Public Sub PedContIV_CheckOplVlst_07()

    CheckOplVlst 7

End Sub

Public Sub PedContIV_CheckOplVlst_08()

    CheckOplVlst 8

End Sub

Public Sub PedContIV_CheckOplVlst_09()

    CheckOplVlst 9

End Sub

Public Sub PedContIV_CheckOplVlst_10()

    CheckOplVlst 10

End Sub

Public Sub PedContIV_CheckOplVlst_11()

    CheckOplVlst 11

End Sub

Public Sub PedContIV_CheckOplVlst_12()

    CheckOplVlst 12

End Sub

Public Sub PedContIV_CheckOplVlst_13()

    CheckOplVlst 13

End Sub

Public Sub PedContIV_CheckOplVlst_14()

    CheckOplVlst 14

End Sub

Public Sub PedContIV_CheckOplVlst_15()

    CheckOplVlst 15

End Sub

Private Function CalculateStandByDose(ByVal intN As Integer, ByVal intMed As Integer, ByVal dblDose As Double) As Double

    Dim strN As String
    Dim dblFactor As Double
    Dim dblWeight As Double
    Dim dblOplQty As Double
    Dim dblStand As Double
    Dim dblQty As Double
    Dim intPrec As Integer
    Dim strMed As String
    
    strN = IntNToStrN(intN)
    dblFactor = GetMedicamentItemWithIndex(intMed, constFactorIndex)
    dblOplQty = ModRange.GetRangeValue(constMedIVOplVol & strN, 0)
    dblOplQty = IIf(dblOplQty = 0, ModExcel.Excel_Index("Tbl_Ped_BerMedCont", intN, 9), dblOplQty)
    dblQty = ModRange.GetRangeValue(constMedIVSterkte & strN, 0)
    dblQty = IIf(dblQty = 0, ModExcel.Excel_Index("Tbl_Ped_BerMedCont", intN, 8), dblQty)
    dblWeight = ModPatient.Patient_GetWeight()
    strMed = GetMedicamentItemWithIndex(intMed, 1)
    
    If dblQty > 0 Then dblStand = dblDose * dblOplQty * dblWeight / (dblQty * dblFactor)
    
    CalculateStandByDose = dblStand

End Function

Private Function GetMedicamentItemWithIndex(ByVal intMed As Integer, ByVal intIndex As Integer) As Variant

    Dim objTblMed As Range
    
    Set objTblMed = ModRange.GetRange(constTblMed)

    GetMedicamentItemWithIndex = objTblMed.Cells(intMed, intIndex).Value2
    
End Function

Private Sub SetStandByDose(ByVal intN As Integer)

    Dim strN As String
    Dim frmDose As FormInvoerNumeriek
    Dim intMed As Integer
    Dim strMed As String
    Dim dblRate As Double
    Dim dblDose As Double
    Dim varDose As Variant
    Dim strEenheid As String
    
    strN = IntNToStrN(intN)
    intMed = ModRange.GetRangeValue(constMedIVKeuze & strN, vbNullString)
    strMed = ModExcel.Excel_Index(constTblMed, intMed, 1)
    
    If strMed = vbNullString Or ModString.ContainsCaseInsensitive(strMed, "EPI") Then Exit Sub
    
    strEenheid = ModExcel.Excel_Index("Tbl_Ped_BerMedCont", intN, 38)
    
    varDose = ModExcel.Excel_Index("Tbl_Ped_BerMedCont", intN, 10)
    dblDose = ModString.StringToDouble(varDose)
    
    Set frmDose = New FormInvoerNumeriek
    
    With frmDose
        .lblText.Caption = "Voer dosering in voor " & strMed
        .SetValue vbNullString, "Dose:", dblDose, strEenheid, vbNullString
        
        .Show
        
        If Not .txtWaarde.Value = vbNullString Then
            dblDose = StringToDouble(.txtWaarde.Value)
            dblRate = CalculateStandByDose(intN, intMed, dblDose)
            dblRate = ModExcel.Excel_RoundBy(dblRate, 0.1)
            ModRange.SetRangeValue constMedIVStand & strN, dblRate * 10
        End If
    End With

End Sub

Public Sub PedContIV_SetRate_01()

    SetStandByDose 1

End Sub

Public Sub PedContIV_SetRate_02()

    SetStandByDose 2

End Sub

Public Sub PedContIV_SetRate_03()

    SetStandByDose 3

End Sub

Public Sub PedContIV_SetRate_04()

    SetStandByDose 4

End Sub


Public Sub PedContIV_SetRate_05()

    SetStandByDose 5

End Sub

Public Sub PedContIV_SetRate_06()

    SetStandByDose 6

End Sub

Public Sub PedContIV_SetRate_07()

    SetStandByDose 7

End Sub

Public Sub PedContIV_SetRate_08()

    SetStandByDose 8

End Sub

Public Sub PedContIV_SetRate_09()

    SetStandByDose 9

End Sub

Public Sub PedContIV_SetRate_10()

    SetStandByDose 10

End Sub

Public Sub PedContIV_SetRate_11()

    SetStandByDose 11

End Sub
Public Sub PedContIV_SetRate_12()

    SetStandByDose 12

End Sub
Public Sub PedContIV_SetRate_13()

    SetStandByDose 13

End Sub
Public Sub PedContIV_SetRate_14()

    SetStandByDose 14

End Sub
Public Sub PedContIV_SetRate_15()

    SetStandByDose 15

End Sub


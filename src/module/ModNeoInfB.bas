Attribute VB_Name = "ModNeoInfB"
Option Explicit

Private Const constIntakeAdvice As String = "B7"

Private Const constTblMedIV As String = "Tbl_Neo_MedIV"
Private Const constTblNeoOpl As String = "Tbl_Neo_OplVlst"
Private Const constMedIVMax As Integer = 10
Private Const constDoseEenheidIndx As Integer = 3

Private Const constActInfB As String = "Actuele Afspraken"
Private Const const1700InfB As String = "17:00 uur Afspraken"
Private Const constInfbVersie As String = "B3"

Private Const constInfBVa As String = "Var_Neo_InfB"
Private Const constInfBDataAct As String = "_Neo_InfB"
Private Const constInfBData1700 As String = "_Neo_1700"

Private Const constCont As String = "_Cont_"
Private Const constEnt As String = "_Ent_"
Private Const constTPN As String = "_TPN_"

Private Const constIntakePerKg As String = "Var_Neo_InfB_TPN_IntakePerKg"
Private Const constTPNParEnt As String = "Var_Neo_InfB_TPN_ParEnteraal"

Private Const constLipidKeuze As String = "Var_Neo_InfB_TPN_IntralipidSmof"
Private Const constLipidStand As String = "Var_Neo_InfB_TPN_IntraLipid"
Private Const constLipidAdvies As String = "Var_Neo_LipidAdv"

Private Const constNaCl As String = "Var_Neo_InfB_TPN_NaCl"
Private Const constKAct As String = "Var_Neo_InfB_TPN_KCl"
Private Const constCaCl As String = "Var_Neo_InfB_TPN_CaCl2"
Private Const constMgCl As String = "Var_Neo_InfB_TPN_MgCl2"
Private Const constSolu As String = "Var_Neo_InfB_TPN_Soluvit"
Private Const constPrim As String = "Var_Neo_InfB_TPN_Primene"

Private Const constNICUmix As String = "Var_Neo_InfB_TPN_NICUMix"
Private Const constSSTB As String = "Var_Neo_InfB_TPN_SSTB"

Private Const constMedIndex As Integer = 1
Private Const constUnitIndex As Integer = 2
Private Const constAdvMinIndex As Integer = 6
Private Const constAdvOplIndex As Integer = 11
Private Const constDefHoevIndex As Integer = 13
Private Const constDefStandIndex As Integer = 14
Private Const constFactorIndex As Integer = 20

Private Const constMedKeuze As String = "Var_Neo_InfB_Cont_MedKeuze_"
Private Const constMedSterkte As String = "Var_Neo_InfB_Cont_MedSterkte_"
Private Const constOplHoev As String = "Var_Neo_InfB_Cont_OplHoev_"
Private Const constOplossing As String = "Var_Neo_InfB_Cont_Oplossing_"
Private Const constStand As String = "Var_Neo_InfB_Cont_Stand_"
Private Const constExtra As String = "Var_Neo_InfB_Cont_VochtExtra_"

Private Const constPrevVoed As String = "Var_Neo_PrevVoed"
Private Const constEntVoed As String = "Var_Neo_InfB_Ent_Voeding"
Private Const constEntFreq As String = "Var_Neo_InfB_Ent_Frequentie_"
Private Const constEntHoev As String = "Var_Neo_InfB_Ent_Hoeveelheid_"
Private Const constEntToev As String = "Var_Neo_InfB_Ent_Toevoeging_"
Private Const constEntPerc As String = "Var_Neo_InfB_Ent_PercentageKeuze_"
Private Const constEntExtra As String = "Var_Neo_InfB_Ent_Extra"

Private Const constArtLijn As String = "Var_Neo_InfB_Cont_ArtLijn"

Private Const constMedText As String = "Var_Neo_InfB_Cont_MedTekst_"

Private Const constTblVoeding As String = "Tbl_Neo_Voed"
Private Const constVoedingCount As Integer = 1
Private Const constTblToevoegMM As String = "Tbl_Neo_PoedMM"
Private Const constToevoegMMCount As Integer = 4
Private Const constTblToevoegKV As String = "Tbl_Neo_PoedKV"
Private Const constToevoegKVCount As Integer = 4

Private Const constTPNGluc As String = "AL50"

Private Const constStandLipid As String = "B40"
Private Const constStandNaCl As String = "B42"
Private Const constStandKAcet As String = "B43"
Private Const constStandCaCl2 As String = "B44"
Private Const constStandMgCl2 As String = "B45"
Private Const constStandSoluvit As String = "B46"
Private Const constStandPrimene As String = "B47"
Private Const constStandNICUmix As String = "B48"
Private Const constStandSSTB As String = "B49"


Public Function IsEpiduraal(ByVal strText As String) As Boolean

    IsEpiduraal = ModString.ContainsCaseSensitive(strText, "EPI")

End Function

Private Function Is1700() As Boolean

    Is1700 = shtNeoBerInfB.Range(constInfbVersie) = const1700InfB

End Function

Public Function NeoInfB_Is1700() As Boolean

    NeoInfB_Is1700 = Is1700()

End Function


Private Sub CopyVarData(ByVal bln1700 As Boolean, ByVal blnToVar As Boolean, ByVal blnShowProgress As Boolean)

    Dim objName As Name
    Dim strStartsWith As String
    Dim varVarValue As Variant
    Dim varDataValue As Variant
    Dim strVarName As String
    Dim strDataName As String
    
    Dim intN As Integer
    Dim intC As Integer
    
    strStartsWith = IIf(bln1700, constInfBData1700, constInfBDataAct)
    
    intN = 1
    intC = WbkAfspraken.Names.Count
    For Each objName In WbkAfspraken.Names
        If ModString.StartsWith(objName.Name, strStartsWith) Then
            strDataName = objName.Name
            strVarName = "Var" & IIf(bln1700, Strings.Replace(objName.Name, "1700", "InfB"), objName.Name)
            
            varVarValue = ModRange.GetRangeValue(strVarName, vbNullString)
            varDataValue = ModRange.GetRangeValue(strDataName, vbNullString)
            
            If blnToVar Then
                ModRange.SetRangeValue strVarName, varDataValue
            Else
                ModRange.SetRangeValue strDataName, varVarValue
            End If
            
            If blnShowProgress Then
                ModProgress.SetJobPercentage "Data verplaatsen", intC, intN
                intN = intN + 1
            End If
            
        End If
    Next

End Sub

Public Sub CopyCurrentInfVarToData(ByVal blnShowProgress As Boolean)

    CopyVarData Is1700(), False, blnShowProgress
    
End Sub

Public Sub CopyCurrentInfDataToVar(ByVal blnShowProgress As Boolean)

    CopyVarData Is1700(), True, blnShowProgress
    
End Sub

Private Sub TestCopyVarData()

    CopyVarData True, True, True

End Sub

Public Sub NeoInfB_SelectInfB(ByVal bln1700 As Boolean, ByVal blnStartProgress As Boolean)

    If bln1700 And Is1700() Then                 ' InfB is same as 1700
        ModSheet.GoToSheet shtNeoGuiInfB, "A9"
        
    ElseIf Not bln1700 And Not Is1700() Then     ' InfB is same as act
        ModSheet.GoToSheet shtNeoGuiInfB, "A9"
        
    ElseIf bln1700 And Not Is1700() Then         ' Infb is currently act
        If blnStartProgress Then ModProgress.StartProgress "Infuus brief klaar maken"
        
        CopyVarData False, False, True   ' First copy var data to act data
        CopyVarData True, True, True     ' Then copy 1700 data to var data
        shtNeoBerInfB.Range(constInfbVersie).Value2 = const1700InfB
        
        If blnStartProgress Then ModProgress.FinishProgress
        
        ModSheet.GoToSheet shtNeoGuiInfB, "A9"
    
    ElseIf Not bln1700 And Is1700() Then         ' Infb is currently 1700
        If blnStartProgress Then ModProgress.StartProgress "Infuus brief klaar maken"
        
        CopyVarData True, False, True   ' First copy var data to act data
        CopyVarData False, True, True   ' Then copy act data to var data
        shtNeoBerInfB.Range(constInfbVersie).Value2 = constActInfB
        
        If blnStartProgress Then ModProgress.FinishProgress
        
        ModSheet.GoToSheet shtNeoGuiInfB, "A9"
    
    End If

End Sub

' ToDo: Add comment
Public Sub NeoInfB_CopyActTo1700()

    Dim bln1700 As Boolean

    ModProgress.StartProgress "Overzetten Actueel naar 17:00 uur"

    bln1700 = Is1700()
    If bln1700 Then
        ModNeoInfB.NeoInfB_SelectInfB False, False
    End If
    
    CopyCurrentInfVarToData True

    ModProgress.SetJobPercentage "Voeding", 3, 1
    CopyRangeNamesToRangeNames NeoInfB_GetVoedingItems(), ChangeTo1700(NeoInfB_GetVoedingItems())
    ModProgress.SetJobPercentage "Continue IV", 3, 2
    CopyRangeNamesToRangeNames NeoInfB_GetIVAfsprItems(), ChangeTo1700(NeoInfB_GetIVAfsprItems())
    ModProgress.SetJobPercentage "TPN", 3, 3
    CopyRangeNamesToRangeNames NeoInfB_GetTPNItems(), ChangeTo1700(NeoInfB_GetTPNItems())
    
    If bln1700 Then
        ModNeoInfB.NeoInfB_SelectInfB True, False
    End If
    
    ModProgress.FinishProgress
    
End Sub

Private Function GetItems(ByVal strGrp As String) As String()

    Dim arrItems() As String
    Dim objName As Name
    Dim strStartsWith As String
    
    strStartsWith = constInfBDataAct & strGrp
    
    For Each objName In WbkAfspraken.Names
        If ModString.StartsWith(objName.Name, strStartsWith) Then ModArray.AddItemToStringArray arrItems, objName.Name
    Next
    
    GetItems = arrItems

End Function

Public Function NeoInfB_GetVoedingItems() As String()

    NeoInfB_GetVoedingItems = GetItems(constEnt)

End Function

Private Sub Test_NeoInfB_GetVoedingItems()

    Dim strItem As Variant
    
    For Each strItem In NeoInfB_GetVoedingItems()
        MsgBox strItem
    Next

End Sub

Public Function NeoInfB_GetIVAfsprItems() As String()
    
    NeoInfB_GetIVAfsprItems = GetItems(constCont)

End Function

Public Function NeoInfB_GetTPNItems() As String()
    
    NeoInfB_GetTPNItems = GetItems(constTPN)
    
End Function

Private Function ChangeTo1700(arrItems() As String) As String()
    
    Dim arr1700Items() As String
    Dim varItem As Variant
    Dim intN As Integer
    
    ReDim arr1700Items(UBound(arrItems))
    
    For Each varItem In arrItems
        arr1700Items(intN) = Replace(varItem, "InfB", "1700")
        intN = intN + 1
    Next varItem
    
    ChangeTo1700 = arr1700Items

End Function

' Drie speciale situaties
' 1. SST gluc 10% 1700 = Actueel
' 1. SST gluc 10% 1700 > Actueel
' 1. SST gluc 10% 1700 < Actueel
Private Sub Copy1700ToAct(ByVal blnAlles As Boolean, ByVal blnVoeding As Boolean, ByVal blnContMed As Boolean, ByVal blnTPN As Boolean)
    
    Dim dblGluc As Double
    Dim dblVocht As Double
    
    If ModPatient.GetGewichtFromRange() = 0 Then Exit Sub
    
    dblGluc = shtNeoBerInfB.Range(constTPNGluc).Value2
    
    ModProgress.StartProgress "1700 Afspraken overnemen"
    
    If blnAlles Then
        blnVoeding = True
        blnContMed = True
        blnTPN = True
    End If
    
    If blnVoeding Then
        VoedingOvernemen
    End If
    
    If blnContMed Then
        ContMedOvernemen
    End If
    
    If blnTPN Then
        TPNOvernemen
    End If
    
    ModNeoInfB.NeoInfB_SelectInfB False, True
    
    If dblGluc = shtNeoBerInfB.Range(constTPNGluc).Value2 Then
        ' Do nothing
    Else
        ' Correct for increase or decrease in TPN glucose
        ' Increase or decrease vocht intake
        dblVocht = ModRange.GetRangeValue(constIntakePerKg, 0)
        dblVocht = dblVocht + ((dblGluc - shtNeoBerInfB.Range(constTPNGluc)) / ModPatient.GetGewichtFromRange())
        ModRange.SetRangeValue constIntakePerKg, dblVocht
    End If
    
    ModProgress.FinishProgress

End Sub

Private Sub VoedingOvernemen()

    Dim arrTo() As String
    Dim arrFrom() As String
    
    arrTo = NeoInfB_GetVoedingItems()
    arrFrom = ChangeTo1700(arrTo)
    
    CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

Private Sub ContMedOvernemen()
    
    Dim arrTo() As String
    Dim arrFrom() As String
    
    arrTo = NeoInfB_GetIVAfsprItems()
    arrFrom = ChangeTo1700(arrTo)
    
    CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

Private Sub TPNOvernemen()
    
    Dim arrTo() As String
    Dim arrFrom() As String
    
    arrTo = NeoInfB_GetTPNItems()
    arrFrom = ChangeTo1700(arrTo)
    
    CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

Public Sub NeoInfB_ShowFormCopy1700ToAct()

    Dim frmCopy1700 As FormCopy1700
    Dim bln1700 As Boolean
    
    bln1700 = Is1700()
    
    Set frmCopy1700 = New FormCopy1700
    frmCopy1700.Show
    
    If frmCopy1700.lblAction.Caption = "OK" Then
        Copy1700ToAct frmCopy1700.optAlles.Value, frmCopy1700.chkVoeding.Value, frmCopy1700.chkContinueMedicatie.Value, frmCopy1700.chkTPN.Value
    End If

    NeoInfB_SelectInfB bln1700, True
    
End Sub

Private Sub Test()
    
    Dim varItem As Variant
    Dim arr1700Items() As String
    Dim intN As Integer
    
    arr1700Items = ChangeTo1700(NeoInfB_GetIVAfsprItems())
    For Each varItem In NeoInfB_GetIVAfsprItems()
        MsgBox varItem & ", " & arr1700Items(intN)
        intN = intN + 1
    Next varItem

End Sub

Private Function GetMedicamentItemWithIndex(ByVal intMed As Integer, ByVal intIndex As Integer) As Variant

    Dim objTblMed As Range
    
    Set objTblMed = ModRange.GetRange(constTblMedIV)

    GetMedicamentItemWithIndex = objTblMed.Cells(intMed, intIndex).Value2
    
End Function

Private Sub TestGetMinimumAdvice()

    MsgBox GetMedicamentItemWithIndex(2, constAdvMinIndex)

End Sub

Private Function CalculateMedicamentQtyByDose(ByVal strN As String, ByVal intMed As Integer, ByVal dblDose As Double) As Double

    Dim dblFactor As Double
    Dim dblWeight As Double
    Dim dblOplQty As Double
    Dim dblStand As Double
    Dim dblQty As Double
    Dim dblMaxConc As Double
    Dim intPrec As Integer
    Dim strMed As String
    
    dblFactor = GetMedicamentItemWithIndex(intMed, constFactorIndex)
    dblOplQty = ModRange.GetRangeValue(constOplHoev & strN, 0)
    dblStand = ModRange.GetRangeValue(constStand & strN, 0) / 10
    dblWeight = ModPatient.GetGewichtFromRange()
    dblMaxConc = GetMedicamentItemWithIndex(intMed, 10)
    strMed = GetMedicamentItemWithIndex(intMed, 1)
    
    ' Medicatie doxapram puur geen oplosvolume
    If ModString.ContainsCaseInsensitive(strMed, "doxapram") Then
        dblQty = dblMaxConc * dblOplQty
    ElseIf dblStand * dblFactor > 0 And dblDose > 0 Then
        dblQty = dblDose * dblOplQty * dblWeight / (dblStand * dblFactor)
        dblQty = CorrectMedQty(strN, intMed, dblQty)
    Else
        dblQty = 0
    End If
    
    CalculateMedicamentQtyByDose = dblQty

End Function

Private Sub ChangeMedContIV(ByVal intN As Integer, ByVal blnRemove As Boolean)

    Dim intMedIndx As Integer
    Dim strMedName As String
    Dim strN As String
      
    Dim dblOplQty As Double
    Dim dblMedQty As Double
    Dim dblStand As Double
    
    Dim objTblMed As Range
    
    Set objTblMed = ModRange.GetRange(constTblMedIV)
    strN = IntNToStrN(intN)
    
    intMedIndx = IIf(blnRemove, 1, ModRange.GetRangeValue(constMedKeuze & strN, vbNullString))
    
    dblOplQty = GetMedicamentItemWithIndex(intMedIndx, constDefHoevIndex)
    dblStand = GetMedicamentItemWithIndex(intMedIndx, constDefStandIndex) * 10
    
    If blnRemove Then ModRange.SetRangeValue constMedKeuze & strN, 1
    
    
    ModRange.SetRangeValue constOplHoev & strN, IIf(blnRemove, 0, dblOplQty)
    ModRange.SetRangeValue constStand & strN, IIf(blnRemove, 0, dblStand)
    ModRange.SetRangeValue constExtra & strN, False
    
    ModRange.SetRangeValue constOplossing & strN, Application.VLookup(objTblMed.Cells(intMedIndx, 1), objTblMed, constAdvOplIndex, False)
    If Not IsNumeric(ModRange.GetRangeValue(constOplossing & strN, vbNullString)) Then
        ModRange.SetRangeValue constOplossing & strN, 1
    End If
    
    If blnRemove Then
        ModRange.SetRangeValue constMedSterkte & strN, 0
    Else
        strMedName = ModMedicatie.GetNeoMedContIVName(intMedIndx)
        If IsEpiduraal(strMedName) Then
            dblMedQty = ModMedicatie.Medicatie_CalcEpiQty(ModPatient.GetGewichtFromRange())
            dblMedQty = CorrectMedQty(strN, intMedIndx, dblMedQty)
        ElseIf strMedName = "doxapram" Then
            dblMedQty = CalculateMedicamentQtyByDose(strN, intMedIndx, 0)
        Else
            dblMedQty = 0
        End If
    End If
    
    ModRange.SetRangeValue constMedSterkte & strN, dblMedQty * 10


End Sub

Public Sub NeoInfB_ChangeMedContIV_01()

    ChangeMedContIV 1, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_02()
    
    ChangeMedContIV 2, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_03()
    
    ChangeMedContIV 3, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_04()
    
    ChangeMedContIV 4, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_05()
    
    ChangeMedContIV 5, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_06()
    
    ChangeMedContIV 6, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_07()
    
    ChangeMedContIV 7, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_08()
    
    ChangeMedContIV 8, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_09()
    
    ChangeMedContIV 9, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_10()
    
    ChangeMedContIV 10, False

End Sub

Public Sub NeoInfB_RemoveMedContIV_01()

    ChangeMedContIV 1, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_02()
    
    ChangeMedContIV 2, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_03()
    
    ChangeMedContIV 3, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_04()
    
    ChangeMedContIV 4, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_05()
    
    ChangeMedContIV 5, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_06()
    
    ChangeMedContIV 6, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_07()
    
    ChangeMedContIV 7, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_08()
    
    ChangeMedContIV 8, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_09()
    
    ChangeMedContIV 9, True

End Sub

Public Sub NeoInfB_RemoveMedContIV_10()
    
    ChangeMedContIV 10, True

End Sub

Public Sub NeoInfB_RemvoveArtLijn()

    ModRange.SetRangeValue constArtLijn, 0
    ModRange.SetRangeValue constExtra + "01", False

End Sub

Public Sub NeoInfB_IntakeAdvies()

    ModRange.SetRangeValue constIntakePerKg, shtNeoBerInfB.Range(constIntakeAdvice).Value2

End Sub

Public Sub NeoInfB_IntakePerKg()

    Dim frmInvoer As FormInvoerNumeriek
    
    Set frmInvoer = New FormInvoerNumeriek
    
    With frmInvoer
        .Caption = "Vocht per Kg/dag"
        .lblText = "Voer vocht per Kg/dag in"
        .lblParameter = "Vocht"
        .lblEenheid = "ml/kg/dag"
        .txtWaarde = ModRange.GetRangeValue(constIntakePerKg, 0)
        .Show
        If IsNumeric(.txtWaarde) Then
            ModRange.SetRangeValue constIntakePerKg, .txtWaarde
        End If
    End With

End Sub

Private Function CorrectMedQty(ByVal strN As String, ByVal intMed As Integer, ByVal dblQty As Double) As Double

    Dim dblMultiple As Double
    Dim intFactor As Integer
    Dim dblOplQty As Double
    Dim dblMaxConc As Double
    Dim dblMinConc As Double
    Dim dblConc As Double
    Dim dblFilter As Double
    Dim intN As Integer
    
    dblOplQty = ModRange.GetRangeValue(constOplHoev & strN, 0)
    
    If dblOplQty = 0 Then
        dblQty = 0
    Else
        ' Haal de minimale en maximale concentratie op
        dblMinConc = GetMedicamentItemWithIndex(intMed, 9)
        dblMaxConc = GetMedicamentItemWithIndex(intMed, 10)
        ' Corrigeer de hoeveelheid naar min of max concentratie
        dblConc = dblQty / dblOplQty
        If dblMinConc > 0 Then dblQty = IIf(dblConc < dblMinConc, dblMinConc * dblOplQty, dblQty)
        If dblMaxConc > 0 Then dblQty = IIf(dblConc > dblMaxConc, dblMaxConc * dblOplQty, dblQty)
        ' Bepaal of hoevelheid een veelvoud van hele, tiende of honderste milliliters is
        dblMultiple = ModExcel.Excel_Index(constTblMedIV, intMed, 4)
        intFactor = 1                                                   ' >=10ml  hele getallen geen decimalen
        intFactor = IIf(dblQty / dblMultiple < 10, 10, intFactor)       ' >=1,0  <10 ml   1 decimaal nauwkeurig 0,1
        intFactor = IIf(dblQty / dblMultiple < 1, 100, intFactor)       ' >=0,1 < 1,0 ml:     2 decimalen nauwkeurig 0,01
        intFactor = IIf(dblQty / dblMultiple <= 0.1, 100, intFactor)    ' <0,1ml  2 decimalen nauwkeurig 0,01 + verdunningstekst
        dblMultiple = dblMultiple / intFactor
        ' Corrigeer de hoeveelheid naar een veelvoud
        If dblMultiple > 0 Then
            dblQty = ModExcel.Excel_RoundBy(dblQty, dblMultiple)
            If dblQty = 0 Then dblQty = dblMultiple
        End If
        ' Check opnieuw of de minimale concentratie wordt overschreden
        ' Voeg anders steeds 1 veelvoud van de hoeveelheid toe
        dblConc = dblQty / dblOplQty
        Do While dblConc < dblMinConc And Not dblMinConc = 0
            dblQty = dblQty + dblMultiple
            dblConc = dblQty / dblOplQty
        Loop
        ' Check opnieuw of de maximale concentratie wordt overschreden
        ' Haal anders steeds 1 veelvoud van de hoeveelheid af
        dblConc = dblQty / dblOplQty
        Do While dblConc > dblMaxConc And Not dblMaxConc = 0
            dblQty = dblQty - dblMultiple
            dblConc = dblQty / dblOplQty
        Loop
    
    End If
    
    CorrectMedQty = dblQty
    
End Function

Private Sub MedSterkte(ByVal intN As Integer)

    Dim frmInvoer As FormInvoerNumeriek
    Dim intMed As Integer
    Dim strMed As String
    Dim dblQty As Double
    Dim strUnit As String
    
    Dim objTblMed As Range
    
    intMed = ModRange.GetRangeValue(constMedKeuze & IntNToStrN(intN), 1)
    strMed = ModExcel.Excel_Index(constTblMedIV, intMed, 1)

    If intMed <= 1 Then Exit Sub
    
    Set objTblMed = Range(constTblMedIV)
    strUnit = Application.VLookup(objTblMed.Cells(intMed, 1), objTblMed, constUnitIndex, False)
    
    Set frmInvoer = New FormInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament " & intN
        .lblText.Caption = "Voer de medicament sterkte in voor " & strMed
        .lblParameter = "Sterkte"
        .lblEenheid = strUnit
        .txtWaarde = ModRange.GetRangeValue(constMedSterkte & IntNToStrN(intN), 0) / 10
        .Show
        
        If IsNumeric(.txtWaarde) Then
            dblQty = ModString.StringToDouble(.txtWaarde)
            SetMedSterkteNeoInfB intN, dblQty
        End If
        
    End With
    
End Sub

Public Function SetMedSterkteNeoInfB(intN As Integer, dblQty As Double) As Boolean

    Dim strN As String
    Dim intMed As Integer
    Dim strMed As String
    
    strN = IntNToStrN(intN)
    intMed = ModRange.GetRangeValue(constMedKeuze & strN, 1)
    strMed = GetMedicamentItemWithIndex(intMed, 1)

    If Not IsEpiduraal(strMed) Then dblQty = CorrectMedQty(strN, intMed, dblQty)
    SetMedSterkteNeoInfB = ModRange.SetRangeValue(constMedSterkte & strN, dblQty * 10)
    
End Function

Public Sub NeoInfB_MedConc_01()

    MedSterkte 1

End Sub

Public Sub NeoInfB_MedConc_02()

    MedSterkte 2

End Sub

Public Sub NeoInfB_MedConc_03()

    MedSterkte 3

End Sub

Public Sub NeoInfB_MedConc_04()

    MedSterkte 4

End Sub

Public Sub NeoInfB_MedConc_05()

        MedSterkte 5

End Sub

Public Sub NeoInfB_MedConc_06()

    MedSterkte 6

End Sub

Public Sub NeoInfB_MedConc_07()

    MedSterkte 7

End Sub

Public Sub NeoInfB_MedConc_08()

    MedSterkte 8

End Sub

Public Sub NeoInfB_MedConc_09()

    MedSterkte 9

End Sub

Public Sub NeoInfB_MedConc_10()

    MedSterkte 10

End Sub

Private Sub ChangeIV(ByVal intRegel As Integer, ByVal blnRemove As Boolean)

    Dim strStand As String
    Dim strExtra As String

    strStand = constStand & intRegel
    strExtra = constExtra & intRegel + 1
    blnRemove = ModRange.GetRangeValue(constOplossing & intRegel, 1) = 1
    
    If blnRemove Then
        ModRange.SetRangeValue constOplossing & intRegel, 1
        ModRange.SetRangeValue strStand, 0
        ModRange.SetRangeValue strExtra, False
    End If
    
End Sub

Public Sub NeoInfB_ChangeIV_11()
    
    ChangeIV 11, False

End Sub

Public Sub NeoInfB_ChangeIV_12()
    
    ChangeIV 12, False

End Sub

Public Sub NeoInfB_ChangeIV_13()
    
    ChangeIV 13, False

End Sub

Public Sub NeoInfB_RemoveIV_11()
    
    ChangeIV 11, True

End Sub

Public Sub NeoInfB_RemoveIV_12()
    
    ChangeIV 12, True

End Sub

Public Sub NeoInfB_RemoveIV_13()
    
    ChangeIV 13, True

End Sub

Public Sub NeoInfB_TPNAdvice()

    NeoInfB_StandardTPN
    NeoInfB_StandardLipid
    
    ModSheet.GoToSheet shtNeoGuiInfB, "A9"

End Sub

Private Sub EnterText(ByVal strCaption As String, ByVal strName As String, ByVal strRange As String)

    Dim frmInvoer As FormTekstInvoer
    
    Set frmInvoer = New FormTekstInvoer
    
    With frmInvoer
        .Caption = strCaption
        .lblNaam.Caption = strName
        .Tekst = ModRange.GetRangeValue(strRange, vbNullString)
        .Show
        If .IsOK Then ModRange.SetRangeValue strRange, .Tekst
    End With
    
End Sub

Public Sub NeoInfB_MedText_1()

    EnterText "Voer tekst in ...", "Tekst voor medicatie 14", constMedText & "1"

End Sub

Public Sub NeoInfB_MedText_2()

    EnterText "Voer tekst in ...", "Tekst voor medicatie 15", constMedText & "2"
    
End Sub

Public Sub NeoInfB_RemoveMedText_1()

    ModRange.SetRangeValue constMedText & "1", vbNullString

End Sub

Public Sub NeoInfB_RemoveMedText_2()

    ModRange.SetRangeValue constMedText & "2", vbNullString

End Sub

Private Sub ChangeVoeding(ByVal blnRemove As Boolean)

    Dim intN As Integer

    If blnRemove Then
        ModRange.SetRangeValue constEntVoed, 1
        ModRange.SetRangeValue constPrevVoed, 1
    Else
        ModRange.SetRangeValue constPrevVoed, ModRange.GetRangeValue(constEntVoed, 1)
    End If
    
    ModRange.SetRangeValue constEntFreq & 1, 0
    ModRange.SetRangeValue constEntFreq & 2, 0
    ModRange.SetRangeValue constEntHoev & 1, 0
    ModRange.SetRangeValue constEntHoev & 2, 0
    
    ModRange.SetRangeValue constEntExtra, 2
    
    ModRange.SetRangeValue constEntPerc & 0, 1
    For intN = 1 To 8
        ModRange.SetRangeValue constEntToev & intN, 1
        ModRange.SetRangeValue constEntPerc & intN, 1
    Next intN
    
End Sub

Public Sub NeoInfB_ChangeVoed()

    Dim intPrev As Integer
    Dim intVoed As Integer
    
    intPrev = ModRange.GetRangeValue(constPrevVoed, 1)
    intVoed = ModRange.GetRangeValue(constEntVoed, 1)
    
    If intPrev = 1 Then
        ChangeVoeding False
    ElseIf intVoed = 1 Then
        ChangeVoeding True
    End If
    
End Sub

Public Sub NeoInfB_RemoveVoed()

    ChangeVoeding True
    
End Sub

Private Sub ChangeToevoeg(ByVal blnMM As Boolean, ByVal intN As Integer, ByVal blnRemove As Boolean)

    intN = IIf(blnMM, intN, intN + 4)
    blnRemove = blnRemove Or (ModRange.GetRangeValue(constEntToev & intN, 1) = 1)
    
    ModRange.SetRangeValue constEntPerc & intN, 1
    If blnRemove Then ModRange.SetRangeValue constEntToev & intN, 1
    

End Sub

Public Sub NeoInfB_ChangeToevMM_1()

    ChangeToevoeg True, 1, False

End Sub

Public Sub NeoInfB_ChangeToevMM_2()

    ChangeToevoeg True, 2, False

End Sub

Public Sub NeoInfB_ChangeToevMM_3()

    ChangeToevoeg True, 3, False

End Sub

Public Sub NeoInfB_ChangeToevMM_4()

    ChangeToevoeg True, 4, False

End Sub

Public Sub NeoInfB_ChangeToevKV_1()

    ChangeToevoeg False, 1, False

End Sub

Public Sub NeoInfB_ChangeToevKV_2()

    ChangeToevoeg False, 2, False

End Sub

Public Sub NeoInfB_ChangeToevKV_3()

    ChangeToevoeg False, 3, False

End Sub

Public Sub NeoInfB_ChangeToevKV_4()

    ChangeToevoeg False, 4, False

End Sub

Private Sub ChangeIntakePerKg(ByVal blnRemove As Boolean)

    If blnRemove Then
        ModRange.SetRangeValue constIntakePerKg, 0
        ModRange.SetRangeValue constTPNParEnt, False
    Else
        ModRange.SetRangeValue constTPNParEnt, True
    End If

End Sub

Public Sub NeoInfB_ChangeIntakePerKg()

    ChangeIntakePerKg False

End Sub

Public Sub NeoInfB_RemoveIntakePerKg()

    ChangeIntakePerKg True

End Sub

Public Sub NeoInfB_StandardLipid()

    ModRange.SetRangeValue constLipidStand, shtNeoBerInfB.Range(constStandLipid).Value2 + 5000

End Sub

Public Sub NeoInfB_RemoveLipid()

    Dim lngNullVal As Long
    
    ModRange.SetRangeValue constLipidStand, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN_NaCl()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constNaCl, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN_KAct()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constKAct, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN_CaCl()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constCaCl, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN_MgCl()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constMgCl, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN_Solu()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constSolu, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN_Prim()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constPrim, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN_NICUMix()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constNICUmix, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN_SSTB()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constSSTB, lngNullVal

End Sub

Public Sub NeoInfB_RemoveTPN()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constNaCl, lngNullVal
    ModRange.SetRangeValue constKAct, lngNullVal
    ModRange.SetRangeValue constCaCl, lngNullVal
    ModRange.SetRangeValue constMgCl, lngNullVal
    ModRange.SetRangeValue constSolu, lngNullVal
    ModRange.SetRangeValue constPrim, lngNullVal
    ModRange.SetRangeValue constNICUmix, lngNullVal
    ModRange.SetRangeValue constSSTB, lngNullVal

End Sub

Public Sub NeoInfB_StandardTPN()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constNaCl, shtNeoBerInfB.Range(constStandNaCl).Value2 * 10 + lngNullVal
    ModRange.SetRangeValue constKAct, shtNeoBerInfB.Range(constStandKAcet).Value2 * 10 + lngNullVal
    ModRange.SetRangeValue constCaCl, shtNeoBerInfB.Range(constStandCaCl2).Value2 * 10 + lngNullVal
    ModRange.SetRangeValue constMgCl, shtNeoBerInfB.Range(constStandMgCl2).Value2 * 10 + lngNullVal
    ModRange.SetRangeValue constSolu, shtNeoBerInfB.Range(constStandSoluvit).Value2 * 10 + lngNullVal
    ModRange.SetRangeValue constPrim, shtNeoBerInfB.Range(constStandPrimene).Value2 + lngNullVal
    ModRange.SetRangeValue constNICUmix, shtNeoBerInfB.Range(constStandNICUmix).Value2 + lngNullVal
    ModRange.SetRangeValue constSSTB, shtNeoBerInfB.Range(constStandSSTB).Value2 + lngNullVal

End Sub

' Copy paste function cannot be reused because of private clear method
Private Sub ShowMedIVPickList(ByVal strTbl As String, ByVal strRange As String, ByVal intStart As Integer, ByVal intMax As Integer)

    Dim frmPickList As FormNeoMedIVPickList
    Dim colTbl As Collection
    Dim intN As Integer
    Dim strN As String
    Dim intKeuze As Integer
    
    Set colTbl = ModRange.CollectionFromRange(strTbl, intStart)
    
    Set frmPickList = New FormNeoMedIVPickList
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
                    ChangeMedContIV intN, True ' Remove is specific to PedContIV replace with appropriate sub when copy paste
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
                    ChangeMedContIV intN, False
                    Exit For
                End If
            Next intN
        Loop
    
    End If
    
End Sub


Public Sub NeoInfB_ShowMedIVPickList()

    ShowMedIVPickList constTblMedIV, constMedKeuze, 2, constMedIVMax

End Sub

Public Sub NeoInfB_ShowVoedingPickList()

    Dim frmPickList As FormNeoEntPickList
    Dim colVoeding As Collection
    Dim colToevoegMM As Collection
    Dim colToevoegKV As Collection
    Dim intN As Integer
    Dim intVoeding As Integer
    Dim intToevoegMM As Integer
    Dim intToevoegKV As Integer
    
    Dim intMaxLoop As Integer
    Dim intLoop As Integer
    
    intMaxLoop = 100
    
    Set colVoeding = ModRange.CollectionFromRange(constTblVoeding, 2)
    Set colToevoegMM = ModRange.CollectionFromRange(constTblToevoegMM, 2)
    Set colToevoegKV = ModRange.CollectionFromRange(constTblToevoegKV, 2)
    
    Set frmPickList = New FormNeoEntPickList
    frmPickList.LoadVoedingen colVoeding
    frmPickList.LoadToevoegMM colToevoegMM
    frmPickList.LoadToevoegKV colToevoegKV
    
    intVoeding = ModRange.GetRangeValue(constEntVoed, 1)
    If intVoeding > 1 Then frmPickList.SelectVoeding intVoeding
    
    For intN = 1 To constToevoegMMCount
        intToevoegMM = ModRange.GetRangeValue(constEntToev & intN, 1)
        If intToevoegMM > 1 Then frmPickList.SelectToevoegMM intToevoegMM
    Next intN
    
    For intN = 1 To constToevoegKVCount
        intToevoegKV = ModRange.GetRangeValue(constEntToev & intN + 4, 1)
        If intToevoegKV > 1 Then frmPickList.SelectToevoegKV intToevoegKV
    Next intN
    
    frmPickList.Show
    
    If frmPickList.GetAction = vbNullString Then
    
        ' -- Process Voeding
    
        ' First remove nonselected items
        intVoeding = ModRange.GetRangeValue(constEntVoed, 1)
        If intVoeding > 1 Then
            If frmPickList.IsVoedingSelected(intVoeding) Then
                frmPickList.UnselectVoeding (intVoeding)
            Else
                ChangeVoeding True
            End If
        End If
        
        Do While frmPickList.HasSelectedVoedingen()  ' Then add selected items
            intVoeding = ModRange.GetRangeValue(constEnt, 1)
            If intVoeding <= 1 Then
                intVoeding = frmPickList.GetFirstSelectedVoeding(True)
                ModRange.SetRangeValue constEntVoed, intVoeding
            End If
        Loop
    
        ' -- Process Toevoegingen Moedermelk = Var_Neo_InFfB_Toevoeging_ 1 t/m 4
    
        For intN = 1 To constToevoegMMCount         ' First remove nonselected items
            intToevoegMM = ModRange.GetRangeValue(constEntToev & intN, 1)
            If intToevoegMM > 1 Then
                If frmPickList.IsToevoegMMSelected(intToevoegMM) Then
                    frmPickList.UnselectToevoegMM intToevoegMM
                Else
                    ChangeToevoeg True, intN, True  ' Remove toevoeging
                End If
            End If
        Next intN
        
        intLoop = 1
        Do While frmPickList.HasSelectedToevMM()    ' Then add selected items
            For intN = 1 To constToevoegMMCount
                intToevoegMM = ModRange.GetRangeValue(constEntToev & intN, 1)
                If intToevoegMM <= 1 Then
                    intToevoegMM = frmPickList.GetFirstSelectedToevoegMM(True)
                    ModRange.SetRangeValue constEntToev & intN, intToevoegMM
                    Exit For
                End If
            Next intN
            
            intLoop = intLoop + 1
            If intLoop > intMaxLoop Then GoTo NeoInfB_ShowVoedingPickListError
        Loop
    
        ' -- Process Toevoegingen Kunstvoeding = Var_Neo_InFfB_Toevoeging_ 5 t/m 8
    
        For intN = 1 To constToevoegKVCount     ' First remove nonselected items
            intToevoegKV = ModRange.GetRangeValue(constEntToev & intN + 4, 1)
            If intToevoegKV > 1 Then
                If frmPickList.IsToevoegKVSelected(intToevoegKV) Then
                    frmPickList.UnselectToevoegKV intToevoegKV
                Else
                    ChangeToevoeg False, intN, True
                End If
            End If
        Next intN
        
        intLoop = 1
        Do While frmPickList.HasSelectedToevKV()    ' Then add selected items
            For intN = 1 To constToevoegKVCount
                intToevoegKV = ModRange.GetRangeValue(constEntToev & intN + 4, 1)
                If intToevoegKV <= 1 Then
                    intToevoegKV = frmPickList.GetFirstSelectedToevoegKV(True)
                    ModRange.SetRangeValue constEntToev & intN + 4, intToevoegKV
                    Exit For
                End If
            Next intN
            
            intLoop = intLoop + 1
            If intLoop > intMaxLoop Then GoTo NeoInfB_ShowVoedingPickListError
        Loop
    
    End If
    
    Exit Sub
    
NeoInfB_ShowVoedingPickListError:

    ModMessage.ShowMsgBoxError "Zit in een loop"
    ModLog.LogError "Loop error for NeoInfB_ShowVoedingPickList"
    
    Set frmPickList = Nothing
    
End Sub

Private Sub ResetOplVlst(ByVal strOpl, ByVal intOpl As Integer, blnShowWarn)

    If blnShowWarn Then
        ModMessage.ShowMsgBoxInfo "Ongeldige oplossing vloeistof voor dit medicament"
    End If
    
    ModRange.SetRangeValue strOpl, intOpl

End Sub

Public Sub NeoInfB_TestCheckOplVlst(ByVal intN As Integer)

    CheckOplVlst intN, False

End Sub

Private Sub CheckOplVlst(ByVal intN As Integer, ByVal blnShowWarn As Boolean)
    
    Dim strN As String
    Dim intMed As Integer
    Dim intOplVlst As Integer
    Dim intAdvVlst As Integer
    
    strN = ModString.IntNToStrN(intN)
    intMed = ModRange.GetRangeValue(constMedKeuze & strN, 0)
    If intMed > 0 Then
        intAdvVlst = GetMedicamentItemWithIndex(intMed, constAdvOplIndex)
        intOplVlst = ModRange.GetRangeValue(constOplossing & strN, 0)
        'Geen oplossing vloeistof
        If intAdvVlst = 1 And Not intOplVlst = 1 Then
            ResetOplVlst constOplossing & strN, intAdvVlst, blnShowWarn
        End If
        'Oplossing vloeistof is NaCl
        If intAdvVlst = 12 And Not intOplVlst = 12 Then
            ResetOplVlst constOplossing & strN, intAdvVlst, blnShowWarn
        End If
        'Oplossing vloeistof is glucose
        If intAdvVlst > 1 And intAdvVlst < 12 And (intOplVlst = 1 Or intOplVlst > 11) Then
            ResetOplVlst constOplossing & strN, intAdvVlst, blnShowWarn
        End If
                
    End If
    
End Sub

Public Sub NeoInfB_CheckOplVlst_01()

    CheckOplVlst 1, True

End Sub

Public Sub NeoInfB_CheckOplVlst_02()

    CheckOplVlst 2, True

End Sub

Public Sub NeoInfB_CheckOplVlst_03()

    CheckOplVlst 3, True

End Sub

Public Sub NeoInfB_CheckOplVlst_04()

    CheckOplVlst 4, True

End Sub

Public Sub NeoInfB_CheckOplVlst_05()

    CheckOplVlst 5, True

End Sub

Public Sub NeoInfB_CheckOplVlst_06()

    CheckOplVlst 6, True

End Sub

Public Sub NeoInfB_CheckOplVlst_07()

    CheckOplVlst 7, True

End Sub

Public Sub NeoInfB_CheckOplVlst_08()

    CheckOplVlst 8, True

End Sub

Public Sub NeoInfB_CheckOplVlst_09()

    CheckOplVlst 9, True

End Sub

Public Sub NeoInfB_CheckOplVlst_10()

    CheckOplVlst 10, True

End Sub

Public Sub NeoInfB_SetTestDose(ByVal strN As String, ByVal dblDose As Double)

    Dim dblQty As Double
    Dim intMed As Integer
    
    intMed = ModRange.GetRangeValue(constMedKeuze & strN, vbNullString)
    If intMed <= 1 Then Exit Sub
    
    If Not dblDose = 0 Then
        dblQty = CalculateMedicamentQtyByDose(strN, intMed, dblDose)
        ModRange.SetRangeValue constMedSterkte & strN, dblQty * 10
    End If

End Sub

Private Sub SetDose(ByVal intN As Integer)

    Dim strN As String
    Dim frmDose As FormInvoerNumeriek
    Dim intMed As Integer
    Dim strMed As String
    Dim dblDose As Double
    Dim dblQty As Double
    Dim strEenheid As String
    Dim blnNotSetDose
    
    strN = IntNToStrN(intN)
    intMed = ModRange.GetRangeValue(constMedKeuze & strN, vbNullString)
    strMed = ModExcel.Excel_Index(constTblMedIV, intMed, 1)
    
    blnNotSetDose = intMed <= 1
    blnNotSetDose = blnNotSetDose Or ModString.ContainsCaseSensitive(strMed, "EPI")
    blnNotSetDose = blnNotSetDose Or ModString.ContainsCaseInsensitive(strMed, "doxapram")
    If blnNotSetDose Then Exit Sub
    
    strEenheid = ModExcel.Excel_Index(constTblMedIV, intMed, constDoseEenheidIndx)
    dblDose = ModExcel.Excel_Index("Tbl_Neo_BerMedCont", intN, 7)
    
    Set frmDose = New FormInvoerNumeriek
    
    With frmDose
        .lblText.Caption = "Voer dosering in voor " & strMed
        .SetValue vbNullString, "Dose:", dblDose, strEenheid, vbNullString
        
        .Show
        
        If Not .txtWaarde.Value = vbNullString Then
            dblDose = StringToDouble(.txtWaarde.Value)
            dblQty = CalculateMedicamentQtyByDose(strN, intMed, dblDose)
            ModRange.SetRangeValue constMedSterkte & strN, dblQty * 10
        End If
    End With
    
End Sub

Public Sub NeoInfB_SetDose_01()

    SetDose 1

End Sub

Public Sub NeoInfB_SetDose_02()

    SetDose 2

End Sub

Public Sub NeoInfB_SetDose_03()

    SetDose 3

End Sub

Public Sub NeoInfB_SetDose_04()

    SetDose 4

End Sub

Public Sub NeoInfB_SetDose_05()

    SetDose 5

End Sub

Public Sub NeoInfB_SetDose_06()

    SetDose 6

End Sub

Public Sub NeoInfB_SetDose_07()

    SetDose 7

End Sub

Public Sub NeoInfB_SetDose_08()

    SetDose 8

End Sub

Public Sub NeoInfB_SetDose_09()

    SetDose 9

End Sub

Public Sub NeoInfB_SetDose_10()

    SetDose 10

End Sub

Public Function NeoInfB_GetNeoOplVlst() As Range

    Dim objTable As Range
    
    Set objTable = ModRange.GetRange(constTblNeoOpl)
    Set NeoInfB_GetNeoOplVlst = objTable

End Function

Public Function NeoInfB_IsValidContMed() As Boolean

    Dim objRange As Range
    Dim objCell As Range
    Dim blnValid As Boolean
    
    Set objRange = shtNeoBerInfB.Range("Tbl_Neo_ValidContMed")
    
    blnValid = True
    For Each objCell In objRange
        If objCell.Value2 = True Then
            blnValid = False
            Exit For
        End If
    Next

    NeoInfB_IsValidContMed = blnValid

End Function


Public Function NeoInfB_IsValidTPN() As Boolean

    Dim objRange As Range
    Dim objCell As Range
    Dim blnValid As Boolean
    
    Set objRange = shtNeoBerInfB.Range("Tbl_NeoInfB_ValidTPN")
    
    blnValid = True
    For Each objCell In objRange
        If objCell.Value2 = True Then
            blnValid = False
            Exit For
        End If
    Next

    NeoInfB_IsValidTPN = blnValid

End Function


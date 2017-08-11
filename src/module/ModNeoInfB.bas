Attribute VB_Name = "ModNeoInfB"
Option Explicit

Private Const constTblMedIV As String = "Tbl_Neo_MedIV"
Private Const constMedIVMax As Integer = 12

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
Private Const constTPNParEnt As String = "Var_Neo_InfB_TPN_Parenteraal"

Private Const constLipidKeuze As String = "Var_Neo_InfB_TPN_IntralipidSmof"
Private Const constLipidStand As String = "Var_Neo_InfB_TPN_IntraLipid"
Private Const constLipidAdvies As String = "Var_Neo_LipidAdv"

Private Const constNaCl As String = "Var_Neo_InfB_TPN_NaCl"
Private Const constKAct As String = "Var_Neo_InfB_TPN_KCl"
Private Const constCaCl As String = "Var_Neo_InfB_TPN_CaCl2"
Private Const constMgCl As String = "Var_Neo_InfB_TPN_MgCl2"
Private Const constSolu As String = "Var_Neo_InfB_TPN_Soluvit"
Private Const constPrim As String = "Var_Neo_InfB_TPN_Primene"

Private Const constNICUMixStand As String = "Var_Neo_InfB_TPN_NICUMix"
Private Const constNICUMixAdv As String = "Var_Neo_NICUMixAdv"
Private Const constSSTBStand As String = "Var_Neo_InfB_TPN_SSTB"
Private Const constSSTBAdv As String = "Var_Neo_SSTBAdv"

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

Public Function IsEpiduraal(ByVal strText As String) As Boolean

    IsEpiduraal = ModString.ContainsCaseInsensitive(strText, "epiduraal")

End Function

Private Function Is1700() As Boolean

    Is1700 = shtNeoBerInfB.Range(constInfbVersie) = const1700InfB

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

    ModProgress.StartProgress "Overzetten Actueel naar 17:00 uur"

    CopyCurrentInfVarToData True

    ModProgress.SetJobPercentage "Voeding", 3, 1
    CopyRangeNamesToRangeNames NeoInfB_GetVoedingItems(), ChangeTo1700(NeoInfB_GetVoedingItems())
    ModProgress.SetJobPercentage "Continue IV", 3, 2
    CopyRangeNamesToRangeNames NeoInfB_GetIVAfsprItems(), ChangeTo1700(NeoInfB_GetIVAfsprItems())
    ModProgress.SetJobPercentage "TPN", 3, 3
    CopyRangeNamesToRangeNames NeoInfB_GetTPNItems(), ChangeTo1700(NeoInfB_GetTPNItems())
    
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

Public Sub NeoInfB_Copy1700ToAct(ByVal blnAlles As Boolean, ByVal blnVoeding As Boolean, ByVal blnContMed As Boolean, ByVal blnTPN As Boolean)
    
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

    Set frmCopy1700 = Nothing
    
    NeoInfB_SelectInfB bln1700, True
    
End Sub

Private Sub test()
    
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
    
    Set objTblMed = Nothing

End Function

Private Sub TestGetMinimumAdvice()

    MsgBox GetMedicamentItemWithIndex(2, constAdvMinIndex)

End Sub


Private Function GetMedicamentMinQty(ByVal intMed As Integer) As Double

    Dim dblAdvMin As Double
    Dim dblFactor As Double
    Dim dblWeight As Double
    Dim dblOplQty As Double
    Dim dblStand As Double
    Dim dblQty As Double
    Dim dblMaxConc As Double
    Dim strMed As String
    
    dblAdvMin = GetMedicamentItemWithIndex(intMed, constAdvMinIndex)
    dblFactor = GetMedicamentItemWithIndex(intMed, constFactorIndex)
    dblOplQty = GetMedicamentItemWithIndex(intMed, constDefHoevIndex)
    dblStand = GetMedicamentItemWithIndex(intMed, constDefStandIndex)
    dblWeight = ModPatient.GetGewichtFromRange()
    dblMaxConc = GetMedicamentItemWithIndex(intMed, 10)
    strMed = GetMedicamentItemWithIndex(intMed, 1)
    
    ' Medicatie doxapram puur geen oplosvolume
    If ModString.ContainsCaseInsensitive(strMed, "doxapram") Then
        dblQty = dblMaxConc * dblOplQty
        
    ' dblAdvMin = dblStand * dblFactor * (dblQty / dblOplQty) / dblWeight
    ' dblQty = dblAdvMin * dblOplQty * dblWeight / (dblStand * dblFactor)
    ElseIf dblStand * dblFactor > 0 Then
        dblQty = dblAdvMin * dblOplQty * dblWeight / (dblStand * dblFactor)
        dblQty = IIf(dblQty / dblOplQty > dblMaxConc, dblMaxConc * dblOplQty, dblQty)
    Else
        dblQty = 0
    End If
    
    GetMedicamentMinQty = ModString.FixPrecision(dblQty, 1)

End Function

Private Sub TestGetMedicamentMinQty()

    MsgBox GetMedicamentMinQty(9)

End Sub

Private Sub ChangeMedContIV(ByVal intRegel As Integer, ByVal blnRemove As Boolean)

    Dim strMedVar As String
    Dim varMedIndex As Variant
    Dim strMedName As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim strStand As String
    Dim strExtra As String
    Dim strRegel As String
      
    Dim dblOplQty As Double
    Dim dblStand As Double
    
    Dim objTblMed As Range
    
    Set objTblMed = ModRange.GetRange(constTblMedIV)
    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    
    strMedVar = constMedKeuze & strRegel
    strMedSterkte = constMedSterkte & strRegel
    strOplHoev = constOplHoev & strRegel
    
    strOplossing = constOplossing & strRegel
    strStand = constStand & strRegel
    strExtra = constExtra & IIf(intRegel + 1 < 10, "0" & (intRegel + 1), intRegel + 1)

    varMedIndex = IIf(blnRemove, 1, ModRange.GetRangeValue(strMedVar, vbNullString))
    
    If blnRemove Then ModRange.SetRangeValue strMedVar, 1
    
    If blnRemove Then
        ModRange.SetRangeValue strMedSterkte, 0
    Else
        strMedName = ModMedicatie.GetNeoMedContIVName(varMedIndex)
        If IsEpiduraal(strMedName) Then
            ModRange.SetRangeValue strMedSterkte, ModMedicatie.Medicatie_CalcEpiQty(ModPatient.GetGewichtFromRange()) * 10
        Else
            ModRange.SetRangeValue strMedSterkte, GetMedicamentMinQty(varMedIndex) * 10
        End If
    End If
    
    dblOplQty = GetMedicamentItemWithIndex(varMedIndex, constDefHoevIndex)
    dblStand = GetMedicamentItemWithIndex(varMedIndex, constDefStandIndex) * 10
    
    ModRange.SetRangeValue strOplHoev, IIf(blnRemove, 0, dblOplQty)
    ModRange.SetRangeValue strStand, IIf(blnRemove, 0, dblStand)
    ModRange.SetRangeValue strExtra, False
    
    ModRange.SetRangeValue strOplossing, Application.VLookup(objTblMed.Cells(varMedIndex, 1), objTblMed, constAdvOplIndex, False)
    If Not IsNumeric(ModRange.GetRangeValue(strOplossing, vbNullString)) Then
        ModRange.SetRangeValue strOplossing, 1
    End If
    
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

Private Sub MedSterkte(ByVal intN As Integer)

    Dim frmInvoer As FormInvoerNumeriek
    Dim strSterkte As String
    Dim strN As String
    Dim intMedKeuze As Integer
    Dim strUnit As String
    
    Dim objTblMed As Range
    
    strN = IIf(intN < 10, "0" & intN, intN)
    intMedKeuze = ModRange.GetRangeValue(constMedKeuze & strN, 1)
    If intMedKeuze <= 1 Then Exit Sub
    
    Set objTblMed = Range(constTblMedIV)
    strUnit = Application.VLookup(objTblMed.Cells(intMedKeuze, 1), objTblMed, constUnitIndex, False)
    strSterkte = constMedSterkte & strN
    
    Set frmInvoer = New FormInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament " & intN
        .lblParameter = "Sterkte"
        .lblEenheid = strUnit
        .txtWaarde = ModRange.GetRangeValue(strSterkte, 0) / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            ModRange.SetRangeValue strSterkte, .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

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
    
    If blnRemove Then ModRange.SetRangeValue constOplossing & intRegel, 1
    ModRange.SetRangeValue strStand, 0
    ModRange.SetRangeValue strExtra, vbNullString
    
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

    ModRange.SetRangeValue "_DagKeuze", IIf(ModRange.GetRangeValue("Dag", 0) < 4, 1, 2)
    ModRange.SetRangeValue "_IntakePerKg", 5000
    ModRange.SetRangeValue "_IntraLipid", 5000
    ModRange.SetRangeValue "_NaCl", 5000
    ModRange.SetRangeValue "_KCl", 5000
    ModRange.SetRangeValue "_CaCl2", 5000
    ModRange.SetRangeValue "_MgCl2", 5000
    ModRange.SetRangeValue "_SoluVit", 5000
    ModRange.SetRangeValue "_Primene", 5000
    ModRange.SetRangeValue "_NICUMix", 5000
    ModRange.SetRangeValue "_SSTB", 5000
    
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
    
    Set frmInvoer = Nothing

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
        ModRange.SetRangeValue constIntakePerKg, 5000
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

Private Function CalcAdviceNullValue(ByVal strRange As String) As Long

    CalcAdviceNullValue = 5000 - ModRange.GetRangeValue(strRange, 0)

End Function

Public Sub NeoInfB_StandardLipid()

    ModRange.SetRangeValue constLipidStand, 5000

End Sub

Public Sub NeoInfB_RemoveLipid()

    Dim lngNullVal As Long
    
    lngNullVal = CalcAdviceNullValue(constLipidAdvies)
    ModRange.SetRangeValue constLipidStand, lngNullVal

End Sub

' ToDo Implement and assign
Public Sub NeoInfB_RemoveTPN()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constNaCl, lngNullVal
    ModRange.SetRangeValue constKAct, lngNullVal
    ModRange.SetRangeValue constCaCl, lngNullVal
    ModRange.SetRangeValue constMgCl, lngNullVal
    ModRange.SetRangeValue constSolu, lngNullVal
    ModRange.SetRangeValue constPrim, lngNullVal
    ModRange.SetRangeValue constNICUMixStand, CalcAdviceNullValue(constNICUMixAdv)
    ModRange.SetRangeValue constSSTBStand, CalcAdviceNullValue(constSSTBAdv)

End Sub

' ToDo Implement and assign
Public Sub NeoInfB_StandardTPN()

    Dim lngNullVal As Long
        
    lngNullVal = 5000
    ModRange.SetRangeValue constNaCl, lngNullVal
    ModRange.SetRangeValue constKAct, lngNullVal
    ModRange.SetRangeValue constNICUMixStand, lngNullVal
    ModRange.SetRangeValue constSSTBStand, lngNullVal

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
        strN = IIf(intMax > 9, IIf(intN < 10, "0" & intN, intN), intN)
        intKeuze = ModRange.GetRangeValue(strRange & strN, 1)
        If intKeuze > 1 Then frmPickList.SelectMedicament intKeuze
    Next intN
    
    frmPickList.Show
    
    If frmPickList.GetAction = vbNullString Then
    
        For intN = 1 To intMax                 ' First remove nonselected items
            strN = IIf(intN < 10, "0" & intN, intN)
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
                strN = IIf(intN < 10, "0" & intN, intN)
                intKeuze = ModRange.GetRangeValue(strRange & strN, 1)
                If intKeuze <= 1 Then
                    intKeuze = frmPickList.GetFirstSelectedMedicament(True)
                    ModRange.SetRangeValue strRange & strN, intKeuze
                    Exit For
                End If
            Next intN
        Loop
    
    End If
    
    Set frmPickList = Nothing

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
    
    Set frmPickList = Nothing
    
    Exit Sub
    
NeoInfB_ShowVoedingPickListError:

    ModMessage.ShowMsgBoxError "Zit in een loop"
    ModLog.LogError "Loop error for NeoInfB_ShowVoedingPickList"
    
    Set frmPickList = Nothing
    
End Sub


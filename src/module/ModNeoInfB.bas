Attribute VB_Name = "ModNeoInfB"
Option Explicit

Private Const constTblMedIV As String = "tbl_Neo_MedIV"

Private Const constActInfB = "Actuele Afspraken"
Private Const const1700InfB As String = "17.00 uur Afspraken"
Private Const constInfbVersie = "B2"

Private Const constInfBVa = "Var_Neo_InfB"
Private Const constInfBDataAct = "_Neo_InfB"
Private Const constInfBData1700 = "_Neo_1700"

Private Function Is1700() As Boolean

    Is1700 = shtNeoBerInfB.Range(constInfbVersie) = const1700InfB

End Function

Private Sub CopyVarData(ByVal bln1700 As Boolean, ByVal blnToVar As Boolean, blnShowProgress As Boolean)

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

Private Sub TestCopyVarData()

    CopyVarData True, True, True

End Sub

Public Sub NeoInfB_SelectInfB(ByVal bln1700 As Boolean)

    If bln1700 And Is1700() Then                ' InfB is same as 1700
        ModSheet.GoToSheet shtNeoGuiInfB, "A9"
        
    ElseIf Not bln1700 And Not Is1700() Then     ' InfB is same as act
        ModSheet.GoToSheet shtNeoGuiInfB, "A9"
        
    ElseIf bln1700 And Not Is1700() Then         ' Infb is currently act
        ModProgress.StartProgress "Infuus brief klaar maken"
        
        CopyVarData False, False, True   ' First copy var data to act data
        CopyVarData True, True, True   ' Then copy 1700 data to var data
        shtNeoBerInfB.Range(constInfbVersie).Value2 = const1700InfB
        
        ModProgress.FinishProgress
        
        ModSheet.GoToSheet shtNeoGuiInfB, "A9"
    
    ElseIf Not bln1700 And Is1700() Then         ' Infb is currently 1700
        ModProgress.StartProgress "Infuus brief klaar maken"
        
        CopyVarData True, False, True ' First copy var data to act data
        CopyVarData False, True, True   ' Then copy act data to var data
        shtNeoBerInfB.Range(constInfbVersie).Value2 = constActInfB
        
        ModProgress.FinishProgress
        
        ModSheet.GoToSheet shtNeoGuiInfB, "A9"
    
    End If

End Sub

' ToDo: Add comment
Public Sub NeoInfB_CopyActTo1700()

    CopyRangeNamesToRangeNames NeoInfB_GetVoedingItems(), ChangeTo1700(NeoInfB_GetVoedingItems())
    CopyRangeNamesToRangeNames NeoInfB_GetIVAfsprItems(), ChangeTo1700(NeoInfB_GetIVAfsprItems())
    CopyRangeNamesToRangeNames NeoInfB_GetTPNItems(), ChangeTo1700(NeoInfB_GetTPNItems())
    
End Sub

Public Function NeoInfB_GetVoedingItems() As String()

    Dim arrItems() As String
    ReDim arrItems(0)
        
    arrItems(0) = "_InfB_Voeding"
    AddItemsToArray arrItems, "_Frequentie", 1, 2
    AddItemsToArray arrItems, "_Fototherapie", 1, 1
    AddItemsToArray arrItems, "_Parenteraal", 1, 1
    AddItemsToArray arrItems, "_Toevoeging", 1, 8
    AddItemsToArray arrItems, "_PercentageKeuze", 0, 8
    AddItemsToArray arrItems, "_IntakePerKg", 1, 1
    AddItemsToArray arrItems, "_Extra", 1, 1
    
    NeoInfB_GetVoedingItems = arrItems

End Function

Public Function NeoInfB_GetIVAfsprItems() As String()

    Dim arrItems() As String
    ReDim arrItems(0)
        
    arrItems(0) = "_InfB_ArtLijn"
    AddItemsToArray arrItems, "_Medicament", 1, 9
    AddItemsToArray arrItems, "_MedSterkte", 1, 9
    AddItemsToArray arrItems, "_OplHoev", 1, 9
    AddItemsToArray arrItems, "_Oplossing", 1, 12
    AddItemsToArray arrItems, "_Stand", 1, 12
    AddItemsToArray arrItems, "_Extra", 1, 12
    AddItemsToArray arrItems, "_MedTekst", 1, 2
    
    NeoInfB_GetIVAfsprItems = arrItems

End Function

Public Function NeoInfB_GetTPNItems() As String()

    Dim arrItems() As String
    ReDim arrItems(0)
    
    arrItems(0) = "_InfB_Parenteraal"
    AddItemsToArray arrItems, "_IntraLipid", 1, 1
    AddItemsToArray arrItems, "_DagKeuze", 1, 1
    
    AddItemsToArray arrItems, "_NaCl", 1, 1
    AddItemsToArray arrItems, "_KCl", 1, 1
    AddItemsToArray arrItems, "_CaCl2", 1, 1
    AddItemsToArray arrItems, "_MgCl2", 1, 1
    AddItemsToArray arrItems, "_SoluVit", 1, 1
    AddItemsToArray arrItems, "_Primene", 1, 1
    AddItemsToArray arrItems, "_NICUMix", 1, 1
    AddItemsToArray arrItems, "_SSTB", 1, 1
    AddItemsToArray arrItems, "_GlucSterkte", 1, 1
    
    NeoInfB_GetTPNItems = arrItems
    
End Function

Private Sub AddItemsToArray(ByRef arrItems() As String, ByVal strItem As String, ByVal intStart As Integer, ByVal intStop As Boolean)

    Dim intC As Integer
    Dim intU As Integer
    Dim strInfB As String
    Dim strN As String
    
    strInfB = "_Neo_InfB"
    
    If intStart = intStop Then
        ModArray.AddItemToStringArray arrItems, strItem
    Else
        intU = UBound(arrItems)
        ReDim Preserve arrItems(0 To intU + intStop - intStart + 1)
        
        For intC = intStart To intStop
            intU = intU + 1
            strN = IIf(intStop > 9 And intC < 10, "0" & intC, intC)
            arrItems(intU) = strInfB & strItem & "_" & strN
        Next intC
    End If
    
End Sub

Private Function ChangeTo1700(ByRef arrItems() As String) As String()
    
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
    
    NeoInfB_SelectInfB bln1700
    
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

Private Sub ChangeMedContIV(ByVal intRegel As Integer, ByVal bln1700 As Boolean)

    Dim strMedicament As String
    Dim varMedicament As Variant
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim strStand As String
    Dim strExtra As String
    Dim strRegel As String
    
    Dim objTblMed As Range
    
    Set objTblMed = Range(constTblMedIV)
    strRegel = IIf(intRegel < 10, "0" & intRegel, intRegel)
    
    strMedicament = IIf(bln1700, "_Neo_1700_Medicament_" & intRegel, "_Neo_InfB_Medicament_" & intRegel)
    strMedSterkte = IIf(bln1700, "_Neo_1700_MedSterkte_" & intRegel, "_Neo_InfB_MedSterkte_" & intRegel)
    strOplHoev = IIf(bln1700, "_Neo_1700_OplHoev_" & intRegel, "_Neo_InfB_OplHoev_" & intRegel)
    
    strOplossing = IIf(bln1700, "_Neo_1700_Oplossing_" & strRegel, "_Neo_InfB_Oplossing_" & strRegel)
    strStand = IIf(bln1700, "_Neo_1700_Stand_" & strRegel, "_Neo_InfB_Stand_" & strRegel)
    strExtra = IIf(bln1700, "_Neo_1700_VochtExtra_" & strRegel, "_Neo_InfB_VochtExtra_" & strRegel)

    varMedicament = ModRange.GetRangeValue(strMedicament, vbNullString)
    
    ModRange.SetRangeValue strMedSterkte, 0
    ModRange.SetRangeValue strOplHoev, 0
    ModRange.SetRangeValue strStand, 0
    ModRange.SetRangeValue strExtra, vbNullString
    
    ModRange.SetRangeValue strOplossing, Application.VLookup(objTblMed.Cells(varMedicament, 1), objTblMed, 10, False)
    If Not IsNumeric(ModRange.GetRangeValue(strOplossing, vbNullString)) Then
        ModRange.SetRangeValue strOplossing, 1
    End If
    
End Sub

Public Sub NeoInfB_ChangeMedContIV_1()

    ChangeMedContIV 1, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_2()
    
    ChangeMedContIV 2, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_3()
    
    ChangeMedContIV 3, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_4()
    
    ChangeMedContIV 4, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_5()
    
    ChangeMedContIV 5, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_6()
    
    ChangeMedContIV 6, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_7()
    
    ChangeMedContIV 7, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_8()
    
    ChangeMedContIV 8, False

End Sub

Public Sub NeoInfB_ChangeMedContIV_9()
    
    ChangeMedContIV 9, False

End Sub

Private Sub MedSterkte(ByVal intRegel As Integer, ByVal bln1700 As Boolean)

    Dim frmInvoer As FormInvoerNumeriek
    Dim strSterkte As String
    
    Set frmInvoer = New FormInvoerNumeriek
    strSterkte = IIf(bln1700, "_Neo_1700_MedSterkte_" & intRegel, "_Neo_InfB_MedSterkte_" & intRegel)
    
    With frmInvoer
        .Caption = "Medicament " & intRegel
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = ModRange.GetRangeValue(strSterkte, 0) / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            ModRange.SetRangeValue strSterkte, .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub NeoInfB_MedConc_1()

    MedSterkte 1, False

End Sub

Public Sub NeoInfB_MedConc_2()

    MedSterkte 2, False

End Sub

Public Sub NeoInfB_MedConc_3()

    MedSterkte 3, False

End Sub

Public Sub NeoInfB_MedConc_4()

    MedSterkte 4, False

End Sub

Public Sub NeoInfB_MedConc_5()

        MedSterkte 5, False

End Sub

Public Sub NeoInfB_MedConc_6()

    MedSterkte 6, False

End Sub

Public Sub NeoInfB_MedConc_7()

    MedSterkte 7, False

End Sub

Public Sub NeoInfB_MedConc_8()

    MedSterkte 8, False

End Sub

Public Sub NeoInfB_MedConc_9()

    MedSterkte 9, False

End Sub


Private Sub RemoveIV(ByVal intRegel As Integer)

    Dim strStand As String
    Dim strExtra As String

    strStand = "_Neo_InfB_Stand_" & intRegel
    strExtra = "_Neo_InfB_VochtExtra_" & intRegel + 1
    
    ModRange.SetRangeValue strStand, 0
    ModRange.SetRangeValue strExtra, vbNullString
    
End Sub

Public Sub NeoInfB_RemoveIV_10()
    
    RemoveIV 10

End Sub

Public Sub NeoInfB_RemoveIV_11()
    
    RemoveIV 11

End Sub

Public Sub NeoInfB_RemoveIV_12()
    
    RemoveIV 12

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

    EnterText "Voer tekst in ...", "Tekst voor medicatie 13", "_MedTekst_1"

End Sub

Public Sub NeoInfB_MedText_2()

    EnterText "Voer tekst in ...", "Tekst voor medicatie 14", "_MedTekst_2"
    
End Sub

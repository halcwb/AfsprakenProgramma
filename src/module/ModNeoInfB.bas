Attribute VB_Name = "ModNeoInfB"
Option Explicit

' ToDo: Add comment
Public Sub NeoInfB_CopyAfspraken()

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

Private Sub AddItemsToArray(arrItems() As String, strItem As String, intStart As Integer, intStop)

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

Public Sub NeoInfB_CopyInfB(blnAlles As Boolean, blnVoeding As Boolean, blnContMed As Boolean, blnTPN As Boolean)
    
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

Public Sub CopyToActueel()

    Dim frmCopy1700 As New FormCopy1700
    
    frmCopy1700.Show

    Set frmCopy1700 = Nothing
    
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

Private Sub RemoveContIV(intRegel As Integer, bln1700 As Boolean)

    Dim strMedicament As String
    Dim varMedicament As Variant
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim strStand As String
    Dim strExtra As String
    Dim strRegel As String
    
    Dim objTblMed As Range
    
    Set objTblMed = Range(ModConst.CONST_RANGE_NEOMED)
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

Public Sub NeoInfB_RemoveContIV_1()

    RemoveContIV 1, False

End Sub

Public Sub NeoInfB_RemoveContIV_2()
    
    RemoveContIV 2, False

End Sub

Public Sub NeoInfB_RemoveContIV_3()
    
    RemoveContIV 3, False

End Sub

Public Sub NeoInfB_RemoveContIV_4()
    
    RemoveContIV 4, False

End Sub

Public Sub NeoInfB_RemoveContIV_5()
    
    RemoveContIV 5, False

End Sub

Public Sub NeoInfB_RemoveContIV_6()
    
    RemoveContIV 6, False

End Sub

Public Sub NeoInfB_RemoveContIV_7()
    
    RemoveContIV 7, False

End Sub

Public Sub NeoInfB_RemoveContIV_8()
    
    RemoveContIV 8, False

End Sub

Public Sub NeoInfB_RemoveContIV_9()
    
    RemoveContIV 9, False

End Sub

Public Sub NeoInfB_RemoveContIV1700_1()

    RemoveContIV 1, True

End Sub

Public Sub NeoInfB_RemoveContIV1700_2()
    
    RemoveContIV 2, True

End Sub

Public Sub NeoInfB_RemoveContIV1700_3()
    
    RemoveContIV 3, True

End Sub

Public Sub NeoInfB_RemoveContIV1700_4()
    
    RemoveContIV 4, True

End Sub

Public Sub NeoInfB_RemoveContIV1700_5()
    
    RemoveContIV 5, True

End Sub

Public Sub NeoInfB_RemoveContIV1700_6()
    
    RemoveContIV 6, True

End Sub

Public Sub NeoInfB_RemoveContIV1700_7()
    
    RemoveContIV 7, True

End Sub

Public Sub NeoInfB_RemoveContIV1700_8()
    
    RemoveContIV 8, True

End Sub

Public Sub NeoInfB_RemoveContIV1700_9()
    
    RemoveContIV 9, True

End Sub

Private Sub MedSterkte(intRegel As Integer, bln1700 As Boolean)

    Dim frmInvoer As New FormInvoerNumeriek
    Dim strSterkte As String
    
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

Public Sub NeoInfB_MedConc1700_1()
    
    MedSterkte 1, True

End Sub

Public Sub NeoInfB_MedConc1700_2()
    
    MedSterkte 2, True

End Sub

Public Sub NeoInfB_MedConc1700_3()
    
    MedSterkte 3, True

End Sub

Public Sub NeoInfB_MedConc1700_4()
    
    MedSterkte 4, True

End Sub

Public Sub NeoInfB_MedConc1700_5()
    
    MedSterkte 5, True

End Sub

Public Sub NeoInfB_MedConc1700_6()
    
    MedSterkte 6, True

End Sub

Public Sub NeoInfB_MedConc1700_7()
    
    MedSterkte 7, True

End Sub

Public Sub NeoInfB_MedConc1700_8()
    
    MedSterkte 8, True

End Sub

Public Sub NeoInfB_MedConc1700_9()
    
    MedSterkte 9, True

End Sub

Private Sub RemoveIV(intRegel As Integer)

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

Private Sub EnterText(strCaption As String, strName As String, strRange As String)

    Dim frmInvoer As New FormTekstInvoer
    
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


Public Sub NeoInfB_MedText1700_1()

    EnterText "Voer tekst in ...", "Tekst voor medicatie 13", "_MedTekst1700_1"
    
End Sub

Public Sub NeoInfB_MedText1700_2()

    EnterText "Voer tekst in ...", "Tekst voor medicatie 14", "_MedTekst1700_2"
    
End Sub

Attribute VB_Name = "ModInfuusbrief"
Option Explicit

' ToDo: Add comment
Public Sub CopyAfspraken()

    CopyRangeNamesToRangeNames GetVoedingItems(), Get1700Items(GetVoedingItems())
    CopyRangeNamesToRangeNames GetIVAfsprItems(), Get1700Items(GetIVAfsprItems())
    CopyRangeNamesToRangeNames GetTPNItems(), Get1700Items(GetTPNItems())
    
End Sub

Public Function GetVoedingItems() As String()

    Dim arrItems() As String
    ReDim arrItems(0)
        
    arrItems(0) = "_Voeding"
    AddItemsToArray arrItems, "_Frequentie", 1, 2
    AddItemsToArray arrItems, "_Fototherapie", 1, 1
    AddItemsToArray arrItems, "_Parenteraal", 1, 1
    AddItemsToArray arrItems, "_Toevoeging", 1, 8
    AddItemsToArray arrItems, "_PercentageKeuze", 0, 8
    AddItemsToArray arrItems, "_IntakePerKg", 1, 1
    AddItemsToArray arrItems, "_Extra", 1, 1
    
    GetVoedingItems = arrItems

End Function

Public Function GetIVAfsprItems() As String()

    Dim arrItems() As String
    ReDim arrItems(0)
        
    arrItems(0) = "_ArtLijn"
    AddItemsToArray arrItems, "_Medicament", 1, 9
    AddItemsToArray arrItems, "_MedSterkte", 1, 9
    AddItemsToArray arrItems, "_OplHoev", 1, 9
    AddItemsToArray arrItems, "_Oplossing", 1, 12
    AddItemsToArray arrItems, "_Stand", 1, 12
    AddItemsToArray arrItems, "_Extra", 1, 12
    AddItemsToArray arrItems, "_MedTekst", 1, 2
    
    GetIVAfsprItems = arrItems

End Function

Public Function GetTPNItems() As String()

    Dim arrItems() As String
    ReDim arrItems(0)
    
    arrItems(0) = "_Parenteraal"
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
    
    GetTPNItems = arrItems
    
End Function

Public Sub AddItemToArray(arrItems() As String, strItem As String)

    Dim intU As Integer
    
    intU = UBound(arrItems) + 1
    ReDim Preserve arrItems(0 To intU)
    
    arrItems(intU) = strItem

End Sub

Public Sub AddItemsToArray(arrItems() As String, strItem As String, intStart As Integer, intStop)

    Dim intC As Integer
    Dim intU As Integer
    
    If intStart = intStop Then
        AddItemToArray arrItems, strItem
    Else
        intU = UBound(arrItems)
        ReDim Preserve arrItems(0 To intU + intStop - intStart + 1)
        
        For intC = intStart To intStop
            intU = intU + 1
            arrItems(intU) = strItem & "_" & intC
        Next intC
    End If
    
End Sub

Public Function Get1700Items(arrItems() As String) As String()
    
    Dim arr1700Items() As String
    Dim varItem As Variant
    Dim arrSplit() As String
    Dim strAfspr, strAfspr1700 As String
    Dim strNum As String
    Dim intN As Integer
    
    ReDim arr1700Items(UBound(arrItems))
    
    For Each varItem In arrItems
        arrSplit = Split(varItem, "_")
        strAfspr = arrSplit(1)
        
        If UBound(arrSplit) = 2 Then
            strNum = arrSplit(2)
        Else
            strNum = ""
        End If
        
        If strNum = vbNullString Then
            strAfspr1700 = "_" & strAfspr & "1700"
        Else
            strAfspr1700 = "_" & strAfspr & "1700" & "_" & strNum
        End If
        
        If strAfspr1700 = vbNullString Then Err.Raise 1004, "Get1700Items", "Afspraken 1700 cannot be empty string"
        
        arr1700Items(intN) = strAfspr1700
        intN = intN + 1
        
    Next varItem
    
    Get1700Items = arr1700Items

End Function

Public Sub CopyRangeNamesToRangeNames(arrFrom() As String, arrTo() As String)
    
    Dim intN As Integer
    
    For intN = 0 To UBound(arrFrom)
        ModRange.SetRangeValue arrTo(intN), ModRange.GetRangeValue(arrFrom(intN), vbNullString)
    Next intN
    
End Sub

Public Sub AfsprakenOvernemen(blnAlles As Boolean, blnVoeding As Boolean, blnContMed As Boolean, blnTPN As Boolean)
    
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
    
    arrTo = GetVoedingItems()
    arrFrom = Get1700Items(arrTo)
    
    CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

Private Sub ContMedOvernemen()
    
    Dim arrTo() As String
    Dim arrFrom() As String
    
    arrTo = GetIVAfsprItems()
    arrFrom = Get1700Items(arrTo)
    
    CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

Private Sub TPNOvernemen()
    
    Dim arrTo() As String
    Dim arrFrom() As String
    
    arrTo = GetTPNItems()
    arrFrom = Get1700Items(arrTo)
    
    CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

Public Sub CopyToActueel()

    Dim frmCopy1700 As New FormCopy1700
    ' Show selectie form
    frmCopy1700.Show
    ' Copy by block

    Set frmCopy1700 = Nothing
    
End Sub

Private Sub Test()
    
    Dim varItem As Variant
    Dim arr1700Items() As String
    Dim intN As Integer
    
    arr1700Items = Get1700Items(GetIVAfsprItems())
    For Each varItem In GetIVAfsprItems()
        MsgBox varItem & ", " & arr1700Items(intN)
        intN = intN + 1
    Next varItem

End Sub

Sub VerwijderContInfuus(intRegel As Integer, bln1700 As Boolean)

    Dim strMedicament As String
    Dim varMedicament As Variant
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim strStand As String
    Dim strExtra As String
    
    Dim objMeds As Range
    
    Set objMeds = Range(ModConst.CONST_RANGE_NEOMED)
    
    If bln1700 Then
        strMedicament = "_Medicament1700_" & intRegel
        strMedSterkte = "_MedSterkte1700_" & intRegel
        strOplHoev = "_OplHoev1700_" & intRegel
        strOplossing = "_Oplossing1700_" & intRegel
        strStand = "_Stand1700_" & intRegel
        strExtra = "_Extra1700_" & intRegel
    Else
        strMedicament = "_Medicament_" & intRegel
        strMedSterkte = "_MedSterkte_" & intRegel
        strOplHoev = "_OplHoev_" & intRegel
        strOplossing = "_Oplossing_" & intRegel
        strStand = "_Stand_" & intRegel
        strExtra = "_Extra_" & intRegel
    End If

    varMedicament = ModRange.GetRangeValue(strMedicament, vbNullString)
    
    ModRange.SetRangeValue strMedSterkte, 0
    ModRange.SetRangeValue strOplHoev, 0
    ModRange.SetRangeValue strStand, 0
    ModRange.SetRangeValue strExtra, 0
    
    ModRange.SetRangeValue strOplossing, Application.VLookup(objMeds.Cells(varMedicament, 1), objMeds, 10, False)
    If Not IsNumeric(ModRange.GetRangeValue(strOplossing, vbNullString)) Then
        ModRange.SetRangeValue strOplossing, 1
    End If
    
End Sub

Sub VerwijderContInfuus1700_1()
    VerwijderContInfuus 1, True
End Sub

Sub VerwijderContInfuus1700_2()
    VerwijderContInfuus 2, True
End Sub

Sub VerwijderContInfuus1700_3()
    VerwijderContInfuus 3, True
End Sub

Sub VerwijderContInfuus1700_4()
    VerwijderContInfuus 4, True
End Sub

Sub VerwijderContInfuus1700_5()
    VerwijderContInfuus 5, True
End Sub

Sub VerwijderContInfuus1700_6()
    VerwijderContInfuus 6, True
End Sub

Sub VerwijderContInfuus1700_7()
    VerwijderContInfuus 7, True
End Sub

Sub VerwijderContInfuus1700_8()
    VerwijderContInfuus 8, True
End Sub

Sub VerwijderContInfuus1700_9()
    VerwijderContInfuus 9, True
End Sub

Private Sub MedSterkte(intRegel As Integer, bln1700 As Boolean)

    Dim frmInvoer As New FormInvoerNumeriek
    Dim strSterkte As String
    
    If bln1700 Then
        strSterkte = "_MedSterkte1700_" & intRegel
    Else
        strSterkte = "_MedSterkte_" & intRegel
    End If
    
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

Public Sub MedSterkte1700_1()
    MedSterkte 1, True
End Sub

Public Sub MedSterkte1700_2()
    MedSterkte 2, True
End Sub

Public Sub MedSterkte1700_3()
    MedSterkte 3, True
End Sub

Public Sub MedSterkte1700_4()
    MedSterkte 4, True
End Sub

Public Sub MedSterkte1700_5()
    MedSterkte 5, True
End Sub

Public Sub MedSterkte1700_6()
    MedSterkte 6, True
End Sub

Public Sub MedSterkte1700_7()
    MedSterkte 7, True
End Sub

Public Sub MedSterkte1700_8()
    MedSterkte 8, True
End Sub

Public Sub MedSterkte1700_9()
    MedSterkte 9, True
End Sub

Public Sub Afspraken_Vervolgkeuzelijst1_BijWijzigen()

    VerwijderContInfuus 1, False

End Sub

Public Sub Vervolgkeuzelijst2_BijWijzigen()
    
    VerwijderContInfuus 2, False

End Sub

Public Sub Vervolgkeuzelijst3_BijWijzigen()
    
    VerwijderContInfuus 3, False

End Sub

Public Sub Vervolgkeuzelijst4_BijWijzigen()
    
    VerwijderContInfuus 4, False

End Sub

Public Sub Vervolgkeuzelijst5_BijWijzigen()
    
    VerwijderContInfuus 5, False

End Sub

Public Sub Vervolgkeuzelijst6_BijWijzigen()
    
    VerwijderContInfuus 6, False

End Sub

Public Sub Vervolgkeuzelijst7_BijWijzigen()
    
    VerwijderContInfuus 7, False

End Sub

Public Sub Vervolgkeuzelijst67_BijWijzigen()
    
    VerwijderContInfuus 8, False

End Sub

Public Sub Vervolgkeuzelijst97_BijWijzigen()
    
    VerwijderContInfuus 9, False

End Sub

Private Sub VerwijderZijlijn(intRegel As Integer)

    Dim strStand As String
    Dim strExtra As String

    strStand = "_Stand_" & intRegel
    strExtra = "_Extra_" & intRegel + 1
    
    ModRange.SetRangeValue strStand, 0
    ModRange.SetRangeValue strExtra, 0
    
End Sub

Public Sub Vervolgkeuzelijst101_BijWijzigen()
    
    VerwijderZijlijn 10

End Sub

Public Sub Vervolgkeuzelijst104_BijWijzigen()
    
    VerwijderZijlijn 11

End Sub

Public Sub Vervolgkeuzelijst108_BijWijzigen()
    
    VerwijderZijlijn 12

End Sub

Public Sub Med1Sterkte()

    MedSterkte 1, False

End Sub

Public Sub Med2Sterkte()

    MedSterkte 2, False

End Sub

Public Sub Med3Sterkte()

    MedSterkte 3, False

End Sub

Public Sub Med4Sterkte()

    MedSterkte 4, False

End Sub

Public Sub Med5Sterkte()

        MedSterkte 5, False

End Sub

Public Sub Med6Sterkte()

    MedSterkte 6, False

End Sub

Public Sub Med7Sterkte()

    MedSterkte 7, False

End Sub

Public Sub Med8Sterkte()

    MedSterkte 8, False

End Sub

Public Sub Med9Sterkte()

    MedSterkte 9, False

End Sub

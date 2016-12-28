Attribute VB_Name = "ModAfspraken"
Option Explicit

' ToDo: Add comment
Public Sub AfsprakenOvernemen()

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
        Range(arrTo(intN)).Value = Range(arrFrom(intN)).Value
    Next intN
    
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


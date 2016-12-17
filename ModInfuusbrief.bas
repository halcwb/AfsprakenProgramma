Attribute VB_Name = "ModInfuusbrief"
Option Explicit

Sub VerwijderContInfuus(regel As Integer)
Dim nrMedicament As String
Dim nrMedSterkte As String
Dim nrOplHoev As String
Dim nrOplossing As String
Dim nrStand As String
Dim nrExtra As String

    nrMedicament = "_Medicament_" & regel
    nrMedSterkte = "_MedSterkte_" & regel
    nrOplHoev = "_OplHoev_" & regel
    nrOplossing = "_Oplossing_" & regel
    nrStand = "_Stand_" & regel
    nrExtra = "_Extra_" & regel
    
    Range(nrMedSterkte).Value = 0
    Range(nrOplHoev).Value = 0
    Range(nrStand).Value = 0
    Range(nrExtra).Value = 0
    
    Range(nrOplossing).Value = Application.VLookup(Range("Medicamenten").Cells(Range(nrMedicament).Value, 1), Range("Medicamenten"), 10, False)
    If Not IsNumeric(Range(nrOplossing).Value) Then
        Range(nrOplossing).Value = 1
    End If
End Sub

Sub Afspraken_Vervolgkeuzelijst1_BijWijzigen()
    VerwijderContInfuus 1
End Sub
Sub Vervolgkeuzelijst2_BijWijzigen()
    VerwijderContInfuus 2
End Sub
Sub Vervolgkeuzelijst3_BijWijzigen()
    VerwijderContInfuus 3
End Sub
Sub Vervolgkeuzelijst4_BijWijzigen()
    VerwijderContInfuus 4
End Sub
Sub Vervolgkeuzelijst5_BijWijzigen()
    VerwijderContInfuus 5
End Sub
Sub Vervolgkeuzelijst6_BijWijzigen()
    VerwijderContInfuus 6
End Sub
Sub Vervolgkeuzelijst7_BijWijzigen()
    VerwijderContInfuus 7
End Sub
Sub Vervolgkeuzelijst67_BijWijzigen()
    VerwijderContInfuus 8
End Sub
Sub Vervolgkeuzelijst97_BijWijzigen()
    VerwijderContInfuus 9
End Sub

Sub VerwijderZijlijn(regel As Integer)
Dim nrStand As String
Dim nrExtra As String

    nrStand = "_Stand_" & regel
    nrExtra = "_Extra_" & regel + 1
    
    Range(nrStand).Value = 0
    Range(nrExtra).Value = 0
    
End Sub

Sub Vervolgkeuzelijst101_BijWijzigen()
    VerwijderZijlijn 10
End Sub

Sub Vervolgkeuzelijst104_BijWijzigen()
    VerwijderZijlijn 11
End Sub

Sub Vervolgkeuzelijst108_BijWijzigen()
    VerwijderZijlijn 12
End Sub

Public Sub Med1Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 1"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_1").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_1").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med2Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 2"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_2").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_2").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med3Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 3"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_3").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_3").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med4Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 4"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_4").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_4").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med5Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 5"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_5").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_5").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med6Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 6"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_6").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_6").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med7Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 7"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_7").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_7").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med8Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 8"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_8").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_8").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med9Sterkte()

    Dim frmInvoer As New frmInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament 9"
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_9").Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte_9").Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub



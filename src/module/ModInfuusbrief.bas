Attribute VB_Name = "ModInfuusbrief"
Option Explicit

Public Sub VerwijderContInfuus(intRegel As Integer)

    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim strStand As String
    Dim strExtra As String

    strMedicament = "_Medicament_" & intRegel
    strMedSterkte = "_MedSterkte_" & intRegel
    strOplHoev = "_OplHoev_" & intRegel
    strOplossing = "_Oplossing_" & intRegel
    strStand = "_Stand_" & intRegel
    strExtra = "_Extra_" & intRegel
    
    Range(strMedSterkte).Value = 0
    Range(strOplHoev).Value = 0
    Range(strStand).Value = 0
    Range(strExtra).Value = 0
    
    Range(strOplossing).Value = Application.VLookup(Range("Medicamenten").Cells(Range(strMedicament).Value, 1), Range("Medicamenten"), 10, False)
    
    If Not IsNumeric(Range(strOplossing).Value) Then
        Range(strOplossing).Value = 1
    End If

End Sub

Public Sub Afspraken_Vervolgkeuzelijst1_BijWijzigen()

    VerwijderContInfuus 1

End Sub

Public Sub Vervolgkeuzelijst2_BijWijzigen()
    
    VerwijderContInfuus 2

End Sub

Public Sub Vervolgkeuzelijst3_BijWijzigen()
    
    VerwijderContInfuus 3

End Sub

Public Sub Vervolgkeuzelijst4_BijWijzigen()
    
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

Private Sub VerwijderZijlijn(intRegel As Integer)

    Dim strStand As String
    Dim strExtra As String

    strStand = "_Stand_" & intRegel
    strExtra = "_Extra_" & intRegel + 1
    
    Range(strStand).Value = 0
    Range(strExtra).Value = 0
    
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

Private Sub MedSterkte(intN As Integer)

    Dim frmInvoer As New FormInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament " & intN
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte_" & intN).Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
           Range("_MedSterkte_" & intN).Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub


Public Sub Med1Sterkte()

    MedSterkte (1)

End Sub

Public Sub Med2Sterkte()

    MedSterkte (2)

End Sub

Public Sub Med3Sterkte()

    MedSterkte (3)

End Sub

Public Sub Med4Sterkte()

    MedSterkte (4)

End Sub

Public Sub Med5Sterkte()

        MedSterkte (5)

End Sub

Public Sub Med6Sterkte()

    MedSterkte (6)

End Sub

Public Sub Med7Sterkte()

    MedSterkte (7)

End Sub

Public Sub Med8Sterkte()

    MedSterkte (8)

End Sub

Public Sub Med9Sterkte()

    MedSterkte (9)

End Sub



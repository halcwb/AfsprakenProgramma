Attribute VB_Name = "modPedContIV"
Option Explicit

Sub WijzigContIVMedicament(regel As Integer)
Dim nrMedicament As String
Dim nrMedSterkte As String
Dim nrOplHoev As String
Dim nrOplossing As String
Dim nrStand As String

    nrMedicament = "MedIVKeuze_" & regel
    nrMedSterkte = "MedIVSterkte_" & regel
    nrOplHoev = "MedIVMlOpl_" & regel
    nrOplossing = "MedIVOplVlst_" & regel
    nrStand = "MedIVStand_" & regel
    
    Range(nrMedSterkte).Value = 0
    Range(nrOplHoev).Value = 0
    Range(nrStand).Value = 0
    
    If Range(nrMedicament).Value = 1 Then
        Range(nrOplossing).Value = 1
    Else
        Range(nrOplossing).Value = Application.VLookup(Range("tblMedicationContIV").Cells(Range(nrMedicament).Value, 1), Range("tblMedicationContIV"), 22, False)
    End If

End Sub

Sub MedicatieIV_Vervolgkeuzelijst2_BijWijzigen()
    WijzigContIVMedicament 1
End Sub

Sub MedicatieIV_Vervolgkeuzelijst3_BijWijzigen()
    WijzigContIVMedicament 2
End Sub
Sub MedicatieIV_Vervolgkeuzelijst4_BijWijzigen()
    WijzigContIVMedicament 3
End Sub
Sub MedicatieIV_Vervolgkeuzelijst5_BijWijzigen()
    WijzigContIVMedicament 4
End Sub
Sub MedicatieIV_Vervolgkeuzelijst6_BijWijzigen()
    WijzigContIVMedicament 5
End Sub
Sub MedicatieIV_Vervolgkeuzelijst7_BijWijzigen()
    WijzigContIVMedicament 6
End Sub
Sub Vervolgkeuzelijst8_BijWijzigen()
    WijzigContIVMedicament 7
End Sub
Sub Vervolgkeuzelijst9_BijWijzigen()
    WijzigContIVMedicament 8
End Sub
Sub Vervolgkeuzelijst10_BijWijzigen()
    WijzigContIVMedicament 9
End Sub
Sub Vervolgkeuzelijst76_BijWijzigen()
    WijzigContIVMedicament 10
End Sub

Public Sub PedMed1Sterkte()
    PedMedSterkte 1
End Sub

Public Sub PedMed2Sterkte()
    PedMedSterkte 2
End Sub

Public Sub PedMed3Sterkte()
    PedMedSterkte 3
End Sub

Public Sub PedMed4Sterkte()
    PedMedSterkte 4
End Sub

Public Sub PedMed5Sterkte()
    PedMedSterkte 5
End Sub

Public Sub PedMed6Sterkte()
    PedMedSterkte 6
End Sub

Public Sub PedMed7Sterkte()
    PedMedSterkte 7
End Sub

Public Sub PedMed8Sterkte()
    PedMedSterkte 8
End Sub

Public Sub PedMed9Sterkte()
    PedMedSterkte 9
End Sub

Public Sub PedMed10Sterkte()
    PedMedSterkte 10
End Sub

Public Sub PedMedSterkte(regel As Integer)

    Dim frmInvoer As New frmInvoerNumeriek
    Dim y As Variant
    With frmInvoer
        .Caption = "Medicament " & regel
        .lblParameter = "Sterkte"
        .lblEenheid = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & regel), 4)
        If (Range("MedIVSterkte_" & regel).Value = 0) Then
            .txtWaarde = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & regel), 11)
        Else
            .txtWaarde = Range("MedIVSterkte_" & regel).Value
        End If
        .Show
        If IsNumeric(.txtWaarde) Then
            If CDbl(.txtWaarde) = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & regel), 11) Then
                Range("MedIVSterkte_" & regel).Formula = 0
            Else
                Range("MedIVSterkte_" & regel).Formula = .txtWaarde
            End If
        End If
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub PedMed1Oplossing()
    PedMedOplossing 1
End Sub

Public Sub PedMed2Oplossing()
    PedMedOplossing 2
End Sub

Public Sub PedMed3Oplossing()
    PedMedOplossing 3
End Sub

Public Sub PedMed4Oplossing()
    PedMedOplossing 4
End Sub

Public Sub PedMed5Oplossing()
    PedMedOplossing 5
End Sub

Public Sub PedMed6Oplossing()
    PedMedOplossing 6
End Sub

Public Sub PedMed7Oplossing()
    PedMedOplossing 7
End Sub

Public Sub PedMed8Oplossing()
    PedMedOplossing 8
End Sub

Public Sub PedMed9Oplossing()
    PedMedOplossing 9
End Sub

Public Sub PedMed10Oplossing()
    PedMedOplossing 10
End Sub

Public Sub PedMedOplossing(regel As Integer)

    Dim frmInvoer As New frmInvoerNumeriek
    Dim y As Variant
    With frmInvoer
        .Caption = "Medicament " & regel
        .lblParameter = "Oplossing"
        .lblEenheid = "ml"
        If (Range("MedIVMlOpl_" & regel).Value = 0) Then
            .txtWaarde = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & regel), 12)
        Else
            .txtWaarde = Range("MedIVMlOpl_" & regel).Value
        End If
        .Show
        If IsNumeric(.txtWaarde) Then
            If CDbl(.txtWaarde) = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & regel), 12) Then
                Range("MedIVMlOpl_" & regel).Formula = 0
            Else
                Range("MedIVMlOpl_" & regel).Formula = .txtWaarde
            End If
        End If
    End With
    
    Set frmInvoer = Nothing

End Sub



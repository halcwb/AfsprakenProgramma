Attribute VB_Name = "modPedContIV"
Option Explicit

Sub WijzigContIVMedicament(intRegel As Integer)

    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim varOplossing As Variant
    Dim strStand As String
    Dim blnLog As Boolean
    
    blnLog = ModSetting.GetEnableLogging
    
    On Error GoTo ChangeMedIVError

    strMedicament = "MedIVKeuze_" & intRegel
    strMedSterkte = "MedIVSterkte_" & intRegel
    strOplHoev = "MedIVMlOpl_" & intRegel
    strOplossing = "MedIVOplVlst_" & intRegel
    strStand = "MedIVStand_" & intRegel
    
    ModRange.SetRangeValue strMedSterkte, 0
    ModRange.SetRangeValue strOplHoev, 0
    ModRange.SetRangeValue strStand, 0
    
    If ModRange.GetRangeValue(strMedicament, 0) = 1 Then
        ModRange.SetRangeValue strOplossing, 1
    Else
        varOplossing = Application.VLookup(Range("tblMedicationContIV").Cells(Range(strMedicament).Value, 1), Range("tblMedicationContIV"), 22, False)
        ModRange.SetRangeValue strOplossing, varOplossing
    End If
    
    Exit Sub
    
ChangeMedIVError:

    ModLog.EnableLogging
    ModLog.LogToFile ModSetting.GetLogPath, Error, "ModPedContIV: " & " Error for intRegel " & intRegel
    If Not blnLog Then ModLog.DisableLogging

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

Public Sub PedMedSterkte(intRegel As Integer)

    Dim frmInvoer As New FormInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament " & intRegel
        .lblParameter = "Sterkte"
        .lblEenheid = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & intRegel), 4)
        If (Range("MedIVSterkte_" & intRegel).Value = 0) Then
            .txtWaarde = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & intRegel), 11)
        Else
            .txtWaarde = Range("MedIVSterkte_" & intRegel).Value
        End If
        .Show
        If IsNumeric(.txtWaarde) Then
            If CDbl(.txtWaarde) = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & intRegel), 11) Then
                Range("MedIVSterkte_" & intRegel).Formula = 0
            Else
                Range("MedIVSterkte_" & intRegel).Formula = .txtWaarde
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

Public Sub PedMedOplossing(intRegel As Integer)

    Dim frmInvoer As New FormInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament " & intRegel
        .lblParameter = "Oplossing"
        .lblEenheid = "ml"
        If (Range("MedIVMlOpl_" & intRegel).Value = 0) Then
            .txtWaarde = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & intRegel), 12)
        Else
            .txtWaarde = Range("MedIVMlOpl_" & intRegel).Value
        End If
        .Show
        If IsNumeric(.txtWaarde) Then
            If CDbl(.txtWaarde) = Application.WorksheetFunction.Index(Range("tblMedicationContIV"), Range("MedIVKeuze_" & intRegel), 12) Then
                Range("MedIVMlOpl_" & intRegel).Formula = 0
            Else
                Range("MedIVMlOpl_" & intRegel).Formula = .txtWaarde
            End If
        End If
    End With
    
    Set frmInvoer = Nothing

End Sub



Attribute VB_Name = "ModAfspraken"
Option Explicit

Public Sub AfsprakenOvernemen()
    VoedingOvernemen
    ContMedOvernemen
    TPNOvernemen
End Sub

Private Sub VoedingOvernemen()
    Range("_Voeding1700").Value = Range("_Voeding").Value
    
    Range("_Frequentie1700_1").Value = Range("_Frequentie_1").Value
    Range("_Hoeveelheid1700_1").Value = Range("_Hoeveelheid_1").Value
    Range("_Frequentie1700_2").Value = Range("_Frequentie_2").Value
    Range("_Hoeveelheid1700_2").Value = Range("_Hoeveelheid_2").Value

    Range("_Fototherapie1700").Value = Range("_Fototherapie").Value
    Range("_Parenteraal1700").Value = Range("_Parenteraal").Value

    Range("_Toevoeging1700_1").Value = Range("_Toevoeging_1").Value
    Range("_Toevoeging1700_2").Value = Range("_Toevoeging_2").Value
    Range("_Toevoeging1700_3").Value = Range("_Toevoeging_3").Value
    Range("_Toevoeging1700_4").Value = Range("_Toevoeging_4").Value

    Range("_Toevoeging1700_5").Value = Range("_Toevoeging_5").Value
    Range("_Toevoeging1700_6").Value = Range("_Toevoeging_6").Value
    Range("_Toevoeging1700_7").Value = Range("_Toevoeging_7").Value
    Range("_Toevoeging1700_8").Value = Range("_Toevoeging_8").Value

    Range("_PercentageKeuze1700_0").Value = Range("_PercentageKeuze_0").Value
    Range("_PercentageKeuze1700_1").Value = Range("_PercentageKeuze_1").Value
    Range("_PercentageKeuze1700_2").Value = Range("_PercentageKeuze_2").Value
    Range("_PercentageKeuze1700_3").Value = Range("_PercentageKeuze_3").Value
    Range("_PercentageKeuze1700_4").Value = Range("_PercentageKeuze_4").Value

    Range("_PercentageKeuze1700_5").Value = Range("_PercentageKeuze_5").Value
    Range("_PercentageKeuze1700_6").Value = Range("_PercentageKeuze_6").Value
    Range("_PercentageKeuze1700_7").Value = Range("_PercentageKeuze_7").Value
    Range("_PercentageKeuze1700_8").Value = Range("_PercentageKeuze_8").Value

    Range("_IntakePerKg1700").Value = Range("_IntakePerKg").Value

    Range("_Extra1700").Value = Range("_Extra").Value
End Sub

Private Sub ContMedOvernemen()
    Range("_ArtLijn1700").Value = Range("_ArtLijn").Value
    
    Range("_Medicament1700_1").Value = Range("_Medicament_1").Value
    Range("_Medicament1700_2").Value = Range("_Medicament_2").Value
    Range("_Medicament1700_3").Value = Range("_Medicament_3").Value
    Range("_Medicament1700_4").Value = Range("_Medicament_4").Value
    Range("_Medicament1700_5").Value = Range("_Medicament_5").Value
    Range("_Medicament1700_6").Value = Range("_Medicament_6").Value
    Range("_Medicament1700_7").Value = Range("_Medicament_7").Value
    Range("_Medicament1700_8").Value = Range("_Medicament_8").Value
    Range("_Medicament1700_9").Value = Range("_Medicament_9").Value

    Range("_MedSterkte1700_1").Value = Range("_MedSterkte_1").Value
    Range("_MedSterkte1700_2").Value = Range("_MedSterkte_2").Value
    Range("_MedSterkte1700_3").Value = Range("_MedSterkte_3").Value
    Range("_MedSterkte1700_4").Value = Range("_MedSterkte_4").Value
    Range("_MedSterkte1700_5").Value = Range("_MedSterkte_5").Value
    Range("_MedSterkte1700_6").Value = Range("_MedSterkte_6").Value
    Range("_MedSterkte1700_7").Value = Range("_MedSterkte_7").Value
    Range("_MedSterkte1700_8").Value = Range("_MedSterkte_8").Value
    Range("_MedSterkte1700_9").Value = Range("_MedSterkte_9").Value
    
    Range("_OplHoev1700_1").Value = Range("_OplHoev_1").Value
    Range("_OplHoev1700_2").Value = Range("_OplHoev_2").Value
    Range("_OplHoev1700_3").Value = Range("_OplHoev_3").Value
    Range("_OplHoev1700_4").Value = Range("_OplHoev_4").Value
    Range("_OplHoev1700_5").Value = Range("_OplHoev_5").Value
    Range("_OplHoev1700_6").Value = Range("_OplHoev_6").Value
    Range("_OplHoev1700_7").Value = Range("_OplHoev_7").Value
    Range("_OplHoev1700_8").Value = Range("_OplHoev_8").Value
    Range("_OplHoev1700_9").Value = Range("_OplHoev_9").Value

    Range("_Oplossing1700_1").Value = Range("_Oplossing_1").Value
    Range("_Oplossing1700_2").Value = Range("_Oplossing_2").Value
    Range("_Oplossing1700_3").Value = Range("_Oplossing_3").Value
    Range("_Oplossing1700_4").Value = Range("_Oplossing_4").Value
    Range("_Oplossing1700_5").Value = Range("_Oplossing_5").Value
    Range("_Oplossing1700_6").Value = Range("_Oplossing_6").Value
    Range("_Oplossing1700_7").Value = Range("_Oplossing_7").Value
    Range("_Oplossing1700_8").Value = Range("_Oplossing_8").Value
    Range("_Oplossing1700_9").Value = Range("_Oplossing_9").Value
    Range("_Oplossing1700_10").Value = Range("_Oplossing_10").Value
    Range("_Oplossing1700_11").Value = Range("_Oplossing_11").Value
    Range("_Oplossing1700_12").Value = Range("_Oplossing_12").Value

    Range("_Stand1700_1").Value = Range("_Stand_1").Value
    Range("_Stand1700_2").Value = Range("_Stand_2").Value
    Range("_Stand1700_3").Value = Range("_Stand_3").Value
    Range("_Stand1700_4").Value = Range("_Stand_4").Value
    Range("_Stand1700_5").Value = Range("_Stand_5").Value
    Range("_Stand1700_6").Value = Range("_Stand_6").Value
    Range("_Stand1700_7").Value = Range("_Stand_7").Value
    Range("_Stand1700_8").Value = Range("_Stand_8").Value
    Range("_Stand1700_9").Value = Range("_Stand_9").Value
    Range("_Stand1700_10").Value = Range("_Stand_10").Value
    Range("_Stand1700_11").Value = Range("_Stand_11").Value
    Range("_Stand1700_12").Value = Range("_Stand_12").Value

    Range("_Extra1700_1").Value = Range("_Extra_1").Value
    Range("_Extra1700_2").Value = Range("_Extra_2").Value
    Range("_Extra1700_3").Value = Range("_Extra_3").Value
    Range("_Extra1700_4").Value = Range("_Extra_4").Value
    Range("_Extra1700_5").Value = Range("_Extra_5").Value
    Range("_Extra1700_6").Value = Range("_Extra_6").Value
    Range("_Extra1700_7").Value = Range("_Extra_7").Value
    Range("_Extra1700_8").Value = Range("_Extra_8").Value
    Range("_Extra1700_9").Value = Range("_Extra_9").Value
    Range("_Extra1700_10").Value = Range("_Extra_10").Value
    Range("_Extra1700_11").Value = Range("_Extra_11").Value
    Range("_Extra1700_12").Value = Range("_Extra_12").Value

    Range("_MedTekst1700_1").Value = Range("_MedTekst_1").Value
    Range("_MedTekst1700_2").Value = Range("_MedTekst_2").Value
End Sub

Private Sub TPNOvernemen()
    Range("_Parenteraal1700").Value = Range("_Parenteraal").Value
    Range("_IntraLipid1700").Value = Range("_IntraLipid").Value
    
    Range("_DagKeuze1700").Value = Range("_DagKeuze").Value
    
    Range("_NaCl1700").Value = Range("_NaCl").Value
    Range("_KCl1700").Value = Range("_KCl").Value
    Range("_CaCl21700").Value = Range("_CaCl2").Value
    Range("_MgCl21700").Value = Range("_MgCl2").Value
    Range("_SoluVit1700").Value = Range("_SoluVit").Value
    Range("_Primene1700").Value = Range("_Primene").Value
    Range("_NICUMix1700").Value = Range("_NICUMix").Value
    Range("_SSTB1700").Value = Range("_SSTB").Value
    Range("_GlucSterkte1700").Value = Range("_GlucSterkte").Value
End Sub



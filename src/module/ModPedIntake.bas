Attribute VB_Name = "ModPedIntake"
Option Explicit

Public Sub SpecialeVoeding()
    
    Dim frmSpecialeVoeding As New FormSpecialeVoeding

    frmSpecialeVoeding.Show
    
    Set frmSpecialeVoeding = Nothing

End Sub

Public Sub ChangeEntVoedKeuze()

    If Range("_Ped_Ent_Keuze_1").Value = 1 Then
    
        ModRange.SetRangeValue "_Ped_Ent_Keuze_2", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Keuze_3", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Keuze_4", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Freq_1", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Freq_2", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Freq_3", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Freq_4", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_1", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_2", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_3", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_4", vbNullString
        
    End If

End Sub
Public Sub ChangeEntToevKeuze_1()

    If ModRange.GetRangeValue("_Ped_Ent_Keuze_2", 0) = 1 Then
        ModRange.SetRangeValue "_Ped_Ent_Freq_2", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_2", vbNullString
    End If

End Sub

Public Sub ChangeEntToevKeuze_2()
    
    If ModRange.GetRangeValue("_Ped_Ent_Keuze_3", 0) = 1 Then
        ModRange.SetRangeValue "_Ped_Ent_Freq_3", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_3", vbNullString
    End If

End Sub

Public Sub ChangeEntToevKeuze_3()
    
    If ModRange.GetRangeValue("_Ped_Ent_Keuze_4", 0) = 1 Then
        ModRange.SetRangeValue "_Ped_Ent_Freq_4", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_4", vbNullString
    End If

End Sub

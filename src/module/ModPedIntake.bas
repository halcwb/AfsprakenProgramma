Attribute VB_Name = "ModPedIntake"
Option Explicit

Sub Intake_Vervolgkeuzelijst4_BijWijzigen()

    If Range("KeuzePO_1").Value = 1 Then
    
        ModRange.SetRangeValue "KeuzePO_2", vbNullString
        ModRange.SetRangeValue "KeuzePO_3", vbNullString
        ModRange.SetRangeValue "KeuzePO_4", vbNullString
        ModRange.SetRangeValue "FreqPO_1", vbNullString
        ModRange.SetRangeValue "FreqPO_2", vbNullString
        ModRange.SetRangeValue "FreqPO_3", vbNullString
        ModRange.SetRangeValue "FreqPO_4", vbNullString
        ModRange.SetRangeValue "VolPO_1", vbNullString
        ModRange.SetRangeValue "VolPO_2", vbNullString
        ModRange.SetRangeValue "VolPO_3", vbNullString
        ModRange.SetRangeValue "VolPO_4", vbNullString
        
    End If

End Sub
Sub Vervolgkeuzelijst57_BijWijzigen()

    If ModRange.GetRangeValue("KeuzePO_2", 0) = 1 Then
        ModRange.SetRangeValue "FreqPO_2", vbNullString
        ModRange.SetRangeValue "VolPO_2", vbNullString
    End If

End Sub

Sub Vervolgkeuzelijst62_BijWijzigen()
    
    If Range("KeuzePO_3").Value = 1 Then
        ModRange.SetRangeValue "FreqPO_3", vbNullString
        ModRange.SetRangeValue "VolPO_3", vbNullString
    End If

End Sub

Sub Vervolgkeuzelijst65_BijWijzigen()
    
    If Range("KeuzePO_4").Value = 1 Then
        ModRange.SetRangeValue "FreqPO_4", vbNullString
        ModRange.SetRangeValue "VolPO_4", vbNullString
    End If

End Sub

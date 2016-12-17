Attribute VB_Name = "ModPedIntake"
Option Explicit

Sub Intake_Vervolgkeuzelijst4_BijWijzigen()
    If Range("KeuzePO_1").Value = 1 Then
        Range("KeuzePO_2").Value = vbNullString
        Range("KeuzePO_3").Value = vbNullString
        Range("KeuzePO_4").Value = vbNullString
        Range("FreqPO_1").Value = vbNullString
        Range("FreqPO_2").Value = vbNullString
        Range("FreqPO_3").Value = vbNullString
        Range("FreqPO_4").Value = vbNullString
        Range("VolPO_1").Value = vbNullString
        Range("VolPO_2").Value = vbNullString
        Range("VolPO_3").Value = vbNullString
        Range("VolPO_4").Value = vbNullString
    End If
End Sub
Sub Vervolgkeuzelijst57_BijWijzigen()
    If Range("KeuzePO_2").Value = 1 Then
        Range("FreqPO_2").Value = vbNullString
        Range("VolPO_2").Value = vbNullString
    End If
End Sub
Sub Vervolgkeuzelijst62_BijWijzigen()
    If Range("KeuzePO_3").Value = 1 Then
        Range("FreqPO_3").Value = vbNullString
        Range("VolPO_3").Value = vbNullString
    End If
End Sub
Sub Vervolgkeuzelijst65_BijWijzigen()
    If Range("KeuzePO_4").Value = 1 Then
        Range("FreqPO_4").Value = vbNullString
        Range("VolPO_4").Value = vbNullString
    End If
End Sub

Attribute VB_Name = "ModAfspraken1700"
Option Explicit

Sub VerwijderContInfuus1700(intRegel As Integer)

    Dim strMedicament As String
    Dim strMedSterkte As String
    Dim strOplHoev As String
    Dim strOplossing As String
    Dim strStand As String
    Dim strExtra As String

    strMedicament = "_Medicament1700_" & intRegel
    strMedSterkte = "_MedSterkte1700_" & intRegel
    strOplHoev = "_OplHoev1700_" & intRegel
    strOplossing = "_Oplossing1700_" & intRegel
    strStand = "_Stand1700_" & intRegel
    strExtra = "_Extra1700_" & intRegel
    
    Range(strMedSterkte).Value = 0
    Range(strOplHoev).Value = 0
    Range(strStand).Value = 0
    Range(strExtra).Value = 0
    
    Range(strOplossing).Value = Application.VLookup(Range("Medicamenten").Cells(Range(strMedicament).Value, 1), Range("Medicamenten"), 10, False)
    If Not IsNumeric(Range(strOplossing).Value) Then
        Range(strOplossing).Value = 1
    End If
    
End Sub

Sub VerwijderContInfuus1700_1()
    VerwijderContInfuus1700 1
End Sub
Sub VerwijderContInfuus1700_2()
    VerwijderContInfuus1700 2
End Sub
Sub VerwijderContInfuus1700_3()
    VerwijderContInfuus1700 3
End Sub
Sub VerwijderContInfuus1700_4()
    VerwijderContInfuus1700 4
End Sub
Sub VerwijderContInfuus1700_5()
    VerwijderContInfuus1700 5
End Sub
Sub VerwijderContInfuus1700_6()
    VerwijderContInfuus1700 6
End Sub
Sub VerwijderContInfuus1700_7()
    VerwijderContInfuus1700 7
End Sub
Sub VerwijderContInfuus1700_8()
    VerwijderContInfuus1700 8
End Sub
Sub VerwijderContInfuus1700_9()
    VerwijderContInfuus1700 9
End Sub

Private Sub MedSterkte1700(intRegel As Integer)

    Dim frmInvoer As New FormInvoerNumeriek
    
    With frmInvoer
        .Caption = "Medicament " & intRegel
        .lblParameter = "Sterkte"
        .lblEenheid = "mg"
        .txtWaarde = Range("_MedSterkte1700_" & intRegel).Value / 10
        .Show
        If IsNumeric(.txtWaarde) Then _
            Range("_MedSterkte1700_" & intRegel).Formula = .txtWaarde * 10
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub MedSterkte1700_1()
    MedSterkte1700 1
End Sub

Public Sub MedSterkte1700_2()
    MedSterkte1700 2
End Sub

Public Sub MedSterkte1700_3()
    MedSterkte1700 3
End Sub

Public Sub MedSterkte1700_4()
    MedSterkte1700 4
End Sub

Public Sub MedSterkte1700_5()
    MedSterkte1700 5
End Sub

Public Sub MedSterkte1700_6()
    MedSterkte1700 6
End Sub

Public Sub MedSterkte1700_7()
    MedSterkte1700 7
End Sub

Public Sub MedSterkte1700_8()
    MedSterkte1700 8
End Sub

Public Sub MedSterkte1700_9()
    MedSterkte1700 9
End Sub

Public Sub CopyToActueel()

    Dim frmCopy1700 As New FormCopy1700
'Show selectie form
    frmCopy1700.Show
'Copy by block

    Set frmCopy1700 = Nothing
    
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
    
    arrTo = ModAfspraken.GetVoedingItems()
    arrFrom = ModAfspraken.Get1700Items(arrTo)
    
    ModAfspraken.CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

Private Sub ContMedOvernemen()
    
    Dim arrTo() As String
    Dim arrFrom() As String
    
    arrTo = ModAfspraken.GetIVAfsprItems()
    arrFrom = ModAfspraken.Get1700Items(arrTo)
    
    ModAfspraken.CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

Private Sub TPNOvernemen()
    
    Dim arrTo() As String
    Dim arrFrom() As String
    
    arrTo = ModAfspraken.GetTPNItems()
    arrFrom = ModAfspraken.Get1700Items(arrTo)
    
    ModAfspraken.CopyRangeNamesToRangeNames arrFrom, arrTo

End Sub

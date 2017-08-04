Attribute VB_Name = "ModNeoInfB_Tests"
Option Explicit

Private Const constTestStart As Integer = 3
Private Const constTestCount As Integer = 100
Private Const constTestNum As String = "A"
Private Const constSetupGewicht As String = "B"
Private Const constSetupMedicament As String = "C"
Private Const constSetupHoeveelheid As String = "D"
Private Const constSetupOplosmiddel As String = "E"
Private Const constSetupOploshoeveelheid As String = "F"
Private Const constSetupInfuusStand As String = "G"

Private Const constActGewicht As String = "S"
Private Const constActMedicament As String = "T"
Private Const constActHoeveelheid As String = "U"
Private Const constActEenheid As String = "V"
Private Const constActOplosmiddel As String = "W"
Private Const constActOploshoeveelheid As String = "X"
Private Const constActInfuusStand As String = "Y"
Private Const constActDosis As String = "Z"
Private Const constActDosisEenheid As String = "AA"
Private Const constActNormaalWaarde As String = "AB"
Private Const constActInloopTijd As String = "AC"

Private Const constGewicht As String = "_Pat_Gewicht"
Private Const constMedicament As String = "Var_Neo_InfB_Cont_MedKeuze_"
Private Const constHoeveelheid As String = "Var_Neo_InfB_Cont_MedSterkte_"
Private Const constOplosmiddel As String = "Var_Neo_InfB_Cont_Oplossing_"
Private Const constOploshoeveelheid As String = "Var_Neo_InfB_Cont_OplHoev_"
Private Const constInfuusStand As String = "Var_Neo_InfB_Cont_Stand_"
Private Const constDosis As String = "G"
Private Const constDosisEenheid As String = "R"
Private Const constNormaalWaarde As String = "S"
Private Const constInloopTijd As String = "Z"

Private Const constTblMedIV As String = "Tbl_Neo_MedIV"
Private Const constTblOpl As String = "Tbl_Neo_OplVlst"

Public Sub Test_NeoInfB_ContMed()

    Dim wbkTests As Workbook
    Dim shtTests As Worksheet
    Dim intN As Integer
    Dim intM As Integer
    Dim strM As String
    Dim strNo As String
    Dim varVal As Variant
    Dim dblGew As Double
    Dim intMed As Integer
    Dim dblHoev As Double
    Dim intOpl As Integer
    Dim dblOplHoev As Double
    Dim dblStand As Double
    Dim blnPass As Boolean
    
    On Error GoTo Test_NeoInfB_ContMedError
    
    ModProgress.StartProgress "Neo Infuusbrief Continue Medicatie Tests"
    
    'ModPatient.PatientClearAll False, True
    
    Set wbkTests = Workbooks.Open(WbkAfspraken.Path & "/tests/Tests.xlsx")
    Set shtTests = wbkTests.Sheets("NICU_ContMed")

    blnPass = True
    For intN = constTestStart To constTestCount
        strNo = shtTests.Range(constTestNum & intN).Value2
        ModProgress.SetJobPercentage "Testing", constTestCount, intN
        
        ' Gewicht
        dblGew = ModString.StringToDouble(shtTests.Range(constSetupGewicht & intN).Value2)
        blnPass = blnPass And ModRange.SetRangeValue(constGewicht, dblGew * 10)
        
        ' Medicament
        varVal = shtTests.Range(constSetupMedicament & intN).Value2
        If IsEmpty(varVal) Then
            intMed = 1
        Else
            intMed = ModExcel.Excel_VLookup(varVal, constTblMedIV, 21)
        End If
        
        ' Hoeveelheid
        dblHoev = ModString.StringToDouble(shtTests.Range(constSetupHoeveelheid & intN).Value2)
        
        ' Oplosmiddel
        varVal = shtTests.Range(constSetupOplosmiddel & intN).Value2
        If IsEmpty(varVal) Then
            intOpl = 1
        Else
            intOpl = ModExcel.Excel_VLookup(varVal, constTblOpl, 2)
        End If
        
        ' Oplos hoeveelheid
        dblOplHoev = ModString.StringToDouble(shtTests.Range(constSetupOploshoeveelheid & intN).Value2)
        
        ' Infuus stand
        dblStand = ModString.StringToDouble(shtTests.Range(constSetupInfuusStand & intN).Value2)
        
        ' Voer testen uit
        For intM = 1 To 10
            ' Voer testcase in
            strM = IIf(intM < 10, "0" & intM, intM)
            blnPass = blnPass And ModRange.SetRangeValue(constMedicament & strM, intMed)
            ChangeMedIV intM
            If dblHoev > 0 Then blnPass = blnPass And ModRange.SetRangeValue(constHoeveelheid & strM, dblHoev)
            If intOpl > 1 Then blnPass = blnPass And ModRange.SetRangeValue(constOplosmiddel & strM, intOpl)
            If dblOplHoev > 0 Then blnPass = blnPass And ModRange.SetRangeValue(constOploshoeveelheid & strM, dblOplHoev)
            If dblStand > 0 Then blnPass = blnPass And ModRange.SetRangeValue(constInfuusStand & strM, dblStand)
        Next
        
        ' Schrijf gewicht weg
        shtTests.Range(constActGewicht & intN).Value2 = ModRange.GetRangeValue(constGewicht, 0) / 10
        
        ' Check medicament
        varVal = shtNeoPrtWerkbr.Range("C24").Value2
        For intM = 0 To 9
            'Check werkbrief
            blnPass = blnPass And varVal = shtNeoPrtWerkbr.Range("C" & intM * 3 + 24).Value2
            ' Check apotheek print
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            blnPass = blnPass And varVal = shtNeoPrtApoth.Range("D5").Value2
        Next
        
        ' Schrijf medicament
        shtTests.Range(constActMedicament & intN).Value2 = varVal
        
        
    Next

    ModProgress.FinishProgress

    If blnPass Then
        ModMessage.ShowMsgBoxInfo "Alle testen geslaagt"
    Else
        ModMessage.ShowMsgBoxExclam "Niet alle testen geslaagt"
    End If
    
    wbkTests.Close True
    Set shtTests = Nothing
    Set wbkTests = Nothing
    
    Exit Sub
    
Test_NeoInfB_ContMedError:

    On Error Resume Next
    
    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxError "Kan tests niet uitvoeren: " & Err.Number & " " & Err.Description
    
    wbkTests.Close False
    Set shtTests = Nothing
    Set wbkTests = Nothing
End Sub


Private Sub ChangeMedIV(ByVal intN As Integer)

    Select Case intN
        Case 1
        NeoInfB_ChangeMedContIV_01
        Case 2
        NeoInfB_ChangeMedContIV_02
        Case 3
        NeoInfB_ChangeMedContIV_03
        Case 4
        NeoInfB_ChangeMedContIV_04
        Case 5
        NeoInfB_ChangeMedContIV_05
        Case 6
        NeoInfB_ChangeMedContIV_06
        Case 7
        NeoInfB_ChangeMedContIV_07
        Case 8
        NeoInfB_ChangeMedContIV_08
        Case 9
        NeoInfB_ChangeMedContIV_09
        Case 10
        NeoInfB_ChangeMedContIV_10
    End Select

End Sub

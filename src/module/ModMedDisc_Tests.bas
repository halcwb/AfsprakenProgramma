Attribute VB_Name = "ModMedDisc_Tests"
Option Explicit

' --- Medicament ---
Private Const constGPK As String = "_Glob_MedDisc_GPK_" ' GPK code
Private Const constATC As String = "_Glob_MedDisc_ATC_" ' ATC code
Private Const constGeneric As String = "_Glob_MedDisc_Generic_" ' Generiek
Private Const constVorm As String = "_Glob_MedDisc_Vorm_" ' Medicament vorm
Private Const constConc As String = "_Glob_MedDisc_Sterkte_" ' Sterkte
Private Const constConcUnit As String = "_Glob_MedDisc_SterkteEenh_" ' Sterkte eenheid
Private Const constLabel As String = "_Glob_MedDisc_Etiket_" ' Etiket
Private Const constStandDose As String = "_Glob_MedDisc_StandDose_" ' Dose standaard
Private Const constDoseUnit As String = "_Glob_MedDisc_DoseEenh_" ' Dose eenheid
Private Const constRoute As String = "_Glob_MedDisc_Toed_" ' Toediening route
Private Const constIndic As String = "_Glob_MedDisc_Ind_" ' Indicatie
Private Const constMedicament As String = "AQ"   ' Medicament

' --- Voorschrift ---
Private Const constPRN As String = "_Glob_MedDisc_PRN_" ' PRN
Private Const constPRNText As String = "_Glob_MedDisc_PRNText_" ' PRN tekst
Private Const constFreq As String = "_Glob_MedDisc_Freq_" ' Frequentie
Private Const constDoseQty As String = "_Glob_MedDisc_DoseHoev_" ' Dose hoeveelheid
Private Const constSolNo As String = "_Glob_MedDisc_OplKeuze_" ' Oplossing vloeistof
Private Const constSolVol As String = "_Glob_MedDisc_OplVol_" ' Oplossing volume
Private Const constTime As String = "_Glob_MedDisc_Inloop_" ' Inloop tijd
Private Const constText As String = "_Glob_MedDisc_Opm_" ' Opmerking

Private Const constSetupGewicht As String = "B"
Private Const constSetupMedicament As String = "C"
Private Const constSetupAfronding As String = "D"
Private Const constSetupAfrondingEenheid As String = "E"
Private Const constSetupToediening As String = "F"
Private Const constSetupIndicatie As String = "G"
Private Const constSetupFreq As String = "H"
Private Const constSetupHoeveelheid As String = "I"
Private Const constSetupOplossing As String = "J"
Private Const constSetupOplossingHoeveelheid As String = "K"
Private Const constSetupTijd As String = "L"
Private Const constSetupOpmerking As String = "M"
Private Const constSetupPRN As String = "N"
Private Const constSetupPRNTekst As String = "O"
Private Const constActDosering As String = "R"
Private Const constActConcentratie As String = "S"

Private Const constVerw As String = "AL"
Private Const constMedCount As Integer = 30
Private Const constMedTime As Integer = 25

Public Sub Test_MedDisc()

    Dim wbkTests As Workbook
    Dim shtTests As Worksheet

    Dim objForm As ClassFormularium
    Dim objMed As ClassMedicatieDisc
    Dim colMed As Collection
    Dim intN As Integer
    Dim intC As Integer
    Dim strN As String
    Dim colRoute As Collection
    Dim strRoute As String
    Dim strIndic As String
    Dim colIndic As Collection
    Dim intTime  As Integer
    
    Set objForm = New ClassFormularium
    Set colMed = objForm.GetMedicamenten(False)
    
    Set wbkTests = Workbooks.Open(WbkAfspraken.Path & "/tests/Tests.xlsx")
    Set shtTests = wbkTests.Sheets("DiscMed")
    
    On Error GoTo TestError
    
    ModProgress.StartProgress "Testing Discontinue Medicatie"
    
    intTime = 1
    intN = 1
    intC = 1
    For Each objMed In colMed
        intN = IIf(intN > constMedCount, 1, intN)
        intTime = IIf(intTime > constMedTime, 1, intTime)
        
        strN = IntNToStrN(intN)
        
        Set colRoute = objMed.GetRouteList()
        If colRoute.Count = 1 Then
            strRoute = colRoute.Item(1)
        Else
            strRoute = vbNullString
        End If
        
        Set colIndic = objMed.GetIndicatieList()
        If colIndic.Count = 1 Then
            strIndic = colIndic.Item(1)
        Else
            strIndic = vbNullString
        End If
        
        ModMedDisc.MedDisc_SetMed objMed, strN
        ModRange.SetRangeValue constFreq & strN, intTime
        ModRange.SetRangeValue constDoseQty & strN, intN
        
        WriteTestResults shtTests, intC, intN
        
        intN = intN + 1
        intTime = intTime + 1
        intC = intC + 1
        
        ModProgress.SetJobPercentage objMed.Etiket, colMed.Count, intC
        
        If intN = 30 Then
            intN = 30
        End If
    Next
    
    
    ModProgress.FinishProgress
    
    wbkTests.Close True
    Set shtTests = Nothing
    Set wbkTests = Nothing
    
    Exit Sub
    
TestError:
    
    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxExclam "Kan tests niet uitvoeren: " & Err.Source & " " & Err.Description
    
    On Error Resume Next
        
    wbkTests.Close True
    Set shtTests = Nothing
    Set wbkTests = Nothing
    
End Sub

Private Sub WriteTestResults(shtTests As Worksheet, ByVal intC As Integer, ByVal intN As Integer)
    
    Dim strN As String
    Dim intT As Integer
    Dim intB As Integer
    
    strN = IntNToStrN(intN)
    intT = intC + 2
    intB = intN + 1
    
    shtTests.Range("A" & intT).Value2 = intC
    
    shtTests.Range(constSetupGewicht & intT).Value2 = ModPatient.GetGewichtFromRange()
    shtTests.Range(constSetupMedicament & intT).Value2 = shtGlobBerMedDisc.Range(constMedicament & intB).Value2
    shtTests.Range(constSetupAfronding & intT).Value2 = ModRange.GetRangeValue(constStandDose & strN, vbNullString)
    shtTests.Range(constSetupAfrondingEenheid & intT).Value2 = ModRange.GetRangeValue(constDoseUnit & strN, vbNullString)
    shtTests.Range(constSetupToediening & intT).Value2 = ModRange.GetRangeValue(constRoute & strN, vbNullString)
    shtTests.Range(constSetupIndicatie & intT).Value2 = ModRange.GetRangeValue(constIndic & strN, vbNullString)
    shtTests.Range(constSetupFreq & intT).Value2 = shtGlobBerMedDisc.Range("X" & intB).Value2
    shtTests.Range(constSetupHoeveelheid & intT).Value2 = shtGlobBerMedDisc.Range("Y" & intB).Value2
    shtTests.Range(constSetupOplossing & intT).Value2 = shtGlobBerMedDisc.Range("AA" & intB).Value2
    shtTests.Range(constSetupOplossingHoeveelheid & intT).Value2 = shtGlobBerMedDisc.Range("O" & intB).Value2
    
    shtTests.Range(constActDosering & intT).Value2 = shtGlobBerMedDisc.Range("AN" & intB).Value2
    
End Sub

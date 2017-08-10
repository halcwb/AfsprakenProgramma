Attribute VB_Name = "ModNeoInfB_Tests"
Option Explicit

Public Const CONST_TEST_ERROR As Long = vbObjectError + 1

Private Const constTestStart As Integer = 3
Private Const constTestCount As Integer = 16
Private Const constTestNum As String = "A"
Private Const constSetupGewicht As String = "B"
Private Const constSetupMedicament As String = "C"
Private Const constSetupHoeveelheid As String = "D"
Private Const constSetupOplosmiddel As String = "E"
Private Const constSetupOploshoeveelheid As String = "F"
Private Const constSetupInfuusStand As String = "G"

Private Const constActGewicht As String = "T"
Private Const constActMedicament As String = "U"
Private Const constActHoeveelheid As String = "V"
Private Const constActEenheid As String = "W"
Private Const constActOplosmiddel As String = "X"
Private Const constActTotaalVolume As String = "Y"
Private Const constActInfuusStand As String = "Z"
Private Const constActDosis As String = "AA"
Private Const constActNormaalWaarde As String = "AB"
Private Const constActInloopTijd As String = "AC"
Private Const constActMedVolume As String = "AD"
Private Const constActOplVolume As String = "AE"

Private Const constGewicht As String = "_Pat_Gewicht"
Private Const constMedicament As String = "Var_Neo_InfB_Cont_MedKeuze_"
Private Const constHoeveelheid As String = "Var_Neo_InfB_Cont_MedSterkte_"
Private Const constOplosmiddel As String = "Var_Neo_InfB_Cont_Oplossing_"
Private Const constOploshoeveelheid As String = "Var_Neo_InfB_Cont_OplHoev_"
Private Const constInfuusStand As String = "Var_Neo_InfB_Cont_Stand_"

Private Const constTblMedIV As String = "Tbl_Neo_MedIV"
Private Const constTblOpl As String = "Tbl_Neo_OplVlst"
Private Const constTblVerwacht As String = "T3:AE102"
Private Const constTblTekst As String = "AS3:AT102"

Private Const constAfsprTekst As String = "AS"
Private Const constEtiketTekst As String = "AT"

Private Const constTestResult As String = "AF103"

Public Sub Test_NeoInfB_ContMed()

    Dim wbkTests As Workbook
    Dim shtTests As Worksheet
    Dim intN As Integer
    Dim intM As Integer
    Dim strM As String
    Dim strNo As String
    Dim varVal As Variant
    Dim dblVal As Double
    Dim dblGew As Double
    Dim intMed As Integer
    Dim dblHoev As Double
    Dim intOpl As Integer
    Dim dblOplHoev As Double
    Dim dblStand As Double
    Dim intSpuitN As Integer
    Dim intI As Integer
    Dim strTekst As String
    Dim blnPass As Boolean
    Dim blnShowMsg As Boolean
    
    On Error GoTo Test_NeoInfB_ContMedError
    
    ModProgress.StartProgress "Neo Infuusbrief Continue Medicatie Tests"
    
    'ModPatient.PatientClearAll False, True
    
    Set wbkTests = Workbooks.Open(WbkAfspraken.Path & "/tests/Tests.xlsx")
    Set shtTests = wbkTests.Sheets("NICU_ContMed")
    
    shtTests.Range(constTblVerwacht).ClearContents
    shtTests.Range(constTblTekst).ClearContents

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
        
        ' Totale volume
        dblOplHoev = ModString.StringToDouble(shtTests.Range(constSetupOploshoeveelheid & intN).Value2)
        
        ' Infuus stand
        dblStand = ModString.StringToDouble(shtTests.Range(constSetupInfuusStand & intN).Value2)
        
        ' Voer testcase in
        For intM = 1 To 10
            strM = IIf(intM < 10, "0" & intM, intM)
            blnPass = blnPass And ModRange.SetRangeValue(constMedicament & strM, intMed)
            ChangeMedIV intM
            If dblHoev > 0 Then blnPass = blnPass And ModRange.SetRangeValue(constHoeveelheid & strM, dblHoev * 10)
            If intOpl > 1 Then blnPass = blnPass And ModRange.SetRangeValue(constOplosmiddel & strM, intOpl)
            If dblOplHoev > 0 Then blnPass = blnPass And ModRange.SetRangeValue(constOploshoeveelheid & strM, dblOplHoev)
            If dblStand > 0 Then blnPass = blnPass And ModRange.SetRangeValue(constInfuusStand & strM, dblStand * 10)
        Next
        
        ' Schrijf gewicht
        shtTests.Range(constActGewicht & intN).Value2 = ModRange.GetRangeValue(constGewicht, 0) / 10
        
        
        ' ================ Check medicament ================
        varVal = shtNeoPrtWerkbr.Range("C24").Value2
        For intM = 0 To 9
            'Check werkbrief
            blnPass = blnPass And Equals(varVal, shtNeoPrtWerkbr.Range("C" & intM * 3 + 24).Value2)
            ' Check apotheek print
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            blnPass = blnPass And varVal = shtNeoPrtApoth.Range("D5").Value2
            'Check afspraken print
            If Not (IsEmpty(varVal) Or varVal = "") Then blnPass = blnPass And ModString.ContainsCaseInsensitive(shtNeoPrtAfspr.Range("B" & intM + 33).Value2, varVal)
        Next
        ' Schrijf actuele medicament
        shtTests.Range(constActMedicament & intN).Value2 = varVal
        
        
        ' ================ Check hoeveelheid medicament ================
        varVal = shtNeoPrtWerkbr.Range("E24").Value2
        For intM = 0 To 9
            blnShowMsg = True
            
            'Check werkbrief
            blnPass = blnPass And Equals(varVal, shtNeoPrtWerkbr.Range("E" & intM * 3 + 24).Value2)
            If Not blnPass And blnShowMsg Then
                ModMessage.ShowMsgBoxExclam "Werkbrief print niet goed voor test " & intN - constTestStart & " no: " & intM + 1
                blnShowMsg = False
            End If
            
            ' Check apotheek print
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            blnPass = blnPass And Equals(varVal, IIf(ModRange.GetRangeValue("Var_Neo_InfB_Cont_DubbeleHoev", False), shtNeoPrtApoth.Range("G5").Value2 / 2, shtNeoPrtApoth.Range("G5").Value2))
            If Not blnPass And blnShowMsg Then
                ModMessage.ShowMsgBoxExclam "Apotheek print niet goed voor test " & intN - constTestStart & " no: " & intM + 1
                blnShowMsg = False
            End If
            
            'Check afspraken print
            If Not (IsEmpty(varVal) Or varVal = "" Or varVal = "0") Then
                blnPass = blnPass And ModString.ContainsCaseInsensitive(shtNeoPrtAfspr.Range("B" & intM + 33).Value2, varVal)
                If Not blnPass And blnShowMsg Then
                    ModMessage.ShowMsgBoxExclam "Afspraken print niet goed voor test " & intN - constTestStart & " no: " & intM + 1
                    blnShowMsg = False
                End If
            Else
                blnPass = blnPass
            End If
        Next
        ' Schrijf actuele medicament hoeveelheid
        shtTests.Range(constActHoeveelheid & intN).Value2 = varVal
        
        
        ' ================ Check medicament eenheid ================
        varVal = shtNeoPrtWerkbr.Range("F24").Value2
        For intM = 0 To 9
            'Check werkbrief
            blnPass = blnPass And Equals(varVal, shtNeoPrtWerkbr.Range("F" & intM * 3 + 24).Value2)
            If Not blnPass And blnShowMsg Then
                ModMessage.ShowMsgBoxExclam "Werkbrief print niet goed voor test " & intN - constTestStart & " no: " & intM + 1
                blnShowMsg = False
            End If
            
            ' Check apotheek print
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            blnPass = blnPass And Equals(varVal, shtNeoPrtApoth.Range("H5").Value2)
            If Not blnPass And blnShowMsg Then
                ModMessage.ShowMsgBoxExclam "Apotheek print niet goed voor test " & intN - constTestStart & " no: " & intM + 1
                blnShowMsg = False
            End If
            
            'Check afspraken print
            If Not (IsEmpty(varVal) Or varVal = "" Or varVal = "0") Then
                blnPass = blnPass And ModString.ContainsCaseInsensitive(shtNeoPrtAfspr.Range("B" & intM + 33).Value2, varVal)
                If Not blnPass And blnShowMsg Then
                    ModMessage.ShowMsgBoxExclam "Afspraken print niet goed voor test " & intN - constTestStart & " no: " & intM + 1
                    blnShowMsg = False
                End If
            Else
                blnPass = blnPass
            End If
        Next
        ' Schrijf medicament eenheid
        shtTests.Range(constActEenheid & intN).Value2 = varVal
                
        
        ' ================ Check oplosmiddel ================
        varVal = shtNeoPrtWerkbr.Range("J25").Value2
        For intM = 0 To 9
            'Check werkbrief
            blnPass = blnPass And Equals(varVal, shtNeoPrtWerkbr.Range("J" & intM * 3 + 25).Value2)
            ' Check apotheek print
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            blnPass = blnPass And varVal = shtNeoPrtApoth.Range("J6").Value2
            'Check afspraken print
            If Not (IsEmpty(varVal) Or varVal = "") Then blnPass = blnPass And ModString.ContainsCaseInsensitive(shtNeoPrtAfspr.Range("B" & intM + 33).Value2, varVal)
        Next
        ' Schrijf oplosmiddel
        shtTests.Range(constActOplosmiddel & intN).Value2 = varVal
        
        
        ' ================ Check totaal volume ================
        varVal = AddNumDenumCells(shtNeoPrtWerkbr.Range("M26"), shtNeoPrtWerkbr.Range("N26"))
        For intM = 0 To 9
            blnShowMsg = True
            
            'Check werkbrief
            blnPass = blnPass And Equals(varVal, AddNumDenumCells(shtNeoPrtWerkbr.Range("M" & intM * 3 + 26), shtNeoPrtWerkbr.Range("N" & intM * 3 + 26)))
            
            ' Check apotheek print
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            dblVal = AddNumDenumCells(shtNeoPrtApoth.Range("N7"), shtNeoPrtApoth.Range("O7"))
            dblVal = IIf(ModRange.GetRangeValue("Var_Neo_InfB_Cont_DubbeleHoev", False), dblVal / 2, dblVal)
            blnPass = blnPass And Equals(varVal, dblVal)
            
            'Check afspraken print
            If Not (IsEmpty(varVal) Or varVal = "" Or varVal = "0") Then
                blnPass = blnPass And ModString.ContainsCaseInsensitive(shtNeoPrtAfspr.Range("B" & intM + 33).Value2, varVal)
            Else
                blnPass = blnPass
            End If
        Next
        ' Schrijf actuele oplossing hoeveelheid
        shtTests.Range(constActTotaalVolume & intN).Value2 = varVal
        
        
        '  ================ Check infuus stand ================
        ' ToDo: fix empty test cases
        varVal = shtNeoPrtWerkbr.Range("A24").Value2
        If Not (IsEmpty(varVal) Or varVal = "") Then
            For intM = 0 To 9
                'Check werkbrief
                blnPass = blnPass And Equals(varVal, shtNeoPrtWerkbr.Range("A" & intM * 3 + 24).Value2)
                ' Check apotheek print
                shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
                blnPass = blnPass And varVal = Replace(shtNeoPrtApoth.Range("D6").Value2, " ml/uur", "")
                'Check afspraken print
                If Not (IsEmpty(varVal) Or varVal = "") Then blnPass = blnPass And ModString.ContainsCaseInsensitive(shtNeoPrtAfspr.Range("B" & intM + 33).Value2, varVal)
            Next
        End If
        ' Schrijf infuus stand
        shtTests.Range(constActInfuusStand & intN).Value2 = varVal
        
        
        ' ================ Check dosis ================
        ' ToDo: fix empty test cases
        varVal = Trim(Replace(shtNeoPrtWerkbr.Range("E26").Value2, "= ", ""))
        If Not (IsEmpty(varVal) Or varVal = "") Then
            For intM = 0 To 9
                'Check werkbrief
                blnPass = blnPass And Equals(varVal, Replace(shtNeoPrtWerkbr.Range("E" & intM * 3 + 26).Value2, "= ", ""))
                ' Check apotheek print
                shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
                blnPass = blnPass And Equals(varVal, shtNeoPrtApoth.Range("F6").Value2)
                'Check afspraken print
                If Not (IsEmpty(varVal) Or varVal = "") Then blnPass = blnPass And ModString.ContainsCaseInsensitive(shtNeoPrtAfspr.Range("B" & intM + 33).Value2, varVal)
            Next
        End If
        ' Schrijf dosis
        shtTests.Range(constActDosis & intN).Value2 = varVal
        
        
        ' ================ Check normaal waarde ================
        varVal = shtNeoGuiInfB.Range("O28").Value2
        For intM = 0 To 9
            ' Check apotheek print
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            blnPass = blnPass And Equals(varVal, shtNeoPrtApoth.Range("F7").Value2)
            If Not blnPass And blnShowMsg Then
                ModMessage.ShowMsgBoxExclam "Apotheek print niet goed voor test " & intN - constTestStart & " no: " & intM + 1
                blnShowMsg = False
            End If
            
            'Check infuusbrief gui
            blnPass = blnPass And Equals(varVal, shtNeoGuiInfB.Range("O" & intM + 28).Value2)
        Next
        ' Schrijf normaal waarde
        shtTests.Range(constActNormaalWaarde & intN).Value2 = varVal
        
        
        ' ================ Check inloop tijd ================
        varVal = shtNeoGuiInfB.Range("R28").Value2
        For intM = 0 To 9
            'Check infuusbrief gui
            blnPass = blnPass And Equals(varVal, shtNeoGuiInfB.Range("R" & intM + 28).Value2)
        Next
        ' Schrijf inloop tijd
        shtTests.Range(constActInloopTijd & intN).Value2 = varVal
                
        
        ' ================ Check medicament volume ================
        varVal = AddNumDenumCells(shtNeoPrtWerkbr.Range("M24"), shtNeoPrtWerkbr.Range("N24"))
        For intM = 0 To 9
            'Check werkbrief
            blnPass = blnPass And Equals(varVal, AddNumDenumCells(shtNeoPrtWerkbr.Range("M" & intM * 3 + 24), shtNeoPrtWerkbr.Range("N" & intM * 3 + 24)))
            'Check apotheek brief
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            dblVal = AddNumDenumCells(shtNeoPrtApoth.Range("N5"), shtNeoPrtApoth.Range("O5"))
            dblVal = IIf(ModRange.GetRangeValue("Var_Neo_InfB_Cont_DubbeleHoev", False), dblVal / 2, dblVal)
            blnPass = blnPass And Equals(varVal, dblVal)
        Next
        ' Schrijf medicament volume
        shtTests.Range(constActMedVolume & intN).Value2 = varVal
                
        
        ' ================ Check oplossing volume ================
        varVal = AddNumDenumCells(shtNeoPrtWerkbr.Range("M25"), shtNeoPrtWerkbr.Range("N25"))
        For intM = 0 To 9
            'Check werkbrief
            blnPass = blnPass And Equals(varVal, AddNumDenumCells(shtNeoPrtWerkbr.Range("M" & intM * 3 + 25), shtNeoPrtWerkbr.Range("N" & intM * 3 + 25)))
            'Check apotheek brief
            shtNeoPrtApoth.Range("Var_Neo_PrintApothNo").Value2 = intM + 1
            dblVal = AddNumDenumCells(shtNeoPrtApoth.Range("N6"), shtNeoPrtApoth.Range("O6"))
            dblVal = IIf(ModRange.GetRangeValue("Var_Neo_InfB_Cont_DubbeleHoev", False), dblVal / 2, dblVal)
            blnPass = blnPass And Equals(varVal, dblVal)
        Next
        ' Schrijf oplossing volume
        shtTests.Range(constActOplVolume & intN).Value2 = varVal
        
        
        ' ================ Check Afspraak Tekst ================
        strTekst = shtNeoPrtAfspr.Range("B33").Value
        For intM = 0 To 9
            blnPass = blnPass And Equals(strTekst, shtNeoPrtAfspr.Range("B" & intM + 33))
        Next
        shtTests.Range(constAfsprTekst & intN).Value2 = strTekst
        
        
        ' ================ Check Etiket Tekst ================
        strTekst = vbNullString
        For intM = 0 To 9
            If IsNumeric(shtNeoPrtApoth.Range("A17").Value2) Then
                intSpuitN = shtNeoPrtApoth.Range("A17").Value2
                strTekst = GetEtiketTekst(1)
                For intI = 1 To intSpuitN + 1
                    blnPass = blnPass And Equals(strTekst, GetEtiketTekst(intI))
                    strTekst = GetEtiketTekst(intI)
                Next
            Else
                strTekst = vbNullString
            End If
        Next
        shtTests.Range(constEtiketTekst & intN).Value2 = strTekst
        
        
        If Not blnPass Then
            Err.Raise CONST_TEST_ERROR, "NeoInfB_Tests", "Test no: " & intN - constTestStart & " did not pass"
        End If
        
    Next

    ModProgress.FinishProgress
    
    blnPass = blnPass And shtTests.Range(constTestResult).Value

    If blnPass Then
        ModMessage.ShowMsgBoxInfo "Alle testen geslaagt"
    Else
        ModMessage.ShowMsgBoxExclam "Niet alle testen geslaagt: " & intN - constTestStart
    End If
    
    wbkTests.Close True
    Set shtTests = Nothing
    Set wbkTests = Nothing
    
    Exit Sub
    
Test_NeoInfB_ContMedError:

    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxExclam "Kan tests niet uitvoeren: " & Err.Source & " " & Err.Description
    
    On Error Resume Next
        
    wbkTests.Close True
    Set shtTests = Nothing
    Set wbkTests = Nothing
End Sub

Private Function Equals(ByVal strVal1 As Variant, ByVal strVal2 As Variant)

    Dim strVal_1 As String
    Dim strVal_2 As String
    
    strVal_1 = Trim(strVal1)
    strVal_2 = Trim(strVal2)
    
    strVal_1 = IIf(strVal_1 = "0", vbNullString, strVal_1)
    strVal_2 = IIf(strVal_2 = "0", vbNullString, strVal_2)
    
    If Not (strVal_1 = strVal_2) Then MsgBox "Not Equal: " & strVal_1 & ", " & strVal_2
    
    Equals = strVal_1 = strVal_2

End Function

Private Sub Test_Equals()
    
    Dim varVal
    
    MsgBox Equals("0", varVal)

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

Private Function AddNumDenumCells(objRange1 As Range, objRange2 As Range) As Double

    Dim strVal1 As String
    Dim strVal2 As String
    
    strVal1 = objRange1.Value
    strVal1 = IIf(strVal1 = vbNullString, "0", strVal1)
    strVal2 = objRange2.Value
    strVal2 = IIf(strVal2 = vbNullString, "0", strVal2)

    AddNumDenumCells = ModString.StringToDouble(strVal1) + ModString.StringToDouble(strVal2) / 100

End Function

Private Sub Test_AddNumDenumCells()

    MsgBox AddNumDenumCells(shtNeoPrtApoth.Range("N5"), shtNeoPrtApoth.Range("O5"))

End Sub

Private Function GetEtiketTekst(ByVal intNum As Integer) As String

    Dim strTekst As String
    Dim strVal As String
    Dim objRange As Range
    Dim objCell As Range
    
    On Error Resume Next
    
    Select Case intNum
    
    Case 1
    Set objRange = shtNeoPrtApoth.Range("A45").Resize(6, 5)
    
    Case 2
    Set objRange = shtNeoPrtApoth.Range("G45").Resize(6, 5)
    
    Case 3
    Set objRange = shtNeoPrtApoth.Range("M45").Resize(6, 5)
    
    Case 4
    Set objRange = shtNeoPrtApoth.Range("A56").Resize(6, 5)
    
    Case 5
    Set objRange = shtNeoPrtApoth.Range("G56").Resize(6, 5)
    
    Case 6
    Set objRange = shtNeoPrtApoth.Range("M56").Resize(6, 5)
        
    End Select
    
    For Each objCell In objRange
        strVal = Trim(objCell.Value)
        strVal = IIf(strVal = "0", vbNullString, strVal)
        strTekst = strTekst & " " & strVal
    Next
    
    GetEtiketTekst = strTekst

End Function

Private Sub Test_GetEtiketTekst()

    MsgBox GetEtiketTekst(3)

End Sub


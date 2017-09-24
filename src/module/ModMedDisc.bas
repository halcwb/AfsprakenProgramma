Attribute VB_Name = "ModMedDisc"
Option Explicit

Private Const constFreqTable = "Tbl_Glob_MedFreq"
Private Const constTblMedOpdr As String = "Tbl_Glob_MedOpdr"

' --- Medicament ---
Private Const constGPK As String = "_Glob_MedDisc_GPK_"                 ' GPK code
Private Const constATC As String = "_Glob_MedDisc_ATC_"                 ' ATC code
Private Const constGeneric As String = "_Glob_MedDisc_Generic_"         ' Generiek
Private Const constVorm As String = "_Glob_MedDisc_Vorm_"               ' Medicament vorm
Private Const constConc As String = "_Glob_MedDisc_Sterkte_"            ' Sterkte
Private Const constConcUnit As String = "_Glob_MedDisc_SterkteEenh_"    ' Sterkte eenheid
Private Const constLabel As String = "_Glob_MedDisc_Etiket_"            ' Etiket
Private Const constStandDose As String = "_Glob_MedDisc_StandDose_"     ' Dose standaard
Private Const constDoseUnit As String = "_Glob_MedDisc_DoseEenh_"       ' Dose eenheid
Private Const constRoute As String = "_Glob_MedDisc_Toed_"              ' Toediening route
Private Const constIndic As String = "_Glob_MedDisc_Ind_"               ' Indicatie

' --- Voorschrift ---
Private Const constPRN As String = "_Glob_MedDisc_PRN_"                 ' PRN
Private Const constPRNText As String = "_Glob_MedDisc_PRNText_"         ' PRN tekst
Private Const constFreq As String = "_Glob_MedDisc_Freq_"               ' Frequentie
Private Const constDoseQty As String = "_Glob_MedDisc_DoseHoev_"        ' Dose hoeveelheid
Private Const constSolNo As String = "_Glob_MedDisc_OplKeuze_"          ' Oplossing vloeistof
Private Const constSolVol As String = "_Glob_MedDisc_OplVol_"           ' Oplossing volume
Private Const constTime As String = "_Glob_MedDisc_Inloop_"             ' Inloop tijd
Private Const constText As String = "_Glob_MedDisc_Opm_"                ' Opmerking

'--- Medicatie Controle ---
Private Const constFreqList As String = "_Glob_MedDisc_FreqList_"       ' Lijst van frequenties
Private Const constNormDose As String = "_Glob_MedDisc_NormDose_"       ' Normale dosering
Private Const constMinDose As String = "_Glob_MedDisc_Onder_"           ' Dagdosering ondergrens
Private Const constMaxDose As String = "_Glob_MedDisc_Boven_"           ' Dag dosering bovengrens
Private Const constAbsDose As String = "_Glob_MedDisc_MaxDose_"         ' Maximale dosering

Private Const constMaxConc As String = "_Glob_MedDisc_MaxConc_"         ' Maximale concentratie
Private Const constOplVlst As String = "_Glob_MedDisc_OplVlst_"         ' Verplichte oplos vloeistof
Private Const constMinTijd As String = "_Glob_MedDisc_MinTijd_"         ' Minimale inloop tijd

Private Const constFreqText As String = "Var_MedDisc_FreqText_"

Private Const constVerw As String = "AL"
Private Const constMedCount As Integer = 30


' Copy paste function cannot be reused because of private clear method
Private Sub ShowPickList(colTbl As Collection, ByVal strRange As String, ByVal intStart As Integer, ByVal intMax As Integer)

    Dim frmPickList As FormMedDiscPickList
    Dim intN As Integer
    Dim strN As String
    Dim strKeuze As String
    
    Set frmPickList = New FormMedDiscPickList
    frmPickList.LoadMedicamenten colTbl
    
    For intN = 1 To intMax
        strN = IIf(intMax > 9, IntNToStrN(intN), intN)
        strKeuze = ModRange.GetRangeValue(strRange & strN, 1)
        If Not strKeuze = vbNullString Then frmPickList.SelectMedicament strKeuze
    Next intN
    
    frmPickList.Show
    
    If frmPickList.GetAction = vbNullString Then
    
        For intN = 1 To intMax                   ' First remove nonselected items
            strN = IntNToStrN(intN)
            strKeuze = ModRange.GetRangeValue(strRange & strN, 1)
            If Not strKeuze = vbNullString Then
                If frmPickList.IsMedicamentSelected(strKeuze) Then
                    frmPickList.UnselectMedicament (strKeuze)
                Else
                    Clear intN                   ' Remove is specific to PedContIV replace with appropriate sub when copy paste
                End If
            End If
        Next intN
        
        Do While frmPickList.HasSelectedMedicamenten() ' Then add selected items
            For intN = 1 To intMax
                strN = IntNToStrN(intN)
                strKeuze = ModRange.GetRangeValue(strRange & strN, 1)
                If strKeuze = vbNullString Then
                    strKeuze = frmPickList.GetFirstSelectedMedicament(True)
                    ModRange.SetRangeValue strRange & strN, strKeuze
                    Exit For
                End If
            Next intN
        Loop
    
    End If
    
    Set frmPickList = Nothing


End Sub

Public Sub MedDisc_ShowPickList()

    Dim objForm As ClassFormularium
    Dim objGenCol As Collection
    Dim objTable As Range
    Dim varGen As Variant
    
    Set objForm = Formularium_GetFormularium
    Set objGenCol = New Collection
    Set objTable = ModRange.GetRange(constTblMedOpdr)
    
    For Each varGen In objTable
        varGen = CStr(varGen)
        If Not varGen = vbNullString Then
            varGen = Split(varGen, " ")(0)
            varGen = LCase(varGen)
            If CollectionContains(varGen, objForm.GetGenerieken()) Then
                If Not CollectionContains(varGen, objGenCol) Then objGenCol.Add varGen
            End If
        End If
    Next
    
    ShowPickList objGenCol, constGeneric, 1, constMedCount
    
End Sub

Private Sub Clear(ByVal intN As Integer)

    Dim strN As String
    
    strN = IntNToStrN(intN)

    ModRange.SetRangeValue constGPK & strN, 0
    ModRange.SetRangeValue constATC & strN, vbNullString
    ModRange.SetRangeValue constGeneric & strN, vbNullString
    ModRange.SetRangeValue constVorm & strN, vbNullString
    ModRange.SetRangeValue constConc & strN, 0
    ModRange.SetRangeValue constConcUnit & strN, vbNullString
    ModRange.SetRangeValue constLabel & strN, vbNullString
    ModRange.SetRangeValue constStandDose & strN, 0
    ModRange.SetRangeValue constDoseUnit & strN, vbNullString
    ModRange.SetRangeValue constRoute & strN, vbNullString
    ModRange.SetRangeValue constIndic & strN, vbNullString
    
    ModRange.SetRangeValue constPRN & strN, False
    ModRange.SetRangeValue constPRNText & strN, vbNullString
    ModRange.SetRangeValue constDoseQty & strN, 0
    ModRange.SetRangeValue constFreq & strN, 1
    ModRange.SetRangeValue constSolNo & strN, 1
    ModRange.SetRangeValue constSolVol & strN, 0
    ModRange.SetRangeValue constTime & strN, 0
    ModRange.SetRangeValue constText & strN, vbNullString
    
    ModRange.SetRangeValue constFreqList & strN, vbNullString
    ModRange.SetRangeValue constNormDose & strN, 0
    ModRange.SetRangeValue constMinDose & strN, 0
    ModRange.SetRangeValue constMaxDose & strN, 0
    ModRange.SetRangeValue constAbsDose & strN, 0
    
    ModRange.SetRangeValue constMaxConc & strN, 0
    ModRange.SetRangeValue constOplVlst & strN, vbNullString
    ModRange.SetRangeValue constMinTijd & strN, 0

End Sub

Public Sub MedDisc_Clear_01()

    Clear 1

End Sub

Public Sub MedDisc_Clear_02()

    Clear 2

End Sub

Public Sub MedDisc_Clear_03()

    Clear 3

End Sub

Public Sub MedDisc_Clear_04()

    Clear 4

End Sub

Public Sub MedDisc_Clear_05()

    Clear 5

End Sub

Public Sub MedDisc_Clear_06()

    Clear 6

End Sub

Public Sub MedDisc_Clear_07()

    Clear 7

End Sub

Public Sub MedDisc_Clear_08()

    Clear 8

End Sub

Public Sub MedDisc_Clear_09()

    Clear 9

End Sub

Public Sub MedDisc_Clear_10()

    Clear 10

End Sub

Public Sub MedDisc_Clear_11()

    Clear 11

End Sub

Public Sub MedDisc_Clear_12()

    Clear 12

End Sub

Public Sub MedDisc_Clear_13()

    Clear 13

End Sub

Public Sub MedDisc_Clear_14()

    Clear 14

End Sub

Public Sub MedDisc_Clear_15()

    Clear 15

End Sub

Public Sub MedDisc_Clear_16()

    Clear 16

End Sub

Public Sub MedDisc_Clear_17()

    Clear 17

End Sub

Public Sub MedDisc_Clear_18()

    Clear 18

End Sub

Public Sub MedDisc_Clear_19()

    Clear 19

End Sub

Public Sub MedDisc_Clear_20()

    Clear 20

End Sub

Public Sub MedDisc_Clear_21()

    Clear 21

End Sub

Public Sub MedDisc_Clear_22()

    Clear 22

End Sub

Public Sub MedDisc_Clear_23()

    Clear 23

End Sub

Public Sub MedDisc_Clear_24()

    Clear 24

End Sub

Public Sub MedDisc_Clear_25()

    Clear 25

End Sub

Public Sub MedDisc_Clear_26()

    Clear 26

End Sub

Public Sub MedDisc_Clear_27()

    Clear 27

End Sub

Public Sub MedDisc_Clear_28()

    Clear 28

End Sub

Public Sub MedDisc_Clear_29()

    Clear 29

End Sub

Public Sub MedDisc_Clear_30()

    Clear 30

End Sub

Public Function GetMedicationFreqs() As Dictionary

    Dim objTable As Range
    Dim dictFreq As Dictionary
    Dim strFreq As String
    Dim dblFactor As Double
    Dim intN As Integer
    Dim intC As Integer
    
    Set objTable = ModRange.GetRange(constFreqTable)
    Set dictFreq = New Dictionary
    
    intC = objTable.Rows.Count
    For intN = 2 To intC
        strFreq = objTable.Cells(intN, 1).Value2
        dblFactor = objTable.Cells(intN, 2).Value2
        dictFreq.Add strFreq, dblFactor
    Next
    
    Set GetMedicationFreqs = dictFreq

End Function

Private Sub Test_GetMedicationFreqs()

    GetMedicationFreqs

End Sub


Private Sub MedicamentInvoeren(ByVal intN As Integer)

    Dim objMed As ClassMedicatieDisc
    Dim objForm As FormMedDisc
    Dim strN As String
    Dim blnLoad As Boolean
      
    strN = IntNToStrN(intN)
    
    Set objForm = New FormMedDisc
    With objForm
    
        .Caption = "Kies een medicament voor regel " & strN
        
        blnLoad = False
        If ModRange.GetRangeValue(constGPK & strN, 0) > 0 Then ' Drug from formularium
            blnLoad = .LoadGPK(CStr(ModRange.GetRangeValue(constGPK & strN, vbNullString)))
        End If
        
        If Not blnLoad Then                      ' Manually entered drug
            .SetNoFormMed
            .cboGeneriek.Text = ModRange.GetRangeValue(constGeneric & strN, vbNullString)
            .SetComboBoxIfNotEmpty .cboVorm, ModRange.GetRangeValue(constVorm & strN, vbNullString)
            .SetTextBoxIfNotEmpty .txtSterkte, ModRange.GetRangeValue(constConc & strN, 0)
            .SetComboBoxIfNotEmpty .cboSterkteEenheid, ModRange.GetRangeValue(constConcUnit & strN, vbNullString)
            .SetComboBoxIfNotEmpty .cboDosisEenheid, ModRange.GetRangeValue(constDoseUnit & strN, vbNullString)
        End If
        
        ' Edited details
        .SetComboBoxIfNotEmpty .cboDosisEenheid, ModRange.GetRangeValue(constDoseUnit & strN, vbNullString)
        .SetComboBoxIfNotEmpty .cboRoute, ModRange.GetRangeValue(constRoute & strN, vbNullString)
        .SetComboBoxIfNotEmpty .cboIndicatie, ModRange.GetRangeValue(constIndic & strN, vbNullString)
        .SetComboBoxIfNotEmpty .cboFreq, ModRange.GetRangeValue(constFreqText & strN, vbNullString)
        
        .SetTextBoxIfNotEmpty .txtNormDose, ModRange.GetRangeValue(constNormDose & strN, vbNullString)
        .SetTextBoxIfNotEmpty .txtMinDose, ModRange.GetRangeValue(constMinDose & strN, vbNullString)
        .SetTextBoxIfNotEmpty .txtMaxDose, ModRange.GetRangeValue(constMaxDose & strN, vbNullString)
        .SetTextBoxIfNotEmpty .txtAbsMax, ModRange.GetRangeValue(constAbsDose & strN, vbNullString)
        
        .Show
        
        If .GetClickedButton = "OK" Then
            If .HasSelectedMedicament() Then
            
                Set objMed = .GetSelectedMedicament()
                ' -- Medicament --
                MedDisc_SetMed objMed, strN
                
            End If

        Else
            If .GetClickedButton = "Clear" Then
                Clear intN
            End If
        End If
    End With
    
    Set objForm = Nothing

End Sub

Private Function CalculateOplossingVolume(ByVal dblSterkte As Double, ByVal dblMaxConc As Double) As Double
    
    Dim dblVolume
    
    If dblMaxConc > 0 Then
        ' maxoonc = sterkte / volume -> volume = sterkte / maxconc
        dblVolume = dblSterkte / dblMaxConc
    End If

    CalculateOplossingVolume = dblVolume

End Function


Public Sub MedDisc_SetMed(objMed As ClassMedicatieDisc, strN As String)
    
    Dim intFreq As Integer
    Dim varFreq As Variant
    Dim dictFreq As Dictionary
    Dim intDoseQty As Integer
    Dim dblOplVol As Double
    
    ModRange.SetRangeValue constGPK & strN, objMed.GPK
    ModRange.SetRangeValue constATC & strN, objMed.ATC
    ModRange.SetRangeValue constGeneric & strN, objMed.Generiek
    ModRange.SetRangeValue constVorm & strN, objMed.Vorm
    ModRange.SetRangeValue constConc & strN, objMed.Sterkte
    ModRange.SetRangeValue constConcUnit & strN, objMed.SterkteEenheid
    ModRange.SetRangeValue constLabel & strN, objMed.Etiket
    ModRange.SetRangeValue constStandDose & strN, objMed.DeelDose
    ModRange.SetRangeValue constDoseUnit & strN, objMed.DoseEenheid
    ModRange.SetRangeValue constRoute & strN, objMed.Route
    ModRange.SetRangeValue constIndic & strN, objMed.Indicatie
    
    ModRange.SetRangeValue constNormDose & strN, objMed.NormDose
    ModRange.SetRangeValue constMinDose & strN, objMed.MinDose
    ModRange.SetRangeValue constMaxDose & strN, objMed.MaxDose
    ModRange.SetRangeValue constAbsDose & strN, objMed.AbsDose
      
    ModRange.SetRangeValue constMaxConc & strN, objMed.MaxConc
    ModRange.SetRangeValue constOplVlst & strN, objMed.OplVlst
    ModRange.SetRangeValue constMinTijd & strN, objMed.MinTijd
    
    If objMed.OplVlst = "NaCl" Then
        ModRange.SetRangeValue constSolNo & strN, 2
    ElseIf objMed.OplVlst = "glucose" Then
        ModRange.SetRangeValue constSolNo & strN, 4
    End If
    
    If Not objMed.Freq = vbNullString Then
        Set dictFreq = GetMedicationFreqs()
        intFreq = 2
        For Each varFreq In dictFreq
            If varFreq = objMed.Freq Then Exit For
            intFreq = intFreq + 1
        Next
        ModRange.SetRangeValue constFreq & strN, intFreq
    End If
    
    ModRange.SetRangeValue constFreqList & strN, objMed.GetFreqListString
    
    If Not objMed.CalcDose = 0 And Not intFreq < 2 Then
        intDoseQty = objMed.CalcDose * ModPatient.GetGewichtFromRange() / ModExcel.Excel_Index(constFreqTable, intFreq, 2) / objMed.DeelDose
        ModRange.SetRangeValue constDoseQty & strN, intDoseQty
        
        dblOplVol = CalculateOplossingVolume(intDoseQty * objMed.DeelDose, objMed.MaxConc)
    
        If Not dblOplVol = 0 Then
            ModRange.SetRangeValue constSolVol & strN, dblOplVol
        End If
    
    End If

End Sub

Public Sub MedDisc_EnterMed_01()

    MedicamentInvoeren 1

End Sub

Public Sub MedDisc_EnterMed_02()

    MedicamentInvoeren 2

End Sub

Public Sub MedDisc_EnterMed_03()

    MedicamentInvoeren 3

End Sub

Public Sub MedDisc_EnterMed_04()

    MedicamentInvoeren 4

End Sub

Public Sub MedDisc_EnterMed_05()

    MedicamentInvoeren 5

End Sub

Public Sub MedDisc_EnterMed_06()

    MedicamentInvoeren 6

End Sub

Public Sub MedDisc_EnterMed_07()

    MedicamentInvoeren 7

End Sub

Public Sub MedDisc_EnterMed_08()

    MedicamentInvoeren 8

End Sub

Public Sub MedDisc_EnterMed_09()

    MedicamentInvoeren 9

End Sub

Public Sub MedDisc_EnterMed_10()

    MedicamentInvoeren 10

End Sub

Public Sub MedDisc_EnterMed_11()

    MedicamentInvoeren 11

End Sub

Public Sub MedDisc_EnterMed_12()

    MedicamentInvoeren 12

End Sub

Public Sub MedDisc_EnterMed_13()

    MedicamentInvoeren 13

End Sub

Public Sub MedDisc_EnterMed_14()

    MedicamentInvoeren 14

End Sub

Public Sub MedDisc_EnterMed_15()

    MedicamentInvoeren 15

End Sub

Public Sub MedDisc_EnterMed_16()

    MedicamentInvoeren 16

End Sub

Public Sub MedDisc_EnterMed_17()

    MedicamentInvoeren 17

End Sub

Public Sub MedDisc_EnterMed_18()

    MedicamentInvoeren 18

End Sub

Public Sub MedDisc_EnterMed_19()

    MedicamentInvoeren 19

End Sub

Public Sub MedDisc_EnterMed_20()

    MedicamentInvoeren 20

End Sub

Public Sub MedDisc_EnterMed_21()

    MedicamentInvoeren 21

End Sub

Public Sub MedDisc_EnterMed_22()

    MedicamentInvoeren 22

End Sub

Public Sub MedDisc_EnterMed_23()

    MedicamentInvoeren 23

End Sub

Public Sub MedDisc_EnterMed_24()

    MedicamentInvoeren 24

End Sub

Public Sub MedDisc_EnterMed_25()

    MedicamentInvoeren 25

End Sub

Public Sub MedDisc_EnterMed_26()

    MedicamentInvoeren 26

End Sub

Public Sub MedDisc_EnterMed_27()

    MedicamentInvoeren 27

End Sub

Public Sub MedDisc_EnterMed_28()

    MedicamentInvoeren 28

End Sub

Public Sub MedDisc_EnterMed_29()

    MedicamentInvoeren 29

End Sub

Public Sub MedDisc_EnterMed_30()

    MedicamentInvoeren 30

End Sub

Private Sub OpmMedDisc(ByVal intN As Integer)
    
    Dim frmOpmerking As FormOpmerking
    Dim strRange As String
    
    Set frmOpmerking = New FormOpmerking
    
    strRange = constText
    strRange = constText & IntNToStrN(intN)

    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(strRange, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue strRange, frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub MedDisc_EnterText_01()
    
    OpmMedDisc 1
    
End Sub

Public Sub MedDisc_EnterText_02()
    
    OpmMedDisc 2

End Sub

Public Sub MedDisc_EnterText_03()
    
    OpmMedDisc 3

End Sub

Public Sub MedDisc_EnterText_04()
    
    OpmMedDisc 4

End Sub

Public Sub MedDisc_EnterText_05()
    
    OpmMedDisc 5

End Sub

Public Sub MedDisc_EnterText_06()
    
    OpmMedDisc 6

End Sub

Public Sub MedDisc_EnterText_07()
    
    OpmMedDisc 7

End Sub

Public Sub MedDisc_EnterText_08()
    
    OpmMedDisc 8

End Sub

Public Sub MedDisc_EnterText_09()
    
    OpmMedDisc 9

End Sub

Public Sub MedDisc_EnterText_10()
    
    OpmMedDisc 10

End Sub

Public Sub MedDisc_EnterText_11()
    
    OpmMedDisc 11

End Sub

Public Sub MedDisc_EnterText_12()
    
    OpmMedDisc 12

End Sub

Public Sub MedDisc_EnterText_13()
    
    OpmMedDisc 13

End Sub

Public Sub MedDisc_EnterText_14()
    
    OpmMedDisc 14

End Sub

Public Sub MedDisc_EnterText_15()
    
    OpmMedDisc 15

End Sub

Public Sub MedDisc_EnterText_16()
    
    OpmMedDisc 16

End Sub

Public Sub MedDisc_EnterText_17()
    
    OpmMedDisc 17

End Sub

Public Sub MedDisc_EnterText_18()
    
    OpmMedDisc 18

End Sub

Public Sub MedDisc_EnterText_19()
    
    OpmMedDisc 19

End Sub

Public Sub MedDisc_EnterText_20()
    
    OpmMedDisc 20

End Sub

Public Sub MedDisc_EnterText_21()
    
    OpmMedDisc 21

End Sub

Public Sub MedDisc_EnterText_22()
    
    OpmMedDisc 22

End Sub

Public Sub MedDisc_EnterText_23()
    
    OpmMedDisc 23

End Sub

Public Sub MedDisc_EnterText_24()
    
    OpmMedDisc 24

End Sub

Public Sub MedDisc_EnterText_25()
    
    OpmMedDisc 25

End Sub

Public Sub MedDisc_EnterText_26()
    
    OpmMedDisc 26

End Sub

Public Sub MedDisc_EnterText_27()
    
    OpmMedDisc 27

End Sub

Public Sub MedDisc_EnterText_28()
    
    OpmMedDisc 28

End Sub

Public Sub MedDisc_EnterText_29()
    
    OpmMedDisc 29

End Sub

Public Sub MedDisc_EnterText_30()
    
    OpmMedDisc 30

End Sub

Public Function GetFormulariumDatabasePath() As String
    
    GetFormulariumDatabasePath = WbkAfspraken.Path & ModSetting.GetFormDbDir()

End Function

Private Sub Test_GetFormulariumDatabasePath()

    MsgBox GetFormulariumDatabasePath()

End Sub

Private Sub OpenPRNForm(ByVal intN As Integer)

    Dim frmPrn As FormPRN
    Dim strTitle As String
    Dim strVerw As String
    
    strVerw = shtGlobBerMedDisc.Range(constVerw & intN + 1)
    If strVerw = vbNullString Then Exit Sub
    
    Set frmPrn = New FormPRN
    strTitle = "PRN voor medicament " & intN & " instellen"
    frmPrn.SetMedicamentNo intN
    frmPrn.Caption = strTitle
    frmPrn.Show
    
    Set frmPrn = Nothing

End Sub

Public Sub MedDisc_PRN_01()
    
    OpenPRNForm 1
    
End Sub

Public Sub MedDisc_PRN_02()
    
    OpenPRNForm 2

End Sub

Public Sub MedDisc_PRN_03()
    
    OpenPRNForm 3

End Sub

Public Sub MedDisc_PRN_04()
    
    OpenPRNForm 4

End Sub

Public Sub MedDisc_PRN_05()
    
    OpenPRNForm 5

End Sub

Public Sub MedDisc_PRN_06()
    
    OpenPRNForm 6

End Sub

Public Sub MedDisc_PRN_07()
    
    OpenPRNForm 7

End Sub

Public Sub MedDisc_PRN_08()
    
    OpenPRNForm 8

End Sub

Public Sub MedDisc_PRN_09()
    
    OpenPRNForm 9

End Sub

Public Sub MedDisc_PRN_10()
    
    OpenPRNForm 10

End Sub

Public Sub MedDisc_PRN_11()
    
    OpenPRNForm 11

End Sub

Public Sub MedDisc_PRN_12()
    
    OpenPRNForm 12

End Sub

Public Sub MedDisc_PRN_13()
    
    OpenPRNForm 13

End Sub

Public Sub MedDisc_PRN_14()
    
    OpenPRNForm 14

End Sub

Public Sub MedDisc_PRN_15()
    
    OpenPRNForm 15

End Sub

Public Sub MedDisc_PRN_16()
    
    OpenPRNForm 16

End Sub

Public Sub MedDisc_PRN_17()
    
    OpenPRNForm 17

End Sub

Public Sub MedDisc_PRN_18()
    
    OpenPRNForm 18

End Sub

Public Sub MedDisc_PRN_19()
    
    OpenPRNForm 19

End Sub

Public Sub MedDisc_PRN_20()
    
    OpenPRNForm 20

End Sub

Public Sub MedDisc_PRN_21()
    
    OpenPRNForm 21

End Sub

Public Sub MedDisc_PRN_22()
    
    OpenPRNForm 22

End Sub

Public Sub MedDisc_PRN_23()
    
    OpenPRNForm 23

End Sub

Public Sub MedDisc_PRN_24()
    
    OpenPRNForm 24

End Sub

Public Sub MedDisc_PRN_25()
    
    OpenPRNForm 25

End Sub

Public Sub MedDisc_PRN_26()
    
    OpenPRNForm 26

End Sub

Public Sub MedDisc_PRN_27()
    
    OpenPRNForm 27

End Sub

Public Sub MedDisc_PRN_28()
    
    OpenPRNForm 28

End Sub

Public Sub MedDisc_PRN_29()
    
    OpenPRNForm 29

End Sub

Public Sub MedDisc_PRN_30()
    
    OpenPRNForm 30

End Sub


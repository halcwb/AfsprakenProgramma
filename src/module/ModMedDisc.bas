Attribute VB_Name = "ModMedDisc"
Option Explicit

Private Const constFreqTable As String = "Tbl_Glob_MedFreq"
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
Private Const constHasDose As String = "_Glob_MedDisc_HasDose_"         ' Prescription has dose
Private Const constATCGroup As String = "_Glob_MedDisc_Group_"          ' ATC groep

' --- Voorschrift ---
Private Const constPRN As String = "_Glob_MedDisc_PRN_"                 ' PRN
Private Const constPRNText As String = "_Glob_MedDisc_PRNText_"         ' PRN tekst
Private Const constFreq As String = "_Glob_MedDisc_Freq_"               ' Frequentie
Private Const constDoseSubst As String = "_Glob_MedDisc_DoseSubst_"     ' Dose stof
Private Const constDoseQty As String = "_Glob_MedDisc_DoseHoev_"        ' Dose hoeveelheid
Private Const constProdQty As String = "_Glob_MedDisc_ProductDose_"     ' Product hoeveelheid
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
Private Const constMaxKeer As String = "_Glob_MedDisc_MaxKeer_"         ' Maximale dosering per keer
Private Const constPerDose As String = "_Glob_MedDisc_PerDose_"         ' Bereken per dose
Private Const constPerKg As String = "_Glob_MedDisc_PerKg_"             ' Bereken per kg
Private Const constPerM2 As String = "_Glob_MedDisc_PerM2_"             ' Bereken per m2

Private Const constMaxConc As String = "_Glob_MedDisc_MaxConc_"         ' Maximale concentratie
Private Const constOplVlst As String = "_Glob_MedDisc_OplVlst_"         ' Verplichte oplos vloeistof
Private Const constOplVol As String = "_Glob_MedDisc_OplVol_"           ' Minimaal oplos volume
Private Const constMinTijd As String = "_Glob_MedDisc_MinTijd_"         ' Minimale inloop tijd

Private Const constFreqText As String = "Var_MedDisc_FreqText_"

Private Const constVerw As String = "AL"
Private Const constMedCount As Integer = 30

Private m_Formularium As ClassFormularium

Public Function MedDisc_CanonGen(ByVal strGeneriek As String) As String

    strGeneriek = Trim(LCase(strGeneriek))
    strGeneriek = Replace(strGeneriek, "/", "+")
    strGeneriek = Replace(strGeneriek, " ", "-")
    
    MedDisc_CanonGen = strGeneriek

End Function

' Copy paste function cannot be reused because of private clear method
Private Sub ShowPickList(colTbl As Collection, ByVal strRange As String, ByVal intStart As Integer, ByVal intMax As Integer)

    Dim frmPickList As FormMedDiscPickList
    Dim intN As Integer
    Dim strN As String
    Dim strKeuze As String
    Dim strMsg As String
    
    On Error GoTo ShowPickListError
    
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
    
    Exit Sub

ShowPickListError:

    strMsg = "De ingevoerde lijst bevat medicamenten die niet in MetaVision bekend zijn."
    strMsg = strMsg & vbNewLine & "Daarom kan deze functie niet worden gebruikt."
    ModMessage.ShowMsgBoxInfo strMsg
    
End Sub

Public Sub MedDisc_ShowPickList()

    Dim objForm As ClassFormularium
    Dim objGenCol As Collection
    Dim objTable As Range
    Dim varGen As Variant
    
    If m_Formularium Is Nothing Then Set m_Formularium = Formularium_GetFormularium
    
    Set objGenCol = New Collection
    Set objTable = ModRange.GetRange(constTblMedOpdr)
    
    ' Use only generieken from MetaVision
    For Each varGen In objTable
        varGen = Trim(CStr(varGen))
        If Not varGen = vbNullString Then
            varGen = Split(varGen, " ")(0)
            varGen = LCase(varGen)
            If CollectionContains(varGen, m_Formularium.GetGenerieken()) Then
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
    ModRange.SetRangeValue constATCGroup & strN, vbNullString
    
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
    ModRange.SetRangeValue constMaxKeer & strN, 0
    ModRange.SetRangeValue constPerDose & strN, False
    
    ModRange.SetRangeValue constMaxConc & strN, 0
    ModRange.SetRangeValue constOplVlst & strN, vbNullString
    ModRange.SetRangeValue constOplVol & strN, 0
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

Public Sub MedDisc_ClearAll(ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    
    For intN = 1 To 30
        Clear intN
        If blnShowProgress Then ModProgress.SetJobPercentage "Medicatie verwijderen", 30, intN
        
    Next

End Sub

Public Function MedDisc_GetMedicationFreqs() As Dictionary

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
        dblFactor = objTable.Cells(intN, 3).Value2
        dictFreq.Add strFreq, dblFactor
    Next
    
    Set MedDisc_GetMedicationFreqs = dictFreq

End Function

Private Sub Test_MedDisc_GetMedicationFreqs()

    MedDisc_GetMedicationFreqs

End Sub

Private Sub Util_OpenMedDiscForm(ByVal intN As Integer)

    Dim objMed As ClassMedDisc
    Dim objForm As FormMedDisc
    Dim strN As String
    Dim intOplVlst As Integer
    Dim strOplVlst As String
    Dim dblOplVol As Double
    Dim blnLoad As Boolean
    Dim dblKeer As Double
    Dim dblWght As Double
    Dim dblFact As Double
    Dim strFreq As String
    Dim strTime As String
      
    On Error GoTo ErrorHandler
      
    strN = IntNToStrN(intN)
    
    dblWght = ModPatient.Patient_GetWeight()
    
    If dblWght = 0 Then
        ModMessage.ShowMsgBoxExclam "Kan geen medicatie invoeren zonder een gewicht"
        Exit Sub
    End If
    
    Set objForm = New FormMedDisc
    With objForm
    
        .Caption = "Kies een medicament voor regel " & strN
        
        blnLoad = False
        If ModRange.GetRangeValue(constGPK & strN, 0) > 0 Then ' Drug from formularium
            blnLoad = .LoadGPK(CStr(ModRange.GetRangeValue(constGPK & strN, vbNullString)))
        End If
        
        If Not blnLoad Then                      ' Manually entered drug
            .SetNoFormMed
            .cboGeneric.Text = ModRange.GetRangeValue(constGeneric & strN, vbNullString)
            .SetComboBoxIfNotEmpty .cboShape, ModRange.GetRangeValue(constVorm & strN, vbNullString)
            .SetTextBoxIfNotEmpty .txtGenericQuantity, ModRange.GetRangeValue(constConc & strN, 0)
            .SetComboBoxIfNotEmpty .cboGenericQuantityUnit, ModRange.GetRangeValue(constConcUnit & strN, vbNullString)
        End If
        
        ' Edited details
        .SetComboBoxIfNotEmpty .cboRoute, ModRange.GetRangeValue(constRoute & strN, vbNullString)
        .SetComboBoxIfNotEmpty .cboIndication, ModRange.GetRangeValue(constIndic & strN, vbNullString)
        strFreq = ModRange.GetRangeValue(constFreqText & strN, vbNullString)
        
        If Not strFreq = vbNullString Then strTime = ModExcel.Excel_VLookup(strFreq, "Tbl_Glob_MedFreq", 2)
        ' Change of dose time clears all dose details!
        .SetComboBoxIfNotEmpty .cboFreqTime, strTime
        ' Therefore, now set the freq
        .SetComboBoxIfNotEmpty .cboFreq, strFreq
        
        .chkDose.Value = ModRange.GetRangeValue(constHasDose & strN, True)
        .SetTextBoxIfNotEmpty .txtMultipleQuantity, ModRange.GetRangeValue(constStandDose & strN, vbNullString)
        .SetComboBoxIfNotEmpty .cboMultipleQuantityUnit, ModRange.GetRangeValue(constDoseUnit & strN, vbNullString)
        
        dblKeer = shtGlobBerMedDisc.Range("H" & intN + 1).Value2
        .SetTextBoxIfNotEmpty .txtAdminDose, dblKeer
        
        .SetTextBoxIfNotEmpty .txtNormDose, ModRange.GetRangeValue(constNormDose & strN, vbNullString)
        .SetTextBoxIfNotEmpty .txtMinDose, ModRange.GetRangeValue(constMinDose & strN, vbNullString)
        .SetTextBoxIfNotEmpty .txtMaxDose, ModRange.GetRangeValue(constMaxDose & strN, vbNullString)
        .SetTextBoxIfNotEmpty .txtAbsMaxDose, ModRange.GetRangeValue(constAbsDose & strN, vbNullString)
        .SetTextBoxIfNotEmpty .txtMaxPerDose, ModRange.GetRangeValue(constMaxKeer & strN, vbNullString)
        
        .chkPerDose.Value = ModRange.GetRangeValue(constPerDose & strN, False)
        .optKg.Value = ModRange.GetRangeValue(constPerKg & strN, False)
        .optM2.Value = ModRange.GetRangeValue(constPerM2 & strN, False)
        .optNone.Value = (Not .optKg.Value) And (Not .optM2.Value)
        
        intOplVlst = ModRange.GetRangeValue(constSolNo & strN, 1)
        strOplVlst = ModExcel.Excel_Index("Tbl_Glob_MedDisc_OplVl", intOplVlst, 1)
        .SetComboBoxIfNotEmpty .cboSolutions, strOplVlst
        
        .SetTextBoxIfNotEmpty .txtTijd, ModRange.GetRangeValue(constTime & strN, vbNullString)
        
        dblOplVol = ModRange.GetRangeValue(constOplVol & strN, 0)
        dblOplVol = ModString.FixPrecision(dblOplVol, 1)
        If dblOplVol > 0 Then
            .SetToVolume
            .SetTextBoxIfNotEmpty .txtSolutionVolume, dblOplVol
        End If
        
        dblFact = .GetFactorByFreq(.cboFreq)
        If Not dblFact = 0 And Not .txtCalcDose.Value = dblKeer * dblFact / dblWght Then .CaclculateWithKeerDose dblKeer
        .Show
        
        If .GetClickedButton = "OK" Then
            If .HasSelectedMedicament() Then
            
                Set objMed = .GetSelectedMedicament()
                ' -- Medicament --
                MedDisc_SetMed objMed, strN
                
                If .Mail Then
                    ModRange.SetRangeValue "Var_Glob_MedDiscPrtNo", intN
                    SendApotheekMedDisc
                End If
                
            End If

        Else
            If .GetClickedButton = "Clear" Then
                Clear intN
            End If
        End If
        
    End With
    
    Exit Sub
    
ErrorHandler:
        
    ModLog.LogError Err, "Util_OpenMedDiscForm Failed"
        
End Sub

Public Sub MedDisc_SetMed(objMed As ClassMedDisc, strN As String)
    
    Dim intFreq As Integer
    Dim dblFactor As Double
    Dim varFreq As Variant
    Dim dictFreq As Dictionary
    Dim intDoseQty As Integer
    Dim dblOplVol As Double
    Dim dblAdjust As Double
    
    ModRange.SetRangeValue constGPK & strN, objMed.GPK
    ModRange.SetRangeValue constATC & strN, objMed.ATC
    ModRange.SetRangeValue constGeneric & strN, objMed.Generic
    ModRange.SetRangeValue constVorm & strN, objMed.Shape
    ModRange.SetRangeValue constConc & strN, objMed.GenericQuantity
    ModRange.SetRangeValue constConcUnit & strN, objMed.GenericUnit
    ModRange.SetRangeValue constLabel & strN, objMed.Label
    ModRange.SetRangeValue constATCGroup & strN, objMed.MainGroup

    ModRange.SetRangeValue constRoute & strN, objMed.Route
    ModRange.SetRangeValue constIndic & strN, objMed.Indication
    
    ModRange.SetRangeValue constHasDose & strN, objMed.HasDose
    If objMed.HasDose Then
        ModRange.SetRangeValue constDoseSubst & strN, objMed.Substance
        ModRange.SetRangeValue constStandDose & strN, objMed.MultipleQuantity
        ModRange.SetRangeValue constDoseUnit & strN, objMed.MultipleUnit
        ModRange.SetRangeValue constProdQty & strN, objMed.ProductDose
        
        ModRange.SetRangeValue constNormDose & strN, objMed.NormDose
        ModRange.SetRangeValue constMinDose & strN, objMed.MinDose
        ModRange.SetRangeValue constMaxDose & strN, objMed.MaxDose
        ModRange.SetRangeValue constAbsDose & strN, objMed.AbsMaxDose
        ModRange.SetRangeValue constMaxKeer & strN, objMed.MaxPerDose
          
        ModRange.SetRangeValue constMaxConc & strN, objMed.MaxConc
        ModRange.SetRangeValue constOplVlst & strN, objMed.Solution
        ModRange.SetRangeValue constOplVol & strN, objMed.SolutionVolume
        ModRange.SetRangeValue constMinTijd & strN, objMed.MinInfusionTime
        
        
        ModRange.SetRangeValue constPerDose & strN, objMed.PerDose
        ModRange.SetRangeValue constPerKg & strN, objMed.PerKg
        ModRange.SetRangeValue constPerM2 & strN, objMed.PerM2
        
        If objMed.Solution = "NaCl 0,9%" Then
            ModRange.SetRangeValue constSolNo & strN, 2
        ElseIf objMed.Solution = "glucose 5%" Then
            ModRange.SetRangeValue constSolNo & strN, 3
        ElseIf objMed.Solution = "glucose 10%" Then
            ModRange.SetRangeValue constSolNo & strN, 4
        End If
        
        ModRange.SetRangeValue constTime & strN, objMed.MinInfusionTime
        
        If Not objMed.Freq = vbNullString Then
            Set dictFreq = MedDisc_GetMedicationFreqs()
            intFreq = 2
            For Each varFreq In dictFreq
                If varFreq = objMed.Freq Then Exit For
                intFreq = intFreq + 1
            Next
            ModRange.SetRangeValue constFreq & strN, intFreq
        End If
        
        ModRange.SetRangeValue constFreqList & strN, objMed.GetFreqListString
        
        If Not objMed.MultipleQuantity = 0 And Not intFreq < 2 Then
            dblAdjust = 1
            dblAdjust = IIf(objMed.PerKg, ModPatient.Patient_GetWeight(), dblAdjust)
            dblAdjust = IIf(objMed.PerM2, ModPatient.CalculateBSA(), dblAdjust)
            
            dblFactor = IIf(objMed.PerDose, 1, ModExcel.Excel_Index(constFreqTable, intFreq, 3))
            ' intDoseQty = objMed.CalcDose * dblAdjust / dblFactor / objMed.MultipleQuantity
            intDoseQty = objMed.AdminDose / objMed.MultipleQuantity
            ModRange.SetRangeValue constDoseQty & strN, intDoseQty
            
        End If
        
        If Not objMed.DoseText = vbNullString Then
            ModRange.SetRangeValue constText & strN, objMed.DoseText
        End If
    End If
    
End Sub

Public Sub MedDisc_EnterMed_01()

    Util_OpenMedDiscForm 1

End Sub

Public Sub MedDisc_EnterMed_02()

    Util_OpenMedDiscForm 2

End Sub

Public Sub MedDisc_EnterMed_03()

    Util_OpenMedDiscForm 3

End Sub

Public Sub MedDisc_EnterMed_04()

    Util_OpenMedDiscForm 4

End Sub

Public Sub MedDisc_EnterMed_05()

    Util_OpenMedDiscForm 5

End Sub

Public Sub MedDisc_EnterMed_06()

    Util_OpenMedDiscForm 6

End Sub

Public Sub MedDisc_EnterMed_07()

    Util_OpenMedDiscForm 7

End Sub

Public Sub MedDisc_EnterMed_08()

    Util_OpenMedDiscForm 8

End Sub

Public Sub MedDisc_EnterMed_09()

    Util_OpenMedDiscForm 9

End Sub

Public Sub MedDisc_EnterMed_10()

    Util_OpenMedDiscForm 10

End Sub

Public Sub MedDisc_EnterMed_11()

    Util_OpenMedDiscForm 11

End Sub

Public Sub MedDisc_EnterMed_12()

    Util_OpenMedDiscForm 12

End Sub

Public Sub MedDisc_EnterMed_13()

    Util_OpenMedDiscForm 13

End Sub

Public Sub MedDisc_EnterMed_14()

    Util_OpenMedDiscForm 14

End Sub

Public Sub MedDisc_EnterMed_15()

    Util_OpenMedDiscForm 15

End Sub

Public Sub MedDisc_EnterMed_16()

    Util_OpenMedDiscForm 16

End Sub

Public Sub MedDisc_EnterMed_17()

    Util_OpenMedDiscForm 17

End Sub

Public Sub MedDisc_EnterMed_18()

    Util_OpenMedDiscForm 18

End Sub

Public Sub MedDisc_EnterMed_19()

    Util_OpenMedDiscForm 19

End Sub

Public Sub MedDisc_EnterMed_20()

    Util_OpenMedDiscForm 20

End Sub

Public Sub MedDisc_EnterMed_21()

    Util_OpenMedDiscForm 21

End Sub

Public Sub MedDisc_EnterMed_22()

    Util_OpenMedDiscForm 22

End Sub

Public Sub MedDisc_EnterMed_23()

    Util_OpenMedDiscForm 23

End Sub

Public Sub MedDisc_EnterMed_24()

    Util_OpenMedDiscForm 24

End Sub

Public Sub MedDisc_EnterMed_25()

    Util_OpenMedDiscForm 25

End Sub

Public Sub MedDisc_EnterMed_26()

    Util_OpenMedDiscForm 26

End Sub

Public Sub MedDisc_EnterMed_27()

    Util_OpenMedDiscForm 27

End Sub

Public Sub MedDisc_EnterMed_28()

    Util_OpenMedDiscForm 28

End Sub

Public Sub MedDisc_EnterMed_29()

    Util_OpenMedDiscForm 29

End Sub

Public Sub MedDisc_EnterMed_30()

    Util_OpenMedDiscForm 30

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

Public Function MedDisc_GetOplVlstCol() As Collection

    Dim objCol As Collection
    Dim objTable As Range
    Dim varItem As Variant
    
    Set objTable = ModRange.GetRange("Tbl_Glob_MedDisc_OplVl")
    Set objCol = New Collection
    
    For Each varItem In objTable
        objCol.Add varItem
    Next

    Set MedDisc_GetOplVlstCol = objCol
    
End Function

Public Sub SendApotheekMedDisc()

    Dim blnInValid As Boolean
    Dim intNo As Integer
    Dim intMed As Integer
    Dim strNo As String
    Dim blnAsk As Boolean
    Dim blnPrint As Boolean
    Dim strUser As String
    Dim vbAnswer As Integer
    
    Dim objMsg As Object
    Dim strTo As String
    Dim strCc As String
    Dim strFrom As String
    Dim strSubject As String
    Dim strHTML As String
    
    Dim strFile As String
    Dim strPDF As String
    
    Dim strMail As String
    
    On Error GoTo ErrorHandler
       
    blnInValid = ModRange.GetRangeValue("Var_Glob_ValidMedDiscPrescr", True)
    If blnInValid Then
        ModProgress.FinishProgress
        ModMessage.ShowMsgBoxExclam "Medicatie voorshrift niet compleet." & "Kan de apotheek recept niet verzenden!"
        Exit Sub
    End If
                   
    strMail = "wkz-algemeen@umcutrecht.nl"
    If Not ModSetting.IsProductionDir() Then strMail = ModMessage.ShowInputBox("Voer een email adres in", vbNullString)
    
    If strMail = vbNullString Then
        ModMessage.ShowMsgBoxExclam "Er moet een email adres worden ingevoerd." & vbNewLine & "Kan de apotheekbrief niet verzenden!"
        Exit Sub
    End If
        
    ModProgress.StartProgress "Medicatie naar de apotheek verzenden"

    strTo = strMail
'     strCc = "vbassneo@umcutrecht.nl"
    strFrom = "FunctioneelBeheerMetavision@umcutrecht.nl"
    strSubject = "Medicatie recept voor " & ModPatient.Patient_GetHospitalNumber & " " & ModPatient.Patient_GetLastName & ", " & ModPatient.Patient_GetFirstName
    strHTML = vbNullString
    
    Set objMsg = CreateObject("CDO.Message")
    With objMsg
         
        .To = CStr(strTo)
'         .Cc = CStr(strCc)
        .From = CStr(strFrom)
        .Subject = CStr(strSubject)
        .HTMLBody = CStr(strHTML)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPickup=1, cdoSendUsingPort=2, cdoSendUsingExchange=3
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.umcutrecht.nl"
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Configuration.Fields.Update
        
        strFile = Environ("TEMP") & "\MedDisc_" & ModPatient.Patient_GetHospitalNumber
        strPDF = PrintMedDiscPrev(False, strFile)
        .AddAttachment strPDF
                
        .Send
    
    End With
        
    Set objMsg = Nothing
    
    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxInfo "Medicatie recept is verstuurd naar de apotheek"
    
    Exit Sub

ErrorHandler:

    ModLog.LogError Err, "SendApotheekMedDisc"
    
    On Error Resume Next
    
    ModMessage.ShowMsgBoxError "Medicatie recept is niet verstuurd naar de apotheek, een foutmelding is verzonden naar functioneel beheer"
        
    Set objMsg = Nothing
    
    ModProgress.FinishProgress

End Sub

Private Function PrintMedDiscPrev(ByVal blnPrev As Boolean, ByVal strFile As String) As String

    Dim strPDF As String
    
    If strFile = vbNullString Then
        PrintSheet shtGlobBerMedDiscMail, 1, False, blnPrev
    Else
        strPDF = strFile & ".pdf"
        SaveSheetAsPDF shtGlobBerMedDiscMail, strPDF, True
    End If
    
    PrintMedDiscPrev = strPDF
    
End Function

Public Sub MedDisc_SortTableMedDisc()
    
    Dim strColumn As String
    Dim strRange As String
    
    strColumn = "BW2:BW31"
    strRange = "Tbl_Glob_SortMedDisc"
    
    ImprovePerf True ' Prevent cycling through all windows when sheets are processed
    
    shtGlobBerMedDisc.Sort.SortFields.Clear
    shtGlobBerMedDisc.Sort.SortFields.Add Key:=Range( _
        strColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With shtGlobBerMedDisc.Sort
        .SetRange Range(strRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ImprovePerf False
    
End Sub


Public Sub SendApotheekMedDiscValidation()

    Dim blnAsk As Boolean
    Dim blnPrint As Boolean
    Dim strUser As String
    Dim vbAnswer As Integer
    
    Dim objMsg As Object
    Dim strTo As String
    Dim strCc As String
    Dim strFrom As String
    Dim strSubject As String
    Dim strHTML As String
    
    Dim strFile As String
    Dim strPDF As String
    
    Dim strMail As String
    
    On Error GoTo ErrorHandler
                          
    strMail = "wkz-algemeen@umcutrecht.nl"
    If Not ModSetting.IsProductionDir() Then strMail = ModMessage.ShowInputBox("Voer een email adres in", vbNullString)
    
    If strMail = vbNullString Then
        ModMessage.ShowMsgBoxExclam "Er moet een email adres worden ingevoerd." & vbNewLine & "Kan de apotheekbrief niet verzenden!"
        Exit Sub
    End If
        
    ModProgress.StartProgress "Discontinue medicatie voor validatie naar de apotheek verzenden"

    strTo = strMail
'     strCc = "vbassneo@umcutrecht.nl"
    strFrom = "FunctioneelBeheerMetavision@umcutrecht.nl"
    strSubject = "Medicatie validatie voor " & ModPatient.Patient_GetHospitalNumber & " " & ModPatient.Patient_GetLastName & ", " & ModPatient.Patient_GetFirstName
    strHTML = vbNullString
    
    Set objMsg = CreateObject("CDO.Message")
    With objMsg
         
        .To = CStr(strTo)
'         .Cc = CStr(strCc)
        .From = CStr(strFrom)
        .Subject = CStr(strSubject)
        .HTMLBody = CStr(strHTML)
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPickup=1, cdoSendUsingPort=2, cdoSendUsingExchange=3
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.umcutrecht.nl"
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Configuration.Fields.Update
        
        strFile = Environ("TEMP") & "\MedDiscValidation_" & ModPatient.Patient_GetHospitalNumber
        strPDF = PrintMedDiscValidationPrev(False, strFile)
        .AddAttachment strPDF
                
        .Send
    
    End With
        
    Set objMsg = Nothing
    
    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxInfo "Medicatie validatie is verstuurd naar de apotheek"
    
    Exit Sub

ErrorHandler:

    ModLog.LogError Err, "SendApotheekMedDiscValidation"
    
    On Error Resume Next
    
    ModMessage.ShowMsgBoxError "Medicatie validatie is niet verstuurd naar de apotheek, een foutmelding is verzonden naar functioneel beheer"
        
    Set objMsg = Nothing
    
    ModProgress.FinishProgress

End Sub

Private Function PrintMedDiscValidationPrev(ByVal blnPrev As Boolean, ByVal strFile As String) As String

    Dim strPDF As String
    
    If strFile = vbNullString Then
        PrintSheet shtGlobPrtMedDisc, 1, False, blnPrev
    Else
        strPDF = strFile & ".pdf"
        SaveSheetAsPDF shtGlobPrtMedDisc, strPDF, True
    End If
    
    PrintMedDiscValidationPrev = strPDF
    
End Function

Public Sub MedDisc_ImportFromHix()

    Dim intAnswer As VbMsgBoxResult
    Dim strMsg As String
    
    strMsg = "Om medicatie uit HIX te importeren moeten de volgende stappen worden genomen:"
    strMsg = strMsg & vbNewLine & "1. Selecteer de medicatielijst van de betreffende patient in HIX"
    strMsg = strMsg & vbNewLine & "2. Zorg dat alle actieve medicatie regels zijn aangevinkt (links boven in de medicatielijst)"
    strMsg = strMsg & vbNewLine & "3. Rechtsklik in de medicatie lijst en kies de optie kopieren"
    strMsg = strMsg & vbNewLine & ""
    strMsg = strMsg & vbNewLine & "Als je dit gedaan hebt en je de medicatie wilt importeren klik dan op Ja"
    
    intAnswer = ModMessage.ShowMsgBoxYesNo(strMsg)
    
    If intAnswer = vbYes Then
        Application.DisplayAlerts = False
        
        shtGlobTemp.UsedRange.ClearContents
        shtGlobTemp.Range("A1").PasteSpecial xlPasteAll
        Application.DisplayAlerts = True
    
        ImportCopiedMedications
        shtGlobTemp.Cells.ClearContents
    End If
    

End Sub

Private Sub ImportCopiedMedications()

    Dim objForm As ClassFormularium
    Dim objRange As Range
    Dim objMed As ClassMedDisc
    Dim colMed As Collection
    Dim intN As Integer
    Dim intC As Integer
    Dim intJ As Integer
    Dim intK As Integer
    Dim intM As Integer
    Dim strK As String
    Dim strLabel As String
    Dim arrLabel() As String
    Dim strMissed As String
    Dim blnSet As Boolean
    Dim blnEqs As Boolean
    Dim strMsg As String
    Dim strGen As String
    Dim strProd As String
    Dim strTooMany As String
    
    Set objForm = ModFormularium.Formularium_GetNewFormularium()
    Set colMed = objForm.GetMedicationCollection(False)
    Set objRange = shtGlobTemp.Range("A1").CurrentRegion()
        
    ModProgress.StartProgress "Importeren van medicatie uit HIX"
        
    ImprovePerf True
    MedDisc_ClearAll True
    ImprovePerf False
        
    intC = objRange.Rows.Count + 10
    For intN = 2 To intC
       strLabel = Trim(objRange.Cells(intN, 3).Value)
       If Not (strLabel = vbNullString Or strLabel = "Geaccordeerd") Then
            intM = intM + 1
            blnSet = False
            For Each objMed In colMed
                 arrLabel = Split(objMed.Label, " ")
                 blnEqs = True
                 
                 strGen = arrLabel(0) & " "
                 strProd = Split(objMed.Product, " ")(0) & " "
                 
                 blnEqs = blnEqs And ModString.ContainsInclSpace(strLabel, strGen)
                 If Not blnEqs Then
                    blnEqs = ModString.ContainsInclSpace(strLabel, strProd)
                    arrLabel = Split(objMed.Product, " ")
                 End If
                 
                 If blnEqs Then
                    For intJ = 1 To UBound(arrLabel)
                       blnEqs = blnEqs And (ModString.ContainsCaseSensitive(strLabel, arrLabel(intJ)))
                    Next
                 End If
                 
                 'Make sure that no more than 30 MO's are added
                 blnEqs = blnEqs And (intK < 30)
                 
                 If blnEqs Then
                    intK = intK + 1
                    strK = ModString.IntNToStrN(intK)
                    objMed.Indication = vbNullString
                    
                    ModRange.SetRangeValue constText & strK, objRange.Cells(intN, 4).Value & " " & objRange.Cells(intN, 5).Value
                    MedDisc_SetMed objMed, strK
                    
                    blnSet = True
                    Exit For
                 End If
            Next
            
            If intK > 30 Then
                strTooMany = strTooMany & vbNewLine & strLabel
            Else
                If Not blnSet Then
                    strMissed = strMissed & vbNewLine & strLabel
                    
                    intK = intK + 1
                    strK = ModString.IntNToStrN(intK)
                    ModRange.SetRangeValue constText & strK, strLabel & " " & objRange.Cells(intN, 4).Value & " " & objRange.Cells(intN, 5).Value
                End If
            End If
       
       End If
       
       ModProgress.SetJobPercentage "Medicatie import", intC, intN
    Next
    
    MedDisc_SortTableMedDisc
    
    If strMissed = vbNullString And strTooMany = vbNullString Then
        strMsg = "Alle (" & intM & ") medicamenten werden geimporteerd"
    Else
        If Not strMissed = vbNullString Then
            strMsg = "De volgende medicamenten konden niet worden geimporteerd: " & vbNewLine & strMissed
        End If
        If Not strTooMany = vbNullString Then
            strMsg = IIf(strMsg = vbNullString, strMsg, strMsg & vbNewLine & vbNewLine)
            strMsg = strMsg & "De volgende medicamenten waren te veel opdrachten: " & vbNewLine & strTooMany
        End If
    End If
    
    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxInfo strMsg
    
    Set colMed = Nothing
    Set objMed = Nothing
    Set objForm = Nothing
    Set objRange = Nothing

End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedDisc 
   Caption         =   "Kies een medicament ..."
   ClientHeight    =   14940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   OleObjectBlob   =   "FormMedDisc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMedDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Med As ClassMedDisc
Private m_TherapieGroep As String
Private m_SubGroep As String
Private m_Etiket As String
Private m_Product As String

Private m_IsGPK As Boolean
Private m_LoadGPK As Boolean

Private m_Freq As Dictionary
Private m_Keer As Boolean
Private m_Conc As Boolean
Private m_Mail As Boolean
Private m_ValidMed As Boolean
Private m_CalcVol As Boolean
Private m_CalcDose As Boolean
Private m_HandlingEvent As Boolean

Private Enum Events
    GenericChange = 1
    MedicationChange = 2
    DoseChange = 3
    SolutionChange = 4
    DoseRulesChange = 5
    HasDoseChange = 6
End Enum

Private Sub HandleEvent(ByVal enmEvent As Events, _
                        Optional ByVal strValid As String = vbNullString, _
                        Optional ByVal blnMsg As Boolean = False)

    If m_HandlingEvent Then Exit Sub
    m_HandlingEvent = True
    
    Debug.Print "Handling event: " & enmEvent

    Select Case enmEvent
    
        Case GenericChange
            ApplyGenericChange
            SetDoseUnit
            Validate vbNullString, False
            If Not m_ValidMed Then
                ClearDose
            Else
                ClearDose
                ApplyDoseRule
                LoadMedicament False, True
            End If
    
        Case MedicationChange
            Validate strValid, blnMsg
            If Not m_ValidMed Then
                ClearDose
            Else
                ClearDose
                ApplyDoseRule
                LoadMedicament False, True
            End If
            
        Case DoseChange
            If frmDose.Visible Then
                If strValid = vbNullString Then
                    SetDoseUnit
                    CalculateDose
                    CalculateVolume
                    CalculateLabel
                    SetProductDose
                    Validate vbNullString, True
                Else
                    Validate strValid, False
                End If
            End If
            
        Case HasDoseChange
            If Not frmDose.Visible Then ClearDose
            
            frmDose.Visible = chkDose.Value
            frmSolution.Visible = chkDose.Value
            txtMultipleQuantity.Visible = chkDose.Value
            cboMultipleQuantityUnit.Visible = chkDose.Value
        
            Validate vbNullString, False
        
        Case DoseRulesChange
            If frmDose.Visible And m_ValidMed Then
                LoadMedicament False, False
                SelectDoseRule
                CalculateDose
                CalculateVolume
            End If
            
        Case SolutionChange
            If frmDose.Visible And m_ValidMed Then CalculateVolume
    
    End Select
    
    m_HandlingEvent = False

End Sub

Public Property Get Mail() As Boolean

    Mail = m_Mail

End Property

Public Sub SetToVolume()

    cmdVol_Click

End Sub

Public Sub SetNoFormMed()

    m_IsGPK = False

End Sub

Private Function IsAbsMaxInvalid() As Boolean
    
    IsAbsMaxInvalid = txtAbsMaxDose.Value = vbNullString And ModPatient.Patient_GetWeight() > 50 And txtNormDose.Value = vbNullString And txtMaxDose.Value = vbNullString

End Function

Private Function IsDoseControlInValid() As Boolean

    IsDoseControlInValid = txtNormDose.Value = vbNullString And txtMaxDose.Value = vbNullString

End Function

Private Sub Validate(ByVal strValid As String, ByVal blnMsg As Boolean)
    
    If strValid = vbNullString Then
    
        If frmDose.Visible Then
            strValid = IIf(cboMultipleQuantityUnit.Value = vbNullString, "Voer dosering eenheid in", strValid)
            strValid = IIf(txtMultipleQuantity.Value = vbNullString, "Voer een deelbaarheid in", strValid)
        End If
        
        strValid = IIf(cboIndication.Value = vbNullString, "Kies een indicatie", strValid)
        strValid = IIf(cboRoute.Value = vbNullString, "Kies een route", strValid)
        
        strValid = IIf(cboGenericQuantityUnit.Value = vbNullString, "Voer sterkte eenheid in", strValid)
        strValid = IIf(txtGenericQuantity.Value = vbNullString, "Voer sterkte in", strValid)
        
        strValid = IIf(cboShape.Value = vbNullString, "Voer een vorm in", strValid)
        strValid = IIf(cboGeneric.Value = vbNullString, "Kies een generiek", strValid)
                
        m_ValidMed = strValid = vbNullString
        
    End If
    
    cmdOK.Enabled = strValid = vbNullString
    cmdQuery.Visible = strValid = vbNullString
    
    If frmDose.Visible And strValid = vbNullString Then
        strValid = IIf(IsDoseControlInValid, "Voer of een advies dosering in en/of een max (en evt. min en abs max) dosering", strValid)
        strValid = IIf(IsAbsMaxInvalid, "Gewicht boven de 50 kg, voer een absolute maximum dosering in (of een advies dosering of max dosering)", strValid)
    
    End If
    
    lblValid.Caption = strValid

End Sub

Public Function HasSelectedMedicament() As Boolean

    HasSelectedMedicament = Not m_Med Is Nothing

End Function

Public Function GetSelectedMedicament() As ClassMedDisc

    Set GetSelectedMedicament = m_Med

End Function

Public Function GetClickedButton() As String

    GetClickedButton = lblButton.Caption

End Function

Private Sub SetToGPKMode(ByVal blnIsGPK As Boolean)

    Dim varItem As Variant
    
    m_IsGPK = blnIsGPK
    
    Me.txtGenericQuantity.Enabled = Not blnIsGPK
    Me.cboGenericQuantityUnit.Enabled = Not blnIsGPK
    Me.cboShape.Enabled = Not blnIsGPK
    
    cmdFormularium.Enabled = blnIsGPK
    
    If Not blnIsGPK Then
        lblGPK.Caption = vbNullString
        lblATC.Caption = vbNullString
        
        FillCombo cboShape, Formularium_GetFormularium().GetShapeCollection()
        FillCombo cboGenericQuantityUnit, Formularium_GetFormularium().GetGenericQuantityUnitCollection()
        FillCombo cboMultipleQuantityUnit, Formularium_GetFormularium.GetDoseUnitCollection()
        FillCombo cboRoute, Formularium_GetFormularium.GetRoutes()
        
        cboIndication.Clear
        
        LoadFreq
        
        FillCombo cboSolutions, MedDisc_GetOplVlstCol()
    End If

End Sub

Private Sub cboMultipleQuantityUnit_Change()

    ' SetDoseUnit
    HandleEvent DoseChange

End Sub

Private Sub cboFreq_Change()

    Dim strValid As String
    
    strValid = ValidateCombo(cboFreq)
    
    HandleEvent DoseChange, strValid
        
End Sub

Private Sub cboFreqTime_Change()

    HandleEvent DoseRulesChange

End Sub

Private Sub SelectDoseRule()

    Dim objRule As ClassDoseRule
    Dim strTime As String
    Dim strFreq As String
    
    
    If m_Med.DoseRules.Count > 0 Then
    
        ClearDose
        For Each objRule In m_Med.DoseRules
            strFreq = objRule.Freq
            If strFreq = "antenoctum" Then
                strTime = "dag"
            Else
                strFreq = Replace(strFreq, "antenoctum||", "")
            End If
            
            strTime = Replace(strTime, "antenoctum||", "")
            strTime = Trim(Split(Split(strFreq, "||")(0), "/")(1))
            
            If cboFreqTime.Value = vbNullString Then cboFreqTime.Value = strTime
            If cboFreqTime.Value = strTime And cboSubstance.Value = objRule.Substance Then
                With m_Med
                    .PerKg = objRule.PerKg
                    .PerM2 = objRule.PerM2
                    
                    .SetFreqList objRule.Freq
                    
                    .NormDose = objRule.NormDose
                    .MinDose = objRule.MinDose
                    .MaxDose = objRule.MaxDose
                    .MaxPerDose = objRule.MaxPerDose
                    .AbsMaxDose = objRule.AbsMaxDose
                    
                    If Not .GetFreqListString = vbNullString Then
                        FillCombo cboFreq, .GetFreqList()
                    Else
                        LoadFreq
                    End If
                    
                    If Not .Freq = vbNullString Then cboFreq.Value = .Freq
                    
                    optNone = True
                    optKg = .PerKg
                    optM2 = .PerM2
                    
                    chkPerDose = .PerDose
                    
                    SetTextBoxNumericValue txtNormDose, .NormDose
                    SetTextBoxNumericValue txtMinDose, .MinDose
                    SetTextBoxNumericValue txtMaxDose, .MaxDose
                    SetTextBoxNumericValue txtAbsMaxDose, .AbsMaxDose
                    SetTextBoxNumericValue txtMaxPerDose, .MaxPerDose
                End With
                
            End If
        Next
    End If
    
End Sub


Private Sub cboGeneric_Click()

    cboGeneric_Change

End Sub

Private Sub cboGeneric_Change()

    HandleEvent GenericChange

End Sub

Private Sub ApplyGenericChange()

    Dim objMed As ClassMedDisc
    
    If m_LoadGPK Then Exit Sub

    If cboGeneric.ListIndex > -1 Then
        SetToGPKMode True
        Set m_Med = Formularium_GetFormularium.GetMedication(cboGeneric.ListIndex + 1)
        LoadMedicament True, False
    Else
        SetToGPKMode False
        ClearForm False
    End If
    

End Sub

Public Function GetGPK() As String
    Dim strGPK As String
    
    strGPK = "0"
    If Not m_Med Is Nothing Then strGPK = m_Med.GPK
    
    GetGPK = strGPK

End Function

Public Function LoadGPK(ByVal strGPK As String) As Boolean
    
    Dim blnLoad As Boolean

    blnLoad = True
    
    Set m_Med = Formularium_GetFormularium.GPK(strGPK)
    
    If m_Med Is Nothing Then
        SetToGPKMode False
        blnLoad = False
    Else
        SetToGPKMode True
        LoadMedicament True, False
        m_LoadGPK = True
        cboGeneric.Text = m_Med.Generic
        m_LoadGPK = False
    End If
    
    LoadGPK = blnLoad

End Function

Private Function GetAdjust() As String
    
    Dim strAdjust As String
    
    strAdjust = vbNullString
    strAdjust = IIf(optKg.Value, "kg", strAdjust)
    strAdjust = IIf(optM2.Value, "m2", strAdjust)

    GetAdjust = strAdjust

End Function

Private Sub SetDoseUnit()

    Dim strAdjust As String
    Dim strUnit As String
    Dim strTime As String
    
    strAdjust = GetAdjust()
    strUnit = Trim(cboMultipleQuantityUnit.Text)
    strTime = IIf(chkPerDose.Value, "dosis", GetTimeByFreq())
    
    If Not strUnit = vbNullString Then
        If Not strTime = vbNullString Then
            cboDoseUnit.Text = strUnit & (IIf(strAdjust = "", "/", "/" & strAdjust & "/")) & strTime
        Else
            cboDoseUnit.Text = strUnit & (IIf(strAdjust = "", "", "/" & strAdjust))
        End If
        
        cboAdminDoseUnit.Text = strUnit
        lblDoseUnit.Caption = cboDoseUnit.Value
        lblMinEenheid.Caption = cboDoseUnit.Value
        lblMaxEenheid.Caption = cboDoseUnit.Value
        
        If Not strTime = vbNullString Then
            lblAbsMaxEenheid.Caption = strUnit & "/" & strTime
        Else
            lblAbsMaxEenheid.Caption = strUnit
        End If
        
        lblMaxKeerUnit.Caption = strUnit
        lblConcUnit.Caption = strUnit & "/ml"
    Else
        lblDoseUnit.Caption = ""
        lblMinEenheid.Caption = ""
        lblMaxEenheid.Caption = ""
        lblAbsMaxEenheid.Caption = ""
        lblMaxKeerUnit.Caption = ""
        lblConcUnit.Caption = ""
    End If

End Sub

Private Sub LoadSolution()

    With m_Med
        SetTextBoxNumericValue txtMaxConc, .MaxConc
        SetTextBoxNumericValue txtSolutionVolume, .SolutionVolume
        If (.SolutionVolume > 0 And m_Conc) Then
            ToggleConc
        ElseIf .MaxConc > 0 And .SolutionVolume = 0 And Not m_Conc Then
            ToggleConc
        End If
        
        SetTextBoxNumericValue txtTijd, .MinInfusionTime
        cboSolutions.Value = .Solution
        
    End With
    
End Sub

Private Sub LoadMedicament(ByVal blnReload As Boolean, ByVal blnApplyDose As Boolean)
    
    Dim intN As Integer
    Dim strFreq As String
    Dim arrFreq() As String
    Dim strTime As String
    Dim objRule As ClassDoseRule

    With m_Med
    
        lblMainGroup.Caption = .MainGroup
        lblSubGroup.Caption = .SubGroup
        lblEtiket.Caption = .Label
        lblProduct.Caption = .Product
        
        lblGPK.Caption = .GPK
        lblATC.Caption = .ATC
        
        cboShape.Value = .Shape
        
        txtGenericQuantity.Text = .GenericQuantity
        cboGenericQuantityUnit.Text = .GenericUnit
        
        If blnReload Or cboRoute.Text = "" Then FillCombo cboRoute, .GetRouteList()
        If blnReload Or cboIndication.Text = "" Then FillCombo cboIndication, .GetIndicationList()
        
        If .MultipleQuantity = 0 Then
            chkDose.Value = False
            chkDose_Click
            
            Exit Sub
        End If
        
        FillCombo cboSubstConc, GetSubstances()
        
        txtMultipleQuantity.Text = .MultipleQuantity
        
        FillCombo cboMultipleQuantityUnit, GetDoseUnitCollection()
        cboMultipleQuantityUnit.Text = .MultipleUnit
        
        If Not .GetFreqListString = vbNullString Then
            FillCombo cboFreq, .GetFreqList()
        Else
            LoadFreq
        End If
        
        If Not .Freq = vbNullString Then cboFreq.Value = .Freq
        If Not .Substance = vbNullString Then cboSubstance.Value = .Substance
        
        optNone = True
        optKg = .PerKg
        optM2 = .PerM2
        
        chkPerDose = .PerDose
        
        SetTextBoxNumericValue txtNormDose, .NormDose
        SetTextBoxNumericValue txtMinDose, .MinDose
        SetTextBoxNumericValue txtMaxDose, .MaxDose
        SetTextBoxNumericValue txtAbsMaxDose, .AbsMaxDose
        SetTextBoxNumericValue txtMaxPerDose, .MaxPerDose
        
        LoadSolution
        
        If Not blnApplyDose And m_Med.DoseRules.Count > 1 Then
            lblFreqTime.Visible = True
            cboFreqTime.Visible = True
            
            For Each objRule In m_Med.DoseRules
                ' Remove antenoctum to enable string to time parsing
                strFreq = objRule.Freq
                strFreq = IIf(strFreq = "antenoctum", "1 x /dag", strFreq)
                strFreq = Replace(strFreq, "antenoctum||", "")
                
                strTime = Trim(Split(Split(strFreq, "||")(0), "/")(1))
                If Not ComboContainsStringValue(cboFreqTime, strTime) Then cboFreqTime.AddItem strTime
                If Not ComboContainsStringValue(cboSubstance, objRule.Substance) Then cboSubstance.AddItem objRule.Substance
            Next
            
            SetComboIfCountOne cboFreqTime
            SetComboIfCountOne cboSubstance
        Else
            lblFreqTime.Visible = False
            cboFreqTime.Visible = False
        End If
        
    End With

End Sub

Private Function ComboContainsStringValue(objCombo As MSForms.ComboBox, strVal As String) As Boolean

    Dim varVal As Variant
    
    If objCombo.ListCount = 0 Then
        ComboContainsStringValue = False
        Exit Function
    End If
    
    For Each varVal In objCombo.List
        If varVal = strVal Then
            ComboContainsStringValue = True
            Exit Function
        End If
    Next
    
    ComboContainsStringValue = False

End Function

Public Sub SetComboBoxIfNotEmpty(cboBox As MSForms.ComboBox, ByVal strValue As String)

    If Not ModString.StringIsZeroOrEmpty(strValue) Then cboBox.Value = strValue

End Sub

Public Sub SetTextBoxIfNotEmpty(txtBox As MSForms.TextBox, ByVal strValue As String)

    If Not ModString.StringIsZeroOrEmpty(strValue) Then txtBox.Value = strValue

End Sub

Public Sub SetTextBoxNumericValue(txtBox As MSForms.TextBox, ByVal strValue As String)
    
    If Not IsNumeric(strValue) Then strValue = vbNullString
    strValue = ModString.StringToDouble(strValue)
    txtBox.Value = IIf(strValue = "0", vbNullString, strValue)
    
End Sub

Private Sub TextBoxStringNumericValue(txtBox As MSForms.TextBox)

    SetTextBoxNumericValue txtBox, txtBox.Value

End Sub

Private Function GetDoseUnitCollection() As Collection

    Set GetDoseUnitCollection = Formularium_GetFormularium.GetDoseUnitCollection()

End Function

Private Function GetSubstances() As Collection

    Dim colSubst As Collection
    Dim objSubst As ClassSubstance
    
    Set colSubst = New Collection
    For Each objSubst In m_Med.Substances
        colSubst.Add objSubst.Substance
    Next
    
    Set GetSubstances = colSubst

End Function

Private Sub FillCombo(objCombo As ComboBox, colItems As Collection)

    Dim varItem As Variant

    objCombo.Text = vbNullString
    objCombo.Clear
    
    For Each varItem In colItems
        varItem = CStr(varItem)
        If Not varItem = vbNullString Then objCombo.AddItem CStr(varItem)
    Next varItem
    
    SetComboIfCountOne objCombo
    
End Sub

Private Sub SetComboIfCountOne(objCombo As ComboBox)
    
    If objCombo.ListCount = 1 Then
        objCombo.Text = objCombo.List(0)
    End If
    

End Sub

Public Sub ClearForm(ByVal blnClearGeneric As Boolean)

    lblMainGroup.Caption = m_TherapieGroep
    lblSubGroup.Caption = m_SubGroep
    lblEtiket.Caption = m_Etiket
    lblProduct.Caption = m_Product
    
    If blnClearGeneric Then cboGeneric.Value = vbNullString
    cboShape.Value = vbNullString
    
    txtMultipleQuantity.Value = vbNullString
    cboMultipleQuantityUnit.Clear
    cboMultipleQuantityUnit.Value = vbNullString
    
    txtGenericQuantity.Text = vbNullString
    cboGenericQuantityUnit.Text = vbNullString
    
    FillCombo cboRoute, Formularium_GetFormularium.GetRoutes()
    
    cboIndication.Clear
    cboIndication.Value = vbNullString
    
    cboFreq.Value = vbNullString
    txtNormDose.Value = vbNullString
    txtMinDose.Value = vbNullString
    txtMaxDose.Value = vbNullString
    txtAbsMaxDose.Value = vbNullString
    txtMaxPerDose.Value = vbNullString
    txtCalcDose.Value = vbNullString
    cboDoseUnit.Value = vbNullString
    lblCalcDose.Caption = vbNullString
    lblDoseUnit.Caption = vbNullString
    
    cboGeneric.SetFocus
    
    Set m_Med = Nothing

End Sub

Private Function ValidateCombo(objCombo As MSForms.ComboBox) As String

    Dim strValid As String
    
    strValid = vbNullString
    If objCombo.ListCount > 0 And Not objCombo.MatchFound Then
        objCombo.Value = vbNullString
        strValid = "Kies een item uit de lijst"
    End If
    
    ValidateCombo = strValid

End Function

Private Sub cboIndication_Change()

    If Not m_Med Is Nothing Then m_Med.Indication = Trim(LCase(cboIndication.Text))
    
    HandleEvent MedicationChange

End Sub

Private Sub ApplySolution()

    Dim objSol As ClassSolution
    Dim dblQty As Double
    Dim strDep As String
    
    ClearSolution
    
    If Not m_CalcVol Or m_Med Is Nothing Then Exit Sub
    
    strDep = IIf(MetaVision_IsPICU, "PICU", "NICU")
    
    dblQty = StringToDouble(txtAdminDose.Text)
    Set objSol = m_Med.GetSolution(strDep, dblQty)
    
    If Not objSol Is Nothing Then
        With objSol
            m_Med.Solution = objSol.Solutions
            m_Med.SolutionVolume = objSol.SolutionVolume
            m_Med.MaxConc = objSol.MaxConc
            m_Med.MinInfusionTime = objSol.MinInfusionTime
            
            m_CalcVol = False
            LoadSolution
            m_CalcVol = True
            CalculateVolume
        End With
    End If

End Sub

Private Sub ApplyDoseRule()

    Dim objDose As ClassDose
    Dim dblAge As Double
    Dim dblWeight As Double
    Dim lngGest As Long
    
    If Not m_ValidMed Then Exit Sub
    
    dblAge = Patient_GetAgeInDays() / 30
    lngGest = Patient_GestationalAgeInDays()
    dblWeight = Patient_GetWeight()
    
    Set objDose = m_Med.GetDose("", "", dblAge, lngGest, dblWeight)
    
    If Not objDose Is Nothing Then
        
        With objDose
            m_Med.NormDose = objDose.NormDose
            m_Med.MinDose = objDose.MinDose
            m_Med.MaxDose = objDose.MaxDose
            m_Med.AbsMaxDose = objDose.AbsMaxDose
            m_Med.MaxPerDose = objDose.MaxPerDose
            
            m_Med.PerKg = objDose.IsDosePerKg
            m_Med.PerM2 = objDose.IsDosePerM2
            m_Med.SetFreqList objDose.Frequencies
        End With
        
    End If

End Sub

Private Sub cboRoute_Change()

    Dim strValid As String
    
    If Not m_Med Is Nothing Then m_Med.Route = cboRoute.Text
    strValid = ValidateCombo(cboRoute)
    HandleEvent MedicationChange, strValid
    
End Sub

Public Function GetTimeByFreq() As String

    Dim varFreq As Variant
    Dim strFreq As String
    Dim strPrev As String
    
    If cboFreq.Text = vbNullString Then
        For Each varFreq In cboFreq.List
            strPrev = strFreq
            strFreq = IIf(Not IsNull(varFreq), ModExcel.Excel_VLookup(varFreq, "Tbl_Glob_MedFreq", 2), strFreq)
            If strPrev <> vbNullString And strFreq <> strPrev Then
                GetTimeByFreq = vbNullString
                Exit Function
            End If
        Next
    Else
        strFreq = ModExcel.Excel_VLookup(cboFreq.Text, "Tbl_Glob_MedFreq", 2)
    End If
    
    GetTimeByFreq = strFreq

End Function

Public Function GetFactorByFreq(ByVal strFreq As String) As Double

    If Not m_Freq Is Nothing Then
        GetFactorByFreq = m_Freq.Item(strFreq)
    Else
        GetFactorByFreq = 0
    End If

End Function

Private Sub CalculateDose()

    Dim dblDose As Double
    Dim dblWght As Double
    Dim dblM2 As Double
    Dim dblAdjust As Double
    Dim dblVal As Double
    Dim dblCalc As Double
    Dim dblFact As Double
    Dim dblDeel As Double
    Dim dblKeer As Double
    
    If Not m_CalcDose Or Not m_ValidMed Then Exit Sub
    
    dblFact = IIf(chkPerDose.Value, 1, GetFactorByFreq(cboFreq.Text))
    dblDose = StringToDouble(txtNormDose.Value)
    dblWght = ModPatient.Patient_GetWeight()
    dblM2 = ModPatient.CalculateBSA()
    dblDeel = StringToDouble(txtMultipleQuantity.Value)
    dblKeer = StringToDouble(txtAdminDose.Value)
    dblKeer = IIf(dblDeel > 0, ModExcel.Excel_RoundBy(dblKeer, dblDeel), dblKeer)
    
    dblAdjust = IIf(optNone.Value, 1, dblWght)
    dblAdjust = IIf(optM2.Value, dblM2, dblAdjust)
    
    If dblFact = 0 Or (Not m_Keer And dblDose = 0) Or (m_Keer And dblKeer = 0) Or dblAdjust = 0 Or dblDeel = 0 Then
        If Not m_Keer Then SetTextBoxNumericValue txtAdminDose, 0
        SetTextBoxNumericValue txtCalcDose, 0
        Exit Sub
    End If
    
    If m_Keer Then
        
        dblCalc = dblKeer * dblFact / dblAdjust
    
    Else
    
        dblVal = dblDose * dblAdjust / dblFact
        dblVal = ModExcel.Excel_RoundBy(dblVal, txtMultipleQuantity.Value)
        
        dblCalc = dblVal * dblFact / dblAdjust
        dblKeer = dblCalc * dblAdjust / dblFact
    
    End If
        
    If m_Med Is Nothing Then Set m_Med = New ClassMedDisc
    m_Med.CalcDose = dblCalc
    m_Med.Freq = cboFreq.Text
    
    txtCalcDose.Value = ModString.FixPrecision(dblCalc, 2)
    If Not m_Keer Then txtAdminDose.Value = dblKeer
        

End Sub

Private Sub CalculateLabel()

    lblCalcDose.Caption = "Berekende dosering met deelbaarheid: " & txtMultipleQuantity.Value & " " & cboMultipleQuantityUnit.Text

End Sub

Private Sub CalculateVolume()

    Dim dblKeer As Double
    Dim dblConc As Double
    Dim dblCalc As Double
    Dim dblVol As Double
    
    dblKeer = StringToDouble(txtAdminDose.Value)
    dblConc = StringToDouble(txtMaxConc.Value)
    dblVol = StringToDouble(txtSolutionVolume.Value)
    
    If Not dblKeer > 0 Or Not m_CalcVol Or Not m_ValidMed Then Exit Sub
    
    If m_Conc And dblConc > 0 Then
        dblCalc = dblKeer / dblConc
        dblCalc = FixPrecision(dblCalc, 1)
        Do While dblKeer / dblCalc > dblConc
            dblCalc = dblCalc + 1
        Loop
        
        txtSolutionVolume.Value = dblCalc
    ElseIf Not m_Conc And dblVol > 0 Then
        dblCalc = dblKeer / dblVol
        dblCalc = ModString.FixPrecision(dblCalc, 1)
        
        If Not dblConc = 0 And dblCalc > dblConc Then
            cmdConc_Click
        Else
            txtMaxConc.Value = dblCalc
        End If
    End If

End Sub

Private Sub cboSubstance_Change()

    HandleEvent DoseRulesChange

End Sub

Private Sub cboSubstConc_Change()

    Dim objSubst As ClassSubstance

    lblSubstConc.Caption = cboGenericQuantityUnit.Text
    If cboSubstConc.Text = cboGeneric.Text Then txtSubstConc.Value = txtGenericQuantity.Value
    
    For Each objSubst In m_Med.Substances
        If objSubst.Substance = cboSubstConc.Text Then
            txtSubstConc.Text = objSubst.Concentration
            
            Exit For
        End If
    Next
    cboSubstance.Text = cboSubstConc.Text

End Sub

Private Sub chkDose_Click()
    
    HandleEvent HasDoseChange

End Sub

Private Sub ClearSolution()

    If Not m_Med Is Nothing Then
        m_Med.Solution = ""
        m_Med.SolutionVolume = 0
        m_Med.MinInfusionTime = 0
        m_Med.MaxConc = 0
    End If
    
    cboSolutions.Value = ""
    txtSolutionVolume.Value = ""
    txtTijd.Value = ""

End Sub

Private Sub ClearDose()

    If Not m_Med Is Nothing Then

        m_Med.SetFreqList ""
        m_Med.Freq = ""
        Set m_Freq = Nothing
        LoadFreq
        
        m_Med.PerDose = False
        m_Med.PerKg = True
        m_Med.PerM2 = False
        
        chkPerDose.Value = False
        optKg.Value = True
        optM2.Value = False
        
        m_Med.NormDose = 0
        m_Med.MinDose = 0
        m_Med.MaxDose = 0
        m_Med.AbsMaxDose = 0
        m_Med.MaxPerDose = 0
        
        txtNormDose.Value = ""
        txtMinDose.Value = ""
        txtMaxDose.Value = ""
        txtAbsMaxDose.Value = ""
        txtMaxPerDose = ""
        
        cboSolutions.Value = ""
        txtSolutionVolume.Value = ""
        txtTijd.Value = ""
        
        txtAdminDose.Value = ""
        txtCalcDose.Value = ""
        
    End If
        
End Sub

Private Sub chkPerDose_Click()

    ' SetDoseUnit
    ' CalculateDose
    
    HandleEvent DoseChange

End Sub

Private Sub cmdFormularium_Click()
    
    Dim strUrl As String
    
    strUrl = "https://www.kinderformularium.nl/"
    If Not m_Med Is Nothing Then
        If Not m_Med.ATC = vbNullString Then
            strUrl = strUrl & "geneesmiddelen?atc_code=" + m_Med.ATC
        Else
            strUrl = strUrl & "geneesmiddelen?name=" + cboGeneric.Text
        End If
    End If

    ActiveWorkbook.FollowHyperlink strUrl

End Sub

Private Sub CloseForm(ByVal strAction As String)

    lblButton.Caption = strAction
    Me.Hide

End Sub

Private Sub cmdCancel_Click()
    
    CloseForm "Cancel"

End Sub

Private Sub cmdClear_Click()

    ClearForm True
    CloseForm "Clear"
    
End Sub

Private Sub cmdGStand_Click()
    
    Dim strUrl As String
    Dim dblAge As Double
    Dim dblWeight As Double
    Dim strRoute As String
    Dim strGPK As String
    Dim strGen As String
    Dim strSHP As String
    Dim strUNT As String
    Dim strMsg As String
    
    strUrl = "http://vpxap-meta01.ds.umcutrecht.nl/GenForm/html?"
    
    If Not Patient_BirthDate() = ModDate.EmptyDate Then
        dblAge = Patient_CorrectedAgeInMo()
    Else
        strMsg = "Patient heeft geen geboortedatum."
    End If
    
    dblWeight = Patient_GetWeight()
    
    If Not cboRoute.Value = vbNullString Then
        strRoute = cboRoute.Value
    Else
        strMsg = "Kies eerst een route."
    End If
    
    If Not m_Med Is Nothing Then
        strGPK = m_Med.GPK
        strGen = m_Med.Generic
        strSHP = m_Med.Shape
        strUNT = m_Med.MultipleUnit
    Else
        strMsg = "Kies eerst een medicament."
    End If
    
    If Not strMsg = vbNullString Then
        ModMessage.ShowMsgBoxInfo strMsg
        Exit Sub
    End If
    
    strUrl = strUrl & "age=" & dblAge
    strUrl = strUrl & "&wht=" & dblWeight
    strUrl = strUrl & "&gpk=" & strGPK
    strUrl = strUrl & "&gen=" & strGen
    strUrl = strUrl & "&shp=" & strSHP
    strUrl = strUrl & "&rte=" & strRoute
    strUrl = strUrl & "&unt=" & strUNT
        
    ModUtils.CopyToClipboard strUrl
    ActiveWorkbook.FollowHyperlink strUrl

End Sub

Private Sub cmdAdminDose_Click()

    Dim objPic As StdPicture

    m_Keer = Not m_Keer
    Set objPic = cmdAdminDose.Picture
    cmdAdminDose.Picture = cmdNormDose.Picture
    cmdNormDose.Picture = objPic
    
    ToggleDoseText
    HandleEvent DoseChange
    ' CalculateDose
    
End Sub

Private Sub cmdKompas_Click()

    Dim strUrl As String
    Dim strFirst As String
    Dim strGeneric As String
    
    On Error GoTo ErrorHandler
    
    strUrl = "https://www.farmacotherapeutischkompas.nl/bladeren/preparaatteksten/{FIRST}/{GENERIC}"
    
    strGeneric = Trim(LCase(cboGeneric.Text))
    
    If strGeneric = vbNullString Then Exit Sub
    
    strGeneric = Replace(strGeneric, "-", "_")
    strGeneric = Replace(strGeneric, "+", "_")
    If Not strGeneric = vbNullString Then
        strFirst = Left(strGeneric, 1)
        strUrl = Replace(strUrl, "{FIRST}", strFirst)
        strUrl = Replace(strUrl, "{GENERIC}", strGeneric)
    End If
    
    ActiveWorkbook.FollowHyperlink strUrl
    
    Exit Sub
    
ErrorHandler:

    ActiveWorkbook.FollowHyperlink "https://www.farmacotherapeutischkompas.nl"

End Sub

Private Sub cmdMail_Click()

    m_Mail = True
    cmdOK_Click

End Sub

Private Sub cmdNormDose_Click()

    Dim objPic As StdPicture

    m_Keer = Not m_Keer
    Set objPic = cmdNormDose.Picture
    cmdNormDose.Picture = cmdAdminDose.Picture
    cmdAdminDose.Picture = objPic
    
    ToggleDoseText
    HandleEvent DoseChange
    '    CalculateDose

End Sub

Private Sub cmdConc_Click()

    ToggleConc
    HandleEvent SolutionChange
    ' CalculateVolume
    
End Sub

Private Sub cmdQuery_Click()

    If Not m_Med.GPK = "" Then
        ClearDose

        m_Med.Route = cboRoute.Text
        m_Med.MultipleUnit = cboMultipleQuantityUnit.Text
        
        ModWeb.Web_RetrieveMedicationRules m_Med
        HandleEvent DoseRulesChange
    End If

End Sub

Private Sub cmdVol_Click()

    ToggleConc
    CalculateVolume

End Sub

Private Sub cmdOK_Click()

    Dim intAnswer As Integer
    Dim strMsg As String
    Dim intIndx As Integer
    
    If Not lblValid.Caption = vbNullString Then
        strMsg = "Voorschrift is nog niet goed."
        strMsg = strMsg & vbNewLine & lblValid.Caption
        strMsg = strMsg & vbNewLine & "Toch invoeren?"
        intAnswer = ModMessage.ShowMsgBoxYesNo(strMsg)
        If intAnswer = vbNo Then Exit Sub
    End If

    If Not m_IsGPK Then
    
        Set m_Med = New ClassMedDisc
        
        m_Med.Generic = cboGeneric.Value
        m_Med.GenericQuantity = StringToDouble(txtGenericQuantity.Value)
        m_Med.GenericUnit = cboGenericQuantityUnit.Value
        m_Med.Shape = cboShape.Value
        
    End If
    
    m_Med.Route = cboRoute.Value
    m_Med.Indication = cboIndication.Value
    
    m_Med.HasDose = chkDose.Value
    If m_Med.HasDose Then
        m_Med.MultipleQuantity = StringToDouble(txtMultipleQuantity.Value)
        m_Med.MultipleUnit = cboMultipleQuantityUnit.Value
        
        m_Med.Freq = cboFreq.Value
        m_Med.PerDose = chkPerDose.Value
        m_Med.PerKg = optKg.Value
        m_Med.PerDose = optM2.Value
        
        m_Med.NormDose = StringToDouble(txtNormDose.Value)
        m_Med.MinDose = StringToDouble(txtMinDose.Value)
        m_Med.MaxDose = StringToDouble(txtMaxDose.Value)
        m_Med.AbsMaxDose = StringToDouble(txtAbsMaxDose.Value)
        m_Med.CalcDose = StringToDouble(txtCalcDose.Value)
        
        m_Med.Substance = cboSubstance.Text
        m_Med.AdminDose = StringToDouble(txtAdminDose.Value)
        
        m_Med.Solution = cboSolutions.Value
        m_Med.MaxConc = StringToDouble(txtMaxConc.Value)
        m_Med.SolutionVolume = StringToDouble(txtSolutionVolume.Value)
        m_Med.MinInfusionTime = StringToDouble(txtTijd.Value)
        
    End If
    
    CloseForm "OK"

End Sub

Private Sub SetProductDose()

    If Not m_Med Is Nothing Then
        m_Med.Substance = cboSubstance.Text
        m_Med.AdminDose = StringToDouble(txtAdminDose.Value)
    End If
    
End Sub

Private Function GetProductDose() As String

    Dim strText As String
    Dim strSubst As String
    Dim dblConc As Double
    Dim strConc As String
    Dim arrConc() As String
    Dim strUnit As String
    Dim dblQty As Double
    Dim objSubst As ClassSubstance
    
    strSubst = cboSubstance.Text
    For Each objSubst In m_Med.Substances
        If objSubst.Substance = strSubst Then
            dblConc = objSubst.Concentration
            strConc = cboGenericQuantityUnit.Text
            arrConc = Split(strConc, "/")
            If UBound(arrConc) = 1 Then strUnit = arrConc(1)
        End If
        
        If dblConc > 0 And Not strUnit = "" Then
            dblQty = dblConc * StringToDouble(txtAdminDose.Text)
            strText = strText & " (= " & cboFreq.Text & " " & dblQty & " " & strUnit & ")"
        End If
    Next

    GetProductDose = strText

End Function


Private Sub cmdParEnt_Click()

    Dim strUrl As String
    
    On Error Resume Next
    
    strUrl = "https://infoland-prod.umcutrecht.nl/iprova/Portaal/Handboek_Parenteralia/Zoeken/?Query=" & m_Med.Generic
    
    ActiveWorkbook.FollowHyperlink strUrl

End Sub

Private Sub optKg_Change()

    SetDoseUnit
    CalculateDose

End Sub

Private Sub optNone_Change()

    SetDoseUnit
    CalculateDose

End Sub

Private Sub optM2_Change()

    SetDoseUnit
    CalculateDose

End Sub

Private Sub txtAbsMaxDose_AfterUpdate()
    
    TextBoxStringNumericValue txtAbsMaxDose

End Sub

Private Sub txtAbsMaxDose_Change()

    HandleEvent DoseChange

End Sub

Private Sub txtAbsMaxDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtCalcDose_Change()

    If txtCalcDose.Value = 0 Then
        txtCalcDose.ForeColor = vbRed
        lblCalcDose.ForeColor = vbRed
    Else
        txtCalcDose.ForeColor = vbBlack
        lblCalcDose.ForeColor = vbBlack
    End If

End Sub

Private Sub txtCalcDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModMessage.ShowMsgBoxInfo "Deze waarde is berekend en kan niet worden gewijzigd!"
    KeyAscii = 0

End Sub

Private Sub txtMultipleQuantity_AfterUpdate()

    TextBoxStringNumericValue txtMultipleQuantity

End Sub

Private Sub txtMultipleQuantity_Change()

    CalculateDose

End Sub

Private Sub txtMultipleQuantity_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate vbNullString, False

End Sub

Private Sub cboShape_Change()

    Dim strValid As String
    
    strValid = ValidateCombo(cboShape)
    HandleEvent MedicationChange, strValid

End Sub

Private Sub txtMultipleQuantity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtAdminDose_AfterUpdate()

    Dim dblDeel As Double
    Dim dblKeer As Double
    
    dblDeel = StringToDouble(txtMultipleQuantity.Value)
    dblKeer = StringToDouble(txtAdminDose.Value)
    dblKeer = IIf(dblDeel > 0, ModExcel.Excel_RoundBy(dblKeer, dblDeel), dblKeer)

    SetTextBoxNumericValue txtAdminDose, dblKeer
    
    If m_Keer Then CalculateDose
    
    If dblKeer = 0 Then
        ModMessage.ShowMsgBoxExclam "Deze keerdosering kan niet met een deelbaarheid van " & txtMultipleQuantity.Text
    End If
    
    Validate vbNullString, False
    ApplySolution

End Sub

Private Sub txtAdminDose_Change()

    If m_Keer Then CalculateDose
    ApplySolution
    Validate vbNullString, True

End Sub

Private Sub txtAdminDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtMaxConc_Change()

    CalculateVolume

End Sub

Private Sub txtSolutionVolume_Change()
    
    CalculateVolume

End Sub

Private Sub txtMaxDose_AfterUpdate()

    TextBoxStringNumericValue txtMaxDose

End Sub

Private Sub txtMaxDose_Change()

    Validate vbNullString, True

End Sub

Private Sub txtMaxDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtMinDose_AfterUpdate()

    TextBoxStringNumericValue txtMinDose

End Sub

Private Sub txtMinDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtNormDose_AfterUpdate()

    TextBoxStringNumericValue txtNormDose
    
End Sub

Private Sub txtNormDose_Change()

    If Not m_Keer Then CalculateDose
    Validate vbNullString, True

End Sub

Private Sub txtNormDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtGenericQuantity_Change()

    Dim strGen As String
    Dim strSub As String
    
    strSub = Trim(LCase(cboSubstConc.Text))
    strGen = Trim(LCase(cboGeneric.Text))
    
    If strSub = strGen Then
        txtSubstConc.Value = txtGenericQuantity.Value
    End If
    

End Sub

Private Sub txtSubstConc_Change()

    Dim objSubst As ClassSubstance
    
    For Each objSubst In m_Med.Substances
        If objSubst.Substance = cboSubstConc.Text Then
            objSubst.Concentration = StringToDouble(txtSubstConc.Text)
        End If
    Next

End Sub

Private Sub txtSubstConc_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtGenericQuantity_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate vbNullString, False

End Sub

Private Sub txtGenericQuantity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub cboGenericQuantityUnit_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim strValid As String
    
    strValid = ValidateCombo(cboGenericQuantityUnit)
    HandleEvent MedicationChange, strValid

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()
    
    CenterForm
    
    Validate vbNullString, False
    
    cboGeneric.SetFocus

End Sub

Private Sub LoadFreq()

    Dim varKey As Variant
    
    cboFreq.Clear
    cboFreqTime.Clear
    
    If m_Freq Is Nothing Then
        Set m_Freq = ModMedDisc.MedDisc_GetMedicationFreqs()
    End If
    
    For Each varKey In m_Freq.Keys
        cboFreq.AddItem varKey
    Next

End Sub

Public Sub CaclculateWithKeerDose(ByVal dblKeer As Double)

    txtAdminDose.Value = dblKeer
    cmdAdminDose_Click

End Sub

Private Sub UserForm_Initialize()

    Dim strTitle As String
    Dim objMedCol As Collection
    Dim objMed As ClassMedDisc
    
    m_LoadGPK = False
    m_Keer = False
    m_Mail = False
    
    m_TherapieGroep = lblMainGroup.Caption
    m_SubGroep = lblSubGroup.Caption
    m_Etiket = lblEtiket.Caption
    m_Product = lblProduct.Caption
        
    Set objMedCol = Formularium_GetFormularium.GetMedicationCollection(False)
    For Each objMed In objMedCol
        cboGeneric.AddItem objMed.Generic
    Next
    
    LoadFreq
    FillCombo cboSolutions, MedDisc_GetOplVlstCol()
    
    optKg.Value = True
    
    cboGeneric.TabIndex = 0
    cboShape.TabIndex = 1
    txtGenericQuantity.TabIndex = 2
    cboGenericQuantityUnit.TabIndex = 3
    txtMultipleQuantity.TabIndex = 4
    cboMultipleQuantityUnit.TabIndex = 5
    cboRoute.TabIndex = 6
    cboIndication.TabIndex = 7
    
    cboFreq.TabIndex = 8
    txtAdminDose.TabIndex = 9
    txtNormDose.TabIndex = 10
    txtMinDose.TabIndex = 11
    txtMaxDose.TabIndex = 12
    txtAbsMaxDose.TabIndex = 13
    
    cmdFormularium.TabIndex = 14
    cmdOK.TabIndex = 15
    cmdClear.TabIndex = 16
    cmdCancel.TabIndex = 17
    
    cboDoseUnit.TabStop = False
    cboAdminDoseUnit.TabStop = False
    txtCalcDose.TabStop = False
       
    cboDoseUnit.Enabled = False
    txtAdminDose.Enabled = False
    cboAdminDoseUnit.Enabled = False
    txtMaxConc.Enabled = False
    
    chkDose.Value = True
    
    m_CalcVol = True
    m_CalcDose = True
       
End Sub

Private Sub ToggleDoseText()

    txtAdminDose.Enabled = m_Keer
    txtNormDose.Enabled = Not m_Keer

End Sub

Private Sub ToggleConc()

    Dim objPic As StdPicture

    m_Conc = Not m_Conc
    
    If m_Conc Then
        txtSolutionVolume.Value = 0
    Else
        txtMaxConc.Value = 0
    End If
    
    Set objPic = cmdConc.Picture
    cmdConc.Picture = cmdVol.Picture
    cmdVol.Picture = objPic

    txtMaxConc.Enabled = m_Conc
    txtSolutionVolume.Enabled = Not m_Conc

End Sub

Private Sub UserForm_QueryClose(intCancel As Integer, intMode As Integer)
    
    intCancel = True
    cmdCancel_Click

End Sub


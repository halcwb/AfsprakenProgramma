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

Private m_Med As ClassMedicatieDisc
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
    
    IsAbsMaxInvalid = txtAbsMax.Value = vbNullString And ModPatient.Patient_GetWeight() > 50 And txtNormDose.Value = vbNullString And txtMaxDose.Value = vbNullString

End Function

Private Function IsDoseControlInValid() As Boolean

    IsDoseControlInValid = txtNormDose.Value = vbNullString And txtMaxDose.Value = vbNullString

End Function

Private Sub Validate(ByVal strValid As String)
    
    If strValid = vbNullString Then
    
        If frmDose.Visible Then
            strValid = IIf(cboDosisEenheid.Value = vbNullString, "Voer dosering eenheid in", strValid)
            strValid = IIf(txtDeelDose.Value = vbNullString, "Voer een deelbaarheid in", strValid)
        End If
        
        strValid = IIf(cboIndicatie.Value = vbNullString, "Kies een indicatie", strValid)
        strValid = IIf(cboRoute.Value = vbNullString, "Kies een route", strValid)
        
        strValid = IIf(cboSterkteEenheid.Value = vbNullString, "Voer sterkte eenheid in", strValid)
        strValid = IIf(txtSterkte.Value = vbNullString, "Voer sterkte in", strValid)
        
        strValid = IIf(cboVorm.Value = vbNullString, "Voer een vorm in", strValid)
        strValid = IIf(cboGeneriek.Value = vbNullString, "Kies een generiek", strValid)
    
    End If
    
    cmdOK.Enabled = strValid = vbNullString
    
    If frmDose.Visible And strValid = vbNullString Then
        strValid = IIf(IsDoseControlInValid, "Voer of een advies dosering in en/of een max (en evt. min en abs max) dosering", strValid)
        strValid = IIf(IsAbsMaxInvalid, "Gewicht boven de 50 kg, voer een absolute maximum dosering in (of een advies dosering of max dosering)", strValid)
    
    End If
    
    lblValid.Caption = strValid

End Sub

Public Function HasSelectedMedicament() As Boolean

    HasSelectedMedicament = Not m_Med Is Nothing

End Function

Public Function GetSelectedMedicament() As ClassMedicatieDisc

    Set GetSelectedMedicament = m_Med

End Function

Public Function GetClickedButton() As String

    GetClickedButton = lblButton.Caption

End Function

Private Sub SetToGPKMode(ByVal blnIsGPK As Boolean)

    Dim varItem As Variant
    
    m_IsGPK = blnIsGPK
    
    Me.txtSterkte.Enabled = Not blnIsGPK
    Me.cboSterkteEenheid.Enabled = Not blnIsGPK
    Me.cboVorm.Enabled = Not blnIsGPK
    
    cmdFormularium.Enabled = blnIsGPK
    
    If Not blnIsGPK Then
        lblGPK.Caption = vbNullString
        lblATC.Caption = vbNullString
        
        FillCombo cboVorm, Formularium_GetFormularium().GetVormen()
        FillCombo cboSterkteEenheid, Formularium_GetFormularium().GetSterkteEenheden()
        FillCombo cboDosisEenheid, Formularium_GetFormularium.GetDosisEenheden()
        FillCombo cboRoute, Formularium_GetFormularium.GetRoutes()
        
        cboIndicatie.Clear
        
        LoadFreq
        
        FillCombo cboOplVlst, MedDisc_GetOplVlstCol()
    End If

End Sub

Private Sub cboDosisEenheid_Change()

    SetDoseUnit

End Sub

Private Sub cboFreq_Change()

    Dim strValid As String
    
    strValid = ValidateCombo(cboFreq)
    
    If strValid = vbNullString Then
        SetDoseUnit
        CalculateDose
        Validate vbNullString
    Else
        Validate strValid
    End If
    
End Sub

Private Sub cboFreqTime_Change()

    SelectDoseRule

End Sub

Private Sub SelectDoseRule()

    Dim objRule As ClassDoseRule
    Dim strTime As String
    Dim strFreq As String
    
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
                
                chkPerDosis = .PerDose
                
                SetTextBoxNumericValue txtNormDose, .NormDose
                SetTextBoxNumericValue txtMinDose, .MinDose
                SetTextBoxNumericValue txtMaxDose, .MaxDose
                SetTextBoxNumericValue txtAbsMax, .AbsMaxDose
                SetTextBoxNumericValue txtMaxPerDose, .MaxPerDose
            End With
            
        End If
    Next
    
    CalculateDose

End Sub


Private Sub cboGeneriek_Click()

    cboGeneriek_Change

End Sub

Private Sub cboGeneriek_Change()
    
    If m_LoadGPK Then Exit Sub

    If cboGeneriek.ListIndex > -1 Then
        SetToGPKMode True
        Set m_Med = Formularium_GetFormularium.Item(cboGeneriek.ListIndex + 1)
        LoadMedicament True
    Else
        SetToGPKMode False
        ClearForm False
    End If
    
    SetDoseUnit
    Validate vbNullString

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
        LoadMedicament True
        m_LoadGPK = True
        cboGeneriek.Text = m_Med.Generic
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
    strUnit = Trim(cboDosisEenheid.Text)
    strTime = IIf(chkPerDosis.Value, "dosis", GetTimeByFreq())
    
    If Not strUnit = vbNullString Then
        If Not strTime = vbNullString Then
            cboDoseUnit.Text = strUnit & (IIf(strAdjust = "", "/", "/" & strAdjust & "/")) & strTime
        Else
            cboDoseUnit.Text = strUnit & (IIf(strAdjust = "", "", "/" & strAdjust))
        End If
        
        cboKeerUnit.Text = strUnit
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

Private Sub LoadMedicament(ByVal blnReload As Boolean)
    
    Dim intN As Integer
    Dim strFreq As String
    Dim arrFreq() As String
    Dim strTime As String
    Dim objRule As ClassDoseRule

    With m_Med
    
        lblTherapieGroep.Caption = .MainGroup
        lblSubGroep.Caption = .SubGroup
        lblEtiket.Caption = .Label
        lblProduct.Caption = .Product
        
        lblGPK.Caption = .GPK
        lblATC.Caption = .ATC
        
        cboVorm.Value = .Shape
        
        txtSterkte.Text = .GenericQuantity
        cboSterkteEenheid.Text = .GenericUnit
        
        If blnReload Or cboRoute.Text = "" Then FillCombo cboRoute, .GetRouteList()
        If blnReload Or cboIndicatie.Text = "" Then FillCombo cboIndicatie, .GetIndicationList()
        
        If .MultipleQuantity = 0 Then
            chkDose.Value = False
            chkDose_Click
            
            Exit Sub
        End If
        
        FillCombo cboSubstConc, GetSubstances()
        
        txtDeelDose.Text = .MultipleQuantity
        
        FillCombo cboDosisEenheid, GetDosisEenheden()
        cboDosisEenheid.Text = .MultipleUnit
        
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
        
        chkPerDosis = .PerDose
        
        SetTextBoxNumericValue txtNormDose, .NormDose
        SetTextBoxNumericValue txtMinDose, .MinDose
        SetTextBoxNumericValue txtMaxDose, .MaxDose
        SetTextBoxNumericValue txtAbsMax, .AbsMaxDose
        SetTextBoxNumericValue txtMaxPerDose, .MaxPerDose
        
        SetTextBoxNumericValue txtMaxConc, .MaxConc
        If .MaxConc > 0 And Not m_Conc Then cmdConc_Click
        
        SetTextBoxNumericValue txtOplVol, .Solution
        SetTextBoxNumericValue txtTijd, .MinInfusionTime
        
        cboOplVlst.Value = .Solution
        
        If m_Med.DoseRules.Count > 1 Then
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

Private Function GetDosisEenheden() As Collection

    Set GetDosisEenheden = Formularium_GetFormularium.GetDosisEenheden()

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

    lblTherapieGroep.Caption = m_TherapieGroep
    lblSubGroep.Caption = m_SubGroep
    lblEtiket.Caption = m_Etiket
    lblProduct.Caption = m_Product
    
    If blnClearGeneric Then cboGeneriek.Value = vbNullString
    cboVorm.Value = vbNullString
    
    txtDeelDose.Value = vbNullString
    cboDosisEenheid.Clear
    cboDosisEenheid.Value = vbNullString
    
    txtSterkte.Text = vbNullString
    cboSterkteEenheid.Text = vbNullString
    
    FillCombo cboRoute, Formularium_GetFormularium.GetRoutes()
    
    cboIndicatie.Clear
    cboIndicatie.Value = vbNullString
    
    cboFreq.Value = vbNullString
    txtNormDose.Value = vbNullString
    txtMinDose.Value = vbNullString
    txtMaxDose.Value = vbNullString
    txtAbsMax.Value = vbNullString
    txtMaxPerDose.Value = vbNullString
    txtCalcDose.Value = vbNullString
    cboDoseUnit.Value = vbNullString
    lblCalcDose.Caption = vbNullString
    lblDoseUnit.Caption = vbNullString
    
    cboGeneriek.SetFocus
    
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

Private Sub cboIndicatie_Change()

    Dim strValid As String
    
    ' strValid = ValidateCombo(cboIndicatie, False)
    Validate vbNullString

End Sub

Private Sub cboRoute_Change()

    Dim strValid As String
    
    strValid = ValidateCombo(cboRoute)
    Validate strValid

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
    
    dblFact = IIf(chkPerDosis.Value, 1, GetFactorByFreq(cboFreq.Text))
    dblDose = StringToDouble(txtNormDose.Value)
    dblWght = ModPatient.Patient_GetWeight()
    dblM2 = ModPatient.CalculateBSA()
    dblDeel = StringToDouble(txtDeelDose.Value)
    dblKeer = StringToDouble(txtKeerDose.Value)
    dblKeer = IIf(dblDeel > 0, ModExcel.Excel_RoundBy(dblKeer, dblDeel), dblKeer)
    
    dblAdjust = IIf(optNone.Value, 1, dblWght)
    dblAdjust = IIf(optM2.Value, dblM2, dblAdjust)
    
    If dblFact = 0 Or (Not m_Keer And dblDose = 0) Or (m_Keer And dblKeer = 0) Or dblAdjust = 0 Or dblDeel = 0 Then
        SetTextBoxNumericValue txtKeerDose, 0
        SetTextBoxNumericValue txtCalcDose, 0
        Exit Sub
    End If
    
    If m_Keer Then
        
        dblCalc = dblKeer * dblFact / dblAdjust
    
    Else
    
        dblVal = dblDose * dblAdjust / dblFact
        dblVal = ModExcel.Excel_RoundBy(dblVal, txtDeelDose.Value)
        
        dblCalc = dblVal * dblFact / dblAdjust
        dblKeer = dblCalc * dblAdjust / dblFact
    
    End If
        
    If m_Med Is Nothing Then Set m_Med = New ClassMedicatieDisc
    m_Med.CalcDose = dblCalc
    m_Med.Freq = cboFreq.Text
    
    txtCalcDose.Value = ModString.FixPrecision(dblCalc, 2)
    If Not m_Keer Then txtKeerDose.Value = dblKeer
        
    CalculateVolume
    CalculateLabel
    SetProductDose

End Sub

Private Sub CalculateLabel()

    lblCalcDose.Caption = "Berekende dosering met deelbaarheid: " & txtDeelDose.Value & " " & cboDosisEenheid.Text

End Sub

Private Sub CalculateVolume()

    Dim dblKeer As Double
    Dim dblConc As Double
    Dim dblVol As Double
    
    dblKeer = StringToDouble(txtKeerDose.Value)
    dblConc = StringToDouble(txtMaxConc.Value)
    dblVol = StringToDouble(txtOplVol.Value)
    If m_Conc And dblConc > 0 Then
        dblVol = dblKeer / dblConc
        txtOplVol.Value = ModExcel.Excel_RoundBy(dblVol, 1)
    ElseIf Not m_Conc And dblVol > 0 Then
        dblConc = dblKeer / dblVol
        dblConc = ModString.FixPrecision(dblConc, 1)
        txtMaxConc.Value = dblConc
    End If

End Sub

Private Sub cboSubstance_Change()

    SelectDoseRule

End Sub

Private Sub cboSubstConc_Change()

    Dim objSubst As ClassSubstance

    lblSubstConc.Caption = cboSterkteEenheid.Text
    If cboSubstConc.Text = cboGeneriek.Text Then txtSubstConc.Value = txtSterkte.Value
    
    For Each objSubst In m_Med.Substances
        If objSubst.Substance = cboSubstConc.Text Then
            txtSubstConc.Text = objSubst.Concentration
            
            Exit For
        End If
    Next
    cboSubstance.Text = cboSubstConc.Text

End Sub

Private Sub chkDose_Click()
    
    If Not frmDose.Visible Then ClearDose
    
    frmDose.Visible = chkDose.Value
    frmOpl.Visible = chkDose.Value
    txtDeelDose.Visible = chkDose.Value
    cboDosisEenheid.Visible = chkDose.Value

    Validate vbNullString

End Sub

Private Sub ClearDose()

        m_Med.SetFreqList ""
        Set m_Freq = Nothing
        LoadFreq
        
        m_Med.PerDose = False
        m_Med.PerKg = True
        m_Med.PerM2 = False
        
        chkPerDosis.Value = False
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
        txtAbsMax.Value = ""
        txtMaxPerDose = ""
        
        cboOplVlst.Value = ""
        txtOplVol.Value = ""
        txtTijd.Value = ""
        
        txtKeerDose.Value = ""
        txtCalcDose.Value = ""
        
End Sub

Private Sub chkPerDosis_Click()

    SetDoseUnit
    CalculateDose

End Sub

Private Sub cmdFormularium_Click()
    
    Dim strUrl As String
    
    strUrl = "https://www.kinderformularium.nl/"
    If Not m_Med Is Nothing Then
        If Not m_Med.ATC = vbNullString Then
            strUrl = strUrl & "geneesmiddelen?atc_code=" + m_Med.ATC
        Else
            strUrl = strUrl & "geneesmiddelen?name=" + cboGeneriek.Text
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
    
    strUrl = "http://iis2503.ds.umcutrecht.nl/GenForm/html?"
    
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

Private Sub cmdKeerDose_Click()

    Dim objPic As StdPicture

    m_Keer = Not m_Keer
    Set objPic = cmdKeerDose.Picture
    cmdKeerDose.Picture = cmdNormDose.Picture
    cmdNormDose.Picture = objPic
    
    ToggleDoseText

    CalculateDose
    
End Sub

Private Sub cmdMail_Click()

    m_Mail = True
    cmdOK_Click

End Sub

Private Sub cmdNormDose_Click()

    Dim objPic As StdPicture

    m_Keer = Not m_Keer
    Set objPic = cmdNormDose.Picture
    cmdNormDose.Picture = cmdKeerDose.Picture
    cmdKeerDose.Picture = objPic
    
    ToggleDoseText
    
    CalculateDose

End Sub

Private Sub cmdConc_Click()

    Dim objPic As StdPicture

    m_Conc = Not m_Conc
    Set objPic = cmdConc.Picture
    cmdConc.Picture = cmdVol.Picture
    cmdVol.Picture = objPic
    
    ToggleConcText

    CalculateDose
    
End Sub

Private Sub cmdQuery_Click()

    If Not m_Med.GPK = "" Then
        m_Med.Route = cboRoute.Text
        m_Med.MultipleUnit = cboDosisEenheid.Text
        
        ModWeb.Web_RetrieveMedicationRules m_Med
        LoadMedicament False
    End If

End Sub

Private Sub cmdVol_Click()

    Dim objPic As StdPicture

    m_Conc = Not m_Conc
    Set objPic = cmdVol.Picture
    cmdVol.Picture = cmdConc.Picture
    cmdConc.Picture = objPic
    
    ToggleConcText
    
    CalculateDose

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
    
        Set m_Med = New ClassMedicatieDisc
        
        m_Med.Generic = cboGeneriek.Value
        m_Med.GenericQuantity = StringToDouble(txtSterkte.Value)
        m_Med.GenericUnit = cboSterkteEenheid.Value
        m_Med.Shape = cboVorm.Value
        
    End If
    
    m_Med.Route = cboRoute.Value
    m_Med.Indication = cboIndicatie.Value
    
    m_Med.HasDose = chkDose.Value
    If m_Med.HasDose Then
        m_Med.MultipleQuantity = StringToDouble(txtDeelDose.Value)
        m_Med.MultipleUnit = cboDosisEenheid.Value
        
        m_Med.Freq = cboFreq.Value
        m_Med.PerDose = chkPerDosis.Value
        m_Med.PerKg = optKg.Value
        m_Med.PerDose = optM2.Value
        
        m_Med.NormDose = StringToDouble(txtNormDose.Value)
        m_Med.MinDose = StringToDouble(txtMinDose.Value)
        m_Med.MaxDose = StringToDouble(txtMaxDose.Value)
        m_Med.AbsMaxDose = StringToDouble(txtAbsMax.Value)
        m_Med.CalcDose = StringToDouble(txtCalcDose.Value)
        
        m_Med.Substance = cboSubstance.Text
        m_Med.KeerDose = StringToDouble(txtKeerDose.Value)
        
        m_Med.Solution = cboOplVlst.Value
        m_Med.MaxConc = StringToDouble(txtMaxConc.Value)
        m_Med.SolutionVolume = StringToDouble(txtOplVol.Value)
        m_Med.MinInfusionTime = StringToDouble(txtTijd.Value)
        
    End If
    
    CloseForm "OK"

End Sub

Private Sub SetProductDose()

    m_Med.Substance = cboSubstance.Text
    m_Med.KeerDose = StringToDouble(txtKeerDose.Value)

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
            strConc = cboSterkteEenheid.Text
            arrConc = Split(strConc, "/")
            If UBound(arrConc) = 1 Then strUnit = arrConc(1)
        End If
        
        If dblConc > 0 And Not strUnit = "" Then
            dblQty = dblConc * StringToDouble(txtKeerDose.Text)
            strText = strText & " (= " & cboFreq.Text & " " & dblQty & " " & strUnit & ")"
        End If
    Next

    GetProductDose = strText

End Function


Private Sub cmdParEnt_Click()

    Dim strUrl As String
    
    If MetaVision_IsNeonatologie() Then
        strUrl = "https://neonaten-umcutrecht.parenteralia.nl/"
    Else
        strUrl = "https://kinderen-ic-umcutrecht.parenteralia.nl/"
    End If
    
    ActiveWorkbook.FollowHyperlink strUrl

End Sub

Private Sub optKg_Change()

    SetDoseUnit

End Sub

Private Sub optNone_Change()

    SetDoseUnit

End Sub

Private Sub optM2_Change()

    SetDoseUnit

End Sub

Private Sub txtAbsMax_AfterUpdate()
    
    TextBoxStringNumericValue txtAbsMax

End Sub

Private Sub txtAbsMax_Change()

    Validate vbNullString

End Sub

Private Sub txtAbsMax_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
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

Private Sub txtDeelDose_AfterUpdate()

    TextBoxStringNumericValue txtDeelDose

End Sub

Private Sub txtDeelDose_Change()

    CalculateDose

End Sub

Private Sub txtDeelDose_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate vbNullString

End Sub

Private Sub cboVorm_Change()

    Dim strValid As String
    
    strValid = ValidateCombo(cboVorm)
    Validate strValid

End Sub

Private Sub txtDeelDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtKeerDose_AfterUpdate()

    Dim dblDeel As Double
    Dim dblKeer As Double
    
    dblDeel = StringToDouble(txtDeelDose.Value)
    dblKeer = StringToDouble(txtKeerDose.Value)
    dblKeer = IIf(dblDeel > 0, ModExcel.Excel_RoundBy(dblKeer, dblDeel), dblKeer)

    SetTextBoxNumericValue txtKeerDose, dblKeer
    
    If dblKeer = 0 Then
        ModMessage.ShowMsgBoxExclam "Deze keerdosering kan niet met een deelbaarheid van " & txtDeelDose.Text
    End If
    

End Sub

Private Sub txtKeerDose_Change()

    If m_Keer Then CalculateDose
    Validate vbNullString

End Sub

Private Sub txtKeerDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtMaxConc_Change()

    CalculateVolume

End Sub

Private Sub txtOplVol_Change()
    
    CalculateVolume

End Sub

Private Sub txtMaxDose_AfterUpdate()

    TextBoxStringNumericValue txtMaxDose

End Sub

Private Sub txtMaxDose_Change()

    Validate vbNullString

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
    Validate vbNullString

End Sub

Private Sub txtNormDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtSterkte_Change()

    Dim strGen As String
    Dim strSub As String
    
    strSub = Trim(LCase(cboSubstConc.Text))
    strGen = Trim(LCase(cboGeneriek.Text))
    
    If strSub = strGen Then
        txtSubstConc.Value = txtSterkte.Value
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

Private Sub txtSterkte_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate vbNullString

End Sub

Private Sub txtSterkte_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub cboSterkteEenheid_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim strValid As String
    
    strValid = ValidateCombo(cboSterkteEenheid)
    Validate strValid

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()
    
    CenterForm
    
    Validate vbNullString
    
    cboGeneriek.SetFocus

End Sub

Private Sub LoadFreq()

    Dim varKey As Variant
    
    cboFreq.Clear
    
    If m_Freq Is Nothing Then
        Set m_Freq = ModMedDisc.GetMedicationFreqs()
    End If
    
    For Each varKey In m_Freq.Keys
        cboFreq.AddItem varKey
    Next

End Sub

Public Sub CaclculateWithKeerDose(ByVal dblKeer As Double)

    txtKeerDose.Value = dblKeer
    cmdKeerDose_Click

End Sub

Private Sub UserForm_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    m_LoadGPK = False
    m_Keer = False
    m_Mail = False
    
    m_TherapieGroep = lblTherapieGroep.Caption
    m_SubGroep = lblSubGroep.Caption
    m_Etiket = lblEtiket.Caption
    m_Product = lblProduct.Caption
    
    intC = Formularium_GetFormularium.MedicamentCount + 1
    
    For intN = 1 To intC
        cboGeneriek.AddItem Formularium_GetFormularium.Item(intN).Generic
    Next intN
    
    LoadFreq
    FillCombo cboOplVlst, MedDisc_GetOplVlstCol()
    
    optKg.Value = True
    
    cboGeneriek.TabIndex = 0
    cboVorm.TabIndex = 1
    txtSterkte.TabIndex = 2
    cboSterkteEenheid.TabIndex = 3
    txtDeelDose.TabIndex = 4
    cboDosisEenheid.TabIndex = 5
    cboRoute.TabIndex = 6
    cboIndicatie.TabIndex = 7
    
    cboFreq.TabIndex = 8
    txtKeerDose.TabIndex = 9
    txtNormDose.TabIndex = 10
    txtMinDose.TabIndex = 11
    txtMaxDose.TabIndex = 12
    txtAbsMax.TabIndex = 13
    
    cmdFormularium.TabIndex = 14
    cmdOK.TabIndex = 15
    cmdClear.TabIndex = 16
    cmdCancel.TabIndex = 17
    
    cboDoseUnit.TabStop = False
    cboKeerUnit.TabStop = False
    txtCalcDose.TabStop = False
       
    cboDoseUnit.Enabled = False
    txtKeerDose.Enabled = False
    cboKeerUnit.Enabled = False
    txtMaxConc.Enabled = False
    
    chkDose.Value = True
       
End Sub

Private Sub ToggleDoseText()

    txtKeerDose.Enabled = m_Keer
    txtNormDose.Enabled = Not m_Keer

End Sub

Private Sub ToggleConcText()

    txtMaxConc.Enabled = m_Conc
    txtOplVol.Enabled = Not m_Conc

End Sub

Private Sub UserForm_QueryClose(intCancel As Integer, intMode As Integer)
    
    intCancel = True
    cmdCancel_Click

End Sub


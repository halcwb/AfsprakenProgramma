VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedDisc 
   Caption         =   "Kies een medicament ..."
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
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
Private m_Formularium As ClassFormularium
Private m_TherapieGroep As String
Private m_SubGroep As String
Private m_Etiket As String

Private m_IsGPK As Boolean
Private m_LoadGPK As Boolean

Private m_Freq As Dictionary


Public Sub SetNoFormMed()

    m_IsGPK = False

End Sub

Private Function IsAbsMaxInvalid() As Boolean
    
    IsAbsMaxInvalid = txtAbsMax.Value = vbNullString And ModPatient.GetGewichtFromRange() > 50

End Function

Private Function IsDoseControlInValid() As Boolean

    IsDoseControlInValid = txtNormDose.Value = vbNullString And txtMaxDose.Value = vbNullString

End Function

Private Sub Validate(ByVal strValid As String)
    
    If strValid = vbNullString Then
    
        strValid = IIf(IsDoseControlInValid And Not IsAbsMaxInvalid, "Voer of een norm dosering in en/of een max (en evt. min en abs max) dosering", strValid)
        strValid = IIf(IsAbsMaxInvalid, "Gewicht boven de 50 kg, voer een absolute maximum dosering in", strValid)
    
        strValid = IIf(cboIndicatie.Value = vbNullString, "Kies een indicatie", strValid)
        strValid = IIf(cboRoute.Value = vbNullString, "Kies een route", strValid)
        
        strValid = IIf(cboDosisEenheid.Value = vbNullString, "Voer dosering grootte in", strValid)
        strValid = IIf(txtDeelDose.Value = vbNullString, "Voer dosering eenheid in", strValid)
        strValid = IIf(cboSterkteEenheid.Value = vbNullString, "Voer sterkte eenheid in", strValid)
        strValid = IIf(txtSterkte.Value = vbNullString, "Voer sterkte in", strValid)
        
        strValid = IIf(cboVorm.Value = vbNullString, "Voer een vorm in", strValid)
        strValid = IIf(cboGeneriek.Value = vbNullString, "Kies een generiek", strValid)
    
    End If
    
    lblValid.Caption = strValid
    cmdCalc.Enabled = Not cboFreq.Value = vbNullString And Not txtNormDose.Value = vbNullString
    cmdOK.Enabled = strValid = vbNullString

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
        
        FillCombo cboVorm, m_Formularium.GetVormen()
        FillCombo cboSterkteEenheid, m_Formularium.GetSterkteEenheden()
        FillCombo cboDosisEenheid, m_Formularium.GetDosisEenheden()
        FillCombo cboRoute, m_Formularium.GetRoutes()
        
        cboIndicatie.Clear
        
        LoadFreq
        
    End If

End Sub

Private Sub cboFreq_Change()

    Validate vbNullString

End Sub

Private Sub cboGeneriek_Change()
    
    If m_LoadGPK Then Exit Sub

    If cboGeneriek.ListIndex > -1 Then
        SetToGPKMode True
        Set m_Med = m_Formularium.Item(cboGeneriek.ListIndex + 1)
        LoadMedicament
    Else
        SetToGPKMode False
        ClearForm False
    End If
    
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
    
    Set m_Med = m_Formularium.GPK(strGPK)
    
    If m_Med Is Nothing Then
        SetToGPKMode False
        blnLoad = False
    Else
        SetToGPKMode True
        LoadMedicament
        m_LoadGPK = True
        cboGeneriek.Text = m_Med.Generiek
        m_LoadGPK = False
    End If
    
    LoadGPK = blnLoad

End Function

Private Sub LoadMedicament()
    
    Dim intN As Integer
    Dim arrFreq() As String

    With m_Med
    
        lblTherapieGroep.Caption = .TherapieGroep
        lblSubGroep.Caption = .TherapieSubgroep
        lblEtiket.Caption = .Etiket
        
        lblGPK.Caption = .GPK
        lblATC.Caption = .ATC
        
        cboVorm.Value = .Vorm
        
        txtSterkte.Text = .Sterkte
        cboSterkteEenheid.Text = .SterkteEenheid
        
        txtDeelDose.Text = .DeelDose
        
        FillCombo cboDosisEenheid, GetDosisEenheden()
        cboDosisEenheid.Text = .DoseEenheid
        
        FillCombo cboRoute, .GetRouteList()
        FillCombo cboIndicatie, .GetIndicatieList()
        
        If Not .DoseEenheid = vbNullString Then cboDoseUnit.Text = .DoseEenheid & "/kg/dag"
        
        If Not .GetFreqListString = vbNullString Then
            FillCombo cboFreq, .GetFreqList()
        Else
            LoadFreq
        End If
        
        If Not .Freq = vbNullString Then cboFreq.Value = .Freq
        
        SetTextBoxNumericValue txtNormDose, .NormDose
        SetTextBoxNumericValue txtMinDose, .MinDose
        SetTextBoxNumericValue txtMaxDose, .MaxDose
        SetTextBoxNumericValue txtAbsMax, .AbsDose
        
    End With

End Sub

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

    Set GetDosisEenheden = m_Formularium.GetDosisEenheden()

End Function

Private Sub FillCombo(objCombo As ComboBox, colItems As Collection)

    Dim varItem As Variant

    objCombo.Clear
    
    For Each varItem In colItems
        varItem = CStr(varItem)
        If Not varItem = vbNullString Then objCombo.AddItem CStr(varItem)
    Next varItem
    
    If objCombo.ListCount = 1 Then
        objCombo.Text = objCombo.List(0)
    End If
    
End Sub

Public Sub ClearForm(ByVal blnClearGeneric As Boolean)

    lblTherapieGroep.Caption = m_TherapieGroep
    lblSubGroep.Caption = m_SubGroep
    lblEtiket.Caption = m_Etiket
    
    If blnClearGeneric Then cboGeneriek.Value = vbNullString
    cboVorm.Value = vbNullString
    
    txtDeelDose.Value = vbNullString
    cboDosisEenheid.Clear
    cboDosisEenheid.Value = vbNullString
    
    txtSterkte.Text = vbNullString
    cboSterkteEenheid.Text = vbNullString
    
    FillCombo cboRoute, m_Formularium.GetRoutes()
    
    cboIndicatie.Clear
    cboIndicatie.Value = vbNullString
    
    cboFreq.Value = vbNullString
    txtNormDose.Value = vbNullString
    txtMinDose.Value = vbNullString
    txtMaxDose.Value = vbNullString
    txtAbsMax.Value = vbNullString
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

Private Sub cmdCalc_Click()

    Dim dblDose As Double
    Dim dblWght As Double
    Dim dblVal As Double
    Dim dblCalc As Double
    Dim dblFact As Double
    
    dblFact = m_Freq.Item(cboFreq.Text)
    dblDose = IIf(txtNormDose.Value = vbNullString, 0, txtNormDose.Value)
    dblWght = ModPatient.GetGewichtFromRange()
    
    If dblFact = 0 Or dblDose = 0 Or dblWght = 0 Then Exit Sub
    
    dblVal = dblDose * dblWght / dblFact
    dblVal = ModExcel.Excel_RoundBy(dblVal, txtDeelDose.Value)
    
    dblCalc = dblVal * dblFact / dblWght
    
    m_Med.CalcDose = dblCalc
    m_Med.Freq = cboFreq.Text
    
    txtCalcDose.Value = ModString.FixPrecision(dblCalc, 2)
    lblCalcDose.Caption = "Berekende dosering met deelbaarheid: " & txtDeelDose.Value & " " & cboDosisEenheid.Text
    lblDoseUnit.Caption = cboDoseUnit.Value

End Sub

Private Sub cmdFormularium_Click()
    Dim strUrl As String
    
    strUrl = "https://www.kinderformularium.nl/"
    If Not m_Med.ATC = vbNullString Then
        strUrl = strUrl & "geneesmiddelen?atc_code=" + m_Med.ATC
    Else
        strUrl = strUrl & "geneesmiddelen?name=" + cboGeneriek.Text
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

Private Sub cmdOK_Click()

    If Not m_IsGPK Then
    
        Set m_Med = New ClassMedicatieDisc
        
        m_Med.DoseEenheid = cboDosisEenheid.Value
        m_Med.Generiek = cboGeneriek.Value
        m_Med.Sterkte = StringToDouble(txtSterkte.Value)
        m_Med.SterkteEenheid = cboSterkteEenheid.Value
        m_Med.Vorm = cboVorm.Value
        
    End If
    
    m_Med.DeelDose = StringToDouble(txtDeelDose.Value)
    m_Med.Route = cboRoute.Value
    m_Med.Indicatie = cboIndicatie.Value
    m_Med.Freq = cboFreq.Value
    
    m_Med.NormDose = StringToDouble(txtNormDose.Value)
    m_Med.MinDose = StringToDouble(txtMinDose.Value)
    m_Med.MaxDose = StringToDouble(txtMaxDose.Value)
    m_Med.CalcDose = StringToDouble(txtCalcDose.Value)

    CloseForm "OK"

End Sub

Private Sub txtAbsMax_AfterUpdate()
    
    TextBoxStringNumericValue txtAbsMax

End Sub

Private Sub txtAbsMax_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    ModUtils.OnlyNumericAscii KeyAscii

End Sub

Private Sub txtDeelDose_AfterUpdate()

    TextBoxStringNumericValue txtDeelDose

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

    Validate vbNullString

End Sub

Private Sub txtNormDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

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

Private Sub UserForm_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    m_LoadGPK = False
    
    m_TherapieGroep = lblTherapieGroep.Caption
    m_SubGroep = lblSubGroep.Caption
    m_Etiket = lblEtiket.Caption
    
    Set m_Formularium = ModFormularium.Formularium_GetFormularium
    intC = m_Formularium.MedicamentCount
    
    For intN = 1 To intC
        cboGeneriek.AddItem m_Formularium.Item(intN).Generiek
    Next intN
    
    LoadFreq
    
    cboGeneriek.TabIndex = 0
    cboVorm.TabIndex = 1
    txtSterkte.TabIndex = 2
    cboSterkteEenheid.TabIndex = 3
    txtDeelDose.TabIndex = 4
    cboDosisEenheid.TabIndex = 5
    cboRoute.TabIndex = 6
    cboIndicatie.TabIndex = 7
    
    cboFreq.TabIndex = 8
    txtNormDose.TabIndex = 9
    cboDoseUnit.TabIndex = 10
    txtMinDose.TabIndex = 11
    txtMaxDose.TabIndex = 12
    txtAbsMax.TabIndex = 13
    
    cmdCalc.TabIndex = 14
    cmdFormularium.TabIndex = 15
    cmdOK.TabIndex = 16
    cmdClear.TabIndex = 17
    cmdCancel.TabIndex = 18
       
End Sub

Private Sub UserForm_QueryClose(intCancel As Integer, intMode As Integer)
    
    intCancel = True
    cmdCancel_Click

End Sub


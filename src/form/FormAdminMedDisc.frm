VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAdminMedDisc 
   Caption         =   "Medicament Configuratie"
   ClientHeight    =   12330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20115
   OleObjectBlob   =   "FormAdminMedDisc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAdminMedDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Med As ClassMedDiscConfig
Private m_TherapieGroep As String
Private m_SubGroep As String
Private m_Etiket As String
Private m_Product As String

Private m_IsGPK As Boolean
Private m_LoadGPK As Boolean

Private m_Freq As Dictionary

Public Sub SetNoFormMed()

    m_IsGPK = False

End Sub

Private Function IsNeoAbsMaxInvalid() As Boolean
    
    IsNeoAbsMaxInvalid = txtAbsMax.Value = vbNullString And ModPatient.Patient_GetWeight() > 50 And txtNeoNormDose.Value = vbNullString And txtNeoMaxDose.Value = vbNullString

End Function

Private Function IsPedAbsMaxInvalid() As Boolean
    
    IsPedAbsMaxInvalid = txtAbsMax.Value = vbNullString And ModPatient.Patient_GetWeight() > 50 And txtPedNormDose.Value = vbNullString And txtPedMaxDose.Value = vbNullString

End Function

Private Function IsNeoDoseControlInValid() As Boolean

    IsNeoDoseControlInValid = txtNeoNormDose.Value = vbNullString And txtNeoMaxDose.Value = vbNullString

End Function

Private Function IsPedDoseControlInValid() As Boolean

    IsPedDoseControlInValid = txtPedNormDose.Value = vbNullString And txtPedMaxDose.Value = vbNullString

End Function

Private Sub Validate(ByVal strValid As String)
    
    Dim strNeoValid As String
    Dim strPedValid As String
    
    If strValid = vbNullString Then
    
        strNeoValid = IIf(IsNeoDoseControlInValid, "Voer of een norm dosering in en/of een max (en evt. min en abs max) dosering", strValid)
        strNeoValid = IIf(IsNeoAbsMaxInvalid, "Gewicht boven de 50 kg, voer een absolute maximum dosering in (of een norm dosering of max dosering)", strValid)
    
        strPedValid = IIf(IsPedDoseControlInValid, "Voer of een norm dosering in en/of een max (en evt. min en abs max) dosering", strValid)
        strPedValid = IIf(IsPedAbsMaxInvalid, "Gewicht boven de 50 kg, voer een absolute maximum dosering in (of een norm dosering of max dosering)", strValid)
          
        strValid = IIf(cboDosisEenheid.Value = vbNullString, "Voer dosering eenheid in", strValid)
        strValid = IIf(txtDeelDose.Value = vbNullString, "Voer een deelbaarheid in", strValid)
        strValid = IIf(cboSterkteEenheid.Value = vbNullString, "Voer sterkte eenheid in", strValid)
        strValid = IIf(txtSterkte.Value = vbNullString, "Voer sterkte in", strValid)
        
        strValid = IIf(cboVorm.Value = vbNullString, "Voer een vorm in", strValid)
        strValid = IIf(cboGeneriek.Value = vbNullString, "Kies een generiek", strValid)
    
    End If
    
    lblNeoValid.Caption = strNeoValid
    lblPedValid.Caption = strPedValid
    lblValid.Caption = strValid
    
    cmdOK.Enabled = strValid = vbNullString And strNeoValid = vbNullString And strPedValid = vbNullString

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
    Me.lbxRoute.Enabled = Not blnIsGPK
    Me.lbxIndicatie.Enabled = Not blnIsGPK
    
    cmdFormularium.Enabled = blnIsGPK
    
    If Not blnIsGPK Then
        lblGPK.Caption = vbNullString
        lblATC.Caption = vbNullString
        
        FillCombo cboVorm, Formularium_GetFormConfig().GetVormen()
        FillCombo cboSterkteEenheid, Formularium_GetFormConfig().GetSterkteEenheden()
        FillCombo cboDosisEenheid, Formularium_GetFormConfig.GetDosisEenheden()
        FillListBox lbxRoute, Formularium_GetFormConfig.GetRoutes()
        
        lbxIndicatie.Clear
        
        LoadFreq
        
    End If

End Sub

Private Sub cboDosisEenheid_Change()

    SetDoseUnit

End Sub

Private Sub cboFreq_Change()

    Validate vbNullString

End Sub

Private Sub cboGeneriek_Change()
    
    If m_LoadGPK Then Exit Sub

    If cboGeneriek.ListIndex > -1 Then
        SetToGPKMode True
        Set m_Med = Formularium_GetFormConfig.Item(cboGeneriek.ListIndex + 1)
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
    
    Set m_Med = Formularium_GetFormConfig.GPK(strGPK)
    
    If m_Med Is Nothing Then
        SetToGPKMode False
        blnLoad = False
    Else
        SetToGPKMode True
        LoadMedicament
        m_LoadGPK = True
        cboGeneriek.Text = m_Med.Generic
        m_LoadGPK = False
    End If
    
    LoadGPK = blnLoad

End Function

Private Sub SetDoseUnit()

    Dim strUnit As String
    
    strUnit = Trim(cboDosisEenheid.Text)
    
    If Not strUnit = vbNullString Then
        cboNeoDoseUnit.Text = strUnit & "/kg/dag"
        lblNeoMinDoseUnit.Caption = cboNeoDoseUnit.Value
        lblNeoMaxDoseUnit.Caption = cboNeoDoseUnit.Value
        
        cboPedDoseUnit.Text = strUnit & "/kg/dag"
        lblPedMinDoseUnit.Caption = cboPedDoseUnit.Value
        lblPedMaxDoseUnit.Caption = cboPedDoseUnit.Value
        
        cboAbsMaxUnit.Value = strUnit & "/dag"
        
        lblNeoConcUnit.Caption = strUnit & "/ml"
        lblPedConcUnit.Caption = strUnit & "/ml"
    End If

End Sub

Private Sub LoadMedicament()
    
    Dim arrFreq() As String
    Dim varFreq As Variant
    Dim intN As Integer
    Dim intC As Integer

    With m_Med
    
        lblTherapieGroep.Caption = .MainGroup
        lblSubGroep.Caption = .SubGroup
        lblEtiket.Caption = .Label
        lblProduct.Caption = .Product
        
        lblGPK.Caption = .GPK
        lblATC.Caption = .ATC
        
        txtSynon.Value = cboGeneriek.Value
        cboVorm.Value = .Shape
        
        txtSterkte.Text = .GenericQuantity
        cboSterkteEenheid.Text = .GenericUnit
        
        txtDeelDose.Text = .MultipleQuantity
        
        FillCombo cboDosisEenheid, GetDosisEenheden()
        cboDosisEenheid.Text = .MultipleUnit
        
        FillListBox lbxRoute, .GetRouteList()
        FillListBox lbxIndicatie, .GetIndicationList()
        
        LoadFreq
        If Not .GetFreqListString = vbNullString Then
            intC = lbxFreq.ListCount
            For Each varFreq In .GetFreqList()
                For intN = 0 To intC - 1
                    If lbxFreq.List(intN) = varFreq Then lbxFreq.Selected(intN) = True
                Next
            Next
        End If
        
        SetTextBoxNumericValue txtNeoNormDose, .NeoNormDose
        SetTextBoxNumericValue txtNeoMinDose, .NeoMinDose
        SetTextBoxNumericValue txtNeoMaxDose, .NeoMaxDose
        
        SetTextBoxNumericValue txtPedNormDose, .PedNormDose
        SetTextBoxNumericValue txtPedMinDose, .PedMinDose
        SetTextBoxNumericValue txtPedMaxDose, .PedMaxDose
        
        SetTextBoxNumericValue txtAbsMax, .PedAbsMaxDose
        
        cboPedOplVlst.Value = .PedSolution
        SetTextBoxNumericValue txtPedConc, .PedMaxConc
        SetTextBoxNumericValue txtPedVol, .PedSolutionVolume
        SetTextBoxNumericValue txtPedTijd, .PedMinInfusionTime
        
    End With

End Sub

Private Sub FillListBox(objLbx As MSForms.ListBox, ByVal objCol As Collection)

    Dim objItem As Variant
    
    objLbx.Clear
    For Each objItem In objCol
        objLbx.AddItem CStr(objItem)
    Next

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

    Set GetDosisEenheden = Formularium_GetFormConfig.GetDosisEenheden()

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
    lblProduct.Caption = m_Product
    
    If blnClearGeneric Then
        cboGeneriek.Value = vbNullString
    End If
    txtSynon.Value = vbNullString
    
    cboVorm.Value = vbNullString
    
    txtDeelDose.Value = vbNullString
    cboDosisEenheid.Clear
    cboDosisEenheid.Value = vbNullString
    
    txtSterkte.Text = vbNullString
    cboSterkteEenheid.Text = vbNullString
    
    FillListBox lbxRoute, Formularium_GetFormConfig.GetRoutes()
    
    lbxIndicatie.Clear
    
    txtNeoNormDose.Value = vbNullString
    txtNeoMinDose.Value = vbNullString
    txtNeoMaxDose.Value = vbNullString
    cboNeoDoseUnit.Value = vbNullString
    
    txtPedNormDose.Value = vbNullString
    txtPedMinDose.Value = vbNullString
    txtPedMaxDose.Value = vbNullString
    cboPedDoseUnit.Value = vbNullString
    
    txtAbsMax.Value = vbNullString
    
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
    
    ' strValid = ValidateCombo(cboRoute)
    Validate strValid

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

    Dim objConfig As ClassMedDiscConfig
    
    Dim intN As Integer
    Dim intC As Integer
    Dim strMsg As String
    Dim vbAnswer As Integer
    
    If Not Formularium_IsInitialized Then
        ModMessage.ShowMsgBoxInfo "Open eerst een keer een medicament om het formularium te laden, probeer het daarna opnieuw"
        Exit Sub
    End If
        
    If Not cboGeneriek.Value = txtSynon.Value Then
        vbAnswer = vbNo
    Else
        vbAnswer = ModMessage.ShowMsgBoxYesNo("Instellingen doorvoeren voor alle medicamenten met dezelfde generiek/route?")
    End If
    
    intC = ModFormularium.Formularium_GetFormConfig.MedicamentCount
    For intN = 1 To intC
        If vbAnswer = vbYes Then
            Set objConfig = Formularium_GetFormConfig.Item(intN)
            If Matches(objConfig) Then
                SetConfig objConfig
            End If
        Else
            Set objConfig = Formularium_GetFormConfig.GPK(m_Med.GPK)
            SetConfig objConfig
            cboGeneriek.List(cboGeneriek.ListIndex) = objConfig.Generic
        End If
    Next
    
    strMsg = "Configuratie toegepast voor:" & vbNewLine
    If vbAnswer = vbYes Then
        strMsg = strMsg & cboGeneriek.Value & " "
        strMsg = strMsg & Replace(ListBoxToString(lbxRoute, True), "||", ", ")
    Else
        strMsg = strMsg & m_Med.Product
    End If
    
    ModMessage.ShowMsgBoxInfo strMsg

End Sub

Private Sub SetConfig(objConfig As ClassMedDiscConfig)

    Dim objMed As ClassMedicatieDisc
    Dim blnIsPed As Boolean

    blnIsPed = MetaVision_IsPediatrie()
    
    objConfig.Generic = txtSynon.Value
    txtSynon.Value = objConfig.Generic
    
    objConfig.MultipleQuantity = StringToDouble(txtDeelDose.Value)
    objConfig.MultipleUnit = cboDosisEenheid.Value

    objConfig.SetFreqList ListBoxToString(lbxFreq, False)
    objConfig.NeoNormDose = StringToDouble(txtNeoNormDose.Value)
    objConfig.NeoMinDose = StringToDouble(txtNeoMinDose.Value)
    objConfig.NeoMaxDose = StringToDouble(txtNeoMaxDose.Value)

    objConfig.PedNormDose = StringToDouble(txtPedNormDose.Value)
    objConfig.PedMinDose = StringToDouble(txtPedMinDose.Value)
    objConfig.PedMaxDose = StringToDouble(txtPedMaxDose.Value)
    
    objConfig.PedAbsMaxDose = StringToDouble(txtAbsMax.Value)
    
    objConfig.PedSolution = cboPedOplVlst.Value
    objConfig.PedMaxConc = StringToDouble(txtPedConc.Value)
    objConfig.PedSolutionVolume = StringToDouble(txtPedVol.Value)
    objConfig.PedMinInfusionTime = StringToDouble(txtPedTijd.Value)
    
    Set objMed = Formularium_GetFormularium.GPK(objConfig.GPK)
    
    objMed.Generic = objConfig.Generic
    objMed.MultipleQuantity = objConfig.MultipleQuantity
    objMed.MultipleUnit = objConfig.MultipleUnit

    objMed.SetFreqList objConfig.GetFreqListString
    objMed.NormDose = IIf(blnIsPed, objConfig.PedNormDose, objConfig.NeoNormDose)
    objMed.MinDose = IIf(blnIsPed, objConfig.PedMinDose, objConfig.NeoMinDose)
    objMed.MaxDose = IIf(blnIsPed, objConfig.PedMaxDose, objConfig.NeoMaxDose)
            
    objMed.AbsMaxDose = objConfig.PedAbsMaxDose
    
    objMed.Solution = objConfig.PedSolution
    objMed.MaxConc = objConfig.PedMaxConc
    objMed.MinInfusionTime = objConfig.PedMinInfusionTime

End Sub

Private Function ListBoxToString(objLbx As MSForms.ListBox, ByVal blnAll As Boolean) As String

    Dim strText As String
    Dim strDel As String
    Dim varItem As Variant
    Dim intN As Integer
    Dim intC As Integer
    
    strDel = "||"
    intC = objLbx.ListCount - 1
    For intN = 0 To intC
        If objLbx.Selected(intN) Or blnAll Then
            varItem = objLbx.List(intN)
            strText = IIf(strText = vbNullString, varItem, strText & strDel & varItem)
        End If
    Next
    
    ListBoxToString = strText

End Function

Private Function Matches(objMed As ClassMedDiscConfig) As Boolean

    Dim blnMatch As Boolean
    Dim blnRoute As Boolean
    Dim varRoute As Variant
    Dim intN As Integer
    Dim intC As Integer
    
    blnMatch = objMed.Generic = cboGeneriek.Value
    
    If blnMatch Then
        For Each varRoute In objMed.GetRouteList
            blnMatch = blnMatch And ListBoxContains(lbxRoute, varRoute)
            If Not blnMatch Then Exit For
        Next
    End If
    
    Matches = blnMatch

End Function

Private Function ListBoxContains(objLbx As MSForms.ListBox, ByVal strItem As String) As Boolean
    
    Dim blnContains As Boolean
    Dim intN As Integer
    Dim intC As Integer
    
    blnContains = False
    intC = objLbx.ListCount - 1
    For intN = 0 To intC
        If objLbx.List(intN) = strItem Then
            blnContains = True
            Exit For
        End If
    Next
    
    ListBoxContains = blnContains

End Function

Private Sub cmdSave_Click()

    Me.Hide
    
    ModProgress.StartProgress "Medicatie configuratie opslaan"
    ModFormularium.Formularium_ExportMedDiscConfig (True)
    ModProgress.FinishProgress

End Sub

Private Sub cboVorm_Change()

    Dim strValid As String
    
    strValid = ValidateCombo(cboVorm)
    Validate strValid

End Sub

Private Sub txtDeelDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtDeelDose_AfterUpdate()

    TextBoxStringNumericValue txtDeelDose

End Sub

Private Sub txtDeelDose_Change()

    Validate vbNullString

End Sub

Private Sub txtNeoNormDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtNeoNormDose_AfterUpdate()

    TextBoxStringNumericValue txtNeoNormDose
    
End Sub

Private Sub txtNeoNormDose_Change()

    Validate vbNullString
    
End Sub

Private Sub txtNeoMinDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtNeoMinDose_AfterUpdate()

    TextBoxStringNumericValue txtNeoMinDose
    
End Sub

Private Sub txtNeoMinDose_Change()

    Validate vbNullString
    
End Sub

Private Sub txtNeoMaxDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtNeoMaxDose_AfterUpdate()

    TextBoxStringNumericValue txtNeoMaxDose
    
End Sub

Private Sub txtNeoMaxDose_Change()

    Validate vbNullString
    
End Sub

Private Sub txtPedNormDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtPedNormDose_AfterUpdate()

    TextBoxStringNumericValue txtPedNormDose
    
End Sub

Private Sub txtPedNormDose_Change()

    Validate vbNullString
    
End Sub

Private Sub txtPedMinDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtPedMinDose_AfterUpdate()

    TextBoxStringNumericValue txtPedMinDose
    
End Sub

Private Sub txtPedMinDose_Change()

    Validate vbNullString
    
End Sub

Private Sub txtPedMaxDose_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtPedMaxDose_AfterUpdate()

    TextBoxStringNumericValue txtPedMaxDose
    
End Sub

Private Sub txtPedMaxDose_Change()

    Validate vbNullString
    
End Sub

Private Sub txtAbsMax_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtAbsMax_AfterUpdate()

    TextBoxStringNumericValue txtAbsMax
    
End Sub

Private Sub txtAbsMax_Change()

    Validate vbNullString
    
End Sub

Private Sub txtPedConc_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtPedConc_AfterUpdate()

    TextBoxStringNumericValue txtPedConc
    
End Sub

Private Sub txtPedConc_Change()

    Validate vbNullString
    
End Sub

Private Sub txtPedTijd_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ModUtils.CorrectNumberAscii KeyAscii

End Sub

Private Sub txtPedTijd_AfterUpdate()

    TextBoxStringNumericValue txtPedTijd
    
End Sub

Private Sub txtPedTijd_Change()

    Validate vbNullString
    
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

    lbxFreq.Clear

    If m_Freq Is Nothing Then
        Set m_Freq = ModMedDisc.GetMedicationFreqs()
    End If

    For Each varKey In m_Freq.Keys
        lbxFreq.AddItem varKey
    Next

End Sub

Private Sub UserForm_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String
    
    If Not Formularium_IsInitialized Then Formularium_Initialize

    m_LoadGPK = False
    
    m_TherapieGroep = lblTherapieGroep.Caption
    m_SubGroep = lblSubGroep.Caption
    m_Etiket = lblEtiket.Caption
    m_Product = lblProduct.Caption
    
    intC = Formularium_GetFormConfig.MedicamentCount + 1
    
    For intN = 1 To intC
        cboGeneriek.AddItem Formularium_GetFormConfig.Item(intN).Generic
    Next intN
       
    SetTabOrder2 ' GetTabControls()
    
    FillCombo cboPedOplVlst, ModMedDisc.MedDisc_GetOplVlstCol
       
End Sub

Private Sub UserForm_QueryClose(intCancel As Integer, intMode As Integer)
    
    intCancel = True
    cmdCancel_Click

End Sub

Private Sub SetTabOrder2()

    frmDetails.TabIndex = 0
    cboGeneriek.TabIndex = 0
    txtSynon.TabIndex = 1
    txtDeelDose.TabIndex = 2
    cboDosisEenheid.TabIndex = 3
    
    frmFreq.TabIndex = 1
    lbxFreq.TabIndex = 0
    
    frmDoseNeo.TabIndex = 2
    txtNeoNormDose.TabIndex = 0
    txtNeoMinDose.TabIndex = 1
    txtNeoMaxDose.TabIndex = 2
    
    frmPedDose.TabIndex = 3
    txtPedNormDose.TabIndex = 0
    txtPedMinDose.TabIndex = 1
    txtPedMaxDose.TabIndex = 2
    
    frmAbsMax.TabIndex = 4
    txtAbsMax.TabIndex = 0
    
    frmPedOplossing.TabIndex = 5
    cboPedOplVlst.TabIndex = 0
    txtPedConc.TabIndex = 1
    txtPedTijd.TabIndex = 2
    
    cmdFormularium.TabIndex = 6
    cmdOK.TabIndex = 7
    cmdSave.TabIndex = 8
    cmdCancel.TabIndex = 9

End Sub


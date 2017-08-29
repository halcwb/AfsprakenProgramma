VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedDisc 
   Caption         =   "Kies een medicament ..."
   ClientHeight    =   7485
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

Private m_Medicament As ClassMedicatieDisc
Private m_Formularium As ClassFormularium
Private m_TherapieGroep As String
Private m_SubGroep As String
Private m_Etiket As String

Private m_IsGPK As Boolean
Private m_LoadGPK As Boolean

Public Sub SetNoFormMed()

    m_IsGPK = False

End Sub

Private Sub Validate(ByVal strValid As String)

    
    If strValid = vbNullString Then
    
        strValid = IIf(cboIndicatie.Value = vbNullString, "Kies een indicatie", vbNullString)
        strValid = IIf(cboRoute.Value = vbNullString, "Kies een route", strValid)
        
        strValid = IIf(cboDosisEenheid.Value = vbNullString, "Voer dosering grootte in", strValid)
        strValid = IIf(txtDosis.Value = vbNullString, "Voer dosering eenheid in", strValid)
        strValid = IIf(cboSterkteEenheid.Value = vbNullString, "Voer sterkte eenheid in", strValid)
        strValid = IIf(txtSterkte.Value = vbNullString, "Voer sterkte in", strValid)
        
        strValid = IIf(cboVorm.Value = vbNullString, "Voer een vorm in", strValid)
        strValid = IIf(cboGeneriek.Value = vbNullString, "Kies een generiek", strValid)
    
    End If
    
    lblValid.Caption = strValid
    cmdOK.Enabled = strValid = vbNullString

End Sub

Public Function GetSelectedDosisEenheid() As String

    GetSelectedDosisEenheid = cboDosisEenheid.Value

End Function

Public Function GetSelectedRoute() As String

    GetSelectedRoute = cboRoute.Value

End Function

Public Function GetSelectedIndication() As String

    GetSelectedIndication = cboIndicatie.Value

End Function

Public Function HasSelectedMedicament() As Boolean

    HasSelectedMedicament = Not m_Medicament Is Nothing

End Function

Public Function GetSelectedMedicament() As ClassMedicatieDisc

    Set GetSelectedMedicament = m_Medicament

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
        
        cboVorm.Clear
        For Each varItem In m_Formularium.GetVormen
            cboVorm.AddItem varItem
        Next
        
        cboSterkteEenheid.Clear
        For Each varItem In m_Formularium.GetSterkteEenheden()
            cboSterkteEenheid.AddItem varItem
        Next
    
        cboDosisEenheid.Clear
        For Each varItem In m_Formularium.GetDosisEenheden
            cboDosisEenheid.AddItem varItem
        Next
        
        cboRoute.Clear
        For Each varItem In m_Formularium.GetRoutes
            cboRoute.AddItem varItem
        Next
        
        cboIndicatie.Clear
    End If

End Sub

Private Sub cboGeneriek_Change()
    
    If m_LoadGPK Then Exit Sub

    If cboGeneriek.ListIndex > -1 Then
        SetToGPKMode True
        Set m_Medicament = m_Formularium.Item(cboGeneriek.ListIndex + 1)
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
    If Not m_Medicament Is Nothing Then strGPK = m_Medicament.GPK
    
    GetGPK = strGPK

End Function

Public Function LoadGPK(ByVal strGPK As String) As Boolean
    
    Dim blnLoad As Boolean

    blnLoad = True
    
    Set m_Medicament = m_Formularium.GPK(strGPK)
    
    If m_Medicament Is Nothing Then
        SetToGPKMode False
        blnLoad = False
    Else
        SetToGPKMode True
        LoadMedicament
        m_LoadGPK = True
        cboGeneriek.Text = m_Medicament.Generiek
        m_LoadGPK = False
    End If
    
    LoadGPK = blnLoad

End Function

Private Sub LoadMedicament()

    With m_Medicament
    
        lblTherapieGroep.Caption = .TherapieGroep
        lblSubGroep.Caption = .TherapieSubgroep
        lblEtiket.Caption = .Etiket
        
        lblGPK.Caption = .GPK
        lblATC.Caption = .ATC
        
        cboVorm.Value = .Vorm
        
        txtSterkte.Text = .Sterkte
        cboSterkteEenheid.Text = .SterkteEenheid
        
        txtDosis.Text = .Dosis
        
        FillCombo cboDosisEenheid, GetDosisEenheden()
        cboDosisEenheid.Text = .DosisEenheid
        
        FillCombo cboRoute, .GetRoutes()
        FillCombo cboIndicatie, .GetIndicaties()
        
    End With

End Sub

Private Function GetDosisEenheden() As String()

    Dim arrEenheden() As String
    
    ModArray.StringArrayAddAllFromCol m_Formularium.GetDosisEenheden, arrEenheden
        
    GetDosisEenheden = arrEenheden

End Function

Private Sub FillCombo(objCombo As ComboBox, arrItems() As String)

    Dim varItem As Variant

    objCombo.Clear
    
    For Each varItem In arrItems
        objCombo.AddItem CStr(varItem)
    Next varItem
    
    If UBound(arrItems) = 0 Then objCombo.Text = arrItems(0)
    
End Sub

Public Sub ClearForm(ByVal blnClearGeneric As Boolean)

    lblTherapieGroep.Caption = m_TherapieGroep
    lblSubGroep.Caption = m_SubGroep
    lblEtiket.Caption = m_Etiket
    
    If blnClearGeneric Then cboGeneriek.Value = vbNullString
    cboVorm.Value = vbNullString
    
    txtDosis.Value = vbNullString
    cboDosisEenheid.Clear
    cboDosisEenheid.Value = vbNullString
    
    txtSterkte.Text = vbNullString
    cboSterkteEenheid.Text = vbNullString
    
    cboRoute.Clear
    cboRoute.Value = vbNullString
    cboIndicatie.Clear
    cboIndicatie.Value = vbNullString
    
    cboGeneriek.SetFocus
    
    Set m_Medicament = Nothing

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

Private Sub cmdFormularium_Click()
    Dim strUrl As String
    strUrl = "https://www.kinderformularium.nl/geneesmiddelen?name=" + cboGeneriek.Text

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
    
        Set m_Medicament = New ClassMedicatieDisc
        
        m_Medicament.Dosis = StringToDouble(txtDosis.Value)
        m_Medicament.DosisEenheid = cboDosisEenheid.Value
        m_Medicament.Generiek = cboGeneriek.Value
        m_Medicament.Indicaties = cboIndicatie.Value
        m_Medicament.Routes = cboRoute.Value
        m_Medicament.Sterkte = StringToDouble(txtSterkte.Value)
        m_Medicament.SterkteEenheid = cboSterkteEenheid.Value
        m_Medicament.Vorm = cboVorm.Value
        
    Else
        
        m_Medicament.Dosis = StringToDouble(txtDosis.Value)
        m_Medicament.Routes = cboRoute.Value
        m_Medicament.Indicaties = cboIndicatie.Value
    
    End If

    CloseForm "OK"

End Sub

Private Sub txtAfronding_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii = 46 Or KeyAscii = 44 Then
                KeyAscii = 44
            Else
                'Bij elke andere waarde negeren
                KeyAscii = 0
                Beep
            End If
    End If
        

End Sub

Private Sub txtDosis_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate vbNullString

End Sub

Private Sub cboVorm_Change()

    Dim strValid As String
    
    strValid = ValidateCombo(cboVorm)
    Validate strValid

End Sub

Private Sub txtSterkte_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate vbNullString

End Sub

Private Sub txtSterkte_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii = 46 Or KeyAscii = 44 Then
                KeyAscii = 44
            Else
                'Bij elke andere waarde negeren
                KeyAscii = 0
                Beep
            End If
    End If

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

Private Sub UserForm_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    strTitle = "Formularium wordt geladen, een ogenblik geduld a.u.b. ..."
    
    ModProgress.StartProgress strTitle
    
    m_LoadGPK = False
    
    m_TherapieGroep = lblTherapieGroep.Caption
    m_SubGroep = lblSubGroep.Caption
    m_Etiket = lblEtiket.Caption
    
    Set m_Formularium = New ClassFormularium
    m_Formularium.GetMedicamenten (True)
    
    intC = m_Formularium.MedicamentCount
    For intN = 1 To intC
        cboGeneriek.AddItem m_Formularium.Item(intN).Generiek
        
        ModProgress.SetJobPercentage "Generieken toevoegen", intC, intN
    Next intN
    
    cboGeneriek.TabIndex = 0
    cboVorm.TabIndex = 1
    txtSterkte.TabIndex = 2
    cboSterkteEenheid.TabIndex = 3
    txtDosis.TabIndex = 4
    cboDosisEenheid.TabIndex = 5
    cboRoute.TabIndex = 6
    cboIndicatie.TabIndex = 7
    
    cmdFormularium.TabIndex = 8
    cmdOK.TabIndex = 9
    cmdClear.TabIndex = 10
    cmdCancel.TabIndex = 11
       
    ModProgress.FinishProgress

End Sub

Private Sub UserForm_QueryClose(intCancel As Integer, intMode As Integer)
    
    intCancel = True
    cmdCancel_Click

End Sub

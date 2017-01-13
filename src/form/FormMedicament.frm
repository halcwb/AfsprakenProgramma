VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedicament 
   Caption         =   "Kies een medicament ..."
   ClientHeight    =   5446
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   OleObjectBlob   =   "FormMedicament.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMedicament"
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

Private blnNoFormMed As Boolean

Public Sub SetNoFormMed()

    blnNoFormMed = True

End Sub

Private Sub Validate()

    Dim strValid As String
    
    strValid = IIf(cboIndicatie.Value = vbNullString, "Kies een indicatie", vbNullString)
    strValid = IIf(cboRoute.Value = vbNullString, "Kies een route", strValid)
    
    strValid = IIf(txtDosisEenheid.Value = vbNullString, "Voer dosering grootte in", strValid)
    strValid = IIf(txtDosis.Value = vbNullString, "Voer dosering eenheid in", strValid)
    strValid = IIf(txtSterkteEenheid.Value = vbNullString, "Voer sterkte eenheid in", strValid)
    strValid = IIf(txtSterkte.Value = vbNullString, "Voer sterkte in", strValid)
    
    strValid = IIf(txtShape.Value = vbNullString, "Voer een vorm in", strValid)
    strValid = IIf(cboGeneriek.Value = vbNullString, "Kies een generiek", strValid)
    
    lblValid.Caption = strValid
    cmdOK.Enabled = strValid = vbNullString

End Sub

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

Private Sub cboGeneriek_Change()

    If cboGeneriek.ListIndex > -1 Then
        blnNoFormMed = False
        Set m_Medicament = m_Formularium.Item(cboGeneriek.ListIndex + 1)
        LoadMedicament
    Else
        If Not blnNoFormMed Then ' Only clear form once when no form med is discovered
            blnNoFormMed = True
            ClearForm False
        End If
    End If
    
    Validate

End Sub

Public Function GetGPK() As String
    Dim strGPK As String
    
    strGPK = "0"
    If Not m_Medicament Is Nothing Then strGPK = m_Medicament.GPK
    
    GetGPK = strGPK

End Function

Public Sub LoadGPK(ByVal strGPK As String)

    Set m_Medicament = m_Formularium.GPK(strGPK)
    LoadMedicament
    cboGeneriek.Value = m_Medicament.Generiek

End Sub

Private Sub LoadMedicament()

    With m_Medicament
        lblTherapieGroep.Caption = .TherapieGroep
        lblSubGroep.Caption = .TherapieSubgroep
        lblEtiket.Caption = .Etiket
        
        txtShape.Value = .Vorm
        
        txtSterkte.Text = .Sterkte
        txtSterkteEenheid.Text = .SterkteEenheid
        
        txtDosis.Text = .Dosis
        txtDosisEenheid.Text = .DosisEenheid
        
        FillCombo cboRoute, .GetRoutes()
        FillCombo cboIndicatie, .GetIndicaties()
    End With

End Sub

Private Sub FillCombo(ByRef objCombo As ComboBox, ByRef arrItems() As String)

    Dim varItem As Variant

    objCombo.Clear
    
    For Each varItem In arrItems
        objCombo.AddItem CStr(varItem)
    Next varItem
    
    If UBound(arrItems) = 0 Then objCombo.Text = arrItems(0)
    
End Sub

Public Sub ClearForm(ByVal blnGeneric As Boolean)

    lblTherapieGroep.Caption = m_TherapieGroep
    lblSubGroep.Caption = m_SubGroep
    lblEtiket.Caption = m_Etiket
    
    If blnGeneric Then cboGeneriek.Value = vbNullString
    txtShape.Value = vbNullString
    
    txtDosis.Value = vbNullString
    txtDosisEenheid.Value = vbNullString
    
    txtSterkte.Text = vbNullString
    txtSterkteEenheid.Text = vbNullString
    
    cboRoute.Clear
    cboRoute.Value = vbNullString
    cboIndicatie.Clear
    cboIndicatie.Value = vbNullString
    
    Set m_Medicament = Nothing

End Sub

Private Sub cboIndicatie_Change()

    Validate

End Sub

Private Sub cboRoute_Change()

    Validate

End Sub

Private Sub cmdFormularium_Click()
    Dim strUrl As String
    strUrl = "https://www.kinderformularium.nl/geneesmiddelen?name=" + cboGeneriek.Text

    ActiveWorkbook.FollowHyperlink strUrl

End Sub

Private Sub CloseForm(ByVal strAction As String)

    lblButton.Caption = strAction
    blnNoFormMed = False
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

    If blnNoFormMed Then
        Set m_Medicament = New ClassMedicatieDisc
        
        m_Medicament.Dosis = Val(txtDosis.Value)
        m_Medicament.DosisEenheid = txtDosisEenheid.Value
        m_Medicament.Generiek = cboGeneriek.Value
        m_Medicament.Indicaties = cboIndicatie.Value
        m_Medicament.Routes = cboRoute.Value
        m_Medicament.Sterkte = Val(txtSterkte.Value)
        m_Medicament.SterkteEenheid = txtSterkteEenheid.Value
        m_Medicament.Vorm = txtShape.Value
        
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

    Validate

End Sub

Private Sub txtShape_Change()

    Validate

End Sub

Private Sub txtSterkte_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate

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

Private Sub txtSterkteEenheid_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Validate

End Sub

Private Sub UserForm_Activate()

    Validate

End Sub

Private Sub UserForm_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    strTitle = "Formularium wordt geladen, een ogenblik geduld a.u.b. ..."
    
    ModProgress.StartProgress strTitle
    
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
    txtShape.TabIndex = 1
    txtSterkte.TabIndex = 2
    txtSterkteEenheid.TabIndex = 3
    txtDosis.TabIndex = 4
    txtDosisEenheid.TabIndex = 5
    cboRoute.TabIndex = 6
    cboIndicatie.TabIndex = 7
    
    cmdFormularium.TabIndex = 8
    cmdOK.TabIndex = 9
    cmdClear.TabIndex = 10
    cmdCancel.TabIndex = 11
    
    ModProgress.FinishProgress

End Sub

Private Sub UserForm_QueryClose(intCancel As Integer, intMode As Integer)
    
    ' ModMessage.ShowMsgBoxExclam "Dit formulier kan alleen worden afgesloten met 'OK', 'Clear' of 'Cancel'"
    intCancel = True
    cmdCancel_Click

End Sub

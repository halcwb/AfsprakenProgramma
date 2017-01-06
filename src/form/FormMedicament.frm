VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedicament 
   Caption         =   "Kies een medicament ..."
   ClientHeight    =   5068
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   10647
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
        Set m_Medicament = m_Formularium.Item(cboGeneriek.ListIndex + 1)
        LoadMedicament
    Else
        ClearForm
    End If

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

End Sub

Private Sub LoadMedicament()

    With m_Medicament
        lblTherapieGroep.Caption = .TherapieGroep
        lblSubGroep.Caption = .TherapieSubgroep
        lblEtiket.Caption = .Etiket
        cboGeneriek.Value = .Generiek
        FillCombo cboIndicatie, .GetIndicaties()
        txtSterkte.Text = .Sterkte
        txtSterkteEenheid.Text = .SterkteEenheid
        txtDosis.Text = .Dosis
        txtDosisEenheid.Text = .DosisEenheid
        FillCombo cboRoute, .GetRoutes()
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

Private Sub ClearForm()

    cboGeneriek.Text = vbNullString
    
    lblTherapieGroep.Caption = m_TherapieGroep
    lblSubGroep.Caption = m_SubGroep
    lblEtiket.Caption = m_Etiket
    
    cboGeneriek.Value = vbNullString
    cboIndicatie.Clear
    txtSterkte.Text = vbNullString
    txtSterkteEenheid.Text = vbNullString
    cboRoute.Clear
    chkPRN.Value = False
    txtPRN.Value = vbNullString
    
    Set m_Medicament = Nothing

End Sub

Private Sub chkPRN_Change()

    txtPRN.Visible = chkPRN.Value

End Sub


Private Sub cmdFormularium_Click()
    Dim strUrl As String
    strUrl = "https://www.kinderformularium.nl/geneesmiddelen?name=" + cboGeneriek.Text

    ActiveWorkbook.FollowHyperlink strUrl

End Sub

Private Sub cmdCancel_Click()
    
    lblButton.Caption = "Cancel"
    Me.Hide

End Sub

Private Sub cmdClear_Click()

    ClearForm
    lblButton.Caption = "Clear"
    Me.Hide
    
End Sub

Private Sub cmdOk_Click()

    lblButton.Caption = "OK"
    Me.Hide

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

Private Sub UserForm_Activate()

    If Not HasSelectedMedicament() Then ClearForm

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
    
    ModProgress.FinishProgress

End Sub

Private Sub UserForm_QueryClose(intCancel As Integer, intMode As Integer)
    
    ModMessage.ShowMsgBoxExclam "Dit formulier kan alleen worden afgesloten met 'OK', 'Clear' of 'Cancel'"
    intCancel = True

End Sub

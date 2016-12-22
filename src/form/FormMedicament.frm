VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedicament 
   Caption         =   "Kies een medicament ..."
   ClientHeight    =   6426
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   12397
   OleObjectBlob   =   "FormMedicament.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMedicament"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intN As Integer
Private objMedicament As ClassMedicatieDisc
Private objFormularium As ClassFormularium

Private Sub cboGeneriek_Change()

    If cboGeneriek.ListIndex > -1 Then
        Set objMedicament = objFormularium.Item(cboGeneriek.ListIndex + 1)
        LoadMedicament objMedicament
    Else
        ClearForm
    End If

End Sub

Public Function GetGPK() As String
    Dim strGPK As String
    
    strGPK = "0"
    If Not objMedicament Is Nothing Then strGPK = objMedicament.GPK
    
    GetGPK = strGPK

End Function

Public Sub LoadGPK(strGPK As String)

    Set objMedicament = objFormularium.GPK(strGPK)
    LoadMedicament objMedicament

End Sub

Private Sub LoadMedicament(objMedicament As ClassMedicatieDisc)

        With objMedicament
            lblTherapieGroep.Caption = .TherapieGroep
            lblSubGroep.Caption = .TherapieSubgroep
            lblEtiket.Caption = .Etiket
            FillCombo cboIndicatie, .GetIndicaties()
            txtSterkte.Text = .Sterkte
            txtSterkteEenheid.Text = .SterkteEenheid
            txtDosis.Text = .Dosis
            txtDosisEenheid.Text = .DosisEenheid
            FillCombo cboRoute, .GetRoutes()
        End With

End Sub

Private Sub FillCombo(objCombo As ComboBox, arrItems() As String)
    Dim varItem As Variant

    objCombo.Clear
    
    For Each varItem In arrItems
        objCombo.AddItem CStr(varItem)
    Next varItem
    
    If UBound(arrItems) = 0 Then objCombo.Text = arrItems(0)
    
End Sub

Private Sub ClearForm()

    cboGeneriek.Text = vbNullString
    
    lblTherapieGroep.Caption = vbNullString
    lblSubGroep.Caption = vbNullString
    lblEtiket.Caption = vbNullString
    cboIndicatie.Clear
    txtSterkte.Text = vbNullString
    txtSterkteEenheid.Text = vbNullString
    cboRoute.Clear
    
    Set objMedicament = Nothing

End Sub

Private Sub cmdFormularium_Click()
    Dim strUrl As String
    strUrl = "https://www.kinderformularium.nl/geneesmiddelen?name=" + cboGeneriek.Text

    ActiveWorkbook.FollowHyperlink strUrl

End Sub

Private Sub cmdCancel_Click()
    
    Me.Hide
    lblCancel.Caption = "Cancel"

End Sub

Private Sub cmdClear_Click()

    Me.Hide
    lblCancel.Caption = "Clear"
    
End Sub

Private Sub cmdOk_Click()

    Me.Hide
    lblCancel.Caption = "OK"

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

Private Sub UserForm_Initialize()

    MsgBox prompt:="Formularium wordt geladen, een ogenblik geduld a.u.b. ...", _
    Buttons:=vbInformation, Title:="Informedica 2016"
    
    Set objFormularium = New ClassFormularium
    
    For intN = 1 To objFormularium.MedicamentCount
        cboGeneriek.AddItem objFormularium.Item(intN).Generiek
    Next intN

End Sub

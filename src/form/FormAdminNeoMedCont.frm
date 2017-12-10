VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAdminNeoMedCont 
   Caption         =   "Neo Continue Medicatie Configuratie"
   ClientHeight    =   11685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19575
   OleObjectBlob   =   "FormAdminNeoMedCont.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAdminNeoMedCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_MedCol As Collection
Private m_PrevSel As Integer ' Holds medicament index is 1 based, note! lbxMedicamenten is 0 based
Private m_SetAdvice As Boolean

Private Function Validate(ByVal strMsg As String) As Boolean

    Dim blnValid As Boolean

    strMsg = IIf(MinSmallerThanMax(txtMinConc.Value, txtMaxConc.Value), strMsg, "Minimum concentratie kan niet groter dan maximum concentratie zijn")
    strMsg = IIf(MinSmallerThanMax(txtMaxConc.Value, txtConc.Value), strMsg, "Maximum concentratie kan niet groter dan ampul concentratie zijn")
    strMsg = IIf(MinSmallerThanMax(txtMinDose.Value, txtMaxDose.Value), strMsg, "Minimum dosering kan niet groter dan maximum dosering zijn")
    strMsg = IIf(MinSmallerThanMax(txtMaxDose.Value, txtAbsMax.Value), strMsg, "Absolute max dosering kan niet groter dan maximum dosering zijn")
    strMsg = IIf(MinSmallerThanMax(txtMinDose.Value, txtAbsMax.Value), strMsg, "Minimum dosering kan niet groter dan absolute maximum dosering zijn")

    lblValid.Caption = strMsg
    blnValid = strMsg = vbNullString
    cmdOK.Enabled = blnValid
    
    Validate = blnValid

End Function

Private Sub LoadMedicamenten()

    If m_MedCol Is Nothing Then Set m_MedCol = ModAdmin.Admin_GetNeoMedCont()

End Sub

Private Function MinSmallerThanMax(ByVal txtMin As String, ByVal txtMax As String) As Boolean

    Dim blnMinSmaller As Boolean
    Dim dblMin As Double
    Dim dblMax As Double
    
    dblMin = ModString.StringToDouble(txtMin)
    dblMax = ModString.StringToDouble(txtMax)
    
    blnMinSmaller = dblMin <= dblMax
    blnMinSmaller = blnMinSmaller Or dblMax = 0

    MinSmallerThanMax = blnMinSmaller
    
End Function

Private Sub LoadOplVlst()

    Dim objVlst As Range
    Dim strValue As String
    Dim intN As Integer
    Dim intC As Integer
    
    Set objVlst = NeoInfB_GetNeoOplVlst()
    intC = objVlst.Rows.Count
    
    For intN = 1 To intC
        strValue = objVlst.Cells(intN, 1).Value2
        cboOplVlst.AddItem strValue
    Next

End Sub

Private Sub cboDoseUnit_Change()

    lblMinDoseUnit.Caption = cboDoseUnit.Text
    lblMaxDoseUnit.Caption = cboDoseUnit.Text
    lblAbsMaxUnit.Caption = cboDoseUnit.Text

End Sub

Private Sub cboUnit_Change()

    Dim strUnit As String
    
    strUnit = cboUnit.Text & "/ml"
    
    lblConcUnit.Caption = strUnit
    lblMinConcUnit.Caption = strUnit
    lblMaxConcUnit.Caption = strUnit

End Sub

Private Sub cmdCancel_Click()

    lblButton.Caption = "Cancel"
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Hide
    lblButton.Caption = "OK"
    lbxMedicamenten_Click
    Admin_SetNeoMedCont m_MedCol, txtVerdunning.Value

End Sub

Private Sub cmdPrint_Click()

    shtNeoTblMedIV.Visible = xlSheetVisible
    ModSheet.PrintSheetAllPortrait shtNeoTblMedIV
    shtNeoTblMedIV.Visible = xlSheetVeryHidden

End Sub

Private Sub cmdSave_Click()

    Me.Hide
    lblButton.Caption = "OK"
    lbxMedicamenten_Click
    Admin_SetNeoMedCont m_MedCol, txtVerdunning
    Application_SaveNeoMedContConfig

End Sub

Private Sub lbxMedicamenten_Click()

    Dim intSel As Integer
    
    intSel = lbxMedicamenten.ListIndex
    
    LoadMedicament intSel

End Sub

Private Sub txtAbsMax_Change()

    Validate vbNullString

End Sub

Private Sub txtAbsMax_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtConc_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtVol_KeyPress(ByVal intKey As MSForms.ReturnInteger)
    
    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtOplVol_KeyPress(ByVal intKey As MSForms.ReturnInteger)
    
    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtRate_KeyPress(ByVal intKey As MSForms.ReturnInteger)
    
    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtMinConc_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtMaxConc_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtMinDose_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtMaxDose_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtHoudbaar_KeyPress(ByVal intKey As MSForms.ReturnInteger)
    
    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtAdvice_Change()

    m_SetAdvice = Not txtAdvice.Text = vbNullString

End Sub

Private Sub txtConc_Change()
    
    Validate vbNullString

End Sub

Private Sub txtMinConc_Change()
    
    Validate vbNullString

End Sub

Private Sub txtMaxConc_Change()
    
    Validate vbNullString

End Sub

Private Sub txtMinDose_Change()
    
    Validate vbNullString

End Sub

Private Sub txtMaxDose_Change()
    
    Validate vbNullString

End Sub

Private Sub UserForm_Activate()

    CenterForm

End Sub

Private Sub UserForm_Initialize()

    Dim objMed As ClassNeoMedCont
    
    LoadMedicamenten
    LoadOplVlst
    
    For Each objMed In m_MedCol
        lbxMedicamenten.AddItem objMed.Name
    Next
    
    txtVerdunning.Text = Admin_GetNeoMedVerdunning()
    lbxMedicamenten.ListIndex = 0

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub LoadMedicament(ByVal intSel As Integer)

    Dim objMed As ClassNeoMedCont
    
    If intSel = -1 Then Exit Sub
    
    UpdatePreviousSelection
    Set objMed = m_MedCol(intSel + 1)
    
    With objMed
        lblMed.Caption = .Name
        cboUnit.Text = .Unit
        cboDoseUnit.Text = .DoseUnit
        txtConc.Text = .Conc
        txtVol.Text = .Volume
        cboOplVlst.Text = cboOplVlst.List(.OplVlst - 1) ' NeoInfB_GetNeoOplVlst(.OplVlst)
        txtOplVol.Text = .OplVol
        txtRate.Text = .Rate
        txtMinConc.Text = .MinConc
        txtMaxConc.Text = .MaxConc
        txtMinDose.Text = .MinDose
        txtMaxDose.Text = .MaxDose
        txtAbsMax.Text = .AbsMax
        txtAdvice.Text = IIf(.Advice = vbNullString, GetAdvice(objMed), .Advice)
        txtProduct.Text = .Product
        txtHoudbaar.Text = .Houdbaar
        txtBewaar.Text = .Bewaar
        txtBereiding.Text = .Tekst
    End With
    
    m_PrevSel = intSel + 1

End Sub

Private Sub UpdatePreviousSelection()

    Dim objMed As ClassNeoMedCont
    
    If m_PrevSel = 0 Then Exit Sub
    
    Set objMed = m_MedCol(m_PrevSel)
    
    With objMed
        .Unit = cboUnit.Text
        .DoseUnit = cboDoseUnit.Text
        .Conc = txtConc.Text
        .Volume = txtVol.Text
        .OplVlst = IIf(cboOplVlst.ListIndex = -1, 0, cboOplVlst.ListIndex) + 1
        .OplVol = txtOplVol.Text
        .Rate = txtRate.Text
        .MinConc = txtMinConc.Text
        .MaxConc = txtMaxConc.Text
        .MinDose = txtMinDose.Text
        .MaxDose = txtMaxDose.Text
        .AbsMax = txtAbsMax.Text
        .Advice = IIf(m_SetAdvice, txtAdvice.Text, vbNullString)
        .Product = txtProduct.Text
        .Houdbaar = txtHoudbaar.Text
        .Bewaar = txtBewaar.Text
        .Tekst = txtBereiding.Text
    End With

End Sub

Private Function GetAdvice(objMed As ClassNeoMedCont) As String

    Dim strAdv As String
    
    strAdv = objMed.MinDose & " - " & objMed.MaxDose & " " & objMed.DoseUnit
    
    GetAdvice = strAdv

End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If Cancel = 0 Then cmdCancel_Click

End Sub
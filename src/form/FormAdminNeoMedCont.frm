VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAdminNeoMedCont 
   Caption         =   "Neo Continue Medicatie Configuratie"
   ClientHeight    =   11685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22695
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
Private m_SelectedVersion As Integer

Private Const ConstCaption As String = "Neonatale Continue Medicatie Configuratie"

Private Sub ClearMedDetails()

        lblMed.Caption = ""
        cboUnit.Text = ""
        cboDoseUnit.Text = ""
        txtConc.Text = ""
        txtVol.Text = ""
        cboOplVlst.Text = ""
        chkSolReq.Value = False
        txtOplVol.Text = ""
        txtRate.Text = ""
        txtMinConc.Text = ""
        txtMaxConc.Text = ""
        txtMinDose.Text = ""
        txtMaxDose.Text = ""
        txtAbsMax.Text = ""
        txtAdvice.Text = ""
        txtProduct.Text = ""
        txtHoudbaar.Text = ""
        txtBewaar.Text = ""
        txtBereiding.Text = ""
        txtVerdunning.Text = ""
        
        m_PrevSel = 0

End Sub


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

Private Sub LoadMedicationCollection()

    Dim objMed As ClassNeoMedCont
    
    If Setting_UseDatabase Then
        Caption = ConstCaption & IIf(m_SelectedVersion = 0, "", " Versie: " & m_SelectedVersion)
        
        Set m_MedCol = ModDatabase.Database_GetNeoConfigMedCont(m_SelectedVersion)
    Else
        Set m_MedCol = ModAdmin.Admin_MedContNeoGetCollection()
    End If
    
    lbxMedicamenten.Clear
    For Each objMed In m_MedCol
        lbxMedicamenten.AddItem objMed.Generic
    Next
    
    ClearMedDetails
    
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

Private Sub LoadSolution()

    Dim objSol As Range
    Dim strValue As String
    Dim intN As Integer
    Dim intC As Integer
    
    Set objSol = NeoInfB_GetNeoOplVlst()
    intC = objSol.Rows.Count
    
    For intN = 1 To intC
        strValue = objSol.Cells(intN, 1).Value2
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

Private Sub cmdImport_Click()

    Dim objConfigWbk As Workbook
    Dim objSrc As Range
    Dim objDst As Range
    Dim lngErr As Long
    Dim strFile As String
        
    Dim objMed As ClassNeoMedCont
    
    strFile = ModFile.GetFileWithDialog
    
    Dim strMsg As String
    
    On Error GoTo HandleError
       
    Application.DisplayAlerts = False
        
    Set objConfigWbk = Workbooks.Open(strFile, True, True)
    Set objSrc = objConfigWbk.Sheets(CONST_TBL_MEDCONT_NEO).Range(CONST_TBL_MEDCONT_NEO)
    Set objDst = ModRange.GetRange(CONST_TBL_MEDCONT_NEO)
        
    Sheet_CopyRangeFormulaToDst objSrc, objDst
    Sheet_CopyRangeFormulaToDst objConfigWbk.Sheets(CONST_MEDCONTVERDUNNING_NEO).Range("A1"), ModRange.GetRange(CONST_MEDCONTVERDUNNING_NEO)
    
    Set m_MedCol = ModAdmin.Admin_MedContNeoGetCollection()
    
    lbxMedicamenten.Clear
    For Each objMed In m_MedCol
        lbxMedicamenten.AddItem objMed.Generic
    Next
    
    ClearMedDetails
    
    objConfigWbk.Close
    Application.DisplayAlerts = True
    
    Exit Sub
    
HandleError:

    objConfigWbk.Close
    Application.DisplayAlerts = True
    ModLog.LogError Err, "Could not import: " & strFile

End Sub

Private Sub cmdOK_Click()

    Me.Hide
    lblButton.Caption = "OK"
    lbxMedicamenten_Click
    Admin_MedContNeoSetCollection m_MedCol, txtVerdunning.Value

End Sub

Private Sub cmdPrint_Click()

    shtNeoTblMedIV.Visible = xlSheetVisible
    ModSheet.PrintSheetAllPortrait shtNeoTblMedIV
    shtNeoTblMedIV.Visible = xlSheetVeryHidden

End Sub

Private Sub cboVersions_Change()

    If Not cboVersions.Value = vbNullString Then
        Set m_MedCol = Nothing
        m_SelectedVersion = Database_GetVersionIDFromString(cboVersions.Value)
        LoadMedicationCollection
    End If
    
End Sub

Private Sub cmdSave_Click()

    Me.Hide
    lblButton.Caption = "OK"
    lbxMedicamenten_Click
    
    Admin_MedContNeoSetCollection m_MedCol, txtVerdunning
    
    If Setting_UseDatabase Then
        Database_SaveNeoConfigMedCont
    Else
        App_SaveNeoMedContConfig
    End If
    
End Sub

Private Sub lbxMedicamenten_Click()

    Dim intSel As Integer
    
    intSel = lbxMedicamenten.ListIndex
    
    LoadMedicationDetails intSel

End Sub

Private Sub optLastVersion_Click()

    ToggleVersionSelect False

End Sub

Private Sub LoadVersions()

    Dim colVersions As Collection
    Dim objVersion As ClassVersion
    Dim intN As Integer
        
    cboVersions.Clear
    Set colVersions = Database_GetConfigMedContVersions(CONST_DEP_NICU)
    
    For Each objVersion In colVersions
        cboVersions.AddItem objVersion.ToString()
    Next
    
End Sub

Private Sub ToggleVersionSelect(ByVal blnVisible As Boolean)

    If Not blnVisible Then
        m_SelectedVersion = 0
        LoadMedicationCollection
    Else
        LoadVersions
    End If
    
    cboVersions.Visible = blnVisible
    lblCboVersions.Visible = blnVisible

End Sub

Private Sub optSpecificVersion_Click()

    ToggleVersionSelect True

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

    optLastVersion.Value = True
    LoadSolution
    
    If Not Setting_UseDatabase Then txtVerdunning.Text = Admin_MedContNeoGetVerdunning()
    lbxMedicamenten.ListIndex = 0

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub LoadMedicationDetails(ByVal intSel As Integer)

    Dim objMed As ClassNeoMedCont
    
    If intSel = -1 Then Exit Sub
    
    UpdatePreviousSelection
    Set objMed = m_MedCol(intSel + 1)
    
    With objMed
        lblMed.Caption = .Generic
        cboUnit.Text = .GenericUnit
        cboDoseUnit.Text = .DoseUnit
        txtConc.Text = .GenericQuantity
        txtVol.Text = .GenericVolume
        cboOplVlst.Text = cboOplVlst.List(.Solution - 1) ' NeoInfB_GetNeoOplVlst(.OplVlst)
        txtOplVol.Text = .SolutionVolume
        chkSolReq.Value = .SolutionRequired
        txtRate.Text = .DripQuantity
        txtMinConc.Text = .MinConcentration
        txtMaxConc.Text = .MaxConcentration
        txtMinDose.Text = .MinDose
        txtMaxDose.Text = .MaxDose
        txtAbsMax.Text = .AbsMaxDose
        txtAdvice.Text = IIf(.DoseAdvice = vbNullString, GetAdvice(objMed), .DoseAdvice)
        txtProduct.Text = .Product
        txtHoudbaar.Text = .ShelfLife
        txtBewaar.Text = .ShelfCondition
        txtBereiding.Text = .PreparationText
        If Setting_UseDatabase Then txtVerdunning.Text = .DilutionText
    End With
    
    m_PrevSel = intSel + 1

End Sub

Private Sub UpdatePreviousSelection()

    Dim objMed As ClassNeoMedCont
    
    If m_PrevSel = 0 Then Exit Sub
    
    Set objMed = m_MedCol(m_PrevSel)
    
    With objMed
        .GenericUnit = cboUnit.Text
        .DoseUnit = cboDoseUnit.Text
        .GenericQuantity = txtConc.Text
        .GenericVolume = txtVol.Text
        .Solution = IIf(cboOplVlst.ListIndex = -1, 0, cboOplVlst.ListIndex) + 1
        .SolutionVolume = txtOplVol.Text
        .SolutionRequired = chkSolReq.Value
        .DripQuantity = txtRate.Text
        .MinConcentration = txtMinConc.Text
        .MaxConcentration = txtMaxConc.Text
        .MinDose = txtMinDose.Text
        .MaxDose = txtMaxDose.Text
        .AbsMaxDose = txtAbsMax.Text
        .DoseAdvice = IIf(txtAdvice = GetAdvice(objMed), vbNullString, txtAdvice.Text)
        .Product = txtProduct.Text
        .ShelfLife = txtHoudbaar.Text
        .ShelfCondition = txtBewaar.Text
        .PreparationText = txtBereiding.Text
        
        If Setting_UseDatabase Then
            For Each objMed In m_MedCol
                .DilutionText = txtVerdunning.Text
            Next objMed
        End If
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

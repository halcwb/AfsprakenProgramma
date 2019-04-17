VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAdminParEnt 
   Caption         =   "Parentaralia configuratie"
   ClientHeight    =   10335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14205
   OleObjectBlob   =   "FormAdminParent.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAdminParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ParEntCol As Collection
Private m_Versions As Collection
Private m_SelectedVersion As Integer
Private m_PrevSel As Integer

Private Const ConstCaption As String = "Parenteralia Configuratie"

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub LoadVersions()

    Dim objVersion As ClassVersion
    
    cboVersions.Clear
    Set m_Versions = Database_GetConfigParEntVersions()
    
    For Each objVersion In m_Versions
        cboVersions.AddItem objVersion.ToString()
    Next
    
End Sub

Private Sub ToggleVersionSelect(ByVal blnVisible As Boolean)

    If Not blnVisible Then
        m_SelectedVersion = 0
        Set m_Versions = Nothing
        LoadParEntCollection
    Else
        LoadVersions
    End If
    
    cboVersions.Visible = blnVisible
    lblCboVersions.Visible = blnVisible

End Sub

Private Sub cboVersions_Change()

    If Not cboVersions.Value = vbNullString Then
        Set m_ParEntCol = Nothing
        m_SelectedVersion = Database_GetVersionIDFromString(cboVersions.Value)
        LoadParEntCollection
    End If
    
End Sub

Private Sub cmdImport_Click()

    Dim objConfigWbk As Workbook
    Dim objSrc As Range
    Dim objDst As Range
    Dim lngErr As Long
    Dim strFile As String
        
    Dim objParEnt As ClassParent
    
    strFile = ModFile.GetFileWithDialog
    If strFile = "" Then Exit Sub
    
    Dim strMsg As String
    
    On Error GoTo HandleError
       
    Application.DisplayAlerts = False
        
    Set objConfigWbk = Workbooks.Open(strFile, True, True)
    Set objSrc = objConfigWbk.Sheets(constGlobParEntTbl).Range(constGlobParEntTbl)
    Set objDst = ModRange.GetRange(constGlobParEntTbl)
        
    Sheet_CopyRangeFormulaToDst objSrc, objDst
    
    Set m_ParEntCol = ModAdmin.Admin_GetParEnt()
    
    lbxParenteralia.Clear
    For Each objParEnt In m_ParEntCol
        lbxParenteralia.AddItem objParEnt.Name
    Next
    
    ClearParEntDetails
    
    objConfigWbk.Close
    Application.DisplayAlerts = True
    
    Exit Sub
    
HandleError:

    objConfigWbk.Close
    Application.DisplayAlerts = True
    ModLog.LogError Err, "Could not import: " & strFile
End Sub

Private Sub optLastVersion_Click()

    ToggleVersionSelect False

End Sub

Private Sub optSpecificVersion_Click()

    ToggleVersionSelect True

End Sub

Private Sub UserForm_Initialize()
    
    optLastVersion.Value = True
    lbxParenteralia.ListIndex = 0
    
End Sub

Private Sub LoadParEntCollection()

    Dim objParEnt As ClassParent
    
    If Setting_UseDatabase Then
        Caption = ConstCaption & IIf(m_SelectedVersion = 0, "", " Versie: " & m_SelectedVersion)

        Set m_ParEntCol = ModDatabase.Database_GetConfigParEnt(m_SelectedVersion)
    Else
        Set m_ParEntCol = ModAdmin.Admin_GetParEnt()
    End If
    
    lbxParenteralia.Clear
    For Each objParEnt In m_ParEntCol
        lbxParenteralia.AddItem objParEnt.Name
    Next
    
    ClearParEntDetails

End Sub

Private Sub LoadParEntDetails(ByVal intSel As Integer)

    Dim objParEnt As ClassParent
    
    If intSel = -1 Then Exit Sub
    
    UpdatePreviousSelection
    Set objParEnt = m_ParEntCol(intSel + 1)
    
    With objParEnt
        lblName.Caption = .Name
        txtEnergy.Value = .Energy
        txtEiwit.Value = .Eiwit
        txtKH.Value = .KH
        txtVet.Value = .Vet
        txtNa.Value = .Na
        txtK.Value = .K
        txtCa.Value = .Ca
        txtP.Value = .P
        txtMg.Value = .Mg
        txtFe.Value = .Fe
        txtVitD.Value = .VitD
        txtCl.Value = .Cl
        txtProduct.Value = .Product
    End With
    
    m_PrevSel = intSel + 1

End Sub

Private Sub ClearParEntDetails()

        lblName.Caption = ""
        txtEnergy.Value = ""
        txtEiwit.Value = ""
        txtKH.Value = ""
        txtVet.Value = ""
        txtNa.Value = ""
        txtK.Value = ""
        txtCa.Value = ""
        txtP.Value = ""
        txtMg.Value = ""
        txtFe.Value = ""
        txtVitD.Value = ""
        txtCl.Value = ""
        txtProduct.Value = ""
        
        m_PrevSel = 0

End Sub

Private Function TextToNum(ByVal strText As String) As Double

    On Error GoTo ErrorHandler
    
    TextToNum = CDbl(strText)
    
    Exit Function
    
ErrorHandler:

    TextToNum = 0

End Function

Private Sub UpdatePreviousSelection()

    Dim objParEnt As ClassParent
    
    If m_PrevSel = 0 Then Exit Sub
    
    Set objParEnt = m_ParEntCol(m_PrevSel)
    
    With objParEnt
        objParEnt.Energy = TextToNum(txtEnergy.Value)
        objParEnt.Eiwit = TextToNum(txtEiwit.Value)
        objParEnt.KH = TextToNum(txtKH.Value)
        objParEnt.Vet = TextToNum(txtVet.Value)
        objParEnt.Na = TextToNum(txtNa.Value)
        objParEnt.K = TextToNum(txtK.Value)
        objParEnt.Ca = TextToNum(txtCa.Value)
        objParEnt.P = TextToNum(txtP.Value)
        objParEnt.Mg = TextToNum(txtMg.Value)
        objParEnt.Fe = TextToNum(txtFe.Value)
        objParEnt.VitD = TextToNum(txtVitD.Value)
        objParEnt.Cl = TextToNum(txtCl.Value)
        objParEnt.Product = txtProduct.Value
    End With

End Sub

Private Sub lbxParenteralia_Click()

    Dim intSel As Integer
    
    intSel = lbxParenteralia.ListIndex
    
    LoadParEntDetails intSel

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If Cancel = 0 Then cmdCancel_Click

End Sub

Private Sub cmdCancel_Click()

    lblButton.Caption = "Cancel"
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Hide
    
    lblButton.Caption = "OK"
    lbxParenteralia_Click
    Admin_SetParEnt m_ParEntCol

End Sub

Private Sub cmdSave_Click()

    Me.Hide
    
    lblButton.Caption = "OK"
    lbxParenteralia_Click
    Admin_SetParEnt m_ParEntCol
    
    If Setting_UseDatabase() Then
        Database_SaveConfigParEnt
    Else
        Application_SaveParEntConfig
    End If
    
End Sub

Private Sub txtEnergy_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtEiwit_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtKH_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtVet_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtNa_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtK_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtCa_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtP_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtMg_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtFe_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtVitD_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

Private Sub txtCl_KeyPress(ByVal intKey As MSForms.ReturnInteger)

    intKey = ModUtils.CorrectNumberAscii(intKey)

End Sub

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
Private m_Versions() As String
Private m_SelectedVersion As String
Private m_PrevSel As Integer

Private Const ConstCaption As String = "Parenteralia Configuratie"

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub LoadVersions()

    Dim arrVersions() As String
    Dim intN As Integer
        
    cboVersions.Clear
    arrVersions = Database_GetConfigParEntVersions()
    
    If Not ModArray.ArrayIsEmpty(arrVersions) Then
        For intN = 0 To UBound(arrVersions)
            cboVersions.AddItem arrVersions(intN)
        Next
    End If
    
    
End Sub

Private Sub ToggleVersionSelect(ByVal blnVisible As Boolean)

    If Not blnVisible Then
        m_SelectedVersion = vbNullString
        LoadParEnteralia
    Else
        LoadVersions
    End If
    
    cboVersions.Visible = blnVisible
    lblCboVersions.Visible = blnVisible

End Sub

Private Sub cboVersions_Change()

    If Not cboVersions.Value = vbNullString Then
        Set m_ParEntCol = Nothing
        m_SelectedVersion = cboVersions.Value
        LoadParEnteralia
    End If
    
End Sub

Private Sub optLastVersion_Click()

    ToggleVersionSelect False

End Sub

Private Sub optSpecificVersion_Click()

    ToggleVersionSelect True

End Sub

Private Sub UserForm_Initialize()
    
    LoadParEnteralia
    ToggleVersionSelect False
    
    lbxParenteralia.ListIndex = 0
    
End Sub

Private Sub LoadParEnteralia()

    Dim objParEnt As ClassParent
    Dim strVersion As String
    
    If Setting_UseDatabase Then
        Caption = ConstCaption & IIf(m_SelectedVersion = vbNullString, "", " Versie: " & m_SelectedVersion)
        
        If Not m_SelectedVersion = vbNullString Then
            strVersion = FormatDateTimeSeconds(CDate(m_SelectedVersion))
        End If
        
        Set m_ParEntCol = ModDatabase.Database_GetConfigParEnt(strVersion)
    Else
        Set m_ParEntCol = ModAdmin.Admin_GetParEnt()
    End If
    
    lbxParenteralia.Clear
    For Each objParEnt In m_ParEntCol
        lbxParenteralia.AddItem objParEnt.Name
    Next
    
    ClearSelectedParEnt

End Sub

Private Sub LoadParEnt(ByVal intSel As Integer)

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
Private Sub ClearSelectedParEnt()

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
    
    LoadParEnt intSel

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

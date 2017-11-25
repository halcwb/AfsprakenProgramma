VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAdminParEnt 
   Caption         =   "Parentaralia configuratie"
   ClientHeight    =   10335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10980
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
Private m_PrevSel As Integer

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Initialize()

    Dim objParEnt As ClassParent
    
    LoadParEnteralia
    
    For Each objParEnt In m_ParEntCol
        lbxParenteralia.AddItem objParEnt.Name
    Next
    
    lbxParenteralia.ListIndex = 0
    
End Sub

Private Sub LoadParEnteralia()

    If m_ParEntCol Is Nothing Then Set m_ParEntCol = ModAdmin.Admin_GetParEnt()

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

Private Sub UpdatePreviousSelection()

    Dim objParEnt As ClassParent
    
    If m_PrevSel = 0 Then Exit Sub
    
    Set objParEnt = m_ParEntCol(m_PrevSel)
    
    With objParEnt
        objParEnt.Energy = txtEnergy.Value
        objParEnt.Eiwit = txtEiwit.Value
        objParEnt.KH = txtKH.Value
        objParEnt.Vet = txtVet.Value
        objParEnt.Na = txtNa.Value
        objParEnt.K = txtK.Value
        objParEnt.Ca = txtCa.Value
        objParEnt.P = txtP.Value
        objParEnt.Mg = txtMg.Value
        objParEnt.Fe = txtFe.Value
        objParEnt.VitD = txtVitD.Value
        objParEnt.Cl = txtCl.Value
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
    Application_SaveParEntConfig

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

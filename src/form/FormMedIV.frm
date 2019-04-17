VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedIV 
   Caption         =   "Voer medicament in ..."
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   OleObjectBlob   =   "FormMedIV.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMedIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Validate(ByVal strValid As String)

    Dim strText As String
    
    strText = IIf(cboDoseUnit.Text = vbNullString, "Geef een doseer eenheid op", vbNullString)
    strText = IIf(cboUnit.Text = vbNullString, "Geef een eenheid op", strText)
    strText = IIf(txtSterkte.Text = vbNullString Or txtSterkte.Text = "0", "Voer een sterkte in", strText)
    strText = IIf(txtMedicament.Text = vbNullString, "Medicament moet een naam hebben", strText)
    
    strText = IIf(strValid = vbNullString, strText, strValid)
    
    lblValid.Caption = strText
    cmdOK.Enabled = strText = vbNullString

End Sub

Private Sub Clear()

    txtMedicament.Text = vbNullString
    txtSterkte.Text = vbNullString
    txtSolVol.Text = vbNullString
    cboUnit.Text = vbNullString
    cboDoseUnit.Text = vbNullString
    
    Validate vbNullString

End Sub

Private Sub cmdCancel_Click()
    
    Clear
    Me.Hide

End Sub

Private Sub cmdOK_Click()
    
    Me.Hide

End Sub

Private Sub cboDoseUnit_Change()

    Validate vbNullString

End Sub

Private Sub cboUnit_Change()

    Validate vbNullString

End Sub

Private Sub txtMedicament_Change()

    Validate vbNullString

End Sub

Private Sub txtSterkte_Change()

    Validate vbNullString

End Sub

Private Sub txtSolVol_KeyPress(ByVal intKeyAscii As MSForms.ReturnInteger)

    intKeyAscii = ModUtils.CorrectNumberAscii(intKeyAscii)

End Sub

Private Sub txtSterkte_KeyPress(ByVal intKeyAscii As MSForms.ReturnInteger)

    intKeyAscii = ModUtils.CorrectNumberAscii(intKeyAscii)

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm
    
End Sub

Private Sub FillUnitCombo()

    Dim objRange As Range
    Dim intN As Integer
    Dim intC As Integer
    
    Set objRange = ModRange.GetRange("Tbl_Glob_Conv_EenhCont")

    intC = objRange.Columns.Count
    For intN = 3 To intC
        cboUnit.AddItem objRange.Cells(1, intN).Value
    Next

End Sub

Private Sub FillDoseUnitCombo()

    Dim objRange As Range
    Dim intN As Integer
    Dim intC As Integer
    
    Set objRange = ModRange.GetRange("Tbl_Glob_Conv_EenhCont")

    intC = objRange.Rows.Count
    For intN = 2 To intC
        cboDoseUnit.AddItem objRange.Cells(intN, 1).Value
    Next

End Sub

Private Sub UserForm_Initialize()

    txtMedicament.TabIndex = 0
    txtSterkte.TabIndex = 1
    cboUnit.TabIndex = 3
    txtSolVol.TabIndex = 4
    cboDoseUnit.TabIndex = 5
    
    cmdOK.TabIndex = 6
    cmdCancel.TabIndex = 7
    
    FillUnitCombo
    FillDoseUnitCombo
    
    Validate vbNullString

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Clear
    Cancel = False
    
End Sub



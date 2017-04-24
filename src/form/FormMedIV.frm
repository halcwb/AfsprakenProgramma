VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedIV 
   Caption         =   "Voer medicament in ..."
   ClientHeight    =   1905
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

Private Sub Validate(strValid As String)

    Dim strText As String
    
    strText = IIf(txtEenheid.Text = vbNullString, "Geef een eenheid op", vbNullString)
    strText = IIf(txtSterkte.Text = vbNullString Or txtSterkte.Text = "0", "Voer een sterkte in", strText)
    strText = IIf(txtMedicament.Text = vbNullString, "Medicament moet een naam hebben", strText)
    
    strText = IIf(strValid = vbNullString, strText, strValid)
    
    lblValid.Caption = strText
    cmdOK.Enabled = strText = vbNullString

End Sub

Private Sub Clear()

    txtMedicament.Text = vbNullString
    txtSterkte.Text = vbNullString
    txtEenheid.Text = vbNullString

    Validate vbNullString

End Sub

Private Sub cmdCancel_Click()
    
    Clear
    Me.Hide

End Sub

Private Sub cmdOK_Click()
    
    Me.Hide

End Sub

Private Sub txtEenheid_Change()

    Validate vbNullString

End Sub

Private Sub txtMedicament_Change()

    Validate vbNullString

End Sub

Private Sub txtSterkte_Change()

    Validate vbNullString

End Sub

Private Sub txtSterkte_KeyPress(ByVal intKeyAscii As MSForms.ReturnInteger)

    intKeyAscii = ModUtils.CorrectNumberAscii(intKeyAscii)

End Sub

Private Sub UserForm_Initialize()

    txtMedicament.TabIndex = 0
    txtSterkte.TabIndex = 1
    txtEenheid.TabIndex = 2
    
    cmdOK.TabIndex = 3
    cmdCancel.TabIndex = 4
    
    Validate vbNullString

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Clear
    Cancel = False
    
End Sub



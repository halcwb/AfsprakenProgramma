VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormFontPicker 
   Caption         =   "Kies een font"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4950
   OleObjectBlob   =   "FormFontPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormFontPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Validate()

    Dim strValid As String
    
    strValid = IIf(cboSize.Value = vbNullString, "Kies een grootte", strValid)
    strValid = IIf(cboFont.Value = vbNullString, "Kies een font", vbNullString)
    
    cmdOK.Enabled = strValid = vbNullString
    lblValid.Caption = strValid

End Sub

Private Sub cboFont_Change()

    If Not cboFont.Value = vbNullString Then
        lblExample.Font.Name = cboFont.Value
    End If

    Validate

End Sub

Private Sub cboSize_Change()

    If Not cboSize.Value = vbNullString Then
        lblExample.Font.Size = Int(cboSize.Value)
    End If
    
    Validate
    
End Sub

Private Sub chkBold_Click()

    lblExample.Font.Bold = chkBold.Value

End Sub

Private Sub chkItalic_Click()

    lblExample.Font.Italic = chkItalic.Value

End Sub

Private Sub cmdClear_Click()

    cboFont.Value = vbNullString
    cboSize.Value = vbNullString
    chkBold.Value = False
    chkItalic.Value = False

End Sub

Private Sub cmdOK_Click()

    Me.Hide

End Sub

Private Sub cmdCancel_Click()

    lblValid.Caption = "Cancel"
    Me.Hide

End Sub

Private Sub UserForm_Activate()

    Validate

End Sub

Private Sub UserForm_Initialize()

    Dim varFonts() As Variant
    Dim varN As Variant
    
    cboSize.AddItem 8
    cboSize.AddItem 9
    cboSize.AddItem 10
    cboSize.AddItem 11
    cboSize.AddItem 12
    cboSize.AddItem 14

    For Each varN In ModColors.GetFontNames
        cboFont.AddItem varN
    Next
        
End Sub

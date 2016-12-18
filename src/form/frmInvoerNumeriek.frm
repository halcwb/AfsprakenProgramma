VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInvoerNumeriek 
   ClientHeight    =   1183
   ClientLeft      =   42
   ClientTop       =   329
   ClientWidth     =   3199
   OleObjectBlob   =   "frmInvoerNumeriek.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInvoerNumeriek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Explicit

Private Sub cmdCancel_Click()
    Me.txtWaarde = vbNullString
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub txtWaarde_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii >= 48 And KeyAscii <= 57 Then
    'Numerieke waarde is OK
Else
    If KeyAscii = 46 Or KeyAscii = 44 Then
        'Decimale punt >> zet om in komma indien toetsenbord Engels en
        'User Interface Nederlands
'        If Application.LanguageSettings.LanguageID(msoLanguageIDExeMode) _
'        <> Application.LanguageSettings.LanguageID(msoLanguageIDUI) And _
'        Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1043 Then
            KeyAscii = 44
        Else
            'Bij elke andere waarde negeren
            KeyAscii = 0
            Beep
        End If
'    End If
End If
        

End Sub

Private Sub UserForm_Activate()

    Me.txtWaarde.SetFocus

End Sub

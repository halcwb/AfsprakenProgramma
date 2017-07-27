VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPedMedIVPickList 
   Caption         =   "Kies medicamenten"
   ClientHeight    =   7125
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   5835
   OleObjectBlob   =   "FormPedMedIVPickList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPedMedIVPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const constMaxSelected As Integer = 15

Private Function Validate() As Boolean
    
    Dim strValid As String
    
    strValid = IIf(GetSelectedMedicatieCount > constMaxSelected, "Kies maximaal " & constMaxSelected & " medicamenten", vbNullString)
    
    lblValid.Caption = strValid
    Validate = strValid = vbNullString

End Function

Public Function GetSelectedMedicatieCount() As Integer

    GetSelectedMedicatieCount = ModList.GetSelectedListCount(lstMedicamenten)

End Function

Public Sub LoadMedicamenten(colMeds As Collection)

    ModList.LoadListItems lstMedicamenten, colMeds

End Sub

Public Sub SelectMedicament(ByVal intN As Integer)

    ModList.SelectListItem lstMedicamenten, intN

End Sub

Public Function IsMedicamentSelected(ByVal intN As Integer) As Boolean

    IsMedicamentSelected = ModList.IsListItemSelected(lstMedicamenten, intN)

End Function

Public Sub UnselectMedicament(ByVal intN As Integer)

    ModList.UnselectListItem lstMedicamenten, intN

End Sub

Public Function HasSelectedMedicamenten() As Boolean
    
    HasSelectedMedicamenten = ModList.HasSelectedListItems(lstMedicamenten)

End Function

Public Function GetFirstSelectedMedicament(ByVal blnUnSelect As Boolean) As Integer

    GetFirstSelectedMedicament = ModList.GetFirstSelectedListItem(lstMedicamenten, blnUnSelect)

End Function

Public Function GetAction() As String

    If lblValid.Caption = "Cancel" Then
        GetAction = "Cancel"
    Else
        GetAction = vbNullString
    End If
    
End Function

Private Sub cmdCancel_Click()

    lblValid.Caption = "Cancel"
    Me.Hide

End Sub

Private Sub cmdClear_Click()

    Dim intN As Integer
    Dim intC As Integer
    
    intC = lstMedicamenten.ListCount - 1
    For intN = 0 To intC
        lstMedicamenten.Selected(intN) = False
    Next intN
    
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Hide
    
End Sub

Private Sub lstMedicamenten_Click()

    cmdOK.Enabled = Validate()

End Sub

Private Sub lstMedicamenten_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    cmdOK.Enabled = Validate()

End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm

End Sub

Private Sub UserForm_Initialize()

    lstMedicamenten.Clear

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    cmdCancel_Click

End Sub


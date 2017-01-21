VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPedEntPickList 
   Caption         =   "Kies Voedingen ..."
   ClientHeight    =   7875
   ClientLeft      =   14
   ClientTop       =   315
   ClientWidth     =   6706
   OleObjectBlob   =   "FormPedEntPickList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPedEntPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const constMaxVoedingSelected As Integer = 1
Private Const constMaxToevoegingSelected As Integer = 3

Private Function Validate() As Boolean
    
    Dim strValid As String
    
    strValid = IIf(GetSelectedToevoegingCount > constMaxToevoegingSelected, "Kies maximaal " & constMaxToevoegingSelected & " toevoegingen", vbNullString)
    strValid = IIf(GetSelectedVoedingCount > constMaxVoedingSelected, "Kies maximaal " & constMaxVoedingSelected & " voeding", strValid)
    
    lblValid.Caption = strValid
    Validate = strValid = vbNullString

End Function

Public Function GetSelectedToevoegingCount() As Integer

    GetSelectedToevoegingCount = ModList.GetSelectedListCount(lstToeVoeging)

End Function

Public Function GetSelectedVoedingCount() As Integer

    GetSelectedVoedingCount = ModList.GetSelectedListCount(lstVoeding)

End Function

Public Sub LoadToevoegingen(ByRef colToevoeging As Collection)

    ModList.LoadListItems lstToeVoeging, colToevoeging

End Sub

Public Sub LoadVoedingen(ByRef colVoeding As Collection)

    ModList.LoadListItems lstVoeding, colVoeding

End Sub

Public Sub SelectToevoeging(ByVal intN As Integer)

    ModList.SelectListItem lstToeVoeging, intN

End Sub

Public Sub SelectVoeding(ByVal intN As Integer)

    ModList.SelectListItem lstVoeding, intN

End Sub

Public Function IsToevoegingSelected(ByVal intN As Integer) As Boolean

    IsToevoegingSelected = ModList.IsListItemSelected(lstToeVoeging, intN)

End Function

Public Function IsVoedingSelected(ByVal intN As Integer) As Boolean

    IsVoedingSelected = ModList.IsListItemSelected(lstVoeding, intN)

End Function

Public Sub UnselectToevoeging(ByVal intN As Integer)

    ModList.UnselectListItem lstToeVoeging, intN

End Sub
Public Sub UnselectVoeding(ByVal intN As Integer)

    ModList.UnselectListItem lstVoeding, intN

End Sub

Public Function HasSelectedToevoegingen() As Boolean
    
    HasSelectedToevoegingen = ModList.HasSelectedListItems(lstToeVoeging)

End Function

Public Function HasSelectedVoedingen() As Boolean
    
    HasSelectedVoedingen = ModList.HasSelectedListItems(lstVoeding)

End Function

Public Function GetFirstSelectedToevoeging(ByVal blnUnSelect As Boolean) As Integer

    GetFirstSelectedToevoeging = ModList.GetFirstSelectedListItem(lstToeVoeging, blnUnSelect)

End Function

Public Function GetFirstSelectedVoeding(ByVal blnUnSelect As Boolean) As Integer

    GetFirstSelectedVoeding = ModList.GetFirstSelectedListItem(lstVoeding, blnUnSelect)

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
    
    intC = lstVoeding.ListCount - 1
    For intN = 0 To intC
        lstVoeding.Selected(intN) = False
    Next intN
    
    intC = lstToeVoeging.ListCount - 1
    For intN = 0 To intC
        lstToeVoeging.Selected(intN) = False
    Next intN
    
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Hide
    
End Sub

Private Sub lstToevoeging_Click()

    cmdOK.Enabled = Validate()

End Sub

Private Sub lstVoeding_Click()

    cmdOK.Enabled = Validate()

End Sub

Private Sub lstToevoeging_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    cmdOK.Enabled = Validate()

End Sub

Private Sub lstVoeding_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    cmdOK.Enabled = Validate()

End Sub

Private Sub UserForm_Initialize()

    lstVoeding.Clear
    lstToeVoeging.Clear

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    cmdCancel_Click

End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormNeoEntPickList 
   Caption         =   "Kies Voedingen ..."
   ClientHeight    =   5625
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   10920
   OleObjectBlob   =   "FormNeoEntPickList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormNeoEntPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const constMaxVoedSelected As Integer = 1
Private Const constMaxToevoegMMSelected As Integer = 4
Private Const constMaxToevoegKVSelected As Integer = 4

Private Function Validate() As Boolean
    
    Dim strValid As String
    
    strValid = IIf(GetSelectedToevoegKVCount > constMaxToevoegKVSelected, "Kies maximaal " & constMaxToevoegKVSelected & " kunstvoeding toevoegingen", vbNullString)
    strValid = IIf(GetSelectedToevoegMMCount > constMaxToevoegMMSelected, "Kies maximaal " & constMaxToevoegMMSelected & " moedermelk toevoegingen", vbNullString)
    strValid = IIf(GetSelectedVoedingCount > constMaxVoedSelected, "Kies maximaal " & constMaxVoedSelected & " voeding", strValid)
    
    lblValid.Caption = strValid
    Validate = strValid = vbNullString

End Function

Public Function GetSelectedToevoegMMCount() As Integer

    GetSelectedToevoegMMCount = ModList.GetSelectedListCount(lstToevoegMM)

End Function

Public Function GetSelectedToevoegKVCount() As Integer

    GetSelectedToevoegKVCount = ModList.GetSelectedListCount(lstToevoegKV)

End Function

Public Function GetSelectedVoedingCount() As Integer

    GetSelectedVoedingCount = ModList.GetSelectedListCount(lstVoeding)

End Function

Public Sub LoadToevoegMM(ByRef colToevoeging As Collection)

    ModList.LoadListItems lstToevoegMM, colToevoeging

End Sub

Public Sub LoadToevoegKV(ByRef colToevoeging As Collection)

    ModList.LoadListItems lstToevoegKV, colToevoeging

End Sub

Public Sub LoadVoedingen(ByRef colVoeding As Collection)

    ModList.LoadListItems lstVoeding, colVoeding

End Sub

Public Sub SelectToevoegMM(ByVal intN As Integer)

    ModList.SelectListItem lstToevoegMM, intN

End Sub

Public Sub SelectToevoegKV(ByVal intN As Integer)

    ModList.SelectListItem lstToevoegKV, intN

End Sub

Public Sub SelectVoeding(ByVal intN As Integer)

    ModList.SelectListItem lstVoeding, intN

End Sub

Public Function IsToevoegMMSelected(ByVal intN As Integer) As Boolean

    IsToevoegMMSelected = ModList.IsListItemSelected(lstToevoegMM, intN)

End Function
Public Function IsToevoegKVSelected(ByVal intN As Integer) As Boolean

    IsToevoegKVSelected = ModList.IsListItemSelected(lstToevoegKV, intN)

End Function

Public Function IsVoedingSelected(ByVal intN As Integer) As Boolean

    IsVoedingSelected = ModList.IsListItemSelected(lstVoeding, intN)

End Function

Public Sub UnselectToevoegMM(ByVal intN As Integer)

    ModList.UnselectListItem lstToevoegMM, intN

End Sub

Public Sub UnselectToevoegKV(ByVal intN As Integer)

    ModList.UnselectListItem lstToevoegKV, intN

End Sub

Public Sub UnselectVoeding(ByVal intN As Integer)

    ModList.UnselectListItem lstVoeding, intN

End Sub

Public Function HasSelectedToevKV() As Boolean
    
    HasSelectedToevKV = ModList.HasSelectedListItems(lstToevoegKV)

End Function

Public Function HasSelectedToevMM() As Boolean
    
    HasSelectedToevMM = ModList.HasSelectedListItems(lstToevoegMM)

End Function

Public Function HasSelectedVoedingen() As Boolean
    
    HasSelectedVoedingen = ModList.HasSelectedListItems(lstVoeding)

End Function

Public Function GetFirstSelectedToevoegMM(ByVal blnUnSelect As Boolean) As Integer

    GetFirstSelectedToevoegMM = ModList.GetFirstSelectedListItem(lstToevoegMM, blnUnSelect)

End Function

Public Function GetFirstSelectedToevoegKV(ByVal blnUnSelect As Boolean) As Integer

    GetFirstSelectedToevoegKV = ModList.GetFirstSelectedListItem(lstToevoegKV, blnUnSelect)

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
    
    intC = lstToevoegMM.ListCount - 1
    For intN = 0 To intC
        lstToevoegMM.Selected(intN) = False
    Next intN
    
    intC = lstToevoegKV.ListCount - 1
    For intN = 0 To intC
        lstToevoegKV.Selected(intN) = False
    Next intN
    
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Hide
    
End Sub

Private Sub lstToevoegKV_Click()

    cmdOK.Enabled = Validate()

End Sub

Private Sub lstToeVoegMM_Click()
    
    cmdOK.Enabled = Validate()


End Sub

Private Sub lstToevoegMM_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    cmdOK.Enabled = Validate()

End Sub

Private Sub lstToevoegKV_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    cmdOK.Enabled = Validate()

End Sub

Private Sub lstVoeding_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    cmdOK.Enabled = Validate()

End Sub

Private Sub UserForm_Initialize()

    lstVoeding.Clear
    lstToevoegMM.Clear

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    cmdCancel_Click

End Sub

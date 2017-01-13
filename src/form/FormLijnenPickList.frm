VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormLijnenPickList 
   Caption         =   "Kies lijnen ..."
   ClientHeight    =   7065
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   6705
   OleObjectBlob   =   "FormLijnenPickList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormLijnenPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const constMaxSelected As Integer = 6

Private Function Validate() As Boolean
    
    Dim strValid As String
    
    strValid = IIf(GetSelectedLijnCount > constMaxSelected, "Kies maximaal " & constMaxSelected & " lijnen", vbNullString)
    
    lblValid.Caption = strValid
    Validate = strValid = vbNullString

End Function

Public Function GetSelectedLijnCount() As Integer

    GetSelectedLijnCount = ModList.GetSelectedListCount(lstLijnen)

End Function

Public Sub LoadLijnen(ByRef colLijnen As Collection)

    ModList.LoadListItems lstLijnen, colLijnen

End Sub

Public Sub SelectLijn(ByVal intN As Integer)

    ModList.SelectListItem lstLijnen, intN

End Sub

Public Function IsLijnSelected(ByVal intN As Integer) As Boolean

    IsLijnSelected = ModList.IsListItemSelected(lstLijnen, intN)

End Function

Public Sub UnselectLijn(ByVal intN As Integer)

    ModList.UnselectListItem lstLijnen, intN

End Sub

Public Function HasSelectedLijnen() As Boolean
    
    HasSelectedLijnen = ModList.HasSelectedListItems(lstLijnen)

End Function

Public Function GetFirstSelectedLijn(ByVal blnUnSelect As Boolean) As Integer

    GetFirstSelectedLijn = ModList.GetFirstSelectedListItem(lstLijnen, blnUnSelect)

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
    
    intC = lstLijnen.ListCount - 1
    For intN = 0 To intC
        lstLijnen.Selected(intN) = False
    Next intN
    
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Hide
    
End Sub

Private Sub lstLijnen_Click()

    cmdOK.Enabled = Validate()

End Sub

Private Sub lstLijnen_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    cmdOK.Enabled = Validate()

End Sub

Private Sub UserForm_Initialize()

    lstLijnen.Clear

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    cmdCancel_Click

End Sub



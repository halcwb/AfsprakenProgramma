VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCopy1700 
   Caption         =   "17.00 uur Afspraken overnemen naar actuele afspraken"
   ClientHeight    =   11123
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   16597
   OleObjectBlob   =   "FormCopy1700.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCopy1700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const constVoeding As String = "Var_Neo_#_ContIV"
Private Const constIVCont As String = "Var_Neo_#_Voeding"

Private Sub cmdCancel_Click()
    
    Me.Hide

End Sub

Private Sub cmdOk_Click()
    
    ModNeoInfB.NeoInfB_Copy1700ToAct Me.optAlles.Value, Me.chkVoeding.Value, Me.chkContinueMedicatie.Value, Me.chkTPN.Value
    Me.Hide

End Sub

Private Sub optAlles_Click()
    
    frmVoeding.Enabled = False
    frmContMed.Enabled = False
    frmTPN.Enabled = False

End Sub

Private Sub optPerBlok_Click()
    
    frmVoeding.Enabled = True
    frmContMed.Enabled = True
    frmTPN.Enabled = True

End Sub

Private Sub AddItemToList(ByVal strList As String, ByVal strItem As String, ByVal intN As Integer, ByVal bln1700 As Boolean)

    strList = IIf(bln1700, Replace(strList, "Act", "1700"), strList)
    strItem = IIf(bln1700, Replace(strItem, "#", "1700"), Replace(strItem, "#", "InfB"))
    strItem = IIf(intN < 10, strItem & "_0" & intN, strItem & "_" & intN)
    
    Me.Controls(strList).AddItem ModRange.GetRangeValue(strItem, vbNullString)
    
End Sub

Private Sub UserForm_Activate()

    Dim intN As Integer
    Dim strList As String
    Dim strItem As String
    
    strList = "lstActVoed"
    strItem = constVoeding
    For intN = 1 To 15
        AddItemToList strList, strItem, intN, True
        AddItemToList strList, strItem, intN, False
    Next intN

    strList = "lstActMed"
    strItem = constIVCont
    For intN = 1 To 15
        AddItemToList strList, strItem, intN, True
        AddItemToList strList, strItem, intN, False
    Next intN

    strList = "lstActTPN"
    strItem = constIVCont
    For intN = 16 To 27
        AddItemToList strList, strItem, intN, True
        AddItemToList strList, strItem, intN, False
    Next intN

End Sub

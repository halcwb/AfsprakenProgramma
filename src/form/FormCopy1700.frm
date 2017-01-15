VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCopy1700 
   Caption         =   "17.00 uur Afspraken overnemen naar actuele afspraken"
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16590
   OleObjectBlob   =   "FormCopy1700.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCopy1700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const constVoeding As String = "Txt_Neo_InfB_ContIV"
Private Const constIVCont As String = "Txt_Neo_InfB_Voeding"

Private Sub cmdCancel_Click()
    
    Me.Hide

End Sub

Private Sub cmdOK_Click()
    
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
    strItem = IIf(intN < 10, strItem & "_0" & intN, strItem & "_" & intN)
    
    On Error Resume Next ' ToDo Improve error handling in Rubberduck
    Me.Controls(strList).AddItem ModRange.GetRangeValue(strItem, vbNullString)
    
End Sub

Private Sub UserForm_Initialize()

    Dim intN As Integer
    Dim strList As String
    Dim strItem As String
    
    
    ModProgress.StartProgress "Afspraken laden"
    
    ' First get the actual items
    ModNeoInfB.NeoInfB_SelectInfB False
    
    strList = "lstActVoed"
    strItem = constVoeding
    For intN = 1 To 15
        AddItemToList strList, strItem, intN, False
    Next intN

    strList = "lstActMed"
    strItem = constIVCont
    For intN = 1 To 15
        AddItemToList strList, strItem, intN, False
    Next intN

    strList = "lstActTPN"
    strItem = constIVCont
    For intN = 16 To 27
        AddItemToList strList, strItem, intN, False
    Next intN
    
    ' Then get the 1700 items
    ModNeoInfB.NeoInfB_SelectInfB True
    
    strList = "lstActVoed"
    strItem = constVoeding
    For intN = 1 To 15
        AddItemToList strList, strItem, intN, True
    Next intN

    strList = "lstActMed"
    strItem = constIVCont
    For intN = 1 To 15
        AddItemToList strList, strItem, intN, True
    Next intN

    strList = "lstActTPN"
    strItem = constIVCont
    For intN = 16 To 27
        AddItemToList strList, strItem, intN, True
    Next intN
    
    ModProgress.FinishProgress

End Sub

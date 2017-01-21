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

Private Const constVoeding As String = "Txt_Neo_InfB_Voeding_"
Private Const constVoedingCount As Integer = 10
Private Const constIVCont As String = "Txt_Neo_InfB_ContIV_"
Private Const constContIVCount As Integer = 16
Private Const constTPN As String = "Txt_Neo_InfB_TPN_"
Private Const constTPNCount As Integer = 13

Private Sub cmdCancel_Click()
    
    Me.Hide

End Sub

Private Sub cmdOK_Click()
    
    ModNeoInfB.NeoInfB_Copy1700ToAct Me.optAlles.value, Me.chkVoeding.value, Me.chkContinueMedicatie.value, Me.chkTPN.value
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

Private Sub RemoveDoubles(ByVal strList)

    Dim intActN As Integer
    Dim intActC As Integer
    Dim int1700N As Integer
    Dim int1700C As Integer
    Dim strAct As String
    Dim str1700 As String
    Dim objListAct As MSForms.ListBox
    Dim objList1700 As MSForms.ListBox
    
    Set objListAct = Me.Controls(strList)
    strList = Replace(strList, "Act", "1700")
    Set objList1700 = Me.Controls(strList)
    
    intActC = objListAct.ListCount - 1
    int1700C = objList1700.ListCount - 1
    
    For intActN = 0 To intActC
        strAct = objListAct.List(intActN)
        
        If Not strAct = vbNullString Then
            For int1700N = 0 To int1700C
                str1700 = objList1700.List(int1700N)
                If strAct = str1700 Then
                    objListAct.List(intActN) = vbNullString
                    objList1700.List(int1700N) = vbNullString
                    
                    Exit For
                End If
            Next
        End If
    Next

End Sub

Private Sub AddItemToList(ByVal strList As String, ByVal strItem As String, ByVal intN As Integer, ByVal bln1700 As Boolean)

    strList = IIf(bln1700, Replace(strList, "Act", "1700"), strList)
    strItem = IIf(intN < 10, strItem & "0" & intN, strItem & intN)
    
    Me.Controls(strList).AddItem ModRange.GetRangeValue(strItem, vbNullString)
    
End Sub

Private Sub AddItems(ByVal bln1700)

    Dim intN As Integer
    Dim strList As String
    Dim strItem As String
        
    strList = "lstActVoed"
    strItem = constVoeding
    For intN = 1 To constVoedingCount
        AddItemToList strList, strItem, intN, bln1700
    Next intN

    strList = "lstActMed"
    strItem = constIVCont
    For intN = 1 To constContIVCount
        AddItemToList strList, strItem, intN, bln1700
    Next intN

    strList = "lstActTPN"
    strItem = constTPN
    For intN = 1 To constTPNCount
        AddItemToList strList, strItem, intN, bln1700
    Next intN
    
End Sub

Private Sub UserForm_Initialize()

    Dim intN As Integer
    Dim strList As String
    Dim strItem As String
    
    ModProgress.StartProgress "Afspraken laden"
    
    ' First get the actual items
    ModNeoInfB.NeoInfB_SelectInfB False
    AddItems False
    
    ' Then get the 1700 items
    ModNeoInfB.NeoInfB_SelectInfB True
    AddItems True
    
    RemoveDoubles "lstActVoed"
    RemoveDoubles "lstActMed"
    RemoveDoubles "lstActTPN"
    
    ModProgress.FinishProgress

End Sub

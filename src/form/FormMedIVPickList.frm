VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMedIVPickList 
   Caption         =   "Kies medicamenten"
   ClientHeight    =   7112
   ClientLeft      =   21
   ClientTop       =   322
   ClientWidth     =   5831
   OleObjectBlob   =   "FormMedIVPickList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormMedIVPickList"
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

    Dim intN As Integer
    Dim intC As Integer
    
    intC = 0
    For intN = 0 To lstMedicamenten.ListCount - 1
        If lstMedicamenten.Selected(intN) Then intC = intC + 1
    Next intN
    
    GetSelectedMedicatieCount = intC

End Function

Public Sub LoadMedicamenten(colMeds As Collection)

    Dim varMed As Variant
    
    For Each varMed In colMeds
        lstMedicamenten.AddItem varMed
    Next varMed

End Sub

Public Sub SelectMedicament(ByVal intN As Integer)

    lstMedicamenten.Selected(intN - 2) = True

End Sub

Public Function IsMedicamentSelected(ByVal intN As Integer) As Boolean

    IsMedicamentSelected = lstMedicamenten.Selected(intN - 2)

End Function

Public Sub UnselectMedicament(ByVal intN As Integer)

    lstMedicamenten.Selected(intN - 2) = False

End Sub

Public Function HasSelectedMedicamenten() As Boolean
    
    HasSelectedMedicamenten = Not GetFirstSelectedMedicament(False) = 1

End Function

Public Function GetFirstSelectedMedicament(ByVal blnUnSelect As Boolean) As Integer

    Dim intN As Integer
    Dim intC As Integer
    
    intC = lstMedicamenten.ListCount - 1
    For intN = 0 To intC
        If lstMedicamenten.Selected(intN) Then
            If blnUnSelect Then lstMedicamenten.Selected(intN) = False
            GetFirstSelectedMedicament = intN + 2
            Exit Function
        End If
    Next intN
    
    GetFirstSelectedMedicament = 1

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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = True
    cmdCancel_Click

End Sub

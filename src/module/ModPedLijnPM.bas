Attribute VB_Name = "ModPedLijnPM"
Option Explicit

Private Const constOpm As String = "_Ped_Lijn_Opm"
Private Const constTblLijn As String = "Tbl_Ped_lijnen"
Private Const constLijnCount As Integer = 6
Private Const constLijnKeuze As String = "_Ped_Lijn_"
Private Const constPM As String = "_Ped_PM_"
Private Const constTblPMStand As String = "Tbl_Ped_PMStandaard"
Private Const constTblPMSet As String = "Tbl_Ped_PMInstelling"

Public Sub PedLijnPM_ShowPickList()

    Dim frmPickList As FormPedLijnenPickList
    Dim colLijnen As Collection
    Dim intN As Integer
    Dim intC As Integer
    Dim intKeuze As Integer
    
    Set colLijnen = New Collection
    intC = shtPedTblIV.Range(constTblLijn).Rows.Count
    For intN = 2 To intC
        colLijnen.Add shtPedTblIV.Range(constTblLijn).Cells(intN, 1)
    Next intN
    
    Set frmPickList = New FormPedLijnenPickList
    frmPickList.LoadLijnen colLijnen
    
    For intN = 1 To constLijnCount
        intKeuze = ModRange.GetRangeValue(constLijnKeuze & intN, 1)
        If intKeuze > 1 Then frmPickList.SelectLijn intKeuze
    Next intN
    
    frmPickList.Show
    
    If frmPickList.GetAction = vbNullString Then
    
        For intN = 1 To constLijnCount                 ' First remove nonselected items
            intKeuze = ModRange.GetRangeValue(constLijnKeuze & intN, 1)
            If intKeuze > 1 Then
                If frmPickList.IsLijnSelected(intKeuze) Then
                    frmPickList.UnselectLijn (intKeuze)
                Else
                    Clear intN
                End If
            End If
        Next intN
        
        Do While frmPickList.HasSelectedLijnen()  ' Then add selected items
            For intN = 1 To constLijnCount
                intKeuze = ModRange.GetRangeValue(constLijnKeuze & intN, 1)
                If intKeuze <= 1 Then
                    intKeuze = frmPickList.GetFirstSelectedLijn(True)
                    ModRange.SetRangeValue constLijnKeuze & intN, intKeuze
                    Exit For
                End If
            Next intN
        Loop
    
    End If
    
End Sub

Private Sub Clear(ByVal intN As Integer)

    ModRange.SetRangeValue constLijnKeuze & intN, 1

End Sub

Public Sub PedLijnPM_Clear_1()

    Clear 1
    
End Sub

Public Sub PedLijnPM_Clear_2()

    Clear 2
    
End Sub

Public Sub PedLijnPM_Clear_3()

    Clear 3
    
End Sub

Public Sub PedLijnPM_Clear_4()

    Clear 4
    
End Sub

Public Sub PedLijnPM_Clear_5()

    Clear 5
    
End Sub

Public Sub PedLijnPM_Clear_6()

    Clear 6
    
End Sub

Private Sub EnterOpm()

    Dim frmOpmerking As FormOpmerking
    
    Set frmOpmerking = New FormOpmerking
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(constOpm, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue constOpm, frmOpmerking.txtOpmerking.Text
    End If
    
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub

Public Sub PedLijnPM_Clear_Opm()

    ModRange.SetRangeValue constOpm, vbNullString

End Sub

Public Sub PedLijnPM_EnterText()
    
    EnterOpm
    
End Sub

Public Sub PedLijnPM_ClearPM()

    ModProgress.StartProgress "Verwijder PM"
    ModPatient.Patient_ClearData constPM, False, True
    ModProgress.FinishProgress

End Sub

Public Sub PedLijnPM_PaceMaker()

    shtPedBerIVenPM.Range(constTblPMStand).Copy
    shtPedBerIVenPM.Range(constTblPMSet).PasteSpecial xlPasteValues

End Sub


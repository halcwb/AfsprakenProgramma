Attribute VB_Name = "ModPedAfspr"
Option Explicit

Private Const constOverige = "_Ped_AfsprOverig"
Private Const constCompens = "_Ped_AfsprD_Verliezen" 'ToDo Remove duplicate name _Ped_AfsprD_VerliezenStof
Private Const constPedAfsprB = "_Ped_AfsprB_"
Private Const constPedAfsprD = "_Ped_AfsprD_"

Public Sub PedAfspr_Clear()

    ModProgress.StartProgress "Verwijder Ped Afspraken"
    ModPatient.ClearPatientData constPedAfsprB, False, True
    ModPatient.ClearPatientData constPedAfsprD, False, True
    ModProgress.FinishProgress

End Sub

Private Sub EnterText(strCaption As String, strName As String, strRange As String)

    Dim frmInvoer As New FormTekstInvoer
    
    With frmInvoer
        .Caption = strCaption
        .lblNaam.Caption = strName
        .Tekst = ModRange.GetRangeValue(strRange, vbNullString)
        .Show
        If .IsOK Then ModRange.SetRangeValue strRange, .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub PedAfspr_OverigeText()

    EnterText "Voer tekst in ...", "Voer overige aanvullende afspraken in", constOverige
    
End Sub

Public Sub PedAfspr_CompensateText()

    EnterText "Voer tekst in ...", "Voer compensatie vloeistof in", constCompens

End Sub

Attribute VB_Name = "ModNeoAfspr"
Option Explicit

Private Const constWondKweek As String = "_Neo_AfsprD_Wondkweek"
Private Const constOverig As String = "_Neo_AfsprD_OverigTekst"
Private Const constNeoAfsprB As String = "_Neo_AfsprB_"
Private Const constNeoAfsprD As String = "_Neo_AfsprD_"

Public Sub NeoAfspr_Clear()

    ModProgress.StartProgress "Verwijder Neo Afspraken"
    ModPatient.Patient_ClearData constNeoAfsprB, False, True
    ModPatient.Patient_ClearData constNeoAfsprD, False, True
    ModProgress.FinishProgress

End Sub

Private Sub EnterText(ByVal strCaption As String, ByVal strName As String, ByVal strRange As String)

    Dim frmInvoer As FormTekstInvoer
    
    Set frmInvoer = New FormTekstInvoer
    
    With frmInvoer
        .Caption = strCaption
        .lblNaam.Caption = strName
        .Tekst = ModRange.GetRangeValue(strRange, vbNullString)
        .Show
        If .IsOK Then ModRange.SetRangeValue strRange, .Tekst
    End With

End Sub

Public Sub NeoAfspr_WondText()

    EnterText "Voer tekst in ...", "Voer locatie wond(en) in", constWondKweek

End Sub

Public Sub NeoAfspr_OverigText()

    EnterText "Voer tekst in ...", "Voer een opmerking in voor overige afspraken", constOverig

End Sub


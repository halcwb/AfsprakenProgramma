Attribute VB_Name = "ModPedAfspr"
Option Explicit

Private Sub TekstInvoer(strCaption As String, strName As String, strRange As String)

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


Public Sub MedTekstWondkweek()

    TekstInvoer "Voer tekst in ...", "Voer locatie wond(en) in", "Aanvullend_WondkweekTekst"

End Sub

Public Sub MedTekstVerliezenCompenseren()

    TekstInvoer "Voer tekst in ...", "Voer verliezen compenseren in", "Aanvullend_VerliezenTekst"

End Sub

Public Sub MedTekstAanvullendeAfsprakenOverigePed()

    TekstInvoer "Voer tekst in ...", "Voer overige aanvullende afspraken in", "Aanvullend_Overige_Ped"
    
End Sub


Public Sub MedTekstAanvullendeAfsprakenVerliezenPed()

    TekstInvoer "Voer tekst in ...", "Voer compensatie vloeistof in", "Aanvullend_Verliezen_Ped"

End Sub

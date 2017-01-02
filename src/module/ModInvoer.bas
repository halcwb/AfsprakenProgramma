Attribute VB_Name = "ModInvoer"
Option Explicit


Public Sub NaamGeven()
    
    Dim frmNaam As New FormNaamGeven
    
    frmNaam.Show
    
    Set frmNaam = Nothing

End Sub




Private Sub EnterOpmAfspr(intN As Integer)

    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue("opmAfsprBlad__" & intN, vbNullString)
    
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue "opmAfsprBlad__" & intN, frmOpmerking.txtOpmerking.Text
    End If
    
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub OpmAfsprInfusen()
    
    EnterOpmAfspr 2
    
End Sub

Public Sub opmIntakeMedIV()
    
    EnterOpmAfspr 8

End Sub

Public Sub opmOverig_1()
    
    EnterOpmAfspr 9
    
End Sub

Public Sub opmVoedPO()
    
    EnterOpmAfspr 14
    
End Sub


Public Sub Nutriflex()

    shtPedBerTPN.Range("TPNVol").Value = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))

End Sub

Private Sub EnterHoeveelheid(strItem As String)

    Dim dblVol As Double

    dblVol = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))
        
    ModRange.SetRangeValue strItem, dblVol

End Sub

Public Sub NaCL()

    EnterHoeveelheid "NaClVol"

End Sub
Public Sub KCl()

    EnterHoeveelheid "KClVol"

End Sub

Public Sub CaGlucVol()

    EnterHoeveelheid "CaGlucVol"

End Sub
Public Sub MgCl()

    EnterHoeveelheid "MgClVol"

End Sub

Public Sub PaceMaker()

    shtPedBerIVenPM.Range("PM_Standaard").Copy
    shtPedBerIVenPM.Range("PM_Instelling").PasteSpecial xlPasteValues

End Sub

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

Public Sub Med11Tekst()

    TekstInvoer "Voer tekst in ...", "Tekst voor medicatie 13", "_MedTekst_1"

End Sub

Public Sub Med12Tekst()

    TekstInvoer "Voer tekst in ...", "Tekst voor medicatie 14", "_MedTekst_2"
    
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

Public Sub OpmLabNeo()
    
    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue("LabNeoOpmerkingen", vbNullString)
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue "LabNeoOpmerkingen", frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing
    
End Sub

Public Sub MedTekstAanvullendeAfsprakenVerliezenPed()

    TekstInvoer "Voer tekst in ...", "Voer compensatie vloeistof in", "Aanvullend_Verliezen_Ped"

End Sub

Public Sub Med11Tekst1700()

    TekstInvoer "Voer tekst in ...", "Tekst voor medicatie 13", "_MedTekst1700_1"
    
End Sub

Public Sub Med12Tekst1700()

    TekstInvoer "Voer tekst in ...", "Tekst voor medicatie 14", "_MedTekst1700_2"
    
End Sub


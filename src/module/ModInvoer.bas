Attribute VB_Name = "ModInvoer"
Option Explicit
Dim intN As Integer, dblVol As Double

Public Sub VoerGewichtIn()

    Dim frmGewichtInvoer As New FormInvoerNumeriek
    Dim objPatient As New ClassPatient
    
    With frmGewichtInvoer
        .Caption = "Gewicht invoeren ..."
        .lblParameter.Caption = "Gewicht:"
        .lblEenheid = "kg"
        .txtWaarde = Range("Gewicht").Value / 10
        .Show
        If .txtWaarde.Text <> vbNullString Then
            objPatient.Gewicht = .txtWaarde.Text
            If Not IsNull(objPatient.Gewicht) Then
                Range("Gewicht") = objPatient.Gewicht * 10
                Range("_Gewicht") = CDbl(objPatient.Gewicht)
                
            End If
        End If
        .txtWaarde = vbNullString
    End With
    
    SelectTPN
    
    Set objPatient = Nothing
    Set frmGewichtInvoer = Nothing

End Sub

Public Sub VoerLengteIn()

    Dim frmLengteInvoer As New FormInvoerNumeriek
    Dim objPatient As New ClassPatient
    
    With frmLengteInvoer
        .Caption = "Lengte invoeren ..."
        .lblParameter.Caption = "Lengte:"
        .lblEenheid = "cm"
        .txtWaarde = Range("Lengte").Value
        .Show
        If .txtWaarde.Text <> vbNullString Then
            objPatient.Lengte = .txtWaarde.Text
            If Not IsNull(objPatient.Lengte) Then
                Range("Lengte") = objPatient.Lengte
            End If
        End If
        .txtWaarde = vbNullString
    End With
    
    Set objPatient = Nothing
    Set frmLengteInvoer = Nothing

End Sub

Public Sub NaamGeven()
    
    Dim frmNaam As New FormNaamGeven
    
    frmNaam.Show
    
    Set frmNaam = Nothing

End Sub

Private Sub MedicamentInvoeren(intN)

    Dim frmMedicament As New FormMedicament
    Dim strMed As String
    Dim strGeneric As String
    
    With frmMedicament
        
        If Range("RecNo_" & intN).Value > 0 Then
            .LoadGPK CStr(Range("RecNo_" & intN).Value)
        Else
            .cboGeneriek.Text = Range("Generic_" & intN).Value
            .txtSterkte = vbNullString
            .txtSterkteEenheid = vbNullString
            
            
        End If
        .txtDosisEenheid = Range("Eenheid_" & intN).Value
        .txtDosis = Range("StandDos_" & intN).Value
        .cboRoute = Range("MedToed_" & intN).Value
        .Show
        
        If .lblCancel.Caption = "OK" Then
            strMed = .lblEtiket.Caption
            strGeneric = .cboGeneriek.Text
            If strMed = vbNullString And .txtSterkte <> vbNullString Then
                strMed = strGeneric & " " & .txtSterkte & " " & .txtSterkteEenheid
            End If
            Range("MedKeuze_" & intN).Value = strMed
            Range("Generic_" & intN).Value = strGeneric
            Range("StandDos_" & intN).Value = Val(Replace(.txtDosis.Value, ",", "."))
            Range("Eenheid_" & intN).Value = .txtDosisEenheid.Text
            Range("medtoed_" & intN).Value = .cboRoute.Text
            Range("RecNo_" & intN).Value = CLng(.GetGPK())
            
        Else
            If .lblCancel.Caption = "Clear" Then
                Range("MedKeuze_" & intN).Value = vbNullString
                Range("StandDos_" & intN).Value = vbNullString
                Range("Eenheid_" & intN).Value = vbNullString
                Range("MedToed_" & intN).Value = vbNullString
                Range("OpmMedDisc__" & intN).Value = vbNullString
                Range("DosHoev_" & intN).Value = vbNullString
                Range("MedTijden_" & intN).Value = 1
                Range("MedOplVol_" & intN).Value = 0
                Range("MedOpl_" & intN).Value = 0
                Range("MedInloop_" & intN).Value = 0
                Range("RecNo_" & intN).Value = 0
            End If
        End If
    End With
    
    Set frmMedicament = Nothing

End Sub

Public Sub Medicament_16()

    MedicamentInvoeren (16)

End Sub

Public Sub Medicament_17()

    MedicamentInvoeren (17)

End Sub

Public Sub Medicament_18()

    MedicamentInvoeren (18)

End Sub

Public Sub Medicament_19()

    MedicamentInvoeren (19)

End Sub

Public Sub Medicament_15()

    MedicamentInvoeren (15)

End Sub

Public Sub Medicament_14()

    MedicamentInvoeren (14)

End Sub

Public Sub Medicament_13()

    MedicamentInvoeren (13)

End Sub

Public Sub Medicament_12()

    MedicamentInvoeren (12)

End Sub

Public Sub Medicament_11()

    MedicamentInvoeren (11)

End Sub

Public Sub Medicament_10()

    MedicamentInvoeren (10)

End Sub

Public Sub Medicament_9()

    MedicamentInvoeren (9)

End Sub

Public Sub Medicament_8()

    MedicamentInvoeren (8)

End Sub

Public Sub Medicament_7()

    MedicamentInvoeren (7)

End Sub

Public Sub Medicament_6()

    MedicamentInvoeren (6)

End Sub

Public Sub Medicament_5()

    MedicamentInvoeren (5)

End Sub

Public Sub Medicament_4()

    MedicamentInvoeren (4)

End Sub

Public Sub Medicament_3()

    MedicamentInvoeren (3)

End Sub

Public Sub Medicament_2()

    MedicamentInvoeren (2)

End Sub

Public Sub Medicament_1()

    MedicamentInvoeren (1)

End Sub


Public Sub Medicament_20()

    MedicamentInvoeren (20)

End Sub

Public Sub Medicament_21()

    MedicamentInvoeren (21)

End Sub

Public Sub Medicament_22()

    MedicamentInvoeren (22)

End Sub

Public Sub Medicament_23()

    MedicamentInvoeren (23)

End Sub

Public Sub Medicament_24()

    MedicamentInvoeren (24)

End Sub

Public Sub Medicament_25()

    MedicamentInvoeren (25)

End Sub

Public Sub Medicament_26()

    MedicamentInvoeren (26)

End Sub

Public Sub Medicament_27()

    MedicamentInvoeren (27)

End Sub

Public Sub Medicament_28()

    MedicamentInvoeren (28)

End Sub

Public Sub Medicament_29()

    MedicamentInvoeren (29)

End Sub

Public Sub Medicament_30()

    MedicamentInvoeren (30)

End Sub

Private Sub MedIVInvoer(intN As Integer)

    Dim strMed As String, strSterkte As String
    Dim frmMedIV As New FormMedIV
    
    frmMedIV.Show
    
    strMed = frmMedIV.txtMedicament.Text
    strSterkte = frmMedIV.txtSterkte.Text
    Range("MedIVKeuze_" & intN).Value = strMed
    Range("MedIVSterkte_" & intN).Value = strSterkte
    
    Set frmMedIV = Nothing
        
End Sub

Public Sub MedIV_11()

    MedIVInvoer (11)
        
End Sub

Public Sub MedIV_12()
    
    MedIVInvoer (12)

End Sub

Public Sub MedIV_13()
    
    MedIVInvoer (13)

End Sub

Public Sub MedIV_14()
    
    MedIVInvoer (14)

End Sub

Public Sub MedIV_15()
    
    MedIVInvoer (15)

End Sub

Private Sub EnterOpmAfspr(intN As Integer)

    Dim strN As String
    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__" & intN).Value
    
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__" & intN).Value = frmOpmerking.txtOpmerking.Text
    End If
    
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub OpmAfsprInfusen()
    
    EnterOpmAfspr (2)
    
End Sub

Public Sub OpmAfsprMedIV_1()
    
    EnterOpmAfspr (3)

End Sub

Public Sub OpmAfsprMedIV_2()
    
    EnterOpmAfspr (4)

End Sub

Public Sub OpmAfsprMedIV_3()
    
    EnterOpmAfspr (5)

End Sub

Public Sub OpmAfsprMedIV_4()
    
    EnterOpmAfspr (6)
    
End Sub

Public Sub OpmAfsprMedIV_5()
    
    EnterOpmAfspr (7)

End Sub

Public Sub opmIntakeMedIV()
    
    EnterOpmAfspr (8)

End Sub

Public Sub opmOverig_1()
    
    EnterOpmAfspr (9)
    
End Sub

Public Sub opmOverig_2()
    
    EnterOpmAfspr (10)

End Sub

Public Sub opmOverig_3()
    
    EnterOpmAfspr (11)
    
End Sub

Public Sub opmOverig_4()
    
    EnterOpmAfspr (12)
    
End Sub

Public Sub opmOverig_5()
    
    EnterOpmAfspr (13)
    
End Sub

Public Sub opmVoedPO()
    
    EnterOpmAfspr (14)
    
End Sub

Public Sub opmOverig_6()
    
    EnterOpmAfspr (15)
    
End Sub

Private Sub OpmMedDisc(intN As Integer)
    
    Dim frmOpmerking As New FormOpmerking
    Dim strRange As String
    
    strRange = shtGlobBerOpm.Name & "!c" & intN

    frmOpmerking.txtOpmerking.Text = Range(strRange).Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range(strRange).Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub

Public Sub OpmMedDisc_1()
    
    OpmMedDisc (16)
    
End Sub

Public Sub OpmMedDisc_2()
    
    OpmMedDisc (17)

End Sub

Public Sub OpmMedDisc_3()
    
    OpmMedDisc (18)

End Sub
Public Sub OpmMedDisc_4()
    
    OpmMedDisc (19)

End Sub

Public Sub OpmMedDisc_5()
    
    OpmMedDisc (20)

End Sub
Public Sub OpmMedDisc_6()
    
    OpmMedDisc (21)

End Sub
Public Sub OpmMedDisc_7()
    
    OpmMedDisc (22)

End Sub
Public Sub OpmMedDisc_8()
    
    OpmMedDisc (23)

End Sub
Public Sub OpmMedDisc_9()
    
    OpmMedDisc (24)

End Sub
Public Sub OpmMedDisc_10()
    
    OpmMedDisc (25)

End Sub
Public Sub OpmMedDisc_11()
    
    OpmMedDisc (26)

End Sub
Public Sub OpmMedDisc_12()
    
    OpmMedDisc (27)

End Sub
Public Sub OpmMedDisc_13()
    
    OpmMedDisc (28)

End Sub
Public Sub OpmMedDisc_14()
    
    OpmMedDisc (29)

End Sub
Public Sub OpmMedDisc_15()
    
    OpmMedDisc (30)

End Sub
Public Sub OpmMedDisc_16()
    
    OpmMedDisc (31)

End Sub
Public Sub OpmMedDisc_17()
    
    OpmMedDisc (32)

End Sub
Public Sub OpmMedDisc_18()
    
    OpmMedDisc (33)

End Sub
Public Sub OpmMedDisc_19()
    
    OpmMedDisc (34)

End Sub
Public Sub OpmMedDisc_20()
    
    OpmMedDisc (35)

End Sub

Public Sub OpmMedDisc_21()
    
    OpmMedDisc (36)

End Sub

Public Sub OpmMedDisc_22()
    
    OpmMedDisc (37)

End Sub

Public Sub OpmMedDisc_23()
    
    OpmMedDisc (38)

End Sub

Public Sub OpmMedDisc_24()
    
    OpmMedDisc (39)

End Sub

Public Sub OpmMedDisc_25()
    
    OpmMedDisc (40)

End Sub

Public Sub OpmMedDisc_26()
    
    OpmMedDisc (41)

End Sub

Public Sub OpmMedDisc_27()
    
    OpmMedDisc (42)

End Sub

Public Sub OpmMedDisc_28()
    
    OpmMedDisc (43)

End Sub

Public Sub OpmMedDisc_29()
    
    OpmMedDisc (44)

End Sub

Public Sub OpmMedDisc_30()
    
    OpmMedDisc (45)

End Sub

Public Sub Nutriflex()

    shtPedBerTPN.Range("TPNVol").Value = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))

End Sub

Private Sub EnterHoeveelheid(strItem As String)

    Dim dblVol As Double

    dblVol = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))
        
    shtPedBerTPN.Range(strItem).Value = dblVol

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
        .Tekst = Range(strRange).Value
        .Show
        If .IsOK Then Range(strRange) = .Tekst
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
    
    frmOpmerking.txtOpmerking.Text = Range("LabNeoOpmerkingen").Formula
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("LabNeoOpmerkingen").Value = frmOpmerking.txtOpmerking.Text
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


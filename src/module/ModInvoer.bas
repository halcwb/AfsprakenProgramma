Attribute VB_Name = "ModInvoer"
Option Explicit
Dim intN As Integer, dblVol As Double

Public Sub InGewicht()

    Dim frmGewichtInvoer As New frmInvoerNumeriek
    Dim objPatient As New clsPatient
    
    With frmInvoerNumeriek
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

Public Sub InLengte()

    Dim frmLengteInvoer As New frmInvoerNumeriek
    Dim objPatient As New clsPatient
    
    With frmInvoerNumeriek
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
    
    frmNaamGeven.Show

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

Public Sub MedicamentInvoeren(intN)
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
                Range("opmMedDisc__" & intN).Value = vbNullString
                Range("DosHoev_" & intN).Value = vbNullString
                Range("MedTijden_" & intN).Value = 1
                Range("MedOplVol_" & intN).Value = 0
                Range("MedOpl_" & intN).Value = 0
                Range("MedInloop_" & intN).Value = 0
                Range("RecNo_" & intN).Value = 0
            End If
        End If
    End With

End Sub

Public Sub MedIV_11()

        Dim strMed As String, strSterkte As String
        frmMedIV.Show
        strMed = frmMedIV.txtMedicament.Text
        strSterkte = frmMedIV.txtSterkte.Text
        Range("MedIVKeuze_11").Value = strMed
        Range("MedIVSterkte_11").Value = strSterkte
        
End Sub

Public Sub MedIV_12()
    
    Dim strMed As String, strSterkte As String
    frmMedIV.Show
    strMed = frmMedIV.txtMedicament.Text
    strSterkte = frmMedIV.txtSterkte.Text
    Range("MedIVKeuze_12").Value = strMed
    Range("MedIVSterkte_12").Value = strSterkte

End Sub

Public Sub MedIV_13()
    
    Dim strMed As String, strSterkte As String
    frmMedIV.Show
    strMed = frmMedIV.txtMedicament.Text
    strSterkte = frmMedIV.txtSterkte.Text
    Range("MedIVKeuze_13").Value = strMed
    Range("MedIVSterkte_13").Value = strSterkte

End Sub

Public Sub MedIV_14()
    
    Dim strMed As String, strSterkte As String
    frmMedIV.Show
    strMed = frmMedIV.txtMedicament.Text
    strSterkte = frmMedIV.txtSterkte.Text
    Range("MedIVKeuze_14").Value = strMed
    Range("MedIVSterkte_14").Value = strSterkte

End Sub

Public Sub MedIV_15()
    
    Dim strMed As String, strSterkte As String
    frmMedIV.Show
    strMed = frmMedIV.txtMedicament.Text
    strSterkte = frmMedIV.txtSterkte.Text
    Range("MedIVKeuze_15").Value = strMed
    Range("MedIVSterkte_15").Value = strSterkte

End Sub

Public Sub opmAfsprBeademing()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__1").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__1").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmAfsprInfusen()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__2").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__2").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmAfsprMedIV_1()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__3").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__3").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmAfsprMedIV_2()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__4").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__4").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmAfsprMedIV_3()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__5").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__5").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmAfsprMedIV_4()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__6").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__6").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmAfsprMedIV_5()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__7").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__7").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmIntakeMedIV()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__8").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__8").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmOverig_1()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__9").Formula
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__9").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub
Public Sub opmOverig_2()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__10").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__10").Value = frmOpmerking.txtOpmerking.Text
        
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmOverig_3()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__11").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__11").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub
Public Sub opmOverig_4()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__12").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__12").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub
Public Sub opmOverig_5()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__13").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__13").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub
Public Sub opmOverig_6()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__15").Formula
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__15").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub
Public Sub opmVoedPO()
    
    frmOpmerking.txtOpmerking.Text = Range("opmAfsprBlad__14").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("opmAfsprBlad__14").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub

Public Sub medDiscBijz()
    
    With Sheets("ber_med_disc")
'        frmMedDiscBijz.txtMedDiscBijz.text = Sheets("dbform").Cells(.Cells(25, 2).Value + 1, 9).Value
        frmMedDiscBijz.Show
        If frmMedDiscBijz.txtMedDiscBijz.Text <> vbNullString Then
            Sheets("dbform").Cells(.Cells(25, 2).Value + 1, 9).Formula = _
            frmMedDiscBijz.txtMedDiscBijz.Text
        End If
    End With

End Sub

Public Sub medDiscDos()
    
    With Sheets("ber_med_disc")
'        frmMedDiscDos.txtMedDiscDos.text = Sheets("dbform").Cells(.Cells(25, 2).Value + 1, 8).Value
        frmMedDiscDos.Show
        If frmMedDiscDos.txtMedDiscDos.Text <> vbNullString Then
            Sheets("dbform").Cells(.Cells(25, 2).Value + 1, 8).Formula = _
            frmMedDiscDos.txtMedDiscDos.Text
        End If
    End With

End Sub

Public Sub opmMedDisc_1()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c16").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c16").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_2()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c17").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c17").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_3()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c18").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c18").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_4()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c19").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c19").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_5()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c20").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c20").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_6()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c21").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c21").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_7()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c22").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c22").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_8()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c23").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c23").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_9()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c24").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c24").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_10()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c25").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c25").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_11()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c26").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c26").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_12()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c27").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c27").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_13()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c28").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c28").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_14()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c29").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c29").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_15()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c30").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c30").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_16()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c31").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c31").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_17()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c32").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c32").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_18()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c33").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c33").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_19()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c34").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c34").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub
Public Sub opmMedDisc_20()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c35").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c35").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_21()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c37").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c37").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_22()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c38").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c38").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_23()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c39").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c39").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_24()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c40").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c40").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_25()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c41").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c41").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_26()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c42").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c42").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_27()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c43").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c43").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_28()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c44").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c44").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_29()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c45").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c45").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub opmMedDisc_30()
    
    frmOpmerking.txtOpmerking.Text = Range("ber_opm!c46").Value
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("Ber_Opm!c46").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString

End Sub

Public Sub medDiscToeding()
    
    With Sheets("ber_med_disc")
        frmMedDiscToediening.txtMedDiscToediening.Text = _
        Sheets("dbform").Cells(.Cells(25, 2).Value + 1, 7).Value
        frmMedDiscToediening.Show
        If frmMedDiscToediening.txtMedDiscToediening.Text <> vbNullString Then
            Sheets("dbform").Cells(.Cells(25, 2).Value + 1, 7).Formula _
            = frmMedDiscToediening.txtMedDiscToediening.Text
        End If
    End With

End Sub

Public Sub Nutriflex()

    shtBerTPN.Range("TPNVol").Value = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))

End Sub
Public Sub NaCL()

    dblVol = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))
    If dblVol < 5 Then
        shtBerTPN.Range("NaClVol").Value = dblVol
    Else
        shtBerTPN.Range("NaClVol").Value = dblVol
    End If

End Sub
Public Sub KCl()

    dblVol = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))
    If dblVol < 5 Then
        shtBerTPN.Range("KClVol").Value = dblVol
    Else
        shtBerTPN.Range("KClVol").Value = dblVol
    End If

End Sub

Public Sub CaGlucVol()

    dblVol = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))
    If dblVol < 5 Then
        shtBerTPN.Range("CaGlucVol").Value = dblVol
    Else
        shtBerTPN.Range("CaGlucVol").Value = dblVol
    End If

End Sub
Public Sub MgCl()

    dblVol = Val(InputBox(prompt:="Voer de hoeveelheid in ...", _
    Title:="Informedica 2000"))
    If dblVol < 5 Then
        shtBerTPN.Range("MgClVol").Value = dblVol
    Else
        shtBerTPN.Range("MgClVol").Value = dblVol
    End If

End Sub

Public Sub PaceMaker()

    shtBerInfusen.Range("PM_Standaard").Copy
    shtBerInfusen.Range("PM_Instelling").PasteSpecial xlPasteValues

End Sub

'Public Sub Med11Tekst()
'    frmOpmerking.Caption = "Voer tekst in voor medicatie 13"
'    frmOpmerking.txtOpmerking.Text = Range("_MedTekst_1").Formula
'    frmOpmerking.Show
'    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
'        Range("_MedTekst_1").Value = frmOpmerking.txtOpmerking.Text
'    End If
'    frmOpmerking.txtOpmerking.Text = ""
'End Sub
'
'Public Sub Med12Tekst()
'    frmOpmerking.Caption = "Voer tekst in voor medicatie 14"
'    frmOpmerking.txtOpmerking.Text = Range("_MedTekst_2").Formula
'    frmOpmerking.Show
'    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
'        Range("_MedTekst_2").Value = frmOpmerking.txtOpmerking.Text
'    End If
'    frmOpmerking.txtOpmerking.Text = ""
'End Sub

Public Sub Med11Tekst()

    Dim frmInvoer As New frmTekstInvoer
    
    With frmInvoer
        .Caption = "Voer tekst in ..."
        .lblNaam.Caption = "Tekst voor medicatie 13"
        .Tekst = Range("_MedTekst_1").Value
        .Show
        If .IsOK Then Range("_MedTekst_1") = .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med12Tekst()

    Dim frmInvoer As New frmTekstInvoer
    
    With frmInvoer
        .Caption = "Voer tekst in ..."
        .lblNaam.Caption = "Tekst voor medicatie 14"
        .Tekst = Range("_MedTekst_2").Value
        .Show
        If .IsOK Then Range("_MedTekst_2") = .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub MedTekstWondkweek()

    Dim frmInvoer As New frmTekstInvoer
    
    With frmInvoer
        .Caption = "Voer tekst in"
        .lblNaam.Caption = "Voer locatie wond(en) in"
        .Tekst = Range("Aanvullend_WondkweekTekst").Value
        .Show
        If .IsOK Then Range("Aanvullend_WondkweekTekst") = .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub MedTekstVerliezenCompenseren()

    Dim frmInvoer As New frmTekstInvoer
    
    With frmInvoer
        .Caption = "Voer tekst in"
        .lblNaam.Caption = "Voer verliezen compenseren in"
        .Tekst = Range("Aanvullend_VerliezenTekst").Value
        .Show
        If .IsOK Then Range("Aanvullend_VerliezenTekst") = .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub MedTekstAanvullendeAfsprakenOverigePed()

    Dim frmInvoer As New frmTekstInvoer
    
    With frmInvoer
        .Caption = "Voer tekst in"
        .lblNaam.Caption = "Voer overige aanvullende afspraken in"
        .Tekst = Range("Aanvullend_Overige_Ped").Value
        .Show
        If .IsOK Then Range("Aanvullend_Overige_Ped") = .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub opmLabNeo()
    
    frmOpmerking.txtOpmerking.Text = Range("LabNeoOpmerkingen").Formula
    frmOpmerking.Show
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        Range("LabNeoOpmerkingen").Value = frmOpmerking.txtOpmerking.Text
    End If
    frmOpmerking.txtOpmerking.Text = vbNullString
    
End Sub

Public Sub MedTekstAanvullendeAfsprakenVerliezenPed()

    Dim frmInvoer As New frmTekstInvoer
    
    With frmInvoer
        .Caption = "Voer tekst in"
        .lblNaam.Caption = "Voer compensatie vloeistof in"
        .Tekst = Range("Aanvullend_Verliezen_Ped").Value
        .Show
        If .IsOK Then Range("Aanvullend_Verliezen_Ped") = .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med11Tekst1700()

    Dim frmInvoer As New frmTekstInvoer
    
    With frmInvoer
        .Caption = "Voer tekst in ..."
        .lblNaam.Caption = "Tekst voor medicatie 13"
        .Tekst = Range("_MedTekst1700_1").Value
        .Show
        If .IsOK Then Range("_MedTekst1700_1") = .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub

Public Sub Med12Tekst1700()

    Dim frmInvoer As New frmTekstInvoer
    
    With frmInvoer
        .Caption = "Voer tekst in ..."
        .lblNaam.Caption = "Tekst voor medicatie 14"
        .Tekst = Range("_MedTekst1700_2").Value
        .Show
        If .IsOK Then Range("_MedTekst1700_2") = .Tekst
    End With
    
    Set frmInvoer = Nothing

End Sub


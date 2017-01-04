Attribute VB_Name = "ModPedEntTPN"
Option Explicit

Private Const constEntText = "_Ped_Ent_Opm"
Private Const constTpnText = "_Ped_TPN_Opm"
Private Const constNaCl1 = "_Ped_TPN_NaCl1"
Private Const constKCl1 = "_Ped_TPN_KCl1"
Private Const constNaCl2 = "_Ped_TPN_NaCl2"
Private Const constKCl2 = "_Ped_TPN_KCl2"
Private Const constCaGluc = "_Ped_TPN_CaCl"
Private Const constMgCl = "_Ped_TPN_MgCl"
Private Const constKNaFosf = "_Ped_TPN_KNaFosf"
Private Const constNaCl1Vol = "_Ped_TPN_NaClVol1"
Private Const constKCl1Vol = "_Ped_TPN_KClVol1"
Private Const constNaCl2Vol = "_Ped_TPN_NaClVol2"
Private Const constKCl2Vol = "_Ped_TPN_KClVol2"
Private Const constCaGlucVol = "_Ped_TPN_CaGlucVol"
Private Const constMgClVol = "_Ped_TPN_MgClVol"
Private Const constTPNVol = "_Ped_TPN_Vol"

Public Sub PedEntTPN_SelectTPNPrint()

    Dim dblGew As Double
    
    dblGew = Val(ModRange.GetRangeValue(ModConst.CONST_RANGE_GEWICHT, 0)) / 10

    If dblGew >= CONST_TPN_1 And dblGew <= CONST_TPN_2 Then
        shtPedPrtTPN2tot6.Select
    Else
        If dblGew <= CONST_TPN_3 Then
            shtPedPrtTPN7tot15.Select
        Else
            If dblGew <= CONST_TPN_4 Then
                shtPedPrtTPN16tot30.Select
            Else
                If dblGew <= CONST_TPN_5 Then
                    shtPedPrtTPN31tot50.Select
                Else
                    shtPedPrtTPN50.Select
                End If
            End If
        End If
    End If
        
    Range("A1").Select

End Sub

Public Sub PedEntTPN_SelectTPN()

    Dim dblGewicht As Double
    Dim strTPNB As String
    Dim strTPNC As String
    Dim strTPND As String
    Dim strTPNE As String
    Dim strTPNnutri As String
    Dim strSelected As String
    
    strTPNB = "tbl_Ped_tpnB"
    strTPNC = "tbl_Ped_tpnC"
    strTPND = "tbl_Ped_tpnD"
    strTPNE = "tbl_Ped_tpnE"
    strTPNnutri = "tbl_Ped_tpnNutriflex"
    strSelected = "tbl_Ped_tpnSelected"

    dblGewicht = ModPatient.GetGewichtFromRange()

    With shtPedBerTPN
        If dblGewicht >= 2 And dblGewicht < 7 Then
            .Range(strTPNB).Copy
        ElseIf dblGewicht >= 7 And dblGewicht < 15 Then
            .Range(strTPNC).Copy
        ElseIf dblGewicht >= 15 And dblGewicht < 30 Then
            .Range(strTPND).Copy
        ElseIf dblGewicht >= 30 And dblGewicht <= 50 Then
            .Range(strTPNE).Copy
        ElseIf dblGewicht > 50 Then
            .Range(strTPNnutri).Copy
        End If
        .Range(strSelected).PasteSpecial xlPasteValues
    End With
    
    Application.Calculate
    
End Sub

Private Sub TPNAdvies(Dag As Integer)

    Dim dblVol As Double
    Dim dblNaCl As Double
    Dim dblKCl As Double
    Dim dblVitIntra As Double
    Dim dblLipid As Double
    Dim dblSolu As Double
    Dim dblGewicht As Double
    
    dblGewicht = Val(ModRange.GetRangeValue(ModConst.CONST_RANGE_GEWICHT, 0)) / 10

    Select Case dblGewicht
        Case 2 To 6
            ModRange.SetRangeValue "TPN", 2
            
            ModRange.SetRangeValue "NaCl", True
            dblNaCl = 6 * dblGewicht
            ModRange.SetRangeValue "NaClVol", dblNaCl
            
            ModRange.SetRangeValue "KCl", True
            dblKCl = 1 * dblGewicht
            ModRange.SetRangeValue "KClVol", dblKCl
            
            dblVitIntra = dblGewicht
            ModRange.SetRangeValue "VitIntra", True
            ModRange.SetRangeValue "ModSetting", IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            Select Case Dag
                Case 1
                    ModRange.SetRangeValue "SSTglucose", 2
                
                    dblKCl = 1.5 * dblGewicht
                    ModRange.SetRangeValue "KClVol", dblKCl
                    ModRange.SetRangeValue "TPNVol", 15 * dblGewicht
                
                    dblLipid = 6 * dblGewicht / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                    ModRange.SetRangeValue "SSTglucose", 3
                
                    ModRange.SetRangeValue "TPNVol", 25 * dblGewicht
                
                    dblLipid = 11 * dblGewicht / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                    ModRange.SetRangeValue "SSTglucose", 5
                
                    ModRange.SetRangeValue "TPNVol", 35 * dblGewicht
            
                    dblLipid = 16 * dblGewicht / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            
            End Select
            
            dblVol = (150 * dblGewicht - ModRange.GetRangeValue("TPNVol", 0) * 2 - dblNaCl * 2 - dblKCl * 2 - dblLipid * 24) / 24
            
            If dblVol < 5 Then
                ModRange.SetRangeValue "SSTstand", dblVol * 10
            ElseIf dblVol < 146 Then
                ModRange.SetRangeValue "SSTstand", dblVol + 45
            Else
                ModRange.SetRangeValue "SSTstand", (dblVol + 125) / 5
            End If
            
        Case 7 To 15
            ModRange.SetRangeValue "TPN", 2
            
            ModRange.SetRangeValue "NaCl", True
            dblNaCl = 6 * dblGewicht
            ModRange.SetRangeValue "NaClVol", dblNaCl
            
            ModRange.SetRangeValue "KCl", True
            dblKCl = 1.5 * dblGewicht
            ModRange.SetRangeValue "KClVol", dblKCl
            
            dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
            ModRange.SetRangeValue "VitIntra", True
            ModRange.SetRangeValue "VitIntraVol", IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            ModRange.SetRangeValue "SoluVit", True
            ModRange.SetRangeValue "SoluVitVol", IIf(dblSolu < 5, dblSolu * 10, dblSolu + 45)
            
            Select Case Dag
                Case 1
                    ModRange.SetRangeValue "SSTglucose", 2
                
                    dblKCl = 2 * dblGewicht
                    ModRange.SetRangeValue "KClVol", dblKCl
                    ModRange.SetRangeValue "TPNVol", 10 * dblGewicht
                
                    dblLipid = (5 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                    ModRange.SetRangeValue "SSTglucose", 6
                
                    ModRange.SetRangeValue "TPNVol", 20 * dblGewicht
                
                    dblLipid = (10 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                    ModRange.SetRangeValue "SSTglucose", 8
                
                    ModRange.SetRangeValue "TPNVol", 25 * dblGewicht
            
                    dblLipid = (15 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            End Select
            
            dblVol = (90 * dblGewicht + ((15 - dblGewicht) / 8) * 20 * dblGewicht - ModRange.GetRangeValue("TPNVol", 0) * 2 - dblNaCl * 2 - dblKCl * 2 - dblLipid * 24) / 24
            
            If dblVol < 5 Then
                ModRange.SetRangeValue "SSTstand", dblVol * 10
            ElseIf dblVol < 146 Then
                ModRange.SetRangeValue "SSTstand", dblVol + 45
            Else
                ModRange.SetRangeValue "SSTstand", (dblVol + 125) / 5
            End If
    
        Case 16 To 30
            ModRange.SetRangeValue "TPN", 2
            
            ModRange.SetRangeValue "NaCl", True
            dblNaCl = 6 * dblGewicht
            ModRange.SetRangeValue "NaClVol", dblNaCl
            
            ModRange.SetRangeValue "KCl", True
            dblKCl = 1.5 * dblGewicht
            ModRange.SetRangeValue "KClVol", dblKCl
            
            dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
            ModRange.SetRangeValue "VitIntra", True
            ModRange.SetRangeValue "VitIntraVol", IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            ModRange.SetRangeValue "SoluVit", True
            ModRange.SetRangeValue "SoluVitVol", IIf(dblSolu < 5, dblSolu * 10, dblSolu + 45)
            
            ModRange.SetRangeValue "Peditrace", 15
            
            Select Case Dag
                Case 1
                    ModRange.SetRangeValue "SSTglucose", 2
                
                    dblKCl = 2 * dblGewicht
                    ModRange.SetRangeValue "KClVol", dblKCl
                    ModRange.SetRangeValue "TPNVol", 10 * dblGewicht
                
                    dblLipid = (5 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                    ModRange.SetRangeValue "SSTglucose", 6
                
                    ModRange.SetRangeValue "TPNVol", 15 * dblGewicht
                
                    dblLipid = (10 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                    ModRange.SetRangeValue "SSTglucose", 8
                
                    ModRange.SetRangeValue "TPNVol", 20 * dblGewicht
            
                    dblLipid = (15 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            End Select
            
            dblVol = (70 * dblGewicht + ((30 - dblGewicht) / 14) * 10 * dblGewicht - 15 - ModRange.GetRangeValue("TPNVol", 0) * 2 - dblNaCl * 2 - dblKCl * 2 - dblLipid * 24) / 24
            
            If dblVol < 5 Then
                ModRange.SetRangeValue "SSTstand", dblVol * 10
            ElseIf dblVol < 146 Then
                ModRange.SetRangeValue "SSTstand", dblVol + 45
            Else
                ModRange.SetRangeValue "SSTstand", (dblVol + 125) / 5
            End If
        
        Case 31 To 50
            ModRange.SetRangeValue "TPN", 2
            
            ModRange.SetRangeValue "NaCl", True
            dblNaCl = 6 * dblGewicht
            ModRange.SetRangeValue "NaClVol", dblNaCl
            
            ModRange.SetRangeValue "KCl", True
            dblKCl = 1.5 * dblGewicht
            ModRange.SetRangeValue "KClVol", dblKCl
            
            dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
            ModRange.SetRangeValue "VitIntra", True
            ModRange.SetRangeValue "VitIntraVol", IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            ModRange.SetRangeValue "SoluVit", True
            ModRange.SetRangeValue "SoluVitVol", IIf(dblSolu < 5, dblSolu * 10, dblSolu + 45)
            
            ModRange.SetRangeValue "Peditrace", 15
            
            Select Case Dag
                Case 1
                    ModRange.SetRangeValue "SSTglucose", 2
                
                    dblKCl = 2 * dblGewicht
                    ModRange.SetRangeValue "KClVol", dblKCl
                    ModRange.SetRangeValue "TPNVol", 5 * dblGewicht
                
                    dblLipid = (3 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                    ModRange.SetRangeValue "SSTglucose", 6
                
                    ModRange.SetRangeValue "TPNVol", 8 * dblGewicht
                
                    dblLipid = (6 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                    ModRange.SetRangeValue "SSTglucose", IIf(dblGewicht > 35, 9, 7)
                
                    ModRange.SetRangeValue "TPNVol", 12 * dblGewicht
            
                    dblLipid = (10 * dblGewicht + dblVitIntra + dblSolu) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            End Select
            
            dblVol = (50 * dblGewicht + ((50 - dblGewicht) / 19) * 20 * dblGewicht - 15 - ModRange.GetRangeValue("TPNVol", 0) * 2 - dblNaCl * 2 - dblKCl * 2 - dblLipid * 24) / 24
            
            If dblVol < 5 Then
                ModRange.SetRangeValue "SSTstand", dblVol * 10
            ElseIf dblVol < 146 Then
                ModRange.SetRangeValue "SSTstand", dblVol + 45
            Else
                ModRange.SetRangeValue "SSTstand", (dblVol + 125) / 5
            End If
        Case Else
            ModRange.SetRangeValue "TPN", 2
            
            ModRange.SetRangeValue "Nacl", False
            ModRange.SetRangeValue "KCl", False
            
            dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
            ModRange.SetRangeValue "VitIntra", True
            ModRange.SetRangeValue "VitIntraVol", IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            ModRange.SetRangeValue "SoluVit", True
            ModRange.SetRangeValue "SoluVitVol", IIf(dblSolu < 5, dblSolu * 10, dblSolu + 45)
            
            ModRange.SetRangeValue "Peditrace", 15
            ModRange.SetRangeValue "SSTGlucose", 2
            
            Select Case Dag
                Case 1
                
                    ModRange.SetRangeValue "TPNVol", 700
                
                    dblLipid = (150 + 20) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                
                    ModRange.SetRangeValue "TPNVol", 1000
                
                    dblLipid = (300 + 20) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                
                    ModRange.SetRangeValue "TPNVol", 1500
            
                    dblLipid = (500 + 20) / 24
                    ModRange.SetRangeValue "LipidenStand", IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            End Select
            
            dblVol = 0
            
            If dblVol < 5 Then
                ModRange.SetRangeValue "SSTstand", dblVol * 10
            ElseIf dblVol < 146 Then
                ModRange.SetRangeValue "SSTstand", dblVol + 45
            Else
                ModRange.SetRangeValue "SSTstand", (dblVol + 125) / 5
            End If
                
    End Select
    
End Sub

Public Sub PedEntTPN_TPNAdviceDayOne()

    TPNAdvies 1

End Sub

Public Sub PedEntTPN_TPNAdviceDayTwo()

    TPNAdvies 2

End Sub

Public Sub PedEntTPN_TPNAdviceDayThree()

    TPNAdvies 3

End Sub

Private Sub EnterOpmAfspr(strRange As String)

    Dim frmOpmerking As New FormOpmerking
    
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(strRange, vbNullString)
    
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue strRange, frmOpmerking.txtOpmerking.Text
    End If
    
    frmOpmerking.txtOpmerking.Text = vbNullString
    
    Set frmOpmerking = Nothing

End Sub


Public Sub PedEntTPN_EntText()
    
    EnterOpmAfspr constEntText
    
End Sub

Public Sub PedEntTPN_TPNText()
    
    EnterOpmAfspr constTpnText

End Sub

Private Sub EnterHoeveelheid(strRange As String, strItem As String)

    Dim frmInvoer As New FormInvoerNumeriek
    
    frmInvoer.SetValue strRange, strItem, ModRange.GetRangeValue(strRange, 0), "mL"
    frmInvoer.Show
    
    Set frmInvoer = Nothing

End Sub

Public Sub PedEntTPN_TPN()

    EnterHoeveelheid constTPNVol, "TPN"

End Sub

Public Sub PedEntTPN_NaCL1()

    EnterHoeveelheid constNaCl1Vol, "NaCl"
    ModRange.SetRangeValue constNaCl1, True
    
End Sub

Public Sub PedEntTPN_KCl1()

    EnterHoeveelheid constKCl1Vol, "KCl"
    ModRange.SetRangeValue constKCl1, True

End Sub

Public Sub PedEntTPN_NaCL2()

    EnterHoeveelheid constNaCl2Vol, "NaCl"
    ModRange.SetRangeValue constNaCl2, True
    
End Sub

Public Sub PedEntTPN_KCl2()

    EnterHoeveelheid constKCl2Vol, "KCl"
    ModRange.SetRangeValue constKCl2, True

End Sub

Public Sub PedEntTPN_CaGlucVol()

    EnterHoeveelheid constCaGlucVol, "CaGluc"
    ModRange.SetRangeValue constCaGluc, True

End Sub

Public Sub PedEntTPN_MgCl()

    EnterHoeveelheid constMgClVol, "MgCl"
    ModRange.SetRangeValue constMgCl, True

End Sub

Public Sub PedEntTPN_ChangeNaCl1()

    ModRange.SetRangeValue constNaCl1, True

End Sub

Public Sub PedEntTPN_ChangeKCl1()

    ModRange.SetRangeValue constKCl1, True

End Sub

Public Sub PedEntTPN_ChangeNaCl2()

    ModRange.SetRangeValue constNaCl2, True

End Sub

Public Sub PedEntTPN_ChangeKCl2()

    ModRange.SetRangeValue constKCl2, True

End Sub

Public Sub PedEntTPN_ChangeCaGluc()

    ModRange.SetRangeValue constCaGluc, True

End Sub

Public Sub PedEntTPN_ChangeMgCl()

    ModRange.SetRangeValue constMgCl, True

End Sub

Public Sub PedEntTPN_ChangeKNaFosf()

    ModRange.SetRangeValue constKNaFosf, True

End Sub

Public Sub PedEntTPN_SpecVoed()
    
    Dim frmSpecialeVoeding As New FormSpecialeVoeding

    frmSpecialeVoeding.Show
    
    Set frmSpecialeVoeding = Nothing

End Sub

Public Sub PedEntTPN_ChangeEnt()

    If Range("_Ped_Ent_Keuze_1").Value = 1 Then
    
        ModRange.SetRangeValue "_Ped_Ent_Keuze_2", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Keuze_3", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Keuze_4", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Freq_1", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Freq_2", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Freq_3", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Freq_4", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_1", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_2", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_3", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_4", vbNullString
        
    End If

End Sub

Public Sub PedEntTPN_ChangeEntAdd1()

    If ModRange.GetRangeValue("_Ped_Ent_Keuze_2", 0) = 1 Then
        ModRange.SetRangeValue "_Ped_Ent_Freq_2", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_2", vbNullString
    End If

End Sub

Public Sub PedEntTPN_ChangeEntAdd2()
    
    If ModRange.GetRangeValue("_Ped_Ent_Keuze_3", 0) = 1 Then
        ModRange.SetRangeValue "_Ped_Ent_Freq_3", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_3", vbNullString
    End If

End Sub

Public Sub PedEntTPN_ChangeEntAdd3()
    
    If ModRange.GetRangeValue("_Ped_Ent_Keuze_4", 0) = 1 Then
        ModRange.SetRangeValue "_Ped_Ent_Freq_4", vbNullString
        ModRange.SetRangeValue "_Ped_Ent_Vol_4", vbNullString
    End If

End Sub

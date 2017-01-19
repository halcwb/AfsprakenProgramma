Attribute VB_Name = "ModPedEntTPN"
Option Explicit

Private Const CONST_TPN_1 As Integer = 2
Private Const CONST_TPN_2 As Integer = 7
Private Const CONST_TPN_3 As Integer = 16
Private Const CONST_TPN_4 As Integer = 30
Private Const CONST_TPN_5 As Integer = 50

Private Const constTblVoeding As String = "Tbl_Ped_Voeding"
Private Const constTblToevoeging As String = "Tbl_Ped_Poeder"

Private Const constVoedingCount As Integer = 1
Private Const constToevoegingCount As Integer = 3

Private Const constSode As String = "_Ped_Ent_Sonde"
Private Const constVoeding As String = "_Ped_Ent_Keuze_"
Private Const constEntText As String = "_Ped_Ent_Opm"
Private Const constTpnText As String = "_Ped_TPN_Opm"

Private Const constSST1Stand As String = "_Ped_TPN_SST1Stand"
Private Const constSST1Keuze As String = "_Ped_TPN_SST1Keuze"
Private Const constSST2Stand As String = "_Ped_TPN_SST2Stand"
Private Const constSST2Keuze As String = "_Ped_TPN_SST2Keuze"

Private Const constTPN As String = "_Ped_TPN_Keuze"
Private Const constTPNVol As String = "_Ped_TPN_Vol"

Private Const constNaCl1 As String = "_Ped_TPN_NaCl1"
Private Const constNaCl1Vol As String = "_Ped_TPN_NaClVol1"
Private Const constKCl1 As String = "_Ped_TPN_KCl1"
Private Const constKCl1Vol As String = "_Ped_TPN_KClVol1"
Private Const constCaGluc As String = "_Ped_TPN_CaCl"
Private Const constCaGlucVol As String = "_Ped_TPN_CaGlucVol"
Private Const constMgCl As String = "_Ped_TPN_MgCl"
Private Const constMgClVol As String = "_Ped_TPN_MgClVol"

Private Const constNaCl2 As String = "_Ped_TPN_NaCl2"
Private Const constNaCl2Vol As String = "_Ped_TPN_NaClVol2"
Private Const constKCl2 As String = "_Ped_TPN_KCl2"
Private Const constKCl2Vol As String = "_Ped_TPN_KClVol2"

Private Const constKNaFosf As String = "_Ped_TPN_KNaFosf"
Private Const constKNaFosfVol As String = "_Ped_TPN_KNaFosfVol"

Private Const constPediTrace As String = "_Ped_TPN_PediTrace"

Private Const constLipid As String = "_Ped_TPN_LipidStand"

Private Const constSoluvit As String = "_Ped_TPN_Soluvit"
Private Const constSoluvitVol As String = "_Ped_TPN_SoluvitVol"
Private Const constVitIntra As String = "_Ped_TPN_VitIntra"
Private Const constVitIntraVol As String = "_Ped_TPN_VitIntraVol"


Public Sub PedEntTPN_ClearSSt()

    ModRange.SetRangeValue constTPN, 1
    ModRange.SetRangeValue constTPNVol, 0
    
    ModRange.SetRangeValue constSST1Stand, 0
    ModRange.SetRangeValue constSST1Keuze, 1
    
    ModRange.SetRangeValue constNaCl1, False
    ModRange.SetRangeValue constNaCl1Vol, 0
    
    ModRange.SetRangeValue constKCl1, False
    ModRange.SetRangeValue constKCl1Vol, 0
    
    ModRange.SetRangeValue constSST2Stand, 0
    ModRange.SetRangeValue constSST2Keuze, 1
    
    ModRange.SetRangeValue constNaCl2, False
    ModRange.SetRangeValue constNaCl2Vol, 0
    
    ModRange.SetRangeValue constKCl2, False
    ModRange.SetRangeValue constKCl2Vol, 0
    
    ModRange.SetRangeValue constCaGluc, False
    ModRange.SetRangeValue constCaGlucVol, 0
    
    ModRange.SetRangeValue constMgCl, False
    ModRange.SetRangeValue constMgClVol, 0
    
End Sub

Public Sub PedEntTPN_clearPeditrace()

    ModRange.SetRangeValue constPediTrace, 0

End Sub

Public Sub PedEntTPN_ClearCaGluc()

    ModRange.SetRangeValue constCaGluc, False
    ModRange.SetRangeValue constCaGlucVol, 0

End Sub

Public Sub PedEntTPN_ClearKNAP()
    
    ModRange.SetRangeValue constKNaFosf, False
    ModRange.SetRangeValue constKNaFosfVol, 0
    
End Sub

Public Sub PedEntTPN_ClearLipid()

    ModRange.SetRangeValue constLipid, 0
    ModRange.SetRangeValue constSoluvit, False
    ModRange.SetRangeValue constSoluvitVol, 0
    ModRange.SetRangeValue constVitIntra, False
    ModRange.SetRangeValue constVitIntraVol, 0

End Sub

Public Sub PedEntTPN_ClearOpmTPN()

    ModRange.SetRangeValue constTpnText, vbNullString

End Sub

Public Sub PedEntTPN_ShowVoedingPickList()

    Dim frmPickList As FormPedEntPickList
    Dim colVoeding As Collection
    Dim colToevoeging As Collection
    Dim intN As Integer
    Dim intC As Integer
    Dim intVoeding As Integer
    Dim intToevoeging As Integer
    
    Set colVoeding = ModRange.CollectionFromRange(constTblVoeding, 2)
    Set colToevoeging = ModRange.CollectionFromRange(constTblToevoeging, 2)
    
    Set frmPickList = New FormPedEntPickList
    frmPickList.LoadVoedingen colVoeding
    frmPickList.LoadToevoegingen colToevoeging
    
    For intN = 1 To constVoedingCount
        intVoeding = ModRange.GetRangeValue(constVoeding & intN, 1)
        If intVoeding > 1 Then frmPickList.SelectVoeding intVoeding
    Next intN
    
    For intN = 2 To constToevoegingCount + 1
        intToevoeging = ModRange.GetRangeValue(constVoeding & intN, 1)
        If intToevoeging > 1 Then frmPickList.SelectToevoeging intToevoeging
    Next intN
    
    frmPickList.Show
    
    If frmPickList.GetAction = vbNullString Then
    
        ' -- Process Voeding
    
        For intN = 1 To constVoedingCount            ' First remove nonselected items
            intVoeding = ModRange.GetRangeValue(constVoeding & intN, 1)
            If intVoeding > 1 Then
                If frmPickList.IsVoedingSelected(intVoeding) Then
                    frmPickList.UnselectVoeding (intVoeding)
                Else
                    ClearEnt intN
                End If
            End If
        Next intN
        
        Do While frmPickList.HasSelectedVoedingen()  ' Then add selected items
            For intN = 1 To constVoedingCount
                intVoeding = ModRange.GetRangeValue(constVoeding & intN, 1)
                If intVoeding <= 1 Then
                    intVoeding = frmPickList.GetFirstSelectedVoeding(True)
                    ModRange.SetRangeValue constVoeding & intN, intVoeding
                    Exit For
                End If
            Next intN
        Loop
    
        ' -- Process Toevoegingen
    
        For intN = 2 To constToevoegingCount + 1       ' First remove nonselected items
            intToevoeging = ModRange.GetRangeValue(constVoeding & intN, 1)
            If intToevoeging > 1 Then
                If frmPickList.IsToevoegingSelected(intToevoeging) Then
                    frmPickList.UnselectToevoeging (intToevoeging)
                Else
                    ClearEnt intN
                End If
            End If
        Next intN
        
        Do While frmPickList.HasSelectedToevoegingen()  ' Then add selected items
            For intN = 2 To constToevoegingCount + 1
                intToevoeging = ModRange.GetRangeValue(constVoeding & intN, 1)
                If intToevoeging <= 1 Then
                    intToevoeging = frmPickList.GetFirstSelectedToevoeging(True)
                    ModRange.SetRangeValue constVoeding & intN, intToevoeging
                    Exit For
                End If
            Next intN
        Loop
    
    End If
    
    Set frmPickList = Nothing
    
End Sub


Public Sub PedEntTPN_ClearSonde()

    ModRange.SetRangeValue constSode, 1

End Sub

Private Sub ClearEnt(ByVal intN As Integer)

    ModRange.SetRangeValue constVoeding & intN, 1

End Sub

Public Sub PedEntTPN_ClearVoeding_1()

    ClearEnt 1

End Sub

Public Sub PedEntTPN_ClearVoeding_2()

    ClearEnt 2

End Sub

Public Sub PedEntTPN_ClearVoeding_3()

    ClearEnt 3

End Sub

Public Sub PedEntTPN_ClearVoeding_4()

    ClearEnt 4

End Sub

Public Sub PedEntTPN_ClearOpm()

    ModRange.SetRangeValue constEntText, vbNullString

End Sub


Public Sub PedEntTPN_SelectTPNPrint()

    Dim dblGew As Double
    
    dblGew = ModPatient.GetGewichtFromRange

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
        Else          ' Not a valid weight
            Exit Sub  ' So exit sub
        End If
        
        .Range(strSelected).PasteSpecial xlPasteValues
    End With
    
    Application.Calculate
    
End Sub

Private Sub TPNAdvies(ByVal intDag As Integer)

    Dim dblVol As Double
    Dim dblNaCl As Double
    Dim dblKCl As Double
    Dim dblVitIntra As Double
    Dim dblLipid As Double
    Dim dblSolu As Double
    Dim dblGewicht As Double
    
    dblGewicht = ModPatient.GetGewichtFromRange()

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
            
            Select Case intDag
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
            
            Select Case intDag
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
            
            Select Case intDag
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
            
            Select Case intDag
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
            
            Select Case intDag
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

Private Sub EnterOpmAfspr(ByVal strRange As String)

    Dim frmOpmerking As FormOpmerking
    
    Set frmOpmerking = New FormOpmerking
    frmOpmerking.txtOpmerking.Text = ModRange.GetRangeValue(strRange, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue strRange, frmOpmerking.txtOpmerking.Text
    End If
       
    Set frmOpmerking = Nothing

End Sub

Public Sub PedEntTPN_EntText()
    
    EnterOpmAfspr constEntText
    
End Sub

Public Sub PedEntTPN_TPNText()
    
    EnterOpmAfspr constTpnText

End Sub

Private Sub EnterHoeveelheid(ByVal strRange As String, ByVal strItem As String)

    Dim frmInvoer As FormInvoerNumeriek
    
    Set frmInvoer = New FormInvoerNumeriek
    frmInvoer.SetValue strRange, strItem, ModRange.GetRangeValue(strRange, 0), "mL", vbNullString
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


Public Sub PedEntTPN_ChangeVitIntra()

    ModRange.SetRangeValue constVitIntra, True

End Sub

Public Sub PedEntTPN_ChangeSoluvit()

    ModRange.SetRangeValue constSoluvit, True

End Sub


Public Sub PedEntTPN_SpecVoed()
    
    Dim frmSpecialeVoeding As FormSpecialeVoeding

    Set frmSpecialeVoeding = New FormSpecialeVoeding
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

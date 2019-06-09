Attribute VB_Name = "ModPedEntTPN"
Option Explicit

Private Const CONST_TPN_1 As Integer = 2
Private Const CONST_TPN_2 As Integer = 7
Private Const CONST_TPN_3 As Integer = 16
Private Const CONST_TPN_4 As Integer = 31
Private Const CONST_TPN_5 As Integer = 50

Private Const constTblVoeding As String = "Tbl_Ped_Voeding"
Private Const constTblToevoeging As String = "Tbl_Ped_Poeder"

Private Const constVoedingCount As Integer = 1
Private Const constToevoegingCount As Integer = 3

Private Const constSode As String = "_Ped_Ent_Sonde"
Private Const constVoeding As String = "_Ped_Ent_Keuze_"

Private Const constEntText As String = "_Ped_Ent_Opm"
Private Const constTpnText As String = "_Ped_TPN_Opm"

Private Const constLipidStand As String = "_Ped_TPN_LipidStand"
Private Const constSST1Stand As String = "_Ped_TPN_SST1Stand"
Private Const constSST1Vol As String = "_Ped_TPN_SST1Vol"
Private Const constSST1Keuze As String = "_Ped_TPN_SST1Keuze"
Private Const constSST2Stand As String = "_Ped_TPN_SST2Stand"
Private Const constSST2Keuze As String = "_Ped_TPN_SST2Keuze"

Private Const constTPN As String = "_Ped_TPN_Keuze"
Private Const constTPNVol As String = "_Ped_TPN_Vol"
Private Const constTPNDag As String = "_Ped_TPN_DagKeuze"

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

Private Const constLipidVol As String = "_Ped_TPN_LipidVol"

Private Const constSoluvit As String = "_Ped_TPN_Soluvit"
Private Const constSoluvitVol As String = "_Ped_TPN_SoluvitVol"
Private Const constVitIntra As String = "_Ped_TPN_VitIntra"
Private Const constVitIntraVol As String = "_Ped_TPN_VitIntraVol"

Private Const constGluc05 As Integer = 2     'glucose 5
Private Const constGluc10 As Integer = 3     'glucose 10
Private Const constGluc12_5 As Integer = 4   'glucose 12, 5
Private Const constGluc15 As Integer = 5     'glucose 15
Private Const constGluc17_5 As Integer = 6   'glucose 17, 5
Private Const constGluc20 As Integer = 7     'glucose 20
Private Const constGluc25 As Integer = 8     'glucose 25
Private Const constGluc30 As Integer = 9     'glucose 30
Private Const constGluc40 As Integer = 10    'glucose 40
Private Const constGluc50 As Integer = 11    'glucose 50

Private Const constMaxSoluvit = 10
Private Const constMaxVitIntra = 10


Public Sub PedEntTPN_ClearSST1()

    ModRange.SetRangeValue constTPN, 1
    ModRange.SetRangeValue constTPNVol, 0
    
    ModRange.SetRangeValue constSST1Stand, 0
    ModRange.SetRangeValue constSST1Vol, 0
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

Public Sub PedEntTPN_ClearSST2()

    ModRange.SetRangeValue constSST2Stand, 0
    ModRange.SetRangeValue constSST2Keuze, 1
    
    ModRange.SetRangeValue constNaCl2, False
    ModRange.SetRangeValue constNaCl2Vol, 0
    
    ModRange.SetRangeValue constKCl2, False
    ModRange.SetRangeValue constKCl2Vol, 0
    
End Sub

Public Sub PedEntTPN_ClearKCl1()

    ModRange.SetRangeValue constKCl1, False
    ModRange.SetRangeValue constKCl1Vol, 0

End Sub

Public Sub PedEntTPN_ClearNaCl1()

    ModRange.SetRangeValue constNaCl1, False
    ModRange.SetRangeValue constNaCl1Vol, 0

End Sub

Public Sub PedEntTPN_ClearKCl2()

    ModRange.SetRangeValue constKCl2, False
    ModRange.SetRangeValue constKCl2Vol, 0

End Sub

Public Sub PedEntTPN_ClearNaCl2()

    ModRange.SetRangeValue constNaCl2, False
    ModRange.SetRangeValue constNaCl2Vol, 0

End Sub

Public Sub PedEntTPN_ClearVitIntra()

    ModRange.SetRangeValue constVitIntra, False
    ModRange.SetRangeValue constVitIntraVol, 0

End Sub

Public Sub PedEntTPN_ClearSoluvit()

    ModRange.SetRangeValue constSoluvit, False
    ModRange.SetRangeValue constSoluvitVol, 0

End Sub

Public Sub PedEntTPN_ClearPeditrace()

    ModRange.SetRangeValue constPediTrace, 0

End Sub

Public Sub PedEntTPN_ClearCaGluc()

    ModRange.SetRangeValue constCaGluc, False
    ModRange.SetRangeValue constCaGlucVol, 0

End Sub

Public Sub PedEntTPN_ClearMgCl()

    ModRange.SetRangeValue constMgCl, False
    ModRange.SetRangeValue constMgClVol, 0

End Sub

Public Sub PedEntTPN_ClearKNAP()
    
    ModRange.SetRangeValue constKNaFosf, False
    ModRange.SetRangeValue constKNaFosfVol, 0
    
End Sub

Public Sub PedEntTPN_ClearLipid()

    ModRange.SetRangeValue constLipidStand, 0
    ModRange.SetRangeValue constLipidVol, 0
    ModRange.SetRangeValue constSoluvit, False
    ModRange.SetRangeValue constSoluvitVol, 0
    ModRange.SetRangeValue constVitIntra, False
    ModRange.SetRangeValue constVitIntraVol, 0

End Sub

Public Sub PedEntTPN_ClearTPN()

    ModRange.SetRangeValue constTPN, 1

End Sub

Public Sub PedEntTPN_ClearOpmTPN()

    ModRange.SetRangeValue constTpnText, vbNullString

End Sub

Public Sub PedEntTPN_ShowVoedingPickList()

    Dim frmPickList As FormPedEntPickList
    Dim colVoeding As Collection
    Dim colToevoeging As Collection
    Dim intN As Integer
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
    
End Sub


Public Sub PedEntTPN_ClearSonde()

    ModRange.SetRangeValue constSode, 1

End Sub

Private Sub ClearEnt(ByVal intN As Integer)

    ModRange.SetRangeValue constVoeding & intN, 1
    If intN = 1 Then
        PedEntTPN_ChangeEnt_1
    Else
        ChangeEnt intN
    End If
    
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

Public Function GetPedTPNIndexForWeight() As Integer

    Dim dblWeight As Double
    Dim intTPN As Integer
    
    dblWeight = ModPatient.Patient_GetWeight()
    
    intTPN = 1
    intTPN = IIf(dblWeight >= CONST_TPN_1, 3, intTPN)
    intTPN = IIf(dblWeight >= CONST_TPN_2, 4, intTPN)
    intTPN = IIf(dblWeight >= CONST_TPN_3, 5, intTPN)
    intTPN = IIf(dblWeight >= CONST_TPN_4, 6, intTPN)
    intTPN = IIf(dblWeight > CONST_TPN_5, 7, intTPN)
    
    GetPedTPNIndexForWeight = intTPN

End Function

Public Sub PedEntTPN_SelectStandardTPN()

    If ModRange.GetRangeValue(constTPN, 1) > 1 Then ModRange.SetRangeValue constTPN, GetPedTPNIndexForWeight()
    
End Sub

Private Sub SetTPNAdvies(ByVal strRange As String, dblVal As Double)

    ModRange.SetRangeValue strRange, Round(dblVal, 1)

End Sub

Public Sub PedTPN_SetLipidStand()

    Dim dblStand As Double
    
    dblStand = ModRange.GetRangeValue("Var_Ped_TPN_LipidVol", 0) / 24

    ' ALS(C3<100;C3/10;C3-90)
    If dblStand < 10 Then
        dblStand = Round(dblStand, 1)
        SetTPNAdvies constLipidStand, dblStand * 10
    Else
        dblStand = Round(dblStand, 0) + 90
        SetTPNAdvies constLipidStand, dblStand
    End If

End Sub
Public Sub PedTPN_SetSST1Stand()

    Dim dblStand As Double
    
    dblStand = ModRange.GetRangeValue("Var_Ped_TPN_SST1Vol", 0) / 24

    ' ALS(C3<100;C3/10;C3-90)
    If dblStand < 10 Then
        dblStand = Round(dblStand, 1)
        SetTPNAdvies constSST1Stand, dblStand * 10
    Else
        dblStand = Round(dblStand, 0) + 90
        SetTPNAdvies constSST1Stand, dblStand
    End If

End Sub

Public Sub PedTPN_SetSST2Stand()

    Dim dblStand As Double
    
    dblStand = ModRange.GetRangeValue("Var_Ped_TPN_SST2Vol", 0) / 24

    ' ALS(C3<100;C3/10;C3-90)
    If dblStand < 10 Then
        dblStand = Round(dblStand, 1)
        SetTPNAdvies constSST2Stand, dblStand * 10
    Else
        dblStand = Round(dblStand, 0) + 90
        SetTPNAdvies constSST2Stand, dblStand
    End If

End Sub

Private Sub ClearRangeColor(objRange As Range)

    With objRange.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Private Sub SetRangeColor(objRange As Range, xlColor As XlThemeColor)

    With objRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlColor
        .TintAndShade = -0.14996795556505
        .PatternTintAndShade = 0
    End With

End Sub

Private Sub TPNAdvies(ByVal intDag As Integer, Optional ByVal varTPN As Variant)

    Dim dblGewicht As Double
    Dim intTPN As Integer
    Dim dblTPNVol As Double
    Dim intSSTKeuze As Integer
    Dim dblSSTVol As Double
    Dim blnNaCl As Boolean
    Dim dblNaCl As Double
    Dim blnKCl As Boolean
    Dim dblKCl As Double
    Dim dblLipid As Double
    Dim blnVitIntra As Boolean
    Dim dblVitIntra As Double
    Dim blnSolu As Boolean
    Dim dblSolu As Double
    Dim dblPediTrace As Double
    Dim objSheet As Worksheet
    Dim objRange As Range
    Dim strRange As String
    
    dblGewicht = ModPatient.Patient_GetWeight()
    
    intTPN = ModRange.GetRangeValue(constTPN, 1)
    intTPN = IIf(IsMissing(varTPN), intTPN, CInt(varTPN))
    
    PedEntTPN_ClearSST1
    PedEntTPN_ClearLipid
        
    If intDag = 4 Or intDag = 0 Then
        Exit Sub
    End If

    If dblGewicht < 2 Then Exit Sub
    
    Select Case Int(dblGewicht)
    Case 2 To 6
        Set objSheet = shtPedPrtTPN2tot6
    
        intTPN = IIf(intTPN = 2, 2, 3)
        blnNaCl = True
        dblNaCl = 6 * dblGewicht
        blnKCl = True
        dblKCl = 1 * dblGewicht
        blnVitIntra = True
        dblVitIntra = dblGewicht
            
        Select Case intDag
        Case 1
            dblKCl = 1.5 * dblGewicht
            dblTPNVol = 15 * dblGewicht
            dblLipid = 5 * dblGewicht
            intSSTKeuze = constGluc10
            dblSSTVol = 120 * dblGewicht
                
        Case 2
            dblTPNVol = 25 * dblGewicht
            dblLipid = 10 * dblGewicht
            intSSTKeuze = constGluc12_5
            dblSSTVol = 105 * dblGewicht
                
        Case 3
            dblTPNVol = 35 * dblGewicht
            dblLipid = 15 * dblGewicht
            intSSTKeuze = constGluc17_5
            dblSSTVol = 90 * dblGewicht
            
        End Select
            
    Case 7 To 15
        Set objSheet = shtPedPrtTPN7tot15
        
        blnNaCl = True
        dblNaCl = 6 * dblGewicht
        blnKCl = True
        dblKCl = 1.5 * dblGewicht
        dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
        blnVitIntra = True
        blnSolu = True
        dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            
        Select Case intDag
        Case 1
            dblKCl = 2 * dblGewicht
            dblTPNVol = 10 * dblGewicht
            dblLipid = 5 * dblGewicht
            intSSTKeuze = constGluc10
            dblSSTVol = 65 * dblGewicht
                
        Case 2
            dblTPNVol = 20 * dblGewicht
            dblLipid = 10 * dblGewicht
            intSSTKeuze = constGluc20
            dblSSTVol = 50 * dblGewicht
                
        Case 3
            dblTPNVol = 25 * dblGewicht
            dblLipid = 15 * dblGewicht
            intSSTKeuze = constGluc30
            dblSSTVol = IIf(dblGewicht < 11, 55 * dblGewicht, 40 * dblGewicht)
        
        End Select
                            
    Case 16 To 30
        Set objSheet = shtPedPrtTPN16tot30
    
        blnNaCl = True
        dblNaCl = 6 * dblGewicht
        blnKCl = True
        dblKCl = 1.5 * dblGewicht
        blnVitIntra = True
        dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
        blnSolu = True
        dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
        dblPediTrace = 15
            
        Select Case intDag
        Case 1
            dblKCl = 2 * dblGewicht
            dblTPNVol = 10 * dblGewicht
            dblLipid = 5 * dblGewicht
            intSSTKeuze = constGluc10
            dblSSTVol = 50 * dblGewicht
                
        Case 2
            dblTPNVol = 15 * dblGewicht
            dblLipid = 10 * dblGewicht
            intSSTKeuze = constGluc20
            dblSSTVol = 40 * dblGewicht
                
        Case 3
            dblTPNVol = 20 * dblGewicht
            dblLipid = 15 * dblGewicht
            intSSTKeuze = constGluc30
            dblSSTVol = 30 * dblGewicht
            
        End Select
                    
    Case 31 To 50
        Set objSheet = shtPedPrtTPN31tot50
    
        blnNaCl = True
        dblNaCl = 6 * dblGewicht
        blnKCl = True
        dblKCl = 1.5 * dblGewicht
        blnVitIntra = True
        dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
        blnSolu = True
        dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
        dblPediTrace = 15
            
        Select Case intDag
        Case 1
            dblKCl = 2 * dblGewicht
            dblTPNVol = 5 * dblGewicht
            dblLipid = 3 * dblGewicht
            intSSTKeuze = constGluc10
            dblSSTVol = 35 * dblGewicht
                
        Case 2
            dblTPNVol = 8 * dblGewicht
            dblLipid = 6 * dblGewicht
            intSSTKeuze = constGluc20
            dblSSTVol = 30 * dblGewicht
                
        Case 3
            dblTPNVol = 12 * dblGewicht
            dblLipid = 10 * dblGewicht
            intSSTKeuze = IIf(ModPatient.Patient_GetWeight >= 40, constGluc40, constGluc25)
            dblSSTVol = IIf(dblGewicht < 36, 40 * dblGewicht, 20 * dblGewicht)
            
        End Select
            
    Case Else
        Set objSheet = shtPedPrtTPN50

        blnNaCl = False
        blnKCl = False
        blnVitIntra = True
        dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
        blnSolu = True
        dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
        dblPediTrace = 15
        intSSTKeuze = constGluc10
            
        Select Case intDag
        Case 1
            dblTPNVol = 700
            dblLipid = 150
                
        Case 2
            dblTPNVol = 1000
            dblLipid = 300
                
        Case 3
            dblTPNVol = 1500
            dblLipid = 500
        End Select
                
    End Select
    
    objSheet.Unprotect CONST_PASSWORD
    
    Set objRange = objSheet.Range("D11:F11")
    ClearRangeColor objRange
    
    strRange = IIf(intDag = 1, "D11", vbNullString)
    strRange = IIf(intDag = 2, "E11", strRange)
    strRange = IIf(intDag = 3, "F11", strRange)
    Set objRange = objSheet.Range(strRange)
    SetRangeColor objRange, xlThemeColorDark1
    
    If Not ModSetting.IsDevelopmentMode Then objSheet.Protect CONST_PASSWORD
    
    ModRange.SetRangeValue constTPN, intTPN
    SetTPNAdvies constTPNVol, dblTPNVol
    
    ModRange.SetRangeValue constPediTrace, dblPediTrace
    
    ModRange.SetRangeValue constSST1Keuze, intSSTKeuze
    SetTPNAdvies constSST1Vol, dblSSTVol
    
    ModRange.SetRangeValue constNaCl1, blnNaCl
    SetTPNAdvies constNaCl1Vol, dblNaCl
    
    ModRange.SetRangeValue constKCl1, blnKCl
    SetTPNAdvies constKCl1Vol, dblKCl
    
    SetTPNAdvies constLipidVol, dblLipid
    
    ModRange.SetRangeValue constVitIntra, blnVitIntra
    SetTPNAdvies constVitIntraVol, dblVitIntra
    
    ModRange.SetRangeValue constSoluvit, blnSolu
    SetTPNAdvies constSoluvitVol, dblSolu
    
    PedTPN_SetSST1Stand
    PedTPN_SetLipidStand
    
End Sub

Public Sub PedEntTPN_TPNAdviceDay_0()

    TPNAdvies 0

End Sub

Public Sub PedEntTPN_TPNAdviceDay_1()

    TPNAdvies 1

End Sub

Public Sub PedEntTPN_TPNAdviceDay_2()

    TPNAdvies 2

End Sub

Public Sub PedEntTPN_TPNAdviceDay_3()

    TPNAdvies 3

End Sub

Private Sub EnterOpmAfspr(ByVal strRange As String)

    Dim frmOpmerking As FormOpmerking
    
    Set frmOpmerking = New FormOpmerking
    frmOpmerking.SetText ModRange.GetRangeValue(strRange, vbNullString)
    frmOpmerking.Show
    
    If frmOpmerking.txtOpmerking.Text <> "Cancel" Then
        ModRange.SetRangeValue strRange, frmOpmerking.txtOpmerking.Text
    End If
       
End Sub

Public Sub PedEntTPN_EntText()
    
    EnterOpmAfspr constEntText
    
End Sub

Public Sub PedEntTPN_TPNText()
    
    EnterOpmAfspr constTpnText

End Sub

Private Sub EnterHoeveelheid(ByVal strRange As String, ByVal strItem As String)

    Dim frmInvoer As FormInvoerNumeriek
    Dim dblValue As Double
    
    Set frmInvoer = New FormInvoerNumeriek
    frmInvoer.lblText.Caption = "Voer hoeveelheid in voor " & strItem
    
    dblValue = ModRange.GetRangeValue(strRange, 0)
    frmInvoer.SetValue strRange, strItem, dblValue, "mL", vbNullString
    
    frmInvoer.Show
    
End Sub

Public Sub PedEntTPN_TPN()

    EnterHoeveelheid constTPNVol, "TPN"
    If ModRange.GetRangeValue(constTPNVol, 0) = 0 Then
        ModRange.SetRangeValue constTPN, 1
    End If

End Sub

Public Sub PedEntTPN_ChangeTPNVol()

    If ModRange.GetRangeValue(constTPNVol, 0) = 0 Then
        ModRange.SetRangeValue constTPN, 1
    End If

End Sub

Public Sub PedEntTPN_SST1()

    EnterHoeveelheid "_Ped_TPN_SST1Vol", "SST1"
    ModRange.SetRangeValue constNaCl1, True
    
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

Public Sub PedEntTPN_Phosph()

    EnterHoeveelheid constKNaFosfVol, "Fosfaat"
    ModRange.SetRangeValue constKNaFosf, True

End Sub

Public Sub PedEntTPN_IntraLipid()

    EnterHoeveelheid constLipidVol, "Lipiden"

End Sub

Public Sub PedEntTPN_VitIntra()

    EnterHoeveelheid constVitIntraVol, "Vitintra"
    ModRange.SetRangeValue constVitIntra, True

End Sub
Public Sub PedEntTPN_Soluvit()

    EnterHoeveelheid constSoluvitVol, "Soluvit"
    ModRange.SetRangeValue constSoluvit, True

End Sub
Public Sub PedEntTPN_ChangeTPN()

    Dim intDag As Integer
    Dim intTPN As Integer
    
    intDag = ModRange.GetRangeValue(constTPNDag, 0)
    intTPN = ModRange.GetRangeValue(constTPN, 1)
    If intTPN > 1 And (intDag = 4 Or intDag = 0) Then
        ModMessage.ShowMsgBoxInfo "Kies een TPN dag"
    End If
    
    TPNAdvies intDag, intTPN

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
    
End Sub

Private Sub ChangeEnt(ByVal intN As Integer)
    
    Dim strKeuze As String
    Dim strFreq As String
    Dim strVol As String
    
    strKeuze = "_Ped_Ent_Keuze_" & intN
    strFreq = "_Ped_Ent_Freq_" & intN
    strVol = "_Ped_Ent_Vol_" & intN
    
    If ModRange.GetRangeValue(strKeuze, 0) > 1 Then
        If ModRange.GetRangeValue(strFreq, 0) > 0 And ModRange.GetRangeValue(strVol, 0) = 0 Then ModRange.SetRangeValue strVol, 1
        If ModRange.GetRangeValue(strVol, 0) > 0 And ModRange.GetRangeValue(strFreq, 0) = 0 Then ModRange.SetRangeValue strFreq, 1
    Else
        ModRange.SetRangeValue strFreq, 0
        ModRange.SetRangeValue strVol, 0
    End If

End Sub

Public Sub PedEntTPN_ChangeEnt_1()

    If ModRange.GetRangeValue("_Ped_Ent_Keuze_1", 0) = 1 Then
    
        ModRange.SetRangeValue "_Ped_Ent_Keuze_2", 1
        ModRange.SetRangeValue "_Ped_Ent_Keuze_3", 1
        ModRange.SetRangeValue "_Ped_Ent_Keuze_4", 1
        
        ChangeEnt 1
        ChangeEnt 2
        ChangeEnt 3
        ChangeEnt 4
    
    Else
        ChangeEnt 1
    End If

End Sub

Public Sub PedEntTPN_ChangeEnt_2()

    ChangeEnt 2

End Sub

Public Sub PedEntTPN_ChangeEnt_3()
    
    ChangeEnt 3

End Sub

Public Sub PedEntTPN_ChangeEnt_4()
    
    ChangeEnt 4

End Sub

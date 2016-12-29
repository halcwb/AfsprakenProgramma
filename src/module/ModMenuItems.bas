Attribute VB_Name = "ModMenuItems"
Option Explicit

Public Sub StandaardInstellingen()

    Dim intN As Integer
    
    On Error GoTo StandaarInstellingenError
        
    GoTo StandaarInstellingenError
    
    Application.Cursor = xlWait
    
    With shtPatData
        For intN = 3 To 25
            Range(.Cells(intN, 1).Value).Formula = .Cells(intN, 3).Formula
        Next intN
    End With
    
    With shtPatData
        For intN = 105 To 150
            Range(.Cells(intN, 1).Value).Formula = .Cells(intN, 3).Formula
        Next intN
    End With
    
    With shtPatData
        For intN = 370 To 392
            Range(.Cells(intN, 1).Value).Formula = .Cells(intN, 3).Formula
        Next intN
    End With
    
    Application.Cursor = xlDefault
    
    Exit Sub
    
StandaarInstellingenError:

    ModMessage.ShowMsgBoxExclam "Dit werkt nog niet"

End Sub

Public Sub InstellingenKlein()

    Dim intN As Integer
    
    On Error GoTo InstellingenKleinError
    
    GoTo InstellingenKleinError
        
    Application.Cursor = xlWait
    
    With shtPatData
        For intN = 3 To 25
            Range(.Cells(intN, 1).Value).Formula = .Cells(intN, 4).Formula
        Next intN
    End With
    
    With shtPatData
        For intN = 105 To 150
            Range(.Cells(intN, 1).Value).Formula = .Cells(intN, 3).Formula
        Next intN
    End With
    
    With shtPatData
        For intN = 370 To 392
            Range(.Cells(intN, 1).Value).Formula = .Cells(intN, 4).Formula
        Next intN
    End With
    
    Application.Cursor = xlDefault
    
    Exit Sub

InstellingenKleinError:

    ModMessage.ShowMsgBoxExclam "Dit werkt nog niet"

End Sub

Public Sub SelectPedTPNPrint()

    Dim dblGew As Double
    
    dblGew = Val(ModRange.GetRangeValue("Gewicht", 0)) / 10

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

Public Sub SaveAndPrintAfspraken()

    Dim frmPrintAfspraken As New FormPrintAfspraken
    
    ModBed.SluitBed
    frmPrintAfspraken.Show
    
    Set frmPrintAfspraken = Nothing
    
End Sub

Public Sub SpecialeVoeding()
    
    Dim frmSpecialeVoeding As New FormSpecialeVoeding

    frmSpecialeVoeding.Show
    
    Set frmSpecialeVoeding = Nothing

End Sub

Public Sub SelectTPN()

    Dim dblGewicht As Double

    dblGewicht = Val(ModRange.GetRangeValue("Gewicht", 0)) / 10

    With shtPedBerTPN
        If dblGewicht >= 2 And dblGewicht < 7 Then
            .Range("tpnB").Copy
        ElseIf dblGewicht >= 7 And dblGewicht < 15 Then
            .Range("tpnC").Copy
        ElseIf dblGewicht >= 15 And dblGewicht < 30 Then
            .Range("tpnD").Copy
        ElseIf dblGewicht >= 30 And dblGewicht <= 50 Then
            .Range("tpnE").Copy
        ElseIf dblGewicht > 50 Then
            .Range("tpnNutriflex").Copy
        End If
        .Range("tpnSelected").PasteSpecial xlPasteValues
    End With
    
    Application.Calculate
    
End Sub

Private Sub ClearContentsSheetRange(shtSheet As Worksheet, strRange As String)

    Dim blnLog As Boolean
    Dim strError As String
    Dim blnIsDevelop As Boolean
    Dim strPw As String
    
    On Error GoTo ClearContentError
    
    blnIsDevelop = ModSetting.IsDevelopmentMode()
    strPw = ModConst.CONST_PASSWORD
    
    shtSheet.Unprotect strPw
    shtSheet.Visible = xlSheetVisible
    
    Application.Goto Reference:=strRange
    Selection.ClearContents
    
    If strRange = ModConst.CONST_RANGE_NEOMRI Then
        Selection.Value = 50
    End If
    
    If Not blnIsDevelop Then
        shtSheet.Visible = xlSheetVeryHidden
        shtSheet.Protect strPw
    End If
    
    Exit Sub
    
ClearContentError:
    
    strError = "ClearContentSheetRange Sheet: " & shtSheet.Name & " could  not clear content of Range: " & strRange
    ModLog.LogError strError

End Sub

Public Sub ClearLab()
    
    ClearContentsSheetRange shtPedBerLab, ModConst.CONST_RANGE_PEDLAB
    ClearContentsSheetRange shtNeoBerLab, ModConst.CONST_RANGE_NEOLAB
    
End Sub

Public Sub ClearAfspraken()

    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_RANGE_NEOBOOL
    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_RANGE_NEODATA
    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_RANGE_NEOMRI
    
    ClearContentsSheetRange shtPedBerExtraAfspr, ModConst.CONST_RANGE_PEDBOOL
    ClearContentsSheetRange shtPedBerExtraAfspr, ModConst.CONST_RANGE_PEDDATA

End Sub

Private Sub TPNAdvies(Dag As Integer)

    Dim dblVol As Double
    Dim dblNaCl As Double
    Dim dblKCl As Double
    Dim dblVitIntra As Double
    Dim dblLipid As Double
    Dim dblSolu As Double
    Dim dblGewicht As Double
    
    dblGewicht = Val(ModRange.GetRangeValue("Gewicht", 0)) / 10

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

Public Sub TPNAdviesDagEen()

    TPNAdvies (1)

End Sub

Public Sub TPNAdviesDagTwee()

    TPNAdvies (2)

End Sub

Public Sub TPNAdviesDagDrie()

    TPNAdvies (3)

End Sub

Public Sub PrintLabAanvragen()

    With Application
        .DisplayAlerts = False
        'TODO: Link controleren op werking
        .Workbooks.Open "G:\Zorgeenh\Pelikaan\ICAP Data\LabAanvragen.xls", True, True
        .ActiveWorkbook.Sheets("Unit 1").PrintOut
        .ActiveWorkbook.Sheets("Unit 2").PrintOut
        .Workbooks("LabAanvragen.xls").Close
    End With

End Sub

Public Sub AfsprakenPrinten()

    shtNeoPrtAfspr.PrintPreview

End Sub

Public Sub WerkBriefPrinten()
        
    With shtNeoPrtWerkbr
        .Unprotect ModConst.CONST_PASSWORD
        .PrintPreview
        .Protect ModConst.CONST_PASSWORD
    End With

End Sub

Public Sub TPNAdviesNEO()

    ModRange.SetRangeValue "_DagKeuze", IIf(ModRange.GetRangeValue("Dag", 0) < 4, 1, 2)
    ModRange.SetRangeValue "_IntakePerKg", 5000
    ModRange.SetRangeValue "_IntraLipid", 5000
    ModRange.SetRangeValue "_NaCl", 5000
    ModRange.SetRangeValue "_KCl", 5000
    ModRange.SetRangeValue "_CaCl2", 5000
    ModRange.SetRangeValue "_MgCl2", 5000
    ModRange.SetRangeValue "_SoluVit", 5000
    ModRange.SetRangeValue "_Primene", 5000
    ModRange.SetRangeValue "_NICUMix", 5000
    ModRange.SetRangeValue "_SSTB", 5000
    
    ModSheet.GoToSheet shtNeoGuiAfspraken, "A9"

End Sub

' Shows the frmNaamGeven to give a range a
' sequential naming of "Name_" + a number
' When runnig this from the visual basic editor
' it works as expected. When running from the ribbon
' menu, the selection in the sheet is not visible.
' But it works as otherwise.
Public Sub GiveNameToRange()

    Dim frmNaamGeven As New FormNaamGeven
    
    frmNaamGeven.Show vbModal
    
    Set frmNaamGeven = Nothing

End Sub


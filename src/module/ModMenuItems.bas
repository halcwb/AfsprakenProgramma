Attribute VB_Name = "ModMenuItems"
Option Explicit

Dim BAXA As Boolean

Public Sub NieuwePatient()

    Dim frmPatient As New FormPatient
    
    frmPatient.Show
    
    Set frmPatient = Nothing

End Sub

Public Sub PatientenLijst()

    Dim frmPatLijst As New FormPatLijst

    frmPatLijst.Show
    
    Set frmPatLijst = Nothing

End Sub

Public Function SavePatient() As Boolean

    Dim intN As Integer
    
    On Error GoTo Hell
    
    With Sheets("Patienten")
        For intN = 2 To .Range("A1").CurrentRegion.Rows.Count
            .Cells(intN, 4).Formula = Range(.Cells(intN, 1).Value).Formula
        Next intN
    End With
    
    SavePatient = True
    
Hell:
    Resume Next
End Function

Public Sub ClearPatient(blnShowWarn As Boolean)
    
    Dim intN As Integer, objResult As VbMsgBoxResult
            
    If blnShowWarn Then
        objResult = MsgBox("Afspraken echt verwijderen?", vbYesNo, "Informedica")
    Else
        objResult = vbYes
    End If
    
    If objResult = vbYes Then
        Application.Cursor = xlWait
        With shtPatData
            For intN = 2 To .Range("A1").CurrentRegion.Rows.Count
                Range(.Cells(intN, 1).Value).Formula = .Cells(intN, 3).Formula
            Next intN
        End With
        VerwijderLab
        VerwijderAanvullendeAfspraken
        Application.Cursor = xlDefault
    End If
    
    Range("AfspraakDatum").FormulaLocal = "=Vandaag()"
    SetApplTitle
    
End Sub

Public Sub StandaardInstellingen()

    Dim intN As Integer
    
    On Error GoTo Hell
        
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

Hell:
Resume Next

End Sub

Public Sub InstellingenKlein()

    Dim intN As Integer
    
    On Error GoTo Hell
        
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

Hell:
Resume Next

End Sub

Public Sub SelectPedTPNPrint()

        If Val(Range("Gewicht").Text) / 10 >= CONST_TPN_1 And Val(Range("Gewicht").Text) / 10 <= CONST_TPN_2 Then
            shtPedPrtTPN2tot6.Select
        Else
            If Val(Range("Gewicht").Text) / 10 <= CONST_TPN_3 Then
                shtPedPrtTPN7tot15.Select
            Else
                If Val(Range("Gewicht").Text) / 10 <= CONST_TPN_4 Then
                    shtPedPrtTPN16tot30.Select
                Else
                    If Val(Range("Gewicht").Text) / 10 <= CONST_TPN_5 Then
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
    
    ModBedden.SluitBed
    frmPrintAfspraken.Show
    
    Set frmPrintAfspraken = Nothing
    
End Sub

Public Sub AfsprakenVerw()
    'TODO: Geeft compilatiefout
    'TODO: controleren of dit nog gebruikt wordt
    'clearPat (4)

End Sub

Public Sub SpecialeVoeding()
    
    Dim frmSpecialeVoeding As New FormSpecialeVoeding

    frmSpecialeVoeding.Show
    
    Set frmSpecialeVoeding = Nothing

End Sub

Public Sub SelectTPN()

    Dim dblGewicht As Double

    dblGewicht = shtPatDetails.Range("Gewicht").Value / 10

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

    Dim blnIsDevelop As Boolean
    Dim strPw As String
    
    blnIsDevelop = ModSettings.IsDevelopmentMode()
    strPw = ModConst.CONST_PASSWORD
    
    shtSheet.Unprotect strPw
    shtSheet.Visible = xlSheetVisible
    
    Application.GoTo Reference:=strRange
    Selection.ClearContents
    
    If strRange = ModConst.CONST_AANVULLEND_MRI_VERTREKTIJD Then
        Selection.Value = 50
    End If
    
    If Not blnIsDevelop Then
        shtSheet.Visible = xlSheetVeryHidden
        shtSheet.Protect strPw
    End If


End Sub

Public Sub VerwijderLab()
    
    ClearContentsSheetRange shtPedBerLab, ModConst.CONST_LABDATA
    ClearContentsSheetRange shtNeoBerLab, ModConst.CONST_LABDATA_NEO
    
End Sub

Public Sub VerwijderAanvullendeAfspraken()

    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_AANVULLEND_BOOLEANS
    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_AANVULLEND_DATA
    ClearContentsSheetRange shtNeoBerAfspr, ModConst.CONST_AANVULLEND_MRI_VERTREKTIJD
    
    ClearContentsSheetRange shtPedBerExtraAfspr, ModConst.CONST_AANVULLEND_BOOLEANS_PED
    ClearContentsSheetRange shtPedBerExtraAfspr, ModConst.CONST_AANVULLEND_DATA_PED

End Sub

Private Sub TPNAdvies(Dag As Integer)

    Dim dblVol As Double
    Dim dblNaCl As Double
    Dim dblKCl As Double
    Dim dblVitIntra As Double
    Dim dblLipid As Double
    Dim dblSolu As Double
    Dim dblGewicht As Double
    
    dblGewicht = Range("Gewicht").Value / 10

    Select Case dblGewicht
        Case 2 To 6
            Range("TPN").Value = 2
            
            Range("NaCl").Value = True
            dblNaCl = 6 * dblGewicht
            Range("NaClVol").Value = dblNaCl
            
            Range("KCl").Value = True
            dblKCl = 1 * dblGewicht
            Range("KClVol").Value = dblKCl
            
            dblVitIntra = dblGewicht
            Range("VitIntra").Value = True
            Range("VitIntraVol").Value = IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            Select Case Dag
                Case 1
                    Range("SSTglucose").Value = 2
                
                    dblKCl = 1.5 * dblGewicht
                    Range("KClVol").Value = dblKCl
                    Range("TPNVol") = 15 * dblGewicht
                
                    dblLipid = 6 * dblGewicht / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                    Range("SSTglucose").Value = 3
                
                    Range("TPNVol") = 25 * dblGewicht
                
                    dblLipid = 11 * dblGewicht / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                    Range("SSTglucose").Value = 5
                
                    Range("TPNVol") = 35 * dblGewicht
            
                    dblLipid = 16 * dblGewicht / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            
            End Select
            
            dblVol = (150 * dblGewicht - _
            Range("TPNVol").Value * 2 - _
            dblNaCl * 2 - _
            dblKCl * 2 - _
            dblLipid * 24) / 24
            
            If dblVol < 5 Then
                Range("SSTstand").Value = dblVol * 10
            ElseIf dblVol < 146 Then
                Range("SSTstand").Value = dblVol + 45
            Else
                Range("SSTstand").Value = (dblVol + 125) / 5
            End If
            
        Case 7 To 15
            Range("TPN").Value = 2
            
            Range("NaCl").Value = True
            dblNaCl = 6 * dblGewicht
            Range("NaClVol").Value = dblNaCl
            
            Range("KCl").Value = True
            dblKCl = 1.5 * dblGewicht
            Range("KClVol").Value = dblKCl
            
            dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
            Range("VitIntra").Value = True
            Range("VitIntraVol").Value = IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            Range("SoluVit").Value = True
            Range("SoluVitVol").Value = IIf(dblSolu < 5, dblSolu * 10, dblSolu + 45)
            
            Select Case Dag
                Case 1
                    Range("SSTglucose").Value = 2
                
                    dblKCl = 2 * dblGewicht
                    Range("KClVol").Value = dblKCl
                    Range("TPNVol") = 10 * dblGewicht
                
                    dblLipid = (5 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                    Range("SSTglucose").Value = 6
                
                    Range("TPNVol") = 20 * dblGewicht
                
                    dblLipid = (10 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                    Range("SSTglucose").Value = 8
                
                    Range("TPNVol") = 25 * dblGewicht
            
                    dblLipid = (15 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            End Select
            
            dblVol = (90 * dblGewicht + _
            ((15 - dblGewicht) / 8) * 20 * dblGewicht - _
            Range("TPNVol").Value * 2 - _
            dblNaCl * 2 - _
            dblKCl * 2 - _
            dblLipid * 24) / 24
            
            If dblVol < 5 Then
                Range("SSTstand").Value = dblVol * 10
            ElseIf dblVol < 146 Then
                Range("SSTstand").Value = dblVol + 45
            Else
                Range("SSTstand").Value = (dblVol + 125) / 5
            End If
    
        Case 16 To 30
            Range("TPN").Value = 2
            
            Range("NaCl").Value = True
            dblNaCl = 6 * dblGewicht
            Range("NaClVol").Value = dblNaCl
            
            Range("KCl").Value = True
            dblKCl = 1.5 * dblGewicht
            Range("KClVol").Value = dblKCl
            
            dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
            Range("VitIntra").Value = True
            Range("VitIntraVol").Value = IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            Range("SoluVit").Value = True
            Range("SoluVitVol").Value = IIf(dblSolu < 5, dblSolu * 10, dblSolu + 45)
            
            Range("Peditrace").Value = 15
            
            Select Case Dag
                Case 1
                    Range("SSTglucose").Value = 2
                
                    dblKCl = 2 * dblGewicht
                    Range("KClVol").Value = dblKCl
                    Range("TPNVol") = 10 * dblGewicht
                
                    dblLipid = (5 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                    Range("SSTglucose").Value = 6
                
                    Range("TPNVol") = 15 * dblGewicht
                
                    dblLipid = (10 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                    Range("SSTglucose").Value = 8
                
                    Range("TPNVol") = 20 * dblGewicht
            
                    dblLipid = (15 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            End Select
            
            dblVol = (70 * dblGewicht + _
            ((30 - dblGewicht) / 14) * 10 * dblGewicht - 15 - _
            Range("TPNVol").Value * 2 - _
            dblNaCl * 2 - _
            dblKCl * 2 - _
            dblLipid * 24) / 24
            
            If dblVol < 5 Then
                Range("SSTstand").Value = dblVol * 10
            ElseIf dblVol < 146 Then
                Range("SSTstand").Value = dblVol + 45
            Else
                Range("SSTstand").Value = (dblVol + 125) / 5
            End If
        
        Case 31 To 50
            Range("TPN").Value = 2
            
            Range("NaCl").Value = True
            dblNaCl = 6 * dblGewicht
            Range("NaClVol").Value = dblNaCl
            
            Range("KCl").Value = True
            dblKCl = 1.5 * dblGewicht
            Range("KClVol").Value = dblKCl
            
            dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
            Range("VitIntra").Value = True
            Range("VitIntraVol").Value = IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            Range("SoluVit").Value = True
            Range("SoluVitVol").Value = IIf(dblSolu < 5, dblSolu * 10, dblSolu + 45)
            
            Range("Peditrace").Value = 15
            
            Select Case Dag
                Case 1
                    Range("SSTglucose").Value = 2
                
                    dblKCl = 2 * dblGewicht
                    Range("KClVol").Value = dblKCl
                    Range("TPNVol") = 5 * dblGewicht
                
                    dblLipid = (3 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                    Range("SSTglucose").Value = 6
                
                    Range("TPNVol") = 8 * dblGewicht
                
                    dblLipid = (6 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                    Range("SSTglucose").Value = IIf(dblGewicht > 35, 9, 7)
                
                    Range("TPNVol") = 12 * dblGewicht
            
                    dblLipid = (10 * dblGewicht + dblVitIntra + dblSolu) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            End Select
            
            dblVol = (50 * dblGewicht + _
            ((50 - dblGewicht) / 19) * 20 * dblGewicht - 15 - _
            Range("TPNVol").Value * 2 - _
            dblNaCl * 2 - _
            dblKCl * 2 - _
            dblLipid * 24) / 24
            
            If dblVol < 5 Then
                Range("SSTstand").Value = dblVol * 10
            ElseIf dblVol < 146 Then
                Range("SSTstand").Value = dblVol + 45
            Else
                Range("SSTstand").Value = (dblVol + 125) / 5
            End If
        Case Else
            Range("TPN").Value = 2
            
            Range("Nacl").Value = False
            Range("KCl").Value = False
            
            dblVitIntra = IIf(dblGewicht > 10, 10, dblGewicht)
            Range("VitIntra").Value = True
            Range("VitIntraVol").Value = IIf(dblVitIntra < 5, dblVitIntra * 10, dblVitIntra + 45)
            
            dblSolu = IIf(dblGewicht > 10, 10, dblGewicht)
            Range("SoluVit").Value = True
            Range("SoluVitVol").Value = IIf(dblSolu < 5, dblSolu * 10, dblSolu + 45)
            
            Range("Peditrace").Value = 15
            Range("SSTGlucose").Value = 2
            
            Select Case Dag
                Case 1
                
                    Range("TPNVol") = 700
                
                    dblLipid = (150 + 20) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 2
                
                    Range("TPNVol") = 1000
                
                    dblLipid = (300 + 20) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
                
                Case 3
                
                    Range("TPNVol") = 1500
            
                    dblLipid = (500 + 20) / 24
                    Range("LipidenStand").Value = IIf(dblLipid < 5, dblLipid * 10, dblLipid + 45)
            End Select
            
            dblVol = 0
            
            If dblVol < 5 Then
                Range("SSTstand").Value = dblVol * 10
            ElseIf dblVol < 146 Then
                Range("SSTstand").Value = dblVol + 45
            Else
                Range("SSTstand").Value = (dblVol + 125) / 5
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

Public Sub ClearPat2()
    'TODO: Samenvoegen met clearPat
    Dim i As Integer, vResp As Variant
    Dim shtPatienten As Worksheet

    On Error GoTo Hell

    vResp = MsgBox("Afspraken echt verwijderen?", vbYesNo, "Informedica")

    If vResp = vbYes Then
        Application.Cursor = xlWait
        With shtPatienten
            For i = 2 To .Range("A1").CurrentRegion.Rows.Count
                If Not i = 4 Then _
                    Range(.Cells(i, 1).Value).Formula = .Cells(i, 3).Formula
            Next i
        End With
        Application.Cursor = xlDefault
    End If

    SetApplTitle

    Exit Sub

Hell:

    Resume Next

End Sub

Public Sub Afsluiten2()
    Dim nPiTeller As Integer
    Dim shtTemp As Worksheet

    'TODO: Samenvoegen met Afsluiten
    Application.Cursor = xlWait
    Application.DisplayAlerts = False

On Error GoTo Hell:

'   Werkbalken weer herstellen
    For nPiTeller = 1 To Toolbars.Count
        If shtTemp.Cells(20, nPiTeller).Value Then
            Toolbars(nPiTeller).Visible = True
        End If
    Next nPiTeller

'   Menubalk activeren
    For nPiTeller = 1 To MenuBars(xlWorksheet).Menus.Count
        MenuBars(xlWorksheet).Menus(nPiTeller).Enabled = True
    Next nPiTeller
    
    With ActiveWindow
        .DisplayGridlines = True
        .DisplayHeadings = True
        .DisplayOutline = True
        .DisplayZeros = True
        .DisplayWorkbookTabs = True
    End With
 
    Toolbars("Afspraken").Visible = False
    
    With Application
         .Caption = vbNullString
         .DisplayFormulaBar = True
         .Cursor = xlDefault
         .Quit
    End With
    
Exit Sub
    
Hell:

    Resume Next:

End Sub

Sub Openen2()
'TODO: Samenvoegen met Openen
'''''    Dim bPuGlToolbars() As Boolean
'''''
'''''    On Error Resume Next
'''''
'''''    Application.Cursor = xlWait
'''''
'''''    InterfaceBladenBeveiligen
'''''    RekenBladenVerbergen
'''''    shtAfspraken.Activate
''''''   Knoppen en balken verwijderen
'''''    With Application
'''''         .Caption = "Informedica 2000 NICU programma "
'''''         .DisplayFormulaBar = False
'''''         .DisplayStatusBar = False
'''''         .DisplayScrollBars = True
'''''         .DisplayFormulaBar = False
'''''    End With
'''''
''''''   Werkbalken weghalen behalve de Afspraken balk
'''''    ReDim bPuGlToolbars(Toolbars.Count)
'''''    For nPiTeller = 1 To Toolbars.Count
'''''        If Toolbars(nPiTeller).Visible = True Then
'''''            shtTemp.Cells(20, nPiTeller).Value = True
'''''        Else
'''''            shtTemp.Cells(20, nPiTeller).Value = False
'''''        End If
'''''        If Toolbars(nPiTeller).Name <> "NICU" Then
'''''            Toolbars(nPiTeller).Visible = False
'''''        End If
'''''    Next nPiTeller
'''''    Toolbars("'NICU").Visible = True
'''''
''''''   Menubalk leegmaken
'''''    For nPiTeller = 1 To MenuBars(xlWorksheet).Menus.Count
'''''        MenuBars(xlWorksheet).Menus(nPiTeller).Enabled = False
'''''    Next nPiTeller
'''''
'''''    With ActiveWindow
'''''        .DisplayHorizontalScrollBar = False
'''''        .DisplayVerticalScrollBar = True
'''''        .DisplayGridlines = False
'''''        .DisplayZeros = False
'''''        .DisplayWorkbookTabs = False
'''''        .DisplayHeadings = False
'''''    End With
'''''
'''''    Application.Cursor = xlDefault
'''''    PaPatientenLijst

End Sub

Public Sub SluitBed2()
'TODO: Samenvoegen met BeSluitBed
''''    Dim sFileName As String, sBookName As String, sRange As String, sBed As String
''''    Dim sPrompt As String, vReply As Variant
''''
''''    Dim colPatienten As Collection
''''    Dim oFrmPatientLijst As frmPatLijst
''''
''''    sBed = Range("_Bed").Formula
''''    sFileName = "\\fs04.ds.umcutrecht.nl\PGGroups$\Groups\Infuusbrief\Patient" + sBed + ".xls"
''''    sBookName = "Patient" + sBed + ".xls"
''''
''''    sPrompt = "Patient " & Range("_VoorNaam").Value & ", " & Range("_AchterNaam") _
''''    & " opslaan op bed: " & sBed & "?"
''''    vReply = MsgBox(prompt:=sPrompt, Buttons:=vbYesNo, Title:="Informedica 200")
''''
''''    If vReply = vbYes Then
''''        Application.Cursor = xlWait
''''        If bPuBedOpslaan(sFileName, sBookName) Then
''''            MsgBox "Patient is opgeslagen", vbInformation, "Informedica"
''''        End If
''''        Application.Cursor = xlDefault
''''    Else
''''        vReply = MsgBox("Op een ander bed opslaan?", vbYesNo, "Informedica")
''''        If vReply = vbYes Then
''''            Set colPatienten = oPuPatientenCollectie
''''            Set oFrmPatientLijst = New frmPatLijst
''''            oFrmPatientLijst.Caption = "Selecteer de patient die vervangen moet worden ..."
''''            With oFrmPatientLijst.lstPatienten
''''                .Clear
''''                For nPiTeller = 1 To colPatienten.Count
''''                    .AddItem colPatienten(nPiTeller)
''''                Next nPiTeller
''''                oFrmPatientLijst.Show
''''                If .ListIndex > -1 Then
''''                    sBed = VBA.Left$(.Text, 3)
''''                    Range("_Bed").Value = sBed
''''                    Set colPatienten = Nothing
''''                    Set oFrmPatientLijst = Nothing
''''                    Call BeSluitBed
''''                Else
''''                    Set colPatienten = Nothing
''''                    Set oFrmPatientLijst = Nothing
''''                    Exit Sub
''''                End If
''''            End With
''''        End If
''''    End If
''''
''''    SetApplTitle
''''    shtAfspraken.Select

End Sub

Public Sub SetApplTitle()

    Application.Caption = "Patient: " & Range("_VoorNaam") & " " & Range("_AchterNaam") & ", op bed " & Range("_bed")

End Sub

Public Sub TPNAdviesNEO()

    Range("_DagKeuze").Value = IIf(Range("Dag").Value < 4, 1, 2)
    Range("_IntakePerKg").Value = 5000
    Range("_IntraLipid").Value = 5000
    Range("_NaCl").Value = 5000
    Range("_KCl").Value = 5000
    Range("_CaCl2").Value = 5000
    Range("_MgCl2").Value = 5000
    Range("_SoluVit").Value = 5000
    Range("_Primene").Value = 5000
    Range("_NICUMix").Value = 5000
    Range("_SSTB").Value = 5000
    
    ModSheets.GoToSheet shtNeoGuiAfspraken, "A9"

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


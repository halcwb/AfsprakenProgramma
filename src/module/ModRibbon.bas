Attribute VB_Name = "ModRibbon"
Option Explicit

Public Sub ButtonOnAction(ctrlMenuItem As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Application.ScreenUpdating = False
    
    Select Case ctrlMenuItem.ID
    
        'grpAfspraken                                       ' -- AFSPRAKEN --
        
        Case "btnClose"                                     ' -> Programma Afsluiten
            Afsluiten
        
        Case "btnClear"                                     ' -> Alles Verwijderen
            ClearPatient True
            SelectPedOrNeoStartSheet
        
        'grpBedden                                          ' -- BEDDEN --
        
        Case "btnOpenBed"                                   ' -> Bed Openen
            OpenPatientLijst
        
        Case "btnSaveBed"                                   ' -> Bed Opslaan
            ModBedden.SluitBed
            SelectPedOrNeoStartSheet
        
        Case "btnEnterPatient"                              ' -> Patient Gegevens
            NieuwePatient
            
        'grpPediatrie                                       ' -- PEDIATRIE --
        
        Case "btnPedMedIV"                                  ' -> Continue IV Medicatie
            ModSheets.GoToSheet shtPedGuiMedIV, "A6"
        
        Case "btnPedMedDisc"                                ' -> Discontinue Medicatie
            ModSheets.GoToSheet shtPedGuiMedDisc, "A6"
        
        Case "btnPedIVandPM"                                ' -> Lijnen en PM
            ModSheets.GoToSheet shtPedGuiPMenIV, "A6"
        
        Case "btnPedEntTPN"                                 ' -> Voeding en TPN
            ModSheets.GoToSheet shtPedGuiEntTPN, "A6"
        
        Case "btnPedLab"                                    ' -> Lab Aanvragen
            ModSheets.GoToSheet shtPedGuiLab, "A1"
        
        Case "btnPedExtra"                                  ' -> Afspraken en Controles
            ModSheets.GoToSheet shtPedGuiAfsprExta, "A6"
            
        'grpNeonatologie                                    ' -- NEONATOLOGIE --
        
        Case "btnNeoMedIV"                                  ' -> Infuusbrief
            ModSheets.GoToSheet shtNeoGuiAfspraken, "A9"
        
        Case "btnNeoMedDisc"                                ' -> Discontinue Medicatie
            ModSheets.GoToSheet shtPedGuiMedDisc, "A6"
        
        Case "btnNeoExtra"                                  ' -> Afspraken en Controles
            ModSheets.GoToSheet shtNeoGuiAfsprExtra, "A6"
        
        Case "btnNTPNadvies"
            TPNAdviesNEO
        
        Case "btnNeoLab"                                    ' -> Lab Aanvragen
            ModSheets.GoToSheet shtNeoGuiLab, "A1"
        
        Case "btnNeo1700"                                   ' -> Infuusbrief 17:00
            ModSheets.GoToSheet shtNeoGuiAfspr1700, "A9"
        
        'grpRemoveData                                      ' -- ACTIES --
                
        Case "btnRemoveLab"                                 ' Lab Verwijderen
            VerwijderLab
            SelectPedOrNeoLabSheet
        
        Case "btnRemoveExtra"                               ' Afspraken Controles Verwijderen
            VerwijderAanvullendeAfspraken
            SelectPedOrNeoAfsprExtraSheet
        
        ' grpNeoActions                                     ' -- INFUUSBRIEF OVERZETTEN --
        
        Case "btnCopy1700"                                  ' -> 17:00 uur Overzetten
            ModAfspraken1700.CopyToActueel
        
        Case "btnCopyCurrent"                               ' -> Actueel Overzetten
            ModAfspraken.AfsprakenOvernemen
            
        'grpPedPrint                                        ' -- PRINT PEDIATRIE --
        
        Case "btnPedPrintAcuut"                             ' -> Acute Blad
            ModSheets.GoToSheet shtPedGuiAcuut, "A1"
        
        Case "btnPedPrintMedIV"                             ' -> Medicatie IV
            ModSheets.GoToSheet shtPedPrtAfspr, "A1"
        
        Case "btnPedPrintMedDisc"                           ' -> Medicatie Discontinu
            ModSheets.GoToSheet shtPedPrtMedDisc, "A1"
        
        Case "btnPedPrintTPN"                               ' -> TPN Brief
            SelectPedTPNPrint
            
        'grpNeoPrint                                        ' -- PRINT NEO ---
        
        Case "btnNeoPrintAcuut"                             ' -> Acute Blad
            ModSheets.GoToSheet shtNeoGuiAcuut, "A1"
        
        Case "btnNeoPrintMedIV"                             ' -> Infuus Brief
            ModSheets.GoToSheet shtNeoPrtAfspr, "A1"
        
        Case "btnNeoPrintMedDisc"                           ' -> Medicatie Discontinu
            ModSheets.GoToSheet shtNeoPrtMedDisc, "A1"
        
        Case "btnNTPN"
            TPNAdviesNEO
        
        Case "btnNeoPrintApoth"                             ' -> Apotheek
            ModSheets.GoToSheet shtNeoPrtApoth, "A1"
        
        Case "btnNeoPrintWerkbr"                            ' -> Werkbrief
            ModSheets.GoToSheet shtNeoPrtWerkbr, "A1"
            
        'grpDevelopment                                     ' -- DEVELOPMENT --
        
        Case "btnDevMode"                                   ' -> Development Mode
            SetToDevelopmentMode
        
        Case "btnToggleLogging"                             ' -> Toggle Logging
    
        Case "btnRangeNames"                                ' -> Range Names
            ModMenuItems.GiveNameToRange
        
        Case Else
            MsgBox ctrlMenuItem.ID & " has no select case", vbCritical
            
        
    End Select

    ' Waarom wordt dit aangeroepen na een menu item keuze???
    ' HideBars
    
    Application.ScreenUpdating = True
    
End Sub

Sub GetVisiblePed(control As IRibbonControl, ByRef blnVisible)

    Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_PELI_FOLDERNAME)) > 0 Or InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Sub GetVisibleNeo(control As IRibbonControl, ByRef blnVisible)
    
    Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_NEO_FOLDERNAME)) > 0 Or InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Sub GetVisibleDevelopment(control As IRibbonControl, ByRef blnVisible)

    Dim strPath As String

    strPath = LCase(Application.ActiveWorkbook.Path)
    
    If InStr(1, strPath, LCase(CONST_DEVELOP_FOLDERNAME)) > 0 Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub


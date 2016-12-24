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
            CloseAfspraken
        
        Case "btnClear"                                     ' -> Alles Verwijderen
            ModPatient.ClearPatient True
            ModSheet.SelectPedOrNeoStartSheet
        
        'grpBedden                                          ' -- BEDDEN --
        
        Case "btnOpenBed"                                   ' -> Bed Openen
            ModBed.OpenBed ModPatient.OpenPatientLijst("Kies een patient...")
        
        Case "btnSaveBed"                                   ' -> Bed Opslaan
            ModBed.SluitBed
            ModSheet.SelectPedOrNeoStartSheet
        
        Case "btnEnterPatient"                              ' -> Patient Gegevens
            ModPatient.EnterPatient
            
        'grpPediatrie                                       ' -- PEDIATRIE --
        
        Case "btnPedMedIV"                                  ' -> Continue IV Medicatie
            ModSheet.GoToSheet shtPedGuiMedIV, "A6"
        
        Case "btnPedMedDisc"                                ' -> Discontinue Medicatie
            ModSheet.GoToSheet shtPedGuiMedDisc, "A6"
        
        Case "btnPedIVandPM"                                ' -> Lijnen en PM
            ModSheet.GoToSheet shtPedGuiPMenIV, "A6"
        
        Case "btnPedEntTPN"                                 ' -> Voeding en TPN
            ModSheet.GoToSheet shtPedGuiEntTPN, "A6"
        
        Case "btnPedLab"                                    ' -> Lab Aanvragen
            ModSheet.GoToSheet shtPedGuiLab, "A1"
        
        Case "btnPedExtra"                                  ' -> Afspraken en Controles
            ModSheet.GoToSheet shtPedGuiAfsprExta, "A6"
            
        'grpNeonatologie                                    ' -- NEONATOLOGIE --
        
        Case "btnNeoMedIV"                                  ' -> Infuusbrief
            ModSheet.GoToSheet shtNeoGuiAfspraken, "A9"
        
        Case "btnNeoMedDisc"                                ' -> Discontinue Medicatie
            ModSheet.GoToSheet shtPedGuiMedDisc, "A6"
        
        Case "btnNeoExtra"                                  ' -> Afspraken en Controles
            ModSheet.GoToSheet shtNeoGuiAfsprExtra, "A6"
        
        Case "btnNTPNadvies"
            TPNAdviesNEO
        
        Case "btnNeoLab"                                    ' -> Lab Aanvragen
            ModSheet.GoToSheet shtNeoGuiLab, "A1"
        
        Case "btnNeo1700"                                   ' -> Infuusbrief 17:00
            ModSheet.GoToSheet shtNeoGuiAfspr1700, "A9"
        
        'grpRemoveData                                      ' -- ACTIES --
                
        Case "btnRemoveLab"                                 ' Lab Verwijderen
            VerwijderLab
            ModSheet.SelectPedOrNeoLabSheet
        
        Case "btnRemoveExtra"                               ' Afspraken Controles Verwijderen
            VerwijderAanvullendeAfspraken
            ModSheet.SelectPedOrNeoAfsprExtraSheet
        
        ' grpNeoActions                                     ' -- INFUUSBRIEF OVERZETTEN --
        
        Case "btnCopy1700"                                  ' -> 17:00 uur Overzetten
            ModAfspraken1700.CopyToActueel
        
        Case "btnCopyCurrent"                               ' -> Actueel Overzetten
            ModAfspraken.AfsprakenOvernemen
            
        'grpPedPrint                                        ' -- PRINT PEDIATRIE --
        
        Case "btnPedPrintAcuut"                             ' -> Acute Blad
            ModSheet.GoToSheet shtPedGuiAcuut, "A1"
        
        Case "btnPedPrintMedIV"                             ' -> Medicatie IV
            ModSheet.GoToSheet shtPedPrtAfspr, "A1"
        
        Case "btnPedPrintMedDisc"                           ' -> Medicatie Discontinu
            ModSheet.GoToSheet shtPedPrtMedDisc, "A1"
        
        Case "btnPedPrintTPN"                               ' -> TPN Brief
            SelectPedTPNPrint
            
        'grpNeoPrint                                        ' -- PRINT NEO ---
        
        Case "btnNeoPrintAcuut"                             ' -> Acute Blad
            ModSheet.GoToSheet shtNeoGuiAcuut, "A1"
        
        Case "btnNeoPrintMedIV"                             ' -> Infuus Brief
            ModSheet.GoToSheet shtNeoPrtAfspr, "A1"
        
        Case "btnNeoPrintMedDisc"                           ' -> Medicatie Discontinu
            ModSheet.GoToSheet shtNeoPrtMedDisc, "A1"
        
        Case "btnNTPN"
            TPNAdviesNEO
        
        Case "btnNeoPrintApoth"                             ' -> Apotheek
            ModSheet.GoToSheet shtNeoPrtApoth, "A1"
        
        Case "btnNeoPrintWerkbr"                            ' -> Werkbrief
            ModSheet.GoToSheet shtNeoPrtWerkbr, "A1"
            
        'grpDevelopment                                     ' -- DEVELOPMENT --
        
        Case "btnDevMode"                                   ' -> Development Mode
            ModApplication.SetToDevelopmentMode
        
        Case "btnToggleLogging"                             ' -> Toggle Logging
            ModSetting.ToggleLogging
    
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

    Dim strPath, strPedDir As String
    Dim blnIsDevelop

    blnIsDevelop = ModSetting.IsDevelopmentMode()
    strPath = Application.ActiveWorkbook.Path
    strPedDir = ModSetting.GetPedDir()
    
    If ModString.ContainsCaseInsensitive(strPath, strPedDir) Or blnIsDevelop Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Sub GetVisibleNeo(control As IRibbonControl, ByRef blnVisible)
    
    Dim strPath, strPedDir As String
    Dim blnIsDevelop

    blnIsDevelop = ModSetting.IsDevelopmentMode()
    strPath = Application.ActiveWorkbook.Path
    strPedDir = ModSetting.GetNeoDir()
    
    If ModString.ContainsCaseInsensitive(strPath, strPedDir) Or blnIsDevelop Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Sub GetVisibleDevelopment(control As IRibbonControl, ByRef blnVisible)

    blnVisible = ModSetting.IsDevelopmentMode()
    
End Sub


Attribute VB_Name = "ModRibbon"
Option Explicit

' Module to handle ribbon events
' Note: blnVisible in GetVisible functions is not a boolean!! Using blnVisible As Boolean will result in a type mismatch!

Public Sub ButtonOnAction(ctrlMenuItem As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    
    Select Case ctrlMenuItem.Id
    
        'grpAfspraken                                       ' -- AFSPRAKEN --
        
        Case "btnAbout"                                     ' -> Programma Voorblad
            shtGlobGuiFront.Select
        
        Case "btnClose"                                     ' -> Programma Afsluiten
            App_CloseApplication
        
        Case "btnClear"                                     ' -> Alles Verwijderen
            ModProgress.StartProgress "Patient Data Verwijderen"
            Patient_ClearAll True, True
            ModProgress.FinishProgress
        
        'grpBedden                                          ' -- PATIENTEN --
        
        Case "btnPatientList"                               ' -> Open patient lijst
            If Setting_UseDatabase Then
                Patient_OpenPatientAndAsk
            Else
                ModBed.Bed_OpenBed
            End If
            ModSheet.SelectPedOrNeoStartSheet True
        
        Case "btnSavePatient"                               ' -> Patient opslaan
            If Setting_UseDatabase Then
                Patient_SavePatient
            Else
                ModBed.Bed_CloseBed True
            End If
            ModSheet.SelectPedOrNeoStartSheet True
        
        Case "btnEnterPatient"                              ' -> Patient Gegevens
            ModPatient.Patient_EnterDetails
            
        'grpPediatrie                                       ' -- PEDIATRIE --
        
        Case "btnPedMedIV"                                  ' -> Continue IV Medicatie
            ModSheet.GoToSheet shtPedGuiMedIV, "A6"
        
        Case "btnPedMedDisc"                                ' -> Discontinue Medicatie
            ModSheet.GoToSheet shtGlobGuiMedDisc, "A6"
        
        Case "btnPedIVandPM"                                ' -> Lijnen en PM
            ModSheet.GoToSheet shtPedGuiLijnPM, "A6"
        
        Case "btnPedEntTPN"                                 ' -> Voeding en TPN
            ModSheet.GoToSheet shtPedGuiEntTPN, "A6"
        
        Case "btnPedLab"                                    ' -> Lab Aanvragen
            ModSheet.GoToSheet shtPedGuiLab, "A1"
        
        Case "btnPedExtra"                                  ' -> Afspraken en Controles
            ModSheet.GoToSheet shtPedGuiAfspr, "A6"
            
        'grpNeonatologie                                    ' -- NEONATOLOGIE --
        
        Case "btnNeoMedIV"                                  ' -> Infuusbrief
            ModNeoInfB.NeoInfB_SelectInfB False, True
        
        Case "btnNeoMedDisc"                                ' -> Discontinue Medicatie
            ModSheet.GoToSheet shtGlobGuiMedDisc, "A6"
        
        Case "btnNeoExtra"                                  ' -> Afspraken en Controles
            ModSheet.GoToSheet shtNeoGuiAfspr, "A6"
        
        Case "btnNTPNadvies"
            ModNeoInfB.NeoInfB_TPNAdvice
        
        Case "btnNeoLab"                                    ' -> Lab Aanvragen
            ModSheet.GoToSheet shtNeoGuiLab, "A1"
        
        Case "btnNeo1700"                                   ' -> Infuusbrief 17:00
            ModNeoInfB.NeoInfB_SelectInfB True, True
        
        'grpRemoveData                                      ' -- ACTIES --
                
        Case "btnRemoveAll"                                 ' Alle Afspraken Verwijderen
            ClearAll
            ModSheet.SelectPedOrNeoAfsprSheet
            
        Case "btnRemoveLab"                                 ' Lab Verwijderen
            ClearLab
            ModSheet.SelectPedOrNeoLabSheet
        
        Case "btnRemoveExtra"                               ' Afspraken Controles Verwijderen
            ClearAfspraken
            ModSheet.SelectPedOrNeoAfsprSheet
        
        ' grpNeoActions                                     ' -- INFUUSBRIEF OVERZETTEN --
        
        Case "btnCopy1700"                                  ' -> 17:00 uur Overzetten
            ModNeoInfB.NeoInfB_ShowFormCopy1700ToAct
        
        Case "btnCopyCurrent"                               ' -> Actueel Overzetten
            ModNeoInfB.NeoInfB_CopyActTo1700
            
        'grpPedPrint                                        ' -- PRINT PEDIATRIE --
        
        Case "btnPedPrintAcuut"                             ' -> Acute Blad
            ModPedPrint.PedPrint_PrintAcuteBlad True
        
        Case "btnPedPrintMedIV"                             ' -> Medicatie IV
            ModPedPrint.PedPrint_PrintMedicatieCont True
        
        Case "btnPedPrintMedDisc"                           ' -> Medicatie Discontinu
            ModPedPrint.PedPrint_PrintMedicatieDisc True
        
        Case "btnPedPrintTPN"                               ' -> TPN Brief
            ModPedPrint.PedPrint_PrintTPN True
            
        Case "btnPedSendTPN"                                ' -> TPN Brief
            ModPedPrint.PedPrint_SendTPN
        
        Case "btnPedSendMedVal"                             ' -> Medicatie Validatie
            ModMedDisc.SendApotheekMedDiscValidation
        
        'grpNeoPrint                                        ' -- PRINT NEO ---
        
        Case "btnNeoPrintAcuut"                             ' -> Acute Blad
            ModNeoPrint.NeoPrint_PrintAcuteBlad True
        
        Case "btnNeoPrintMedIV"                             ' -> Infuus Brief
            ModNeoPrint.NeoPrint_PrintMedicatieCont True
        
        Case "btnNeoPrintMedDisc"                           ' -> Medicatie Discontinu
            ModNeoPrint.NeoPrint_PrintMedicatieDisc True
              
        Case "btnNeoSendApoth"                              ' -> Apotheek Versturen
            ModNeoPrint.SendApotheekWerkBrief
        
        Case "btnNeoPrintApoth"                             ' -> Apotheek Printen
            ModNeoPrint.PrintApotheekWerkBrief
        
        Case "btnNeoPrintWerkbr"                            ' -> Werkbrief
            ModNeoPrint.PrintNeoWerkBrief
            
        Case "btnNeoSendMedVal"                             ' -> Medicatie Validatie
            ModMedDisc.SendApotheekMedDiscValidation
        
        'grpDevelopment                                     ' -- DEVELOPMENT --
        
        Case "btnDevMode"                                   ' -> Development Mode
            ModApplication.App_ToggleDevelopmentMode
        
        Case "btnToggleLogging"                             ' -> Toggle Logging
            ModSetting.ToggleLogging
    
        Case "btnRangeNames"                                ' -> Name Range
            ModRange.GiveNameToRange
            
        Case "btnWriteNames"                                ' -> Write Names
            ModRange.WriteNamesToGlobNames
            
        Case "btnReplaceNames"                              ' -> Replace Names
            ModRange.ReplaceRangeNames
            
        Case "btnRefreshPatientData"                        ' -> Refresh Patient Data
            ModRange.RefreshPatientData
            
        Case "btnExportSource"                              ' -> Export To Source
            ModUtils.ExportForSourceControl
        
        'grpFB                                              ' -- ADMISTRATION FUNCTIONEEL BEHEER --
        
        Case "btnOpenSettings"                              ' -> Instellingen
             ModMessage.ShowMsgBoxExclam "Nog niet geimplementeerd"
        
        Case "btnSetColors"                                 ' -> Kleuren Instellen
             ModAdmin.Admin_ShowColorPicker
        
        Case "btnCreatePedData"                             ' -> Pediatrie DataFiles
             ModAdmin.Admin_SetUpPedDataDir
        
        Case "btnCreateNeoData"                             ' -> Neonatologie DataFiles
             ModAdmin.Admin_SetUpNeoDataDir
        
        Case "btnOpenLogFiles"                              ' -> Log files openen
             ModAdmin.Admin_OpenLogFiles
        
        Case "btnRefreshMedOpdr"                            ' -> MetaVision Medicatie Opdrachten Verversen
             ModMetaVision.MetaVision_GetMedicatieOpdrachten
        
        'grpFB                                              ' -- ADMISTRATION APOTHEEK --
        
        Case "btnNeoMedCont"                                ' -> Beheer Continue Medicatie Neo
             ModAdmin.Admin_MedContNeoConfig
             
        Case "btnParent"                                    ' -> Beheer ParEnterale Vloeistoffen
             ModAdmin.Admin_ParEntConfig
        
        Case "btnMedDisc"                                   ' -> Beheer Discontinue medicatie
             ModDatabase.Database_ImportConfigMedDisc
                
        Case "btnNeoMedContImport"                          ' -> Beheer Continue Medicatie Neo Exporteren
            ModAdmin.Admin_MedContNeoImport
        
        Case "btnNeoMedContExport"                          ' -> Beheer Continue Medicatie Neo Exporteren
            ModAdmin.Admin_MedContNeoExport
        
        Case "btnPedMedContImport"                          ' -> Beheer Continue Medicatie Ped Importeren
             ModAdmin.Admin_MedContPedImport
        
        Case "btnPedMedContExport"                          ' -> Beheer Continue Medicatie Ped Exporteren
            ModAdmin.Admin_MedContPedExport
                
        Case "btnParentImport"                              ' -> Beheer Parenteralia Ped Importeren
             ModAdmin.Admin_ParentImport
        
        Case "btnParentExport"                              ' -> Beheer Parenteralia Ped Exporteren
            ModAdmin.Admin_ParentExport
        
        Case "btnDiscMedImport"                             ' -> Beheer Discontinue Medicatie Ped Importeren
             ModAdmin.Admin_MedDiscImport
        
        Case "btnPedMedContExport"                          ' -> Beheer Discontinue Medicatie Ped Exporteren
            ModAdmin.Admin_MedDiscExport
        
        'grpFB                                              ' -- ACCEPTATIE TESTS --
        
        Case "btnNeoMedContTests"                           ' -> Neo Continue Medicatie
             ModNeoInfB_Tests.Test_NeoInfB_ContMed
        
        Case "btnNeoMedPrintTests"                          ' -> Neo Werkbrief en Apotheek prints
             ModNeoInfB_Tests.Test_NeoInfB_Print
                
        Case Else
            ModMessage.ShowMsgBoxError ctrlMenuItem.Id & " has no select case"
            
    End Select
    
End Sub

Public Sub RibbonOnLoad(ByRef objRibbon As IRibbonUI)

    ModLog.LogInfo "RibbonOnLoad"

End Sub

Public Sub GetVisiblePed(ByRef ctrContr As IRibbonControl, ByRef blnVisible As Variant)

    Dim blnIsDevelop As Boolean
    Dim blnIsPed As Boolean
    
    blnIsDevelop = ModSetting.IsDevelopmentDir()
    blnIsPed = MetaVision_IsPediatrie()
    
    If blnIsPed Or blnIsDevelop Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Public Sub GetVisibleNeo(ByRef ctrContr As IRibbonControl, ByRef blnVisible As Variant)
    
    Dim blnIsDevelop As Boolean
    Dim blnIsNeo As Boolean

    blnIsDevelop = ModSetting.IsDevelopmentDir()
    blnIsNeo = MetaVision_IsNeonatologie()
    
    If blnIsNeo Or blnIsDevelop Then
        blnVisible = True
    Else
        blnVisible = False
    End If
    
End Sub

Public Sub GetVisibleDevelopment(ByRef ctrContr As IRibbonControl, ByRef blnVisible As Variant)

    Dim strUserType As String
    
    strUserType = ModRange.GetRangeValue("_User_Type", vbNullString)
    blnVisible = ModSetting.IsDevelopmentDir() Or strUserType = "Beheerders"
    
End Sub

Public Sub GetVisibleAdmin(ByRef ctrContr As IRibbonControl, ByRef blnVisible As Variant)

    Dim strUserType As String
    
    strUserType = ModRange.GetRangeValue("_User_Type", vbNullString)
    blnVisible = ModSetting.IsDevelopmentDir() Or strUserType = "Beheerders" Or strUserType = "Apotheek"
    
End Sub

Private Sub ClearAll()

    If ModSetting.IsDevelopmentDir Then
        ModPatient.Patient_ClearNeoData
        ModPatient.Patient_ClearPedData
    Else
        If MetaVision_IsNeonatologie() Then ModPatient.Patient_ClearNeoData
        If MetaVision_IsPediatrie() Then ModPatient.Patient_ClearPedData
    End If
    
End Sub

Private Sub ClearLab()
    
    If ModSetting.IsDevelopmentDir Then
        ModNeoLab.NeoLab_Clear
        ModPedLab.PedLab_Clear
    Else
        If MetaVision_IsNeonatologie() Then ModNeoLab.NeoLab_Clear
        If MetaVision_IsPediatrie() Then ModPedLab.PedLab_Clear
    End If
    
End Sub

Private Sub ClearAfspraken()

    If ModSetting.IsDevelopmentDir Then
        ModNeoAfspr.NeoAfspr_Clear
        ModPedAfspr.PedAfspr_Clear
    Else
        If MetaVision_IsNeonatologie() Then ModNeoAfspr.NeoAfspr_Clear
        If MetaVision_IsPediatrie() Then ModPedAfspr.PedAfspr_Clear
    End If

End Sub



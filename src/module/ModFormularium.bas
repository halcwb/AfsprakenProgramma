Attribute VB_Name = "ModFormularium"
Option Explicit

Private m_Formularium As ClassFormularium
Private m_FormConfig As ClassFormConfig

Private Const constFormularium As String = "FormulariumDb.xlsm"
Private Const constFormDbStart As Integer = 2

'--- Formularium ---
Private Const constGPKIndx As Integer = 1
Private Const constATCIndx As Integer = 2
Private Const constMainGroupIndx As Integer = 3
Private Const constSubGroupIndx As Integer = 4

Private Const constGenericIndx As Integer = 5
Private Const constProductIndx As Integer = 6
Private Const constLabelIndx As Integer = 7
Private Const constShapeIndx As Integer = 8
Private Const constRouteIndx As Integer = 9
Private Const constGenericQuantityIndx As Integer = 10
Private Const constGenericQuantityUnitIndx As Integer = 11
Private Const constMultipleQuantityIndx As Integer = 12
Private Const constMultipleQuantityUnitIndx As Integer = 13

Private Const constIndicationsIndx As Integer = 14

Private Const constPICU_MaxSolutionConcIndx As Integer = 23
Private Const constPICU_SolutionIndx As Integer = 24
Private Const constPICU_SolutionVolumeIndx As Integer = 25
Private Const constPICU_MinInfusionTimeIndx As Integer = 26

Private Const constNICU_MaxSolutionConcIndx As Integer = 27
Private Const constNICU_SolutionIndx As Integer = 28
Private Const constNICU_SolutionVolumeIndx As Integer = 29
Private Const constNICU_MinInfusionTimeIndx As Integer = 30

Private Const constDoseFreq_AN As String = "antenoctum"
Private Const constDoseFreq_1D As String = "1 x / dag"
Private Const constDoseFreq_2D As String = "2 x / dag"
Private Const constDoseFreq_3D As String = "3 x / dag"
Private Const constDoseFreq_4D As String = "4 x / dag"
Private Const constDoseFreq_5D As String = "5 x / dag"
Private Const constDoseFreq_6D As String = "6 x / dag"
Private Const constDoseFreq_7D As String = "7 x / dag"
Private Const constDoseFreq_8D As String = "8 x / dag"
Private Const constDoseFreq_9D As String = "9 x / dag"
Private Const constDoseFreq_10D As String = "10 x / dag"
Private Const constDoseFreq_11D As String = "11 x / dag"
Private Const constDoseFreq_12D As String = "12 x / dag"
Private Const constDoseFreq_24D As String = "24 x / dag"
Private Const constDoseFreq_1D2 As String = "1 x / 2 dagen"
Private Const constDoseFreq_1D3 As String = "1 x / 3 dagen"
Private Const constDoseFreq_1U36 As String = "1 x / 36 uur"
Private Const constDoseFreq_1W As String = "1 x / week"
Private Const constDoseFreq_2W As String = "2 x / week"
Private Const constDoseFreq_3W As String = "3 x / week"
Private Const constDoseFreq_4W As String = "4 x / week"
Private Const constDoseFreq_1W2 As String = "1 x / 2 weken"
Private Const constDoseFreq_1W4 As String = "1 x / 4 weken"
Private Const constDoseFreq_1W12 As String = "1 x / 12 weken"
Private Const constDoseFreq_1W13 As String = "1 x / 13 weken"
Private Const constDoseFreq_1M As String = "1 x / 1 maand"

Private Const constDoseFreqFormula As String = "=IF({W}{2}<>0,{W}$1,"""")"

Private Const constDoseGenericIndx As Integer = 1
Private Const constDoseShapeIndx As Integer = 2
Private Const constDoseRouteIndx As Integer = 3
Private Const constDoseIndicationIndx As Integer = 4
Private Const constDoseDepartmentIndx As Integer = 5
Private Const constDoseGenderIndx As Integer = 6
Private Const constDoseMinAgeMoIndx As Integer = 7
Private Const constDoseMaxAgeMoIndx As Integer = 8
Private Const constDoseMinWeightKgIndx As Integer = 9
Private Const constDoseMaxWeightKgIndx As Integer = 10
Private Const constDoseMinGestDaysIndx As Integer = 11
Private Const constDoseMaxGestDaysIndx As Integer = 12
Private Const constDoseFrequenciesIndx As Integer = 13
Private Const constDoseUnitIndx As Integer = 14
Private Const constDoseNormDoseIndx As Integer = 15
Private Const constDoseMinDoseIndx As Integer = 16
Private Const constDoseMaxDoseIndx As Integer = 17
Private Const constDoseAbsMaxDoseIndx As Integer = 18
Private Const constDoseMaxPerDoseIndx As Integer = 19
Private Const constDoseIsDosePerKgIndx As Integer = 20
Private Const constDoseIsDosePerM2Indx As Integer = 21

Private Const constSolDepartmentIndx As Integer = 1
Private Const constSolGenericIndx As Integer = 2
Private Const constSolShapeIndx As Integer = 3
Private Const constSolMinAgeIndx As Integer = 4
Private Const constSolMaxAgeIndx As Integer = 5
Private Const constSolMinGestDaysIndx As Integer = 6
Private Const constSolMaxGestDaysIndx As Integer = 7
Private Const constSolMinWeightIndx As Integer = 8
Private Const constSolMaxWeightIndx As Integer = 9
Private Const constSolSolutionIndx As Integer = 10
Private Const constSolMinConcIndx As Integer = 11
Private Const constSolMaxConcIndx As Integer = 12
Private Const constSolSolutionVolumeIndx As Integer = 13
Private Const constSolMinInfusionTimeIndx As Integer = 14

Public Sub Formularium_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    If Not m_Formularium Is Nothing Then Exit Sub

    strTitle = "Formularium wordt geladen, een ogenblik geduld a.u.b. ..."
    
    ModProgress.StartProgress strTitle
       
    Set m_Formularium = New ClassFormularium
    m_Formularium.GetMedicationCollection True
    
    ModProgress.FinishProgress

End Sub

Private Sub FormConfig_Initialize()

    Dim intN As Integer
    Dim intC As Integer
    Dim strTitle As String

    If Not m_FormConfig Is Nothing Then Exit Sub

    strTitle = "Formularium configuratie wordt geladen, een ogenblik geduld a.u.b. ..."
    
    ModProgress.StartProgress strTitle
       
    Set m_FormConfig = New ClassFormConfig
    m_FormConfig.GetMedicamenten True
      
    ModProgress.FinishProgress

End Sub

Public Function Formularium_IsInitialized() As Boolean

    Formularium_IsInitialized = Not m_Formularium Is Nothing

End Function

Public Function Formularium_GetNewFormularium() As ClassFormularium

    Set m_Formularium = Nothing
    Formularium_Initialize
    Set Formularium_GetNewFormularium = m_Formularium

End Function

Public Function Formularium_GetFormularium() As ClassFormularium

    Formularium_Initialize
    Set Formularium_GetFormularium = m_Formularium

End Function

Public Function Formularium_GetFormConfig() As ClassFormConfig

    FormConfig_Initialize
    Set Formularium_GetFormConfig = m_FormConfig

End Function

Private Sub Test_Formularium_Initialize()

    Formularium_Initialize

End Sub

Public Sub Formularium_Import(objFormularium As ClassFormularium, strFileName As String, ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim intD As Integer
    Dim intK As Integer
    
    Dim objFormRange As Range
    Dim arrDose As Variant
    Dim objMedSheet As Worksheet
    Dim objDoseSheet As Worksheet
    Dim objDoseRange As Range
    Dim objWbk As Workbook
    Dim varCell As Variant
    Dim intCell As Integer
    
    Dim strSheet As String
    Dim blnIsPed As Boolean
    
    Dim objMed As ClassMedDisc
    Dim objDose As ClassDose
    Dim objPICUSolution As ClassSolution
    Dim objNICUSolution As ClassSolution
    
    On Error GoTo ErrorHandler
    
    blnIsPed = MetaVision_IsPICU()
    
    strSheet = "Medicatie"
    
    If blnShowProgress Then ModProgress.StartProgress "Formularium importeren"
    
    Application.DisplayAlerts = False
    ImprovePerf True
    
    Set objWbk = Workbooks.Open(strFileName, True, True)
    
    Set objMedSheet = objWbk.Worksheets(strSheet)
    Set objDoseSheet = objWbk.Worksheets("Doseringen")
    
    Set objFormRange = objMedSheet.Range("A1").CurrentRegion
    Set objDoseRange = objDoseSheet.Range("A1").CurrentRegion
    shtGlobTemp.Unprotect ModConst.CONST_PASSWORD
    
        
    intC = objFormRange.Rows.Count
    For intN = constFormDbStart To intC
        Set objMed = New ClassMedDisc
        Set objPICUSolution = New ClassSolution
        Set objNICUSolution = New ClassSolution
        
        With objMed
            
            .GPK = objFormRange.Cells(intN, constGPKIndx).Value2
            .MainGroup = objFormRange.Cells(intN, constMainGroupIndx).Value2
            .SubGroup = objFormRange.Cells(intN, constSubGroupIndx).Value2
            
            .ATC = objFormRange.Cells(intN, constATCIndx).Value2
            .Generic = objFormRange.Cells(intN, constGenericIndx).Value2
            .Product = objFormRange.Cells(intN, constProductIndx).Value2
            .Shape = objFormRange.Cells(intN, constShapeIndx).Value2
            .GenericQuantity = objFormRange.Cells(intN, constGenericQuantityIndx).Value2
            .GenericUnit = objFormRange.Cells(intN, constGenericQuantityUnitIndx).Value2
            .Label = objFormRange.Cells(intN, constLabelIndx).Value2
            .MultipleQuantity = objFormRange.Cells(intN, constMultipleQuantityIndx).Value2
            .MultipleUnit = objFormRange.Cells(intN, constMultipleQuantityUnitIndx).Value2
            
            .SetRouteList objFormRange.Cells(intN, constRouteIndx).Value2
            .SetIndicationList objFormRange.Cells(intN, constIndicationsIndx).Value2
            
            .MaxConc = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_MaxSolutionConcIndx, constNICU_MaxSolutionConcIndx)).Value2
            .Solution = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_SolutionIndx, constNICU_SolutionIndx)).Value2
            .SolutionVolume = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_SolutionVolumeIndx, constNICU_SolutionVolumeIndx)).Value2
            .MinInfusionTime = objFormRange.Cells(intN, IIf(blnIsPed, constPICU_MinInfusionTimeIndx, constNICU_MinInfusionTimeIndx)).Value2
        
            With objPICUSolution
                .Generic = objFormRange.Cells(intN, constGenericIndx).Value2
                .Shape = objFormRange.Cells(intN, constShapeIndx).Value2
                .MaxConc = objFormRange.Cells(intN, constPICU_MaxSolutionConcIndx).Value2
                .Solution = objFormRange.Cells(intN, constPICU_SolutionIndx).Value2
                .SolutionVolume = objFormRange.Cells(intN, constPICU_SolutionVolumeIndx).Value2
                .MinInfusionTime = objFormRange.Cells(intN, constPICU_MinInfusionTimeIndx).Value2
            End With
                    
            With objNICUSolution
                .Generic = objFormRange.Cells(intN, constGenericIndx).Value2
                .Shape = objFormRange.Cells(intN, constShapeIndx).Value2
                .MaxConc = objFormRange.Cells(intN, constNICU_MaxSolutionConcIndx).Value2
                .Solution = objFormRange.Cells(intN, constNICU_SolutionIndx).Value2
                .SolutionVolume = objFormRange.Cells(intN, constNICU_SolutionVolumeIndx).Value2
                .MinInfusionTime = objFormRange.Cells(intN, constNICU_MinInfusionTimeIndx).Value2
            End With
            
            objMed.SetPICUSolution objPICUSolution
            objMed.SetNICUSolution objNICUSolution
            
            With objDoseRange
                .AutoFilter 1, objMed.Generic
                .AutoFilter 2, objMed.Shape
            End With
            
            objDoseRange.SpecialCells(xlCellTypeVisible).Copy
            shtGlobTemp.Range("A1").PasteSpecial xlPasteValues
            arrDose = shtGlobTemp.Range("A1").CurrentRegion.Value
            shtGlobTemp.Range("A1").CurrentRegion.Clear
            
            intK = UBound(arrDose, 1)
            For intD = 2 To intK
                Set objDose = New ClassDose
                
                With objDose
                    .Department = arrDose(intD, constDoseDepartmentIndx)
                    .Generic = arrDose(intD, constDoseGenericIndx)
                    .Shape = arrDose(intD, constDoseShapeIndx)
                    .Route = arrDose(intD, constDoseRouteIndx)
                    .Indication = arrDose(intD, constDoseIndicationIndx)
                    .Gender = arrDose(intD, constDoseGenderIndx)
                    .MinAgeMo = arrDose(intD, constDoseMinAgeMoIndx)
                    .MaxAgeMo = arrDose(intD, constDoseMaxAgeMoIndx)
                    .MinWeightKg = arrDose(intD, constDoseMinWeightKgIndx)
                    .MaxWeightKg = arrDose(intD, constDoseMaxWeightKgIndx)
                    .MinGestDays = arrDose(intD, constDoseMinGestDaysIndx)
                    .MaxGestDays = arrDose(intD, constDoseMaxGestDaysIndx)
                    .Frequencies = arrDose(intD, constDoseFrequenciesIndx)
                    .Unit = objMed.MultipleUnit
                    .NormDose = arrDose(intD, constDoseNormDoseIndx)
                    .MinDose = arrDose(intD, constDoseMinDoseIndx)
                    .MaxDose = arrDose(intD, constDoseMaxDoseIndx)
                    .AbsMaxDose = arrDose(intD, constDoseAbsMaxDoseIndx)
                    .MaxPerDose = arrDose(intD, constDoseMaxPerDoseIndx)
                    .IsDosePerKg = Not (arrDose(intD, constDoseIsDosePerKgIndx) = vbNullString)
                    .IsDosePerM2 = Not (arrDose(intD, constDoseIsDosePerM2Indx) = vbNullString)
                End With
                
                objMed.AddDose objDose
            Next
            
        End With
                
        objFormularium.AddMedication objMed
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", intC, intN
        
    Next intN
    
    If blnShowProgress Then ModProgress.FinishProgress
    
    objWbk.Close
    shtGlobTemp.Protect ModConst.CONST_PASSWORD
    Application.DisplayAlerts = True
    ImprovePerf False
    
    ModMessage.ShowMsgBoxInfo strFileName & " is succesvol geimporteerd."
    
    Exit Sub
    
ErrorHandler:
    
    ModLog.LogError Err, "Could not import formularium from: " & strFileName
    
    On Error Resume Next
    
    objWbk.Close
    shtGlobTemp.Range("A1").CurrentRegion.Clear
    shtGlobTemp.Protect ModConst.CONST_PASSWORD

    Application.DisplayAlerts = True
    ImprovePerf False
    
    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxExclam "Could not retrieve medicament from: " & strFileName
    
End Sub

Public Function Formularium_GetDoses(objMedCol As Collection, blnAll As Boolean) As Collection

    Dim objDose As ClassDose
    Dim objItem As ClassDose
    Dim objDoseCol As Collection
    Dim objMed As ClassMedDisc
    Dim blnContains As Boolean
    
    Set objDoseCol = New Collection
    
    For Each objMed In objMedCol
        For Each objDose In objMed.Doses
        
            If (objDose.AbsMaxDose > 0 Or objDose.MaxDose > 0 Or objDose.MinDose > 0) Or blnAll Then
                blnContains = False
                For Each objItem In objDoseCol
                    blnContains = objItem.Generic = objDose.Generic
                    blnContains = blnContains And (objItem.Shape = objDose.Shape)
                    blnContains = blnContains And (objItem.Department = objDose.Department)
                    blnContains = blnContains And (objItem.Indication = objDose.Indication)
                    blnContains = blnContains And (objItem.Route = objDose.Route)
                    blnContains = blnContains And (objItem.Gender = objDose.Gender)
                    blnContains = blnContains And (objItem.MinAgeMo = objDose.MinAgeMo)
                    blnContains = blnContains And (objItem.MaxAgeMo = objDose.MaxAgeMo)
                    blnContains = blnContains And (objItem.MinWeightKg = objDose.MinWeightKg)
                    blnContains = blnContains And (objItem.MaxWeightKg = objDose.MaxWeightKg)
                    blnContains = blnContains And (objItem.MinGestDays = objDose.MinGestDays)
                    blnContains = blnContains And (objItem.MaxGestDays = objDose.MaxGestDays)
                    
                    If blnContains Then Exit For
                Next
                
                If Not blnContains Then
                    objDoseCol.Add objDose
                End If
            End If
        Next
    Next
    
    Set Formularium_GetDoses = objDoseCol

End Function

Public Function Formularium_GetSolutions(ByVal blnIsPed As Boolean, objMedCol As Collection) As Collection

    Dim objSol As ClassSolution
    Dim objMedSol As ClassSolution
    Dim objSolCol As Collection
    Dim objMed As ClassMedDisc
    Dim blnContains As Boolean
    
    Set objSolCol = New Collection
    
    For Each objMed In objMedCol
        Set objMedSol = IIf(blnIsPed, objMed.PICUSolution, objMed.NICUSolution)
        
        If Not objMedSol.Generic = vbNullString Then
            blnContains = False
            For Each objSol In objSolCol
                If objSol.Generic = objMedSol.Generic And objSol.Shape = objMedSol.Shape Then
                    blnContains = True
                    Exit For
                End If
            Next
            If Not blnContains And (objMedSol.MaxConc > 0 Or objMedSol.SolutionVolume > 0) Then
                objSolCol.Add objMedSol
            End If
        End If
    Next
    
    Set Formularium_GetSolutions = objSolCol

End Function

Private Sub Test_Formularium_GetSolutions()

    Dim objMedCol As Collection
    Dim objSolCol As Collection
    
    ModAdmin.Admin_MedDiscImport
    Set objMedCol = Formularium_GetFormularium.GetMedicationCollection(True)
    Set objSolCol = Formularium_GetSolutions(True, objMedCol)
    
    ModMessage.ShowMsgBoxInfo "Found " & objSolCol.Count & " solutions"

End Sub


Public Sub Formularium_GetMedDiscConfig(objFormularium As ClassFormConfig, ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim objFormRange As Range
    Dim objSheet As Worksheet
    
    Dim strFileName As String
    Dim strName As String
    Dim strSheet As String
    Dim blnIsPed As Boolean
    
    Dim objMed As ClassMedDiscConfig
    
    On Error GoTo GetMedicamentenError
    
    blnIsPed = MetaVision_IsPICU()
    
    strName = constFormularium
    strSheet = "Table"
    
    Application.DisplayAlerts = False
    ImprovePerf True
    
    strFileName = ModMedDisc.GetFormulariumDatabasePath() + strName

    Workbooks.Open strFileName, True, True
    
    Set objSheet = Workbooks(strName).Worksheets(strSheet)
    Set objFormRange = objSheet.Range("A1").CurrentRegion
        
    intC = objFormRange.Rows.Count
    For intN = constFormDbStart To intC
        Set objMed = New ClassMedDiscConfig
        
        With objMed
            .GPK = objFormRange.Cells(intN, constGPKIndx).Value2
            .MainGroup = objFormRange.Cells(intN, constMainGroupIndx).Value2
            .SubGroup = objFormRange.Cells(intN, constSubGroupIndx).Value2
            
            .ATC = objFormRange.Cells(intN, constATCIndx).Value2
            .Generic = objFormRange.Cells(intN, constGenericIndx).Value2
            .Product = objFormRange.Cells(intN, constProductIndx).Value2
            .Shape = objFormRange.Cells(intN, constShapeIndx).Value2
            .GenericQuantity = objFormRange.Cells(intN, constGenericQuantityIndx).Value2
            .GenericUnit = objFormRange.Cells(intN, constGenericQuantityUnitIndx).Value2
            .Label = objFormRange.Cells(intN, constLabelIndx).Value2
            .MultipleQuantity = objFormRange.Cells(intN, constMultipleQuantityIndx).Value2
            .MultipleUnit = objFormRange.Cells(intN, constMultipleQuantityUnitIndx).Value2
            
            .SetRouteList objFormRange.Cells(intN, constRouteIndx).Value2
            .SetIndicationList objFormRange.Cells(intN, constIndicationsIndx).Value2
            
            .PedMaxConc = objFormRange.Cells(intN, constPICU_MaxSolutionConcIndx).Value2
            .PedSolutionVolume = objFormRange.Cells(intN, constPICU_SolutionVolumeIndx).Value2
            .PedSolution = objFormRange.Cells(intN, constPICU_SolutionIndx).Value2
            .PedMinInfusionTime = objFormRange.Cells(intN, constPICU_MinInfusionTimeIndx).Value2
            
            .NeoMaxConc = objFormRange.Cells(intN, constNICU_MaxSolutionConcIndx).Value2
            .NeoSoutionVolume = objFormRange.Cells(intN, constNICU_SolutionVolumeIndx).Value2
            .NeoSolution = objFormRange.Cells(intN, constNICU_SolutionIndx).Value2
            .NeoMinInfustionTime = objFormRange.Cells(intN, constNICU_MinInfusionTimeIndx).Value2
            
        End With
        
        objFormularium.AddMedicament objMed
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", intC, intN
        
    Next intN
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    ImprovePerf False
    
    Exit Sub
    
GetMedicamentenError:
    
    ModLog.LogError Err, "Could not retrieve medicament from: " & strFileName
    
    On Error Resume Next
    
    Workbooks(strName).Close

    Application.DisplayAlerts = True
    ImprovePerf False
    
    ModProgress.FinishProgress
    
End Sub

' ToDo add headings
Public Sub Formularium_Export(ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim intD As Integer
    Dim intS As Integer
    Dim intF As Integer
    Dim intK As Integer
    Dim objWbk As Workbook
    Dim objMedSheet As Worksheet
    Dim objDoseSheet As Worksheet
    Dim objSolSheet As Worksheet
    
    Dim strFreqForm As String
    Dim strConcat As String
    
    Dim strFile As String
    Dim varDir As String
    Dim strName As String
    Dim strSheet As String
    Dim blnIsPed As Boolean
    
    Dim colMed As Collection
    Dim objMed As ClassMedDisc
    Dim colDose As Collection
    Dim objDose As ClassDose
    Dim colSolution As Collection
    Dim objSolution As ClassSolution
    
    On Error GoTo ErrorHandler
    
    strName = "FormulariumDB"
    strSheet = "Medicatie"
    
    varDir = ModFile.GetFolderWithDialog(WbkAfspraken.Path)
    
    If CStr(varDir) = vbNullString Then Exit Sub
        
    Application.DisplayAlerts = False
    ImprovePerf True
        
    strFile = Replace(ModDate.FormatDateTimeSeconds(Now()), ":", "_")
    strFile = Replace(strFile, " ", "_")
    strFile = CStr(varDir) & "\" & strName & "_" & strFile & "_.xlsx"
    
    Set objWbk = Workbooks.Add()
    Set objMedSheet = objWbk.Sheets(1)
    objMedSheet.Name = strSheet
    
    Set objSolSheet = objWbk.Sheets(2)
    objSolSheet.Name = "Oplossingen"
    
    Set objDoseSheet = objWbk.Sheets(3)
    objDoseSheet.Name = "Doseringen"
    
    Set colMed = Formularium_GetFormularium().GetMedicationCollection(False)
    Set colDose = Formularium_GetDoses(colMed, True)
    
    ModProgress.FinishProgress
    ModProgress.StartProgress "Discontinue Medicatie Exporteren"
    
    With objMedSheet
        .Cells(1, constGPKIndx).Value2 = "GPK"
        .Cells(1, constATCIndx).Value2 = "ATC"
        .Cells(1, constMainGroupIndx).Value2 = "Hoofd Groep"
        .Cells(1, constSubGroupIndx).Value2 = "Sub Groep"
        .Cells(1, constGenericIndx).Value2 = "Generiek"
        .Cells(1, constProductIndx).Value2 = "Product"
        .Cells(1, constLabelIndx).Value2 = "Etiket"
        .Cells(1, constShapeIndx).Value2 = "Vorm"
        .Cells(1, constRouteIndx).Value2 = "Routes"
        .Cells(1, constGenericQuantityIndx).Value2 = "Generiek Hoeveelheid"
        .Cells(1, constGenericQuantityUnitIndx).Value2 = "Hoeveelheid Eenheid"
                    
        .Cells(1, constMultipleQuantityIndx).Value2 = "Veelvoud"
        .Cells(1, constMultipleQuantityUnitIndx).Value2 = "Veelvoud Eenheid"
        
        .Cells(1, constIndicationsIndx).Value2 = "Indicaties"
                        
    End With

    With objDoseSheet
        .Cells(1, constDoseGenericIndx).Value2 = "Generic"
        .Cells(1, constDoseShapeIndx).Value2 = "Vorm"
        .Cells(1, constDoseRouteIndx).Value2 = "Route"
        .Cells(1, constDoseIndicationIndx).Value2 = "Indicatie"
        .Cells(1, constDoseDepartmentIndx).Value2 = "Afdeling"
        .Cells(1, constDoseGenderIndx).Value2 = "Geslacht"
        .Cells(1, constDoseMinAgeMoIndx).Value2 = "MinLeeftijdMnd"
        .Cells(1, constDoseMaxAgeMoIndx).Value2 = "MaxLeeftijdMnd"
        .Cells(1, constDoseMinWeightKgIndx).Value2 = "MinGewichtKg"
        .Cells(1, constDoseMaxWeightKgIndx).Value2 = "MaxGewichtKg"
        .Cells(1, constDoseMinGestDaysIndx).Value2 = "MinGestDagen"
        .Cells(1, constDoseMaxGestDaysIndx).Value2 = "MaxGestDagen"
        .Cells(1, constDoseFrequenciesIndx).Value2 = "Frequenties"
        .Cells(1, constDoseUnitIndx).Value2 = "Eenheid"
        .Cells(1, constDoseNormDoseIndx).Value2 = "NormDose"
        .Cells(1, constDoseMinDoseIndx).Value2 = "MinDose"
        .Cells(1, constDoseMaxDoseIndx).Value2 = "MaxDose"
        .Cells(1, constDoseAbsMaxDoseIndx).Value2 = "AbsMaxDose"
        .Cells(1, constDoseMaxPerDoseIndx).Value2 = "MaxPerDose"
        .Cells(1, constDoseIsDosePerKgIndx).Value2 = "DosePerKg"
        .Cells(1, constDoseIsDosePerM2Indx).Value2 = "DosePerM2"
        
        intF = constDoseIsDosePerM2Indx + 2
        .Cells(1, intF).Value2 = constDoseFreq_AN
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_2D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_3D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_4D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_5D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_6D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_7D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_8D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_9D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_10D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_11D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_12D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_24D
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1D2
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1D3
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1U36
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1W
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_2W
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_3W
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_4W
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1W2
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1W4
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1W12
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1W13
        intF = intF + 1
        .Cells(1, intF).Value2 = constDoseFreq_1M
        intF = intF + 1
        
    End With
    
    With objSolSheet
        .Cells(1, constSolDepartmentIndx).Value2 = "Afdeling"
        .Cells(1, constSolGenericIndx).Value2 = "Generiek"
        .Cells(1, constSolShapeIndx).Value2 = "Vorm"
        .Cells(1, constSolMinAgeIndx).Value2 = "MinLeeftijdMnd"
        .Cells(1, constSolMaxAgeIndx).Value2 = "MaxLeeftijdMnd"
        .Cells(1, constSolMinGestDaysIndx).Value2 = "MinGestDagen"
        .Cells(1, constSolMaxGestDaysIndx).Value2 = "MaxGestDagen"
        .Cells(1, constSolMinWeightIndx).Value2 = "MinGewichtKg"
        .Cells(1, constSolMaxWeightIndx).Value2 = "MaxGewichtKg"
        .Cells(1, constSolSolutionIndx).Value2 = "Oplossing"
        .Cells(1, constSolMinConcIndx).Value2 = "MinConcentratie"
        .Cells(1, constSolMaxConcIndx).Value2 = "MaxConcentratie"
        .Cells(1, constSolSolutionVolumeIndx).Value2 = "OplossingVolume"
    End With

    
    intN = constFormDbStart
    intD = 2
    intS = 2
    
    intC = colDose.Count
    For Each objDose In colDose
    
        With objDose
            objDoseSheet.Cells(intD, constDoseDepartmentIndx).Value2 = .Department
            objDoseSheet.Cells(intD, constDoseGenericIndx).Value2 = .Generic
            objDoseSheet.Cells(intD, constDoseShapeIndx).Value2 = .Shape
            objDoseSheet.Cells(intD, constDoseRouteIndx).Value2 = .Route
            objDoseSheet.Cells(intD, constDoseIndicationIndx).Value2 = .Indication
            objDoseSheet.Cells(intD, constDoseGenderIndx).Value2 = .Gender
            If .MinAgeMo > 0 Then objDoseSheet.Cells(intD, constDoseMinAgeMoIndx).Value2 = .MinAgeMo
            If .MaxAgeMo > 0 Then objDoseSheet.Cells(intD, constDoseMaxAgeMoIndx).Value2 = .MaxAgeMo
            If .MinWeightKg > 0 Then objDoseSheet.Cells(intD, constDoseMinWeightKgIndx).Value2 = .MinWeightKg
            If .MaxWeightKg > 0 Then objDoseSheet.Cells(intD, constDoseMaxWeightKgIndx).Value2 = .MaxWeightKg
            If .MinGestDays > 0 Then objDoseSheet.Cells(intD, constDoseMinGestDaysIndx).Value2 = .MinGestDays
            If .MaxGestDays > 0 Then objDoseSheet.Cells(intD, constDoseMaxGestDaysIndx).Value2 = .MaxGestDays
            objDoseSheet.Cells(intD, constDoseFrequenciesIndx).Value2 = .Frequencies
            objDoseSheet.Cells(intD, constDoseUnitIndx).Value2 = .Unit
            
            If .NormDose > 0 Then
                objDoseSheet.Cells(intD, constDoseNormDoseIndx).Value2 = .NormDose
            End If
            If .MinDose > 0 Then
                objDoseSheet.Cells(intD, constDoseMinDoseIndx).Value2 = .MinDose
            End If
            If .MaxDose > 0 Then
                objDoseSheet.Cells(intD, constDoseMaxDoseIndx).Value2 = .MaxDose
            End If
            If .AbsMaxDose > 0 Then
                objDoseSheet.Cells(intD, constDoseAbsMaxDoseIndx).Value2 = .AbsMaxDose
            End If
            
            If .MaxPerDose > 0 Then objDoseSheet.Cells(intD, constDoseMaxPerDoseIndx).Value2 = .MaxPerDose
        
            If Not objDose.Frequencies = vbNullString Then
                objDoseSheet.Cells(intD, constDoseFrequenciesIndx).Value2 = .Frequencies
            Else
                intK = intF
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "W")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "X")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "Y")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "Z")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AA")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AB")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AC")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AD")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AE")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AF")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AG")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AH")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AI")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AJ")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AK")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AL")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AM")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AN")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AO")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AP")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AQ")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AR")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AS")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AT")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AU")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1
                strFreqForm = Replace(Replace(constDoseFreqFormula, "{2}", intD), "{W}", "AV")
                objDoseSheet.Cells(intD, intK).Formula = strFreqForm
                intK = intK + 1

                strConcat = "=IF(AW" & intD & "<>"""",AW" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(AX" & intD & "<>"""",AX" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(AY" & intD & "<>"""",AY" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(AZ" & intD & "<>"""",AZ" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BA" & intD & "<>"""",BA" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BB" & intD & "<>"""",BB" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BC" & intD & "<>"""",BC" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BD" & intD & "<>"""",BD" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BE" & intD & "<>"""",BE" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BF" & intD & "<>"""",BF" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BG" & intD & "<>"""",BG" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BH" & intD & "<>"""",BH" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BI" & intD & "<>"""",BI" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BJ" & intD & "<>"""",BJ" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BK" & intD & "<>"""",BK" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BL" & intD & "<>"""",BL" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BM" & intD & "<>"""",BM" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BN" & intD & "<>"""",BN" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BO" & intD & "<>"""",BO" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BP" & intD & "<>"""",BP" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BQ" & intD & "<>"""",BQ" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BR" & intD & "<>"""",BR" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BS" & intD & "<>"""",BS" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BT" & intD & "<>"""",BT" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BU" & intD & "<>"""",BU" & intD & "&""||"","""")"
                strConcat = strConcat & "&IF(BV" & intD & "<>"""",BV" & intD & "&""||"","""")"
                objDoseSheet.Cells(intD, intK).Formula = strConcat
                
                objDoseSheet.Cells(intD, constDoseFrequenciesIndx).Formula = "=IFERROR(IF(LEN(BW" & intD & ">0),MID(BW" & intD & ",1,LEN(BW" & intD & ")-2),""""),"""")"
            End If
        
        End With
        
        
        intD = intD + 1
        If blnShowProgress Then ModProgress.SetJobPercentage "Export " & intD, intC, intD
    Next
            
    intC = colMed.Count
    For Each objMed In colMed
    
        With objMed
            
            objMedSheet.Cells(intN, constGPKIndx).Value2 = objMed.GPK
            objMedSheet.Cells(intN, constATCIndx).Value2 = objMed.ATC
            objMedSheet.Cells(intN, constMainGroupIndx).Value2 = objMed.MainGroup
            objMedSheet.Cells(intN, constSubGroupIndx).Value2 = objMed.SubGroup
            objMedSheet.Cells(intN, constGenericIndx).Value2 = objMed.Generic
            objMedSheet.Cells(intN, constProductIndx).Value2 = objMed.Product
            objMedSheet.Cells(intN, constLabelIndx).Value2 = objMed.Label
            objMedSheet.Cells(intN, constShapeIndx).Value2 = objMed.Shape
            objMedSheet.Cells(intN, constRouteIndx).Value2 = objMed.Routes
            objMedSheet.Cells(intN, constGenericQuantityIndx).Value2 = objMed.GenericQuantity
            objMedSheet.Cells(intN, constGenericQuantityUnitIndx).Value2 = objMed.GenericUnit
                        
            If .MultipleQuantity > 0 Then objMedSheet.Cells(intN, constMultipleQuantityIndx).Value2 = .MultipleQuantity
            If Not .MultipleUnit = vbNullString Then objMedSheet.Cells(intN, constMultipleQuantityUnitIndx).Value2 = .MultipleUnit
            
            objMedSheet.Cells(intN, constIndicationsIndx).Value2 = objMed.Indications
                        
            With objMed.PICUSolution
                objSolSheet.Cells(intS, constSolDepartmentIndx).Value2 = "PICU"
                objSolSheet.Cells(intS, constSolGenericIndx).Value2 = objMed.Generic
                objSolSheet.Cells(intS, constSolShapeIndx).Value2 = objMed.Shape
                If .MaxConc > 0 Then objSolSheet.Cells(intS, constSolMaxConcIndx).Value2 = .MaxConc
                If .SolutionVolume > 0 Then objSolSheet.Cells(intS, constSolSolutionVolumeIndx).Value2 = .SolutionVolume
                If Not .Solution = vbNullString Then objSolSheet.Cells(intS, constSolSolutionIndx).Value2 = .Solution
                If .MinInfusionTime > 0 Then objSolSheet.Cells(intS, constSolMinInfusionTimeIndx).Value2 = .MinInfusionTime
                
                intS = intS + 1
            End With
            
            With objMed.NICUSolution
                objSolSheet.Cells(intS, constSolDepartmentIndx).Value2 = "NICU"
                objSolSheet.Cells(intS, constSolGenericIndx).Value2 = objMed.Generic
                objSolSheet.Cells(intS, constSolShapeIndx).Value2 = objMed.Shape
                If .MaxConc > 0 Then objSolSheet.Cells(intS, constSolMaxConcIndx).Value2 = .MaxConc
                If .SolutionVolume > 0 Then objSolSheet.Cells(intS, constSolSolutionVolumeIndx).Value2 = .SolutionVolume
                If Not .Solution = vbNullString Then objSolSheet.Cells(intS, constSolSolutionIndx).Value2 = .Solution
                If .MinInfusionTime > 0 Then objSolSheet.Cells(intS, constSolMinInfusionTimeIndx).Value2 = .MinInfusionTime
                
                intS = intS + 1
            End With
            
        End With
        
        intN = intN + 1
        If blnShowProgress Then ModProgress.SetJobPercentage "Export " & intN, intC, intN
    Next
    
    
    
    objWbk.SaveAs strFile
    objWbk.Close
    
    Set m_Formularium = Nothing
    
    ModProgress.FinishProgress

    Application.DisplayAlerts = True
    ImprovePerf False
    
        
    ModMessage.ShowMsgBoxInfo "Het formularium is succesvol geexporteerd naar: " & strFile
    
    Exit Sub
    
ErrorHandler:
    
    FinishProgress
    ModLog.LogError Err, "Could not export medication to: " & strFile
    
    On Error Resume Next
    
    objWbk.Close

    Application.DisplayAlerts = True
    ImprovePerf False
    
    ModProgress.FinishProgress
    
End Sub

Private Sub Test_Formularium_ExportMedDiscConfig()

    Formularium_Export True
    
End Sub

Public Sub Formularium_ShowConfig()

    FormAdminMedDisc.Show
    
End Sub

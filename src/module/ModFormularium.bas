Attribute VB_Name = "ModFormularium"
Option Explicit

Private m_Formularium As ClassFormularium

Private Const constFormularium As String = "FormulariumDb.xlsm"

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
Private Const constHasSolutionsIndx As Integer = 15
Private Const constIsActiveIndx As Integer = 16

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
Private Const constSolMinGenericQuantityIndx As Integer = 4
Private Const constSolMaxGenericQuantityIndx As Integer = 5
Private Const constSolSolutionsIndx As Integer = 6
Private Const constSolMinConcIndx As Integer = 7
Private Const constSolMaxConcIndx As Integer = 8
Private Const constSolSolutionVolumeIndx As Integer = 9
Private Const constSolMinInfusionTimeIndx As Integer = 10

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


Public Function Formularium_IsInitialized() As Boolean

    Formularium_IsInitialized = Not m_Formularium Is Nothing

End Function

Public Function Formularium_GetNewFormularium() As ClassFormularium

    Set m_Formularium = Nothing
    Formularium_Initialize
    Set Formularium_GetNewFormularium = m_Formularium

End Function

Private Sub Test_Formularium_Initialize()

    Formularium_GetNewFormularium
    ModMessage.ShowMsgBoxInfo "Formularium heeft " & m_Formularium.MedicamentCount & " medicamenten"

End Sub

Public Function Formularium_GetFormularium() As ClassFormularium

    Formularium_Initialize
    Set Formularium_GetFormularium = m_Formularium

End Function

Public Sub Formularium_Import(objFormularium As ClassFormularium, strFileName As String, ByVal blnShowProgress As Boolean)

    Dim intN As Integer
    Dim intC As Integer
    Dim intD As Integer
    Dim intS As Integer
    Dim intK As Integer
    
    Dim objFormRange As Range
    Dim arrDose As Variant
    Dim arrSol As Variant
    Dim objMedSheet As Worksheet
    Dim objDoseSheet As Worksheet
    Dim objDoseRange As Range
    Dim objSolSheet As Worksheet
    Dim objSolRange As Range
    Dim objWbk As Workbook
    Dim varCell As Variant
    Dim intCell As Integer
    
    Dim strSheet As String
    Dim blnIsPed As Boolean
    
    Dim objMed As ClassMedDisc
    Dim objDose As ClassDose
    Dim objSol As ClassSolution
    
    On Error GoTo ErrorHandler
    
    blnIsPed = MetaVision_IsPICU()
    
    strSheet = "Medicatie"
    
    If blnShowProgress Then ModProgress.StartProgress "Formularium importeren"
    
    Application.DisplayAlerts = False
    ImprovePerf True
    
    Set objWbk = Workbooks.Open(strFileName, True, True)
    objWbk.Windows(1).Visible = False
    
    Set objMedSheet = objWbk.Worksheets(strSheet)
    objMedSheet.AutoFilterMode = False
    
    Set objDoseSheet = objWbk.Worksheets("Doseringen")
    objDoseSheet.AutoFilterMode = False
    
    Set objSolSheet = objWbk.Worksheets("Oplossingen")
    objSolSheet.AutoFilterMode = False
    
    Set objFormRange = objMedSheet.Range("A1").CurrentRegion
    Set objDoseRange = objDoseSheet.Range("A1").CurrentRegion
    Set objSolRange = objSolSheet.Range("A1").CurrentRegion
        
    If objFormRange.Rows.Count > 1 Then
        With objMedSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=objMedSheet.Range("E2"), Order:=xlAscending
            .SortFields.Add Key:=objMedSheet.Range("H2"), Order:=xlAscending
            .SortFields.Add Key:=objMedSheet.Range("J2"), Order:=xlAscending
            .SetRange objMedSheet.Range("A2:P" & objFormRange.Rows.Count)
            .Apply
        End With
    End If
    
    If objDoseRange.Rows.Count > 1 Then
        With objDoseSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=objDoseSheet.Range("A2"), Order:=xlAscending
            .SortFields.Add Key:=objDoseSheet.Range("D2"), Order:=xlAscending
            .SortFields.Add Key:=objDoseSheet.Range("C2"), Order:=xlAscending
            .SetRange objDoseSheet.Range("A2:BW" & objDoseRange.Rows.Count)
            .Apply
        End With
    End If
    
    If objSolRange.Rows.Count > 1 Then
        With objSolSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=objSolSheet.Range("B2"), Order:=xlAscending
            .SortFields.Add Key:=objSolSheet.Range("C2"), Order:=xlAscending
            .SortFields.Add Key:=objSolSheet.Range("A2"), Order:=xlAscending
            .SetRange objSolSheet.Range("A2:M" & objSolRange.Rows.Count)
            .Apply
        End With
    End If
    
    shtGlobTemp.Unprotect ModConst.CONST_PASSWORD
            
    intC = objFormRange.Rows.Count
    For intN = 2 To intC
        Set objMed = New ClassMedDisc
        
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
            .HasSolutions = IIf(Trim(objFormRange.Cells(intN, constHasSolutionsIndx).Value2) = "x", True, False)
            .IsActive = IIf(Trim(objFormRange.Cells(intN, constIsActiveIndx).Value2) = "x", True, False)
            
            If objMed.HasSolutions Then
            
                With objSolRange
                    .AutoFilter 2, objMed.Generic
                    .AutoFilter 3, objMed.Shape
                End With
                
                objSolRange.SpecialCells(xlCellTypeVisible).Copy
                shtGlobTemp.Range("A1").PasteSpecial xlPasteValues
                arrSol = shtGlobTemp.Range("A1").CurrentRegion.Value
                shtGlobTemp.Range("A1").CurrentRegion.Clear
                
                intK = UBound(arrSol, 1)
                For intS = 2 To intK
                    Set objSol = New ClassSolution
                    
                    With objSol
                        objSol.Department = arrSol(intS, constSolDepartmentIndx)
                        objSol.Generic = objMed.Generic
                        objSol.Shape = objMed.Shape
                        objSol.MinGenericQuantity = arrSol(intS, constSolMinGenericQuantityIndx)
                        objSol.MaxGenericQuantity = arrSol(intS, constSolMaxGenericQuantityIndx)
                        objSol.Solutions = arrSol(intS, constSolSolutionsIndx)
                        objSol.SolutionVolume = arrSol(intS, constSolSolutionVolumeIndx)
                        objSol.MinConc = arrSol(intS, constSolMinConcIndx)
                        objSol.MaxConc = arrSol(intS, constSolMaxConcIndx)
                        objSol.MinInfusionTime = arrSol(intS, constSolMinInfusionTimeIndx)
                    End With
                    
                    objMed.AddSolution objSol
                Next
            
            End If
            
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
                    .IsDosePerKg = Trim(arrDose(intD, constDoseIsDosePerKgIndx)) = "x"
                    .IsDosePerM2 = Trim(arrDose(intD, constDoseIsDosePerM2Indx)) = "x"
                End With
                
                objMed.AddDose objDose
            Next
            
        End With
                
        objFormularium.AddMedication objMed
        
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", intC, intN
        
    Next intN
    
    If blnShowProgress Then ModProgress.FinishProgress
    
    objWbk.Windows(1).Visible = True
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
    shtGlobTemp.Unprotect CONST_PASSWORD
    shtGlobTemp.Range("A1").CurrentRegion.Clear
    shtGlobTemp.Protect CONST_PASSWORD
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
        
            If (objDose.NormDose > 0 Or objDose.MinDose > 0 Or objDose.MaxDose > 0 Or objDose.AbsMaxDose > 0 Or objDose.MaxPerDose > 0) Or blnAll Then
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

Public Function Formularium_GetSolutions(objMedCol As Collection, ByVal blnAll As Boolean) As Collection

    Dim objSol As ClassSolution
    Dim objItem As ClassSolution
    Dim objSolCol As Collection
    Dim objMed As ClassMedDisc
    Dim blnContains As Boolean
    
    Set objSolCol = New Collection
    
    For Each objMed In objMedCol
        If objMed.HasSolutions Then
            For Each objSol In objMed.Solutions
            
                If (objSol.Solutions = vbNullString Or objSol.SolutionVolume > 0 Or objSol.MinConc > 0 Or objSol.MaxConc > 0 Or objSol.MinInfusionTime > 0) Or blnAll Then
                    blnContains = False
                    For Each objItem In objSolCol
                        blnContains = objItem.Generic = objSol.Generic
                        blnContains = blnContains And (objItem.Department = objSol.Department)
                        blnContains = blnContains And (objItem.Shape = objSol.Shape)
                        blnContains = blnContains And (objItem.MinGenericQuantity = objSol.MinGenericQuantity)
                        blnContains = blnContains And (objItem.MaxGenericQuantity = objSol.MaxGenericQuantity)
                        blnContains = blnContains And (objItem.Solutions = objSol.Solutions)
                        blnContains = blnContains And (objItem.SolutionVolume = objSol.SolutionVolume)
                        blnContains = blnContains And (objItem.MinConc = objSol.MinConc)
                        blnContains = blnContains And (objItem.MaxConc = objSol.MaxConc)
                        blnContains = blnContains And (objItem.MinInfusionTime = objSol.MinInfusionTime)
    
                        If blnContains Then Exit For
                    Next
                    
                    If Not blnContains Then
                        objSolCol.Add objSol
                    End If
                End If
            Next
        End If
    Next
    Set Formularium_GetSolutions = objSolCol

End Function

Private Sub Test_Formularium_GetSolutions()

    Dim objMedCol As Collection
    Dim objSolCol As Collection
    
    ModAdmin.Admin_MedDiscImport
    Set objMedCol = Formularium_GetFormularium.GetMedicationCollection(True)
    Set objSolCol = Formularium_GetSolutions(objMedCol, True)
    
    ModMessage.ShowMsgBoxInfo "Found " & objSolCol.Count & " solutions"

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
    
    Dim objFormRange As Range
    Dim objDoseRange As Range
    Dim objSolRange As Range
    
    Dim strFreqForm As String
    Dim strConcat As String
    
    Dim strFile As String
    Dim varDir As String
    Dim strName As String
    Dim strSheet As String
    Dim blnIsPed As Boolean
    
    Dim varFreq As Variant
    Dim arrFreq As Variant
    
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
    objWbk.Windows(1).Visible = False
    
    Set objMedSheet = objWbk.Sheets(1)
    objMedSheet.Name = strSheet
    
    Set objSolSheet = objWbk.Sheets(2)
    objSolSheet.Name = "Oplossingen"
    
    Set objDoseSheet = objWbk.Sheets(3)
    objDoseSheet.Name = "Doseringen"
    
    Set colMed = Formularium_GetFormularium().GetMedicationCollection(True)
    Set colDose = Formularium_GetDoses(colMed, True)
    Set colSolution = Formularium_GetSolutions(colMed, True)
    
    ModProgress.FinishProgress
    ModProgress.StartProgress "Discontinue Medicatie Exporteren"
    
    With objMedSheet
        .Cells(1, constGPKIndx).Value2 = "GPK"
        .Cells(1, constATCIndx).Value2 = "ATC"
        .Cells(1, constMainGroupIndx).Value2 = "Hoofd Groep"
        .Cells(1, constSubGroupIndx).Value2 = "SubGroep"
        .Cells(1, constGenericIndx).Value2 = "Generiek"
        .Cells(1, constProductIndx).Value2 = "Product"
        .Cells(1, constLabelIndx).Value2 = "Etiket"
        .Cells(1, constShapeIndx).Value2 = "Vorm"
        .Cells(1, constRouteIndx).Value2 = "Routes"
        .Cells(1, constGenericQuantityIndx).Value2 = "Sterkte"
        .Cells(1, constGenericQuantityUnitIndx).Value2 = "SterkteEenheid"
                    
        .Cells(1, constMultipleQuantityIndx).Value2 = "Veelvoud"
        .Cells(1, constMultipleQuantityUnitIndx).Value2 = "VeelvoudEenheid"
        
        .Cells(1, constIndicationsIndx).Value2 = "Indicaties"
        .Cells(1, constHasSolutionsIndx).Value2 = "HeeftOplossingen"
        .Cells(1, constIsActiveIndx).Value2 = "InAssortiment"
                        
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
        .Cells(1, constSolMinGenericQuantityIndx).Value2 = "MinGeneriekHoev"
        .Cells(1, constSolMaxGenericQuantityIndx).Value2 = "MaxGeneriekHoev"
        .Cells(1, constSolSolutionsIndx).Value2 = "Oplossingen"
        .Cells(1, constSolMinConcIndx).Value2 = "MinConcentratie"
        .Cells(1, constSolMaxConcIndx).Value2 = "MaxConcentratie"
        .Cells(1, constSolSolutionVolumeIndx).Value2 = "OplossingVolume"
        .Cells(1, constSolMinInfusionTimeIndx).Value2 = "MinInloopTijd"
    End With

    
    intN = 2
    intD = 2
    intS = 2
    
    intC = colSolution.Count
    For Each objSolution In colSolution
        With objSolution
            objSolSheet.Cells(intS, constSolDepartmentIndx).Value2 = .Department
            objSolSheet.Cells(intS, constSolGenericIndx).Value2 = .Generic
            objSolSheet.Cells(intS, constSolShapeIndx).Value2 = .Shape
            objSolSheet.Cells(intS, constSolMinGenericQuantityIndx).Value2 = .MinGenericQuantity
            objSolSheet.Cells(intS, constSolMaxGenericQuantityIndx).Value2 = .MaxGenericQuantity
            objSolSheet.Cells(intS, constSolSolutionsIndx).Value2 = .Solutions
            objSolSheet.Cells(intS, constSolSolutionVolumeIndx).Value2 = .SolutionVolume
            objSolSheet.Cells(intS, constSolMinConcIndx).Value2 = .MinConc
            objSolSheet.Cells(intS, constSolMaxConcIndx).Value2 = .MaxConc
            objSolSheet.Cells(intS, constSolMinInfusionTimeIndx).Value2 = .MinInfusionTime
        End With
        
        intS = intS + 1
        If blnShowProgress Then ModProgress.SetJobPercentage "Export " & intS, intC, intS
    Next
    
    
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
            
            If .NormDose > 0 Then objDoseSheet.Cells(intD, constDoseNormDoseIndx).Value2 = .NormDose
            If .MinDose > 0 Then objDoseSheet.Cells(intD, constDoseMinDoseIndx).Value2 = .MinDose
            If .MaxDose > 0 Then objDoseSheet.Cells(intD, constDoseMaxDoseIndx).Value2 = .MaxDose
            If .AbsMaxDose > 0 Then objDoseSheet.Cells(intD, constDoseAbsMaxDoseIndx).Value2 = .AbsMaxDose
            If .MaxPerDose > 0 Then objDoseSheet.Cells(intD, constDoseMaxPerDoseIndx).Value2 = .MaxPerDose
            objDoseSheet.Cells(intD, constDoseIsDosePerKgIndx).Value2 = IIf(.IsDosePerKg, "x", vbNullString)
            objDoseSheet.Cells(intD, constDoseIsDosePerM2Indx).Value2 = IIf(.IsDosePerM2, "x", vbNullString)
        
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
            
            arrFreq = Split(objDose.Frequencies, "||")
            For Each varFreq In arrFreq
                Select Case varFreq
                    Case constDoseFreq_AN
                        objDoseSheet.Cells(intD, 23).Value2 = "x"
                        
                    Case constDoseFreq_1D
                        objDoseSheet.Cells(intD, 24).Value2 = "x"
                
                    Case constDoseFreq_2D
                        objDoseSheet.Cells(intD, 25).Value2 = "x"
                
                    Case constDoseFreq_3D
                        objDoseSheet.Cells(intD, 26).Value2 = "x"
                
                    Case constDoseFreq_4D
                        objDoseSheet.Cells(intD, 27).Value2 = "x"
                
                    Case constDoseFreq_5D
                        objDoseSheet.Cells(intD, 28).Value2 = "x"
                
                    Case constDoseFreq_6D
                        objDoseSheet.Cells(intD, 29).Value2 = "x"
                
                    Case constDoseFreq_7D
                        objDoseSheet.Cells(intD, 30).Value2 = "x"
                
                    Case constDoseFreq_8D
                        objDoseSheet.Cells(intD, 31).Value2 = "x"
                
                    Case constDoseFreq_9D
                        objDoseSheet.Cells(intD, 32).Value2 = "x"
                
                    Case constDoseFreq_10D
                        objDoseSheet.Cells(intD, 33).Value2 = "x"
                
                    Case constDoseFreq_11D
                        objDoseSheet.Cells(intD, 34).Value2 = "x"
                
                    Case constDoseFreq_12D
                        objDoseSheet.Cells(intD, 35).Value2 = "x"
                
                    Case constDoseFreq_24D
                        objDoseSheet.Cells(intD, 36).Value2 = "x"
                
                    Case constDoseFreq_1D2
                        objDoseSheet.Cells(intD, 37).Value2 = "x"
                
                    Case constDoseFreq_1D3
                        objDoseSheet.Cells(intD, 38).Value2 = "x"
                
                    Case constDoseFreq_1U36
                        objDoseSheet.Cells(intD, 39).Value2 = "x"
                
                    Case constDoseFreq_1W
                        objDoseSheet.Cells(intD, 40).Value2 = "x"
                
                    Case constDoseFreq_2W
                        objDoseSheet.Cells(intD, 41).Value2 = "x"
                
                    Case constDoseFreq_3W
                        objDoseSheet.Cells(intD, 42).Value2 = "x"
                
                    Case constDoseFreq_4W
                        objDoseSheet.Cells(intD, 43).Value2 = "x"
                
                    Case constDoseFreq_1W2
                        objDoseSheet.Cells(intD, 44).Value2 = "x"
                
                    Case constDoseFreq_1W4
                        objDoseSheet.Cells(intD, 45).Value2 = "x"
                
                    Case constDoseFreq_1W12
                        objDoseSheet.Cells(intD, 46).Value2 = "x"
                
                    Case constDoseFreq_1W13
                        objDoseSheet.Cells(intD, 47).Value2 = "x"
                
                    Case constDoseFreq_1M
                        objDoseSheet.Cells(intD, 48).Value2 = "x"
                                
                End Select
                
            Next
        
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
            objMedSheet.Cells(intN, constHasSolutionsIndx).Value2 = IIf(objMed.HasSolutions, "x", " ")
            objMedSheet.Cells(intN, constIsActiveIndx).Value2 = IIf(objMed.IsActive, "x", " ")
            
        End With
        
        intN = intN + 1
        If blnShowProgress Then ModProgress.SetJobPercentage "Export " & intN, intC, intN
    Next
    
    Set objFormRange = objMedSheet.Range("A1").CurrentRegion
    Set objDoseRange = objDoseSheet.Range("A1").CurrentRegion
    Set objSolRange = objSolSheet.Range("A1").CurrentRegion
    
    If objFormRange.Rows.Count > 1 Then
        With objMedSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=objMedSheet.Range("E2"), Order:=xlAscending
            .SortFields.Add Key:=objMedSheet.Range("H2"), Order:=xlAscending
            .SortFields.Add Key:=objMedSheet.Range("J2"), Order:=xlAscending
            .SetRange objMedSheet.Range("A2:P" & objFormRange.Rows.Count)
            .Apply
        End With
    End If
    
    If objDoseRange.Rows.Count > 1 Then
        With objDoseSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=objDoseSheet.Range("A2"), Order:=xlAscending
            .SortFields.Add Key:=objDoseSheet.Range("D2"), Order:=xlAscending
            .SortFields.Add Key:=objDoseSheet.Range("C2"), Order:=xlAscending
            .SetRange objDoseSheet.Range("A2:BW" & objDoseRange.Rows.Count)
            .Apply
        End With
    End If
    
    If objSolRange.Rows.Count > 1 Then
        With objSolSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=objSolSheet.Range("B2"), Order:=xlAscending
            .SortFields.Add Key:=objSolSheet.Range("C2"), Order:=xlAscending
            .SortFields.Add Key:=objSolSheet.Range("A2"), Order:=xlAscending
            .SetRange objSolSheet.Range("A2:M" & objSolRange.Rows.Count)
            .Apply
        End With
    End If
    
    objWbk.Windows(1).Visible = True
    ImprovePerf False
    
    objWbk.SaveAs strFile
    objWbk.Close
    
    ModProgress.FinishProgress
    Application.DisplayAlerts = True
        
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

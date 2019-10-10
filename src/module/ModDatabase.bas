Attribute VB_Name = "ModDatabase"
Option Explicit

Private objConn As ADODB.Connection

Private Const constSecret As String = "secret"

Private Const CONST_CLEARDATABASE = "dbo.ClearDatabase"

Private Const CONST_GET_CONFIG_MEDCONT_VERSIONS = "[dbo].[GetConfigMedContVersionsForDepartment]"
Private Const CONST_GET_VERSION_CONFIG_MEDCONT = "[dbo].[GetConfigMedContForDepartmentWithVersion]"
Private Const CONST_GET_LATEST_CONFIG_MEDCONT = "[dbo].[GetConfigMedContForDepartmentLatest]"
Private Const CONST_INSERT_CONFIG_MEDCONT = "[dbo].[InsertConfigMedCont]"

Private Const CONST_GET_CONFIG_PARENT_VERSIONS = "[dbo].[GetConfigParEntVersions]"
Private Const CONST_GET_VERSION_CONFIG_PARENT = "dbo.GetConfigParEntForVersion"
Private Const CONST_GET_LATEST_CONFIG_PARENT = "[dbo].[GetConfigParEntLatest]"
Private Const CONST_INSERT_CONFIG_PARENT = "[dbo].[InsertConfigParEnt]"

Private Const CONST_GET_PRESCRIPTION_VERSIONS = "[dbo].[GetPrescriptionVersionsForHospitalNumber]"
Private Const CONST_GET_LATEST_PRESCRIPTION_VERSION = "dbo.GetLatestPrescriptionVersionForHospitalNumber"
Private Const CONST_GET_VERSION_PRESCRIPTIONDATA = "dbo.GetPrescriptionDataForVersion"
Private Const CONST_GET_LATEST_PRESCRIPTIONDATA = "dbo.GetPrescriptionDataLatest"
Private Const CONST_INSERT_PRESCRIPTIONDATA = "dbo.InsertPrescriptionData"

Private Const CONST_INSERT_PRESCRIPTIONTEXT = "dbo.InsertPrescriptionText"

Private Const CONST_INSERT_LOG = "dbo.InsertLog"

Private Const CONST_GET_PRESCRIBERS = "dbo.GetPrescribers"
Private Const CONST_INSERT_PRESCRIBER = "InsertPrescriber"
Private Const CONST_UPDATE_PRESCRIBER = "UpdatePrescriber"

Private Const CONST_GET_PATIENTS = "dbo.GetPatients"
Private Const CONST_INSERT_PATIENT = "InsertPatient"
Private Const CONST_UPDATE_PATIENT = "UpdatePatient"

Private Const constMedDiscTbl = "Medicatie"

Private Sub Util_InitConnection()

    Dim strSecret As String
    Dim strUser As String
    Dim strPw As String
    
    Dim strServer As String
    Dim strDatabase As String
    
    On Error GoTo InitConnectionError
    
    strServer = ModSetting.Setting_GetServer()
    strDatabase = ModSetting.Setting_GetDatabase()
    
    strSecret = ModFile.ReadFile(WbkAfspraken.Path & "/" & constSecret)
    
    If strSecret <> vbNullString Then
        strUser = Split(strSecret, vbLf)(0)
        strPw = Split(strSecret, vbLf)(1)
    
        Set objConn = New ADODB.Connection
        
        objConn.ConnectionString = "Provider=SQLOLEDB.1;" _
                 & "Server=" & strServer & ";" _
                 & "Database=" & strDatabase & ";" _
                 & "User ID=" & strUser & ";" _
                 & "Password=" & strPw & ";" _
                 & "DataTypeCompatibility=80;" _
                 & "MARS Connection=True;"
        ' Test de connectie
        objConn.Open
        objConn.Close
    Else
        MsgBox "Geen toegang tot de database!"
        ModLog.LogError Err, "Bestand secret niet aanwezig"
    End If
    
    Exit Sub
    
InitConnectionError:
    MsgBox "Geen toegang tot de database!"
    ModLog.LogError Err, "Util_InitConnection Failed"

End Sub

Private Sub Test_InitConnectionWithAPDB()

    Util_InitConnection

End Sub

Public Function Database_GetPatients() As Collection

    Dim colPatienten As Collection
    Dim strBed As String
    Dim strPN As String
    Dim strVN As String
    Dim strAN As String
    Dim strBD As String

    Dim strSql
    Dim objRs As Recordset
    
    Set colPatienten = New Collection

    Util_InitConnection
    
    objConn.Open
    
    strSql = "SELECT * FROM Patient"
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        strPN = objRs.Fields("HospitalNumber")
        strBed = ""
        strAN = objRs.Fields("LastName")
        strVN = objRs.Fields("FirstName")
        
        colPatienten.Add CreatePatientInfo(strPN, strBed, strAN, strVN, strBD)
        objRs.MoveNext
    Loop
    
    objConn.Close
    Set objRs = Nothing


    Set Database_GetPatients = colPatienten

End Function

Private Sub Test_GetPatients()

    Database_GetPatients

End Sub

Private Function Util_PatientExists(strHospN As String) As Boolean

    Dim strSql As String
    
    strSql = "SELECT * FROM " & CONST_GET_PATIENTS & " ('" & strHospN & "')"
    
    Util_InitConnection
    
    objConn.Open

    Util_PatientExists = Not objConn.Execute(strSql).EOF

End Function

Private Sub Test_Util_PatientExists()

    MsgBox Util_PatientExists("000")

End Sub

Private Function Util_PrescriberExists(strUser As String) As Boolean

    Dim strSql As String
    
    strSql = "SELECT * FROM " & CONST_GET_PRESCRIBERS & " (" & strUser & ")"
    
    Util_InitConnection
    
    objConn.Open

    Util_PrescriberExists = Not objConn.Execute(strSql).EOF

End Function

Private Sub Test_Util_PrescriberExists()

    MsgBox Util_PrescriberExists("000")

End Sub

Private Function Util_WrapString(varItem As Variant) As Variant

    Util_WrapString = "'" & varItem & "'"

End Function

Private Function Util_WrapDateTime(strDateTime As String) As String

    Util_WrapDateTime = "{ts'" & strDateTime & "'}"

End Function

Private Function Util_WrapTransaction(ByVal strSql As String, ByVal strName As String) As String

    Dim strTrans As String
    
    strTrans = "BEGIN TRANSACTION [" & strName & "]" & vbNewLine
    strTrans = strTrans & "BEGIN TRY" & vbNewLine
    strTrans = strTrans & strSql & vbNewLine
    strTrans = strTrans & "COMMIT TRANSACTION [" & strName & "]" & vbNewLine
    strTrans = strTrans & "END TRY" & vbNewLine
    strTrans = strTrans & "BEGIN CATCH" & vbNewLine
    strTrans = strTrans & "ROLLBACK TRANSACTION [" & strName & "]" & vbNewLine
    strTrans = strTrans & "END CATCH"
    
    Util_WrapTransaction = strTrans

End Function

Public Sub Database_SavePatient()

    Dim strHN As String
    Dim strBD As String
    Dim strVN As String
    Dim strAN As String
    Dim strGN As String
    Dim intGW As Integer
    Dim intGD As Integer
    Dim dblBW As Double
    
    Dim strSql As String
    Dim arrSql() As Variant
    
    On Error GoTo SavePatientError
    
    strHN = Util_WrapString(ModPatient.Patient_GetHospitalNumber)
    strBD = Util_WrapString(ModDate.FormatDateYearMonthDay(ModPatient.Patient_BirthDate))
    strAN = Util_WrapString(ModPatient.Patient_GetLastName)
    strVN = Util_WrapString(ModPatient.Patient_GetFirstName)
    strGN = Util_WrapString(ModRange.GetRangeValue(CONST_GENDER_RANGE, Null))
    intGW = ModRange.GetRangeValue(CONST_GESTWEEKS_RANGE, Null)
    intGD = ModRange.GetRangeValue(CONST_GESTDAYS_RANGE, Null)
    dblBW = ModRange.GetRangeValue(CONST_BIRTHWEIGHT_RANGE, Null)
        
    arrSql = Array(strHN, strBD, strAN, strVN, strGN, intGW, intGD, dblBW)
        
    Util_InitConnection
    
    objConn.Open
    
    If Util_PatientExists(ModPatient.Patient_GetHospitalNumber()) Then
        strSql = "EXEC " & CONST_UPDATE_PATIENT & " "
    Else
        strSql = "EXEC " & CONST_INSERT_PATIENT & " "
    End If
    
    strSql = strSql + (Join(arrSql, ", "))
    strSql = Util_WrapTransaction(strSql, "save_patient")
    
    objConn.Execute strSql
    
    objConn.Close
    
    Exit Sub
    
SavePatientError:

    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not save patient details to database: " & strSql
    
End Sub

Public Function Database_GetLastStandardPatientHospNum() As String

    Dim strSql As String
    Dim strHospNum As String
    Dim objRs As Recordset
    
    strSql = "SELECT dbo.GetLastStandardPatientHospNum()"

    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    If Not objRs.EOF Then
        strHospNum = IIf(IsNull(objRs.Fields(0)), vbNullString, objRs.Fields(0).Value)
    Else
        strHospNum = 0
    End If
    
    objConn.Close
    
    Database_GetLastStandardPatientHospNum = strHospNum
    
End Function

Private Sub Test_Database_GetLastStandardPatientHospNum()

    ModMessage.ShowMsgBoxInfo Database_GetLastStandardPatientHospNum()

End Sub

Public Sub Database_SavePrescriber()

    Dim strUser As String
    Dim strLN As String
    Dim strFN As String
    Dim strRole As String
    Dim strPIN As String
    
    Dim strSql As String
    Dim arrSql() As Variant
    
    On Error GoTo SavePrescriberError
    
    strUser = Util_WrapString(ModRange.GetRangeValue("_User_Login", ""))
    strLN = Util_WrapString(ModRange.GetRangeValue("_User_FirstName", ""))
    strFN = Util_WrapString(ModRange.GetRangeValue("_User_LastName", ""))
    strRole = Util_WrapString(ModRange.GetRangeValue("_User_Type", ""))
        
    arrSql = Array(strUser, strLN, strFN, strRole)
        
    Util_InitConnection
    
    objConn.Open
    
    If Util_PrescriberExists(strUser) Then
        strSql = "EXEC " & CONST_UPDATE_PRESCRIBER & " "
    Else
        strSql = "EXEC " & CONST_INSERT_PRESCRIBER & " "
    End If
    
    strSql = strSql & (Join(arrSql, ", "))
    strSql = Util_WrapTransaction(strSql, "save_prescriber")
    
    ModUtils.CopyToClipboard strSql
    objConn.Execute strSql
    
    objConn.Close
    
    Exit Sub
    
SavePrescriberError:

    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not save prescriber details to the database: " & strSql
    
End Sub

Public Sub Database_ClearTestDatabase()

    Dim strSql As String
    
    strSql = "EXEC " & CONST_CLEARDATABASE & "  'UMCU_WKZ_AP_Test'"

    Util_InitConnection
    
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    Exit Sub
    
ClearTestDatabaseError:

    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not clear the database: " & strSql

End Sub

Private Sub Test_SavePrescriber()

    Database_SavePrescriber
    ModMessage.ShowMsgBoxOK Util_PrescriberExists(ModMetaVision.MetaVision_GetUserLogin())

End Sub

Public Function Database_GetLatestPrescriptionVersion(strHospNum) As String

    Dim strSql As String
    Dim objRs As Recordset
    Dim intVersion As Integer
    
    On Error GoTo Database_GetLatestVersionError
    
    strSql = "SELECT " & CONST_GET_LATEST_PRESCRIPTION_VERSION & "('" & strHospNum & "')"
    
    Util_InitConnection
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    If Not objRs.EOF Then
        intVersion = IIf(IsNull(objRs.Fields(0)), 0, objRs.Fields(0).Value)
    Else
        intVersion = 0
    End If

    objConn.Close
    Set objRs = Nothing
    
    Database_GetLatestPrescriptionVersion = intVersion
    
    Exit Function
    
Database_GetLatestVersionError:

    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not get latest version for patient: " & strHospNum & " with SQL: " & vbNewLine & strSql
    objConn.Close

End Function

Private Sub Test_Database_GetLatestPrescriptionVersion()

    ModMessage.ShowMsgBoxOK Database_GetLatestPrescriptionVersion("1234")

End Sub

Public Sub Database_SaveData(ByVal strHospNum As String, ByVal strPrescriber As String, objData As Range, objText As Range, ByVal blnProgress As Boolean)

    Dim strParam As String
    Dim strSql As String
    Dim strLatest As String
    Dim varVal As Variant
    Dim varEmp As Variant
    Dim intVersion As Integer
    
    Dim intC As Integer
    Dim intN As Integer
    
    Dim objBuilder As ClassStringBuilder
        
    On Error GoTo SaveDataError
    
    Set objBuilder = New ClassStringBuilder
    
    objBuilder.Append "DECLARE @RC int" & vbNewLine
    objBuilder.Append "DECLARE @versionID int" & vbNewLine
    objBuilder.Append "DECLARE @versionUTC datetime" & vbNewLine
    objBuilder.Append "DECLARE @versionDate datetime" & vbNewLine
    
    strLatest = "SELECT @versionID = " & CONST_GET_LATEST_PRESCRIPTION_VERSION & "('" & strHospNum & "')"
    strLatest = Util_GetVersionSQL(strLatest) & vbNewLine
    objBuilder.Append vbNewLine & strLatest
       
    intC = objData.Rows.Count
    For intN = 2 To intC
        strParam = objData.Cells(intN, 1).Value2
        varVal = objData.Cells(intN, 2).Value2
        varEmp = objData.Cells(intN, 3).Value2
        
        If Not varVal = varEmp Then
            objBuilder.Append vbNewLine & "EXEC " & CONST_INSERT_PRESCRIPTIONDATA & " '" & strHospNum & "', @versionID, @versionUTC, @versionDate, '" & strPrescriber & "', 0, ' " & strParam & " ', '" & varVal & " '"
        End If
        
        If blnProgress Then ModProgress.SetJobPercentage "Data wegschrijven", intC, intN
    Next intN
    
    objBuilder.Append Util_GetLogSQL("Save patient data", False, strHospNum, "PrescriptionData")
    objBuilder.Append vbNewLine
    objBuilder.Append vbNewLine
    
    intC = objText.Rows.Count
    For intN = 2 To intC
        If Not (Format(objText.Cells(intN, 2).Value2) = vbNullString Or Format(objText.Cells(intN, 2).Value2) = "0") Then
            strParam = objText.Cells(intN, 1).Value2
            varVal = objText.Cells(intN, 2).Value2
            objBuilder.Append vbNewLine & "EXEC " & CONST_INSERT_PRESCRIPTIONTEXT & " '" & strHospNum & "', @versionID, @versionUTC, @versionDate, '" & strPrescriber & "', 0, ' " & strParam & " ', '" & varVal & " '"
        End If
        
        If blnProgress Then ModProgress.SetJobPercentage "Text wegschrijven naar de database", intC, intN
    Next intN
    
    objBuilder.Append Util_GetLogSQL("Save patient data", False, strHospNum, "PrescriptionText")
    objBuilder.Append vbNewLine
    objBuilder.Append vbNewLine
    
    strSql = objBuilder.ToString()
    strSql = ModDatabase.Util_WrapTransaction(strSql, "save_data")
    ModUtils.CopyToClipboard strSql
    
    objConn.Open
    objConn.Execute strSql, adExecuteNoRecords
    objConn.Close
    
    ModBed.Bed_PrescriptionsVersionSet Database_GetLatestPrescriptionVersion(strHospNum)
    
    Exit Sub

SaveDataError:
    
    ModMessage.ShowMsgBoxError "Kan patient data niet opslaan"
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not save patient data to the database"
    
    objConn.Close
    
End Sub

Private Function Util_IsLogical(ByVal varVal As Variant) As Boolean

    Util_IsLogical = LCase(varVal) = "waar" Or LCase(varVal) = "onwaar"
    
End Function

Private Sub Util_GetPatientDataForHospNumAndVersion(ByVal strHospNum, Optional ByVal intVersion As Integer = 0)

    Dim strSql As String
    Dim intC As Long
    Dim intN As Long
    Dim strPar As String
    Dim varVal As Variant
    Dim varEmp As Variant
    Dim objRs As Recordset
    Dim blnVersionSet As Boolean
    Dim blnIsStandard As Boolean
    
    On Error GoTo Database_GetPatientDataError
    
    ImprovePerf True
    
    strSql = strSql & "SELECT * FROM "
    If intVersion = 0 Then
        strSql = strSql & CONST_GET_LATEST_PRESCRIPTIONDATA & "('" & strHospNum & "')"
    Else
        strSql = strSql & CONST_GET_VERSION_PRESCRIPTIONDATA & "('" & strHospNum & "', " & intVersion & ")"
    End If
    
    Util_InitConnection
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    intC = shtPatData.Range("A1").CurrentRegion.Rows.Count
    ' Determine if the current patient is a standard patient and not a standard patient applied to a patient
    blnIsStandard = Patient_IsStandard(strHospNum) And (Patient_GetHospitalNumber() = vbNullString Or Patient_IsStandard(Patient_GetHospitalNumber))
    blnVersionSet = Patient_IsStandard(strHospNum) And Not blnIsStandard ' Patient standard is applied, keep the current version
    Do While Not objRs.EOF
        If Not blnVersionSet Then
            ModRange.SetRangeValue CONST_PRESCRIPTIONS_VERSION, objRs.Fields("VersionID").Value
            blnVersionSet = True
        End If
        
        strPar = Trim(objRs.Fields("Parameter").Value)
        varVal = Trim(objRs.Fields("Data").Value)
        
        If IsNumeric(varVal) Then varVal = CDbl(varVal)
        If Util_IsLogical(varVal) Then varVal = CBool(varVal)
        
        If Patient_IsStandard(strHospNum) And ModString.StartsWith(strPar, "__") And Not blnIsStandard Then
            'Do not change patient details for loading standard data
        Else
            ModRange.SetRangeValue strPar, varVal
        End If
        
        intN = intN + 1
        ModProgress.SetJobPercentage "Patient data laden", intC, intN
        
        objRs.MoveNext
    Loop
    
    objConn.Close
    ImprovePerf False
    
    Exit Sub

Database_GetPatientDataError:
    
    ModMessage.ShowMsgBoxError "Kan patient met ziekenhuis nummer " & strHospNum & " niet laden."
    
    ModLog.LogError Err, "Could not get patient data with hospitalnumber " & strHospNum & " with SQL: " & vbNewLine & strSql
    objConn.Close
    ImprovePerf False

End Sub


Public Sub Database_GetPatientDataForVersion(strHospNum As String, intVersion)

    Util_GetPatientDataForHospNumAndVersion strHospNum, intVersion
    
End Sub

Public Sub Database_GetPatientData(strHospNum As String)

    Util_GetPatientDataForHospNumAndVersion strHospNum
    
End Sub

Private Sub Test_DatabaseGetPatientData()

    ModProgress.StartProgress "Patient data ophalen"
    Database_GetPatientData "0250574"
    ModProgress.FinishProgress

End Sub

Private Function Util_GetVersionSQL(strGetLatest As String) As String

    Dim strSql As String
    
    strSql = strSql & strGetLatest & vbNewLine
    strSql = strSql & "SET @versionID  = COALESCE(@versionID, 0) + 1" & vbNewLine
    strSql = strSql & "SET @versionUTC = GETUTCDATE()" & vbNewLine
    strSql = strSql & "SET @versionDate = GETDATE()" & vbNewLine
    strSql = strSql & "" & vbNewLine

    Util_GetVersionSQL = strSql

End Function

Private Function Util_GetSaveNeoConfigMedContSql(blnIsBatch As Boolean) As String

    Dim strTable As String
    
    Dim strLatest As String
    
    Dim strDepartment As String
    Dim strGeneric As String
    Dim strGenericUnit As String
    Dim dblGenericQuantity As Double
    Dim dblGenericVolume As Double
    Dim dblSolutionVolume As Double
    Dim dblSolution_2_6_Quantity As Double
    Dim dblSolution_2_6_Volume As Double
    Dim dblSolution_6_11_Quantity As Double
    Dim dblSolution_6_11_Volume As Double
    Dim dblSolution_11_40_Quantity As Double
    Dim dblSolution_11_40_Volume As Double
    Dim dblSolution_40_Quantity As Double
    Dim dblSolution_40_Volume As Double
    Dim dblMinConcentration As Double
    Dim dblMaxConcentration As Double
    Dim strSolution As String
    Dim intSolutionRequired As Integer
    Dim dblDripQuantity As Double
    Dim strDoseUnit As String
    Dim dblMinDose As Double
    Dim dblMaxDose As Double
    Dim dblAbsMaxDose As Double
    Dim strDoseAdvice As String
    Dim strProduct As String
    Dim dblShelfLife As Double
    Dim strShelfCondition As String
    Dim strPreparationText As String
    Dim blnSigned As Boolean
    Dim strDilutionText As String
    
    Dim intR As Integer
    Dim intC As Integer
    
    Dim objSrc As Range
    Dim objBuilder As ClassStringBuilder
    Dim strSql
        
    strTable = "Tbl_Admin_NeoMedCont"
    strDepartment = CONST_DEP_NICU
    strDilutionText = ModRange.GetRangeValue("Var_Neo_MedCont_VerdunningTekst", vbNullString)
    
    Set objBuilder = New ClassStringBuilder
    Set objSrc = ModRange.GetRange(strTable)
    If Not blnIsBatch Then
    
        objBuilder.Append "DECLARE @RC int" & vbNewLine
        objBuilder.Append "DECLARE @versionID int" & vbNewLine
        objBuilder.Append "DECLARE @versionUTC datetime" & vbNewLine
        objBuilder.Append "DECLARE @versionDate datetime" & vbNewLine
        
        objBuilder.Append "DECLARE @department nvarchar(60)" & vbNewLine
        objBuilder.Append "DECLARE @generic nvarchar(300)" & vbNewLine
        objBuilder.Append "DECLARE @genericUnit nvarchar(50)" & vbNewLine
        objBuilder.Append "DECLARE @genericQuantity float" & vbNewLine
        objBuilder.Append "DECLARE @genericVolume float" & vbNewLine
        objBuilder.Append "DECLARE @solutionVolume float" & vbNewLine
        objBuilder.Append "DECLARE @solution_2_6_Quantity float" & vbNewLine
        objBuilder.Append "DECLARE @solution_2_6_Volume float" & vbNewLine
        objBuilder.Append "DECLARE @solution_6_11_Quantity float" & vbNewLine
        objBuilder.Append "DECLARE @solution_6_11_Volume float" & vbNewLine
        objBuilder.Append "DECLARE @solution_11_40_Quantity float" & vbNewLine
        objBuilder.Append "DECLARE @solution_11_40_Volume float" & vbNewLine
        objBuilder.Append "DECLARE @solution_40_Quantity float" & vbNewLine
        objBuilder.Append "DECLARE @solution_40_Volume float" & vbNewLine
        objBuilder.Append "DECLARE @minConcentration float" & vbNewLine
        objBuilder.Append "DECLARE @maxConcentration float" & vbNewLine
        objBuilder.Append "DECLARE @solution nvarchar(300)" & vbNewLine
        objBuilder.Append "DECLARE @solutionRequired bit" & vbNewLine
        objBuilder.Append "DECLARE @dripQuantity float" & vbNewLine
        objBuilder.Append "DECLARE @doseUnit nvarchar(50)" & vbNewLine
        objBuilder.Append "DECLARE @minDose float" & vbNewLine
        objBuilder.Append "DECLARE @maxDose float" & vbNewLine
        objBuilder.Append "DECLARE @absMaxDose float" & vbNewLine
        objBuilder.Append "DECLARE @doseAdvice nvarchar(max)" & vbNewLine
        objBuilder.Append "DECLARE @product nvarchar(max)" & vbNewLine
        objBuilder.Append "DECLARE @shelfLife float" & vbNewLine
        objBuilder.Append "DECLARE @shelfCondition nvarchar(50)" & vbNewLine
        objBuilder.Append "DECLARE @preparationText nvarchar(max)" & vbNewLine
        objBuilder.Append "DECLARE @signed bit" & vbNewLine
        objBuilder.Append "DECLARE @dilutionText nvarchar(max)" & vbNewLine
        objBuilder.Append "" & vbNewLine
    
    End If
        
    strLatest = "SET @department  = '" & strDepartment & "'" & vbNewLine
    strLatest = strLatest & "SELECT @versionID = dbo.GetLatestConfigMedContVersionForDepartment(@department)"
    
    objBuilder.Append Util_GetVersionSQL(strLatest)
        
        
    intC = objSrc.Rows.Count
    For intR = 1 To intC
    
        strGeneric = objSrc.Cells(intR, 1).Value2
        strGenericUnit = objSrc.Cells(intR, 2).Value2
        strDoseUnit = objSrc.Cells(intR, 3).Value2
        dblGenericQuantity = objSrc.Cells(intR, 4).Value2
        dblGenericVolume = objSrc.Cells(intR, 5).Value2
        dblMinDose = objSrc.Cells(intR, 6).Value2
        dblMaxDose = objSrc.Cells(intR, 7).Value2
        dblAbsMaxDose = objSrc.Cells(intR, 8).Value2
        dblMinConcentration = objSrc.Cells(intR, 9).Value2
        dblMaxConcentration = objSrc.Cells(intR, 10).Value2
        strSolution = objSrc.Cells(intR, 11).Value2
        intSolutionRequired = 0
        If objSrc.Cells(intR, 19).Value Then intSolutionRequired = 1
        strDoseAdvice = objSrc.Cells(intR, 12).Value2
        dblSolutionVolume = objSrc.Cells(intR, 13).Value2
        dblDripQuantity = objSrc.Cells(intR, 14).Value2
        strProduct = objSrc.Cells(intR, 15).Value2
        dblShelfLife = objSrc.Cells(intR, 16).Value2
        strShelfCondition = objSrc.Cells(intR, 17).Value2
        strPreparationText = objSrc.Cells(intR, 18).Value2
            
        objBuilder.Append "SET @generic  = '" & strGeneric & "'" & vbNewLine
        objBuilder.Append "SET @genericUnit  = '" & strGenericUnit & "'" & vbNewLine
        objBuilder.Append "SET @genericQuantity  =  " & DoubleToString(dblGenericQuantity) & vbNewLine
        objBuilder.Append "SET @genericVolume  =  " & DoubleToString(dblGenericVolume) & vbNewLine
        objBuilder.Append "SET @solutionVolume  =  " & DoubleToString(dblSolutionVolume) & vbNewLine
        objBuilder.Append "SET @solution_2_6_Quantity  =  0" & vbNewLine
        objBuilder.Append "SET @solution_2_6_Volume  =  0" & vbNewLine
        objBuilder.Append "SET @solution_6_11_Quantity  =  0" & vbNewLine
        objBuilder.Append "SET @solution_6_11_Volume  =  0" & vbNewLine
        objBuilder.Append "SET @solution_11_40_Quantity  =  0" & vbNewLine
        objBuilder.Append "SET @solution_11_40_Volume  =  0" & vbNewLine
        objBuilder.Append "SET @solution_40_Quantity  =  0" & vbNewLine
        objBuilder.Append "SET @solution_40_Volume  =  0" & vbNewLine
        objBuilder.Append "SET @minConcentration  = " & DoubleToString(dblMinConcentration) & vbNewLine
        objBuilder.Append "SET @maxConcentration  = " & DoubleToString(dblMaxConcentration) & vbNewLine
        objBuilder.Append "SET @solution  = '" & strSolution & "'" & vbNewLine
        objBuilder.Append "SET @solutionRequired  = " & intSolutionRequired & vbNewLine
        objBuilder.Append "SET @dripQuantity  =  " & DoubleToString(dblDripQuantity) & vbNewLine
        objBuilder.Append "SET @doseUnit  = '" & strDoseUnit & "'" & vbNewLine
        objBuilder.Append "SET @minDose  =  " & DoubleToString(dblMinDose) & vbNewLine
        objBuilder.Append "SET @maxDose  =  " & DoubleToString(dblMaxDose) & vbNewLine
        objBuilder.Append "SET @absMaxDose  =  " & DoubleToString(dblAbsMaxDose) & vbNewLine
        objBuilder.Append "SET @doseAdvice  = '" & strDoseAdvice & "'" & vbNewLine
        objBuilder.Append "SET @product  =  '" & strProduct & "'" & vbNewLine
        objBuilder.Append "SET @shelfLife  =  " & DoubleToString(dblShelfLife) & vbNewLine
        objBuilder.Append "SET @shelfCondition  = '" & strShelfCondition & "'" & vbNewLine
        objBuilder.Append "SET @preparationText  =  '" & strPreparationText & "'" & vbNewLine
        objBuilder.Append "SET @signed = 1" & vbNewLine
        objBuilder.Append "SET @dilutionText  = '" & strDilutionText & "'" & vbNewLine
    
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "EXECUTE @RC = " & CONST_INSERT_CONFIG_MEDCONT & vbNewLine
        objBuilder.Append "   @versionID" & vbNewLine
        objBuilder.Append "  ,@versionUTC" & vbNewLine
        objBuilder.Append "  ,@versionDate" & vbNewLine
        objBuilder.Append "  ,@department" & vbNewLine
        objBuilder.Append "  ,@generic" & vbNewLine
        objBuilder.Append "  ,@genericUnit" & vbNewLine
        objBuilder.Append "  ,@genericQuantity" & vbNewLine
        objBuilder.Append "  ,@genericVolume" & vbNewLine
        objBuilder.Append "  ,@solutionVolume" & vbNewLine
        objBuilder.Append "  ,@solution_2_6_Quantity" & vbNewLine
        objBuilder.Append "  ,@solution_2_6_Volume" & vbNewLine
        objBuilder.Append "  ,@solution_6_11_Quantity" & vbNewLine
        objBuilder.Append "  ,@solution_6_11_Volume" & vbNewLine
        objBuilder.Append "  ,@solution_11_40_Quantity" & vbNewLine
        objBuilder.Append "  ,@solution_11_40_Volume" & vbNewLine
        objBuilder.Append "  ,@solution_40_Quantity" & vbNewLine
        objBuilder.Append "  ,@solution_40_Volume" & vbNewLine
        objBuilder.Append "  ,@minConcentration" & vbNewLine
        objBuilder.Append "  ,@maxConcentration" & vbNewLine
        objBuilder.Append "  ,@solution" & vbNewLine
        objBuilder.Append "  ,@solutionRequired" & vbNewLine
        objBuilder.Append "  ,@dripQuantity" & vbNewLine
        objBuilder.Append "  ,@doseUnit" & vbNewLine
        objBuilder.Append "  ,@minDose" & vbNewLine
        objBuilder.Append "  ,@maxDose" & vbNewLine
        objBuilder.Append "  ,@absMaxDose" & vbNewLine
        objBuilder.Append "  ,@doseAdvice" & vbNewLine
        objBuilder.Append "  ,@product" & vbNewLine
        objBuilder.Append "  ,@shelfLife" & vbNewLine
        objBuilder.Append "  ,@shelfCondition" & vbNewLine
        objBuilder.Append "  ,@preparationText" & vbNewLine
        objBuilder.Append "  ,@signed" & vbNewLine
        objBuilder.Append "  ,@dilutionText" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    
    Next
    
    objBuilder.Append vbNewLine
    objBuilder.Append Util_GetLogSQL("Save neonatal configuration for continuous medication", False, , "ConfigMedCont")

    strSql = objBuilder.ToString()
    Util_GetSaveNeoConfigMedContSql = strSql
    
End Function

Private Sub Test_Util_GetSaveNeoConfigMedContSql()

    ModUtils.CopyToClipboard Util_GetSaveNeoConfigMedContSql(False)

End Sub

Public Sub Database_SaveNeoConfigMedCont()

    Dim strSql As String
    
    On Error GoTo ErrorHandler
     
    ModProgress.StartProgress "Neo Continue Medicatie Configuratie Opslaan"

    strSql = Util_GetSaveNeoConfigMedContSql(False)
    strSql = ModDatabase.Util_WrapTransaction(strSql, "insert_neoconfigmedcont")
    
    Util_InitConnection
    
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:

    objConn.Close
    ModProgress.FinishProgress

    ModUtils.CopyToClipboard strSql
    ModMessage.ShowMsgBoxError "Kon de configuratie voor de neonatologie continue medicatie niet opslaan"
    ModLog.LogError Err, "Database_SaveNeoConfigMedCont with sql: " & vbNewLine & strSql
    
End Sub

Public Sub Database_LoadNeoConfigMedCont()

    Dim strSql As String
    Dim objRs As Recordset
    Dim intC As Integer
    Dim intT As Integer
    Dim intR As Integer
    Dim objSrc As Range
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Configuratie voor NICU Continue Medicatie laden"
    
    Set objSrc = ModRange.GetRange("Tbl_Admin_NeoMedCont")
    intT = objSrc.Rows.Count
    
    Util_InitConnection
    
    strSql = "SELECT * FROM " & CONST_GET_LATEST_CONFIG_MEDCONT & " ('" & CONST_DEP_NICU & "')"

    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    ImprovePerf True
    Application.Calculation = xlCalculationManual
    
    Do While Not objRs.EOF
        intR = intR + 1
        If intR > intT Then GoTo ErrorHandler
        
        objSrc.Cells(intR, 1).Value2 = objRs.Fields("Generic").Value
        objSrc.Cells(intR, 2).Value2 = objRs.Fields("GenericUnit").Value
        objSrc.Cells(intR, 3).Value2 = objRs.Fields("DoseUnit").Value
        objSrc.Cells(intR, 4).Value2 = objRs.Fields("GenericQuantity").Value
        objSrc.Cells(intR, 5).Value2 = objRs.Fields("GenericVolume").Value
        objSrc.Cells(intR, 6).Value2 = objRs.Fields("MinDose").Value
        objSrc.Cells(intR, 7).Value2 = objRs.Fields("MaxDose").Value
        objSrc.Cells(intR, 8).Value2 = objRs.Fields("AbsMaxDose").Value
        objSrc.Cells(intR, 9).Value2 = objRs.Fields("MinConcentration").Value
        objSrc.Cells(intR, 10).Value2 = objRs.Fields("MaxConcentration").Value
        objSrc.Cells(intR, 11).Value2 = objRs.Fields("Solution").Value
        If objRs.Fields("SolutionRequired").Value Then
            objSrc.Cells(intR, 19).Value2 = True
        Else
            objSrc.Cells(intR, 19).Value2 = False
        End If
        objSrc.Cells(intR, 12).Value2 = objRs.Fields("DoseAdvice").Value
        objSrc.Cells(intR, 13).Value2 = objRs.Fields("SolutionVolume").Value
        objSrc.Cells(intR, 14).Value2 = objRs.Fields("DripQuantity").Value
        objSrc.Cells(intR, 15).Value2 = objRs.Fields("Product").Value
        objSrc.Cells(intR, 16).Value2 = objRs.Fields("ShelfLife").Value
        objSrc.Cells(intR, 17).Value2 = objRs.Fields("ShelfCondition").Value
        objSrc.Cells(intR, 18).Value2 = objRs.Fields("PreparationText").Value
        
        ModRange.SetRangeValue "Var_Neo_MedCont_VerdunningTekst", objRs.Fields("DilutionText").Value
        
        ModProgress.SetJobPercentage "Data laden", 24, intR
        objRs.MoveNext
    Loop
    
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ImprovePerf False
    
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ImprovePerf False

    ModProgress.FinishProgress
    objConn.Close

    ModUtils.CopyToClipboard strSql
    ModMessage.ShowMsgBoxError "Kon de configuratie voor de neonatologie continue medicatie niet laden"
    ModLog.LogError Err, "Database_LoadNeoConfigMedCont with sql: " & vbNewLine & strSql

End Sub

Private Function Util_GetSavePediatrieConfigMedContSql(ByVal blnIsBatch As Boolean) As String

    Dim strSql As String
    Dim strLatest As String
    Dim strTable As String
    
    Dim strDepartment As String
    Dim strGeneric As String
    Dim strGenericUnit As String
    Dim dblGenericQuantity As Double
    Dim dblGenericVolume As Double
    Dim dblSolutionVolume As Double
    Dim dblSolution_2_6_Quantity As Double
    Dim dblSolution_2_6_Volume As Double
    Dim dblSolution_6_11_Quantity As Double
    Dim dblSolution_6_11_Volume As Double
    Dim dblSolution_11_40_Quantity As Double
    Dim dblSolution_11_40_Volume As Double
    Dim dblSolution_40_Quantity As Double
    Dim dblSolution_40_Volume As Double
    Dim dblMinConcentration As Double
    Dim dblMaxConcentration As Double
    Dim strSolution As String
    Dim dblDripQuantity As Double
    Dim strDoseUnit As String
    Dim dblMinDose As Double
    Dim dblMaxDose As Double
    Dim dblAbsMaxDose As Double
    Dim strDoseAdvice As String
    Dim strProduct As String
    Dim dblShelfLife As Double
    Dim strShelfCondition As String
    Dim strPreparationText As String
    Dim blnSigned As Boolean
    Dim strDilutionText As String
    
    Dim intR As Integer
    Dim intC As Integer
    
    Dim objSrc As Range
    Dim objBuilder As ClassStringBuilder
    
    Set objBuilder = New ClassStringBuilder
    
    strTable = "Tbl_Admin_PedMedCont"
    strDepartment = CONST_DEP_PICU
    strDilutionText = ""
    
    Set objSrc = ModRange.GetRange(strTable)
    If Not blnIsBatch Then objBuilder.Append "DECLARE @RC int" & vbNewLine
    If Not blnIsBatch Then objBuilder.Append "DECLARE @versionID int" & vbNewLine
    If Not blnIsBatch Then objBuilder.Append "DECLARE @versionUTC datetime" & vbNewLine
    If Not blnIsBatch Then objBuilder.Append "DECLARE @versionDate datetime" & vbNewLine
    objBuilder.Append "DECLARE @department nvarchar(60)" & vbNewLine
    objBuilder.Append "DECLARE @generic nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @genericUnit nvarchar(50)" & vbNewLine
    objBuilder.Append "DECLARE @genericQuantity float" & vbNewLine
    objBuilder.Append "DECLARE @genericVolume float" & vbNewLine
    objBuilder.Append "DECLARE @solutionVolume float" & vbNewLine
    objBuilder.Append "DECLARE @solution_2_6_Quantity float" & vbNewLine
    objBuilder.Append "DECLARE @solution_2_6_Volume float" & vbNewLine
    objBuilder.Append "DECLARE @solution_6_11_Quantity float" & vbNewLine
    objBuilder.Append "DECLARE @solution_6_11_Volume float" & vbNewLine
    objBuilder.Append "DECLARE @solution_11_40_Quantity float" & vbNewLine
    objBuilder.Append "DECLARE @solution_11_40_Volume float" & vbNewLine
    objBuilder.Append "DECLARE @solution_40_Quantity float" & vbNewLine
    objBuilder.Append "DECLARE @solution_40_Volume float" & vbNewLine
    objBuilder.Append "DECLARE @minConcentration float" & vbNewLine
    objBuilder.Append "DECLARE @maxConcentration float" & vbNewLine
    objBuilder.Append "DECLARE @solution nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @solutionRequired bit" & vbNewLine
    objBuilder.Append "DECLARE @dripQuantity float" & vbNewLine
    objBuilder.Append "DECLARE @doseUnit nvarchar(50)" & vbNewLine
    objBuilder.Append "DECLARE @minDose float" & vbNewLine
    objBuilder.Append "DECLARE @maxDose float" & vbNewLine
    objBuilder.Append "DECLARE @absMaxDose float" & vbNewLine
    objBuilder.Append "DECLARE @doseAdvice nvarchar(max)" & vbNewLine
    If Not blnIsBatch Then objBuilder.Append "DECLARE @product nvarchar(max)" & vbNewLine
    objBuilder.Append "DECLARE @shelfLife float" & vbNewLine
    objBuilder.Append "DECLARE @shelfCondition nvarchar(50)" & vbNewLine
    objBuilder.Append "DECLARE @preparationText nvarchar(max)" & vbNewLine
    If Not blnIsBatch Then objBuilder.Append "DECLARE @signed bit" & vbNewLine
    objBuilder.Append "DECLARE @dilutionText nvarchar(max)" & vbNewLine
    objBuilder.Append "" & vbNewLine
    
    strLatest = "SET @department  = '" & strDepartment & "'" & vbNewLine
    strLatest = strLatest & "SELECT @versionID = dbo.GetLatestConfigMedContVersionForDepartment(@department)"
    
    objBuilder.Append Util_GetVersionSQL(strLatest)
        
    intC = objSrc.Rows.Count
    For intR = 1 To intC
    
        strGeneric = objSrc.Cells(intR, 1).Value2
        strGenericUnit = objSrc.Cells(intR, 2).Value2
        strDoseUnit = objSrc.Cells(intR, 3).Value2
                
        dblMinDose = objSrc.Cells(intR, 12).Value2
        dblMaxDose = objSrc.Cells(intR, 13).Value2
        dblAbsMaxDose = objSrc.Cells(intR, 14).Value2
        
        dblMinConcentration = objSrc.Cells(intR, 15).Value2
        dblMaxConcentration = objSrc.Cells(intR, 16).Value2
        
        strSolution = objSrc.Cells(intR, 17).Value2
        
        dblSolution_2_6_Quantity = objSrc.Cells(intR, 4).Value2
        dblSolution_2_6_Volume = objSrc.Cells(intR, 5).Value2
        dblSolution_6_11_Quantity = objSrc.Cells(intR, 6).Value2
        dblSolution_6_11_Volume = objSrc.Cells(intR, 7).Value2
        dblSolution_11_40_Quantity = objSrc.Cells(intR, 8).Value2
        dblSolution_11_40_Volume = objSrc.Cells(intR, 9).Value2
        dblSolution_40_Quantity = objSrc.Cells(intR, 10).Value2
        dblSolution_40_Volume = objSrc.Cells(intR, 11).Value2
        
        strDoseAdvice = objSrc.Cells(intR, 18).Value2
         
        dblGenericQuantity = 0
        dblGenericVolume = 0
        dblSolutionVolume = 0
        
        dblDripQuantity = 0
        
        strProduct = ""
        dblShelfLife = 0
        strShelfCondition = ""
        strPreparationText = ""
            
        objBuilder.Append "SET @department  = '" & strDepartment & "'" & vbNewLine
        objBuilder.Append "SET @generic  = '" & strGeneric & "'" & vbNewLine
        objBuilder.Append "SET @genericUnit  = '" & strGenericUnit & "'" & vbNewLine
        objBuilder.Append "SET @genericQuantity  =  " & DoubleToString(dblGenericQuantity) & vbNewLine
        objBuilder.Append "SET @genericVolume  =  " & DoubleToString(dblGenericVolume) & vbNewLine
        objBuilder.Append "SET @solutionVolume  =  " & DoubleToString(dblSolutionVolume) & vbNewLine
        objBuilder.Append "SET @solution_2_6_Quantity  =  " & DoubleToString(dblSolution_2_6_Quantity) & vbNewLine
        objBuilder.Append "SET @solution_2_6_Volume  =  " & DoubleToString(dblSolution_2_6_Volume) & vbNewLine
        objBuilder.Append "SET @solution_6_11_Quantity  =  " & DoubleToString(dblSolution_6_11_Quantity) & vbNewLine
        objBuilder.Append "SET @solution_6_11_Volume  =  " & DoubleToString(dblSolution_6_11_Volume) & vbNewLine
        objBuilder.Append "SET @solution_11_40_Quantity  =  " & DoubleToString(dblSolution_11_40_Quantity) & vbNewLine
        objBuilder.Append "SET @solution_11_40_Volume  =  " & DoubleToString(dblSolution_11_40_Volume) & vbNewLine
        objBuilder.Append "SET @solution_40_Quantity  =  " & DoubleToString(dblSolution_40_Quantity) & vbNewLine
        objBuilder.Append "SET @solution_40_Volume  =  " & DoubleToString(dblSolution_40_Volume) & vbNewLine
        objBuilder.Append "SET @minConcentration  = " & DoubleToString(dblMinConcentration) & vbNewLine
        objBuilder.Append "SET @maxConcentration  = " & DoubleToString(dblMaxConcentration) & vbNewLine
        objBuilder.Append "SET @solution  = '" & strSolution & "'" & vbNewLine
        objBuilder.Append "SET @solutionRequired  = 0" & vbNewLine
        objBuilder.Append "SET @dripQuantity  =  " & DoubleToString(dblDripQuantity) & vbNewLine
        objBuilder.Append "SET @doseUnit  = '" & strDoseUnit & "'" & vbNewLine
        objBuilder.Append "SET @minDose  =  " & DoubleToString(dblMinDose) & vbNewLine
        objBuilder.Append "SET @maxDose  =  " & DoubleToString(dblMaxDose) & vbNewLine
        objBuilder.Append "SET @absMaxDose  =  " & DoubleToString(dblAbsMaxDose) & vbNewLine
        objBuilder.Append "SET @doseAdvice  = '" & strDoseAdvice & "'" & vbNewLine
        objBuilder.Append "SET @product  =  '" & strProduct & "'" & vbNewLine
        objBuilder.Append "SET @shelfLife  =  " & DoubleToString(dblShelfLife) & vbNewLine
        objBuilder.Append "SET @shelfCondition  = '" & strShelfCondition & "'" & vbNewLine
        objBuilder.Append "SET @preparationText  =  '" & strPreparationText & "'" & vbNewLine
        objBuilder.Append "SET @signed = 1" & vbNewLine
        objBuilder.Append "SET @dilutionText  = '" & strDilutionText & "'" & vbNewLine
    
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "EXECUTE @RC = " & CONST_INSERT_CONFIG_MEDCONT & vbNewLine
        objBuilder.Append "   @versionID" & vbNewLine
        objBuilder.Append "  ,@versionUTC" & vbNewLine
        objBuilder.Append "  ,@versionDate" & vbNewLine
        objBuilder.Append "  ,@department" & vbNewLine
        objBuilder.Append "  ,@generic" & vbNewLine
        objBuilder.Append "  ,@genericUnit" & vbNewLine
        objBuilder.Append "  ,@genericQuantity" & vbNewLine
        objBuilder.Append "  ,@genericVolume" & vbNewLine
        objBuilder.Append "  ,@solutionVolume" & vbNewLine
        objBuilder.Append "  ,@solution_2_6_Quantity" & vbNewLine
        objBuilder.Append "  ,@solution_2_6_Volume" & vbNewLine
        objBuilder.Append "  ,@solution_6_11_Quantity" & vbNewLine
        objBuilder.Append "  ,@solution_6_11_Volume" & vbNewLine
        objBuilder.Append "  ,@solution_11_40_Quantity" & vbNewLine
        objBuilder.Append "  ,@solution_11_40_Volume" & vbNewLine
        objBuilder.Append "  ,@solution_40_Quantity" & vbNewLine
        objBuilder.Append "  ,@solution_40_Volume" & vbNewLine
        objBuilder.Append "  ,@minConcentration" & vbNewLine
        objBuilder.Append "  ,@maxConcentration" & vbNewLine
        objBuilder.Append "  ,@solution" & vbNewLine
        objBuilder.Append "  ,@solutionRequired" & vbNewLine
        objBuilder.Append "  ,@dripQuantity" & vbNewLine
        objBuilder.Append "  ,@doseUnit" & vbNewLine
        objBuilder.Append "  ,@minDose" & vbNewLine
        objBuilder.Append "  ,@maxDose" & vbNewLine
        objBuilder.Append "  ,@absMaxDose" & vbNewLine
        objBuilder.Append "  ,@doseAdvice" & vbNewLine
        objBuilder.Append "  ,@product" & vbNewLine
        objBuilder.Append "  ,@shelfLife" & vbNewLine
        objBuilder.Append "  ,@shelfCondition" & vbNewLine
        objBuilder.Append "  ,@preparationText" & vbNewLine
        objBuilder.Append "  ,@signed" & vbNewLine
        objBuilder.Append "  ,@dilutionText" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    
    Next
    
    objBuilder.Append vbNewLine
    objBuilder.Append Util_GetLogSQL("Save Pediatric Continuous Medication Configuration", False, , "ConfigMedCont")
    
    strSql = objBuilder.ToString()
    Util_GetSavePediatrieConfigMedContSql = strSql

End Function

Private Sub Test_Util_GetSavePediatrieConfigMedContSql()

    ModUtils.CopyToClipboard Util_GetSavePediatrieConfigMedContSql(False)

End Sub

Public Sub Database_SavePedConfigMedCont()

    Dim strSql As String

    On Error GoTo ErrorHandler
     
    ModProgress.StartProgress "PICU Continue Medicatie Configuratie Opslaan"
    
    strSql = Util_GetSavePediatrieConfigMedContSql(False)
    strSql = ModDatabase.Util_WrapTransaction(strSql, "insert_pedconfigmedcont")
    
    Util_InitConnection
    
    objConn.Open
    ModUtils.CopyToClipboard strSql
    objConn.Execute strSql
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:

    objConn.Close
    ModProgress.FinishProgress

    ModUtils.CopyToClipboard strSql
    ModMessage.ShowMsgBoxError "Kon de configuratie voor de pediatrie continue medicatie niet opslaan"
    ModLog.LogError Err, "Database_SavePedConfigMedCont with sql: " & vbNewLine & strSql
    
End Sub

Public Sub Database_LoadPedConfigMedCont()

    Dim strSql As String
    Dim objRs As Recordset
    Dim intC As Integer
    Dim intT As Integer
    Dim intR As Integer
    Dim objSrc As Range
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Configuratie voor PICU Continue Medicatie"
    
    Set objSrc = ModRange.GetRange("Tbl_Admin_PedMedCont")
    intT = objSrc.Rows.Count
    
    Util_InitConnection
    
    strSql = "SELECT * FROM " & CONST_GET_LATEST_CONFIG_MEDCONT & " ('" & CONST_DEP_PICU & "')"

    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    ImprovePerf True
    Application.Calculation = xlCalculationManual
    
    Do While Not objRs.EOF
        intR = intR + 1
        If intR > intT Then GoTo ErrorHandler
        
        objSrc.Cells(intR, 1).Value2 = objRs.Fields("Generic").Value
        objSrc.Cells(intR, 2).Value2 = objRs.Fields("GenericUnit").Value
        objSrc.Cells(intR, 3).Value2 = objRs.Fields("DoseUnit").Value
        objSrc.Cells(intR, 4).Value2 = objRs.Fields("Solution_2_6_Quantity").Value
        objSrc.Cells(intR, 5).Value2 = objRs.Fields("Solution_2_6_Volume").Value
        objSrc.Cells(intR, 6).Value2 = objRs.Fields("Solution_6_11_Quantity").Value
        objSrc.Cells(intR, 7).Value2 = objRs.Fields("Solution_6_11_Volume").Value
        objSrc.Cells(intR, 8).Value2 = objRs.Fields("Solution_11_40_Quantity").Value
        objSrc.Cells(intR, 9).Value2 = objRs.Fields("Solution_11_40_Volume").Value
        objSrc.Cells(intR, 10).Value2 = objRs.Fields("Solution_40_Quantity").Value
        objSrc.Cells(intR, 11).Value2 = objRs.Fields("Solution_40_Volume").Value
        objSrc.Cells(intR, 12).Value2 = objRs.Fields("MinDose").Value
        objSrc.Cells(intR, 13).Value2 = objRs.Fields("MaxDose").Value
        objSrc.Cells(intR, 14).Value2 = objRs.Fields("AbsMaxDose").Value
        objSrc.Cells(intR, 15).Value2 = objRs.Fields("MinConcentration").Value
        objSrc.Cells(intR, 16).Value2 = objRs.Fields("MaxConcentration").Value
        objSrc.Cells(intR, 17).Value2 = objRs.Fields("Solution").Value
        objSrc.Cells(intR, 18).Value2 = objRs.Fields("DoseAdvice").Value
                
        ModProgress.SetJobPercentage "Data laden", 32, intR
        objRs.MoveNext
    Loop
    
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ImprovePerf False
    
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ImprovePerf False

    ModProgress.FinishProgress
    objConn.Close

    ModUtils.CopyToClipboard strSql
    ModMessage.ShowMsgBoxError "Kon de configuratie voor de neonatologie continue medicatie niet laden"
    ModLog.LogError Err, "Database_LoadPedConfigMedCont with sql: " & vbNewLine & strSql

End Sub

Private Function Util_GetSaveConfigParentSql() As String

    Dim strSql As String
    Dim strLatest As String
    Dim objTable As Range
    Dim intC As Integer
    Dim intR As Integer
    
    Dim strName As String
    Dim dblEnergy As Double
    Dim dblProtein As Double
    Dim dblCarbohydrate As Double
    Dim dblLipid As Double
    Dim dblSodium As Double
    Dim dblPotassium As Double
    Dim dblCalcium As Double
    Dim dblPhosphor As Double
    Dim dblMagnesium As Double
    Dim dblIron As Double
    Dim dblVitD As Double
    Dim dblChloride As Double
    Dim strProduct As String
    Dim intSigned As Integer
    
    Dim objBuilder As ClassStringBuilder
    
    Set objBuilder = New ClassStringBuilder
    
    Set objTable = ModRange.GetRange("Tbl_Admin_ParEnt")
    intC = objTable.Rows.Count
    
    objBuilder.Append "DECLARE @RC int" & vbNewLine
    objBuilder.Append "DECLARE @versionID int" & vbNewLine
    objBuilder.Append "DECLARE @versionUTC datetime" & vbNewLine
    objBuilder.Append "DECLARE @versionDate datetime" & vbNewLine
    objBuilder.Append "DECLARE @name nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @energy float" & vbNewLine
    objBuilder.Append "DECLARE @protein float" & vbNewLine
    objBuilder.Append "DECLARE @carbohydrate float" & vbNewLine
    objBuilder.Append "DECLARE @lipid float" & vbNewLine
    objBuilder.Append "DECLARE @sodium float" & vbNewLine
    objBuilder.Append "DECLARE @potassium float" & vbNewLine
    objBuilder.Append "DECLARE @calcium float" & vbNewLine
    objBuilder.Append "DECLARE @phosphor float" & vbNewLine
    objBuilder.Append "DECLARE @magnesium float" & vbNewLine
    objBuilder.Append "DECLARE @iron float" & vbNewLine
    objBuilder.Append "DECLARE @vitD float" & vbNewLine
    objBuilder.Append "DECLARE @chloride float" & vbNewLine
    objBuilder.Append "DECLARE @product nvarchar(max)" & vbNewLine
    objBuilder.Append "DECLARE @signed bit" & vbNewLine
    objBuilder.Append "" & vbNewLine
    
    strLatest = strLatest & "SELECT @versionID = dbo.GetLatestConfigParEntVersion()"
    strLatest = strLatest & Util_GetVersionSQL(strLatest) & vbNewLine
    objBuilder.Append strLatest
    
    For intR = 1 To intC
    
        objBuilder.Append "SET @versionID  = @versionID" & vbNewLine
        objBuilder.Append "SET @versionUTC  = @versionUTC" & vbNewLine
        objBuilder.Append "SET @versionDate  = @versionDate" & vbNewLine
        
        strName = objTable.Cells(intR, 1).Value2
        dblEnergy = objTable.Cells(intR, 2).Value2
        dblProtein = objTable.Cells(intR, 3).Value2
        dblCarbohydrate = objTable.Cells(intR, 4).Value2
        dblLipid = objTable.Cells(intR, 5).Value2
        dblSodium = objTable.Cells(intR, 6).Value2
        dblPotassium = objTable.Cells(intR, 7).Value2
        dblCalcium = objTable.Cells(intR, 8).Value2
        dblPhosphor = objTable.Cells(intR, 9).Value2
        dblMagnesium = objTable.Cells(intR, 10).Value2
        dblIron = objTable.Cells(intR, 11).Value2
        dblVitD = objTable.Cells(intR, 12).Value2
        dblChloride = objTable.Cells(intR, 13).Value2
        strProduct = objTable.Cells(intR, 14).Value2
        
        objBuilder.Append "SET @name  = '" & strName & "'" & vbNewLine
        objBuilder.Append "SET @energy  = " & DoubleToString(dblEnergy) & vbNewLine
        objBuilder.Append "SET @protein  = " & DoubleToString(dblProtein) & vbNewLine
        objBuilder.Append "SET @carbohydrate  = " & DoubleToString(dblCarbohydrate) & vbNewLine
        objBuilder.Append "SET @lipid  = " & DoubleToString(dblLipid) & vbNewLine
        objBuilder.Append "SET @sodium  = " & DoubleToString(dblSodium) & vbNewLine
        objBuilder.Append "SET @potassium  = " & DoubleToString(dblPotassium) & vbNewLine
        objBuilder.Append "SET @calcium  = " & DoubleToString(dblCalcium) & vbNewLine
        objBuilder.Append "SET @phosphor  = " & DoubleToString(dblPhosphor) & vbNewLine
        objBuilder.Append "SET @magnesium  = " & DoubleToString(dblMagnesium) & vbNewLine
        objBuilder.Append "SET @iron  = " & DoubleToString(dblIron) & vbNewLine
        objBuilder.Append "SET @vitD  = " & DoubleToString(dblVitD) & vbNewLine
        objBuilder.Append "SET @chloride  = " & DoubleToString(dblChloride) & vbNewLine
        objBuilder.Append "SET @product  = '" & strProduct & "'" & vbNewLine
        objBuilder.Append "SET @signed = 1" & vbNewLine
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "EXECUTE @RC =  " & CONST_INSERT_CONFIG_PARENT & vbNewLine
        objBuilder.Append "   @versionID" & vbNewLine
        objBuilder.Append "  , @versionUTC" & vbNewLine
        objBuilder.Append "  , @versionDate" & vbNewLine
        objBuilder.Append "  ,@name" & vbNewLine
        objBuilder.Append "  ,@energy" & vbNewLine
        objBuilder.Append "  ,@protein" & vbNewLine
        objBuilder.Append "  ,@carbohydrate" & vbNewLine
        objBuilder.Append "  ,@lipid" & vbNewLine
        objBuilder.Append "  ,@sodium" & vbNewLine
        objBuilder.Append "  ,@potassium" & vbNewLine
        objBuilder.Append "  ,@calcium" & vbNewLine
        objBuilder.Append "  ,@phosphor" & vbNewLine
        objBuilder.Append "  ,@magnesium" & vbNewLine
        objBuilder.Append "  ,@iron" & vbNewLine
        objBuilder.Append "  ,@vitD" & vbNewLine
        objBuilder.Append "  ,@chloride" & vbNewLine
        objBuilder.Append "  ,@product" & vbNewLine
        objBuilder.Append "  ,@signed" & vbNewLine
        objBuilder.Append "" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    
    Next
    
    objBuilder.Append vbNewLine
    objBuilder.Append Util_GetLogSQL("Save configuration for parentaral fluids", False, , "ConfigParEnt")
 
    strSql = objBuilder.ToString()
    Util_GetSaveConfigParentSql = strSql
    
End Function

Private Sub Test_GetSaveConfigParEntSql()

    ModUtils.CopyToClipboard Util_GetSaveConfigParentSql()

End Sub

Public Sub Database_SaveConfigParEnt()

    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Configuratie voor parenteralia"
    
    strSql = Util_GetSaveConfigParentSql()
    strSql = Util_WrapTransaction(strSql, "insert_configparent")
    
    Util_InitConnection
    
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:
    
    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModProgress.FinishProgress
    ModLog.LogError Err, "Database_SaveConfigParEnt"
    

End Sub

Public Sub Database_LoadConfigParEnt()

    Dim strSql As String
    Dim objRs As Recordset
    Dim intC As Integer
    Dim intR As Integer
    Dim objTable As Range
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Parenteralia Configuratie"
    
    strSql = "SELECT * FROM " & CONST_GET_LATEST_CONFIG_PARENT & " ()"

    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Set objTable = ModRange.GetRange("Tbl_Admin_ParEnt")
    
    ImprovePerf True
    Application.Calculation = xlCalculationManual
    
    intC = objTable.Rows.Count
    Do While Not objRs.EOF
        intR = intR + 1
        If intR > intC Then GoTo ErrorHandler
        
        objTable.Cells(intR, 1).Value2 = objRs.Fields("Name").Value
        objTable.Cells(intR, 2).Value2 = objRs.Fields("Energy").Value
        objTable.Cells(intR, 3).Value2 = objRs.Fields("Protein").Value
        objTable.Cells(intR, 4).Value2 = objRs.Fields("Carbohydrate").Value
        objTable.Cells(intR, 5).Value2 = objRs.Fields("Lipid").Value
        objTable.Cells(intR, 6).Value2 = objRs.Fields("Sodium").Value
        objTable.Cells(intR, 7).Value2 = objRs.Fields("Potassium").Value
        objTable.Cells(intR, 8).Value2 = objRs.Fields("Calcium").Value
        objTable.Cells(intR, 9).Value2 = objRs.Fields("Phosphor").Value
        objTable.Cells(intR, 10).Value2 = objRs.Fields("Magnesium").Value
        objTable.Cells(intR, 11).Value2 = objRs.Fields("Iron").Value
        objTable.Cells(intR, 12).Value2 = objRs.Fields("VitD").Value
        objTable.Cells(intR, 13).Value2 = objRs.Fields("Chloride").Value
        objTable.Cells(intR, 14).Value2 = objRs.Fields("Product").Value
        
        ModProgress.SetJobPercentage "Laden", intC, intR
        objRs.MoveNext
    Loop
    
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ImprovePerf False
    
    objConn.Close
    ModProgress.FinishProgress

    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ImprovePerf False
    
    ModProgress.FinishProgress
    objConn.Close
    ModLog.LogError Err, "Database_LoadConfigParEnt"
End Sub

Private Sub Util_LoadVersions(colVersions As Collection, objRs As Recordset)

    Dim objVersion As ClassVersion

    Do While Not objRs.EOF
        Set objVersion = New ClassVersion
        
        objVersion.VersionID = objRs.Fields("VersionID").Value
        objVersion.VersionUTC = objRs.Fields("VersionUTC").Value
        objVersion.VersionDate = objRs.Fields("VersionDate").Value
        
        colVersions.Add objVersion
        objRs.MoveNext
    Loop

End Sub

Public Function Database_GetConfigParEntVersions() As Collection

    Dim colVersions As Collection
    Dim objRs As Recordset
    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    strSql = "SELECT * FROM " & CONST_GET_CONFIG_PARENT_VERSIONS & " ()" & vbNewLine
    strSql = strSql & "ORDER BY [VersionID] DESC "
    
    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    Set colVersions = New Collection
    
    Util_LoadVersions colVersions, objRs
    
    objConn.Close
    
    Set Database_GetConfigParEntVersions = colVersions
    
    Exit Function
    
ErrorHandler:

    ModLog.LogError Err, "Database_GetConfigParEntVersions"
    objConn.Close
    
End Function


Public Function Database_GetConfigParEnt(Optional ByVal intVersion As Integer = 0) As Collection

    Dim objCol As Collection
    Dim objParEnt As ClassParent
        
    Dim strSql As String
    Dim objRs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set objCol = New Collection
    
    If intVersion = 0 Then
        strSql = "SELECT * FROM  " & CONST_GET_LATEST_CONFIG_PARENT & "  ()"
    Else
        strSql = "SELECT * FROM " & CONST_GET_VERSION_CONFIG_PARENT & "(" & intVersion & ")"
    End If
    
    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        Set objParEnt = New ClassParent
    
        objParEnt.Name = objRs.Fields("Name").Value
        objParEnt.Energy = objRs.Fields("Energy").Value
        objParEnt.Eiwit = objRs.Fields("Protein").Value
        objParEnt.KH = objRs.Fields("Carbohydrate").Value
        objParEnt.Vet = objRs.Fields("Lipid").Value
        objParEnt.Na = objRs.Fields("Sodium").Value
        objParEnt.K = objRs.Fields("Potassium").Value
        objParEnt.Ca = objRs.Fields("Calcium").Value
        objParEnt.P = objRs.Fields("Phosphor").Value
        objParEnt.Mg = objRs.Fields("Magnesium").Value
        objParEnt.Fe = objRs.Fields("Iron").Value
        objParEnt.VitD = objRs.Fields("VitD").Value
        objParEnt.Cl = objRs.Fields("Chloride").Value
        objParEnt.Product = objRs.Fields("Product").Value
        
        objCol.Add objParEnt, objParEnt.Name
        
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    Set Database_GetConfigParEnt = objCol
    
    Exit Function
    
ErrorHandler:

    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Database_GetConfigParEnt"

End Function

Public Function Database_GetDataVersions(ByVal strHospNum As String) As Collection

    Dim colVersions As Collection
    Dim objVersion As ClassVersion
    Dim strSql As String
    Dim objRs As Recordset
    
    On Error GoTo ErrorHandler
    
    strSql = "SELECT * FROM " & CONST_GET_PRESCRIPTION_VERSIONS & " ('" & strHospNum & "')"
    strSql = strSql & "ORDER BY [VersionID] Desc"

    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    Set colVersions = New Collection
    
    Do While Not objRs.EOF
        Set objVersion = New ClassVersion
        
        objVersion.VersionID = objRs.Fields("VersionID").Value
        objVersion.VersionUTC = objRs.Fields("VersionUTC").Value
        objVersion.VersionDate = objRs.Fields("VersionDate").Value
        
        colVersions.Add objVersion
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    Set Database_GetDataVersions = colVersions

    Exit Function
    
ErrorHandler:

    ModUtils.CopyToClipboard strSql

    ModLog.LogError Err, "Database_GetDataVersions"
    objConn.Close
    
End Function

Private Sub Test_Database_GetDataVersions()

    Dim objVersion As ClassVersion
    Dim colVersions As Collection
    Dim strHospNum As String
    
    strHospNum = Patient_GetHospitalNumber()
    Set colVersions = Database_GetDataVersions(strHospNum)
    
    For Each objVersion In colVersions
        ModMessage.ShowMsgBoxInfo objVersion.VersionID & " : " & objVersion.VersionDate
    Next
    
End Sub

Public Sub Database_ClearDatabase(blnClearLog As Boolean)

    Dim strDatabase As String
    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    strDatabase = ModSetting.Setting_GetDatabase()
    
    If ModMessage.ShowMsgBoxYesNo("Database " & strDatabase & " leeg maken?") = vbYes Then
        If ModMessage.ShowMsgBoxYesNo("Weet u het zeker dat " & strDatabase & " leeggemaakt moet worden?") Then
            ModProgress.StartProgress "Clear Database"
                        
            strSql = "EXEC " & CONST_CLEARDATABASE & "   " & Util_WrapString(strDatabase) & ", " & IIf(blnClearLog, 1, 0)
            strSql = strSql & vbNewLine & Util_GetSaveConfigParentSql()
            strSql = strSql & vbNewLine & Util_GetSavePediatrieConfigMedContSql(True)
            strSql = strSql & vbNewLine & Util_GetSaveNeoConfigMedContSql(True)
            
            strSql = strSql & Util_GetLogSQL("Clear database", False)
            
            strSql = Util_WrapTransaction(strSql, "cleardatabase_trans")
            
            Util_InitConnection
            objConn.Open
            objConn.Execute strSql
            objConn.Close
            
            ModProgress.FinishProgress
        End If
    End If
    
    Exit Sub

ErrorHandler:

    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not clear database with SQL: " & vbNewLine & strSql
End Sub

Private Sub Test_Database_ClearDatabase()

    Database_ClearDatabase True

End Sub

Private Function Util_GetLogSQL(ByVal strText As String, ByVal blnDeclare As Boolean, Optional ByVal strHospNum As String = "", Optional ByVal strTable As String = "") As String

    Dim strSql As String
    Dim strUser As String
    Dim objBuilder As ClassStringBuilder
    
    Set objBuilder = New ClassStringBuilder
    
    If blnDeclare Then
        objBuilder.Append "DECLARE @versionID int" & vbNewLine
        objBuilder.Append "DECLARE @versionUTC datetime" & vbNewLine
        objBuilder.Append "DECLARE @versionDate datetime" & vbNewLine
        
        objBuilder.Append Util_GetVersionSQL("")
    
    End If
    
    strUser = ModMetaVision.MetaVision_GetUserLogin()
    strUser = Util_WrapString(strUser)
    strHospNum = Util_WrapString(strHospNum)
    strTable = Util_WrapString(strTable)
    strText = Util_WrapString(strText)
    
    objBuilder.Append vbNewLine
    objBuilder.Append "INSERT INTO [dbo].[Log]" & vbNewLine
    objBuilder.Append "( [Prescriber]" & vbNewLine
    objBuilder.Append ", [HospitalNumber]" & vbNewLine
    objBuilder.Append ", [VersionID]" & vbNewLine
    objBuilder.Append ", [VersionUTC]" & vbNewLine
    objBuilder.Append ", [VersionDate]" & vbNewLine
    objBuilder.Append ", [Table]" & vbNewLine
    objBuilder.Append ", [Text])" & vbNewLine
    objBuilder.Append "VALUES" & vbNewLine
    objBuilder.Append "( " & strUser & vbNewLine
    objBuilder.Append ", " & strHospNum & vbNewLine
    objBuilder.Append ", @VersionID" & vbNewLine
    objBuilder.Append ", @versionUTC" & vbNewLine
    objBuilder.Append ", @versionDate " & vbNewLine
    objBuilder.Append ", " & strTable & vbNewLine
    objBuilder.Append ", " & strText & ")" & vbNewLine
    
    strSql = objBuilder.ToString()
    Util_GetLogSQL = strSql

End Function

Private Sub Test_Util_GetLogSQL()

    ModUtils.CopyToClipboard Util_GetLogSQL("Testing", True, "1234", "Test Table")

End Sub

Public Sub Database_LogAction(ByVal strText As String, ByVal strPrescriber As String, ByVal strHospNum As String)

    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    If Not Setting_UseDatabase Then Exit Sub
        
    strSql = Util_GetLogSQL(strText, True, strHospNum, "")
    
    Util_InitConnection
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    Exit Sub
    
ErrorHandler:

    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not log action to database"
    objConn.Close

End Sub

Private Sub Test_Database_LogAction()

    Database_LogAction "Test", "Test User", "Test patient"

End Sub

Public Function Database_GetNeoConfigMedCont(Optional ByVal intVersion As Integer = 0) As Collection

    Dim strSql As String
    Dim objRs As Recordset
    Dim objCol As Collection
    Dim objConfig As ClassNeoMedCont
    
    On Error GoTo ErrorHandler
       
    Util_InitConnection
    
    If intVersion = 0 Then
        strSql = "SELECT * FROM " & CONST_GET_LATEST_CONFIG_MEDCONT & " ('" & CONST_DEP_NICU & "')"
    Else
        strSql = "SELECT * FROM " & CONST_GET_VERSION_CONFIG_MEDCONT & " ('" & CONST_DEP_NICU & "', " & intVersion & ")"
        ModUtils.CopyToClipboard strSql
    End If
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    Set objCol = New Collection
    
    Do While Not objRs.EOF
        Set objConfig = New ClassNeoMedCont
        
        objConfig.Generic = objRs.Fields("Generic").Value
        objConfig.GenericUnit = objRs.Fields("GenericUnit").Value
        objConfig.DoseUnit = objRs.Fields("DoseUnit").Value
        objConfig.GenericQuantity = objRs.Fields("GenericQuantity").Value
        objConfig.GenericVolume = objRs.Fields("GenericVolume").Value
        objConfig.MinDose = objRs.Fields("MinDose").Value
        objConfig.MaxDose = objRs.Fields("MaxDose").Value
        objConfig.AbsMaxDose = objRs.Fields("AbsMaxDose").Value
        objConfig.MinConcentration = objRs.Fields("MinConcentration").Value
        objConfig.MaxConcentration = objRs.Fields("MaxConcentration").Value
        objConfig.Solution = objRs.Fields("Solution").Value
        If objRs.Fields("SolutionRequired").Value Then
            objConfig.SolutionRequired = True
        Else
            objConfig.SolutionRequired = False
        End If
        objConfig.DoseAdvice = objRs.Fields("DoseAdvice").Value
        objConfig.SolutionVolume = objRs.Fields("SolutionVolume").Value
        objConfig.DripQuantity = objRs.Fields("DripQuantity").Value
        objConfig.Product = objRs.Fields("Product").Value
        objConfig.ShelfLife = objRs.Fields("ShelfLife").Value
        objConfig.ShelfCondition = objRs.Fields("ShelfCondition").Value
        objConfig.PreparationText = objRs.Fields("PreparationText").Value
        objConfig.DilutionText = objRs.Fields("DilutionText").Value
        
        objCol.Add objConfig
        
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    Set Database_GetNeoConfigMedCont = objCol
    
    Exit Function
    
ErrorHandler:

    objConn.Close

    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Database_LoadNeoConfigMedCont with sql: " & vbNewLine & strSql
    

End Function

Public Function Database_GetConfigMedContVersions(ByVal strDepartment As String) As Collection

    Dim colVersions As Collection
    Dim objRs As Recordset
    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    strSql = "SELECT * FROM " & CONST_GET_CONFIG_MEDCONT_VERSIONS & "  ('" & strDepartment & "')" & vbNewLine
    strSql = strSql & "ORDER BY [VersionID] DESC "
    
    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Set colVersions = New Collection
    Util_LoadVersions colVersions, objRs
    
    objConn.Close
    
    Set Database_GetConfigMedContVersions = colVersions
    
    Exit Function
    
ErrorHandler:

    ModLog.LogError Err, "Database_GetConfigMedContVersions"
    objConn.Close

End Function

Private Sub Test_Database_GetConfigMedContVersions()

    Dim colVersions As Collection
    Dim objVersion As ClassVersion
    
    Set colVersions = Database_GetConfigMedContVersions(CONST_DEP_NICU)
    
    For Each objVersion In colVersions
        ModMessage.ShowMsgBoxInfo objVersion.VersionID & " : " & objVersion.VersionDate
    Next
    

End Sub

Private Function Util_RemoveQuotes(ByVal strString As String) As String

    Util_RemoveQuotes = Replace(strString, "'", "")

End Function

Private Sub Test_Util_RemoveQuotes()

    ModMessage.ShowMsgBoxInfo Util_RemoveQuotes("Geen 'quotes'")

End Sub

Public Sub Database_SaveConfigMedDisc()
    
    Dim strSql As String
    Dim intVersion As Integer
    
    Dim objMedCol As Collection
    
    Dim objPICUSolCol As Collection
    Dim objNICUSolCol As Collection
    
    Dim objDoseCol As Collection
    
    On Error GoTo ErrorHandler
      
    Set objMedCol = Formularium_GetFormularium.GetMedicationCollection(False)
    Set objPICUSolCol = Formularium_GetSolutions(True, objMedCol)
    Set objNICUSolCol = Formularium_GetSolutions(False, objMedCol)
    
    Set objDoseCol = Formularium_GetDoses(objMedCol, False)
    
    ModProgress.StartProgress "Configuratie voor medicatie discontinue opslaan"
    
    strSql = Util_GetSaveConfigMedDiscSql(objMedCol, objPICUSolCol, objNICUSolCol, objDoseCol)
    strSql = Util_WrapTransaction(strSql, "insert_med_disc_config")
    
    ModUtils.CopyToClipboard strSql
    Util_InitConnection
    
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    ModProgress.FinishProgress
    
    intVersion = Util_GetLatestConfigMedDiscVersion
    ModMessage.ShowMsgBoxInfo "De discontinue medicatie is opgeslagen en de laatste versie is nu: " & intVersion
    
    Exit Sub
    
ErrorHandler:
    
    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModProgress.FinishProgress
    ModLog.LogError Err, "Database_SaveConfigMedDisc"
    

End Sub

Private Function Util_GetLatestConfigMedDiscVersion() As Integer

    Dim intVersion As Integer
    Dim objRs As Recordset
    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    strSql = "SELECT [dbo].[GetLatestConfigMedDiscVersion] ()"
    
    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        intVersion = objRs.Fields(0).Value
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    Util_GetLatestConfigMedDiscVersion = intVersion
    
    Exit Function
    
ErrorHandler:

    ModLog.LogError Err, "Util_GetLatestConfigMedDiscVersion"
    objConn.Close
End Function

Private Sub Test_Util_GetLatestConfigMedDiscVersion()

    ModMessage.ShowMsgBoxInfo Util_GetLatestConfigMedDiscVersion()

End Sub

Public Function Database_GetVersionIDFromString(ByVal strText As String) As Integer

    Database_GetVersionIDFromString = CInt(Split(strText, " : ")(0))

End Function

Public Function Database_GetLatestConfigMedContVersion(ByVal strDepartment As String) As Integer

    Dim colVersions As Collection
    Dim objVersion As ClassVersion
    Dim intVersion As Integer
    
    Set colVersions = Database_GetConfigMedContVersions(strDepartment)
    
    intVersion = 0
    For Each objVersion In colVersions
        intVersion = IIf(objVersion.VersionID > intVersion, objVersion.VersionID, intVersion)
    Next

    Database_GetLatestConfigMedContVersion = intVersion

End Function

Public Function Database_GetLatestConfigParentVersion() As Integer

    Dim colVersions As Collection
    Dim objVersion As ClassVersion
    Dim intVersion As Integer
    
    Set colVersions = Database_GetConfigParEntVersions()
    
    intVersion = 0
    For Each objVersion In colVersions
        intVersion = IIf(objVersion.VersionID > intVersion, objVersion.VersionID, intVersion)
    Next

    Database_GetLatestConfigParentVersion = intVersion

End Function

Private Sub Test_GetLatestConfigMedContVersion()

    ModMessage.ShowMsgBoxInfo Database_GetLatestConfigMedContVersion(CONST_DEP_PICU)

End Sub

Public Sub Database_LoadFormularium(objFormularium As ClassFormularium, ByVal blnShowProgress As Boolean)

    Dim strSql As String
    Dim objRs As Recordset
    Dim intC As Integer
    Dim objMed As ClassMedDisc
    Dim objPICUSol As ClassSolution
    Dim objNICUSol As ClassSolution
    Dim objDose As ClassDose
    Dim arrSubst() As String
    Dim intN As Integer
    Dim blnIsPICU As Boolean
    Dim blnMoved As Boolean
    
    On Error GoTo ErrorHandler
    
    If blnShowProgress Then ModProgress.StartProgress "Formularium"
    
    strSql = "SELECT * FROM [dbo].[GetConfigMedDiscLatest] () AS md" & vbNewLine
    strSql = strSql & "ORDER BY md.Generic, md.Shape, md.GenericQuantity"

    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    blnIsPICU = MetaVision_IsPICU()
    Do While Not objRs.EOF
        Set objMed = New ClassMedDisc
        
        With objMed
            
            .GPK = objRs.Fields("GPK").Value
            .MainGroup = objRs.Fields("MainGroup").Value
            .SubGroup = objRs.Fields("SubGroup").Value
            
            .ATC = objRs.Fields("ATC").Value
            If Not IsNull(objRs.Fields("TallMan").Value) Then .SetTallMan objRs.Fields("TallMan").Value
            .Generic = objRs.Fields("Generic").Value
            .Product = objRs.Fields("Product").Value
            .Shape = objRs.Fields("Shape").Value
            .GenericQuantity = objRs.Fields("GenericQuantity").Value
            .GenericUnit = objRs.Fields("GenericUnit").Value
            .Label = objRs.Fields("Label").Value
            .MultipleQuantity = objRs.Fields("MultipleQuantity").Value
            .MultipleUnit = objRs.Fields("MultipleUnit").Value
            
            .SetRouteList objRs.Fields("Routes").Value
            .SetIndicationList objRs.Fields("Indications").Value
            
            Set objPICUSol = New ClassSolution
            With objPICUSol
                .Generic = objMed.Generic
                .Shape = objMed.Shape
                If Not IsNull(objRs.Fields("PICUSolution")) Then .Solution = objRs.Fields("PICUSolution").Value
                If Not IsNull(objRs.Fields("PICUSolutionVolume")) Then .SolutionVolume = objRs.Fields("PICUSolutionVolume").Value
                If Not IsNull(objRs.Fields("PICUMinConc")) Then .MinConc = objRs.Fields("PICUMinConc").Value
                If Not IsNull(objRs.Fields("PICUMaxConc")) Then .MaxConc = objRs.Fields("PICUMaxConc").Value
                If Not IsNull(objRs.Fields("PICUMinInfusionTime")) Then .MinInfusionTime = objRs.Fields("PICUMinInfusionTime").Value
            End With
            
            Set objNICUSol = New ClassSolution
            With objNICUSol
                .Generic = objMed.Generic
                .Shape = objMed.Shape
                If Not IsNull(objRs.Fields("NICUSolution")) Then .Solution = objRs.Fields("NICUSolution").Value
                If Not IsNull(objRs.Fields("NICUSolutionVolume")) Then .SolutionVolume = objRs.Fields("NICUSolutionVolume").Value
                If Not IsNull(objRs.Fields("NICUMinConc")) Then .MinConc = objRs.Fields("NICUMinConc").Value
                If Not IsNull(objRs.Fields("NICUMaxConc")) Then .MaxConc = objRs.Fields("NICUMaxConc").Value
                If Not IsNull(objRs.Fields("NICUMinInfusionTime")) Then .MinInfusionTime = objRs.Fields("NICUMinInfusionTime").Value
            End With
                        
            If blnIsPICU Then
                .Solution = objPICUSol.Solution
                .SolutionVolume = objPICUSol.SolutionVolume
                .MaxConc = objPICUSol.MaxConc
                .MinInfusionTime = objPICUSol.MinInfusionTime
            Else
                .Solution = objNICUSol.Solution
                .SolutionVolume = objNICUSol.SolutionVolume
                .MaxConc = objNICUSol.MaxConc
                .MinInfusionTime = objNICUSol.MinInfusionTime
            End If
            
            blnMoved = False
            Do While objRs.Fields("GPK").Value = .GPK And Not objRs.EOF
            
                Set objDose = New ClassDose
                With objDose
                    .Generic = objMed.Generic
                    .Shape = objMed.Shape
                    If Not IsNull(objRs.Fields("Route")) Then .Route = objRs.Fields("Route").Value
                    If Not IsNull(objRs.Fields("Indication")) Then .Indication = objRs.Fields("Indication").Value
                    
                    If Not IsNull(objRs.Fields("Gender")) Then .Gender = objRs.Fields("Gender").Value
                    If Not IsNull(objRs.Fields("MinAge")) Then .MinAgeMo = objRs.Fields("MinAge").Value
                    If Not IsNull(objRs.Fields("MaxAge")) Then .MaxAgeMo = objRs.Fields("MaxAge").Value
                    If Not IsNull(objRs.Fields("MinWeight")) Then .MinWeightKg = objRs.Fields("MinWeight").Value
                    If Not IsNull(objRs.Fields("MaxWeight")) Then .MaxWeightKg = objRs.Fields("MaxWeight").Value
                    If Not IsNull(objRs.Fields("MinGestAge")) Then .MinGestDays = objRs.Fields("MinGestAge").Value
                    If Not IsNull(objRs.Fields("MaxGestAge")) Then .MaxGestDays = objRs.Fields("MaxGestAge").Value
                    
                    .Unit = objMed.MultipleUnit
                    If Not IsNull(objRs.Fields("Frequencies")) Then .Frequencies = objRs.Fields("Frequencies").Value
                    If Not IsNull(objRs.Fields("NormDose")) Then .NormDose = objRs.Fields("NormDose").Value
                    If Not IsNull(objRs.Fields("MinDose")) Then .MinDose = objRs.Fields("MinDose").Value
                    If Not IsNull(objRs.Fields("MaxDose")) Then .MaxDose = objRs.Fields("MaxDose").Value
                    If Not IsNull(objRs.Fields("AbsMaxDose")) Then .AbsMaxDose = objRs.Fields("AbsMaxDose").Value
                    If Not IsNull(objRs.Fields("MaxPerDose")) Then .MaxPerDose = objRs.Fields("MaxPerDose").Value
                    If Not IsNull(objRs.Fields("IsDosePerKg")) Then .IsDosePerKg = objRs.Fields("IsDosePerKg").Value
                    If Not IsNull(objRs.Fields("IsDosePerM2")) Then .IsDosePerM2 = objRs.Fields("IsDosePerM2").Value
                End With
                
                objMed.AddDose objDose
                
                objRs.MoveNext
                blnMoved = True
            Loop
            
                        
        End With
                
        arrSubst = Split(objMed.Generic, "+")
        For intN = 0 To UBound(arrSubst)
            objMed.AddSubstance arrSubst(intN), 0
        Next
        
        If objMed.Substances.Count = 1 Then
            objMed.Substances(1).Concentration = objMed.GenericQuantity
        End If
        
        objFormularium.AddMedication objMed
        
        intC = intC + 1
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", 1600, intC
        
        If Not blnMoved Then objRs.MoveNext
    Loop
    
    objConn.Close
    If blnShowProgress Then ModProgress.FinishProgress

    Exit Sub
    
ErrorHandler:

    If blnShowProgress Then ModProgress.FinishProgress
    objConn.Close
    ModLog.LogError Err, "Database_GetMedicamenten"


    Application.DisplayAlerts = True
    ImprovePerf False

End Sub

Private Sub Test_LoadFormularium()
    Dim objForm As ClassFormularium
    
    Set objForm = New ClassFormularium

    Database_LoadFormularium objForm, True

End Sub


Public Function Database_GetStandardPatients() As Collection

    Dim strSql As String
    Dim colPats As Collection
    Dim objPat As ClassPatientDetails
    Dim objRs As Recordset
    
    Set colPats = New Collection
    strSql = "SELECT * FROM dbo.GetStandardPatients()"
    
    Util_InitConnection
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        Set objPat = New ClassPatientDetails
        
        objPat.HospitalNumber = objRs.Fields("HospitalNumber").Value
        objPat.AchterNaam = objRs.Fields("LastName").Value
        objPat.VoorNaam = objRs.Fields("FirstName").Value
        
        colPats.Add objPat
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    Set Database_GetStandardPatients = colPats
    
End Function

Private Function Util_GetSaveConfigMedDiscSql(objMedCol As Collection, objPICUSolCol As Collection, objNICUSolCol As Collection, objDoseCol As Collection) As String

    Dim strSql As String
    Dim strLatest As String
    Dim intC As Integer
    Dim intR As Integer
    
    Dim strGPK As String
    Dim strATC As String
    Dim strMainGroup As String
    Dim strSubGroup As String
    Dim strGeneric As String
    Dim strProduct As String
    Dim strLabel As String
    Dim strShape As String
    Dim strRoutes As String
    Dim dblGenericQuantity As Double
    Dim strGenericUnit As String
    Dim dblMultipleQuantity As Double
    Dim strMultipleUnit As String
    Dim strIndications As String
    Dim intDoseAdjust As Integer
    
    Dim objMed As ClassMedDisc
    Dim objSol As ClassSolution
    Dim objDose As ClassDose
    
    Dim objBuilder As ClassStringBuilder

    Set objBuilder = New ClassStringBuilder
    intC = objMedCol.Count() + objPICUSolCol.Count() + objNICUSolCol.Count() + objDoseCol.Count()
    
    objBuilder.Append "DECLARE @RC int" & vbNewLine
    objBuilder.Append "DECLARE @versionID int" & vbNewLine
    objBuilder.Append "DECLARE @versionUTC datetime" & vbNewLine
    objBuilder.Append "DECLARE @versionDate datetime" & vbNewLine
    objBuilder.Append "DECLARE @GPK int" & vbNewLine
    objBuilder.Append "DECLARE @ATC nvarchar(10)" & vbNewLine
    objBuilder.Append "DECLARE @MainGroup nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @SubGroup nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @Generic nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @Product nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @Label nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @Shape nvarchar(150)" & vbNewLine
    objBuilder.Append "DECLARE @Routes nvarchar(300)" & vbNewLine
    objBuilder.Append "DECLARE @GenericQuantity float" & vbNewLine
    objBuilder.Append "DECLARE @GenericUnit nvarchar(50)" & vbNewLine
    objBuilder.Append "DECLARE @MultipleQuantity float" & vbNewLine
    objBuilder.Append "DECLARE @MultipleUnit nvarchar(50)" & vbNewLine
    objBuilder.Append "DECLARE @Indications nvarchar(max)" & vbNewLine
    objBuilder.Append "DECLARE @IsActive bit" & vbNewLine
    
    objBuilder.Append "DECLARE @Department nvarchar(60)" & vbNewLine
    objBuilder.Append "DECLARE @Solution nvarchar(150)" & vbNewLine
    objBuilder.Append "DECLARE @SolutionVolume float" & vbNewLine
    objBuilder.Append "DECLARE @MinConc float" & vbNewLine
    objBuilder.Append "DECLARE @MaxConc float" & vbNewLine
    objBuilder.Append "DECLARE @MinInfusionTime int" & vbNewLine
    
    objBuilder.Append "DECLARE @Route nvarchar(60)" & vbNewLine
    objBuilder.Append "DECLARE @Indication nvarchar(500)" & vbNewLine
    objBuilder.Append "DECLARE @Gender nvarchar(50)" & vbNewLine
    objBuilder.Append "DECLARE @MinAge float" & vbNewLine
    objBuilder.Append "DECLARE @MaxAge float" & vbNewLine
    objBuilder.Append "DECLARE @MinWeight float" & vbNewLine
    objBuilder.Append "DECLARE @MaxWeight float" & vbNewLine
    objBuilder.Append "DECLARE @MinGestAge float" & vbNewLine
    objBuilder.Append "DECLARE @MaxGestAge float" & vbNewLine
    objBuilder.Append "DECLARE @Frequencies nvarchar(500)" & vbNewLine
    objBuilder.Append "DECLARE @DoseUnit nvarchar(50)" & vbNewLine
    objBuilder.Append "DECLARE @NormDose float" & vbNewLine
    objBuilder.Append "DECLARE @MinDose float" & vbNewLine
    objBuilder.Append "DECLARE @MaxDose float" & vbNewLine
    objBuilder.Append "DECLARE @MaxPerDose float" & vbNewLine
    objBuilder.Append "DECLARE @StartDose float" & vbNewLine
    objBuilder.Append "DECLARE @IsDosePerKg bit" & vbNewLine
    objBuilder.Append "DECLARE @IsDosePerM2 bit" & vbNewLine
    objBuilder.Append "DECLARE @AbsMaxDose float" & vbNewLine
    
    objBuilder.Append "" & vbNewLine
    
    strLatest = "SELECT @versionID = dbo.GetLatestConfigMedDiscVersion()"
    strLatest = Util_GetVersionSQL(strLatest) & vbNewLine
    objBuilder.Append strLatest & vbNewLine
    
    For Each objMed In objMedCol
    
        intR = intR + 1
    
        objBuilder.Append "SET @versionID  = @versionID" & vbNewLine
        objBuilder.Append "SET @versionUTC  = @versionUTC" & vbNewLine
        objBuilder.Append "SET @versionDate  = @versionDate" & vbNewLine
        
        strGPK = objMed.GPK
        strATC = objMed.ATC
        strMainGroup = Util_RemoveQuotes(objMed.MainGroup)
        strSubGroup = Util_RemoveQuotes(objMed.SubGroup)
        strGeneric = Util_RemoveQuotes(objMed.Generic)
        strProduct = Util_RemoveQuotes(objMed.Product)
        strLabel = Util_RemoveQuotes(objMed.Label)
        strShape = objMed.Shape
        strRoutes = objMed.Routes
        dblGenericQuantity = objMed.GenericQuantity
        strGenericUnit = objMed.GenericUnit
        dblMultipleQuantity = objMed.MultipleQuantity
        strMultipleUnit = objMed.MultipleUnit
        strIndications = Util_RemoveQuotes(objMed.Indications)
        
        objBuilder.Append "SET @GPK  = " & strGPK & "" & vbNewLine
        objBuilder.Append "SET @ATC  = '" & strATC & "'" & vbNewLine
        objBuilder.Append "SET @MainGroup  = '" & strMainGroup & "'" & vbNewLine
        objBuilder.Append "SET @SubGroup  = '" & strSubGroup & "'" & vbNewLine
        objBuilder.Append "SET @Generic  = '" & strGeneric & "'" & vbNewLine
        objBuilder.Append "SET @Product  = '" & strProduct & "'" & vbNewLine
        objBuilder.Append "SET @Label  = '" & strLabel & "'" & vbNewLine
        objBuilder.Append "SET @Shape  = '" & strShape & "'" & vbNewLine
        objBuilder.Append "SET @Routes  = '" & strRoutes & "'" & vbNewLine
        objBuilder.Append "SET @GenericQuantity  = " & DoubleToString(dblGenericQuantity) & vbNewLine
        objBuilder.Append "SET @GenericUnit  = '" & strGenericUnit & "'" & vbNewLine
        objBuilder.Append "SET @MultipleQuantity  = " & DoubleToString(dblMultipleQuantity) & vbNewLine
        objBuilder.Append "SET @MultipleUnit  = '" & strMultipleUnit & "'" & vbNewLine
        objBuilder.Append "SET @Indications  = '" & strIndications & "'" & vbNewLine
        objBuilder.Append "SET @IsActive = 1" & vbNewLine
        objBuilder.Append "" & vbNewLine
        
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "EXECUTE @RC = [dbo].[InsertConfigMedDisc] " & vbNewLine
        objBuilder.Append "   @versionID" & vbNewLine
        objBuilder.Append "  ,@versionUTC" & vbNewLine
        objBuilder.Append "  ,@versionDate" & vbNewLine
        objBuilder.Append "  ,@GPK" & vbNewLine
        objBuilder.Append "  ,@ATC" & vbNewLine
        objBuilder.Append "  ,@MainGroup" & vbNewLine
        objBuilder.Append "  ,@SubGroup" & vbNewLine
        objBuilder.Append "  ,@Generic" & vbNewLine
        objBuilder.Append "  ,@Product" & vbNewLine
        objBuilder.Append "  ,@Label" & vbNewLine
        objBuilder.Append "  ,@Shape" & vbNewLine
        objBuilder.Append "  ,@Routes" & vbNewLine
        objBuilder.Append "  ,@GenericQuantity" & vbNewLine
        objBuilder.Append "  ,@GenericUnit" & vbNewLine
        objBuilder.Append "  ,@MultipleQuantity" & vbNewLine
        objBuilder.Append "  ,@MultipleUnit" & vbNewLine
        objBuilder.Append "  ,@Indications" & vbNewLine
        objBuilder.Append "  ,@IsActive" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    
    Next
    
    objBuilder.Append vbNewLine
    
    For Each objSol In objPICUSolCol
        intR = intR + 1
    
        objBuilder.Append "SET @versionID  = @versionID" & vbNewLine
        objBuilder.Append "SET @versionUTC  = @versionUTC" & vbNewLine
        objBuilder.Append "SET @versionDate  = @versionDate" & vbNewLine
        
        objBuilder.Append "SET @Department  = 'PICU'" & vbNewLine
        objBuilder.Append "SET @Generic  = '" & Util_RemoveQuotes(objSol.Generic) & "'" & vbNewLine
        objBuilder.Append "SET @Shape  = '" & Util_RemoveQuotes(objSol.Shape) & "'" & vbNewLine
        objBuilder.Append "SET @Solution  = '" & Util_RemoveQuotes(objSol.Solution) & "'" & vbNewLine
        objBuilder.Append "SET @SolutionVolume  = " & DoubleToString(objSol.SolutionVolume) & vbNewLine
        objBuilder.Append "SET @MinConc  = " & DoubleToString(objSol.MinConc) & vbNewLine
        objBuilder.Append "SET @MaxConc  = " & DoubleToString(objSol.MaxConc) & vbNewLine
        objBuilder.Append "SET @MinInfusionTime  = " & DoubleToString(objSol.MinInfusionTime) & vbNewLine
    
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "EXECUTE @RC = [dbo].[InsertConfigMedDiscSolution] " & vbNewLine
        objBuilder.Append "   @versionID" & vbNewLine
        objBuilder.Append "  ,@versionUTC" & vbNewLine
        objBuilder.Append "  ,@versionDate" & vbNewLine
        objBuilder.Append "  ,@Department" & vbNewLine
        objBuilder.Append "  ,@Generic" & vbNewLine
        objBuilder.Append "  ,@Shape" & vbNewLine
        objBuilder.Append "  ,@Solution" & vbNewLine
        objBuilder.Append "  ,@SolutionVolume" & vbNewLine
        objBuilder.Append "  ,@MinConc" & vbNewLine
        objBuilder.Append "  ,@MaxConc" & vbNewLine
        objBuilder.Append "  ,@MinInfusionTime" & vbNewLine
    
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    Next
    
    objBuilder.Append vbNewLine
        
    For Each objSol In objNICUSolCol
        intR = intR + 1
    
        objBuilder.Append "SET @versionID  = @versionID" & vbNewLine
        objBuilder.Append "SET @versionUTC  = @versionUTC" & vbNewLine
        objBuilder.Append "SET @versionDate  = @versionDate" & vbNewLine
        
        objBuilder.Append "SET @Department  = 'NICU'" & vbNewLine
        
        objBuilder.Append "SET @Generic  = '" & Util_RemoveQuotes(objSol.Generic) & "'" & vbNewLine
        objBuilder.Append "SET @Shape  = '" & Util_RemoveQuotes(objSol.Shape) & "'" & vbNewLine
        objBuilder.Append "SET @Solution  = '" & Util_RemoveQuotes(objSol.Solution) & "'" & vbNewLine
        objBuilder.Append "SET @SolutionVolume  = " & DoubleToString(objSol.SolutionVolume) & vbNewLine
        objBuilder.Append "SET @MinConc  = " & DoubleToString(objSol.MinConc) & vbNewLine
        objBuilder.Append "SET @MaxConc  = " & DoubleToString(objSol.MaxConc) & vbNewLine
        objBuilder.Append "SET @MinInfusionTime  = " & DoubleToString(objSol.MinInfusionTime) & vbNewLine
    
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "EXECUTE @RC = [dbo].[InsertConfigMedDiscSolution] " & vbNewLine
        objBuilder.Append "   @versionID" & vbNewLine
        objBuilder.Append "  ,@versionUTC" & vbNewLine
        objBuilder.Append "  ,@versionDate" & vbNewLine
        objBuilder.Append "  ,@Department" & vbNewLine
        objBuilder.Append "  ,@Generic" & vbNewLine
        objBuilder.Append "  ,@Shape" & vbNewLine
        objBuilder.Append "  ,@Solution" & vbNewLine
        objBuilder.Append "  ,@SolutionVolume" & vbNewLine
        objBuilder.Append "  ,@MinConc" & vbNewLine
        objBuilder.Append "  ,@MaxConc" & vbNewLine
        objBuilder.Append "  ,@MinInfusionTime" & vbNewLine
    
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    Next
    
    objBuilder.Append vbNewLine
    
    For Each objDose In objDoseCol
    
        intR = intR + 1
    
        objBuilder.Append "SET @versionID  = @versionID" & vbNewLine
        objBuilder.Append "SET @versionUTC  = @versionUTC" & vbNewLine
        objBuilder.Append "SET @versionDate  = @versionDate" & vbNewLine
        
        objBuilder.Append "SET @Department  = '" & Util_RemoveQuotes(objDose.Department) & "'" & vbNewLine
        objBuilder.Append "SET @Generic  = '" & Util_RemoveQuotes(objDose.Generic) & "'" & vbNewLine
        objBuilder.Append "SET @Shape  = '" & Util_RemoveQuotes(objDose.Shape) & "'" & vbNewLine
        objBuilder.Append "SET @Route  = '" & Util_RemoveQuotes(objDose.Route) & "'" & vbNewLine
        objBuilder.Append "SET @Indication  = '" & Util_RemoveQuotes(objDose.Indication) & "'" & vbNewLine
        
        objBuilder.Append "SET @Gender  = '" & Util_RemoveQuotes(objDose.Gender) & "'" & vbNewLine
        objBuilder.Append "SET @MinAge = " & DoubleToString(objDose.MinAgeMo) & vbNewLine
        objBuilder.Append "SET @MaxAge = " & DoubleToString(objDose.MaxAgeMo) & vbNewLine
        objBuilder.Append "SET @MinWeight = " & DoubleToString(objDose.MinWeightKg) & vbNewLine
        objBuilder.Append "SET @MaxWeight = " & DoubleToString(objDose.MaxWeightKg) & vbNewLine
        objBuilder.Append "SET @MinGestAge = " & objDose.MinGestDays & vbNewLine
        objBuilder.Append "SET @MaxGestAge = " & objDose.MaxGestDays & vbNewLine
        
        objBuilder.Append "SET @Frequencies = '" & Util_RemoveQuotes(objDose.Frequencies) & "'" & vbNewLine
        objBuilder.Append "SET @DoseUnit = '" & Util_RemoveQuotes(objDose.Unit) & "'" & vbNewLine
        
        objBuilder.Append "SET @NormDose = " & DoubleToString(objDose.NormDose) & vbNewLine
        objBuilder.Append "SET @MinDose = " & DoubleToString(objDose.MinDose) & vbNewLine
        objBuilder.Append "SET @MaxDose = " & DoubleToString(objDose.MaxDose) & vbNewLine
        objBuilder.Append "SET @AbsMaxDose = " & DoubleToString(objDose.AbsMaxDose) & vbNewLine
        
        If objDose.IsDosePerKg Then intDoseAdjust = 1 Else intDoseAdjust = 0
        objBuilder.Append "SET @IsDosePerKg = " & intDoseAdjust & vbNewLine
        If objDose.IsDosePerM2 Then intDoseAdjust = 1 Else intDoseAdjust = 0
        objBuilder.Append "SET @IsDosePerM2 = " & intDoseAdjust & vbNewLine
        
        objBuilder.Append "" & vbNewLine
        objBuilder.Append "EXECUTE @RC = [dbo].[InsertConfigMedDiscDose]" & vbNewLine
        objBuilder.Append "      @VersionID" & vbNewLine
        objBuilder.Append "     ,@VersionUTC" & vbNewLine
        objBuilder.Append "     ,@VersionDate" & vbNewLine
        objBuilder.Append "     ,@Department" & vbNewLine
        objBuilder.Append "     ,@Generic" & vbNewLine
        objBuilder.Append "     ,@Shape" & vbNewLine
        objBuilder.Append "     ,@Route" & vbNewLine
        objBuilder.Append "     ,@Indication" & vbNewLine
        objBuilder.Append "     ,@Gender" & vbNewLine
        objBuilder.Append "     ,@MinAge" & vbNewLine
        objBuilder.Append "     ,@MaxAge" & vbNewLine
        objBuilder.Append "     ,@MinWeight" & vbNewLine
        objBuilder.Append "     ,@MaxWeight" & vbNewLine
        objBuilder.Append "     ,@MinGestAge" & vbNewLine
        objBuilder.Append "     ,@MaxGestAge" & vbNewLine
        objBuilder.Append "     ,@Frequencies" & vbNewLine
        objBuilder.Append "     ,@DoseUnit" & vbNewLine
        objBuilder.Append "     ,@NormDose" & vbNewLine
        objBuilder.Append "     ,@MinDose" & vbNewLine
        objBuilder.Append "     ,@MaxDose" & vbNewLine
        objBuilder.Append "     ,@MaxPerDose" & vbNewLine
        objBuilder.Append "     ,@StartDose" & vbNewLine
        objBuilder.Append "     ,@IsDosePerKg" & vbNewLine
        objBuilder.Append "     ,@IsDosePerM2" & vbNewLine
        objBuilder.Append "     ,@AbsMaxDose" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    Next
    
    objBuilder.Append vbNewLine
    
    objBuilder.Append Util_GetLogSQL("Save configuration for discontinuous medication", False, , "ConfigMedDisc")
 
    Util_GetSaveConfigMedDiscSql = objBuilder.ToString()

End Function

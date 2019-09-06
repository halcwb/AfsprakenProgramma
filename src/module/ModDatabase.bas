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

Private Const constMedDiscTbl = "Table"

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
        
    On Error GoTo SaveDataError
    
    strSql = strSql & "DECLARE @RC int" & vbNewLine
    strSql = strSql & "DECLARE @versionID int" & vbNewLine
    strSql = strSql & "DECLARE @versionUTC datetime" & vbNewLine
    strSql = strSql & "DECLARE @versionDate datetime" & vbNewLine
    
    strLatest = "SELECT @versionID = " & CONST_GET_LATEST_PRESCRIPTION_VERSION & "('" & strHospNum & "')"
    strLatest = Util_GetVersionSQL(strLatest) & vbNewLine
    strSql = strSql & vbNewLine & strLatest
       
    intC = objData.Rows.Count
    For intN = 2 To intC
        strParam = objData.Cells(intN, 1).Value2
        varVal = objData.Cells(intN, 2).Value2
        varEmp = objData.Cells(intN, 3).Value2
        
        If Not varVal = varEmp Then
            strSql = strSql & vbNewLine & "EXEC " & CONST_INSERT_PRESCRIPTIONDATA & " '" & strHospNum & "', @versionID, @versionUTC, @versionDate, '" & strPrescriber & "', 0, ' " & strParam & " ', '" & varVal & " '"
        End If
        
        If blnProgress Then ModProgress.SetJobPercentage "Data wegschrijven", intC, intN
    Next intN
    
    strSql = strSql & Util_GetLogSQL("Save patient data", False, strHospNum, "PrescriptionData")
    strSql = strSql & vbNewLine
    strSql = strSql & vbNewLine
    
    intC = objText.Rows.Count
    For intN = 2 To intC
        If Not (Format(objText.Cells(intN, 2).Value2) = vbNullString Or Format(objText.Cells(intN, 2).Value2) = "0") Then
            strParam = objText.Cells(intN, 1).Value2
            varVal = objText.Cells(intN, 2).Value2
            strSql = strSql & vbNewLine & "EXEC " & CONST_INSERT_PRESCRIPTIONTEXT & " '" & strHospNum & "', @versionID, @versionUTC, @versionDate, '" & strPrescriber & "', 0, ' " & strParam & " ', '" & varVal & " '"
        End If
        
        If blnProgress Then ModProgress.SetJobPercentage "Text wegschrijven naar de database", intC, intN
    Next intN
    
    strSql = strSql & Util_GetLogSQL("Save patient data", False, strHospNum, "PrescriptionText")
    strSql = strSql & vbNewLine
    strSql = strSql & vbNewLine
    
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
    
    Application.ScreenUpdating = False
    
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
    Application.ScreenUpdating = True
    
    Exit Sub

Database_GetPatientDataError:
    
    ModMessage.ShowMsgBoxError "Kan patient met ziekenhuis nummer " & strHospNum & " niet laden."
    
    ModLog.LogError Err, "Could not get patient data with hospitalnumber " & strHospNum & " with SQL: " & vbNewLine & strSql
    objConn.Close
    Application.ScreenUpdating = True

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
    Dim strSql
        
    strTable = "Tbl_Admin_NeoMedCont"
    strDepartment = "Neonatologie"
    strDilutionText = ModRange.GetRangeValue("Var_Neo_MedCont_VerdunningTekst", vbNullString)
    
    Set objSrc = ModRange.GetRange(strTable)
    If Not blnIsBatch Then
    
        strSql = strSql & "DECLARE @RC int" & vbNewLine
        strSql = strSql & "DECLARE @versionID int" & vbNewLine
        strSql = strSql & "DECLARE @versionUTC datetime" & vbNewLine
        strSql = strSql & "DECLARE @versionDate datetime" & vbNewLine
        
        strSql = strSql & "DECLARE @department nvarchar(60)" & vbNewLine
        strSql = strSql & "DECLARE @generic nvarchar(300)" & vbNewLine
        strSql = strSql & "DECLARE @genericUnit nvarchar(50)" & vbNewLine
        strSql = strSql & "DECLARE @genericQuantity float" & vbNewLine
        strSql = strSql & "DECLARE @genericVolume float" & vbNewLine
        strSql = strSql & "DECLARE @solutionVolume float" & vbNewLine
        strSql = strSql & "DECLARE @solution_2_6_Quantity float" & vbNewLine
        strSql = strSql & "DECLARE @solution_2_6_Volume float" & vbNewLine
        strSql = strSql & "DECLARE @solution_6_11_Quantity float" & vbNewLine
        strSql = strSql & "DECLARE @solution_6_11_Volume float" & vbNewLine
        strSql = strSql & "DECLARE @solution_11_40_Quantity float" & vbNewLine
        strSql = strSql & "DECLARE @solution_11_40_Volume float" & vbNewLine
        strSql = strSql & "DECLARE @solution_40_Quantity float" & vbNewLine
        strSql = strSql & "DECLARE @solution_40_Volume float" & vbNewLine
        strSql = strSql & "DECLARE @minConcentration float" & vbNewLine
        strSql = strSql & "DECLARE @maxConcentration float" & vbNewLine
        strSql = strSql & "DECLARE @solution nvarchar(300)" & vbNewLine
        strSql = strSql & "DECLARE @solutionRequired bit" & vbNewLine
        strSql = strSql & "DECLARE @dripQuantity float" & vbNewLine
        strSql = strSql & "DECLARE @doseUnit nvarchar(50)" & vbNewLine
        strSql = strSql & "DECLARE @minDose float" & vbNewLine
        strSql = strSql & "DECLARE @maxDose float" & vbNewLine
        strSql = strSql & "DECLARE @absMaxDose float" & vbNewLine
        strSql = strSql & "DECLARE @doseAdvice nvarchar(max)" & vbNewLine
        strSql = strSql & "DECLARE @product nvarchar(max)" & vbNewLine
        strSql = strSql & "DECLARE @shelfLife float" & vbNewLine
        strSql = strSql & "DECLARE @shelfCondition nvarchar(50)" & vbNewLine
        strSql = strSql & "DECLARE @preparationText nvarchar(max)" & vbNewLine
        strSql = strSql & "DECLARE @signed bit" & vbNewLine
        strSql = strSql & "DECLARE @dilutionText nvarchar(max)" & vbNewLine
        strSql = strSql & "" & vbNewLine
    
    End If
        
    strLatest = "SET @department  = '" & strDepartment & "'" & vbNewLine
    strLatest = strLatest & "SELECT @versionID = dbo.GetLatestConfigMedContVersionForDepartment(@department)"
    
    strSql = strSql & Util_GetVersionSQL(strLatest)
        
        
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
            
        strSql = strSql & "SET @generic  = '" & strGeneric & "'" & vbNewLine
        strSql = strSql & "SET @genericUnit  = '" & strGenericUnit & "'" & vbNewLine
        strSql = strSql & "SET @genericQuantity  =  " & DoubleToString(dblGenericQuantity) & vbNewLine
        strSql = strSql & "SET @genericVolume  =  " & DoubleToString(dblGenericVolume) & vbNewLine
        strSql = strSql & "SET @solutionVolume  =  " & DoubleToString(dblSolutionVolume) & vbNewLine
        strSql = strSql & "SET @solution_2_6_Quantity  =  0" & vbNewLine
        strSql = strSql & "SET @solution_2_6_Volume  =  0" & vbNewLine
        strSql = strSql & "SET @solution_6_11_Quantity  =  0" & vbNewLine
        strSql = strSql & "SET @solution_6_11_Volume  =  0" & vbNewLine
        strSql = strSql & "SET @solution_11_40_Quantity  =  0" & vbNewLine
        strSql = strSql & "SET @solution_11_40_Volume  =  0" & vbNewLine
        strSql = strSql & "SET @solution_40_Quantity  =  0" & vbNewLine
        strSql = strSql & "SET @solution_40_Volume  =  0" & vbNewLine
        strSql = strSql & "SET @minConcentration  = " & DoubleToString(dblMinConcentration) & vbNewLine
        strSql = strSql & "SET @maxConcentration  = " & DoubleToString(dblMaxConcentration) & vbNewLine
        strSql = strSql & "SET @solution  = '" & strSolution & "'" & vbNewLine
        strSql = strSql & "SET @solutionRequired  = " & intSolutionRequired & vbNewLine
        strSql = strSql & "SET @dripQuantity  =  " & DoubleToString(dblDripQuantity) & vbNewLine
        strSql = strSql & "SET @doseUnit  = '" & strDoseUnit & "'" & vbNewLine
        strSql = strSql & "SET @minDose  =  " & DoubleToString(dblMinDose) & vbNewLine
        strSql = strSql & "SET @maxDose  =  " & DoubleToString(dblMaxDose) & vbNewLine
        strSql = strSql & "SET @absMaxDose  =  " & DoubleToString(dblAbsMaxDose) & vbNewLine
        strSql = strSql & "SET @doseAdvice  = '" & strDoseAdvice & "'" & vbNewLine
        strSql = strSql & "SET @product  =  '" & strProduct & "'" & vbNewLine
        strSql = strSql & "SET @shelfLife  =  " & DoubleToString(dblShelfLife) & vbNewLine
        strSql = strSql & "SET @shelfCondition  = '" & strShelfCondition & "'" & vbNewLine
        strSql = strSql & "SET @preparationText  =  '" & strPreparationText & "'" & vbNewLine
        strSql = strSql & "SET @signed = 1" & vbNewLine
        strSql = strSql & "SET @dilutionText  = '" & strDilutionText & "'" & vbNewLine
    
        strSql = strSql & "" & vbNewLine
        strSql = strSql & "" & vbNewLine
        strSql = strSql & "EXECUTE @RC = " & CONST_INSERT_CONFIG_MEDCONT & vbNewLine
        strSql = strSql & "   @versionID" & vbNewLine
        strSql = strSql & "  ,@versionUTC" & vbNewLine
        strSql = strSql & "  ,@versionDate" & vbNewLine
        strSql = strSql & "  ,@department" & vbNewLine
        strSql = strSql & "  ,@generic" & vbNewLine
        strSql = strSql & "  ,@genericUnit" & vbNewLine
        strSql = strSql & "  ,@genericQuantity" & vbNewLine
        strSql = strSql & "  ,@genericVolume" & vbNewLine
        strSql = strSql & "  ,@solutionVolume" & vbNewLine
        strSql = strSql & "  ,@solution_2_6_Quantity" & vbNewLine
        strSql = strSql & "  ,@solution_2_6_Volume" & vbNewLine
        strSql = strSql & "  ,@solution_6_11_Quantity" & vbNewLine
        strSql = strSql & "  ,@solution_6_11_Volume" & vbNewLine
        strSql = strSql & "  ,@solution_11_40_Quantity" & vbNewLine
        strSql = strSql & "  ,@solution_11_40_Volume" & vbNewLine
        strSql = strSql & "  ,@solution_40_Quantity" & vbNewLine
        strSql = strSql & "  ,@solution_40_Volume" & vbNewLine
        strSql = strSql & "  ,@minConcentration" & vbNewLine
        strSql = strSql & "  ,@maxConcentration" & vbNewLine
        strSql = strSql & "  ,@solution" & vbNewLine
        strSql = strSql & "  ,@solutionRequired" & vbNewLine
        strSql = strSql & "  ,@dripQuantity" & vbNewLine
        strSql = strSql & "  ,@doseUnit" & vbNewLine
        strSql = strSql & "  ,@minDose" & vbNewLine
        strSql = strSql & "  ,@maxDose" & vbNewLine
        strSql = strSql & "  ,@absMaxDose" & vbNewLine
        strSql = strSql & "  ,@doseAdvice" & vbNewLine
        strSql = strSql & "  ,@product" & vbNewLine
        strSql = strSql & "  ,@shelfLife" & vbNewLine
        strSql = strSql & "  ,@shelfCondition" & vbNewLine
        strSql = strSql & "  ,@preparationText" & vbNewLine
        strSql = strSql & "  ,@signed" & vbNewLine
        strSql = strSql & "  ,@dilutionText" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    
    Next
    
    strSql = strSql & vbNewLine
    strSql = strSql & Util_GetLogSQL("Save neonatal configuration for continuous medication", False, , "ConfigMedCont")

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
    
    ModProgress.StartProgress "Configuratie voor Neonatologie Continue Medicatie laden"
    
    Set objSrc = ModRange.GetRange("Tbl_Admin_NeoMedCont")
    intT = objSrc.Rows.Count
    
    Util_InitConnection
    
    strSql = "SELECT * FROM " & CONST_GET_LATEST_CONFIG_MEDCONT & " ('Neonatologie')"

    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Application.ScreenUpdating = False
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
    Application.ScreenUpdating = True
    
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True

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
    
    strTable = "Tbl_Admin_PedMedCont"
    strDepartment = "Pediatrie"
    strDilutionText = ""
    
    Set objSrc = ModRange.GetRange(strTable)
    If Not blnIsBatch Then strSql = strSql & "DECLARE @RC int" & vbNewLine
    If Not blnIsBatch Then strSql = strSql & "DECLARE @versionID int" & vbNewLine
    If Not blnIsBatch Then strSql = strSql & "DECLARE @versionUTC datetime" & vbNewLine
    If Not blnIsBatch Then strSql = strSql & "DECLARE @versionDate datetime" & vbNewLine
    strSql = strSql & "DECLARE @department nvarchar(60)" & vbNewLine
    strSql = strSql & "DECLARE @generic nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @genericUnit nvarchar(50)" & vbNewLine
    strSql = strSql & "DECLARE @genericQuantity float" & vbNewLine
    strSql = strSql & "DECLARE @genericVolume float" & vbNewLine
    strSql = strSql & "DECLARE @solutionVolume float" & vbNewLine
    strSql = strSql & "DECLARE @solution_2_6_Quantity float" & vbNewLine
    strSql = strSql & "DECLARE @solution_2_6_Volume float" & vbNewLine
    strSql = strSql & "DECLARE @solution_6_11_Quantity float" & vbNewLine
    strSql = strSql & "DECLARE @solution_6_11_Volume float" & vbNewLine
    strSql = strSql & "DECLARE @solution_11_40_Quantity float" & vbNewLine
    strSql = strSql & "DECLARE @solution_11_40_Volume float" & vbNewLine
    strSql = strSql & "DECLARE @solution_40_Quantity float" & vbNewLine
    strSql = strSql & "DECLARE @solution_40_Volume float" & vbNewLine
    strSql = strSql & "DECLARE @minConcentration float" & vbNewLine
    strSql = strSql & "DECLARE @maxConcentration float" & vbNewLine
    strSql = strSql & "DECLARE @solution nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @solutionRequired bit" & vbNewLine
    strSql = strSql & "DECLARE @dripQuantity float" & vbNewLine
    strSql = strSql & "DECLARE @doseUnit nvarchar(50)" & vbNewLine
    strSql = strSql & "DECLARE @minDose float" & vbNewLine
    strSql = strSql & "DECLARE @maxDose float" & vbNewLine
    strSql = strSql & "DECLARE @absMaxDose float" & vbNewLine
    strSql = strSql & "DECLARE @doseAdvice nvarchar(max)" & vbNewLine
    If Not blnIsBatch Then strSql = strSql & "DECLARE @product nvarchar(max)" & vbNewLine
    strSql = strSql & "DECLARE @shelfLife float" & vbNewLine
    strSql = strSql & "DECLARE @shelfCondition nvarchar(50)" & vbNewLine
    strSql = strSql & "DECLARE @preparationText nvarchar(max)" & vbNewLine
    If Not blnIsBatch Then strSql = strSql & "DECLARE @signed bit" & vbNewLine
    strSql = strSql & "DECLARE @dilutionText nvarchar(max)" & vbNewLine
    strSql = strSql & "" & vbNewLine
    
    strLatest = "SET @department  = '" & strDepartment & "'" & vbNewLine
    strLatest = strLatest & "SELECT @versionID = dbo.GetLatestConfigMedContVersionForDepartment(@department)"
    
    strSql = strSql & Util_GetVersionSQL(strLatest)
        
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
            
        strSql = strSql & "SET @department  = '" & strDepartment & "'" & vbNewLine
        strSql = strSql & "SET @generic  = '" & strGeneric & "'" & vbNewLine
        strSql = strSql & "SET @genericUnit  = '" & strGenericUnit & "'" & vbNewLine
        strSql = strSql & "SET @genericQuantity  =  " & DoubleToString(dblGenericQuantity) & vbNewLine
        strSql = strSql & "SET @genericVolume  =  " & DoubleToString(dblGenericVolume) & vbNewLine
        strSql = strSql & "SET @solutionVolume  =  " & DoubleToString(dblSolutionVolume) & vbNewLine
        strSql = strSql & "SET @solution_2_6_Quantity  =  " & DoubleToString(dblSolution_2_6_Quantity) & vbNewLine
        strSql = strSql & "SET @solution_2_6_Volume  =  " & DoubleToString(dblSolution_2_6_Volume) & vbNewLine
        strSql = strSql & "SET @solution_6_11_Quantity  =  " & DoubleToString(dblSolution_6_11_Quantity) & vbNewLine
        strSql = strSql & "SET @solution_6_11_Volume  =  " & DoubleToString(dblSolution_6_11_Volume) & vbNewLine
        strSql = strSql & "SET @solution_11_40_Quantity  =  " & DoubleToString(dblSolution_11_40_Quantity) & vbNewLine
        strSql = strSql & "SET @solution_11_40_Volume  =  " & DoubleToString(dblSolution_11_40_Volume) & vbNewLine
        strSql = strSql & "SET @solution_40_Quantity  =  " & DoubleToString(dblSolution_40_Quantity) & vbNewLine
        strSql = strSql & "SET @solution_40_Volume  =  " & DoubleToString(dblSolution_40_Volume) & vbNewLine
        strSql = strSql & "SET @minConcentration  = " & DoubleToString(dblMinConcentration) & vbNewLine
        strSql = strSql & "SET @maxConcentration  = " & DoubleToString(dblMaxConcentration) & vbNewLine
        strSql = strSql & "SET @solution  = '" & strSolution & "'" & vbNewLine
        strSql = strSql & "SET @solutionRequired  = 0" & vbNewLine
        strSql = strSql & "SET @dripQuantity  =  " & DoubleToString(dblDripQuantity) & vbNewLine
        strSql = strSql & "SET @doseUnit  = '" & strDoseUnit & "'" & vbNewLine
        strSql = strSql & "SET @minDose  =  " & DoubleToString(dblMinDose) & vbNewLine
        strSql = strSql & "SET @maxDose  =  " & DoubleToString(dblMaxDose) & vbNewLine
        strSql = strSql & "SET @absMaxDose  =  " & DoubleToString(dblAbsMaxDose) & vbNewLine
        strSql = strSql & "SET @doseAdvice  = '" & strDoseAdvice & "'" & vbNewLine
        strSql = strSql & "SET @product  =  '" & strProduct & "'" & vbNewLine
        strSql = strSql & "SET @shelfLife  =  " & DoubleToString(dblShelfLife) & vbNewLine
        strSql = strSql & "SET @shelfCondition  = '" & strShelfCondition & "'" & vbNewLine
        strSql = strSql & "SET @preparationText  =  '" & strPreparationText & "'" & vbNewLine
        strSql = strSql & "SET @signed = 1" & vbNewLine
        strSql = strSql & "SET @dilutionText  = '" & strDilutionText & "'" & vbNewLine
    
        strSql = strSql & "" & vbNewLine
        strSql = strSql & "" & vbNewLine
        strSql = strSql & "EXECUTE @RC = " & CONST_INSERT_CONFIG_MEDCONT & vbNewLine
        strSql = strSql & "   @versionID" & vbNewLine
        strSql = strSql & "  ,@versionUTC" & vbNewLine
        strSql = strSql & "  ,@versionDate" & vbNewLine
        strSql = strSql & "  ,@department" & vbNewLine
        strSql = strSql & "  ,@generic" & vbNewLine
        strSql = strSql & "  ,@genericUnit" & vbNewLine
        strSql = strSql & "  ,@genericQuantity" & vbNewLine
        strSql = strSql & "  ,@genericVolume" & vbNewLine
        strSql = strSql & "  ,@solutionVolume" & vbNewLine
        strSql = strSql & "  ,@solution_2_6_Quantity" & vbNewLine
        strSql = strSql & "  ,@solution_2_6_Volume" & vbNewLine
        strSql = strSql & "  ,@solution_6_11_Quantity" & vbNewLine
        strSql = strSql & "  ,@solution_6_11_Volume" & vbNewLine
        strSql = strSql & "  ,@solution_11_40_Quantity" & vbNewLine
        strSql = strSql & "  ,@solution_11_40_Volume" & vbNewLine
        strSql = strSql & "  ,@solution_40_Quantity" & vbNewLine
        strSql = strSql & "  ,@solution_40_Volume" & vbNewLine
        strSql = strSql & "  ,@minConcentration" & vbNewLine
        strSql = strSql & "  ,@maxConcentration" & vbNewLine
        strSql = strSql & "  ,@solution" & vbNewLine
        strSql = strSql & "  ,@solutionRequired" & vbNewLine
        strSql = strSql & "  ,@dripQuantity" & vbNewLine
        strSql = strSql & "  ,@doseUnit" & vbNewLine
        strSql = strSql & "  ,@minDose" & vbNewLine
        strSql = strSql & "  ,@maxDose" & vbNewLine
        strSql = strSql & "  ,@absMaxDose" & vbNewLine
        strSql = strSql & "  ,@doseAdvice" & vbNewLine
        strSql = strSql & "  ,@product" & vbNewLine
        strSql = strSql & "  ,@shelfLife" & vbNewLine
        strSql = strSql & "  ,@shelfCondition" & vbNewLine
        strSql = strSql & "  ,@preparationText" & vbNewLine
        strSql = strSql & "  ,@signed" & vbNewLine
        strSql = strSql & "  ,@dilutionText" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    
    Next
    
    strSql = strSql & vbNewLine
    strSql = strSql & Util_GetLogSQL("Save Pediatric Continuous Medication Configuration", False, , "ConfigMedCont")
    
    Util_GetSavePediatrieConfigMedContSql = strSql

End Function

Private Sub Test_Util_GetSavePediatrieConfigMedContSql()

    ModUtils.CopyToClipboard Util_GetSavePediatrieConfigMedContSql(False)

End Sub

Public Sub Database_SavePedConfigMedCont()

    Dim strSql As String

    On Error GoTo ErrorHandler
     
    ModProgress.StartProgress "Pediatrie Continue Medicatie Configuratie Opslaan"
    
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
    
    ModProgress.StartProgress "Configuratie voor Pediatrie Continue Medicatie"
    
    Set objSrc = ModRange.GetRange("Tbl_Admin_PedMedCont")
    intT = objSrc.Rows.Count
    
    Util_InitConnection
    
    strSql = "SELECT * FROM " & CONST_GET_LATEST_CONFIG_MEDCONT & " ('Pediatrie')"

    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Application.ScreenUpdating = False
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
    Application.ScreenUpdating = True
    
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True

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
    
    Set objTable = ModRange.GetRange("Tbl_Admin_ParEnt")
    intC = objTable.Rows.Count
    
    strSql = strSql & "DECLARE @RC int" & vbNewLine
    strSql = strSql & "DECLARE @versionID int" & vbNewLine
    strSql = strSql & "DECLARE @versionUTC datetime" & vbNewLine
    strSql = strSql & "DECLARE @versionDate datetime" & vbNewLine
    strSql = strSql & "DECLARE @name nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @energy float" & vbNewLine
    strSql = strSql & "DECLARE @protein float" & vbNewLine
    strSql = strSql & "DECLARE @carbohydrate float" & vbNewLine
    strSql = strSql & "DECLARE @lipid float" & vbNewLine
    strSql = strSql & "DECLARE @sodium float" & vbNewLine
    strSql = strSql & "DECLARE @potassium float" & vbNewLine
    strSql = strSql & "DECLARE @calcium float" & vbNewLine
    strSql = strSql & "DECLARE @phosphor float" & vbNewLine
    strSql = strSql & "DECLARE @magnesium float" & vbNewLine
    strSql = strSql & "DECLARE @iron float" & vbNewLine
    strSql = strSql & "DECLARE @vitD float" & vbNewLine
    strSql = strSql & "DECLARE @chloride float" & vbNewLine
    strSql = strSql & "DECLARE @product nvarchar(max)" & vbNewLine
    strSql = strSql & "DECLARE @signed bit" & vbNewLine
    strSql = strSql & "" & vbNewLine
    
    strLatest = strLatest & "SELECT @versionID = dbo.GetLatestConfigParEntVersion()"
    strLatest = strLatest & Util_GetVersionSQL(strLatest) & vbNewLine
    strSql = strSql & strLatest
    
    For intR = 1 To intC
    
        strSql = strSql & "SET @versionID  = @versionID" & vbNewLine
        strSql = strSql & "SET @versionUTC  = @versionUTC" & vbNewLine
        strSql = strSql & "SET @versionDate  = @versionDate" & vbNewLine
        
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
        
        strSql = strSql & "SET @name  = '" & strName & "'" & vbNewLine
        strSql = strSql & "SET @energy  = " & DoubleToString(dblEnergy) & vbNewLine
        strSql = strSql & "SET @protein  = " & DoubleToString(dblProtein) & vbNewLine
        strSql = strSql & "SET @carbohydrate  = " & DoubleToString(dblCarbohydrate) & vbNewLine
        strSql = strSql & "SET @lipid  = " & DoubleToString(dblLipid) & vbNewLine
        strSql = strSql & "SET @sodium  = " & DoubleToString(dblSodium) & vbNewLine
        strSql = strSql & "SET @potassium  = " & DoubleToString(dblPotassium) & vbNewLine
        strSql = strSql & "SET @calcium  = " & DoubleToString(dblCalcium) & vbNewLine
        strSql = strSql & "SET @phosphor  = " & DoubleToString(dblPhosphor) & vbNewLine
        strSql = strSql & "SET @magnesium  = " & DoubleToString(dblMagnesium) & vbNewLine
        strSql = strSql & "SET @iron  = " & DoubleToString(dblIron) & vbNewLine
        strSql = strSql & "SET @vitD  = " & DoubleToString(dblVitD) & vbNewLine
        strSql = strSql & "SET @chloride  = " & DoubleToString(dblChloride) & vbNewLine
        strSql = strSql & "SET @product  = '" & strProduct & "'" & vbNewLine
        strSql = strSql & "SET @signed = 1" & vbNewLine
        strSql = strSql & "" & vbNewLine
        strSql = strSql & "EXECUTE @RC =  " & CONST_INSERT_CONFIG_PARENT & vbNewLine
        strSql = strSql & "   @versionID" & vbNewLine
        strSql = strSql & "  , @versionUTC" & vbNewLine
        strSql = strSql & "  , @versionDate" & vbNewLine
        strSql = strSql & "  ,@name" & vbNewLine
        strSql = strSql & "  ,@energy" & vbNewLine
        strSql = strSql & "  ,@protein" & vbNewLine
        strSql = strSql & "  ,@carbohydrate" & vbNewLine
        strSql = strSql & "  ,@lipid" & vbNewLine
        strSql = strSql & "  ,@sodium" & vbNewLine
        strSql = strSql & "  ,@potassium" & vbNewLine
        strSql = strSql & "  ,@calcium" & vbNewLine
        strSql = strSql & "  ,@phosphor" & vbNewLine
        strSql = strSql & "  ,@magnesium" & vbNewLine
        strSql = strSql & "  ,@iron" & vbNewLine
        strSql = strSql & "  ,@vitD" & vbNewLine
        strSql = strSql & "  ,@chloride" & vbNewLine
        strSql = strSql & "  ,@product" & vbNewLine
        strSql = strSql & "  ,@signed" & vbNewLine
        strSql = strSql & "" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    
    Next
    
    strSql = strSql & vbNewLine
    strSql = strSql & Util_GetLogSQL("Save configuration for parentaral fluids", False, , "ConfigParEnt")
 
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
    
    Application.ScreenUpdating = False
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
    Application.ScreenUpdating = True
    
    objConn.Close
    ModProgress.FinishProgress

    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    
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
    
    If blnDeclare Then
        strSql = strSql & "DECLARE @versionID int" & vbNewLine
        strSql = strSql & "DECLARE @versionUTC datetime" & vbNewLine
        strSql = strSql & "DECLARE @versionDate datetime" & vbNewLine
        
        strSql = strSql & Util_GetVersionSQL("")
    
    End If
    
    strUser = ModMetaVision.MetaVision_GetUserLogin()
    strUser = Util_WrapString(strUser)
    strHospNum = Util_WrapString(strHospNum)
    strTable = Util_WrapString(strTable)
    strText = Util_WrapString(strText)
    
    strSql = strSql & vbNewLine
    strSql = strSql & "INSERT INTO [dbo].[Log]" & vbNewLine
    strSql = strSql & "( [Prescriber]" & vbNewLine
    strSql = strSql & ", [HospitalNumber]" & vbNewLine
    strSql = strSql & ", [VersionID]" & vbNewLine
    strSql = strSql & ", [VersionUTC]" & vbNewLine
    strSql = strSql & ", [VersionDate]" & vbNewLine
    strSql = strSql & ", [Table]" & vbNewLine
    strSql = strSql & ", [Text])" & vbNewLine
    strSql = strSql & "VALUES" & vbNewLine
    strSql = strSql & "( " & strUser & vbNewLine
    strSql = strSql & ", " & strHospNum & vbNewLine
    strSql = strSql & ", @VersionID" & vbNewLine
    strSql = strSql & ", @versionUTC" & vbNewLine
    strSql = strSql & ", @versionDate " & vbNewLine
    strSql = strSql & ", " & strTable & vbNewLine
    strSql = strSql & ", " & strText & ")" & vbNewLine
    
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
        strSql = "SELECT * FROM " & CONST_GET_LATEST_CONFIG_MEDCONT & " ('Neonatologie')"
    Else
        strSql = "SELECT * FROM " & CONST_GET_VERSION_CONFIG_MEDCONT & " ('Neonatologie', " & intVersion & ")"
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
    
    Set colVersions = Database_GetConfigMedContVersions("Neonatologie")
    
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

Private Function Util_GetSaveConfigMedDiscSql(objTable As Range) As String

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
    
    intC = objTable.Rows.Count
    
    strSql = strSql & "DECLARE @RC int" & vbNewLine
    strSql = strSql & "DECLARE @versionID int" & vbNewLine
    strSql = strSql & "DECLARE @versionUTC datetime" & vbNewLine
    strSql = strSql & "DECLARE @versionDate datetime" & vbNewLine
    strSql = strSql & "DECLARE @GPK int" & vbNewLine
    strSql = strSql & "DECLARE @ATC nvarchar(10)" & vbNewLine
    strSql = strSql & "DECLARE @MainGroup nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @SubGroup nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @Generic nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @Product nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @Label nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @Shape nvarchar(150)" & vbNewLine
    strSql = strSql & "DECLARE @Routes nvarchar(300)" & vbNewLine
    strSql = strSql & "DECLARE @GenericQuantity float" & vbNewLine
    strSql = strSql & "DECLARE @GenericUnit nvarchar(50)" & vbNewLine
    strSql = strSql & "DECLARE @MultipleQuantity float" & vbNewLine
    strSql = strSql & "DECLARE @MultipleUnit nvarchar(50)" & vbNewLine
    strSql = strSql & "DECLARE @Indications nvarchar(max)" & vbNewLine
    strSql = strSql & "DECLARE @IsActive bit" & vbNewLine
    strSql = strSql & "" & vbNewLine
    
    strLatest = strLatest & "SELECT @versionID = dbo.Util_GetLatestConfigMedDiscVersion()"
    strLatest = strLatest & Util_GetVersionSQL(strLatest) & vbNewLine
    strSql = strSql & strLatest
    
    For intR = 3 To intC
    
        strSql = strSql & "SET @versionID  = @versionID" & vbNewLine
        strSql = strSql & "SET @versionUTC  = @versionUTC" & vbNewLine
        strSql = strSql & "SET @versionDate  = @versionDate" & vbNewLine
        
        strGPK = objTable.Cells(intR, 1).Value2
        strATC = objTable.Cells(intR, 2).Value2
        strMainGroup = Util_RemoveQuotes(objTable.Cells(intR, 3).Value2)
        strSubGroup = Util_RemoveQuotes(objTable.Cells(intR, 4).Value2)
        strGeneric = Util_RemoveQuotes(objTable.Cells(intR, 5).Value2)
        strProduct = Util_RemoveQuotes(objTable.Cells(intR, 6).Value2)
        strLabel = Util_RemoveQuotes(objTable.Cells(intR, 7).Value2)
        strShape = objTable.Cells(intR, 8).Value2
        strRoutes = objTable.Cells(intR, 9).Value2
        dblGenericQuantity = objTable.Cells(intR, 10).Value2
        strGenericUnit = objTable.Cells(intR, 11).Value2
        dblMultipleQuantity = objTable.Cells(intR, 12).Value2
        strMultipleUnit = objTable.Cells(intR, 13).Value2
        strIndications = Util_RemoveQuotes(objTable.Cells(intR, 14).Value2)
        
        strSql = strSql & "SET @GPK  = " & strGPK & "" & vbNewLine
        strSql = strSql & "SET @ATC  = '" & strATC & "'" & vbNewLine
        strSql = strSql & "SET @MainGroup  = '" & strMainGroup & "'" & vbNewLine
        strSql = strSql & "SET @SubGroup  = '" & strSubGroup & "'" & vbNewLine
        strSql = strSql & "SET @Generic  = '" & strGeneric & "'" & vbNewLine
        strSql = strSql & "SET @Product  = '" & strProduct & "'" & vbNewLine
        strSql = strSql & "SET @Label  = '" & strLabel & "'" & vbNewLine
        strSql = strSql & "SET @Shape  = '" & strShape & "'" & vbNewLine
        strSql = strSql & "SET @Routes  = '" & strRoutes & "'" & vbNewLine
        strSql = strSql & "SET @GenericQuantity  = " & DoubleToString(dblGenericQuantity) & vbNewLine
        strSql = strSql & "SET @GenericUnit  = '" & strGenericUnit & "'" & vbNewLine
        strSql = strSql & "SET @MultipleQuantity  = " & DoubleToString(dblMultipleQuantity) & vbNewLine
        strSql = strSql & "SET @MultipleUnit  = '" & strMultipleUnit & "'" & vbNewLine
        strSql = strSql & "SET @Indications  = '" & strIndications & "'" & vbNewLine
        strSql = strSql & "SET @IsActive = 1" & vbNewLine
        strSql = strSql & "" & vbNewLine
        
        strSql = strSql & "" & vbNewLine
        strSql = strSql & "EXECUTE @RC = [dbo].[InsertConfigMedDisc] " & vbNewLine
        strSql = strSql & "   @versionID" & vbNewLine
        strSql = strSql & "  ,@versionUTC" & vbNewLine
        strSql = strSql & "  ,@versionDate" & vbNewLine
        strSql = strSql & "  ,@GPK" & vbNewLine
        strSql = strSql & "  ,@ATC" & vbNewLine
        strSql = strSql & "  ,@MainGroup" & vbNewLine
        strSql = strSql & "  ,@SubGroup" & vbNewLine
        strSql = strSql & "  ,@Generic" & vbNewLine
        strSql = strSql & "  ,@Product" & vbNewLine
        strSql = strSql & "  ,@Label" & vbNewLine
        strSql = strSql & "  ,@Shape" & vbNewLine
        strSql = strSql & "  ,@Routes" & vbNewLine
        strSql = strSql & "  ,@GenericQuantity" & vbNewLine
        strSql = strSql & "  ,@GenericUnit" & vbNewLine
        strSql = strSql & "  ,@MultipleQuantity" & vbNewLine
        strSql = strSql & "  ,@MultipleUnit" & vbNewLine
        strSql = strSql & "  ,@Indications" & vbNewLine
        strSql = strSql & "  ,@IsActive" & vbNewLine
        
        ModProgress.SetJobPercentage "Opslaan", intC, intR
    
    Next
    
    strSql = strSql & vbNewLine
    strSql = strSql & Util_GetLogSQL("Save configuration for discontinuous medication", False, , "ConfigMedDisc")
 
    Util_GetSaveConfigMedDiscSql = strSql


End Function

Private Sub Util_SaveConfigMedDisc(objSrc As Range)
    
    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Configuratie voor medicatie discontinue opslaan"
    
    strSql = Util_GetSaveConfigMedDiscSql(objSrc)
    strSql = Util_WrapTransaction(strSql, "insert_med_disc_config")
    
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
    ModLog.LogError Err, "Util_SaveConfigMedDisc"
    

End Sub

Private Function Util_GetLatestConfigMedDiscVersion() As Integer

    Dim intVersion As Integer
    Dim objRs As Recordset
    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    strSql = "SELECT [dbo].[Util_GetLatestConfigMedDiscVersion] ()"
    
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

Public Sub Database_ImportConfigMedDisc()

    Dim objConfigWbk As Workbook
    Dim objSrc As Range
    Dim lngErr As Long
    Dim strFile As String
    Dim intVersion As Integer
    Dim strMsg As String
    Dim vbAnswer
        
    Dim objMed As ClassNeoMedCont
    
    On Error GoTo ErrorHandler
    
    strMsg = "Kies een Excel bestand met de configuratie voor discontinue medicatie"
    ModMessage.ShowMsgBoxInfo strMsg
    
    strFile = ModFile.GetFileWithDialog
        
    strMsg = "Dit bestand importeren?" & vbNewLine & strFile
    If ModMessage.ShowMsgBoxYesNo(strMsg) = vbNo Then Exit Sub
       
    Application.DisplayAlerts = False
        
    Set objConfigWbk = Workbooks.Open(strFile, True, True)
    Set objSrc = objConfigWbk.Sheets(constMedDiscTbl).Range("A2").CurrentRegion()
    
    If objSrc.Columns.Count < 14 Then
        objConfigWbk.Close
        Application.DisplayAlerts = True
        
        strMsg = "Kan dit bestand niet importeren, er moeten minstens 14 kolommen zijn." & vbNewLine
        strMsg = "Dit bestand bevat slechts " & objSrc.Columns.Count & " kolommen."
        ModMessage.ShowMsgBoxExclam strMsg
        
        Exit Sub
    End If
        
    Util_SaveConfigMedDisc objSrc
    
    objConfigWbk.Close
    Application.DisplayAlerts = True
    
    intVersion = Util_GetLatestConfigMedDiscVersion()
    strMsg = "De meest recente versie van de configuratie van disccontinue medicatie is nu: " & intVersion
    ModMessage.ShowMsgBoxInfo strMsg
    
    Exit Sub
        
ErrorHandler:

    objConfigWbk.Close
    Application.DisplayAlerts = True

    ModLog.LogError Err, "Database_ImportConfigMedDisc"

End Sub

Private Sub Test_Database_ImportConfigMedDisc()

    Database_ImportConfigMedDisc

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

Private Sub Test_GetLatestConfigMedContVersion()

    ModMessage.ShowMsgBoxInfo Database_GetLatestConfigMedContVersion("Pediatrie")

End Sub

Public Sub Database_GetMedicamenten(objFormularium As ClassFormularium, ByVal blnShowProgress As Boolean)


    Dim strSql As String
    Dim objRs As Recordset
    Dim intC As Integer
    Dim objMed As ClassMedicatieDisc
    Dim arrSubst() As String
    Dim intN As Integer
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Formularium"
    
    strSql = "SELECT * FROM [dbo].[GetConfigMedDiscLatest] () AS md" & vbNewLine
    strSql = strSql & "ORDER BY md.Generic, md.Shape, md.GenericQuantity"

    Util_InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        Set objMed = New ClassMedicatieDisc
        
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
        
        End With
                
        arrSubst = Split(objMed.Generic, "+")
        For intN = 0 To UBound(arrSubst)
            objMed.AddSubstance arrSubst(intN), 0
        Next
        
        If objMed.Substances.Count = 1 Then
            objMed.Substances(1).Concentration = objMed.GenericQuantity
        End If
        
        objFormularium.AddMedicament objMed
        
        intC = intC + 1
        If blnShowProgress Then ModProgress.SetJobPercentage "Formularium laden", 1600, intC
        
        objRs.MoveNext
    Loop
    
    objConn.Close
    ModProgress.FinishProgress

    Exit Sub
    
ErrorHandler:

    ModProgress.FinishProgress
    objConn.Close
    ModLog.LogError Err, "Database_GetMedicamenten"


    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

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

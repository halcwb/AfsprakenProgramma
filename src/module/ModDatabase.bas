Attribute VB_Name = "ModDatabase"
Option Explicit

Private objConn As ADODB.Connection

Private Const constSecret As String = "secret"

Private Const constVersie As String = "Var_Glob_Versie"

Private Const constPatNum As String = "__0_PatNum"
Private Const constAN As String = "__2_AchterNaam"
Private Const constVN As String = "__3_VoorNaam"
Private Const constGebDatum As String = "__4_GebDatum"
Private Const constOpnDat As String = "_Pat_OpnDatum"
Private Const constGewicht As String = "_Pat_Gewicht"
Private Const constLengte As String = "_Pat_Lengte"
Private Const constGeslacht As String = "_Pat_Geslacht"
Private Const constDagen As String = "_Pat_GestDagen"
Private Const constWeken As String = "_Pat_GestWeken"
Private Const constGebGew As String = "_Pat_GebGew"

Private objDatabase As ClassDatabase

Public Sub InitDatabase()

    If objDatabase Is Nothing Then
        Set objDatabase = New ClassDatabase
        objDatabase.InitConnection ModSetting.Setting_GetServer(), ModSetting.Setting_GetDatabase()
    End If

End Sub

Private Sub Test_InitDatabase()

    InitDatabase

End Sub

Private Sub InitConnection()

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
    ModLog.LogError Err, "InitConnection Failed"

End Sub

Private Sub Test_InitConnectionWithAPDB()

    InitConnection

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

    InitConnection
    
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

Private Function PatientExists(strHospN As String) As Boolean

    Dim strSql As String
    
    strSql = "SELECT * FROM dbo.GetPatients('" & strHospN & "')"
    
    InitConnection
    
    objConn.Open

    PatientExists = Not objConn.Execute(strSql).EOF

End Function

Private Sub Test_PatientExists()

    MsgBox PatientExists("000")

End Sub

Private Function PrescriberExists(strUser As String) As Boolean

    Dim strSql As String
    
    strSql = "SELECT * FROM dbo.GetPrescribers('" & strUser & "')"
    
    InitConnection
    
    objConn.Open

    PrescriberExists = Not objConn.Execute(strSql).EOF

End Function

Private Sub Test_PrescriberExists()

    MsgBox PrescriberExists("000")

End Sub

Private Function WrapString(varItem As Variant) As Variant

    WrapString = "'" & varItem & "'"

End Function

Private Function WrapDateTime(strDateTime As String) As String

    WrapDateTime = "{ts'" & strDateTime & "'}"

End Function

Public Function WrapTransaction(ByVal strSql As String, ByVal strName As String) As String

    Dim strTrans As String
    
    strTrans = "BEGIN TRANSACTION [" & strName & "]" & vbNewLine
    strTrans = strTrans & "BEGIN TRY" & vbNewLine
    strTrans = strTrans & strSql & vbNewLine
    strTrans = strTrans & "COMMIT TRANSACTION [" & strName & "]" & vbNewLine
    strTrans = strTrans & "END TRY" & vbNewLine
    strTrans = strTrans & "BEGIN CATCH" & vbNewLine
    strTrans = strTrans & "ROLLBACK TRANSACTION [" & strName & "]" & vbNewLine
    strTrans = strTrans & "END CATCH"
    
    WrapTransaction = strTrans

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
    
    strHN = WrapString(ModPatient.Patient_GetHospitalNumber)
    strBD = WrapString(ModDate.FormatDateYearMonthDay(ModPatient.Patient_BirthDate))
    strAN = WrapString(ModPatient.Patient_GetLastName)
    strVN = WrapString(ModPatient.Patient_GetFirstName)
    strGN = WrapString(ModRange.GetRangeValue(constGeslacht, Null))
    intGW = ModRange.GetRangeValue(constWeken, Null)
    intGD = ModRange.GetRangeValue(constDagen, Null)
    dblBW = ModRange.GetRangeValue(constGebGew, Null)
        
    arrSql = Array(strHN, strBD, strAN, strVN, strGN, intGW, intGD, dblBW)
        
    InitConnection
    
    objConn.Open
    
    If PatientExists(ModPatient.Patient_GetHospitalNumber()) Then
        strSql = "EXEC UpdatePatient "
    Else
        strSql = "EXEC InsertPatient "
    End If
    
    strSql = strSql + (Join(arrSql, ", "))
    strSql = WrapTransaction(strSql, "save_patient")
    
    objConn.Execute strSql
    
    objConn.Close
    
    Exit Sub
    
SavePatientError:

    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not save patient details to database: " & strSql
    
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
    
    strUser = ModMetaVision.MetaVision_GetUserLogin()
    strLN = WrapString("Bollen")
    strFN = WrapString("")
    strRole = WrapString("")
        
    arrSql = Array(WrapString(strUser), strLN, strFN, strRole, 0)
        
    InitConnection
    
    objConn.Open
    
    If PrescriberExists(strUser) Then
        strSql = "EXEC UpdatePrescriber "
    Else
        strSql = "EXEC InsertPrescriber "
    End If
    
    strSql = strSql & (Join(arrSql, ", "))
    strSql = WrapTransaction(strSql, "save_prescriber")
    
    ModUtils.CopyToClipboard strSql
    objConn.Execute strSql
    
    objConn.Close
    ModUtils.CopyToClipboard strSql
    
    Exit Sub
    
SavePrescriberError:

    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not save prescriber details to the database: " & strSql
    
End Sub

Private Sub ClearTestDatabase()

    Dim strSql As String
    
    strSql = "EXEC dbo.ClearDatabase 'UMCU_WKZ_AP_Test'"

    InitConnection
    
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
    ModMessage.ShowMsgBoxOK PrescriberExists(ModMetaVision.MetaVision_GetUserLogin())

End Sub

Public Function Database_GetLatestVersion(strHospNum) As String

    Dim strSql As String
    Dim objRs As Recordset
    Dim strResult As String
    
    On Error GoTo Database_GetLatestVersionError
    
    strSql = "SELECT dbo.GetLatestPrescriptionDateForHospitalNumber('" & strHospNum & "')"
    
    InitConnection
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    If Not objRs.EOF Then
        strResult = IIf(IsNull(objRs.Fields(0)), "", objRs.Fields(0))
    Else
        strResult = ""
    End If

    objConn.Close
    Set objRs = Nothing
    
    Database_GetLatestVersion = strResult
    
    Exit Function
    
Database_GetLatestVersionError:

    ModLog.LogError Err, "Could not get latest version for patient: " & strHospNum

End Function

Private Sub Test_Database_GetLatestVersion()

    ModMessage.ShowMsgBoxOK Database_GetLatestVersion("1234")

End Sub

Public Sub Database_SaveData(strTimeStamp As String, strHospNum, strPrescriber As String, objData As Range, objText As Range, blnProgress As Boolean)

    InitDatabase
    objDatabase.SaveData strTimeStamp, strHospNum, strPrescriber, objData, objText, blnProgress
    
End Sub

Private Function IsLogical(ByVal varVal As Variant) As Boolean

    IsLogical = LCase(varVal) = "waar" Or LCase(varVal) = "onwaar"
    
End Function

Private Sub GetPatientData(ByVal strHospNum, Optional ByVal strVersion As String = "")

    Dim strSql As String
    Dim intC As Long
    Dim intN As Integer
    Dim strPar As String
    Dim varVal As Variant
    Dim varEmp As Variant
    Dim objRs As Recordset
    Dim blnVersionSet As Boolean
    
    On Error GoTo Database_GetPatientDataError
    
    strSql = strSql & "SELECT * FROM "
    If strVersion = vbNullString Then
        strSql = strSql & "dbo.GetLatestPrescriptionData('" & strHospNum & "')"
    Else
        strVersion = "{ts'" & strVersion & "'}"
        strSql = strSql & "dbo.GetLatestPrescriptionDataForVersion('" & strHospNum & "', " & strVersion & ")"
    End If
    
    InitConnection
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    intC = shtPatData.Range("A1").Rows.Count
    Do While Not objRs.EOF
        If Not blnVersionSet Then
            ModRange.SetRangeValue constVersie, objRs.Fields("DateTime").Value
            blnVersionSet = True
        End If
        
        strPar = Trim(objRs.Fields("Parameter").Value)
        varVal = Trim(objRs.Fields("Data").Value)
        
        If IsNumeric(varVal) Then varVal = CDbl(varVal)
        If IsLogical(varVal) Then varVal = CBool(varVal)
        
        ModRange.SetRangeValue strPar, varVal
        
        intN = intN + 1
        ModProgress.SetJobPercentage "Patient data laden", intC, intN
        
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    Exit Sub

Database_GetPatientDataError:
    
    ModMessage.ShowMsgBoxError "Kan patient met ziekenhuis nummer " & strHospNum & " niet laden."
    
    ModLog.LogError Err, "Could not get patient data with hospitalnumber " & strHospNum & " with SQL: " & vbNewLine & strSql
    objConn.Close
    

End Sub


Public Sub Database_GetPatientDataForVersion(strHospNum As String, strVersion)

    GetPatientData strHospNum, strVersion
    
End Sub

Public Sub Database_GetPatientData(strHospNum As String)

    GetPatientData strHospNum
    
End Sub

Private Sub Test_DatabaseGetPatientData()

    ModProgress.StartProgress "Patient data ophalen"
    Database_GetPatientData "0250574"
    ModProgress.FinishProgress

End Sub

Private Function GetSaveNeoConfigMedContSql(ByVal strVersion, blnIsBatch As Boolean) As String

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
    Dim strSql
        
    strTable = "Tbl_Admin_NeoMedCont"
    strVersion = "{ts'" & strVersion & "'}"
    strDepartment = "Neonatologie"
    strDilutionText = ModRange.GetRangeValue("Var_Neo_MedCont_VerdunningTekst", vbNullString)
    
    Set objSrc = ModRange.GetRange(strTable)
    If Not blnIsBatch Then
    
        strSql = strSql & "DECLARE @RC int" & vbNewLine
        strSql = strSql & "DECLARE @version datetime" & vbNewLine
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
        strDoseAdvice = objSrc.Cells(intR, 12).Value2
        dblSolutionVolume = objSrc.Cells(intR, 13).Value2
        dblDripQuantity = objSrc.Cells(intR, 14).Value2
        strProduct = objSrc.Cells(intR, 15).Value2
        dblShelfLife = objSrc.Cells(intR, 16).Value2
        strShelfCondition = objSrc.Cells(intR, 17).Value2
        strPreparationText = objSrc.Cells(intR, 18).Value2
    
        strSql = strSql & "SET @version = " & strVersion & "" & vbNewLine
        strSql = strSql & "SET @department  = '" & strDepartment & "'" & vbNewLine
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
        strSql = strSql & "EXECUTE @RC = [dbo].[InsertConfigMedCont] " & vbNewLine
        strSql = strSql & "   @version" & vbNewLine
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
    

    GetSaveNeoConfigMedContSql = strSql
    
End Function

'ALTER PROCEDURE [dbo].[InsertConfigMedCont]
'    -- Add the parameters for the stored procedure here
'             @version DATETIME
'           , @department NVARCHAR(60)
'           , @generic NVARCHAR(300)
'           , @genericUnit NVARCHAR(50)
'           , @genericQuantity FLOAT
'           , @genericVolume FLOAT
'           , @solutionVolume FLOAT
'           , @solution_2_6_Quantity FLOAT
'           , @solution_2_6_Volume FLOAT
'           , @solution_6_11_Quantity FLOAT
'           , @solution_6_11_Volume FLOAT
'           , @solution_11_40_Quantity FLOAT
'           , @solution_11_40_Volume FLOAT
'           , @solution_40_Quantity FLOAT
'           , @solution_40_Volume FLOAT
'           , @minConcentration FLOAT
'           , @maxConcentration FLOAT
'           , @solution NVARCHAR(300)
'           , @dripQuantity FLOAT
'           , @doseUnit NVARCHAR(50)
'           , @minDose FLOAT
'           , @maxDose FLOAT
'           , @absMaxDose FLOAT
'           , @doseAdvice NVARCHAR(MAX)
'           , @product NVARCHAR(MAX)
'           , @shelfLife FLOAT
'           , @shelfCondition NVARCHAR(50)
'           , @preparationText NVARCHAR(MAX)
'           , @signed BIT
'           , @dilutionText NVARCHAR(MAX)
Public Sub Database_SaveNeoConfigMedCont()

    Dim strSql As String
    Dim strVersion As String
    
    On Error GoTo ErrorHandler
     
    ModProgress.StartProgress "Neo Continue Medicatie Configuratie Opslaan"

    strVersion = FormatDateTimeSeconds(Now())
    strSql = GetSaveNeoConfigMedContSql(strVersion, False)
    strSql = ModDatabase.WrapTransaction(strSql, "insert_neoconfigmedcont")
    
    InitConnection
    
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    Database_LogAction "Save neonatal configuration for continuous medication", , , strVersion
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:

    objConn.Close
    ModProgress.FinishProgress

    ModUtils.CopyToClipboard strSql
    ModMessage.ShowMsgBoxError "Kon de configuratie voor de neonatologie continue medicatie niet opslaan"
    ModLog.LogError Err, "Database_SaveNeoConfigMedCont with sql: " & vbNewLine & strSql
    
End Sub

'SELECT [Version]
'      ,[Department]
'      ,[Generic]
'      ,[GenericUnit]
'      ,[GenericQuantity]
'      ,[GenericVolume]
'      ,[SolutionVolume]
'      ,[Solution_2_6_Quantity]
'      ,[Solution_2_6_Volume]
'      ,[Solution_6_11_Quantity]
'      ,[Solution_6_11_Volume]
'      ,[Solution_11_40_Quantity]
'      ,[Solution_11_40_Volume]
'      ,[Solution_40_Quantity]
'      ,[Solution_40_Volume]
'      ,[MinConcentration]
'      ,[MaxConcentration]
'      ,[Solution]
'      ,[DripQuantity]
'      ,[DoseUnit]
'      ,[MinDose]
'      ,[MaxDose]
'      ,[AbsMaxDose]
'      ,[DoseAdvice]
'      ,[Product]
'      ,[ShelfLife]
'      ,[ShelfCondition]
'      ,[PreparationText]
'      ,[Signed]
'      ,[DilutionText]
Public Sub Database_LoadNeoConfigMedCont()

    Dim strSql As String
    Dim objRs As Recordset
    Dim intC As Integer
    Dim intR As Integer
    Dim objSrc As Range
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Configuratie voor Neonatologie Continue Medicatie laden"
    
    Set objSrc = ModRange.GetRange("Tbl_Admin_NeoMedCont")
    
    InitConnection
    
    strSql = "SELECT * FROM [dbo].[GetLatestConfigMedContForDepartment] ('Neonatologie')"

    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        intR = intR + 1
        If intR > 24 Then GoTo ErrorHandler
        
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
        objSrc.Cells(intR, 12).Value2 = objRs.Fields("DoseAdvice").Value
        objSrc.Cells(intR, 13).Value2 = objRs.Fields("SolutionVolume").Value
        objSrc.Cells(intR, 14).Value2 = objRs.Fields("DripQuantity").Value
        objSrc.Cells(intR, 15).Value2 = objRs.Fields("Product").Value
        objSrc.Cells(intR, 16).Value2 = objRs.Fields("ShelfLife").Value
        objSrc.Cells(intR, 17).Value2 = objRs.Fields("ShelfCondition").Value
        objSrc.Cells(intR, 18).Value2 = objRs.Fields("PreparationText").Value
        objSrc.Cells(intR, 19).Value2 = objRs.Fields("DoseAdvice").Value
        
        ModRange.SetRangeValue "Var_Neo_MedCont_VerdunningTekst", objRs.Fields("DilutionText").Value
        
        ModProgress.SetJobPercentage "Data laden", 24, intR
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:

    ModProgress.FinishProgress
    objConn.Close

    ModUtils.CopyToClipboard strSql
    ModMessage.ShowMsgBoxError "Kon de configuratie voor de neonatologie continue medicatie niet laden"
    ModLog.LogError Err, "Database_LoadNeoConfigMedCont with sql: " & vbNewLine & strSql

End Sub

Private Function GetSavePediatrieConfigMedContSql(ByVal strVersion As String, ByVal blnIsBatch As Boolean) As String

    Dim strSql As String
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
    strVersion = "{ts'" & strVersion & "'}"
    strDepartment = "Pediatrie"
    strDilutionText = ""
    
    Set objSrc = ModRange.GetRange(strTable)
    If Not blnIsBatch Then strSql = strSql & "DECLARE @RC int" & vbNewLine
    If Not blnIsBatch Then strSql = strSql & "DECLARE @version datetime" & vbNewLine
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
            
        strSql = strSql & "SET @version = " & strVersion & "" & vbNewLine
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
        strSql = strSql & "EXECUTE @RC = [dbo].[InsertConfigMedCont] " & vbNewLine
        strSql = strSql & "   @version" & vbNewLine
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
    
    GetSavePediatrieConfigMedContSql = strSql

End Function

'ALTER PROCEDURE [dbo].[InsertConfigMedCont]
'    -- Add the parameters for the stored procedure here
'             @version DATETIME
'           , @department NVARCHAR(60)
'           , @generic NVARCHAR(300)
'           , @genericUnit NVARCHAR(50)
'           , @genericQuantity FLOAT
'           , @genericVolume FLOAT
'           , @solutionVolume FLOAT
'           , @solution_2_6_Quantity FLOAT
'           , @solution_2_6_Volume FLOAT
'           , @solution_6_11_Quantity FLOAT
'           , @solution_6_11_Volume FLOAT
'           , @solution_11_40_Quantity FLOAT
'           , @solution_11_40_Volume FLOAT
'           , @solution_40_Quantity FLOAT
'           , @solution_40_Volume FLOAT
'           , @minConcentration FLOAT
'           , @maxConcentration FLOAT
'           , @solution NVARCHAR(300)
'           , @dripQuantity FLOAT
'           , @doseUnit NVARCHAR(50)
'           , @minDose FLOAT
'           , @maxDose FLOAT
'           , @absMaxDose FLOAT
'           , @doseAdvice NVARCHAR(MAX)
'           , @product NVARCHAR(MAX)
'           , @shelfLife FLOAT
'           , @shelfCondition NVARCHAR(50)
'           , @preparationText NVARCHAR(MAX)
'           , @signed BIT
'           , @dilutionText NVARCHAR(MAX)
Public Sub Database_SavePediatrieConfigMedCont()

    Dim strSql As String
    Dim strVersion As String

    On Error GoTo ErrorHandler
     
    ModProgress.StartProgress "Pediatrie Continue Medicatie Configuratie Opslaan"
    
    strVersion = FormatDateTimeSeconds(Now())
    strSql = GetSavePediatrieConfigMedContSql(strVersion, False)
    strSql = ModDatabase.WrapTransaction(strSql, "insert_pedconfigmedcont")
    
    InitConnection
    
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    Database_LogAction "Save pediatric configuration for continuous medication", , , strVersion
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:

    objConn.Close
    ModProgress.FinishProgress

    ModUtils.CopyToClipboard strSql
    ModMessage.ShowMsgBoxError "Kon de configuratie voor de pediatrie continue medicatie niet opslaan"
    ModLog.LogError Err, "Database_SavePedConfigMedCont with sql: " & vbNewLine & strSql
    
End Sub

'SELECT [Version]
'      ,[Department]
'      ,[Generic]
'      ,[GenericUnit]
'      ,[GenericQuantity]
'      ,[GenericVolume]
'      ,[SolutionVolume]
'      ,[Solution_2_6_Quantity]
'      ,[Solution_2_6_Volume]
'      ,[Solution_6_11_Quantity]
'      ,[Solution_6_11_Volume]
'      ,[Solution_11_40_Quantity]
'      ,[Solution_11_40_Volume]
'      ,[Solution_40_Quantity]
'      ,[Solution_40_Volume]
'      ,[MinConcentration]
'      ,[MaxConcentration]
'      ,[Solution]
'      ,[DripQuantity]
'      ,[DoseUnit]
'      ,[MinDose]
'      ,[MaxDose]
'      ,[AbsMaxDose]
'      ,[DoseAdvice]
'      ,[Product]
'      ,[ShelfLife]
'      ,[ShelfCondition]
'      ,[PreparationText]
'      ,[Signed]
'      ,[DilutionText]
Public Sub Database_LoadPedConfigMedCont()

    Dim strSql As String
    Dim objRs As Recordset
    Dim intC As Integer
    Dim intR As Integer
    Dim objSrc As Range
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Configuratie voor Pediatrie Continue Medicatie"
    
    Set objSrc = ModRange.GetRange("Tbl_Admin_PedMedCont")
    
    InitConnection
    
    strSql = "SELECT * FROM [dbo].[GetLatestConfigMedContForDepartment] ('Pediatrie')"

    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        intR = intR + 1
        If intR > 31 Then GoTo ErrorHandler
        
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
                
        ModProgress.SetJobPercentage "Data laden", 31, intR
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:

    ModProgress.FinishProgress
    objConn.Close

    ModUtils.CopyToClipboard strSql
    ModMessage.ShowMsgBoxError "Kon de configuratie voor de neonatologie continue medicatie niet laden"
    ModLog.LogError Err, "Database_LoadPedConfigMedCont with sql: " & vbNewLine & strSql

End Sub

Private Function GetSaveConfigParentSql(ByVal strVersion As String) As String

    Dim strSql As String
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
    strSql = strSql & "DECLARE @version datetime" & vbNewLine
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
    
    strVersion = "{ts'" & strVersion & "'}"
    
    For intR = 1 To intC
    
        strSql = strSql & "SET @version  = " & strVersion & vbNewLine
        
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
        strSql = strSql & "EXECUTE @RC = [dbo].[InsertConfigParEnt] " & vbNewLine
        strSql = strSql & "   @version" & vbNewLine
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
 
    GetSaveConfigParentSql = strSql
    
End Function

Public Sub Database_SaveConfigParEnt()

    Dim strSql As String
    Dim strVersion As String
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Configuratie voor parenteralia"
    
    strVersion = FormatDateTimeSeconds(Now())
    strSql = GetSaveConfigParentSql(strVersion)
    strSql = WrapTransaction(strSql, "insert_configparent")
    
    InitConnection
    
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    Database_LogAction "Save configuration for parenteral products", , , strVersion
    ModProgress.FinishProgress
    
    Exit Sub
    
ErrorHandler:
    
    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModProgress.FinishProgress
    ModLog.LogError Err, "Database_SaveConfigParEnt"
    

End Sub

'SELECT [Version]
'      ,[Name]
'      ,[Energy]
'      ,[Protein]
'      ,[Carbohydrate]
'      ,[Lipid]
'      ,[Sodium]
'      ,[Potassium]
'      ,[Calcium]
'      ,[Phosphor]
'      ,[Magnesium]
'      ,[Iron]
'      ,[VitD]
'      ,[Chloride]
'      ,[Product]
'      ,[Signed]
Public Sub Database_LoadConfigParEnt()

    Dim strSql As String
    Dim objRs As Recordset
    Dim intC As Integer
    Dim intR As Integer
    Dim objTable As Range
    
    On Error GoTo ErrorHandler
    
    ModProgress.StartProgress "Parenteralia Configuratie"
    
    strSql = "SELECT * FROM [dbo].[GetLatestConfigParEnt] ()"

    InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Set objTable = ModRange.GetRange("Tbl_Admin_ParEnt")
    
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
    
    objConn.Close
    ModProgress.FinishProgress

    Exit Sub
    
ErrorHandler:

    ModProgress.FinishProgress
    objConn.Close
    ModLog.LogError Err, "Database_LoadConfigParEnt"
End Sub

Public Function Database_GetConfigParEntVersions() As String()

    Dim arrVersions() As String
    Dim objRs As Recordset
    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    strSql = "SELECT * FROM [dbo].[GetConfigParEntVersions] ()" & vbNewLine
    strSql = strSql & "ORDER BY [Version] DESC "
    
    InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        ModArray.AddItemToStringArray arrVersions, objRs.Fields("Version").Value
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    Database_GetConfigParEntVersions = arrVersions
    
    Exit Function
    
ErrorHandler:

    ModLog.LogError Err, "Database_GetConfigParEntVersions"
    objConn.Close
    
End Function

Public Function Database_GetConfigParEnt(Optional ByVal strVersion As String = "") As Collection

    Dim objCol As Collection
    Dim objParEnt As ClassParent
        
    Dim strSql As String
    Dim objRs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set objCol = New Collection
    
    If strVersion = vbNullString Then
        strSql = "SELECT * FROM dbo.GetLatestConfigParEnt()"
    Else
        strSql = "SELECT * FROM dbo.GetConfigParEntForVersion({ts'" & strVersion & "'})"
    End If
    
    InitConnection
    
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

Public Function Database_GetVersions(ByVal strHospNum As String) As String()

    Dim arrVersions() As String
    Dim strSql As String
    Dim objRs As Recordset
    
    On Error GoTo ErrorHandler
    
    strSql = "SELECT * FROM [dbo].[GetPrescriptionVersionsForHospitalNumber] ('" & strHospNum & "')"
    strSql = strSql & "ORDER BY [DateTime] Desc"

    InitConnection
    
    objConn.Open
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        ModArray.AddItemToStringArray arrVersions, objRs.Fields("DateTime").Value
        objRs.MoveNext
    Loop
    
    objConn.Close

    Database_GetVersions = arrVersions

    Exit Function
    
ErrorHandler:

    ModUtils.CopyToClipboard strSql

    ModLog.LogError Err, "Database_GetVersions"
    objConn.Close
    
End Function

Private Sub Test_Database_GetVersions()

    Dim intN As Integer
    Dim arrVersions() As String
    
    arrVersions = Database_GetVersions("0239080")
    
    For intN = 0 To UBound(arrVersions)
        ModMessage.ShowMsgBoxInfo arrVersions(intN)
    Next
    
End Sub

Public Sub Database_ClearDatabase()

    Dim strDatabase As String
    Dim strSql As String
    Dim strVersion As String
    
    On Error GoTo ErrorHandler
    
    strDatabase = ModSetting.Setting_GetDatabase()
    
    If ModMessage.ShowMsgBoxYesNo("Database " & strDatabase & " leeg maken?") = vbYes Then
        If ModMessage.ShowMsgBoxYesNo("Weet u het zeker dat " & strDatabase & " leeggemaakt moet worden?") Then
            ModProgress.StartProgress "Clear Database"
            
            strVersion = FormatDateTimeSeconds(Now())
            
            strSql = "EXEC dbo.ClearDatabase  " & WrapString(strDatabase)
            strSql = strSql & vbNewLine & GetSaveConfigParentSql(strVersion)
            strSql = strSql & vbNewLine & GetSavePediatrieConfigMedContSql(strVersion, True)
            strSql = strSql & vbNewLine & GetSaveNeoConfigMedContSql(strVersion, True)
            strSql = WrapTransaction(strSql, "cleardatabase_trans")
            
            InitConnection
            objConn.Open
            objConn.Execute strSql
            objConn.Close
            
            ModProgress.FinishProgress
        End If
    End If
    
    Database_LogAction "Clear database", , , strVersion

    Exit Sub

ErrorHandler:

    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not clear database with SQL: " & vbNewLine & strSql
End Sub

Public Sub Database_LogAction(ByVal strText As String, Optional strPrescriber As String, Optional ByVal strHospNum As String = "", Optional ByVal strVersion As String = "")

    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    If Not Setting_UseDatabase Then Exit Sub
    
    strHospNum = IIf(strHospNum = vbNullString, ModPatient.Patient_GetHospitalNumber(), strHospNum)
    strVersion = IIf(strVersion = vbNullString, ModDate.FormatDateTimeSeconds(Now()), strVersion)
    strPrescriber = IIf(strPrescriber = vbNullString, ModMetaVision.MetaVision_GetUserLogin(), strPrescriber)

    strSql = "EXEC dbo.InsertLog "
    strSql = strSql & WrapString(strPrescriber) & ", " & WrapString(strHospNum) & ", " & WrapDateTime(strVersion) & ", " & WrapString(strText)
    
    InitConnection
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

    Database_LogAction "Test"

End Sub

Public Function Database_GetNeoConfigMedCont(Optional ByVal strVersion As String = "") As Collection

    Dim strSql As String
    Dim objRs As Recordset
    Dim objCol As Collection
    Dim objConfig As ClassNeoMedCont
    
    On Error GoTo ErrorHandler
       
    InitConnection
    
    If strVersion = vbNullString Then
        strSql = "SELECT * FROM [dbo].[GetLatestConfigMedContForDepartment] ('Neonatologie')"
    Else
        strSql = "SELECT * FROM [dbo].[GetConfigMedContForDepartmentWithVersion] ('Neonatologie'), " & WrapDateTime(strVersion)
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


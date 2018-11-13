Attribute VB_Name = "ModDatabase"
Option Explicit

Private objConn As ADODB.Connection

Private Const constSecret As String = "secret"

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
        objDatabase.InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    End If

End Sub

Private Sub Test_InitDatabase()

    InitDatabase

End Sub

Private Sub InitConnection(ByVal strServer As String, ByVal strDatabase As String)

    Dim strSecret As String
    Dim strUser As String
    Dim strPw As String
    
    On Error GoTo InitConnectionError
    
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
        ModLog.LogError "Bestand secret niet aanwezig"
    End If
    
    Exit Sub
    
InitConnectionError:
    MsgBox "Geen toegang tot de database!"
    ModLog.LogError "InitConnection Failed"

End Sub

Private Sub Test_InitConnectionWithAPDB()

    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"

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

    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    
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
    
    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    
    objConn.Open

    PatientExists = Not objConn.Execute(strSql).EOF

End Function

Private Sub Test_PatientExists()

    MsgBox PatientExists("000")

End Sub

Private Function PrescriberExists(strUser As String) As Boolean

    Dim strSql As String
    
    strSql = "SELECT * FROM dbo.GetPrescribers('" & strUser & "')"
    
    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    
    objConn.Open

    PrescriberExists = Not objConn.Execute(strSql).EOF

End Function

Private Sub Test_PrescriberExists()

    MsgBox PrescriberExists("000")

End Sub

Private Function WrapString(varItem As Variant) As Variant

    WrapString = "'" & varItem & "'"

End Function

Public Function WrapTransaction(strSql As String, strName As String) As String

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
    
    strHN = WrapString(ModPatient.PatientHospNum)
    strBD = WrapString(ModDate.FormatDateYearMonthDay(ModPatient.Patient_BirthDate))
    strAN = WrapString(ModPatient.PatientAchterNaam)
    strVN = WrapString(ModPatient.PatientVoorNaam)
    strGN = WrapString(ModRange.GetRangeValue(constGeslacht, Null))
    intGW = ModRange.GetRangeValue(constWeken, Null)
    intGD = ModRange.GetRangeValue(constDagen, Null)
    dblBW = ModRange.GetRangeValue(constGebGew, Null)
        
    arrSql = Array(strHN, strBD, strAN, strVN, strGN, intGW, intGD, dblBW)
        
    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    
    objConn.Open
    
    If PatientExists(ModPatient.PatientHospNum()) Then
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
    ModLog.LogError "Could not save patient details to database: " & strSql
    
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
        
    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    
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
    ModLog.LogError "Could not save prescriber details to the database: " & strSql
    
End Sub

Private Sub ClearTestDatabase()

    Dim strSql As String
    
    strSql = "EXEC dbo.ClearDatabase 'UMCU_WKZ_AP_Test'"

    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    
    objConn.Open
    objConn.Execute strSql
    objConn.Close
    
    Exit Sub
    
ClearTestDatabaseError:

    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModLog.LogError "Could not clear the database: " & strSql

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
    
    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    
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

    ModLog.LogError "Could not get latest version for patient: " & strHospNum

End Function

Private Sub Test_Database_GetLatestVersion()

    ModMessage.ShowMsgBoxOK Database_GetLatestVersion("1234")

End Sub

Public Sub Database_SaveData(strTimeStamp As String, strHospNum, strPrescriber As String, objData As Range, objText As Range, blnProgress As Boolean)

    InitDatabase
    objDatabase.SaveData strTimeStamp, strHospNum, strPrescriber, objData, objText, blnProgress
    
End Sub

Public Sub Database_GetPatientData(strHospNum As String)

    Dim strSql As String
    Dim intC As Long
    Dim intN As Integer
    Dim strPar As String
    Dim varVal As Variant
    Dim varEmp As Variant
    Dim objRs As Recordset
    
    On Error GoTo Database_GetPatientDataError
    
    strSql = strSql & "SELECT * FROM dbo.GetLatestPrescriptionData('" & strHospNum & "')"
    
    InitConnection "mvtst_wkz", "UMCU_WKZ_AP_Test"
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    intC = shtPatData.Range("A1").Rows.Count
    Do While Not objRs.EOF
        strPar = Trim(objRs.Fields("Parameter").Value)
        varVal = objRs.Fields("Data").Value
        ModRange.SetRangeValue strPar, varVal
        
        intN = intN + 1
        ModProgress.SetJobPercentage "Patient data laden", intC, intN
        
        objRs.MoveNext
    Loop
    
    objConn.Close
    
    Exit Sub

Database_GetPatientDataError:
    
    ModMessage.ShowMsgBoxError "Kan patient met ziekenhuis nummer " & strHospNum & " niet laden."
    
    ModLog.LogError "Could not get patient data with hospitalnumber " & strHospNum & " with SQL: " & vbNewLine & strSql
    objConn.Close
    
End Sub

Private Sub Test_DatabaseGetPatientData()

    ModProgress.StartProgress "Patient data ophalen"
    Database_GetPatientData "0250574"
    ModProgress.FinishProgress

End Sub

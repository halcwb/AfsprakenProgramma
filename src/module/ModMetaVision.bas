Attribute VB_Name = "ModMetaVision"
Option Explicit

Private objConn As ADODB.Connection

Private Const constSecret As String = "secret"

Private Const constBasePath1 As String = "HKCU\SOFTWARE\UMCU\MV\"
Private Const constBasePath2 As String = "HKLM\SOFTWARE\iMD Soft\"
Private Const constBasePath3 As String = "HKEY_CURRENT_USER\Software\Classes\VirtualStore\MACHINE\SOFTWARE\Wow6432Node\iMD Soft\"

Private Const constSettings As String = "Settings"

Private Const constUserId As String = "UserID"
Private Const constUserLogin As String = "UserLogin"

Private Const constCurrentPatient As String = "Current Patient"
Private Const constPatientId As String = "PatientID"

Private Const constConnection As String = "Database Connect"

Private Const constServer As String = "Server"
Private Const constDatabase As String = "Database"

Private Const constEmpiServer As String = "EMPI Server"
Private Const constEMPIDb As String = "EMPI Database"

Private Const constDepartment As String = "Afdeling"
Private Const constDomain As String = "Domain Department"

Private Const constTblMedOpdr As String = "Tbl_Glob_MedOpdr"

Private Enum ParamIds
    Leeftijd = 7225
    Gewicht = 12677
    OpnameReden = 10007
    Allergie = 12688
    Medicatie = 12716
    Beleid = 8211
    Concluse = 12727
    Voorgeschiedenis = 8216
    SumTxtRespiratie = 13288
    SumTxtCirculatie = 7908
    SumTxtNeurologie = 13856
    SumTxtInfectie = 7989
    SumTxtVB = 7911
    SumTxtGI = 7911
    SumTxtLab = 7990
End Enum

Private Enum TableRows
    Patient = 2
    OpnameReden = 3
    Aandachtspunten = 4
    Voorgeschiedenis = 5
    Status = 6
    KorteTermijn = 7
    Medicatie = 8
End Enum

Private Function GetPatientBed(ByVal strPatId As String, ByVal strPatNum As String) As String

    Dim strServer As String
    Dim strDb As String
    Dim objRs As Recordset
    Dim strSql As String
    Dim strBed As String
    Dim strId As String
    Dim blnFound As Boolean
    
    strServer = MetaVision_GetServer()
    strDb = MetaVision_GetDatabase()
    strSql = GetPatientListSQL(strPatId, strPatNum, vbNullString)
    
    If strServer = vbNullString Or strDb = vbNullString Or strSql = vbNullString Then
        strBed = vbNullString
    Else
        InitConnection strServer, strDb
        
        objConn.Open
        
        Set objRs = objConn.Execute(strSql)
        
        blnFound = False
        Do While Not objRs.EOF And Not blnFound
            strId = objRs.Fields("PatientId")
            If strId = strPatId Then
                strBed = Trim(CStr(objRs.Fields("BedName")))
                blnFound = True
            End If
            objRs.MoveNext
        Loop
        
        objConn.Close
        Set objRs = Nothing
        
    End If
    
    GetPatientBed = strBed
    
    
End Function

Private Function GetDatabaseNameSQL(ByVal strDepartment As String) As String

    Dim strSql As String
    
    strSql = strSql & "DECLARE @Name AS nvarchar(60)" & vbNewLine
    
    strSql = strSql & "SET @Name = '" & strDepartment & "'" & vbNewLine
    
    strSql = strSql & "SELECT LU.Name, DatabaseName" & vbNewLine
    strSql = strSql & "FROM Departments AS D" & vbNewLine
    strSql = strSql & "INNER JOIN LogicalUnits AS LU ON LU.DepartmentID = D.ID" & vbNewLine
    strSql = strSql & "WHERE LU.Name = @Name" & vbNewLine
    
    GetDatabaseNameSQL = strSql

End Function

Private Function GetPatientListSQL(ByVal strPatId As String, ByVal strPatNum As String, ByVal strBed As String) As String

    Dim strSql As String
    
    If strPatId = vbNullString And strPatNum = vbNullString And strBed = vbNullString Then
        strSql = vbNullString
    Else
        strSql = strSql & "DECLARE @patId AS int" & vbNewLine
        strSql = strSql & "DECLARE @bed AS nvarchar(100)" & vbNewLine
        strSql = strSql & "DECLARE @patNum AS nvarchar(40)" & vbNewLine
        strSql = strSql & "DECLARE @bd AS int" & vbNewLine
        strSql = strSql & "DECLARE @weightKg AS int" & vbNewLine
        strSql = strSql & "DECLARE @weightGr AS int" & vbNewLine
        strSql = strSql & "DECLARE @bwGr AS int" & vbNewLine
        strSql = strSql & "DECLARE @length AS int" & vbNewLine
        strSql = strSql & "DECLARE @gesl AS int" & vbNewLine
        strSql = strSql & "DECLARE @adD AS int" & vbNewLine
        strSql = strSql & "DECLARE @adW AS int" & vbNewLine
        
        strSql = strSql & "SET @bd = 5372" & vbNewLine
        strSql = strSql & "SET @weightKg = 8365" & vbNewLine
        strSql = strSql & "SET @weightGr = 8456" & vbNewLine
        strSql = strSql & "SET @bwGr = 7734" & vbNewLine
        strSql = strSql & "SET @adD = 10213" & vbNewLine
        strSql = strSql & "SET @adW = 10214" & vbNewLine
        strSql = strSql & "SET @length = 9505" & vbNewLine
        strSql = strSql & "SET @gesl = 5373" & vbNewLine
        
        If Not strBed = vbNullString Then
            strSql = strSql & "SET @Bed = '" & strBed & "'" & vbNewLine
        ElseIf Not strPatNum = vbNullString Then
            strSql = strSql & "SET @patNum = '" & strPatNum & "'" & vbNewLine
        ElseIf Not strPatId = vbNullString And Not strPatId = "-1" Then
            strSql = strSql & "SET @patId = " & strPatId & vbNewLine
        End If
        
        strSql = strSql & "SELECT DISTINCT" & vbNewLine
        strSql = strSql & "pl.PatientID" & vbNewLine
        strSql = strSql & ", pl.HospitalNumber" & vbNewLine
        strSql = strSql & ", pl.LastName" & vbNewLine
        strSql = strSql & ", pl.FirstName" & vbNewLine
        strSql = strSql & ", (SELECT TOP 1 dts.Value" & vbNewLine
        strSql = strSql & "   FROM DateTimeSignals dts" & vbNewLine
        strSql = strSql & "   WHERE dts.PatientID = pl.PatientID AND dts.ParameterID = @bd" & vbNewLine
        strSql = strSql & "   ORDER BY dts.[Time] DESC)" & vbNewLine
        strSql = strSql & "   BirthDate" & vbNewLine
        strSql = strSql & ", (SELECT TOP 1 s.value / 1000 " & vbNewLine
        strSql = strSql & "   FROM Signals s " & vbNewLine
        strSql = strSql & "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @weightKg " & vbNewLine
        strSql = strSql & "   ORDER BY s.Time DESC) WeightKg" & vbNewLine
        strSql = strSql & ", (SELECT TOP 1 s.Value " & vbNewLine
        strSql = strSql & "   FROM Signals s " & vbNewLine
        strSql = strSql & "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @weightGr " & vbNewLine
        strSql = strSql & "   ORDER BY s.Time DESC) WeightGr" & vbNewLine
        strSql = strSql & ", (SELECT TOP 1 s.value * 100 " & vbNewLine
        strSql = strSql & "   FROM Signals s " & vbNewLine
        strSql = strSql & "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @length " & vbNewLine
        strSql = strSql & "   ORDER BY s.Time DESC) LengthCm" & vbNewLine
        strSql = strSql & ", (SELECT TOP 1 s.Value " & vbNewLine
        strSql = strSql & "   FROM Signals s " & vbNewLine
        strSql = strSql & "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @bwGr " & vbNewLine
        strSql = strSql & "   ORDER BY s.Time DESC) BirthWeightGr" & vbNewLine
        strSql = strSql & ", (SELECT TOP 1 s.value / (60 * 24) " & vbNewLine
        strSql = strSql & "   FROM Signals s " & vbNewLine
        strSql = strSql & "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @adD ORDER BY s.Time DESC) PregnDays" & vbNewLine
        strSql = strSql & ", (SELECT TOP 1 s.value / (7 * 60 * 24) FROM Signals s WHERE s.PatientID = pl.PatientID AND s.ParameterID = @adW " & vbNewLine
        strSql = strSql & "   ORDER BY s.Time DESC) PregnWeeks" & vbNewLine
        strSql = strSql & ", (SELECT TOP 1 pt.Text " & vbNewLine
        strSql = strSql & "   FROM ParametersText pt " & vbNewLine
        strSql = strSql & "   INNER JOIN TextSignals ts ON pt.ParameterID = ts.ParameterID AND pt.TextID = ts.TextID " & vbNewLine
        strSql = strSql & "   INNER JOIN Parameters p ON p.ParameterID = ts.ParameterID" & vbNewLine
        strSql = strSql & "   WHERE ts.PatientID = pl.PatientID AND p.ParameterID = @Gesl" & vbNewLine
        strSql = strSql & "   ORDER BY ts.Time DESC) Geslacht" & vbNewLine
        strSql = strSql & ", lu.Name Department" & vbNewLine
        strSql = strSql & ", b.BedName" & vbNewLine
        strSql = strSql & ", pl.LocationFromTime" & vbNewLine
        strSql = strSql & ", pl.TimeLog" & vbNewLine
        strSql = strSql & "FROM PatientLogs pl" & vbNewLine
        strSql = strSql & "INNER JOIN Patients pat ON pat.PatientID = pl.PatientID" & vbNewLine
        strSql = strSql & "LEFT JOIN LogicalUnits lu ON lu.LogicalUnitID = pl.LogicalUnitID" & vbNewLine
        strSql = strSql & "LEFT JOIN Beds b ON b.BedID = pl.BedID" & vbNewLine
        strSql = strSql & "WHERE " & vbNewLine
        strSql = strSql & "(@patId IS NULL OR pl.PatientID = @patId)" & vbNewLine
        strSql = strSql & "AND (@patNum IS NULL OR pl.HospitalNumber = @patNum)" & vbNewLine
        strSql = strSql & "AND (@bed IS NULL OR RTRIM(LTRIM(b.BedName)) = RTRIM(LTRIM(@bed)))" & vbNewLine
        strSql = strSql & "ORDER BY pl.TimeLog DESC" & vbNewLine

    End If
    
    ModUtils.CopyToClipboard strSql
    GetPatientListSQL = strSql

End Function

Private Sub Test_GetPatientSQL()

    ModUtils.CopyToClipboard GetPatientListSQL(vbNullString, vbNullString, "H3.03.04")

End Sub

Public Sub MetaVision_GetPatientDetails(objPat As ClassPatientDetails, ByVal strPatId As String, ByVal strPatNum As String, ByVal strBed As String)

    Dim objRs As Recordset
    Dim strServer As String
    Dim strDatabase As String
    Dim dtmBd As Date
    Dim dtmAdm As Date
    Dim strDep As String
    Dim strSql As String
    
    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    
    If Not strPatNum = vbNullString Then
        strPatId = vbNullString
        strBed = vbNullString
    ElseIf Not strBed = vbNullString Then
        strPatId = vbNullString
    End If
    
    strSql = GetPatientListSQL(strPatId, strPatNum, strBed)
    
    If strServer = vbNullString Or strDatabase = vbNullString Or strSql = vbNullString Then Exit Sub
    
    objPat.Clear
    
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    If Not objRs.EOF Then
        If Not IsNull(objRs.Fields("BirthDate")) Then dtmBd = ModString.StringToDate(objRs.Fields("BirthDate"))
        objPat.PatientId = objRs.Fields("HospitalNumber")
        If Not IsNull(objRs.Fields("BedName")) Then objPat.Bed = Trim(objRs.Fields("BedName"))
        objPat.AchterNaam = objRs.Fields("LastName")
        objPat.VoorNaam = objRs.Fields("FirstName")
        If Not IsNull(objRs.Fields("WeightKg")) Then objPat.Gewicht = ModString.FixPrecision(objRs.Fields("WeightKg"), 2)
        If Not IsNull(objRs.Fields("WeightGr")) Then objPat.Gewicht = ModString.FixPrecision(objRs.Fields("WeightGr") / 1000, 2)
        If Not IsNull(objRs.Fields("LengthCm")) Then objPat.Lengte = Round(objRs.Fields("LengthCm"), 0)
        If Not IsNull(objRs.Fields("Geslacht")) Then objPat.Geslacht = objRs.Fields("Geslacht")
        If Not IsNull(objRs.Fields("BirthWeightGr")) Then objPat.GeboorteGewicht = objRs.Fields("BirthWeightGr")
        If Not IsNull(objRs.Fields("PregnDays")) Then objPat.Days = objRs.Fields("PregnDays")
        If Not IsNull(objRs.Fields("PregnWeeks")) Then objPat.Weeks = objRs.Fields("PregnWeeks")
    
        dtmAdm = ModString.StringToDate(objRs.Fields("LocationFromTime"))
        strDep = objRs.Fields("Department")
        strPatId = objRs.Fields("PatientId")
        
        Do While Not objRs.EOF
            ' MsgBox strDep & ": " & objRs.Fields("Department") & ", " & strPatId & ": " & objRs.Fields("PatientId")
            If strDep = objRs.Fields("Department") And strPatId = objRs.Fields("PatientId") Then
                dtmAdm = ModString.StringToDate(objRs.Fields("LocationFromTime"))
            Else
                Exit Do
            End If
            objRs.MoveNext
        Loop
        
        objPat.SetAdmissionAndBirthDate dtmAdm, dtmBd
        
        If objPat.Gewicht * 1000 < objPat.GeboorteGewicht Then objPat.Gewicht = objPat.GeboorteGewicht / 1000
        
    End If
    
    
    objConn.Close
    
End Sub

Private Sub Test_MetaVision_GetPatientDetails()
    
    Dim objPat As ClassPatientDetails
    Dim strId As String

    ' strId = MetaVision_GetCurrentPatientID()
    strId = vbNullString
    Set objPat = New ClassPatientDetails
    MetaVision_GetPatientDetails objPat, strId, "1234567", vbNullString

    MsgBox objPat.PatientId & ": " & objPat.AchterNaam

End Sub

Public Function MetaVision_GetCurrentBedName() As String

    Dim strPatId As String
    
    strPatId = MetaVision_GetCurrentPatientID()
    
    MetaVision_GetCurrentBedName = GetPatientBed(strPatId, vbNullString)

End Function

Private Sub Test_MetaVision_GetCurrentBedName()

    MsgBox MetaVision_GetCurrentBedName()

End Sub

Public Function MetaVision_GetCurrentPatientID() As String

    Dim strKeyPath As String
    Dim strValue As String
    Dim strBasePath As String
    
    strBasePath = GetBasePath()
    strKeyPath = IIf(strBasePath = constBasePath1, strBasePath, strBasePath & constSettings)
    strValue = IIf(strBasePath = constBasePath1, constPatientId, constCurrentPatient)
    
    MetaVision_GetCurrentPatientID = ModRegistry.ReadRegistryKey(strKeyPath, strValue)

End Function

Private Sub Test_MetaVision_GetCurrentPatient()

    MsgBox MetaVision_GetCurrentPatientID()

End Sub

Public Function MetaVision_GetUserLogin() As String

    Dim strKeyPath As String
    Dim strValue As String
    Dim strBasePath As String
    
    strBasePath = GetBasePath()
    strKeyPath = IIf(strBasePath = constBasePath1, strBasePath, strBasePath & constSettings)
    strValue = IIf(strBasePath = constBasePath1, constUserLogin, constUserId)
    
    MetaVision_GetUserLogin = ModRegistry.ReadRegistryKey(strKeyPath, strValue)

End Function

Private Sub Test_MetaVision_GetUserLogin()

    MsgBox MetaVision_GetUserLogin()

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

Private Function GetBasePath() As String

    Dim strBasePath As String
    
    strBasePath = IIf(RegistryKeyExists(constBasePath1, vbNullString), constBasePath1, constBasePath2)
    
    If strBasePath = vbNullString Then
        ModLog.LogError "No Valid Registry BasePath"
    End If
    
    GetBasePath = strBasePath

End Function

Private Sub Test_BasePath()

    MsgBox GetBasePath()

End Sub

Public Function MetaVision_GetDatabase() As String

    Dim strServer As String
    Dim strEmpi As String
    Dim strDepartment As String
    Dim strSql As String
    Dim strDb As String
    Dim objRs As Recordset
    
    If GetBasePath() = constBasePath1 Then
        
        strDb = ModRegistry.ReadRegistryKey(constBasePath1, constDatabase)
    
    Else
        
        strServer = MetaVision_GetServer()
        strEmpi = GetEmpiDb()
        strDepartment = MetaVision_GetDepartment()
        strSql = GetDatabaseNameSQL(strDepartment)
        
        If strServer <> vbNullString And strEmpi <> vbNullString Then
            InitConnection strServer, strEmpi
            
            If Not objConn Is Nothing Then
                objConn.Open
                
                Set objRs = objConn.Execute(strSql)
                If Not objRs.EOF Then strDb = objRs.Fields("DatabaseName")
                
                objConn.Close
            End If
        End If
    
    End If
    
    MetaVision_GetDatabase = strDb

End Function

Private Sub Test_MetaVision_GetDatabase()

    MsgBox MetaVision_GetDatabase()

End Sub

Private Function GetEmpiDb() As String

    Dim strKeyPath As String
    
    strKeyPath = GetBasePath() & constConnection
    
    GetEmpiDb = ModRegistry.ReadRegistryKey(strKeyPath, constEMPIDb)

End Function

Public Function MetaVision_GetServer() As String

    Dim strKeyPath As String
    Dim strValue As String
    Dim strBasePath As String
    
    strBasePath = GetBasePath()
    strKeyPath = IIf(strBasePath = constBasePath1, strBasePath, strBasePath & constConnection)
    strValue = IIf(strBasePath = constBasePath1, constServer, constEmpiServer)
    
    MetaVision_GetServer = ModRegistry.ReadRegistryKey(strKeyPath, strValue)

End Function

Private Sub Test_MetaVision_GetServer()

    MsgBox MetaVision_GetServer()

End Sub

Public Function MetaVision_GetDepartment() As String

    Dim strKeyPath As String
    Dim strValue As String
    Dim strBasePath As String
    
    strBasePath = GetBasePath()
    strKeyPath = IIf(strBasePath = constBasePath1, strBasePath, strBasePath & constConnection)
    strValue = IIf(strBasePath = constBasePath1, constDepartment, constDomain)
    MetaVision_GetDepartment = ModRegistry.ReadRegistryKey(strKeyPath, strValue)

End Function

Public Function MetaVision_IsPediatrie() As Boolean

    MetaVision_IsPediatrie = Not MetaVision_GetDepartment() = "Neonatologie"

End Function

Public Function MetaVision_IsNeonatologie() As Boolean

    MetaVision_IsNeonatologie = Not MetaVision_IsPediatrie()

End Function

Private Sub Test_MetaVision_GetDepartment()

    MsgBox MetaVision_GetDepartment()
    MsgBox MetaVision_IsPediatrie()

End Sub

Private Sub GetLab(ByVal strHospNum As String)

    Dim objRs As Recordset
    Dim objRange As Range
    Dim objRow As Range
    Dim strRow As String
    Dim strSql As String
    Dim strServer As String
    Dim strDatabase As String

    strSql = strSql & "DECLARE @HospNum AS NVARCHAR(40)" & vbNewLine
    strSql = strSql & "SET @HospNum = '" & strHospNum & "'" & vbNewLine
    strSql = strSql & "SELECT" & vbNewLine
    strSql = strSql & "p.Abbreviation, s.[Time], s.Value / u.Multiplier AS Value, u.UnitName" & vbNewLine
    strSql = strSql & "FROM Signals s" & vbNewLine
    strSql = strSql & "INNER JOIN Parameters p ON p.ParameterID = s.ParameterID" & vbNewLine
    strSql = strSql & "INNER JOIN Units u ON u.UnitID = p.UnitID" & vbNewLine
    strSql = strSql & "INNER JOIN Patients pat ON pat.PatientID = s.PatientID" & vbNewLine
    strSql = strSql & "WHERE pat.HospitalNumber = @HospNum" & vbNewLine
    strSql = strSql & "AND s.Error = 0" & vbNewLine
    strSql = strSql & "AND Datediff(day, s.[Time], GetDate()) <= 1" & vbNewLine
    strSql = strSql & "AND p.ParameterID IN (4199, 4148, 4217, 4136, 4137, 4138, 4142, 4143, 4144, 4263)" & vbNewLine
    strSql = strSql & "ORDER BY p.Abbreviation, s.[Time] DESC" & vbNewLine

    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    Set objRange = shtGlobBerLab.Range("Tbl_Glob_Lab")
    
    If Not (objRs.BOF And objRs.EOF) Then
        For Each objRow In objRange.Rows
            objRs.MoveFirst
            strRow = objRow.Cells(1, 1).Value2
            Do While Not objRs.EOF
                If strRow = objRs.Fields("Abbreviation").Value Then
                    objRow.Cells(1, 2).Value2 = objRs.Fields("Value").Value & " " & objRs.Fields("UnitName").Value
                    Exit Do
                End If
                objRs.MoveNext
            Loop
        Next
    End If
    
    objConn.Close


End Sub

Private Sub Test_GetLab()

    GetLab "8280506"

End Sub

Public Sub GetLeverNierFunctie(ByVal strHospNum As String)

    Dim objRs As Recordset
    Dim objRange As Range
    Dim objRow As Range
    Dim strRow As String
    Dim strSql As String
    Dim strServer As String
    Dim strDatabase As String

    strSql = strSql & "DECLARE @HospNum AS NVARCHAR(40)" & vbNewLine
    strSql = strSql & "SET @HospNum = '" & strHospNum & "'" & vbNewLine
    strSql = strSql & "SELECT" & vbNewLine
    strSql = strSql & "p.Abbreviation, s.[Time], s.Value / u.Multiplier AS Value, u.UnitName" & vbNewLine
    strSql = strSql & "FROM Signals s" & vbNewLine
    strSql = strSql & "INNER JOIN Parameters p ON p.ParameterID = s.ParameterID" & vbNewLine
    strSql = strSql & "INNER JOIN Units u ON u.UnitID = p.UnitID" & vbNewLine
    strSql = strSql & "INNER JOIN Patients pat ON pat.PatientID = s.PatientID" & vbNewLine
    strSql = strSql & "WHERE pat.HospitalNumber = @HospNum" & vbNewLine
    strSql = strSql & "AND s.Error = 0" & vbNewLine
    strSql = strSql & "AND Datediff(month, s.[Time], GetDate()) <= 1" & vbNewLine
    strSql = strSql & "AND p.ParameterID IN (4181, 4182, 4183, 4156)" & vbNewLine
    strSql = strSql & "ORDER BY p.Abbreviation, s.[Time] DESC" & vbNewLine

    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    Set objRange = shtGlobBerLab.Range("Tbl_Glob_Lab")
    
    If Not (objRs.BOF And objRs.EOF) Then
        For Each objRow In objRange.Rows
            objRs.MoveFirst
            strRow = objRow.Cells(1, 1).Value2
            Do While Not objRs.EOF
                If strRow = objRs.Fields("Abbreviation").Value Then
                    objRow.Cells(1, 2).Value2 = objRs.Fields("Value").Value & " " & objRs.Fields("UnitName").Value
                    Exit Do
                End If
                objRs.MoveNext
            Loop
        Next
    End If
    
    objConn.Close
    
End Sub

Public Sub MetaVision_SyncLab()

    Dim strHospNum As String
    Dim objRange As Range
    Dim objRow As Range
    
    Set objRange = shtGlobBerLab.Range("Tbl_Glob_Lab")
    
    For Each objRow In objRange
        objRow.Cells(1, 2).Value2 = vbNullString
    Next
    
    strHospNum = ModPatient.PatientHospNum()
    GetLab strHospNum
    GetLeverNierFunctie strHospNum
    
End Sub

Public Sub MetaVision_GetMedicatieOpdrachten()

    Dim strSql As String
    Dim objRs As Recordset
    Dim strServer As String
    Dim strDatabase As String
    Dim intN As Integer
    Dim intC As Integer
    Dim strMO As String
    Dim objRange As Range
    
    On Error GoTo GetMedicatieOpdrachtenError
    
    ModProgress.StartProgress "Medicatie opdrachten uit MetaVision ophalen"

    strSql = strSql & "SELECT p.ParameterName MO" & vbNewLine
    strSql = strSql & "FROM Parameters p " & vbNewLine
    strSql = strSql & "INNER JOIN ParametersCategories pc ON pc.CategoryID = p.CategoryID" & vbNewLine
    strSql = strSql & "WHERE pc.CategoryName = 'Opdr Medicatie Taken'" & vbNewLine
    strSql = strSql & "ORDER BY p.ParameterName" & vbNewLine

    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    intN = 1
    intC = shtGlobTblMedOpdr.Range("A1").CurrentRegion.Rows.Count
    intC = IIf(intC < 10, 900, intC)
    If Not (objRs.BOF And objRs.EOF) Then
        Do While Not objRs.EOF
            strMO = CStr(objRs.Fields("MO"))
            shtGlobTblMedOpdr.Cells(intN, 1).Value2 = strMO
            intN = intN + 1
            ModProgress.SetJobPercentage strMO, intC, intN
            objRs.MoveNext
        Loop
    End If
    
    Set objRange = shtGlobTblMedOpdr.Range("A1").CurrentRegion
    If ModRange.NameExists(constTblMedOpdr) Then WbkAfspraken.Names(constTblMedOpdr).Delete
    objRange.Name = constTblMedOpdr
    
    ModProgress.FinishProgress
    
    objConn.Close
    Set objConn = Nothing
    
    Exit Sub
    
GetMedicatieOpdrachtenError:

    On Error Resume Next
    
    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxError "Kan medicatie opdrachten niet ophalen"
    
    objConn.Close
    Set objConn = Nothing

End Sub

Public Sub MetaVision_SetUser()

    Dim strSql As String
    Dim strLogin As String
    
    Dim objRs As Recordset
    Dim strServer As String
    Dim strDatabase As String
    
    On Error GoTo SetUser_Error
    
    strLogin = MetaVision_GetUserLogin()

    If strLogin = vbNullString Then Exit Sub
    
    strSql = strSql & "DECLARE @login AS NVARCHAR(255)" & vbNewLine
    strSql = strSql & vbNullString & vbNewLine
    strSql = strSql & "SET @login = '" & strLogin & "' " & vbNewLine
    strSql = strSql & vbNullString & vbNewLine
    strSql = strSql & "SELECT" & vbNewLine
    strSql = strSql & "u.UserID" & vbNewLine
    strSql = strSql & ", u.Login" & vbNewLine
    strSql = strSql & ", u.FirstName" & vbNewLine
    strSql = strSql & ", u.LastName" & vbNewLine
    strSql = strSql & ", ut.UserTypeName" & vbNewLine
    strSql = strSql & "FROM Users u" & vbNewLine
    strSql = strSql & "INNER JOIN t_UsersType ut ON u.UserTypeID = ut.UserTypeID" & vbNewLine
    strSql = strSql & "WHERE u.Login = @login" & vbNewLine
    
    ModUtils.CopyToClipboard strSql
    
    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)

    ModRange.SetRangeValue "_User_Login", strLogin
    If Not objRs.EOF Then
        ModRange.SetRangeValue "_User_FirstName", objRs.Fields("FirstName")
        ModRange.SetRangeValue "_User_LastName", objRs.Fields("LastName")
        ModRange.SetRangeValue "_User_Type", objRs.Fields("UserTypeName")
    End If
    
    objConn.Close
    Set objConn = Nothing
    
    Exit Sub

SetUser_Error:

    ModLog.LogError "SetUser Error"
    
    On Error Resume Next
    
    objConn.Close
    Set objConn = Nothing

End Sub

Private Sub Test_RangeAdress()

    Dim objRange As Range
    
    On Error GoTo GetMedicatieOpdrachtenError
    
    
    Set objRange = shtGlobTblMedOpdr.Range("A1").CurrentRegion
    If ModRange.NameExists(constTblMedOpdr) Then WbkAfspraken.Names(constTblMedOpdr).Delete
    objRange.Name = constTblMedOpdr
    
    Exit Sub
    
GetMedicatieOpdrachtenError:

    On Error Resume Next
    
    
    ModMessage.ShowMsgBoxError "Kan medicatie opdrachten niet ophalen"
End Sub


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
Private Const CONST_PATHOSPNUM_RANGE As String = "__0_PatNum"

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

Public Function MetaVision_GetPatientBed(ByVal strPatId As String, ByVal strPatNum As String) As String

    Dim strServer As String
    Dim strDB As String
    Dim objRs As Recordset
    Dim strSql As String
    Dim strBed As String
    Dim strId As String
    Dim blnFound As Boolean
    
    strServer = MetaVision_GetServer()
    strDB = MetaVision_GetDatabase()
    strSql = GetPatientListSql(strPatId, strPatNum)
    
    If strServer = vbNullString Or strDB = vbNullString Or strSql = vbNullString Then
        strBed = vbNullString
    Else
        InitConnection strServer, strDB
        
        objConn.Open
        
        Set objRs = objConn.Execute(strSql)
        
        blnFound = False
        Do While Not objRs.EOF And Not blnFound
            strId = objRs.Fields("PatientId")
            If strId = strPatId Then
                If Not IsNull(objRs.Fields("BedName")) Then
                    strBed = Trim(CStr(objRs.Fields("BedName")))
                Else
                    strBed = ""
                End If
                blnFound = True
            End If
            objRs.MoveNext
        Loop
        
        objConn.Close
        Set objRs = Nothing
        
    End If
    
    MetaVision_GetPatientBed = strBed
    
    
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


Private Function GetPatientListSql(ByVal strPatId As String, ByVal strPatNum As String) As String

    Dim strSql As String
    Dim objBuilder As ClassStringBuilder
    
    Set objBuilder = New ClassStringBuilder
    
    objBuilder.Append "DECLARE @patId AS int" & vbNewLine
    objBuilder.Append "DECLARE @bed AS nvarchar(100)" & vbNewLine
    objBuilder.Append "DECLARE @dep AS nvarchar(60)" & vbNewLine
    objBuilder.Append "DECLARE @patNum AS nvarchar(40)" & vbNewLine
    objBuilder.Append "DECLARE @bd AS int" & vbNewLine
    objBuilder.Append "DECLARE @weightKg AS int" & vbNewLine
    objBuilder.Append "DECLARE @weightGr AS int" & vbNewLine
    objBuilder.Append "DECLARE @bwGr AS int" & vbNewLine
    objBuilder.Append "DECLARE @length AS int" & vbNewLine
    objBuilder.Append "DECLARE @gesl AS int" & vbNewLine
    objBuilder.Append "DECLARE @adD AS int" & vbNewLine
    objBuilder.Append "DECLARE @adW AS int" & vbNewLine
    
    objBuilder.Append "SET @bd = 5372" & vbNewLine
    objBuilder.Append "SET @weightKg = 8365" & vbNewLine
    objBuilder.Append "SET @weightGr = 8456" & vbNewLine
    objBuilder.Append "SET @bwGr = 7734" & vbNewLine
    objBuilder.Append "SET @adD = 10213" & vbNewLine
    objBuilder.Append "SET @adW = 10214" & vbNewLine
    objBuilder.Append "SET @length = 9505" & vbNewLine
    objBuilder.Append "SET @gesl = 5373" & vbNewLine
    
    If Not strPatNum = vbNullString Then
        objBuilder.Append "SET @patNum = '" & strPatNum & "'" & vbNewLine
    ElseIf Not strPatId = vbNullString And Not strPatId = "-1" Then
        objBuilder.Append "SET @patId = " & strPatId & vbNewLine
    End If
    
    objBuilder.Append "SELECT DISTINCT" & vbNewLine
    objBuilder.Append "pl.PatientID" & vbNewLine
    objBuilder.Append ", pat.HospitalNumber" & vbNewLine
    objBuilder.Append ", pl.LastName" & vbNewLine
    objBuilder.Append ", pl.FirstName" & vbNewLine
    objBuilder.Append ", (SELECT TOP 1 dts.Value" & vbNewLine
    objBuilder.Append "   FROM DateTimeSignals dts" & vbNewLine
    objBuilder.Append "   WHERE dts.PatientID = pl.PatientID AND dts.ParameterID = @bd" & vbNewLine
    objBuilder.Append "   ORDER BY dts.[Time] DESC)" & vbNewLine
    objBuilder.Append "   BirthDate" & vbNewLine
    objBuilder.Append ", (SELECT TOP 1 s.value / 1000 " & vbNewLine
    objBuilder.Append "   FROM Signals s " & vbNewLine
    objBuilder.Append "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @weightKg " & vbNewLine
    objBuilder.Append "   ORDER BY s.Time DESC) WeightKg" & vbNewLine
    objBuilder.Append ", (SELECT TOP 1 s.Value " & vbNewLine
    objBuilder.Append "   FROM Signals s " & vbNewLine
    objBuilder.Append "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @weightGr " & vbNewLine
    objBuilder.Append "   ORDER BY s.Time DESC) WeightGr" & vbNewLine
    objBuilder.Append ", (SELECT TOP 1 s.value * 100 " & vbNewLine
    objBuilder.Append "   FROM Signals s " & vbNewLine
    objBuilder.Append "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @length " & vbNewLine
    objBuilder.Append "   ORDER BY s.Time DESC) LengthCm" & vbNewLine
    objBuilder.Append ", (SELECT TOP 1 s.Value " & vbNewLine
    objBuilder.Append "   FROM Signals s " & vbNewLine
    objBuilder.Append "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @bwGr " & vbNewLine
    objBuilder.Append "   ORDER BY s.Time DESC) BirthWeightGr" & vbNewLine
    objBuilder.Append ", (SELECT TOP 1 s.value / (60 * 24) " & vbNewLine
    objBuilder.Append "   FROM Signals s " & vbNewLine
    objBuilder.Append "   WHERE s.PatientID = pl.PatientID AND s.ParameterID = @adD ORDER BY s.Time DESC) PregnDays" & vbNewLine
    objBuilder.Append ", (SELECT TOP 1 s.value / (7 * 60 * 24) FROM Signals s WHERE s.PatientID = pl.PatientID AND s.ParameterID = @adW " & vbNewLine
    objBuilder.Append "   ORDER BY s.Time DESC) PregnWeeks" & vbNewLine
    objBuilder.Append ", (SELECT TOP 1 pt.Text " & vbNewLine
    objBuilder.Append "   FROM ParametersText pt " & vbNewLine
    objBuilder.Append "   INNER JOIN TextSignals ts ON pt.ParameterID = ts.ParameterID AND pt.TextID = ts.TextID " & vbNewLine
    objBuilder.Append "   INNER JOIN Parameters p ON p.ParameterID = ts.ParameterID" & vbNewLine
    objBuilder.Append "   WHERE ts.PatientID = pl.PatientID AND p.ParameterID = @Gesl" & vbNewLine
    objBuilder.Append "   ORDER BY ts.Time DESC) Geslacht" & vbNewLine
    objBuilder.Append ", lu.Name Department" & vbNewLine
    objBuilder.Append ", b.BedName" & vbNewLine
    objBuilder.Append ", pl.LocationFromTime" & vbNewLine
    objBuilder.Append ", pl.TimeLog" & vbNewLine
    objBuilder.Append "FROM PatientLogs pl" & vbNewLine
    objBuilder.Append "INNER JOIN Patients pat ON pat.PatientID = pl.PatientID" & vbNewLine
    objBuilder.Append "LEFT JOIN LogicalUnits lu ON lu.LogicalUnitID = pl.LogicalUnitID" & vbNewLine
    objBuilder.Append "LEFT JOIN Beds b ON b.BedID = pl.BedID" & vbNewLine
    objBuilder.Append "WHERE " & vbNewLine
    objBuilder.Append "(@patId IS NULL OR pl.PatientID = @patId)" & vbNewLine
    objBuilder.Append "AND (@patNum IS NULL OR pl.HospitalNumber = @patNum)" & vbNewLine
    objBuilder.Append "AND (@bed IS NULL OR RTRIM(LTRIM(b.BedName)) = RTRIM(LTRIM(@bed)))" & vbNewLine
    objBuilder.Append "ORDER BY pat.HospitalNumber, pl.TimeLog DESC" & vbNewLine
    
    strSql = objBuilder.ToString()
    GetPatientListSql = strSql

End Function

Private Sub Test_GetPatientSQL()

    ModUtils.CopyToClipboard GetPatientListSql(vbNullString, vbNullString)

End Sub

Public Sub MetaVision_GetPatientDetails(objPat As ClassPatientDetails, ByVal strPatId As String, ByVal strPatNum As String)

    Dim objRs As Recordset
    Dim strServer As String
    Dim strDatabase As String
    Dim dtmBD As Date
    Dim dtmAdm As Date
    Dim strDep As String
    Dim strSql As String
    
    On Error GoTo ErrorHandler
    
    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    
    If Not strPatNum = vbNullString Then
        strPatId = vbNullString
    ElseIf Not strPatId = vbNullString Then
        strPatNum = vbNullString
    End If
    
    strSql = GetPatientListSql(strPatId, strPatNum)
    
    If strServer = vbNullString Or strDatabase = vbNullString Or strSql = vbNullString Then Exit Sub
    
    objPat.Clear
    
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    If Not objRs.EOF Then
        If Not IsNull(objRs.Fields("BirthDate")) Then dtmBD = ModString.StringToDate(objRs.Fields("BirthDate"))
        objPat.HospitalNumber = objRs.Fields("HospitalNumber")
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
        
        objPat.SetAdmissionAndBirthDate dtmAdm, dtmBD
        
        If objPat.Gewicht * 1000 < objPat.GeboorteGewicht Then objPat.Gewicht = objPat.GeboorteGewicht / 1000
        
    End If
    
    objConn.Close
    
    Database_LogAction "Get MetaVision patient details", ModUser.User_GetCurrent().Login, ModPatient.Patient_GetHospitalNumber()
    
    Exit Sub

ErrorHandler:
    
    objConn.Close
    
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "MetaVision_GetPatientDetails with Sql: " & strSql
    
End Sub

Private Sub Test_MetaVision_GetPatientDetails()
    
    Dim objPat As ClassPatientDetails
    Dim strId As String

    ' strId = MetaVision_GetCurrentPatientID()
    strId = vbNullString
    Set objPat = New ClassPatientDetails
    MetaVision_GetPatientDetails objPat, strId, "1234567"

    MsgBox objPat.HospitalNumber & ": " & objPat.AchterNaam

End Sub

Public Sub MetaVision_GetPatientsForDepartment(objCol As Collection, ByVal strDep As String)

    Dim objRs As Recordset
    Dim objPat As ClassPatientDetails
    Dim strServer As String
    Dim strDatabase As String
    Dim dtmBD As Date
    Dim dtmAdm As Date
    Dim strSql As String
    Dim intN As Integer
    Dim blnDep As Boolean
    
    On Error GoTo HandleError:
    
    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
        
    strSql = GetPatientListSql(vbNullString, vbNullString)
    
    If strServer = vbNullString Or strDatabase = vbNullString Or strSql = vbNullString Then Exit Sub
        
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    Set objRs = objConn.Execute(strSql)
    
    Do While Not objRs.EOF
        Set objPat = New ClassPatientDetails
        intN = intN + 1
        blnDep = False
                
        objPat.HospitalNumber = objRs.Fields("HospitalNumber")
            
        If Not IsNull(objRs.Fields("BirthDate")) Then dtmBD = ModString.StringToDate(objRs.Fields("BirthDate"))
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
        
        If Not IsNull(objRs.Fields("Department")) Then objPat.Afdeling = objRs.Fields("Department")
        
        objPat.SetAdmissionAndBirthDate dtmAdm, dtmBD
        
        If objPat.Gewicht * 1000 < objPat.GeboorteGewicht Then objPat.Gewicht = objPat.GeboorteGewicht / 1000
        
        Do While Not objRs.EOF
            If objRs.Fields("HospitalNumber") = objPat.HospitalNumber Then
                If Not IsNull(objRs.Fields("Department")) Then blnDep = blnDep Or strDep = objRs.Fields("Department")
                objRs.MoveNext
            Else
                Exit Do
            End If
        Loop
        
        If blnDep Then objCol.Add objPat
    Loop
    
    objConn.Close
    
    Exit Sub
    
HandleError:

    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not get patient no: " & intN & " with SQL : " & vbNewLine & strSql
    objConn.Close
    
End Sub

Private Sub Test_MetaVision_GetPatientsForDepartment()

    Dim objCol As Collection
    Dim objPat As ClassPatientDetails
    
    Set objCol = New Collection
    MetaVision_GetPatientsForDepartment objCol, "Pediatrie"
    
    For Each objPat In objCol
        If Not objPat.Bed = vbNullString Then
            ModMessage.ShowMsgBoxInfo objPat.HospitalNumber & ": " & objPat.AchterNaam & ", " & objPat.VoorNaam
        End If
    Next
    
    Set objPat = Nothing
    Set objCol = Nothing

End Sub

Public Function MetaVision_GetCurrentBedName() As String

    Dim strPatId As String
    
    strPatId = MetaVision_GetCurrentPatientID()
    
    MetaVision_GetCurrentBedName = MetaVision_GetPatientBed(strPatId, vbNullString)

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

Private Function GetBasePath() As String

    Dim strBasePath As String
    
    strBasePath = IIf(RegistryKeyExists(constBasePath1, vbNullString), constBasePath1, constBasePath2)
    
    If strBasePath = vbNullString Then
        ModLog.LogError Err, "No Valid Registry BasePath"
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
    Dim strDB As String
    Dim objRs As Recordset
    
    If GetBasePath() = constBasePath1 Then
        
        strDB = ModRegistry.ReadRegistryKey(constBasePath1, constDatabase)
    
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
                If Not objRs.EOF Then strDB = objRs.Fields("DatabaseName")
                
                objConn.Close
            End If
        End If
    
    End If
    
    MetaVision_GetDatabase = strDB

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

Public Function MetaVision_IsPICU() As Boolean

    MetaVision_IsPICU = Not MetaVision_GetDepartment() = "Neonatologie"

End Function

Public Function MetaVision_IsNeonatologie() As Boolean

    MetaVision_IsNeonatologie = Not MetaVision_IsPICU()

End Function

Private Sub Test_MetaVision_GetDepartment()

    MsgBox MetaVision_GetDepartment()
    MsgBox MetaVision_IsPICU()

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

Private Function GetLatestTextSignalInPeriod(ByVal intPatId As Long, ByVal intParId As Long, Optional ByVal dtmFrom As Date, Optional ByVal dtmTo As Date) As ClassSignal

    Dim strFrom As String
    Dim strTo As String
    Dim strSql As String
    Dim strServer As String
    Dim strDatabase As String
    Dim objSignal As ClassSignal
    Dim objRs As Recordset
    
        
    strFrom = WrapDateTime(dtmFrom)
    strTo = WrapDateTime(dtmTo)
    
    strSql = strSql & "DECLARE @parId AS INTEGER" & vbNewLine
    strSql = strSql & "DECLARE @patId AS INTEGER" & vbNewLine
    strSql = strSql & "DECLARE @fromTime AS DATE" & vbNewLine
    strSql = strSql & "DECLARE @toTime AS DATE" & vbNewLine
    strSql = strSql & "" & vbNewLine
    strSql = strSql & "SET @parId = " & intParId & vbNewLine
    strSql = strSql & "SET @patId = " & intPatId & vbNewLine
        
    If Not strFrom = "" Then
        strSql = strSql & "SET @fromTime = " & strFrom & vbNewLine
    End If
    
    If Not strTo = "" Then
        strSql = strSql & "SET @toTime = " & strTo & vbNewLine
    End If
    
    strSql = strSql & "" & vbNewLine
    strSql = strSql & "SELECT TOP 1 " & vbNewLine
    strSql = strSql & "p.Abbreviation AS [Name]" & vbNewLine
    strSql = strSql & ", s.[Time] AS [Time]" & vbNewLine
    strSql = strSql & ", pt.[Text] AS [Value]" & vbNewLine
    strSql = strSql & "FROM TextSignals s" & vbNewLine
    strSql = strSql & "INNER JOIN [Parameters] p ON p.ParameterID = s.ParameterID" & vbNewLine
    strSql = strSql & "INNER JOIN ParametersText pt on pt.ParameterID = s.ParameterID AND pt.TextID = s.TextID" & vbNewLine
    strSql = strSql & "WHERE s.ParameterID = @parId AND" & vbNewLine
    strSql = strSql & "s.PatientID = @patId" & vbNewLine
    
    If Not strFrom = "" Then
        strSql = strSql & "AND" & vbNewLine
        strSql = strSql & "s.Time >= @fromTime" & vbNewLine
    End If
    
    If Not strTo = "" Then
        strSql = strSql & "AND" & vbNewLine
        strSql = strSql & "s.Time <= @toTime  " & vbNewLine
    End If
    
    strSql = strSql & "ORDER BY s.Time DESC" & vbNewLine
        
    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    ModUtils.CopyToClipboard strSql
    Set objRs = objConn.Execute(strSql)
    
    Set objSignal = New ClassSignal
    If Not (objRs.BOF And objRs.EOF) Then
        With objSignal
            .Name = objRs.Fields("Name").Value
            .Time = objRs.Fields("Time").Value
            .Value = objRs.Fields("Value").Value
        End With
    End If
    
    objConn.Close

    Set GetLatestTextSignalInPeriod = objSignal

End Function

Private Sub Test_GetLatestTextSignal()

    Dim dtmFrom As Date
    Dim dtmTo As Date
    Dim objSignal As ClassSignal
    
    dtmTo = Now()
    dtmFrom = DateAdd("yyyy", -5, dtmTo)

    Set objSignal = GetLatestTextSignalInPeriod(31583, 5373)
    ModMessage.ShowMsgBoxInfo objSignal.Name & ": " & objSignal.Value & " " & objSignal.Unit

End Sub

Private Function GetSumNumSignalInPeriod(ByVal intPatId As Long, ByVal strDescr As String, Optional ByVal dtmFrom As Date, Optional ByVal dtmTo As Date) As Double

    Dim strFrom As String
    Dim strTo As String
    Dim strSql As String
    Dim strServer As String
    Dim strDatabase As String
    Dim dblSum As Double
    Dim objRs As Recordset
    
    strFrom = WrapDateTime(dtmFrom)
    strTo = WrapDateTime(dtmTo)
    
    strSql = strSql & "DECLARE @descr AS NVARCHAR(MAX)" & vbNewLine
    strSql = strSql & "DECLARE @patId AS INTEGER" & vbNewLine
    strSql = strSql & "DECLARE @fromTime AS DATETIME" & vbNewLine
    strSql = strSql & "DECLARE @toTime AS DATETIME" & vbNewLine
    strSql = strSql & "" & vbNewLine
    strSql = strSql & "SET @descr = '" & strDescr & "'" & vbNewLine
    strSql = strSql & "SET @patId = " & intPatId & vbNewLine
        
    If Not strFrom = "" Then
        strSql = strSql & "SET @fromTime = " & strFrom & vbNewLine
    End If
    
    If Not strTo = "" Then
        strSql = strSql & "SET @toTime = " & strTo & vbNewLine
    End If
    
    strSql = strSql & "" & vbNewLine
    strSql = strSql & "SELECT  " & vbNewLine
    strSql = strSql & "SUM (s.Value) AS [Sum]" & vbNewLine
    strSql = strSql & "FROM Signals s" & vbNewLine
    strSql = strSql & "WHERE " & vbNewLine
    strSql = strSql & "s.ParameterID IN (SELECT s.ParameterID FROM [Parameters] s WHERE s.Description LIKE @descr)" & vbNewLine
    strSql = strSql & "AND s.PatientID = @patId" & vbNewLine
    
    If Not strFrom = "" Then
        strSql = strSql & "AND" & vbNewLine
        strSql = strSql & "s.Time >= @fromTime" & vbNewLine
    End If
    
    If Not strTo = "" Then
        strSql = strSql & "AND" & vbNewLine
        strSql = strSql & "s.Time <= @toTime  " & vbNewLine
    End If
            
    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    ModUtils.CopyToClipboard strSql
    Set objRs = objConn.Execute(strSql)
    
    If Not (objRs.BOF And objRs.EOF) Then
        If Not IsNull(objRs.Fields("Sum").Value) Then dblSum = objRs.Fields("Sum").Value
    End If
    
    objConn.Close

    GetSumNumSignalInPeriod = dblSum

End Function

Private Sub Test_GetSumNumSignalInPeriod()

    ModMessage.ShowMsgBoxInfo GetSumNumSignalInPeriod(31583, "<diurese>%")

End Sub

Private Function GetLatestNumSignalInPeriod(ByVal intPatId As Long, ByVal intParId As Long, Optional ByVal dtmFrom As Date, Optional ByVal dtmTo As Date) As ClassSignal

    Dim strFrom As String
    Dim strTo As String
    Dim strSql As String
    Dim strServer As String
    Dim strDatabase As String
    Dim objSignal As ClassSignal
    Dim objRs As Recordset
    
        
    strFrom = WrapDateTime(dtmFrom)
    strTo = WrapDateTime(dtmTo)
    
    strSql = strSql & "DECLARE @parId AS INTEGER" & vbNewLine
    strSql = strSql & "DECLARE @patId AS INTEGER" & vbNewLine
    strSql = strSql & "DECLARE @fromTime AS DATE" & vbNewLine
    strSql = strSql & "DECLARE @toTime AS DATE" & vbNewLine
    strSql = strSql & "" & vbNewLine
    strSql = strSql & "SET @parId = " & intParId & vbNewLine
    strSql = strSql & "SET @patId = " & intPatId & vbNewLine
        
    If Not strFrom = "" Then
        strSql = strSql & "SET @fromTime = " & strFrom & vbNewLine
    End If
    
    If Not strTo = "" Then
        strSql = strSql & "SET @toTime = " & strTo & vbNewLine
    End If
    
    strSql = strSql & "" & vbNewLine
    strSql = strSql & "SELECT TOP 1 " & vbNewLine
    strSql = strSql & "p.Abbreviation AS [Name]" & vbNewLine
    strSql = strSql & ", s.Time AS [Time]" & vbNewLine
    strSql = strSql & ", s.Value / u.Multiplier AS [Value]" & vbNewLine
    strSql = strSql & ", u.UnitName AS [Unit]" & vbNewLine
    strSql = strSql & "FROM Signals s" & vbNewLine
    strSql = strSql & "INNER JOIN [Parameters] p ON p.ParameterID = s.ParameterID" & vbNewLine
    strSql = strSql & "INNER JOIN Units u ON u.UnitID = p.UnitID" & vbNewLine
    strSql = strSql & "WHERE s.ParameterID = @parId AND" & vbNewLine
    strSql = strSql & "s.PatientID = @patId" & vbNewLine
    
    If Not strFrom = "" Then
        strSql = strSql & "AND" & vbNewLine
        strSql = strSql & "s.Time >= @fromTime" & vbNewLine
    End If
    
    If Not strTo = "" Then
        strSql = strSql & "AND" & vbNewLine
        strSql = strSql & "s.Time <= @toTime  " & vbNewLine
    End If
    
    strSql = strSql & "ORDER BY s.Time DESC" & vbNewLine
        
    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    ModUtils.CopyToClipboard strSql
    Set objRs = objConn.Execute(strSql)
    
    Set objSignal = New ClassSignal
    If Not (objRs.BOF And objRs.EOF) Then
        With objSignal
            .Name = objRs.Fields("Name").Value
            .Time = objRs.Fields("Time").Value
            .Value = objRs.Fields("Value").Value
            .Unit = objRs.Fields("Unit").Value
        End With
    End If
    
    objConn.Close

    Set GetLatestNumSignalInPeriod = objSignal

End Function

Private Sub Test_GetLatestNumSignal()

    Dim dtmFrom As Date
    Dim dtmTo As Date
    Dim objSignal As ClassSignal
    
    dtmTo = Now()
    dtmFrom = DateAdd("yyyy", -5, dtmTo)

    Set objSignal = GetLatestNumSignalInPeriod(31583, 5473)
    ModMessage.ShowMsgBoxInfo objSignal.Name & ": " & objSignal.Value & " " & objSignal.Unit

End Sub

Private Function WrapDateTime(dtmDate As Date) As String

    Dim strDate As String
    
    strDate = FormatDateTimeSeconds(dtmDate)
    If strDate = "" Then
        WrapDateTime = ""
    Else
        WrapDateTime = "{ts'" & strDate & "'}"
    End If
    
End Function

Private Sub Test_WrapDateTime()


    ModMessage.ShowMsgBoxInfo WrapDateTime(Now())

End Sub

Public Function MetaVision_eGFRWarning() As String

    Dim intPatId As Long

    Dim intCreatId As Integer
    Dim intHeightId As Integer
    Dim intWeightId As Integer
    Dim intGenderId As Integer
    
    Dim intDays As Long
    Dim intValidDays As Integer
    
    Dim dblHeight As Double
    Dim dblWeight As Double
    Dim strGender As String
    
    Dim strSql As String
    Dim strServer As String
    Dim strDatabase As String
    Dim objRs As Recordset
    
    Dim dtmFrom As Date
    Dim dtmTo As Date
    
    Dim blnAKI As Boolean
    Dim blnDiff As Boolean
    Dim blnCreat As Boolean
    Dim blnDiurese As Boolean
    Dim dblCreat As Double
    Dim intEGFR As Integer
    Dim dblVal As Double
    Dim dblMin As Double
    Dim dblMax As Double
    Dim dtmMin As Date
    Dim dtmMax As Date
    Dim dtmTime As Date
    Dim dtmVal As Date
    Dim dblDiurese As Double
    Dim strResult As String
    Dim strTime As String
    Dim dblGender As Double
    
    intCreatId = 4156  ' Creatinine (bl)
    intHeightId = 9505 ' Actuele Lengte (cm)
    intWeightId = 8365 ' Gewicht (kg)
    intGenderId = 5373 ' Patient geslacht 1 = Vrouw, 2 = Man, 3 = Onbekend

    intPatId = MetaVision_GetCurrentPatientID()

    ' Determine valid period to get the height
    intDays = DateDiff("d", Patient_BirthDate(), Now())
    intValidDays = 15
    If intDays > 90 Then intValidDays = 30
    If intDays > 365 Then intValidDays = 90

    ' Get the latest height in the valid period
    dtmFrom = DateAdd("d", Now(), -1 * intValidDays)
    dtmTo = Now()
    dblHeight = GetLatestNumSignalInPeriod(intPatId, intHeightId, dtmFrom, dtmTo).Value
    
    ' Get the latest weight in the valid period
    dtmFrom = DateAdd("d", Now(), -1 * intValidDays)
    dtmTo = Now()
    dblWeight = GetLatestNumSignalInPeriod(intPatId, intWeightId, dtmFrom, dtmTo).Value

    ' Get the latest known gender
    strGender = GetLatestTextSignalInPeriod(intPatId, intGenderId).Value

    'Get the latest creat lab values of the last 90 days
    strSql = strSql & "DECLARE @patId INT" & vbNewLine
    strSql = strSql & "DECLARE @parId INT" & vbNewLine
    strSql = strSql & "" & vbNewLine
    strSql = strSql & "SET @patId = " & intPatId & vbNewLine
    strSql = strSql & "SET @parId = " & intCreatId & vbNewLine
    strSql = strSql & "" & vbNewLine
    strSql = strSql & "SELECT s.[Time], s.Value / u.Multiplier Value FROM Signals s" & vbNewLine
    strSql = strSql & "INNER JOIN Parameters p ON p.ParameterID = s.ParameterID" & vbNewLine
    strSql = strSql & "INNER JOIN Units u ON u.UnitID = p.UnitID" & vbNewLine
    strSql = strSql & "WHERE s.PatientID = @patId" & vbNewLine
    strSql = strSql & "AND s.ParameterID = @parId" & vbNewLine
    strSql = strSql & "AND s.Error = 0" & vbNewLine
    strSql = strSql & "AND datediff(d, s.[Time], getdate()) <= 90" & vbNewLine
    strSql = strSql & "ORDER BY s.[Time] DESC" & vbNewLine

    strServer = MetaVision_GetServer()
    strDatabase = MetaVision_GetDatabase()
    
    InitConnection strServer, strDatabase
    
    objConn.Open
    
    ModUtils.CopyToClipboard strSql
    Set objRs = objConn.Execute(strSql)
        
    blnAKI = False
    dblCreat = 0
    intEGFR = 0
    dblVal = 0
    dblMin = 1
    dtmMin = Now()
    dtmMax = Now()
    dblMax = 0
    Do While Not objRs.EOF
        ' Values not yet initialized, set to the first creat value
        If dblCreat = 0 Then
            dtmTime = CDate(objRs.Fields("Time").Value)
            dblCreat = objRs.Fields("Value").Value
            dblMin = dblCreat
            dblMax = dblCreat
        End If
        ' AKI if the difference with the previous value > 26.5 microg/l
        If dblVal > 0 Then blnDiff = (objRs.Fields("Value").Value - dblVal) > 26.5
        dblVal = objRs.Fields("Value").Value
        ' Calculate min and max creat for the last 7  days
        dtmVal = objRs.Fields("Time").Value
        If DateDiff("d", dtmVal, Now()) <= 7 Then
            If dblMin > dblVal Then
                dblMin = dblVal
                dtmMin = dtmVal
            End If

            If dblMax < dblVal Then
                dblMax = dblVal
                dtmMax = dtmVal
            End If
        End If

        objRs.MoveNext
    Loop
    
    objConn.Close


    ' Determine if there is AKI
    If dtmMax > dtmMin Then blnCreat = (dblMax / dblMin) > 1.5
    If dblWeight > 0 Then
        dtmTo = Now()
        dtmFrom = DateAdd("h", -6, dtmTo)
        dblDiurese = GetSumNumSignalInPeriod(intPatId, "<diurese>%", dtmFrom, dtmTo)
        dblDiurese = dblDiurese / dblWeight
        
        If dblDiurese < 0.5 Then blnDiurese = True
    End If
    blnAKI = (blnAKI Or blnDiff)
    blnAKI = (blnAKI Or blnCreat)
    blnAKI = (blnAKI Or blnDiurese)

    ' Test
    ' intDays = 600
    ' dblCreat = 100
    ' dblHeight = 100
    ' eGFR = 36

    ' Test
    ' intDays = 18 * 365
    ' dblWeight = 50
    ' dblCreat = 120
    ' strGender = "Vrouw"
    ' eGFR =

    If intDays < 365 Then
        strResult = ""
    ElseIf dblWeight < 50 Then
        ' Calculate eGFR from the last known creat value
        ' The Schwartz formule:
        If dblCreat > 0 Then intEGFR = CInt(36.2 * dblHeight / (dblCreat))

        strTime = FormatDateTime(dtmTime, vbShortDate)
        strResult = ""
        If intEGFR > 0 Then
            If intEGFR < 60 Then strResult = "De eGFR = " & intEGFR & " ml/min/1,73 m2  (" & strTime & "), beperkte nierfunctie!"
            If intEGFR < 50 Then strResult = "De eGFR = " & intEGFR & " ml/min/1,73 m2  (" & strTime & "), denk aan evt. dosering aanpassingen!"
            If dblHeight = 0 And dblCreat > 0 And intEGFR = 0 Then strResult = "Geen lengte bekend, kan eGFR niet berekenen. Graag actuele lengte invoeren."
        End If
    Else
        ' Calculate eGFR from the last known creat value
        ' The MDRD formule:
        If strGender = "Vrouw" Then dblGender = 0.742
        If strGender = "Man" Then dblGender = 1
        If dblCreat > 0 Then intEGFR = CInt((175 * ((dblCreat / 88.4) ^ (-1.154))) * ((intDays / 365) ^ -0.203) * (dblGender))
        
        If intEGFR > 0 Then
            strTime = FormatDateTime(dtmTime, vbShortDate)
            If intEGFR < 60 And intEGFR > 0 Then strResult = "De eGFR = " & intEGFR & " ml/min/1,73 m2  (" & strTime & "), beperkte nierfunctie!"
            If intEGFR < 50 Then strResult = "De eGFR = " & intEGFR & " ml/min/1,73 m2  (" & strTime & "), denk aan evt. dosering aanpassingen!"
        End If
    End If


    If blnAKI Then
        If strResult <> "" Then strResult = strResult & " "
        If blnDiff Then strResult = strResult & "Stijging van creatinine waarden van > 26,5 microm/l. "
        If blnCreat Then strResult = strResult & "Stijging van creatinine waarden van > 1,5 keer van de voorgaande waarde. "
        If blnDiurese Then strResult = strResult & "Verminderde diurese van < 0.5 ml/kg/uur. "
        strResult = strResult & vbNewLine & "Patient heeft mogelijk Acute Kidney Injury."
    End If

    MetaVision_eGFRWarning = strResult

End Function

Private Sub Test_MetaVision_eGFRWarning()

    ModMessage.ShowMsgBoxInfo MetaVision_eGFRWarning()

End Sub

Public Sub MetaVision_SyncLab()

    Dim strHospNum As String
    Dim objRange As Range
    Dim objRow As Range
    
    Set objRange = shtGlobBerLab.Range("Tbl_Glob_Lab")
    
    For Each objRow In objRange
        objRow.Cells(1, 2).Value2 = vbNullString
    Next
    
    strHospNum = ModPatient.Patient_GetHospitalNumber()
    GetLab strHospNum
    GetLeverNierFunctie strHospNum
    
    ModRange.SetRangeValue "_Glob_Lab_eGFR", MetaVision_eGFRWarning()
    
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
    Dim objUser As ClassUser
    
    On Error GoTo SetUser_Error
    
    strLogin = MetaVision_GetUserLogin()
    Set objUser = New ClassUser

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

    With objUser
        .Login = strLogin
        If Not objRs.EOF Then
            .FirstName = objRs.Fields("FirstName")
            .LastName = objRs.Fields("LastName")
            .Role = objRs.Fields("UserTypeName")
        End If
    End With
    ModUser.User_SetUser objUser
    
    objConn.Close
    Set objConn = Nothing
    
    Exit Sub

SetUser_Error:

    ModLog.LogError Err, "SetUser Error"
    
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


Private Sub InitConnection(ByVal strServer As String, ByVal strDatabase As String)

    Dim strSecret As String
    Dim strUser As String
    Dim strPw As String
    Dim strMsg As String
    
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
        ModLog.LogError Err, "Bestand secret niet aanwezig"
    End If
    
    Exit Sub
    
InitConnectionError:

    strMsg = "Geen toegang tot " & strServer & ": " & strDatabase & "!"
    ModMessage.ShowMsgBoxExclam strMsg
    ModLog.LogError Err, "InitConnection Failed: " & vbNewLine & strMsg

End Sub


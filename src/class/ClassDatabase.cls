VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objConn As ADODB.Connection
Attribute objConn.VB_VarHelpID = -1

Private strTS As String
Private strPrev As String

Private Const constSecret As String = "secret"

Public Sub InitConnection(ByVal strServer As String, ByVal strDatabase As String)

    Dim strSecret As String
    Dim strUser As String
    Dim strPw As String
    Dim strConn As String
    
    On Error GoTo InitConnectionError
        
    If (Not objConn Is Nothing) Then Exit Sub
    
    strSecret = ModFile.ReadFile(WbkAfspraken.Path & "/" & constSecret)
        
    If strSecret <> vbNullString Then
        strUser = Split(strSecret, vbLf)(0)
        strPw = Split(strSecret, vbLf)(1)
    
        
        strConn = "Provider=SQLOLEDB.1;" _
                 & "Server=" & strServer & ";" _
                 & "Database=" & strDatabase & ";" _
                 & "User ID=" & strUser & ";" _
                 & "Password=" & strPw & ";" _
                 & "DataTypeCompatibility=80;" _
                 & "MARS Connection=True;"
        ' Test de connectie
        
        Set objConn = New ADODB.Connection
        objConn.ConnectionString = strConn
        
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

Private Function ContainsData(varVal As Variant) As Boolean

    ContainsData = Not (Format(varVal) = vbNullString Or Format(varVal) = "0")

End Function

Public Sub SaveData(strTimeStamp As String, strHospNum, strPrescriber As String, objData As Range, objText As Range, blnProgress As Boolean)

    Dim strParam As String
    Dim varVal As Variant
    Dim varEmp As Variant
    Dim strSql As String
    
    Dim intC As Integer
    Dim intN As Integer
        
    strTS = strTimeStamp
    strPrev = ModBed.GetDatabaseVersie()
            
    On Error GoTo SaveDataError
       
    intC = objData.Rows.Count
    For intN = 2 To intC
        strParam = objData.Cells(intN, 1).Value2
        varVal = objData.Cells(intN, 2).Value2
        varEmp = objData.Cells(intN, 3).Value2
        
        If Not varVal = varEmp Then
            strSql = strSql & vbNewLine & "EXEC dbo.InsertPrescriptionData '" & strHospNum & "', {ts' " & strTimeStamp & " '}, '" & strPrescriber & "', 0, ' " & strParam & " ', '" & varVal & " '"
        End If
        
        If blnProgress Then ModProgress.SetJobPercentage "Data wegschrijven", intC, intN
    Next intN
    
    strSql = strSql & vbNewLine
    strSql = strSql & vbNewLine
    
    intC = objText.Rows.Count
    For intN = 2 To intC
        If ContainsData(objText.Cells(intN, 2).Value2) Then
            strParam = objText.Cells(intN, 1).Value2
            varVal = objText.Cells(intN, 2).Value2
            strSql = strSql & vbNewLine & "EXEC dbo.InsertPrescriptionText '" & strHospNum & "', {ts' " & strTimeStamp & " '}, '" & strPrescriber & "', 0, ' " & strParam & " ', '" & varVal & " '"
        End If
        
        If blnProgress Then ModProgress.SetJobPercentage "Text wegschrijven naar de database", intC, intN
    Next intN
    
    strSql = ModDatabase.WrapTransaction(strSql, "save_data")
        
    objConn.Open
    objConn.Execute strSql, adExecuteNoRecords
    objConn.Close
    
    ModBed.SetDatabaseVersie strTS
    
    Exit Sub

SaveDataError:
    
    ModMessage.ShowMsgBoxError "Kan patient data niet opslaan"
    ModUtils.CopyToClipboard strSql
    ModLog.LogError Err, "Could not save patient data to the database"
    
    objConn.Close
    
    ModBed.SetDatabaseVersie strPrev

End Sub

Private Function GetErrorMsg(objErr As ADODB.Error) As String

    Dim strMsg As String
    
    strMsg = "Number: " & objErr.Number & vbNewLine
    strMsg = strMsg & "Source: " & objErr.Source & vbNewLine
    strMsg = strMsg & "Description: " & objErr.Description
    strMsg = strMsg & "Error: " & objErr.NativeError & vbNewLine
    strMsg = strMsg & "SQL: " & objErr.SqlState
    
    GetErrorMsg = strMsg

End Function
'
'Private Sub objConn_ExecuteComplete(ByVal lngAffected As Long, ByVal objErr As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal objComm As ADODB.Command, ByVal objRs As ADODB.Recordset, ByVal objConn As ADODB.Connection)
'
'    If Not objErr Is Nothing Then
'        If objErr.Number <> 0 Then
'            ModMessage.ShowMsgBoxError "Kan data niet wegschrijven in de database"
'            ModLog.LogError Err,  "Could finish transaction to save data with error: " & vbNewLine & GetErrorMsg(objErr)
'
'            ModBed.SetDatabaseVersie strPrev
'        Else
'            ModLog.LogInfo "Saved patient to database: " & strTS
'            ModBed.SetDatabaseVersie strTS
'        End If
'
'    Else
'        ModLog.LogInfo "Saved patient to database: " & strTS
'        ModBed.SetDatabaseVersie strTS
'
'    End If
'
'    ModRange.SetRangeValue constBusy, False
'    objConn.Close
'
'End Sub

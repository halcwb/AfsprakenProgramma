Attribute VB_Name = "ModBed_Tests"
Option Explicit

Public Sub Test_CloseBed()

    Dim objDict As Dictionary
    Dim blnPass As Boolean
    
    Set objDict = New Dictionary
    blnPass = True
    
    On Error GoTo TestError
    
    ModProgress.StartProgress "Test Bed Opslaan"
        
    ModBed.OpenBed
    blnPass = blnPass And ModNeoInfB_Tests.Test_NeoInfB_FillContMed(blnPass)
    If Not blnPass Then
        Err.Raise CONST_TEST_ERROR, , "Test_CloseBed", "Medicatie vullen did not pass"
    End If
    
    FillDictWithAfspr objDict
    
    ModBed.CloseBed False
    ModBed.OpenBed
    
    blnPass = CheckAfsprDict(objDict, blnPass)
    
    ModProgress.FinishProgress
    
    ModMessage.ShowMsgBoxInfo "Alles OK"
    
    Exit Sub
    
TestError:

    ModProgress.FinishProgress
    ModMessage.ShowMsgBoxExclam "Bed opslaan test niet geslaagd" & vbNewLine & Err.Description

End Sub

Private Sub FillDictWithAfspr(objDict As Dictionary)

    Dim strJob As String
    Dim intC As Integer
    Dim intN As Integer
    Dim varValue As Variant
    Dim strKey As String

    With shtPatData
        strJob = "Patient gegevens kopieren"
        intC = .Range("A1").CurrentRegion.Rows.Count
        
        For intN = 2 To intC
            strKey = .Cells(intN, 1).Value2
            varValue = .Cells(intN, 2).Value2
            objDict.Add strKey, varValue
            
            ModProgress.SetJobPercentage strJob, intC, intN
        Next intN
    End With
    
End Sub

Private Function CheckAfsprDict(objDict As Dictionary, blnPass As Boolean) As Boolean

    Dim strJob As String
    Dim intC As Integer
    Dim intN As Integer
    Dim varValue As Variant
    Dim strKey As String
    
    With shtPatData
        strJob = "Patient gegevens checken"
        intC = .Range("A1").CurrentRegion.Rows.Count
        
        For intN = 2 To intC
            strKey = .Cells(intN, 1).Value2
            varValue = .Cells(intN, 2).Value2
            blnPass = blnPass And objDict.Item(strKey) = varValue
            
            If Not blnPass Then
                Err.Raise CONST_TEST_ERROR, , "Key: " & strKey & " Waarde: " & varValue & " <> " & objDict.Item(strKey)
            End If
            
            ModProgress.SetJobPercentage strJob, intC, intN
        Next intN
    End With
    
    CheckAfsprDict = blnPass

End Function

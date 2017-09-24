Attribute VB_Name = "ModExcel"
Option Explicit

Public Function Excel_VLookupExists(ByVal varValue As Variant, ByVal strTable As String) As Boolean
    
    On Error GoTo Excel_VLookupExistsError

    Excel_VLookupExists = Excel_VLookup(varValue, ByVal strTable, 1) = varValue
    
    Exit Function
    
Excel_VLookupExistsError:
    
    Excel_VLookupExists = False

End Function

Private Sub Test_Excel_VLookupExists()

    MsgBox Excel_VLookupExists("dopamine", "Tbl_Neo_MedIV")

End Sub

Public Function Excel_VLookup(ByVal varValue As Variant, ByVal strTable As String, ByVal intColumn As Integer) As Variant

    Dim objTable As Range
    
    Set objTable = WbkAfspraken.Names(strTable).RefersToRange
    Excel_VLookup = Application.VLookup(varValue, objTable, intColumn, False)

End Function

Private Sub Test_Excel_VLookup()

    MsgBox CStr(Excel_VLookup("blah", "Tbl_Neo_MedIV", 1))

End Sub

Public Function Excel_Index(ByVal strTable As String, ByVal intRow As Integer, ByVal intColumn As Integer) As Variant

    Dim objTable As Range
    
    Set objTable = WbkAfspraken.Names(strTable).RefersToRange
    Excel_Index = Application.Index(objTable, intRow, intColumn)

End Function

Private Sub Test_Excel_Index()

    MsgBox Excel_Index("Tbl_Neo_MedIV", 2, 1)
 
End Sub

Public Function Excel_RoundBy(ByVal dblValue, dblStep As Double) As Double

    Dim dblRound As Double
    
    dblRound = Application.WorksheetFunction.MRound(dblValue, dblStep)
    
    Excel_RoundBy = dblRound

End Function

Private Sub Test_Excel_RoundBy()
    
    MsgBox Excel_RoundBy(20, 7)

End Sub




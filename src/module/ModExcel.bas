Attribute VB_Name = "ModExcel"
Option Explicit

Public Function Excel_VLookup(ByVal varValue As Variant, ByVal strTable As String, ByVal intColumn As Integer) As Variant

    Dim objTable As Range
    
    Set objTable = WbkAfspraken.Names(strTable).RefersToRange
    Excel_VLookup = Application.VLookup(varValue, objTable, intColumn, False)

End Function

Private Sub Test_Excel_VLookup()

    MsgBox Excel_VLookup("dopamine", "Tbl_Neo_MedIV", 1)

End Sub

Public Function Excel_Index(ByVal strTable As String, ByVal intRow As Integer, ByVal intColumn As Integer) As Variant

    Dim objTable As Range
    
    Set objTable = WbkAfspraken.Names(strTable).RefersToRange
    Excel_Index = Application.Index(objTable, intRow, intColumn)

End Function

Private Sub Test_Excel_Index()

    MsgBox Excel_Index("Tbl_Neo_MedIV", 2, 1)
 
End Sub



Attribute VB_Name = "ModPublic"
Option Explicit

Public Patienten() As Variant, patRec As Integer
Public BedNummer As Variant

Public Sub PatIndex()

Dim i As Integer, nRec As Integer
With Sheets("Patienten")
nRec = .Range("a1").CurrentRegion.Columns.Count
ReDim Patienten(nRec - 4)
    For i = 4 To nRec
        Patienten(i - 4) = .Cells(2, i).Formula
    Next i
End With
End Sub

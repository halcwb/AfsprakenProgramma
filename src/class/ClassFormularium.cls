VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFormularium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lngCount As Integer
Private m_colMedicamenten As Collection

Public Sub AddMedicament(ByVal v_vntGPK As Variant)

    Dim objMedicament As New ClassMedicatieDisc
    
    If v_vntGPK = vbNullString Then
        Exit Sub
    End If
    
    objMedicament.GPK = v_vntGPK
    m_colMedicamenten.Add objMedicament
    
    m_lngCount = m_lngCount + 1
    
    Set objMedicament = Nothing

End Sub

Public Function GPK(ByVal v_GPK As String) As ClassMedicatieDisc
    Dim objMed As ClassMedicatieDisc
    Dim objRes As ClassMedicatieDisc

    For Each objMed In m_colMedicamenten
        If objMed.GPK = v_GPK Then Set objRes = objMed
    Next objMed
    
    Set GPK = objRes

End Function

Public Function Item(ByVal v_intIndex As Integer) As ClassMedicatieDisc

    Set Item = m_colMedicamenten.Item(v_intIndex)
    
End Function

Private Sub Class_Initialize()

    Set m_colMedicamenten = New Collection
    GetMedicamenten

End Sub

Public Sub GetMedicamenten()

    Dim i As Integer
    Dim oForm As Range
    Dim oSheet As Worksheet
    
    Dim strFileName As String
    Dim sName As String
    Dim sSheet As String
    
    sName = "Formularium.xlsx"
    sSheet = "Table"
    
    Application.DisplayAlerts = False

    strFileName = ModMedDisc.GetFormulariumDatabasePath() + sName

    Workbooks.Open strFileName, True, True
    
    Set oSheet = Workbooks(sName).Worksheets(sSheet)
    Set oForm = oSheet.Range("A1").CurrentRegion
        
    For i = 2 To oForm.Rows.Count
        Me.AddMedicament (oForm.Cells(i, 1))
        With Me.Item(i - 1)
            .ATC = oForm.Cells(i, 2)
            .TherapieGroep = oForm.Cells(i, 3)
            .TherapieSubgroep = oForm.Cells(i, 4)
            .Generiek = oForm.Cells(i, 5)
            .Etiket = oForm.Cells(i, 6)
            .Vorm = oForm.Cells(i, 7)
            .routes = oForm.Cells(i, 8)
            .Sterkte = oForm.Cells(i, 9)
            .SterkteEenheid = oForm.Cells(i, 10)
            .Dosis = oForm.Cells(i, 11)
            .DosisEenheid = oForm.Cells(i, 12)
            .indicaties = oForm.Cells(i, 13)
        End With
    Next i
    
    Workbooks(sName).Close
    Application.DisplayAlerts = True
    
End Sub

Public Property Get MedicamentCount() As Variant

    MedicamentCount = m_lngCount - 1

End Property

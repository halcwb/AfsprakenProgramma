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

Private m_lngCount As Integer
Private m_colMedicamenten As Collection

Public Sub AddMedicament(ByVal v_vntGPK As Variant)

    Dim objMedicament As ClassMedicatieDisc
    
    Set objMedicament = New ClassMedicatieDisc
    
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
        If objMed.GPK = v_GPK Then
            Set objRes = objMed
            Exit For
        End If
    Next objMed
    
    Set GPK = objRes

End Function

Public Function Item(ByVal v_intIndex As Integer) As ClassMedicatieDisc

    Set Item = m_colMedicamenten.Item(v_intIndex)
    
End Function

Private Sub Class_Initialize()

    Set m_colMedicamenten = New Collection

End Sub

Public Sub GetMedicamenten(ByVal blnShowProgress As Boolean)

    ModMedDisc.GetMedicamenten Me, blnShowProgress
    
End Sub

Public Property Get MedicamentCount() As Variant

    MedicamentCount = m_lngCount - 1

End Property

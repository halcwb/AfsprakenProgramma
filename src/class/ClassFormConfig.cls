VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFormConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lngCount As Integer
Private m_colMedicamenten As Collection
Private m_colGeneriek As Collection
Private m_colVorm As Collection
Private m_colSterkteEenheden As Collection
Private m_colDosisEenheden As Collection
Private m_colRoutes As Collection

Public Function HasSterkteEenheid(ByVal strEenheid As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_colSterkteEenheden
        If varItem = strEenheid Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasSterkteEenheid = blnExists

End Function

Public Function GetSterkteEenheden() As Collection

    Set GetSterkteEenheden = m_colSterkteEenheden

End Function

Public Function HasDosisEenheid(ByVal strEenheid As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_colDosisEenheden
        If varItem = strEenheid Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasDosisEenheid = blnExists

End Function

Public Function GetDosisEenheden() As Collection

    Set GetDosisEenheden = m_colDosisEenheden

End Function

Public Function HasRoute(ByVal strRoute As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_colRoutes
        If varItem = strRoute Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasRoute = blnExists

End Function

Public Function GetRoutes() As Collection

    Set GetRoutes = m_colRoutes

End Function

Public Function HasVorm(ByVal strVorm As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_colVorm
        If varItem = strVorm Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasVorm = blnExists

End Function

Public Function GetVormen() As Collection

    Set GetVormen = m_colVorm

End Function

Public Sub AddMedicament(ByVal objMed As ClassMedDiscConfig)

    Dim varRoute As Variant
    
    If objMed.GPK = vbNullString Then
        Exit Sub
    End If
    
    m_colMedicamenten.Add objMed
    
    CollectionAddDistinctStringNotEmpty m_colGeneriek, objMed.Generic
    CollectionAddDistinctStringNotEmpty m_colVorm, objMed.Shape
    CollectionAddDistinctStringNotEmpty m_colSterkteEenheden, objMed.GenericUnit
    CollectionAddDistinctStringNotEmpty m_colDosisEenheden, objMed.MultipleUnit
    
    For Each varRoute In objMed.GetRouteList()
        CollectionAddDistinctStringNotEmpty m_colRoutes, varRoute
    Next
    
    m_lngCount = m_lngCount + 1
    
End Sub

Public Function GPK(ByVal v_GPK As String) As ClassMedDiscConfig

    Dim objMed As ClassMedDiscConfig
    Dim objRes As ClassMedDiscConfig

    For Each objMed In m_colMedicamenten
        If objMed.GPK = v_GPK Then
            Set objRes = objMed
            Exit For
        End If
    Next objMed
    
    Set GPK = objRes

End Function

Public Function Item(ByVal v_intIndex As Integer) As ClassMedDiscConfig

    Set Item = m_colMedicamenten.Item(v_intIndex)
    
End Function

Private Sub Class_Initialize()

    Set m_colMedicamenten = New Collection
    Set m_colGeneriek = New Collection
    Set m_colVorm = New Collection
    Set m_colSterkteEenheden = New Collection
    Set m_colDosisEenheden = New Collection
    Set m_colRoutes = New Collection
    
    m_colDosisEenheden.Add "mmol"

End Sub

Public Function GetMedicamenten(ByVal blnShowProgress As Boolean) As Collection

    Dim objC As ClassStringComparer

    Formularium_GetMedDiscConfig Me, blnShowProgress
    
    Set objC = New ClassStringComparer
    ModCollection.CollectionSort m_colVorm, objC
    ModCollection.CollectionSort m_colSterkteEenheden, objC
    
    Set GetMedicamenten = m_colMedicamenten

End Function

Public Property Get MedicamentCount() As Variant

    MedicamentCount = m_lngCount - 1

End Property

Public Function HasGeneriek(ByVal strGeneriek As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_colGeneriek
        If varItem = strGeneriek Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasGeneriek = blnExists

End Function

Public Function GetGenerieken() As Collection

    Set GetGenerieken = m_colGeneriek

End Function


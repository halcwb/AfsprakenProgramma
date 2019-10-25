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

Private m_MedicationCollection As Collection
Private m_SolutionCollection As Collection
Private m_DoseCollection As Collection
Private m_GenericCollection As Collection
Private m_ShapeCollection As Collection
Private m_GenericQuantityUnitCollection As Collection
Private m_MultipeQuantityUnitCollection As Collection
Private m_RouteCollection As Collection

Public Function Reset()

    Init

End Function

Public Function HasGenericQuantityUnit(ByVal strUnit As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_GenericQuantityUnitCollection
        If varItem = strUnit Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasGenericQuantityUnit = blnExists

End Function

Public Function GetGenericQuantityUnitCollection() As Collection

    Set GetGenericQuantityUnitCollection = m_GenericQuantityUnitCollection

End Function

Public Function HasMultipleQuantityUnit(ByVal strUnit As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_MultipeQuantityUnitCollection
        If varItem = strUnit Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasMultipleQuantityUnit = blnExists

End Function

Public Function GetDoseUnitCollection() As Collection

    Set GetDoseUnitCollection = m_MultipeQuantityUnitCollection

End Function

Public Function HasRoute(ByVal strRoute As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_RouteCollection
        If varItem = strRoute Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasRoute = blnExists

End Function

Public Function GetRoutes() As Collection

    Set GetRoutes = m_RouteCollection

End Function

Public Function HasShape(ByVal strShape As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_ShapeCollection
        If varItem = strShape Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasShape = blnExists

End Function

Public Function GetShapeCollection() As Collection

    Set GetShapeCollection = m_ShapeCollection

End Function

Public Sub AddMedication(ByVal objMed As ClassMedDisc)

    Dim varRoute As Variant
    
    If objMed.GPK = vbNullString Then
        Exit Sub
    End If
    
    m_MedicationCollection.Add objMed
    
    CollectionAddDistinctStringNotEmpty m_GenericCollection, objMed.Generic
    CollectionAddDistinctStringNotEmpty m_ShapeCollection, objMed.Shape
    CollectionAddDistinctStringNotEmpty m_GenericQuantityUnitCollection, objMed.GenericUnit
    CollectionAddDistinctStringNotEmpty m_MultipeQuantityUnitCollection, objMed.MultipleUnit
    
    For Each varRoute In objMed.GetRouteList()
        CollectionAddDistinctStringNotEmpty m_RouteCollection, varRoute
    Next
        
End Sub

Public Function GPK(ByVal v_GPK As String) As ClassMedDisc

    Dim objMed As ClassMedDisc
'    Dim objRes As ClassMedicatieDisc

    For Each objMed In m_MedicationCollection
        If objMed.GPK = v_GPK Then
'            Set objRes = objMed
            Exit For
        End If
    Next objMed
    
    Set GPK = objMed

End Function

Public Function GetMedication(ByVal intIndex As Integer) As ClassMedDisc

    Dim objMed As ClassMedDisc
    
    Set objMed = m_MedicationCollection.Item(intIndex)
    Set GetMedication = objMed.Clone()
    
End Function

Private Sub Init()

    Set m_MedicationCollection = New Collection
    Set m_GenericCollection = New Collection
    Set m_ShapeCollection = New Collection
    Set m_GenericQuantityUnitCollection = New Collection
    Set m_MultipeQuantityUnitCollection = New Collection
    Set m_RouteCollection = New Collection
    
    m_MultipeQuantityUnitCollection.Add "mmol"

End Sub

Private Sub Class_Initialize()

    Init

End Sub

Public Function GetMedicationCollection(ByVal blnShowProgress As Boolean) As Collection

    Dim objC As ClassStringComparer
    
    If m_MedicationCollection Is Nothing Then Set m_MedicationCollection = New Collection
    
    If Not m_MedicationCollection.Count > 0 Then

        Database_LoadFormularium Me, blnShowProgress
        
        Set objC = New ClassStringComparer
        ModCollection.CollectionSort m_ShapeCollection, objC
        ModCollection.CollectionSort m_GenericQuantityUnitCollection, objC
    
    End If
        
    Set GetMedicationCollection = m_MedicationCollection

End Function

Public Property Get MedicamentCount() As Variant

    MedicamentCount = m_MedicationCollection.Count

End Property

Public Function HasGeneriek(ByVal strGeneriek As String) As Boolean

    Dim varItem As Variant
    Dim blnExists As Boolean
    
    blnExists = False
    For Each varItem In m_GenericCollection
        If varItem = strGeneriek Then
            blnExists = True
            Exit For
        End If
    Next
    
    HasGeneriek = blnExists

End Function

Public Function GetGenerieken() As Collection

    Set GetGenerieken = m_GenericCollection

End Function

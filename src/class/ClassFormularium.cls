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

Private Sub AddSterkteEenheid(ByVal strEenheid As String)

    If Not HasSterkteEenheid(strEenheid) Then m_colSterkteEenheden.Add strEenheid

End Sub

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

Private Sub AddDosisEenheid(ByVal strEenheid As String)

    If Not HasDosisEenheid(strEenheid) Then m_colDosisEenheden.Add strEenheid

End Sub

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

Private Sub AddRoute(ByVal strRoute As String)

    If Not HasRoute(strRoute) Then m_colRoutes.Add strRoute

End Sub

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

Private Sub AddVorm(ByVal strVorm As String)

    If Not HasVorm(strVorm) Then m_colVorm.Add strVorm

End Sub

Public Sub AddMedicament(ByVal objMed As ClassMedicatieDisc)

    Dim varRoute As Variant
    
    If objMed.GPK = vbNullString Then
        Exit Sub
    End If
    
    m_colMedicamenten.Add objMed
    
    AddVorm objMed.Vorm
    AddSterkteEenheid objMed.SterkteEenheid
    AddDosisEenheid objMed.DosisEenheid
    
    For Each varRoute In objMed.GetRoutes()
        AddRoute varRoute
    Next
    
    m_lngCount = m_lngCount + 1
    
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
    Set m_colVorm = New Collection
    Set m_colSterkteEenheden = New Collection
    Set m_colDosisEenheden = New Collection
    Set m_colRoutes = New Collection

End Sub

Public Sub GetMedicamenten(ByVal blnShowProgress As Boolean)

    Dim objC As ClassStringComparer

    ModMedDisc.GetMedicamenten Me, blnShowProgress
    
    Set objC = New ClassStringComparer
    ModCollection.CollectionSort m_colVorm, objC
    ModCollection.CollectionSort m_colSterkteEenheden, objC

End Sub

Public Property Get MedicamentCount() As Variant

    MedicamentCount = m_lngCount - 1

End Property

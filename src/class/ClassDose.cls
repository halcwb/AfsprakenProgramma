VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Department As String
Private m_Generic As String
Private m_Shape As String
Private m_Route As String
Private m_Indication As String
Private m_Gender As String
Private m_MinAgeMo As Double
Private m_MaxAgeMo As Double
Private m_MinWeightKg As Double
Private m_MaxWeightKg As Double
Private m_MinGestDays As Integer
Private m_MaxGestDays As Integer
Private m_Frequencies As String
Private m_Unit As String
Private m_NormDose As Double
Private m_MinDose As Double
Private m_MaxDose As Double
Private m_AbsMaxDose As Double
Private m_MaxPerDose As Double
Private m_IsDosePerKg As Boolean
Private m_IsDosePerM2 As Boolean

Public Property Get Department() As String
    
    Department = m_Department

End Property

Public Property Let Department(ByVal strDep As String)

    m_Department = strDep

End Property

Public Property Get Generic() As String

    Generic = m_Generic

End Property

Public Property Let Generic(ByVal strGeneric As String)

    m_Generic = strGeneric

End Property

Public Property Get Route() As String

    Route = m_Route

End Property

Public Property Let Route(ByVal strRoute As String)

    m_Route = strRoute

End Property

Public Property Get Indication() As String

    Indication = m_Indication

End Property

Public Property Let Indication(ByVal strIndication As String)

    m_Indication = strIndication

End Property

Public Property Get Shape() As String

    Shape = m_Shape

End Property

Public Property Let Shape(ByVal strGPK As String)

    m_Shape = strGPK

End Property

Public Property Get Gender() As String

    Gender = m_Gender

End Property

Public Property Let Gender(ByVal strGender As String)

    m_Gender = strGender

End Property

Public Property Get MinAgeMo() As Double

    MinAgeMo = m_MinAgeMo

End Property

Public Property Let MinAgeMo(ByVal dblMinAgeMo As Double)

    m_MinAgeMo = dblMinAgeMo

End Property

Public Property Get MaxAgeMo() As Double

    MaxAgeMo = m_MaxAgeMo

End Property

Public Property Let MaxAgeMo(ByVal dblMaxAgeMo As Double)

    m_MaxAgeMo = dblMaxAgeMo

End Property

Public Property Get MinWeightKg() As Double

    MinWeightKg = m_MinWeightKg

End Property

Public Property Let MinWeightKg(ByVal dblMinWeightKg As Double)

    m_MinWeightKg = dblMinWeightKg

End Property

Public Property Get MaxWeightKg() As Double

    MaxWeightKg = m_MaxWeightKg

End Property

Public Property Let MaxWeightKg(ByVal dblMaxWeightKg As Double)

    m_MaxWeightKg = dblMaxWeightKg

End Property

Public Property Get MinGestDays() As Integer

    MinGestDays = m_MinGestDays

End Property

Public Property Let MinGestDays(ByVal intMinGestDays As Integer)

    m_MinGestDays = intMinGestDays

End Property

Public Property Get MaxGestDays() As Integer

    MaxGestDays = m_MaxGestDays

End Property

Public Property Let MaxGestDays(ByVal intMaxGestDays As Integer)

    m_MaxGestDays = intMaxGestDays

End Property

Public Property Get Frequencies() As String

    Frequencies = m_Frequencies

End Property

Public Property Let Frequencies(ByVal strFreq As String)

    m_Frequencies = strFreq

End Property

Public Property Get Unit() As String

    Unit = m_Unit

End Property

Public Property Let Unit(ByVal strUnit As String)

    m_Unit = strUnit

End Property

Public Property Get NormDose() As Double

    NormDose = m_NormDose

End Property

Public Property Let NormDose(ByVal dblDose As Double)

    m_NormDose = dblDose

End Property

Public Property Get MinDose() As Double
    
    MinDose = m_MinDose
    
End Property

Public Property Let MinDose(ByVal dblDose As Double)

    m_MinDose = dblDose

End Property

Public Property Get MaxDose() As Double

    MaxDose = m_MaxDose

End Property

Public Property Let MaxDose(ByVal dblDose As Double)

    m_MaxDose = dblDose

End Property

Public Property Get AbsMaxDose() As Double
    
    AbsMaxDose = m_AbsMaxDose
    
End Property

Public Property Let AbsMaxDose(ByVal dblDose As Double)

    m_AbsMaxDose = dblDose

End Property

Public Property Get MaxPerDose() As Double

    MaxPerDose = m_MaxPerDose

End Property

Public Property Let MaxPerDose(ByVal dblDose As Double)

    m_MaxPerDose = dblDose

End Property

Public Property Get IsDosePerKg() As Boolean
    
    IsDosePerKg = m_IsDosePerKg
    
End Property

Public Property Let IsDosePerKg(ByVal blnPerKg As Boolean)

    m_IsDosePerKg = blnPerKg

End Property

Public Property Get IsDosePerM2() As Boolean
    
    IsDosePerM2 = m_IsDosePerM2
    
End Property

Public Property Let IsDosePerM2(ByVal blnPerM2 As Boolean)

    m_IsDosePerM2 = blnPerM2

End Property


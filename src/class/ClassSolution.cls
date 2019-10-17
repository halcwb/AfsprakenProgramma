VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Department As String
Private m_Generic As String
Private m_Shape As String
Private m_MinGenericQuantity As Double
Private m_MaxGenericQuantity As Double
Private m_Solutions As String
Private m_MinConc As Double
Private m_MaxConc As Double
Private m_SolutionVolume As Double
Private m_MinInfusionTime As Integer

Public Property Get Department() As String

    Department = m_Department

End Property

Public Property Let Department(ByVal strDepartment As String)

    m_Department = strDepartment

End Property

Public Property Get Generic() As String

    Generic = m_Generic

End Property

Public Property Let Generic(ByVal strGeneric As String)

    m_Generic = strGeneric

End Property

Public Property Get Shape() As String

    Shape = m_Shape

End Property

Public Property Let Shape(ByVal strShape As String)
    
    m_Shape = strShape

End Property

Public Property Get MinGenericQuantity() As Double

    MinGenericQuantity = m_MinGenericQuantity

End Property

Public Property Let MinGenericQuantity(ByVal dblGenericQuantity As Double)

    m_MinGenericQuantity = dblGenericQuantity

End Property

Public Property Get MaxGenericQuantity() As Double

    MaxGenericQuantity = m_MaxGenericQuantity

End Property

Public Property Let MaxGenericQuantity(ByVal dblGenericQuantity As Double)

    m_MaxGenericQuantity = dblGenericQuantity

End Property

Public Property Get MinConc() As Double

    MinConc = m_MinConc

End Property

Public Property Let MinConc(ByVal dblConc As Double)

    m_MinConc = dblConc

End Property

Public Property Get MaxConc() As Double

    MaxConc = m_MaxConc

End Property

Public Property Let MaxConc(ByVal dblConc As Double)

    m_MaxConc = dblConc

End Property

Public Property Get Solutions() As String

    Solutions = m_Solutions

End Property

Public Property Let Solutions(ByVal strSols As String)

    m_Solutions = strSols

End Property

Public Property Get SolutionVolume() As Double

    SolutionVolume = m_SolutionVolume

End Property

Public Property Let SolutionVolume(ByVal dblVol As Double)

    m_SolutionVolume = dblVol

End Property

Public Property Get MinInfusionTime() As Integer

    MinInfusionTime = m_MinInfusionTime

End Property

Public Property Let MinInfusionTime(ByVal intTime As Integer)

    m_MinInfusionTime = intTime

End Property


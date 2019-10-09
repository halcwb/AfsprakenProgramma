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
Private m_MinAge As Double
Private m_MaxAge As Double
Private m_MinGestDays As Integer
Private m_MaxGestDays As Integer
Private m_MinWeight As Double
Private m_MaxWeight As Double
Private m_Generic As String
Private m_Shape As String
Private m_Solution As String
Private m_MinConc As Double
Private m_MaxConc As Double
Private m_SolutionVolume As Double
Private m_MinInfusionTime As Integer

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

Public Property Get Solution() As String

    Solution = m_Solution

End Property

Public Property Let Solution(ByVal strVlst As String)

    m_Solution = strVlst

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


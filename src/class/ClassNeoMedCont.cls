VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassNeoMedCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Generic As String
Private m_GenericUnit As String
Private m_DoseUnit As String
Private m_GenericQuantity As Double
Private m_GenericVolume As Double
Private m_MinDose As Double
Private m_MaxDose As Double
Private m_AbsMaxDose As Double
Private m_MinConcentration As Double
Private m_MaxConcentration As Double
Private m_Solution As String
Private m_DoseAdvice As String
Private m_SolutionVolume As Double
Private m_DripQuantity As Double
Private m_Product As String
Private m_ShelfLife As Integer
Private m_ShelfCondition As String
Private m_PreparationText As String
Private m_NormDose As String
Private m_DilutionText As String

Public Property Get Generic() As String

    Generic = m_Generic

End Property

Public Property Let Generic(ByVal strName As String)

    m_Generic = strName

End Property

Public Property Get GenericUnit() As String

    GenericUnit = m_GenericUnit

End Property

Public Property Let GenericUnit(ByVal strUnit As String)

    m_GenericUnit = strUnit

End Property

Public Property Get DoseUnit() As String

    DoseUnit = m_DoseUnit

End Property

Public Property Let DoseUnit(ByVal strDoseUnit As String)

    m_DoseUnit = strDoseUnit

End Property

Public Property Get GenericQuantity() As Double

    GenericQuantity = m_GenericQuantity

End Property

Public Property Let GenericQuantity(ByVal dblConc As Double)

    m_GenericQuantity = dblConc

End Property

Public Property Get GenericVolume() As Double

    GenericVolume = m_GenericVolume

End Property

Public Property Let GenericVolume(ByVal dblVolume As Double)

    m_GenericVolume = dblVolume

End Property

Public Property Get MinDose() As Double

    MinDose = m_MinDose

End Property

Public Property Let MinDose(ByVal dblMinDose As Double)

    m_MinDose = dblMinDose

End Property

Public Property Get MaxDose() As Double

    MaxDose = m_MaxDose

End Property

Public Property Let MaxDose(ByVal dblMaxDose As Double)

    m_MaxDose = dblMaxDose

End Property

Public Property Get AbsMaxDose() As Double

    AbsMaxDose = m_AbsMaxDose

End Property

Public Property Let AbsMaxDose(ByVal dblAbsMax As Double)

    m_AbsMaxDose = dblAbsMax

End Property

Public Property Get MinConcentration() As Double

    MinConcentration = m_MinConcentration

End Property

Public Property Let MinConcentration(ByVal dblMinConc As Double)

    m_MinConcentration = dblMinConc

End Property

Public Property Get MaxConcentration() As Double

    MaxConcentration = m_MaxConcentration

End Property

Public Property Let MaxConcentration(ByVal dblMaxConc As Double)

    m_MaxConcentration = dblMaxConc

End Property

Public Property Get Solution() As String

    Solution = m_Solution

End Property

Public Property Let Solution(ByVal strOplVlst As String)

    m_Solution = strOplVlst

End Property

Public Property Get DoseAdvice() As String

    DoseAdvice = m_DoseAdvice

End Property

Public Property Let DoseAdvice(ByVal strAdvice As String)

    m_DoseAdvice = strAdvice

End Property

Public Property Get SolutionVolume() As Double

    SolutionVolume = m_SolutionVolume

End Property

Public Property Let SolutionVolume(ByVal dblOplVol As Double)

    m_SolutionVolume = dblOplVol

End Property

Public Property Get DripQuantity() As Double

    DripQuantity = m_DripQuantity

End Property

Public Property Let DripQuantity(ByVal dblRate As Double)

    m_DripQuantity = dblRate

End Property

Public Property Get Product() As String

    Product = m_Product

End Property

Public Property Let Product(ByVal strProduct As String)

    m_Product = strProduct

End Property

Public Property Get ShelfLife() As Integer

    ShelfLife = m_ShelfLife

End Property

Public Property Let ShelfLife(ByVal intHoudbaar As Integer)

    m_ShelfLife = intHoudbaar

End Property

Public Property Get ShelfCondition() As String

    ShelfCondition = m_ShelfCondition

End Property

Public Property Let ShelfCondition(ByVal strBewaar As String)

    m_ShelfCondition = strBewaar

End Property

Public Property Get PreparationText() As String

    PreparationText = m_PreparationText

End Property

Public Property Let PreparationText(ByVal strTekst As String)

    m_PreparationText = strTekst

End Property

Public Property Get DilutionText() As String

    DilutionText = m_DilutionText

End Property

Public Property Let DilutionText(ByVal strTekst As String)

    m_DilutionText = strTekst

End Property



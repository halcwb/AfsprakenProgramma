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

Private m_Name As String
Private m_Unit As String
Private m_DoseUnit As String
Private m_Conc As Double
Private m_Volume As Double
Private m_MinDose As Double
Private m_MaxDose As Double
Private m_AbsMax As Double
Private m_MinConc As Double
Private m_MaxConc As Double
Private m_OplVlst As String
Private m_Advice As String
Private m_OplVol As Double
Private m_Rate As Double
Private m_Product As String
Private m_Houdbaar As Integer
Private m_Bewaar As String
Private m_Tekst As String
Private m_NormDose As String

Public Property Get Name() As String

    Name = m_Name

End Property

Public Property Let Name(ByVal strName As String)

    m_Name = strName

End Property

Public Property Get Unit() As String

    Unit = m_Unit

End Property

Public Property Let Unit(ByVal strUnit As String)

    m_Unit = strUnit

End Property

Public Property Get DoseUnit() As String

    DoseUnit = m_DoseUnit

End Property

Public Property Let DoseUnit(ByVal strDoseUnit As String)

    m_DoseUnit = strDoseUnit

End Property

Public Property Get Conc() As Double

    Conc = m_Conc

End Property

Public Property Let Conc(ByVal dblConc As Double)

    m_Conc = dblConc

End Property

Public Property Get Volume() As Double

    Volume = m_Volume

End Property

Public Property Let Volume(ByVal dblVolume As Double)

    m_Volume = dblVolume

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

Public Property Get AbsMax() As Double

    AbsMax = m_AbsMax

End Property

Public Property Let AbsMax(ByVal dblAbsMax As Double)

    m_AbsMax = dblAbsMax

End Property

Public Property Get MinConc() As Double

    MinConc = m_MinConc

End Property

Public Property Let MinConc(ByVal dblMinConc As Double)

    m_MinConc = dblMinConc

End Property

Public Property Get MaxConc() As Double

    MaxConc = m_MaxConc

End Property

Public Property Let MaxConc(ByVal dblMaxConc As Double)

    m_MaxConc = dblMaxConc

End Property

Public Property Get OplVlst() As String

    OplVlst = m_OplVlst

End Property

Public Property Let OplVlst(ByVal strOplVlst As String)

    m_OplVlst = strOplVlst

End Property

Public Property Get Advice() As String

    Advice = m_Advice

End Property

Public Property Let Advice(ByVal strAdvice As String)

    m_Advice = strAdvice

End Property

Public Property Get OplVol() As Double

    OplVol = m_OplVol

End Property

Public Property Let OplVol(ByVal dblOplVol As Double)

    m_OplVol = dblOplVol

End Property

Public Property Get Rate() As Double

    Rate = m_Rate

End Property

Public Property Let Rate(ByVal dblRate As Double)

    m_Rate = dblRate

End Property

Public Property Get Product() As String

    Product = m_Product

End Property

Public Property Let Product(ByVal strProduct As String)

    m_Product = strProduct

End Property

Public Property Get Houdbaar() As Integer

    Houdbaar = m_Houdbaar

End Property

Public Property Let Houdbaar(ByVal intHoudbaar As Integer)

    m_Houdbaar = intHoudbaar

End Property

Public Property Get Bewaar() As String

    Bewaar = m_Bewaar

End Property

Public Property Let Bewaar(ByVal strBewaar As String)

    m_Bewaar = strBewaar

End Property

Public Property Get Tekst() As String

    Tekst = m_Tekst

End Property

Public Property Let Tekst(ByVal strTekst As String)

    m_Tekst = strTekst

End Property


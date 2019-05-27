VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDoseRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Substance As String
Private m_Freq As String
Private m_PerDose As Boolean
Private m_PerKg As Boolean
Private m_PerM2 As Boolean
Private m_KeerDose As Double
Private m_NormDose As Double
Private m_MinDose As Double
Private m_MaxDose As Double
Private m_AbsMaxDose As Double
Private m_MaxPerDose As Double

Public Property Get PerKg() As Boolean

    PerKg = m_PerKg

End Property

Public Property Let PerKg(ByVal blnPerKg As Boolean)

    If blnPerKg Then m_PerM2 = False
    m_PerKg = blnPerKg

End Property

Public Property Get PerM2() As Boolean

    PerM2 = m_PerM2

End Property

Public Property Let PerM2(ByVal blnPerM2 As Boolean)

    If blnPerM2 Then m_PerKg = False
    m_PerM2 = blnPerM2

End Property

Public Property Get PerDose() As Boolean

    PerDose = m_PerDose

End Property

Public Property Let PerDose(ByVal blnPerDose As Boolean)

    m_PerDose = blnPerDose

End Property

Public Property Get MaxDose() As Double

    MaxDose = m_MaxDose

End Property

Public Property Let MaxDose(ByVal dblDose As Double)

    m_MaxDose = dblDose

End Property

Public Property Get NormDose() As Double

    NormDose = m_NormDose

End Property

Public Property Let NormDose(ByVal dblDose As Double)

    m_NormDose = dblDose

End Property

Public Property Get MaxPerDose() As Double

    MaxPerDose = m_MaxPerDose

End Property

Public Property Let MaxPerDose(ByVal dblDose As Double)

    m_MaxPerDose = dblDose

End Property

Public Property Get Substance() As String

    Substance = m_Substance

End Property

Public Property Let Substance(ByVal strSubstance As String)

    m_Substance = Trim(LCase(strSubstance))

End Property

Public Property Get Freq() As String

    Freq = m_Freq

End Property

Public Property Let Freq(ByVal strFreq As String)

    m_Freq = strFreq

End Property

Public Property Get MinDose() As Double
    
    MinDose = m_MinDose
    
End Property

Public Property Let MinDose(ByVal dblDose As Double)

    m_MinDose = dblDose

End Property

Public Property Get AbsMaxDose() As Double
    
    AbsMaxDose = m_AbsMaxDose
    
End Property

Public Property Let AbsMaxDose(ByVal dblDose As Double)

    m_AbsMaxDose = dblDose

End Property

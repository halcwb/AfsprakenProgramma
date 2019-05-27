VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSubstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Substance As String
Private m_Concentration As Double
Private m_Dose As Double
Private m_DoseUnit As String

Public Property Get Substance() As String

    Substance = m_Substance

End Property

Public Property Let Substance(ByVal strSubst As String)

    m_Substance = Trim(LCase(strSubst))

End Property

Public Property Get Concentration() As Double

    Concentration = m_Concentration

End Property

Public Property Let Concentration(ByVal dblConc As Double)

    m_Concentration = dblConc

End Property

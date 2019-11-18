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
Private m_MinPMDays As Integer
Private m_MaxPMDays As Integer
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

    m_Department = Trim(strDep)

End Property

Public Property Get Generic() As String

    Generic = m_Generic

End Property

Public Property Let Generic(ByVal strGeneric As String)

    m_Generic = Trim(strGeneric)

End Property

Public Property Get Route() As String

    Route = m_Route

End Property

Public Property Let Route(ByVal strRoute As String)

    m_Route = Trim(strRoute)

End Property

Public Property Get Indication() As String

    Indication = m_Indication

End Property

Public Property Let Indication(ByVal strIndication As String)

    m_Indication = Trim(strIndication)

End Property

Public Property Get Shape() As String

    Shape = m_Shape

End Property

Public Property Let Shape(ByVal strShape As String)

    m_Shape = Trim(strShape)

End Property

Public Property Get Gender() As String

    Gender = m_Gender

End Property

Public Property Let Gender(ByVal strGender As String)

    m_Gender = Trim(strGender)

End Property

Public Property Get MinAgeMo() As Double

    MinAgeMo = m_MinAgeMo

End Property

Public Property Let MinAgeMo(ByVal dblMinAgeMo As Double)

    m_MinAgeMo = IIf(dblMinAgeMo < 0, 0, dblMinAgeMo)

End Property

Public Property Get MaxAgeMo() As Double

    MaxAgeMo = m_MaxAgeMo

End Property

Public Property Let MaxAgeMo(ByVal dblMaxAgeMo As Double)

    m_MaxAgeMo = IIf(dblMaxAgeMo < 0, 0, dblMaxAgeMo)

End Property

Public Property Get MinWeightKg() As Double

    MinWeightKg = m_MinWeightKg

End Property

Public Property Let MinWeightKg(ByVal dblMinWeightKg As Double)

    m_MinWeightKg = IIf(dblMinWeightKg < 0, 0, dblMinWeightKg)

End Property

Public Property Get MaxWeightKg() As Double

    MaxWeightKg = m_MaxWeightKg

End Property

Public Property Let MaxWeightKg(ByVal dblMaxWeightKg As Double)

    m_MaxWeightKg = IIf(dblMaxWeightKg < 0, 0, dblMaxWeightKg)

End Property

Public Property Get MinGestDays() As Integer

    MinGestDays = m_MinGestDays

End Property

Public Property Let MinGestDays(ByVal intMinGestDays As Integer)

    m_MinGestDays = IIf(intMinGestDays < 0, 0, intMinGestDays)

End Property

Public Property Get MaxGestDays() As Integer

    MaxGestDays = m_MaxGestDays

End Property

Public Property Let MaxGestDays(ByVal intMaxGestDays As Integer)

    m_MaxGestDays = IIf(intMaxGestDays < 0, 0, intMaxGestDays)

End Property

Public Property Get MinPMDays() As Integer

    MinPMDays = m_MinPMDays

End Property

Public Property Let MinPMDays(ByVal intMinPMDays As Integer)

    m_MinPMDays = IIf(intMinPMDays < 0, 0, intMinPMDays)

End Property

Public Property Get MaxPMDays() As Integer

    MaxPMDays = m_MaxPMDays

End Property

Public Property Let MaxPMDays(ByVal intMaxPMDays As Integer)

    m_MaxPMDays = IIf(intMaxPMDays < 0, 0, intMaxPMDays)

End Property

Public Property Get Frequencies() As String

    Frequencies = m_Frequencies

End Property

Public Property Let Frequencies(ByVal strFreq As String)

    m_Frequencies = Trim(strFreq)

End Property

Public Property Get Unit() As String

    Unit = m_Unit

End Property

Public Property Let Unit(ByVal strUnit As String)

    m_Unit = Trim(strUnit)

End Property

Public Property Get NormDose() As Double

    NormDose = m_NormDose

End Property

Public Property Let NormDose(ByVal dblDose As Double)

    m_NormDose = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get MinDose() As Double
    
    MinDose = m_MinDose
    
End Property

Public Property Let MinDose(ByVal dblDose As Double)

    m_MinDose = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get MaxDose() As Double

    MaxDose = m_MaxDose

End Property

Public Property Let MaxDose(ByVal dblDose As Double)

    m_MaxDose = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get AbsMaxDose() As Double
    
    AbsMaxDose = m_AbsMaxDose
    
End Property

Public Property Let AbsMaxDose(ByVal dblDose As Double)

    m_AbsMaxDose = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get MaxPerDose() As Double

    MaxPerDose = m_MaxPerDose

End Property

Public Property Let MaxPerDose(ByVal dblDose As Double)

    m_MaxPerDose = IIf(dblDose < 0, 0, dblDose)

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

Public Function IsEmpty() As Boolean

    Dim blnHasDose As Boolean
    
    blnHasDose = False
    blnHasDose = blnHasDose Or NormDose > 0
    blnHasDose = blnHasDose Or MinDose > 0
    blnHasDose = blnHasDose Or MaxDose > 0
    blnHasDose = blnHasDose Or AbsMaxDose > 0
    blnHasDose = blnHasDose Or MaxPerDose > 0
    
    IsEmpty = Not blnHasDose

End Function

Public Function ToString() As String

    Dim objBuilder As ClassStringBuilder
    
    Set objBuilder = New ClassStringBuilder
    
    objBuilder.Append "DoseRule for: " & vbNewLine
    objBuilder.Append "Generic: " & Generic & vbNewLine
    objBuilder.Append "Shape: " & Shape & vbNewLine
    objBuilder.Append "Indication: " & Indication & vbNewLine
    objBuilder.Append "Route: " & Route & vbNewLine
    objBuilder.Append "Department: " & Department & vbNewLine
    objBuilder.Append "Gender: " & Gender & vbNewLine
    objBuilder.Append "Age: " & MinAgeMo & " - " & MaxAgeMo & " months" & vbNewLine
    objBuilder.Append "Weight: " & MinWeightKg & " - " & MaxWeightKg & " kg" & vbNewLine
    objBuilder.Append "GestAge: " & MinGestDays & " - " & MaxGestDays & " days" & vbNewLine
    objBuilder.Append "Frequencies: " & Frequencies & vbNewLine
    objBuilder.Append "NormDose: " & NormDose & " " & Unit & IIf(IsDosePerKg, "/kg", "") & vbNewLine
    objBuilder.Append "DoseRange: " & MinDose & " - " & MaxDose & " " & Unit & IIf(IsDosePerKg, "/kg", "") & vbNewLine
    objBuilder.Append "AbsMaxDose: " & AbsMaxDose & " " & Unit & vbNewLine
    objBuilder.Append "MaxPerDose: " & MaxPerDose & Unit & vbNewLine
    
    
    ToString = objBuilder.ToString()
    
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPatientDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ID As String
Private m_AN As String
Private m_VN As String
Private m_BD As Date
Private m_AD As Date
Private m_WT As Double
Private m_BW As Double
Private m_LE As Double
Private m_GS As String
Private m_GW As Integer
Private m_GD As Integer
Private m_LU As String


Public Property Get AchterNaam() As String

    AchterNaam = m_AN

End Property

Public Property Let AchterNaam(ByVal strAN As String)

    m_AN = strAN

End Property

Public Property Get VoorNaam() As String

    VoorNaam = m_VN

End Property

Public Property Let VoorNaam(ByVal strVN As String)

    m_VN = strVN

End Property

Public Property Get Gewicht() As Double

    Gewicht = m_WT

End Property

Public Property Let Gewicht(ByVal dblWeight As Double)

    dblWeight = ModString.FixPrecision(dblWeight, 2)
    If ModPatient.ValidWeightKg(dblWeight) Then m_WT = dblWeight

End Property

Public Property Get GeboorteGewicht() As Double

    GeboorteGewicht = m_BW

End Property

Public Property Let GeboorteGewicht(ByVal dblWeight As Double)

    m_BW = dblWeight

End Property

Public Property Get Lengte() As Double

    Lengte = m_LE

End Property

Public Property Let Lengte(ByVal dblLen As Double)

    m_LE = dblLen

End Property

Public Property Get Oppervlak() As Double

    On Error Resume Next

    If m_LE > 0 And m_WT > 0 Then
        Oppervlak = m_WT ^ 0.425 * m_LE ^ 0.725 * 0.007184
    Else
        Oppervlak = 0
    End If

End Property

Public Property Get Weeks() As Integer

    Weeks = m_GW

End Property

Public Property Let Weeks(ByVal intWeeks As Integer)

    m_GW = intWeeks

End Property

Public Property Get Days() As Integer

    Days = m_GD

End Property

Public Property Let Days(ByVal intDays As Integer)

    m_GD = intDays

End Property

Public Property Get PatientId() As String

    PatientId = m_ID

End Property

Public Property Let PatientId(ByVal strId As String)

    m_ID = strId

End Property

Public Property Get Bed() As String

    Bed = m_LU

End Property

Public Property Let Bed(ByVal strBed As String)

    m_LU = strBed

End Property

Public Property Get Geslacht() As String

    Geslacht = m_GS

End Property

Public Property Let Geslacht(ByVal strGesl As String)
   
    strGesl = Trim(Strings.LCase(strGesl))
    If strGesl = "man" Then
        m_GS = "man"
    ElseIf strGesl = "vrouw" Then
        m_GS = "vrouw"
    ElseIf strGesl = "onbepaald" Or strGesl = "onbekend" Then
        m_GS = "onbepaald"
    End If

End Property

Public Property Get GeboorteDatum() As Date

    GeboorteDatum = m_BD

End Property

Public Property Get OpnameDatum() As Date

    OpnameDatum = m_AD

End Property

Public Function SetAdmissionAndBirthDate(ByVal dtmAdm As Date, ByVal dtmBd As Date) As Boolean

    If ModPatient.ValidBirthDate(dtmBd, dtmAdm) Then
        m_AD = dtmAdm
        m_BD = dtmBd
        
        SetAdmissionAndBirthDate = True
    Else
        m_AD = ModDate.EmptyDate
        m_BD = ModDate.EmptyDate
        
        SetAdmissionAndBirthDate = False
    End If
    
End Function

Public Sub Clear()

 m_ID = vbNullString
 m_AN = vbNullString
 m_VN = vbNullString
 m_BD = ModDate.EmptyDate
 m_AD = ModDate.EmptyDate
 m_WT = 0
 m_BW = 0
 m_LE = 0
 m_GS = vbNullString
 m_GW = 0
 m_GD = 0
 m_LU = vbNullString

End Sub



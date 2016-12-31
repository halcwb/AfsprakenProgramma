VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPatientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ID As String
Private m_Bed As String
Private m_AN As String
Private m_VN As String
Private m_BD As String

Public Property Get Id() As String

    Id = m_ID

End Property

Public Property Let Id(ByVal strID As String)

    m_ID = strID

End Property

Public Property Get Bed() As String

    Bed = m_Bed

End Property

Public Property Let Bed(ByVal strBed As String)

    m_Bed = strBed

End Property

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

Public Property Get BirthDate() As String

    BirthDate = m_BD

End Property

Public Property Let BirthDate(ByVal strBD As String)

    m_BD = strBD

End Property

Public Function ToString() As String

    ToString = Me.Bed() & ": " & Me.Id() & " " & Me.AchterNaam() & ", " & Me.VoorNaam() & " " & Me.BirthDate()

End Function

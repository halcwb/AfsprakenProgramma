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

Private Const constNull As Variant = "0"

Private m_ID As String
Private m_Bed As String
Private m_AN As String
Private m_VN As String
Private m_BD As String

Public Property Get Id() As String

    Id = m_ID

End Property

Public Property Let Id(ByVal strID As String)

    If Not strID = constNull Then m_ID = strID Else m_ID = vbNullString

End Property

Public Property Get Bed() As String

    Bed = m_Bed

End Property

Public Property Let Bed(ByVal strBed As String)

    If Not strBed = constNull Then m_Bed = strBed Else m_Bed = vbNullString

End Property

Public Property Get AchterNaam() As String

    AchterNaam = m_AN

End Property

Public Property Let AchterNaam(ByVal strAN As String)

    If Not constNull = strAN Then m_AN = strAN Else m_AN = vbNullString

End Property

Public Property Get VoorNaam() As String

    VoorNaam = m_VN

End Property

Public Property Let VoorNaam(ByRef strVN As String)

    If Not strVN = constNull Then m_VN = strVN Else strVN = vbNullString

End Property

Public Property Get BirthDate() As String

    BirthDate = m_BD

End Property

Public Property Let BirthDate(ByVal strBD As String)

    If Not strBD = constNull Then m_BD = strBD Else m_BD = vbNullString

End Property

Public Function toString() As String

    Dim strString As String
    
    strString = Me.Bed() & ": "
    strString = strString & Me.Id()
    
    If Me.AchterNaam() <> vbNullString And Me.VoorNaam() <> vbNullString Then
        strString = Strings.Trim(strString) & " " & Me.AchterNaam & ", " & Me.VoorNaam()
    Else
        strString = Strings.Trim(strString) & Me.AchterNaam() & Me.VoorNaam()
    End If
    
    strString = IIf(Me.AchterNaam = vbNullString, strString, Strings.Trim(strString) & " " & Me.BirthDate())

    toString = strString

End Function

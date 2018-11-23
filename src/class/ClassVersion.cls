VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_VersionID As Integer
Private m_VersionUTC As String
Private m_VersionDate As String

Public Property Get VersionID() As Integer

    VersionID = m_VersionID

End Property

Public Property Let VersionID(intID As Integer)

    m_VersionID = intID

End Property


Public Property Get VersionUTC() As String

    VersionUTC = m_VersionUTC

End Property

Public Property Let VersionUTC(strUTC As String)

    m_VersionUTC = strUTC

End Property


Public Property Get VersionDate() As String

    VersionDate = m_VersionDate

End Property

Public Property Let VersionDate(strDate As String)

    m_VersionDate = strDate

End Property

Public Function ToString() As String

    ToString = m_VersionID & " : " & m_VersionDate

End Function

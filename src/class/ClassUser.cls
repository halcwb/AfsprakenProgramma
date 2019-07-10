VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Login As String
Private m_LastName As String
Private m_FirstName As String
Private m_Role As String
Private m_PIN As String

Public Property Get Login() As String

    Login = m_Login

End Property

Public Property Let Login(ByVal strLogin As String)

    If Not Trim(strLogin) = vbNullString Then
        m_Login = strLogin
    End If

End Property

Public Property Get LastName() As String

    LastName = m_LastName

End Property

Public Property Let LastName(ByVal strLastName As String)

    If Not Trim(strLastName) = vbNullString Then
        m_LastName = strLastName
    End If

End Property

Public Property Get FirstName() As String

    FirstName = m_FirstName

End Property

Public Property Let FirstName(ByVal strFirstName As String)

    If Not Trim(strFirstName) = vbNullString Then
        m_FirstName = strFirstName
    End If

End Property

Public Property Get Role() As String

    Role = m_Role

End Property

Public Property Let Role(ByVal strRole As String)

    If Not Trim(strRole) = vbNullString Then
        m_Role = strRole
    End If

End Property

Public Property Get PIN() As String

    PIN = m_PIN

End Property

Public Property Let PIN(ByVal strPIN As String)

    m_PIN = strPIN

End Property

Public Property Get IsValid() As Boolean

    Dim blnValid As Boolean
    
    blnValid = Not m_Login = vbNullString
    blnValid = blnValid And Not m_LastName = vbNullString
    blnValid = blnValid And Not m_FirstName = vbNullString
    blnValid = blnValid And Not m_Role = vbNullString
    
    IsValid = blnValid

End Property

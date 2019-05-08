VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSignal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Value As Variant
Private m_Time As Date
Private m_Unit As String

Public Property Get Name() As String

    Name = m_Name

End Property

Public Property Let Name(ByVal strName As String)

    m_Name = strName

End Property

Public Property Get Value() As Variant

    Value = m_Value

End Property

Public Property Let Value(ByVal varValue As Variant)

    m_Value = varValue

End Property

Public Property Get Time() As Date

    Time = m_Time

End Property

Public Property Let Time(ByVal dtmTime As Date)

    m_Time = dtmTime

End Property

Public Property Get Unit() As String

    Unit = m_Unit

End Property

Public Property Let Unit(ByVal strUnit As String)

    m_Unit = strUnit

End Property







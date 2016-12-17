VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



'=============================================================
' Program      :    Afspraken programma
'
' Filename     :    N/A
'
' Class Name   :    CPatient
'
' Author       :    C.W. Bollen
'                   Informedica
'
' Date         :    2000
'
' Description  :    Business object
'
' Revision History

' 16-feb-2001 C.W. Bollen
' Naming conventions toegepast. Opmaak aangepast
'=============================================================

Option Explicit

Private m_PatientID As Variant
Private m_AchterNaam As Variant
Private m_VoorNaam As Variant
Private m_GeboorteDatum As Variant
Private m_Gewicht As Variant
Private m_Lengte As Variant
Private m_Jaren As Variant
Private m_Maanden As Variant

Const c_MaxJaren = 100
Const c_MinJaren = 0
Const c_MaxMaanden = 12
Const c_MinMaanden = 0
Const c_MaxGewicht = 150
Const c_MinGewicht = 0
Const c_MaxLengte = 250
Const c_MinLengte = 25


Public Property Get AchterNaam() As Variant

    AchterNaam = m_AchterNaam

End Property

Public Property Let AchterNaam(ByVal v_AchterNaam As Variant)

    m_AchterNaam = v_AchterNaam

End Property

Public Property Get VoorNaam() As Variant

    VoorNaam = m_VoorNaam

End Property

Public Property Let VoorNaam(ByVal v_VoorNaam As Variant)

    m_VoorNaam = v_VoorNaam

End Property

Public Property Get Gewicht() As Variant

Gewicht = m_Gewicht

End Property

Public Property Let Gewicht(ByVal v_Gewicht As Variant)

If IsNull(v_Gewicht) Then
    m_Gewicht = Null
Else
    If IsNumeric(v_Gewicht) Then
        If v_Gewicht >= c_MinGewicht And v_Gewicht <= c_MaxGewicht Then
            m_Gewicht = v_Gewicht
        Else
            MsgBox prompt:="Gewicht out of Range in clsPatient", Buttons:=vbExclamation
            m_Gewicht = Null
        End If
    Else
        MsgBox prompt:="Invalid Gewicht entry in clsPatient", Buttons:=vbExclamation
        m_Gewicht = Null
    End If
End If

End Property

Public Property Get Lengte() As Variant

Lengte = m_Lengte

End Property

Public Property Let Lengte(ByVal v_Lengte As Variant)

If IsNull(v_Lengte) Then
    m_Lengte = Null
Else
    If IsNumeric(v_Lengte) Then
        If v_Lengte >= c_MinLengte And v_Lengte <= c_MaxLengte Then
            m_Lengte = v_Lengte
        Else
            MsgBox prompt:="Lengte out of Range in clsPatient", Buttons:=vbExclamation
            m_Lengte = Null
        End If
    Else
        MsgBox prompt:="Invalid Lengte entry in clsPatient", Buttons:=vbExclamation
        m_Lengte = Null
    End If
End If

End Property

Public Property Get Oppervlak() As Variant

On Error Resume Next

If m_Lengte > 0 And m_Gewicht > 0 Then
    Oppervlak = m_Gewicht ^ 0.425 * m_Lengte ^ 0.725 * 0.007184
Else
    Oppervlak = Null
End If

End Property


Public Property Get Jaren() As Variant

Jaren = m_Jaren

End Property

Public Property Let Jaren(ByVal v_Jaren As Variant)

If IsNull(v_Jaren) Then
    m_Jaren = Null
Else
    If IsNumeric(v_Jaren) Then
        If v_Jaren >= c_MinJaren And v_Jaren <= c_MaxJaren Then
            m_Jaren = v_Jaren
        Else
            MsgBox prompt:="Jaren out of Range in clsPatient", Buttons:=vbExclamation
            m_Jaren = Null
        End If
    Else
        MsgBox prompt:="Invalid Jaren entry in clsPatient", Buttons:=vbExclamation
        m_Jaren = Null
    End If
End If

End Property

Public Property Get Maanden() As Variant

Maanden = m_Maanden

End Property

Public Property Let Maanden(ByVal v_Maanden As Variant)

If IsNull(v_Maanden) Then
    m_Maanden = Null
Else
    If IsNumeric(v_Maanden) Then
        If v_Maanden >= c_MinMaanden And v_Maanden <= c_MaxMaanden Then
            m_Maanden = v_Maanden
        Else
            MsgBox prompt:="Maanden out of Range in clsPatient", Buttons:=vbExclamation
            m_Maanden = Null
        End If
    Else
        MsgBox prompt:="Invalid Maanden entry in clsPatient", Buttons:=vbExclamation
        m_Maanden = Null
    End If
End If

End Property

Public Property Get PatientID() As Variant

PatientID = m_PatientID

End Property

Public Property Let PatientID(ByVal v_PatientID As Variant)

m_PatientID = v_PatientID

End Property

Public Property Get GeboorteDatum() As Variant

GeboorteDatum = m_GeboorteDatum

End Property

Public Property Let GeboorteDatum(ByVal v_GeboorteDatum As Variant)

If IsDate(v_GeboorteDatum) And v_GeboorteDatum <= DateTime.Date Then
    m_GeboorteDatum = v_GeboorteDatum
Else
    m_GeboorteDatum = Null
End If

End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassMedicatieDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_GPK As String
Private m_ATC As String
Private m_TherapieGroep As String
Private m_TherapieSubgroep As String
Private m_Generiek As String
Private m_Vorm As String
Private m_Etiket As String
Private m_Sterkte As String
Private m_SterkteEenheid As String
Private m_Dosis As String
Private m_DosisEenheid As String
Private m_Routes As String
Private m_Indicaties As String

Public Property Get GPK() As String

    GPK = m_GPK

End Property

Public Property Let GPK(ByVal strGPK As String)

    m_GPK = strGPK

End Property

Public Property Get ATC() As String

    ATC = m_ATC

End Property

Public Property Let ATC(ByVal strATC As String)

    m_ATC = strATC

End Property

Public Property Get TherapieGroep() As String

    TherapieGroep = m_TherapieGroep

End Property

Public Property Let TherapieGroep(ByVal strTherapieGroep As String)

    m_TherapieGroep = strTherapieGroep

End Property

Public Property Get TherapieSubgroep() As String

    TherapieSubgroep = m_TherapieSubgroep

End Property

Public Property Let TherapieSubgroep(ByVal strTherapieSubgroep As String)

    m_TherapieSubgroep = strTherapieSubgroep

End Property

Public Property Get Generiek() As String

    Generiek = m_Generiek

End Property

Public Property Let Generiek(ByVal strGeneriek As String)

    m_Generiek = strGeneriek

End Property

Private Function CreateEtiket() As String

    Dim strEtiket As String
    
    strEtiket = Join(Array(m_Generiek, m_Vorm, m_Sterkte, m_SterkteEenheid), " ")
    CreateEtiket = strEtiket

End Function

Public Property Get Etiket() As String

    Dim strEtiket As String

    strEtiket = IIf(m_Etiket = vbNullString, CreateEtiket(), m_Etiket)
    Etiket = strEtiket

End Property

Public Property Let Etiket(ByVal strEtiket As String)

    m_Etiket = strEtiket

End Property

Public Property Get Vorm() As String

    Vorm = m_Vorm

End Property

Public Property Let Vorm(ByVal strVorm As String)
    
    m_Vorm = strVorm

End Property

Public Property Get Routes() As String

    Routes = m_Routes

End Property

Public Property Let Routes(ByVal strRoutes As String)

    m_Routes = strRoutes

End Property

Public Property Get Sterkte() As String

    Sterkte = m_Sterkte

End Property

Public Property Let Sterkte(ByVal strSterkte As String)

    m_Sterkte = strSterkte

End Property

Public Property Get SterkteEenheid() As String

    SterkteEenheid = m_SterkteEenheid

End Property

Public Property Let SterkteEenheid(ByVal strSterkteEenheid As String)

    m_SterkteEenheid = strSterkteEenheid

End Property

Public Property Get Dosis() As String

    Dosis = m_Dosis

End Property

Public Property Let Dosis(ByVal strDosis As String)

    m_Dosis = strDosis

End Property

Public Property Get DosisEenheid() As String

    DosisEenheid = m_DosisEenheid

End Property

Public Property Let DosisEenheid(ByVal strDosisEenheid As String)

    m_DosisEenheid = strDosisEenheid

End Property

Public Property Get Indicaties() As String

    Indicaties = m_Indicaties

End Property

Public Property Let Indicaties(ByVal strIndicaties As String)

    m_Indicaties = Strings.LCase(strIndicaties)

End Property

Public Function GetRoutes() As String()

    GetRoutes = Split(m_Routes, "/")

End Function

Public Function GetIndicaties() As String()

    GetIndicaties = Split(m_Indicaties, "/")
    
End Function

Public Function GetMedicamentText() As String

    GetMedicamentText = Me.Generiek & " " & Me.Vorm & " " & Me.Sterkte & " " & Me.SterkteEenheid

End Function

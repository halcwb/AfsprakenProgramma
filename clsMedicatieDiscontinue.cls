VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMedicatieDiscontinue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_GPK As Variant
Private m_ATC As Variant
Private m_TherapieGroep As Variant
Private m_TherapieSubgroep As Variant
Private m_Generiek As Variant
Private m_Etiket As Variant
Private m_Vorm As Variant
Private m_Routes As Variant
Private m_Sterkte As Variant
Private m_SterkteEenheid As Variant
Private m_Dosis As Variant
Private m_DosisEenheid As Variant
Private m_Indicaties As Variant


Public Property Get GPK() As Variant

    GPK = m_GPK

End Property

Public Property Let GPK(ByVal v_GPK As Variant)

    m_GPK = v_GPK

End Property

Public Property Get ATC() As Variant

    ATC = m_ATC

End Property

Public Property Let ATC(ByVal v_ATC As Variant)

    m_ATC = v_ATC

End Property

Public Property Get TherapieGroep() As Variant

    TherapieGroep = m_TherapieGroep

End Property

Public Property Let TherapieGroep(ByVal v_TherapieGroep As Variant)

    m_TherapieGroep = v_TherapieGroep

End Property

Public Property Get TherapieSubgroep() As Variant

    TherapieSubgroep = m_TherapieSubgroep

End Property

Public Property Let TherapieSubgroep(ByVal v_TherapieSubgroep As Variant)

    m_TherapieSubgroep = v_TherapieSubgroep

End Property

Public Property Get Generiek() As Variant

    Generiek = m_Generiek

End Property

Public Property Let Generiek(ByVal v_Generiek As Variant)

    m_Generiek = v_Generiek

End Property

Public Property Get Etiket() As Variant

    Etiket = m_Etiket

End Property

Public Property Let Etiket(ByVal v_Etiket As Variant)

    m_Etiket = v_Etiket

End Property

Public Property Get Vorm() As Variant

    Vorm = m_Vorm

End Property

Public Property Let Vorm(ByVal v_Vorm As Variant)
    
    m_Vorm = v_Vorm

End Property

Public Property Get routes() As Variant

    routes = m_Routes

End Property

Public Property Let routes(ByVal v_Routes As Variant)

    m_Routes = v_Routes

End Property

Public Property Get Sterkte() As Variant

    Sterkte = m_Sterkte

End Property

Public Property Let Sterkte(ByVal v_Sterkte As Variant)

    m_Sterkte = v_Sterkte

End Property

Public Property Get SterkteEenheid() As Variant

    SterkteEenheid = m_SterkteEenheid

End Property

Public Property Let SterkteEenheid(ByVal v_SterkteEenheid As Variant)

    m_SterkteEenheid = v_SterkteEenheid

End Property

Public Property Get Dosis() As Variant

    Dosis = m_Dosis

End Property

Public Property Let Dosis(ByVal v_Dosis As Variant)

    m_Dosis = v_Dosis

End Property

Public Property Get DosisEenheid() As Variant

    DosisEenheid = m_DosisEenheid

End Property

Public Property Let DosisEenheid(ByVal v_DosisEenheid As Variant)

    m_DosisEenheid = v_DosisEenheid

End Property

Public Property Get indicaties() As Variant

    indicaties = m_Indicaties

End Property

Public Property Let indicaties(ByVal v_Indicaties As Variant)

    m_Indicaties = v_Indicaties

End Property

Public Function GetPreparaat() As String

    GetPreparaat = m_Generiek & " " & m_Vorm & " " & m_Sterkte & " " & m_SterkteEenheid

End Function

Public Function GetRoutes() As String()

    GetRoutes = Split(m_Routes, "/")

End Function

Public Function GetIndicaties() As String()

    GetIndicaties = Split(m_Indicaties, "/")
    
End Function

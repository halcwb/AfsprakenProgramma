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
Private m_Product As String
Private m_Vorm As String
Private m_Etiket As String
Private m_Sterkte As Double
Private m_SterkteEenheid As String
Private m_DeelDose As Double
Private m_DoseEenheid As String
Private m_Route As String
Private m_Indicatie As String
Private m_Freq As String
Private m_HasDose As Boolean
Private m_PerDose As Boolean
Private m_PerKg As Boolean
Private m_PerM2 As Boolean
Private m_CalcDose As Double
Private m_KeerDose As Double
Private m_NormDose As Double
Private m_MinDose As Double
Private m_MaxDose As Double
Private m_AbsDose As Double
Private m_MaxKeer As Double
Private m_MaxConc As Double
Private m_OplVlst As String
Private m_OplVol As Double
Private m_MinTijd As Integer

Private m_IndicatieList As Collection
Private m_RouteList As Collection
Private m_FreqList As Collection

Public Property Get PerKg() As Boolean

    PerKg = m_PerKg

End Property

Public Property Let PerKg(ByVal blnPerKg As Boolean)

    If blnPerKg Then m_PerM2 = False
    m_PerKg = blnPerKg

End Property

Public Property Get PerM2() As Boolean

    PerM2 = m_PerM2

End Property

Public Property Let PerM2(ByVal blnPerM2 As Boolean)

    If blnPerM2 Then m_PerKg = False
    m_PerM2 = blnPerM2

End Property

Public Property Get PerDose() As Boolean

    PerDose = m_PerDose

End Property

Public Property Let PerDose(ByVal blnPerDose As Boolean)

    m_PerDose = blnPerDose

End Property

Public Property Get HasDose() As Boolean

    HasDose = m_HasDose

End Property

Public Property Let HasDose(ByVal blnHasDose As Boolean)

    m_HasDose = blnHasDose

End Property

Public Property Get MaxDose() As Double

    MaxDose = m_MaxDose

End Property

Public Property Let MaxDose(ByVal dblDose As Double)

    m_MaxDose = dblDose

End Property

Public Property Get NormDose() As Double

    NormDose = m_NormDose

End Property

Public Property Let NormDose(ByVal dblDose As Double)

    m_NormDose = dblDose

End Property

Public Property Get MaxKeer() As Double

    MaxKeer = m_MaxKeer

End Property

Public Property Let MaxKeer(ByVal dblDose As Double)

    m_MaxKeer = dblDose

End Property

Public Property Get CalcDose() As Double

    CalcDose = m_CalcDose

End Property

Public Property Let CalcDose(ByVal dblDose As Double)

    m_CalcDose = dblDose

End Property

Public Sub SetFreqList(ByVal strFreqList As String)

    FillCollection m_FreqList, strFreqList

End Sub

Public Function GetFreqList() As Collection

    Set GetFreqList = m_FreqList

End Function

Public Function GetFreqListString() As String

    Dim varItem As Variant
    Dim strList As String

    If m_FreqList Is Nothing Then
        GetFreqListString = vbNullString
    Else
        For Each varItem In m_FreqList
            strList = IIf(strList = vbNullString, varItem, strList & " ||" & varItem)
        Next
    End If
    
    GetFreqListString = strList
    
End Function

Public Property Get Freq() As String

    Freq = m_Freq

End Property

Public Property Let Freq(ByVal strFreq As String)

    m_Freq = strFreq

End Property

Public Property Get MaxConc() As Double

    MaxConc = m_MaxConc

End Property

Public Property Let MaxConc(ByVal dblConc As Double)

    m_MaxConc = dblConc

End Property

Public Property Get OplVlst() As String

    OplVlst = m_OplVlst

End Property

Public Property Let OplVlst(ByVal strVlst As String)

    m_OplVlst = strVlst

End Property

Public Property Get OplVol() As Double

    OplVol = m_OplVol

End Property

Public Property Let OplVol(ByVal dblVol As Double)

    m_OplVol = dblVol

End Property

Public Property Get MinTijd() As Integer

    MinTijd = m_MinTijd

End Property

Public Property Let MinTijd(ByVal intTime As Integer)

    m_MinTijd = intTime

End Property

Public Property Get MinDose() As Double
    
    MinDose = m_MinDose
    
End Property

Public Property Let MinDose(ByVal dblDose As Double)

    m_MinDose = dblDose

End Property

Public Property Get AbsDose() As Double
    
    AbsDose = m_AbsDose
    
End Property

Public Property Let AbsDose(ByVal dblDose As Double)

    m_AbsDose = dblDose

End Property

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

Public Property Get Product() As String

    Product = m_Product

End Property

Public Property Let Product(ByVal strProduct As String)

    m_Product = strProduct

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

    m_Generiek = MedDisc_CanonGen(strGeneriek)

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

Public Property Get Route() As String

    Route = m_Route

End Property

Public Property Let Route(ByVal strRoute As String)

    m_Route = strRoute

End Property

Public Property Get Sterkte() As Double

    Sterkte = m_Sterkte

End Property

Public Property Let Sterkte(ByVal strSterkte As Double)

    m_Sterkte = strSterkte

End Property

Public Property Get SterkteEenheid() As String

    SterkteEenheid = m_SterkteEenheid

End Property

Public Property Let SterkteEenheid(ByVal strSterkteEenheid As String)

    m_SterkteEenheid = strSterkteEenheid

End Property

Public Property Get DeelDose() As Double

    DeelDose = m_DeelDose

End Property

Public Property Let DeelDose(ByVal dblDose As Double)

    m_DeelDose = dblDose

End Property

Public Property Get DoseEenheid() As String

    DoseEenheid = m_DoseEenheid

End Property

Public Property Let DoseEenheid(ByVal strDosisEenheid As String)

    m_DoseEenheid = strDosisEenheid

End Property

Public Property Get Indicatie() As String

    Indicatie = m_Indicatie

End Property

Public Property Let Indicatie(ByVal strIndicatie As String)

    m_Indicatie = Strings.LCase(strIndicatie)

End Property

Private Sub FillCollection(objColl As Collection, ByVal strItems As String)

    Dim varItem As Variant
    Dim arrItem() As String
    Dim objC As ClassStringComparer
    
    strItems = IIf(ModString.StringIsZeroOrEmpty(strItems), vbNullString, strItems)
    strItems = Strings.LCase(strItems)
    arrItem = Split(strItems, "||")
    
    Set objColl = New Collection
    Set objC = New ClassStringComparer
    
    For Each varItem In arrItem
        CollectionAddDistinctStringNotEmpty objColl, Trim(CStr(varItem))
    Next
    
    ModCollection.CollectionSort objColl, objC
    If objColl.Count > 0 Then
        If objColl.Item(1) = vbNullString Then objColl.Remove (1)
    End If

End Sub

Public Sub SetRouteList(ByVal strRoutes As String)

    FillCollection m_RouteList, strRoutes

End Sub

Public Function GetRouteList() As Collection
    
    Set GetRouteList = m_RouteList

End Function

Public Sub SetIndicatieList(ByVal strIndicaties As String)
    
    m_Indicatie = strIndicaties
    FillCollection m_IndicatieList, strIndicaties
    
End Sub

Public Function GetIndicatieList() As Collection
    
    Set GetIndicatieList = m_IndicatieList
    
End Function

Public Function GetMedicamentText() As String

    GetMedicamentText = Me.Generiek & " " & Me.Vorm & " " & Me.Sterkte & " " & Me.SterkteEenheid

End Function

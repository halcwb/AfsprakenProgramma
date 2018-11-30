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
Private m_MainGroup As String
Private m_SubGroup As String
Private m_Generic As String
Private m_Product As String
Private m_Shape As String
Private m_Label As String
Private m_GenericQuantity As Double
Private m_GenericUnit As String
Private m_MultipleQuantity As Double
Private m_MultipleUnit As String
Private m_Route As String
Private m_Indication As String
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
Private m_AbsMaxDose As Double
Private m_MaxPerDose As Double
Private m_MaxConc As Double
Private m_Solution As String
Private m_SolutionVolume As Double
Private m_MinInfusionTime As Integer

Private m_IndicationList As Collection
Private m_RouteList As Collection
Private m_FreqList As Collection
Private m_DoseRules As Collection

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

Public Property Get MaxPerDose() As Double

    MaxPerDose = m_MaxPerDose

End Property

Public Property Let MaxPerDose(ByVal dblDose As Double)

    m_MaxPerDose = dblDose

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

Public Property Get Solution() As String

    Solution = m_Solution

End Property

Public Property Let Solution(ByVal strVlst As String)

    m_Solution = strVlst

End Property

Public Property Get SolutionVolume() As Double

    SolutionVolume = m_SolutionVolume

End Property

Public Property Let SolutionVolume(ByVal dblVol As Double)

    m_SolutionVolume = dblVol

End Property

Public Property Get MinInfusionTime() As Integer

    MinInfusionTime = m_MinInfusionTime

End Property

Public Property Let MinInfusionTime(ByVal intTime As Integer)

    m_MinInfusionTime = intTime

End Property

Public Property Get MinDose() As Double
    
    MinDose = m_MinDose
    
End Property

Public Property Let MinDose(ByVal dblDose As Double)

    m_MinDose = dblDose

End Property

Public Property Get AbsMaxDose() As Double
    
    AbsMaxDose = m_AbsMaxDose
    
End Property

Public Property Let AbsMaxDose(ByVal dblDose As Double)

    m_AbsMaxDose = dblDose

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

Public Property Get MainGroup() As String

    MainGroup = m_MainGroup

End Property

Public Property Let MainGroup(ByVal strTherapieGroep As String)

    m_MainGroup = strTherapieGroep

End Property

Public Property Get SubGroup() As String

    SubGroup = m_SubGroup

End Property

Public Property Let SubGroup(ByVal strTherapieSubgroep As String)

    m_SubGroup = strTherapieSubgroep

End Property

Public Property Get Generic() As String

    Generic = m_Generic

End Property

Public Property Let Generic(ByVal strGeneriek As String)

    m_Generic = MedDisc_CanonGen(strGeneriek)

End Property

Private Function CreateLabel() As String

    Dim strEtiket As String
    
    strEtiket = Join(Array(m_Generic, m_Shape, m_GenericQuantity, m_GenericUnit), " ")
    CreateLabel = strEtiket

End Function

Public Property Get Label() As String

    Dim strEtiket As String

    strEtiket = IIf(m_Label = vbNullString, CreateLabel(), m_Label)
    Label = strEtiket

End Property

Public Property Let Label(ByVal strEtiket As String)

    m_Label = strEtiket

End Property

Public Property Get Shape() As String

    Shape = m_Shape

End Property

Public Property Let Shape(ByVal strVorm As String)
    
    m_Shape = strVorm

End Property

Public Property Get Route() As String

    Route = m_Route

End Property

Public Property Let Route(ByVal strRoute As String)

    m_Route = strRoute

End Property

Public Property Get GenericQuantity() As Double

    GenericQuantity = m_GenericQuantity

End Property

Public Property Let GenericQuantity(ByVal strSterkte As Double)

    m_GenericQuantity = strSterkte

End Property

Public Property Get GenericUnit() As String

    GenericUnit = m_GenericUnit

End Property

Public Property Let GenericUnit(ByVal strSterkteEenheid As String)

    m_GenericUnit = strSterkteEenheid

End Property

Public Property Get MultipleQuantity() As Double

    MultipleQuantity = m_MultipleQuantity

End Property

Public Property Let MultipleQuantity(ByVal dblDose As Double)

    m_MultipleQuantity = dblDose

End Property

Public Property Get MultipleUnit() As String

    MultipleUnit = m_MultipleUnit

End Property

Public Property Let MultipleUnit(ByVal strDosisEenheid As String)

    m_MultipleUnit = strDosisEenheid

End Property

Public Property Get Indication() As String

    Indication = m_Indication

End Property

Public Property Let Indication(ByVal strIndicatie As String)

    m_Indication = Strings.LCase(strIndicatie)

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

Public Sub SetIndicationList(ByVal strIndicaties As String)
    
    m_Indication = strIndicaties
    FillCollection m_IndicationList, strIndicaties
    
End Sub

Public Function GetIndicationList() As Collection
    
    Set GetIndicationList = m_IndicationList
    
End Function

Public Function GetMedicamentText() As String

    GetMedicamentText = Me.Generic & " " & Me.Shape & " " & Me.GenericQuantity & " " & Me.GenericUnit

End Function

Public Property Get DoseRules() As Collection

    Set DoseRules = m_DoseRules

End Property

Public Property Let DoseRules(colRules As Collection)

    Set m_DoseRules = colRules

End Property


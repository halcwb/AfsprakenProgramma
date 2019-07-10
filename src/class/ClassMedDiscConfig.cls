VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassMedDiscConfig"
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

Private m_NeoNormDose As Double
Private m_NeoMinDose As Double
Private m_NeoMaxDose As Double

Private m_PedNormDose As Double
Private m_PedMinDose As Double
Private m_PedMaxDose As Double

Private m_PedAbsMaxDose As Double

Private m_PedMaxConc As Double
Private m_PedSolution As String
Private m_PedSolutionVolume As Double
Private m_PedMinInfusionTime As Integer

Private m_NeoMaxConc As Double
Private m_NeoSolution As String
Private m_NeoSoutionVolume As Double
Private m_NeoMinInfustionTime As Integer

Private m_IndicatieList As Collection
Private m_RouteList As Collection
Private m_FreqList As Collection

Public Property Get NeoMinDose() As Double
    
    NeoMinDose = m_NeoMinDose
    
End Property

Public Property Let NeoMinDose(ByVal dblDose As Double)

    m_NeoMinDose = dblDose

End Property

Public Property Get NeoMaxDose() As Double

    NeoMaxDose = m_NeoMaxDose

End Property

Public Property Let NeoMaxDose(ByVal dblDose As Double)

    m_NeoMaxDose = dblDose

End Property

Public Property Get NeoNormDose() As Double

    NeoNormDose = m_NeoNormDose

End Property

Public Property Let NeoNormDose(ByVal dblDose As Double)

    m_NeoNormDose = dblDose

End Property

Public Property Get PedMinDose() As Double
    
    PedMinDose = m_PedMinDose
    
End Property

Public Property Let PedMinDose(ByVal dblDose As Double)

    m_PedMinDose = dblDose

End Property

Public Property Get PedMaxDose() As Double

    PedMaxDose = m_PedMaxDose

End Property

Public Property Let PedMaxDose(ByVal dblDose As Double)

    m_PedMaxDose = dblDose

End Property

Public Property Get PedNormDose() As Double

    PedNormDose = m_PedNormDose

End Property

Public Property Let PedNormDose(ByVal dblDose As Double)

    m_PedNormDose = dblDose

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

    For Each varItem In m_FreqList
        strList = IIf(strList = vbNullString, varItem, strList & " ||" & varItem)
    Next
    
    GetFreqListString = strList
    
End Function

Public Property Get PedMaxConc() As Double

    PedMaxConc = m_PedMaxConc

End Property

Public Property Let PedMaxConc(ByVal dblConc As Double)

    m_PedMaxConc = dblConc

End Property

Public Property Get PedSolution() As String

    PedSolution = m_PedSolution

End Property

Public Property Let PedSolution(ByVal strVlst As String)

    m_PedSolution = strVlst

End Property

Public Property Get PedSolutionVolume() As Double

    PedSolutionVolume = m_PedSolutionVolume

End Property

Public Property Let PedSolutionVolume(ByVal dblConc As Double)

    m_PedSolutionVolume = dblConc

End Property

Public Property Get PedMinInfusionTime() As Integer

    PedMinInfusionTime = m_PedMinInfusionTime

End Property

Public Property Let PedMinInfusionTime(ByVal intTime As Integer)

    m_PedMinInfusionTime = intTime

End Property

Public Property Get NeoMaxConc() As Double

    NeoMaxConc = m_NeoMaxConc

End Property

Public Property Let NeoMaxConc(ByVal dblConc As Double)

    m_NeoMaxConc = dblConc

End Property

Public Property Get NeoSolution() As String

    NeoSolution = m_NeoSolution

End Property

Public Property Let NeoSolution(ByVal strVlst As String)

    m_NeoSolution = strVlst

End Property

Public Property Get NeoSoutionVolume() As Double

    NeoSoutionVolume = m_NeoSoutionVolume

End Property

Public Property Let NeoSoutionVolume(ByVal dblConc As Double)

    m_NeoSoutionVolume = dblConc

End Property

Public Property Get NeoMinInfustionTime() As Integer

    NeoMinInfustionTime = m_NeoMinInfustionTime

End Property

Public Property Let NeoMinInfustionTime(ByVal intTime As Integer)

    m_NeoMinInfustionTime = intTime

End Property

Public Property Get PedAbsMaxDose() As Double
    
    PedAbsMaxDose = m_PedAbsMaxDose
    
End Property

Public Property Let PedAbsMaxDose(ByVal dblDose As Double)

    m_PedAbsMaxDose = dblDose

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

Public Function GetRouteString() As String
    
    Dim strRoute As String
    Dim varRoute As Variant
    
    For Each varRoute In m_RouteList
        strRoute = IIf(strRoute = vbNullString, varRoute, strRoute & "||" & varRoute)
    Next

    GetRouteString = strRoute

End Function

Public Function GetRouteList() As Collection
    
    Set GetRouteList = m_RouteList

End Function

Public Sub SetIndicationList(ByVal strIndicaties As String)

    FillCollection m_IndicatieList, strIndicaties

End Sub

Public Function GetIndicationString() As String

    Dim varIndication As Variant
    Dim strIndication As String
    
    For Each varIndication In m_IndicatieList
        strIndication = IIf(strIndication = vbNullString, varIndication, strIndication & "||" & varIndication)
    Next
    
    GetIndicationString = strIndication

End Function

Public Function GetIndicationList() As Collection

    Set GetIndicationList = m_IndicatieList
    
End Function

Public Function GetMedication() As String

    GetMedication = Me.Generic & " " & Me.Shape & " " & Me.GenericQuantity & " " & Me.GenericUnit

End Function


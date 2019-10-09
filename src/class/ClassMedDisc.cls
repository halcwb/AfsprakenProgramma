VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassMedDisc"
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
Private m_TallMan As String
Private m_Product As String
Private m_Shape As String
Private m_Label As String
Private m_GenericQuantity As Double
Private m_GenericUnit As String
Private m_MultipleQuantity As Double
Private m_MultipleUnit As String
Private m_Route As String
Private m_Routes As String
Private m_Indication As String
Private m_Indications As String
Private m_Substance As String
Private m_Freq As String
Private m_Freqs As String
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
Private m_ProductDose As Double
Private m_DoseText As String

Private m_Substances As Collection
Private m_IndicationList As Collection
Private m_RouteList As Collection
Private m_FreqList As Collection
Private m_DoseRules As Collection
Private m_Doses As Collection
Private m_PICUSolution As ClassSolution
Private m_NICUSolution As ClassSolution

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

Public Property Get KeerDose() As Double

    KeerDose = m_KeerDose

End Property

Public Property Let KeerDose(ByVal dblDose As Double)

    Dim objSubst As ClassSubstance
    
    If m_Substances Is Nothing Then Exit Property
    
    m_KeerDose = dblDose
    For Each objSubst In m_Substances
        If objSubst.Substance = m_Substance And objSubst.Concentration > 0 Then
            ProductDose = dblDose / objSubst.Concentration
        End If
    Next

End Property

Public Property Get ProductDose() As Double

    ProductDose = m_ProductDose

End Property

Public Property Let ProductDose(ByVal dblDose As Double)

    m_ProductDose = dblDose

End Property

Public Sub SetFreqList(ByVal strFreqList As String)

    m_Freqs = strFreqList
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

Public Property Get DoseText() As String

    Dim intSubstCount As Integer
    Dim dblProdDose As Double
    Dim dblSubstDose As Double
    Dim strSubstDose As String
    Dim strSubstText As String
    Dim strSubstUnit As String
    Dim strProdUnit As String
    Dim objSubst As ClassSubstance
    
    If m_KeerDose > 0 Then
        intSubstCount = 0
        For Each objSubst In m_Substances
            If objSubst.Concentration > 0 Then
                intSubstCount = intSubstCount + 1
                If objSubst.Substance = m_Substance Then
                    dblProdDose = m_KeerDose / objSubst.Concentration
                End If
            End If
        Next
        
        If intSubstCount > 1 Then
            m_DoseText = vbNullString
            If UBound(Split(m_GenericUnit, "/")) >= 1 Then
                strSubstUnit = Split(m_GenericUnit, "/")(0)
                strProdUnit = Split(m_GenericUnit, "/")(1)
            End If
            For Each objSubst In m_Substances
                dblSubstDose = dblProdDose * objSubst.Concentration
                strSubstDose = objSubst.Substance & " " & DoubleToString(dblSubstDose) & " " & strSubstUnit
                strSubstText = strSubstText & IIf(strSubstText = vbNullString, strSubstDose, " + " & strSubstDose)
            Next
            
            m_DoseText = "dosering o.b.v. " & m_Substance & ": "
            m_DoseText = m_DoseText & m_Freq & " " & DoubleToString(dblProdDose) & " " & strProdUnit & " = " & strSubstText
        End If
    Else
        m_DoseText = ""
    End If

    DoseText = m_DoseText

End Property

Public Property Let DoseText(ByVal strDoseText As String)

    m_DoseText = strDoseText

End Property

Public Property Get Substance() As String

    Substance = m_Substance

End Property

Public Property Let Substance(ByVal strSubstance As String)

    m_Substance = strSubstance

End Property

Public Property Get Freq() As String

    Freq = m_Freq

End Property

Public Property Get Freqs() As String

    Freqs = m_Freqs

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

    If m_TallMan = vbNullString Then
        Generic = m_Generic
    Else
        Generic = m_TallMan
    End If

End Property

Public Property Let Generic(ByVal strGeneric As String)

    m_Generic = MedDisc_CanonGen(strGeneric)

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

Public Property Let Shape(ByVal strShape As String)
    
    m_Shape = strShape

End Property

Public Property Get Route() As String

    Route = m_Route

End Property

Public Property Let Route(ByVal strRoute As String)

    m_Route = strRoute

End Property

Public Property Get Routes() As String

    Routes = m_Routes

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

Public Property Get Indications() As String

    Indications = m_Indications

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

    m_Routes = strRoutes
    FillCollection m_RouteList, strRoutes

End Sub

Public Function GetRouteList() As Collection
    
    Set GetRouteList = m_RouteList

End Function

Public Sub SetIndicationList(ByVal strIndicaties As String)
    
    m_Indications = strIndicaties
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

    If m_DoseRules Is Nothing Then
        Set DoseRules = New Collection
    Else
        Set DoseRules = m_DoseRules
    End If

End Property

Public Property Let DoseRules(colRules As Collection)

    Set m_DoseRules = colRules
    
End Property

Public Property Get Substances() As Collection
    
    Dim arrSubst() As String
    Dim intI As Integer
    
    If m_Substances Is Nothing Then
        Set m_Substances = New Collection
        arrSubst = Split(Generic, "+")
        For intI = 0 To UBound(arrSubst)
            AddSubstance arrSubst(intI), 0
        Next
    End If
    
    Set Substances = m_Substances

End Property

Public Property Let Substances(colSubst As Collection)
    Dim objSubst As ClassSubstance
    
    For Each objSubst In colSubst
        AddSubstance objSubst.Substance, objSubst.Concentration
    Next
    
End Property

Public Sub AddSubstance(ByVal strSubst As String, ByVal dblConc As Double)
    
    Dim objSubst As ClassSubstance

    If m_Substances Is Nothing Then Set m_Substances = New Collection
    
    If Not HasSubstance(strSubst) Then
        Set objSubst = New ClassSubstance
        objSubst.Substance = strSubst
        objSubst.Concentration = dblConc
        m_Substances.Add objSubst
    End If

End Sub

Public Function HasSubstance(ByVal strSubst As String) As Boolean

    Dim objSubst As ClassSubstance
    
    strSubst = Trim(LCase(strSubst))
    If Not m_Substances Is Nothing Then
        For Each objSubst In m_Substances
            If objSubst.Substance = strSubst Then
                HasSubstance = True
                Exit Function
            End If
        Next
    End If
    
    HasSubstance = False

End Function

Public Sub SetTallMan(ByVal strTallMan As String)

    m_TallMan = strTallMan

End Sub

Public Sub SetPICUSolution(objSol As ClassSolution)
    
    Set m_PICUSolution = objSol

End Sub

Property Get PICUSolution() As ClassSolution

    If m_PICUSolution Is Nothing Then
        Set m_PICUSolution = New ClassSolution
    End If
    
    Set PICUSolution = m_PICUSolution

End Property

Public Sub SetNICUSolution(objSol As ClassSolution)

    Set m_NICUSolution = objSol

End Sub

Property Get NICUSolution() As ClassSolution

    If m_NICUSolution Is Nothing Then
        Set m_NICUSolution = New ClassSolution
    End If
    
    Set NICUSolution = m_NICUSolution

End Property

Public Property Get Doses() As Collection
    Dim objDose As ClassDose
    Dim varIndication As Variant
    Dim varRoute As Variant

    If m_Doses Is Nothing Then
        Set m_Doses = New Collection
        
        For Each varIndication In GetIndicationList()
            For Each varRoute In GetRouteList()
                Set objDose = New ClassDose
                
                objDose.Generic = Me.Generic
                objDose.Shape = Me.Shape
                objDose.Indication = varIndication
                objDose.Route = varRoute
                objDose.Unit = Me.MultipleUnit
                
                m_Doses.Add objDose
                
            Next
        Next
    End If
    
    Set Doses = m_Doses

End Property

Public Sub AddDose(ByVal objDose As ClassDose)

    Dim objItem As ClassDose
    
    Set objItem = GetSimilarDose(objDose.Department, _
                          objDose.Route, _
                          objDose.Indication, _
                          objDose.Gender, _
                          objDose.MinAgeMo, _
                          objDose.MaxAgeMo, _
                          objDose.MinWeightKg, _
                          objDose.MaxWeightKg, _
                          objDose.MinGestDays, _
                          objDose.MaxGestDays)
    
    With objItem
        .Frequencies = objDose.Frequencies
        .Unit = objDose.Unit
        .IsDosePerKg = objDose.IsDosePerKg
        .NormDose = objDose.NormDose
        .MinDose = objDose.MinDose
        .MaxDose = objDose.MaxDose
        .AbsMaxDose = objDose.AbsMaxDose
        .MaxPerDose = objDose.MaxPerDose
    End With
    
    If Not ModString.StringIsZeroOrEmpty(objItem.Generic) _
       And Not ModString.StringIsZeroOrEmpty(objItem.Shape) _
       And Not ModString.StringIsZeroOrEmpty(objItem.Route) _
       And Not ModString.StringIsZeroOrEmpty(objItem.Indication) Then
        m_Doses.Add objItem
    End If

End Sub

Public Function GetDose(ByVal strDep As String, ByVal strGender As String, ByVal dblAge As Double, ByVal intGestDays As Integer, ByVal dblWeight As Double) As ClassDose

    Dim objDose As ClassDose
    Dim blnIsDose
    Dim blnHasDose
    
    For Each objDose In Doses
        blnIsDose = True
        blnIsDose = blnIsDose And (strDep = vbNullString Or LCase(objDose.Department) = LCase(strDep))
        blnIsDose = blnIsDose And (strGender = vbNullString Or LCase(objDose.Gender) = LCase(strGender))
        blnIsDose = blnIsDose And (dblAge = 0 Or dblAge > objDose.MinAgeMo And (objDose.MaxAgeMo = 0 Or dblAge <= objDose.MaxAgeMo))
        blnIsDose = blnIsDose And (dblWeight = 0 Or dblWeight > objDose.MinWeightKg And (objDose.MaxWeightKg = 0 Or dblWeight <= objDose.MaxWeightKg))
        blnIsDose = blnIsDose And (intGestDays = 0 Or intGestDays > objDose.MinGestDays And (objDose.MaxGestDays = 0 Or intGestDays <= objDose.MaxGestDays))
        
        blnIsDose = blnIsDose And (LCase(Me.Indication) = objDose.Indication)
        blnIsDose = blnIsDose And (LCase(Me.Route) = objDose.Route)
                
        If blnIsDose Then
            blnHasDose = False
            blnHasDose = blnHasDose Or objDose.NormDose > 0
            blnHasDose = blnHasDose Or objDose.MinDose > 0
            blnHasDose = blnHasDose Or objDose.MaxDose > 0
            blnHasDose = blnHasDose Or objDose.AbsMaxDose > 0
            blnHasDose = blnHasDose Or objDose.MaxPerDose > 0
            
            If blnHasDose Then
                Set GetDose = objDose
                Exit Function
            End If
        End If
    Next

End Function

Private Function GetSimilarDose(ByVal strDep As String, _
                         ByVal strRoute As String, _
                         ByVal strIndication As String, _
                         ByVal strGender As String, _
                         ByVal dblMinAge As Double, _
                         ByVal dblMaxAge As Double, _
                         ByVal dblMinWeight As Double, _
                         ByVal dblMaxWeight As Double, _
                         ByVal intMinGest As Integer, _
                         ByVal intMaxGest As Integer) As ClassDose

    Dim blnHas As Boolean
    Dim objDose As ClassDose
    
    For Each objDose In Doses
        If objDose.Department = strDep _
           And objDose.Route = strRoute _
           And objDose.Indication = strIndication _
           And objDose.Gender = strGender _
           And objDose.MinAgeMo = dblMinAge _
           And objDose.MaxAgeMo = dblMaxAge _
           And objDose.MinWeightKg = dblMinWeight _
           And objDose.MaxWeightKg = dblMaxWeight _
           And objDose.MinGestDays = intMinGest _
           And objDose.MaxGestDays = intMaxGest Then
           
            Set GetSimilarDose = objDose
            Exit For
        End If
    Next

    Set objDose = New ClassDose
    
    With objDose
        .Generic = Me.Generic
        .Shape = Me.Shape
        
        .Department = strDep
        .Indication = strIndication
        .Route = strRoute
        .Gender = strGender
        .MinAgeMo = dblMinAge
        .MaxAgeMo = dblMaxAge
        .MinWeightKg = dblMinWeight
        .MaxWeightKg = dblMaxWeight
        .MinGestDays = intMinGest
        .MaxGestDays = intMaxGest
    End With
    
    Set GetSimilarDose = objDose
    
End Function
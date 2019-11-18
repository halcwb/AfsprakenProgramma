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
Private m_AdminDose As Double
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
Private m_HasSolutions As Boolean
Private m_IsActive As Boolean

Private m_Substances As Collection
Private m_IndicationList As Collection
Private m_RouteList As Collection
Private m_FreqList As Collection
Private m_DoseRules As Collection
Private m_Doses As Collection
Private m_Solutions As Collection

Public Property Get IsActive() As Boolean

    IsActive = m_IsActive

End Property

Public Property Let IsActive(ByVal blnIsActive As Boolean)

    m_IsActive = blnIsActive

End Property

Public Property Get HasSolutions() As Boolean

    HasSolutions = m_HasSolutions

End Property

Public Property Let HasSolutions(ByVal blnHasSolutions As Boolean)

    m_HasSolutions = blnHasSolutions

End Property

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

    m_MaxPerDose = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get CalcDose() As Double

    CalcDose = m_CalcDose

End Property

Public Property Let CalcDose(ByVal dblDose As Double)

    m_CalcDose = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get AdminDose() As Double

    AdminDose = m_AdminDose

End Property

Public Property Let AdminDose(ByVal dblDose As Double)

    Dim objSubst As ClassSubstance
    
    If m_Substances Is Nothing Or dblDose < 0 Then Exit Property
    
    m_AdminDose = dblDose
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

    m_Freqs = Trim(strFreqList)
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
    
    If m_AdminDose > 0 Then
        intSubstCount = 0
        For Each objSubst In m_Substances
            If objSubst.Concentration > 0 Then
                intSubstCount = intSubstCount + 1
                If objSubst.Substance = m_Substance Then
                    dblProdDose = m_AdminDose / objSubst.Concentration
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

    m_DoseText = Trim(strDoseText)

End Property

Public Property Get Substance() As String

    Substance = m_Substance

End Property

Public Property Let Substance(ByVal strSubstance As String)

    m_Substance = Trim(strSubstance)

End Property

Public Property Get Freq() As String

    Freq = m_Freq

End Property

Public Property Get Freqs() As String

    Freqs = m_Freqs

End Property

Public Property Let Freq(ByVal strFreq As String)

    m_Freq = Trim(strFreq)

End Property

Public Property Get MaxConc() As Double

    MaxConc = m_MaxConc

End Property

Public Property Let MaxConc(ByVal dblConc As Double)

    m_MaxConc = IIf(dblConc < 0, 0, dblConc)

End Property

Public Property Get Solution() As String

    Solution = m_Solution

End Property

Public Property Let Solution(ByVal strVlst As String)

    m_Solution = Trim(strVlst)

End Property

Public Property Get SolutionVolume() As Double

    SolutionVolume = m_SolutionVolume

End Property

Public Property Let SolutionVolume(ByVal dblVol As Double)

    m_SolutionVolume = IIf(dblVol < 0, 0, dblVol)

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

    m_MinDose = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get AbsMaxDose() As Double
    
    AbsMaxDose = m_AbsMaxDose
    
End Property

Public Property Let AbsMaxDose(ByVal dblDose As Double)

    m_AbsMaxDose = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get GPK() As String

    GPK = m_GPK

End Property

Public Property Let GPK(ByVal strGPK As String)

    m_GPK = Trim(strGPK)

End Property

Public Property Get ATC() As String

    ATC = m_ATC

End Property

Public Property Let ATC(ByVal strATC As String)

    m_ATC = Trim(strATC)

End Property

Public Property Get Product() As String

    Product = m_Product

End Property

Public Property Let Product(ByVal strProduct As String)

    m_Product = Trim(strProduct)

End Property

Public Property Get MainGroup() As String

    MainGroup = m_MainGroup

End Property

Public Property Let MainGroup(ByVal strTherapieGroep As String)

    m_MainGroup = Trim(strTherapieGroep)

End Property

Public Property Get SubGroup() As String

    SubGroup = m_SubGroup

End Property

Public Property Let SubGroup(ByVal strTherapieSubgroep As String)

    m_SubGroup = Trim(strTherapieSubgroep)

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
    
    m_Shape = Trim(strShape)

End Property

Public Property Get Route() As String

    Route = m_Route

End Property

Public Property Let Route(ByVal strRoute As String)

    m_Route = Trim(strRoute)

End Property

Public Property Get Routes() As String

    Routes = m_Routes

End Property

Public Property Get GenericQuantity() As Double

    GenericQuantity = m_GenericQuantity

End Property

Public Property Let GenericQuantity(ByVal dblQty As Double)

    m_GenericQuantity = IIf(dblQty < 0, 0, dblQty)

End Property

Public Property Get GenericUnit() As String

    GenericUnit = m_GenericUnit

End Property

Public Property Let GenericUnit(ByVal strUnit As String)

    m_GenericUnit = Trim(strUnit)

End Property

Public Property Get MultipleQuantity() As Double

    MultipleQuantity = m_MultipleQuantity

End Property

Public Property Let MultipleQuantity(ByVal dblDose As Double)

    m_MultipleQuantity = IIf(dblDose < 0, 0, dblDose)

End Property

Public Property Get MultipleUnit() As String

    MultipleUnit = m_MultipleUnit

End Property

Public Property Let MultipleUnit(ByVal strUnit As String)

    m_MultipleUnit = Trim(strUnit)

End Property

Public Property Get Indication() As String

    Indication = m_Indication

End Property

Public Property Get Indications() As String

    Indications = m_Indications

End Property

Public Property Let Indication(ByVal strIndicatie As String)

    m_Indication = Strings.LCase(Trim(strIndicatie))

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

    m_Routes = Trim(strRoutes)
    FillCollection m_RouteList, m_Routes

End Sub

Public Function GetRouteList() As Collection
    
    Set GetRouteList = m_RouteList

End Function

Public Sub SetIndicationList(ByVal strIndicaties As String)
    
    m_Indications = Trim(strIndicaties)
    m_Indication = m_Indications
    FillCollection m_IndicationList, m_Indications
    
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

    m_TallMan = Trim(strTallMan)

End Sub

Public Sub SetSolutions(objSols As Collection)

    Set m_Solutions = objSols

End Sub

Property Get Solutions() As Collection

    Dim objSol As ClassSolution
    
    If m_Solutions Is Nothing Then
        Set m_Solutions = New Collection
        
        If ContainsCaseInsensitive(Routes, "iv") Then
            Set objSol = New ClassSolution
            With objSol
                .Department = "NICU"
                .Generic = Generic
                .Shape = Shape
            End With
            m_Solutions.Add objSol
            
            Set objSol = New ClassSolution
            With objSol
                .Department = "PICU"
                .Generic = Generic
                .Shape = Shape
            End With
            m_Solutions.Add objSol
            
            Set objSol = New ClassSolution
            With objSol
                .Department = "AICU"
                .Generic = Generic
                .Shape = Shape
            End With
            m_Solutions.Add objSol
        End If
    
    End If
    
    Set Solutions = m_Solutions

End Property

Public Sub SetDoses(objDoses As Collection)

    Set m_Doses = objDoses

End Sub


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
    
    If objDose.IsEmpty Then Exit Sub
    
    If objDose.Generic = "meropenem" Then
        Debug.Print "Meropenem dosering"
    End If
    
    Set objItem = GetSimilarDose(objDose.Department, _
                                 objDose.Route, _
                                 objDose.Indication, _
                                 objDose.Gender, _
                                 objDose.MinAgeMo, _
                                 objDose.MaxAgeMo, _
                                 objDose.MinWeightKg, _
                                 objDose.MaxWeightKg, _
                                 objDose.MinGestDays, _
                                 objDose.MaxGestDays, _
                                 objDose.MinPMDays, _
                                 objDose.MaxPMDays)
    
    With objItem
        .Frequencies = objDose.Frequencies
        .Unit = objDose.Unit
        .IsDosePerKg = objDose.IsDosePerKg
        .IsDosePerM2 = objDose.IsDosePerM2
        .NormDose = objDose.NormDose
        .MinDose = objDose.MinDose
        .MaxDose = objDose.MaxDose
        .AbsMaxDose = objDose.AbsMaxDose
        .MaxPerDose = objDose.MaxPerDose
    End With
    

End Sub

Public Function GetDose(ByVal strDep As String, ByVal strGender As String, ByVal dblAge As Double, ByVal lngGestDays As Long, ByVal lngPMAge As Long, ByVal dblWeight As Double) As ClassDose

    Dim objDose As ClassDose
    Dim blnIsDose
    
    Debug.Print "=== START GETTING DOSE ===" & vbNewLine
    Debug.Print "Department: " & strDep
    Debug.Print "Gender: " & strGender
    Debug.Print "Age: " & dblAge
    Debug.Print "GestAge: " & lngGestDays
    Debug.Print "PMAge: " & lngPMAge
    Debug.Print "Weight: " & dblWeight
    Debug.Print "Indicatation: " & Indication
    Debug.Print "Route: " & Route & vbNewLine
    
    For Each objDose In Doses
    
        If Not objDose.IsEmpty() Then
            Debug.Print "Considering: "
            Debug.Print objDose.ToString() & vbNewLine
            
            blnIsDose = True
            blnIsDose = blnIsDose And (strDep = vbNullString Or LCase(objDose.Department) = LCase(strDep))
            blnIsDose = blnIsDose And (strGender = vbNullString Or LCase(objDose.Gender) = LCase(strGender))
            blnIsDose = blnIsDose And (dblAge >= objDose.MinAgeMo And (objDose.MaxAgeMo = 0 Or dblAge < objDose.MaxAgeMo))
            blnIsDose = blnIsDose And (dblWeight = 0 Or dblWeight >= objDose.MinWeightKg And (objDose.MaxWeightKg = 0 Or dblWeight < objDose.MaxWeightKg))
            blnIsDose = blnIsDose And (lngGestDays = 0 Or lngGestDays >= objDose.MinGestDays And (objDose.MaxGestDays = 0 Or lngGestDays < objDose.MaxGestDays))
            blnIsDose = blnIsDose And (lngPMAge = 0 Or lngPMAge >= objDose.MinPMDays And (objDose.MaxPMDays = 0 Or lngPMAge < objDose.MaxPMDays))
            
            blnIsDose = blnIsDose And (Trim(LCase(Me.Indication)) = Trim(LCase(objDose.Indication)))
            blnIsDose = blnIsDose And (Trim(LCase(Me.Route)) = Trim(LCase(objDose.Route)))
                    
            If blnIsDose Then
                
                Debug.Print "=== PICKED DOSE ==" & vbNewLine
                Debug.Print objDose.ToString()
                
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
                                ByVal intMaxGest As Integer, _
                                ByVal intMinPM As Integer, _
                                ByVal intMaxPM As Integer) As ClassDose

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
           And objDose.MaxGestDays = intMaxGest _
           And objDose.MinPMDays = intMinPM _
           And objDose.MaxPMDays = intMaxPM Then
           
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
        .MinPMDays = intMinPM
        .MaxPMDays = intMaxPM
    End With
    
    If Not ModString.StringIsZeroOrEmpty(objDose.Generic) _
       And Not ModString.StringIsZeroOrEmpty(objDose.Shape) _
       And Not ModString.StringIsZeroOrEmpty(objDose.Route) _
       And Not ModString.StringIsZeroOrEmpty(objDose.Indication) Then
        m_Doses.Add objDose
        If objDose.Generic = "meropenem" Then
            Debug.Print "Added: " & objDose.ToString
        End If
    End If
    
    Set GetSimilarDose = objDose
    
End Function

Public Sub AddSolution(ByVal objSol As ClassSolution)

    Dim objItem As ClassSolution
    
    Set objItem = GetSimilarSolution(objSol.Department, objSol.MinGenericQuantity, objSol.MaxGenericQuantity)
    
    With objItem
        .Solutions = objSol.Solutions
        .SolutionVolume = objSol.SolutionVolume
        .MinConc = objSol.MinConc
        .MaxConc = objSol.MaxConc
        .MinInfusionTime = objSol.MinInfusionTime
    End With
    
    If Not ModString.StringIsZeroOrEmpty(objItem.Generic) _
       And Not ModString.StringIsZeroOrEmpty(objItem.Shape) Then
        m_Solutions.Add objItem
    End If

End Sub

Public Function GetSolution(ByVal strDep As String, ByVal dblQty As Double) As ClassSolution

    Dim objSolution As ClassSolution
    Dim blnIsSolution
    Dim blnHasSolution
    
    For Each objSolution In Solutions
        blnIsSolution = True
        blnIsSolution = blnIsSolution And (strDep = vbNullString Or LCase(objSolution.Department) = LCase(strDep))
        blnIsSolution = blnIsSolution And (dblQty = 0 Or dblQty > objSolution.MinGenericQuantity And (objSolution.MaxGenericQuantity = 0 Or dblQty <= objSolution.MaxGenericQuantity))
                
        If blnIsSolution Then
            blnHasSolution = False
            blnHasSolution = blnHasSolution Or objSolution.MinConc > 0
            blnHasSolution = blnHasSolution Or objSolution.MaxConc > 0
            blnHasSolution = blnHasSolution Or objSolution.SolutionVolume > 0
            blnHasSolution = blnHasSolution Or objSolution.MinInfusionTime > 0
            
            If blnHasSolution Then
                Set GetSolution = objSolution
                Exit Function
            End If
        End If
    Next

End Function

Private Function GetSimilarSolution(ByVal strDep As String, ByVal dblMinQty As Double, ByVal dblMaxQty As Double) As ClassSolution

    Dim blnHas As Boolean
    Dim objSol As ClassSolution
    
    For Each objSol In Solutions
        If objSol.Department = strDep _
           And objSol.MinGenericQuantity = dblMinQty _
           And objSol.MaxGenericQuantity = dblMaxQty Then
           
            Set GetSimilarSolution = objSol
            Exit For
        End If
    Next

    Set objSol = New ClassSolution
    
    With objSol
        .Generic = Me.Generic
        .Shape = Me.Shape
        
        .Department = strDep
        .MinGenericQuantity = dblMinQty
        .MaxGenericQuantity = dblMaxQty
    End With
    
    Set GetSimilarSolution = objSol
    
End Function

Public Function Clone() As ClassMedDisc

    Dim objMed As ClassMedDisc
    
    Set objMed = New ClassMedDisc
    
    With objMed
        .AbsMaxDose = AbsMaxDose
        .ATC = ATC
        .DoseRules = DoseRules
        .SetDoses Doses
        .SetFreqList Freqs
        .Generic = Generic
        .GenericQuantity = GenericQuantity
        .GenericUnit = GenericUnit
        .GPK = GPK
        .HasSolutions = HasSolutions
        .SetIndicationList Indications
        .IsActive = IsActive
        .AdminDose = AdminDose
        .Label = Label
        .MainGroup = MainGroup
        .MaxConc = MaxConc
        .MaxDose = MaxDose
        .MaxPerDose = MaxPerDose
        .MinDose = MinDose
        .MinInfusionTime = MinInfusionTime
        .MultipleQuantity = MultipleQuantity
        .MultipleUnit = MultipleUnit
        .NormDose = NormDose
        .PerDose = PerDose
        .PerKg = PerKg
        .PerM2 = PerM2
        .Product = Product
        .ProductDose = ProductDose
        .SetRouteList Routes
        .Shape = Shape
        .Solution = Solution
        .SetSolutions Solutions
        .SolutionVolume = SolutionVolume
        .SubGroup = SubGroup
        .Substances = Substances
    End With
    
    Set Clone = objMed

End Function



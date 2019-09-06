Attribute VB_Name = "ModWeb"
Option Explicit


Private Const constHost = "http://vpxap-meta01.ds.umcutrecht.nl"
Private Const constUrl As String = "/request?age=AGE&wth=WTH&hgt=HGT&gpk=GPK&gen=GEN&shp=SHP&rte=RTE&unt=UNT"

Public Sub Web_RetrieveMedicationRules(objMed As ClassMedicatieDisc)

    If objMed.GPK = "" Then Exit Sub

    Dim strAge As String
    Dim strWTH As String
    Dim strHGT As String
    Dim strGPK As String
    Dim strGen As String
    Dim strSHP As String
    Dim strRTE As String
    Dim strUNT As String
    
    Dim strUrl As String
    
    Dim objClient As New WebClient
    Dim objResponse As WebResponse
    
    objClient.BaseUrl = constHost & "/genform"
        
    strAge = Patient_CorrectedAgeInMo()
    
    strWTH = Val(ModPatient.Patient_GetWeight())
    strHGT = Val(ModPatient.Patient_GetHeight())
    
    strGPK = objMed.GPK
    strGen = objMed.Generic
    strSHP = objMed.Shape
    strRTE = objMed.Route
    strUNT = objMed.MultipleUnit
    
    If strRTE = vbNullString Then
        ModMessage.ShowMsgBoxInfo "Geef de route op"
        Exit Sub
    End If
    
    strUrl = constUrl
    strUrl = Replace(strUrl, "AGE", strAge)
    strUrl = Replace(strUrl, "WTH", strWTH)
    strUrl = Replace(strUrl, "HGT", strHGT)
    strUrl = Replace(strUrl, "GPK", strGPK)
    strUrl = Replace(strUrl, "GEN", strGen)
    strUrl = Replace(strUrl, "SHP", strSHP)
    strUrl = Replace(strUrl, "RTE", strRTE)
    strUrl = Replace(strUrl, "UNT", strUNT)
    
    ModUtils.CopyToClipboard objClient.BaseUrl & strUrl
    
    Set objResponse = objClient.GetJson(strUrl)
    
    If objResponse.StatusCode = OK Then
        ProcessJson objResponse, objMed
    Else
        ModMessage.ShowMsgBoxExclam "Kan de doseer informatie niet ophalen!. Probeer het nog een keer of neem anders contact op met de helpdesk"
        ModLog.LogError Err, "Fout bij ophalen van doseer informatie: " & objResponse.StatusDescription
    End If

End Sub

Private Sub Test_GetJson()

    Dim objClient As New WebClient
    Dim objResponse As WebResponse
    Dim objMed As New ClassMedicatieDisc
    
    objClient.BaseUrl = "http://iis2503.ds.umcutrecht.nl/genform"
    
    Set objResponse = objClient.GetJson("/request?age=0&wth=1.0&hgt=50&gpk=3689&rte=iv")
    
    ModMessage.ShowMsgBoxInfo objResponse.Content
    ProcessJson objResponse, objMed
    ModMessage.ShowMsgBoxInfo objMed.Label

End Sub

Private Sub ProcessJson(objResponse As WebResponse, objMed As ClassMedicatieDisc)

    Dim colRules As Collection
    Dim objRule As ClassDoseRule
        
    Dim objDict As Dictionary
    Dim colJson As Collection
    Dim strJson As String
        
    strJson = objResponse.Content
    Set objDict = JsonConverter.ParseJson(strJson)
        
    objMed.ATC = NotEmpty(objMed.ATC, objDict("atc"))
    objMed.MainGroup = NotEmpty(objMed.MainGroup, objDict("therapyGroup"))
    objMed.SubGroup = NotEmpty(objMed.SubGroup, objDict("therapySubGroup"))
    objMed.Generic = NotEmpty(objMed.Generic, objDict("generic"))
    objMed.Product = NotEmpty(objMed.Product, objDict("tradeProduct"))
    objMed.Shape = NotEmpty(objMed.Shape, objDict("shape"))
    objMed.Label = NotEmpty(objMed.Label, objDict("label"))
    objMed.GenericQuantity = NotEmpty(objMed.GenericQuantity, objDict("concentration"))
    objMed.GenericUnit = NotEmpty(objMed.GenericUnit, objDict("concentrationUnit"))
    objMed.MultipleQuantity = NotEmpty(objMed.MultipleQuantity, objDict("multiple"))
    objMed.MultipleUnit = NotEmpty(objMed.MultipleUnit, objDict("multipleUnit"))
    objMed.Indication = NotEmpty(objMed.Indication, objDict("indication"))
        
    Set colJson = objDict("rules")
        
    Set colRules = New Collection
    For Each objDict In colJson
        Set objRule = New ClassDoseRule
        
        objRule.Substance = objDict("substance")
        objRule.Freq = objDict("frequency")
        
        objRule.NormDose = objDict("normTotalDose")
        objRule.MinDose = objDict("minTotalDose")
        objRule.MaxDose = objDict("maxTotalDose")
        objRule.MaxPerDose = objDict("maxPerDose")
        objRule.AbsMaxDose = objDict("maxTotalDose")
        
        If objDict("normTotalDosePerKg") > 0 Or objDict("minTotalDosePerKg") > 0 Or objDict("maxTotalDosePerKg") > 0 Then
            objRule.PerKg = True
            objRule.NormDose = objDict("normTotalDosePerKg")
            objRule.MinDose = objDict("minTotalDosePerKg")
            objRule.MaxDose = objDict("maxTotalDosePerKg")
        End If
            
        If objDict("normTotalDosePerM2") > 0 Or objDict("minTotalDosePerM2") > 0 Or objDict("maxTotalDosePerM2") > 0 Then
            objRule.PerM2 = True
            objRule.NormDose = objDict("normTotalDosePerM2")
            objRule.MinDose = objDict("minTotalDosePerM2")
            objRule.MaxDose = objDict("maxTotalDosePerM2")
        End If
        
        If Not objMed.HasSubstance(objRule.Substance) Then objMed.AddSubstance objRule.Substance, 0
        colRules.Add objRule
    Next

    If colRules.Count = 1 Then
        Set objRule = colRules(1)
        
        objMed.Substance = objRule.Substance
        
        objMed.PerKg = objRule.PerKg
        objMed.PerM2 = objRule.PerM2
        
        objMed.SetFreqList objRule.Freq
        
        objMed.NormDose = objRule.NormDose
        objMed.MinDose = objRule.MinDose
        objMed.MaxDose = objRule.MaxDose
        objMed.MaxPerDose = objRule.MaxPerDose
        objMed.AbsMaxDose = objRule.AbsMaxDose
        
    ElseIf colRules.Count = 0 Then
    
        objMed.PerKg = True
        
    End If
    
    If objMed.Substances.Count = 1 Then
        objMed.Substances.Item(1).Concentration = objMed.GenericQuantity
    End If
    
    objMed.DoseRules = colRules

End Sub

Private Function NotEmpty(ByVal varVal1, ByVal varVal2) As Variant

    If varVal1 = "" Or varVal1 = 0 Then
        NotEmpty = varVal2
    Else
        NotEmpty = varVal1
    End If

End Function

Private Sub Test_NotEmpty()

    ModMessage.ShowMsgBoxInfo NotEmpty("", "test2")
    ModMessage.ShowMsgBoxInfo NotEmpty("test1", "test2")
    ModMessage.ShowMsgBoxInfo NotEmpty("test1", "")
    ModMessage.ShowMsgBoxInfo NotEmpty(0, 2)
    ModMessage.ShowMsgBoxInfo NotEmpty(1, 2)
    ModMessage.ShowMsgBoxInfo NotEmpty(1, 0)

End Sub

Attribute VB_Name = "ModWeb"
Option Explicit

Private Const constUrl As String = "/request?bty=BTY&btm=BTM&btd=BTD&wth=WTH&hgt=HGT&gpk=GPK&rte=RTE&unt=UNT"

Public Sub Web_RetrieveMedicationRules(objMed As ClassMedicatieDisc)

    If objMed.GPK = "" Then Exit Sub

    Dim strBTY As String
    Dim strBTM As String
    Dim strBTD As String
    Dim strWTH As String
    Dim strHGT As String
    Dim strGPK As String
    Dim strRTE As String
    Dim strUNT As String
    
    Dim strUrl As String
    
    Dim objClient As New WebClient
    Dim objResponse As WebResponse
    
    objClient.BaseUrl = "http://iis2503.ds.umcutrecht.nl/genform"
        
    strBTY = ModDate.DateYear(ModPatient.Patient_BirthDate())
    strBTM = ModDate.DateMonth(ModPatient.Patient_BirthDate())
    strBTD = ModDate.DateDay(ModPatient.Patient_BirthDate())
    
    strWTH = Val(ModPatient.Patient_GetWeight())
    strHGT = Val(ModPatient.Patient_GetHeight())
    
    strGPK = objMed.GPK
    strRTE = objMed.Route
    strUNT = objMed.MultipleUnit
    
    If strRTE = vbNullString Then
        ModMessage.ShowMsgBoxInfo "Geef de route op"
        Exit Sub
    End If
    
    strUrl = constUrl
    strUrl = Replace(strUrl, "BTY", strBTY)
    strUrl = Replace(strUrl, "BTM", strBTM)
    strUrl = Replace(strUrl, "BTD", strBTD)
    strUrl = Replace(strUrl, "WTH", strWTH)
    strUrl = Replace(strUrl, "HGT", strHGT)
    strUrl = Replace(strUrl, "GPK", strGPK)
    strUrl = Replace(strUrl, "RTE", strRTE)
    strUrl = Replace(strUrl, "UNT", strUNT)
    
    ModUtils.CopyToClipboard objClient.BaseUrl & strUrl
    
    Set objResponse = objClient.GetJson(strUrl)
    
    If objResponse.StatusCode = Ok Then
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
    
    objClient.BaseUrl = "http://localhost:8080"
    
    Set objResponse = objClient.GetJson("/request?bty=2017&btm=1&btd=1&wth=10&hgt=70&gpk=9504&rte=or")
    
    ProcessJson objResponse, objMed
    ModMessage.ShowMsgBoxInfo objMed.Label

End Sub

'birthYear 2017
'birthMonth 1
'birthDay 1
'weightKg 10
'birthWeightGram 0
'lengthCm 0
'gender ""
'gestAgeWeeks 0
'gestAgeDays 0
'GPK "9504"
'ATC "N02BE01 "
'therapyGroup "ANALGETICA"
'therapySubGroup "OVERIGE ANALGETICA EN ANTIPYRETICA"
'generic "PARACETAMOL"
'tradeProduct ""
'Shape "STROOP"
'Label "PARACETAMOL 24MG/ML STROOP"
'concentration 0
'concentrationUnit ""
'multiple 0
'multipleUnit ""
'Route "or"
'indication ""
'Frequency "3 x / dag||antenoctum||1 x / dag||2 x / dag||4 x / dag"
'PerDose False
'PerKg True
'PerM2 False
'NormDose 0
'MinDose 0
'MaxDose 90.048
'absMaxTotal 480
'absMaxPerDose 120

Private Sub ProcessJson(objResponse As WebResponse, objMed As ClassMedicatieDisc)

    Dim objDict As Dictionary
    Dim strJson As String
    Dim dblVal As Double
    Dim blnChange As Boolean
    
'    ModMessage.ShowMsgBoxInfo objResponse.Content
    
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
        
    dblVal = objMed.NormDose
    objMed.NormDose = NotEmpty(objMed.NormDose, objDict("normDose"))
    blnChange = Not objMed.NormDose = dblVal
    
    dblVal = objMed.MinDose
    objMed.MinDose = NotEmpty(objMed.MinDose, objDict("minDose"))
    blnChange = blnChange Or (Not dblVal = objMed.MinDose)
    
    dblVal = objMed.MaxDose
    objMed.MaxDose = NotEmpty(objMed.MaxDose, objDict("maxDose"))
    blnChange = blnChange Or (Not dblVal = objMed.MaxDose)

    If blnChange Then
        objMed.PerDose = objDict("perDose")
        objMed.PerKg = objDict("perKg")
        objMed.PerM2 = objDict("perM2")
    End If
    
    objMed.AbsMaxDose = NotEmpty(objMed.AbsMaxDose, objDict("absMaxTotal"))
    objMed.MaxPerDose = NotEmpty(objMed.MaxPerDose, objDict("absMaxPerDose"))

    If objMed.Freq = "" Then objMed.SetFreqList objDict("frequency")
    If objMed.Route = "" Then objMed.SetRouteList objDict("route")

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

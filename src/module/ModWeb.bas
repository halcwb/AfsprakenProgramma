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
    
    objClient.BaseUrl = "http://localhost:8080"
        
    strBTY = ModDate.DateYear(ModPatient.Patient_BirthDate())
    strBTM = ModDate.DateMonth(ModPatient.Patient_BirthDate())
    strBTD = ModDate.DateDay(ModPatient.Patient_BirthDate())
    
    strWTH = ModPatient.GetGewichtFromRange()
    strHGT = ModPatient.GetLengteFromRange()
    
    strGPK = objMed.GPK
    strRTE = objMed.Route
    strUNT = objMed.DoseEenheid
    
    strUrl = constUrl
    strUrl = Replace(strUrl, "BTY", strBTY)
    strUrl = Replace(strUrl, "BTM", strBTM)
    strUrl = Replace(strUrl, "BTD", strBTD)
    strUrl = Replace(strUrl, "WTH", strWTH)
    strUrl = Replace(strUrl, "HGT", strHGT)
    strUrl = Replace(strUrl, "GPK", strGPK)
    strUrl = Replace(strUrl, "RTE", strRTE)
    strUrl = Replace(strUrl, "UNT", strUNT)
    
    Set objResponse = objClient.GetJson(strUrl)
    ProcessJson objResponse, objMed

End Sub


Private Sub Test_GetJson()

    Dim objClient As New WebClient
    Dim objResponse As WebResponse
    Dim objMed As New ClassMedicatieDisc
    
    objClient.BaseUrl = "http://localhost:8080"
    
    Set objResponse = objClient.GetJson("/request?bty=2017&btm=1&btd=1&wth=10&hgt=70&gpk=9504&rte=or")
    
    ProcessJson objResponse, objMed
    ModMessage.ShowMsgBoxInfo objMed.Etiket

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
    
'    ModMessage.ShowMsgBoxInfo objResponse.Content
    
    strJson = objResponse.Content
    Set objDict = JsonConverter.ParseJson(strJson)
        
    objMed.ATC = NotEmpty(objMed.ATC, objDict("atc"))
    objMed.TherapieGroep = NotEmpty(objMed.TherapieGroep, objDict("therapyGroup"))
    objMed.TherapieSubgroep = NotEmpty(objMed.TherapieSubgroep, objDict("therapySubGroup"))
    objMed.Generiek = NotEmpty(objMed.Generiek, objDict("generic"))
    objMed.Product = NotEmpty(objMed.Product, objDict("tradeProduct"))
    objMed.Vorm = NotEmpty(objMed.Vorm, objDict("shape"))
    objMed.Etiket = NotEmpty(objMed.Etiket, objDict("label"))
    objMed.Sterkte = NotEmpty(objMed.Sterkte, objDict("concentration"))
    objMed.SterkteEenheid = NotEmpty(objMed.SterkteEenheid, objDict("concentrationUnit"))
    objMed.DeelDose = NotEmpty(objMed.DeelDose, objDict("multiple"))
    objMed.DoseEenheid = NotEmpty(objMed.DoseEenheid, objDict("multipleUnit"))
    objMed.Route = NotEmpty(objMed.Route, objDict("route"))
    objMed.Indicatie = NotEmpty(objMed.Indicatie, objDict("indication"))
    If objMed.Freq = "" Then objMed.SetFreqList objDict("frequency")
    objMed.PerDose = NotEmpty(objMed.PerDose, objDict("perDose"))
    objMed.PerKg = NotEmpty(objMed.PerKg, objDict("perKg"))
    objMed.PerM2 = NotEmpty(objMed.PerM2, objDict("perM2"))
    objMed.NormDose = NotEmpty(objMed.NormDose, objDict("normDose"))
    objMed.MinDose = NotEmpty(objMed.MinDose, objDict("minDose"))
    objMed.MaxDose = NotEmpty(objMed.MaxDose, objDict("maxDose"))
    objMed.AbsDose = NotEmpty(objMed.AbsDose, objDict("absMaxTotal"))
    objMed.MaxKeer = NotEmpty(objMed.MaxKeer, objDict("absMaxPerDose"))

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

Attribute VB_Name = "ModWeb"
Option Explicit

Private Sub Test_GetJson()

    Dim objClient As New WebClient
    Dim objResponse As WebResponse
    Dim objMed As New ClassMedicatieDisc
    
    objClient.BaseUrl = "http://localhost:8080"
    
    Set objResponse = objClient.GetJson("/")
    
    ProcessJson objResponse, objMed
    ModMessage.ShowMsgBoxInfo objMed.Etiket

End Sub

'{BirthYear = 2017;
' BirthMonth = 3;
' BirthDay = 2;
' WeightKg = 10.0;
' BirthWeightGram = 0.0;
' LengthCm = 70.0;
' Gender = "";
' GestAgeWeeks = 0;
' GestAgeDays = 0;
' GPK = "100331";
' ATC = "A04AA01 ";
' TherapyGroup = "ANTI-EMETICA";
' TherapySubGroup = "ANTI-EMETICA";
' Generic = "ONDANSETRON";
' TradeProduct = "";
' Shape = "STROOP";
' Label = "ONDANSETRON 0,8MG/ML STROOP";
' Concentration = 0.0;
' ConcentrationUnit = "";
' Multiple = 0.0;
' MultipleUnit = "mg";
' Route = "or";
' Indication = "";
' Frequency = "antenoctum||1 x / dag||2 x / dag||3 x / dag";
' PerDose = false;
' PerKg = false;
' PerM2 = true;
' NormDose = 0.0;
' MinDose = 0.0;
' MaxDose = 24.0;
' AbsMaxTotal = 3.0;
' AbsMaxPerDose = 1.0;
' Rules =
'  "287597, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 5. Complicatie(s) medische behandeling, Leeftijd: tot 36 maanden, Freq: 1 per dag, Norm: tot 1.25 ml||287598, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 5. Complicatie(s) medische behandeling, Leeftijd: tot 36 maanden, Freq: 2 per dag, Norm: tot 1.25 ml||287599, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 5. Complicatie(s) medische behandeling, Leeftijd: tot 36 maanden, Freq: 3 per dag, Norm: tot 1.25 ml||287612, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 8. Infectieuze diarree, dysenterie, Leeftijd: 1 - 216 maanden, Gewicht: tot 80 kg, Freq: 1 per dag, Norm/Kg: tot 0.125 ml||287613, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 8. Infectieuze diarree, dysenterie, Leeftijd: 1 - 216 maanden, Gewicht: tot 80 kg, Freq: 2 per dag,
'orm/Kg: tot 0.125 ml||287614, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 8. Infectieuze diarree, dysenterie, Leeftijd: 1 - 216 maanden, Gewicht: tot 80 kg, Freq: 3 per dag, Norm/Kg: tot 0.125 ml||287618, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 4. Algemeen, Leeftijd: 1 - 216 maanden, BSA: tot 1 m2, Freq: 1 per dag, Norm/m2: tot 10 ml||287619, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 4. Algemeen, Leeftijd: 1 - 216 maanden, BSA: tot 1 m2, Freq: 2 per dag, Norm/m2: tot 10 ml||287620, ONDANSETRON STROOP 0,8MG/ML, Groep: intensieve, Type: Standaard, Route: ORAAL, Indicatie: 4. Algemeen, Leeftijd: 1 - 216 maanden, BSA: tot 1 m2, Freq: 3 per dag, Norm/m2: tot 10 ml";}
'


Private Sub ProcessJson(objResponse As WebResponse, objMed As ClassMedicatieDisc)

    Dim objDict As Dictionary
    Dim strJson As String
    
    strJson = objResponse.Content
    strJson = Replace(strJson, "=", ":")
    strJson = Replace(strJson, ";", ",")
    strJson = Replace(strJson, ",}", "}")
    
    JsonConverter.JsonOptions.AllowUnquotedKeys = True
    Set objDict = JsonConverter.ParseJson(strJson)
    
    objMed.GPK = objDict("GPK")
    objMed.ATC = objDict("ATC")
    objMed.Generiek = objDict("Generic")
    objMed.Etiket = objDict("Label")
    objMed.AbsDose = objDict("MaxDose")

End Sub

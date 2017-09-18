Attribute VB_Name = "ModString"
Option Explicit

Public Function FirstPositionInstr(ByVal strString1 As String, ByVal strString2 As String) As Integer

    FirstPositionInstr = InStr(1, strString1, strString2, vbTextCompare)

End Function

Public Function CountFirstInStr(ByVal strString As String, ByVal strFirst As String) As Integer

    Dim intN As Integer
    
    intN = 1
    
    Do While intN <= Len(strString) And Mid(strString, intN, 1) = strFirst
        intN = intN + 1
    Loop
    
    CountFirstInStr = intN - 1

End Function

Private Sub TestCountFirstInStr()

    MsgBox CountFirstInStr("000000012345", "0")

End Sub

' Checks whether strString contains strValue case insensitive, ignores spaces
Public Function ContainsCaseInsensitive(ByVal strString As String, ByVal strValue As String) As Boolean
    strString = Strings.Trim(strString)
    strValue = Strings.Trim(strValue)
    
    If Strings.InStr(1, Strings.LCase(strString), Strings.LCase(strValue)) > 0 Then
        ContainsCaseInsensitive = True
    Else
        ContainsCaseInsensitive = False
    End If
    
End Function

' Checks whether strString contains strValue case sensitive, ignores spaces
Public Function ContainsCaseSensitive(ByVal strString As String, ByVal strValue As String) As Boolean
    strString = Strings.Trim(strString)
    strValue = Strings.Trim(strValue)

    If Strings.InStr(1, strString, strValue) > 0 Then
        ContainsCaseSensitive = True
    Else
        ContainsCaseSensitive = False
    End If
    
End Function

Public Function StartsWith(ByVal strString As String, ByVal strValue As String) As Boolean

    StartsWith = Strings.InStr(1, strString, strValue) = 1

End Function

Private Sub Test()
    Dim strString As String
    Dim strStart As String
    
    strString = "__4_GebDatum"
    strStart = "_"

    MsgBox StartsWith(strString, strStart)

End Sub

Public Function DateToString(ByVal dtmDate As Date) As String

    DateToString = Strings.Format(dtmDate, "dd-mmm-yyyy")

End Function

Private Sub TestDateToString()

    MsgBox DateToString(DateTime.Date)

End Sub

Public Function StringToDate(ByVal strValue As String) As Date

    Dim dtmDate As Date
    
    On Error GoTo StringToDateError
        
    dtmDate = CDate(strValue)
    StringToDate = dtmDate
    
    Exit Function
    
StringToDateError:

    ModLog.LogError "Cannot convert " & strValue & " to a date time"

End Function

Private Sub TestStringToDate()

    MsgBox DateToString(StringToDate("01-02-2017"))

End Sub

Public Function SplitDouble(ByVal dblNum As Double) As String()

    Dim strNum As String
    Dim strDel As String
    Dim arrNum() As String
    
    strNum = CStr(CDec(dblNum))
    strDel = IIf(ContainsCaseInsensitive(strNum, "."), ".", ",")
        
    arrNum = Split(strNum, strDel)

    If UBound(arrNum) = 0 Then ModArray.AddItemToStringArray arrNum, vbNullString
    ModArray.AddItemToStringArray arrNum, strDel
    
    SplitDouble = arrNum

End Function

Public Function DoubleToFractionString(ByVal dblNum As Double) As String

    Dim strNum As String
    Dim intDec As Integer
    Dim strDel As String
    
    strNum = CStr(dblNum)
    strDel = SplitDouble(dblNum)(2)
    intDec = Len(SplitDouble(dblNum)(1))
        
    DoubleToFractionString = Strings.Replace(strNum, strDel, vbNullString) & "/" & Application.WorksheetFunction.Power(10, intDec)

End Function

Private Sub TestDoubleToFractionString()

    MsgBox DoubleToFractionString(12)
    
End Sub

Private Function GetPower(ByVal intN As Integer, ByVal dblNum As Double) As Integer

    Dim intP As Integer
    Dim strN As String
    Dim strD As String
    
    strN = SplitDouble(dblNum)(0)
    strD = SplitDouble(dblNum)(1)
    
    intP = intN - IIf(strN = "0", 0, Len(strN))
    intP = IIf(intP < 0, 0, intP)
    intP = IIf(strD = vbNullString Or dblNum >= 1, intP, CountFirstInStr(strD, "0") + intP)
    
    GetPower = intP

End Function

Private Sub TestGetPrecision()

    MsgBox GetPower(2, 1.002343)

End Sub

Public Function FixPrecision(ByVal dblNum As Double, ByVal intN As Integer) As Double

    Dim dblFix As Double
    Dim intP As Integer
    
    intP = GetPower(intN, dblNum)
    
    With Application.WorksheetFunction
        dblFix = .Round(dblNum * .Power(10, intP), 0)
        dblFix = dblFix / .Power(10, intP)
    End With
    
    FixPrecision = dblFix

End Function

Private Sub TestFixPrecision()

    MsgBox FixPrecision(0.0347, 1)

End Sub

Private Function IsEmptyValue(objCell As Range) As Boolean

    Dim blnEmpty As Boolean
    
    blnEmpty = IsEmpty(objCell.Value2)
    If Not blnEmpty Then blnEmpty = Trim(objCell.Value2) = vbNullString
    If Not blnEmpty Then blnEmpty = Trim(objCell.Value2) = "0"
    
    IsEmptyValue = blnEmpty

End Function

Private Sub IsEmptyValue_Test()

    MsgBox IsEmptyValue(shtPedBerTPN.Range("B36"))

End Sub

Public Function ConcatenateRange(objRange As Range, ByVal strDel As String) As String

    Dim strString As String
    Dim objCell As Range
    
    For Each objCell In objRange.Cells
        If Not IsEmptyValue(objCell) Then
            strString = IIf(strString = vbNullString, objCell.Value2, strString & strDel & objCell.Value2)
        End If
    Next
    
    ConcatenateRange = strString

End Function

Private Sub ConcatenateRange_Test()

    MsgBox ConcatenateRange(shtPedBerTPN.Range("B31:B36"), "##")

End Sub
Public Function StringToDouble(ByVal strDbl As String) As Double

    StringToDouble = Val(Replace(strDbl, ",", "."))

End Function

Private Sub Test_StringToDouble()

    MsgBox StringToDouble("1.5")
    MsgBox StringToDouble("1,5")

End Sub

Private Function RemoveTrailing(ByVal strString As String, ByVal strDel As String) As String
    
    Dim intLeft As Integer

    intLeft = Len(strString) - Len(strDel)
    If intLeft > 0 Then strString = Left(strString, intLeft)
    
    RemoveTrailing = strString

End Function

Public Function ConcatenateKeyValue(objRange As Range, ByVal strRow As String, ByVal strCol As String) As String

    Dim strString As String
    Dim objCell As Range
    Dim strKey As String
    Dim strVal As String
    
    For Each objCell In objRange.Cells
        If strKey = vbNullString Then
            strKey = Trim(objCell.Value2)
        Else
            strVal = Trim(objCell.Value2)
            strVal = IIf(strVal = "0", vbNullString, strVal)
            If Not strVal = vbNullString Then
                strString = strString & strKey & strCol & strVal & strRow
            End If
            strKey = vbNullString
            strVal = vbNullString
        End If
    Next
    
    ConcatenateKeyValue = RemoveTrailing(strString, strRow)

End Function

Public Function IntNToStrN(ByVal intN As Integer) As String

    IntNToStrN = IIf(intN < 10, "0" & intN, intN)

End Function

Public Function StringIsZeroOrEmpty(ByVal strString As String) As Boolean

    StringIsZeroOrEmpty = Trim(strString) = "0" Or strString = vbNullString
    
End Function


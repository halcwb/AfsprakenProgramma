Attribute VB_Name = "ModDate"
Option Explicit

Public Function EmptyDate() As Date

    EmptyDate = DateSerial(1904, 1, 1)

End Function

Public Function IsEmptyDate(ByVal dtmDate As Date) As Boolean

    IsEmptyDate = dtmDate <= EmptyDate()

End Function

Private Sub TestEmptyDate()

    MsgBox Strings.Format(EmptyDate(), "dd-mmm-yyyy")

End Sub

Public Function DateYear(ByVal dtmDate As Date) As Integer

    DateYear = DatePart("yyyy", dtmDate)

End Function


Public Function DateMonth(ByVal dtmDate As Date) As Integer

    DateMonth = DatePart("m", dtmDate)

End Function

Public Function DateDay(ByVal dtmDate As Date) As Integer

    DateDay = DatePart("d", dtmDate)

End Function

Private Sub Test_DateYear()

    ModMessage.ShowMsgBoxInfo DateYear(DateTime.Now)

End Sub

Public Function FormatDateTimeSeconds(dtmDateTime As Date) As String
    
    Dim strFormat As String

    If IsEmptyDate(dtmDateTime) Then
    strFormat = vbNullString
    Else
    strFormat = Format(dtmDateTime, "yyyy-mm-dd hh:mm:ss")
    End If
    
    FormatDateTimeSeconds = strFormat

End Function

Private Sub Test_FormatDateTimeSeconds()

    ModMessage.ShowMsgBoxInfo ModRange.GetRangeValue("Var_Glob_Versie", vbNullString)

End Sub

Public Function FormatDateYearMonthDay(dtmDateTime As Date) As String

    FormatDateYearMonthDay = Format(dtmDateTime, "yyyy-mm-dd")

End Function

Public Function FormatDateDayMonthYears(dtmDateTime As Date) As String

    FormatDateDayMonthYears = Format(dtmDateTime, "dd-mm-yyyy")

End Function

Public Function FormatDateHoursMinutes(dtmDate As Date) As String

    FormatDateHoursMinutes = Format(dtmDate, "hh:mm:ss")

End Function

Private Sub Test_FormatHoursMinutes()

    ModMessage.ShowMsgBoxInfo FormatDateHoursMinutes(Now())

End Sub


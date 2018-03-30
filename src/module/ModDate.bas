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

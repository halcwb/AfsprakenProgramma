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


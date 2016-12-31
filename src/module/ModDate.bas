Attribute VB_Name = "ModDate"
Option Explicit

Public Function EmptyDate() As Date

    EmptyDate = DateSerial(1900, 1, 1)

End Function

Public Function IsEmptyDate(dtmDate As Date) As Boolean

    IsEmptyDate = dtmDate = EmptyDate()

End Function

Private Sub TestEmptyDate()

    MsgBox IsEmptyDate(EmptyDate())

End Sub


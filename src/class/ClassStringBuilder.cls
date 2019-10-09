VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim arrBuffer() As String
Dim lngCurrentIndex As Long
Dim lngMaxIndex As Long

Private Sub Class_Initialize()

    lngCurrentIndex = 0
    lngMaxIndex = 16
    ReDim arrBuffer(1 To lngMaxIndex)

End Sub

'Appends the given Text to this StringBuilder
Public Sub Append(Text As String)

    lngCurrentIndex = lngCurrentIndex + 1

    If lngCurrentIndex > lngMaxIndex Then
        lngMaxIndex = 2 * lngMaxIndex
        ReDim Preserve arrBuffer(1 To lngMaxIndex)
    End If
    arrBuffer(lngCurrentIndex) = Text

End Sub

'Returns the text in this StringBuilder
'Optional Parameter: Separator (default vbNullString) used in joining components
Public Function ToString(Optional Separator As String = vbNullString) As String

    If lngCurrentIndex > 0 Then
        ReDim Preserve arrBuffer(1 To lngCurrentIndex)
        lngMaxIndex = lngCurrentIndex
        ToString = Join(arrBuffer, Separator)
    End If

End Function

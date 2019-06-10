Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Spinner 181")).Select
    With Selection
        .Value = 0
        .Min = 0
        .Max = 30000
        .SmallChange = 1
        .LinkedCell = "GlobBerMedDisc!$G$11"
        .Display3DShading = True
    End With
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    With Selection
        .Value = 0
        .Min = 0
        .Max = 30000
        .SmallChange = 1
        .LinkedCell = "GlobBerMedDisc!$G$31"
        .Display3DShading = True
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("U37").Select
    ActiveSheet.Shapes.Range(Array("Spinner 632")).Select
    With Selection
        .Value = 0
        .Min = 0
        .Max = 30000
        .SmallChange = 1
        .LinkedCell = "GlobBerMedDisc!$G$31"
        .Display3DShading = True
    End With
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Spinner 265")).Select
    With Selection
        .Value = 0
        .Min = 0
        .Max = 30000
        .SmallChange = 1
        .LinkedCell = "GlobBerMedDisc!$G$12"
        .Display3DShading = True
    End With
End Sub

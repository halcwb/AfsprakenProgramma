Attribute VB_Name = "Module1"
Option Explicit

Declare PtrSafe Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (MyDest As Any, MySource As Any, ByVal MySize As Long)

Function CmdToSTr(lngCmd As Long) As String

    Dim arrBuffer() As Byte
    Dim dblLen As Long
    
    If lngCmd Then
        dblLen = lstrlenW(lngCmd) * 2
        
        If dblLen Then
            ReDim arrBuffer(0 To (dblLen - 1)) As Byte
            CopyMemory arrBuffer(0), ByVal lngCmd, dblLen
            CmdToSTr = arrBuffer
        End If
    End If
    
End Function

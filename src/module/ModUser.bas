Attribute VB_Name = "ModUser"
Option Explicit

Private Const constUserFirstName As String = "_User_FirstName"
Private Const constUserLastName As String = "_User_LastName"
Private Const constUserLogin As String = "_User_Login"
Private Const constUserType As String = "_User_Type"

Public Function User_GetCurrent() As ClassUser

    Dim objUser As ClassUser
    
    Set objUser = New ClassUser
    
    With objUser
        .Login = ModRange.GetRangeValue(constUserLogin, vbNullString)
        .LastName = ModRange.GetRangeValue(constUserLastName, vbNullString)
        .FirstName = ModRange.GetRangeValue(constUserFirstName, vbNullString)
        .Role = ModRange.GetRangeValue(constUserType, vbNullString)
    End With
    
    Set User_GetCurrent = objUser

End Function

Public Sub User_SetUser(objUser As ClassUser)

    If objUser.IsValid Then
        With objUser
            ModRange.SetRangeValue constUserLogin, .Login
            ModRange.SetRangeValue constUserLastName, .LastName
            ModRange.SetRangeValue constUserFirstName, .FirstName
            ModRange.SetRangeValue constUserType, .Role
        End With
    Else
    End If

End Sub

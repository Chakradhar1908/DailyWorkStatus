Module modSystem
    Public Function DoBeep(Optional ByVal Count as integer = 1, Optional ByVal BeepStyle as integer = 0) As Boolean
        Dim I as integer
        DoBeep = True
        If IsDevelopment() Then Exit Function
        If Count < 1 Then Exit Function
        If Count > 10 Then Count = 10

        For I = 1 To Count
            Select Case BeepStyle
                Case 0 : Beep()      ' Generic System Sound
                Case Else              ' NONE!!
            End Select
        Next
    End Function

End Module

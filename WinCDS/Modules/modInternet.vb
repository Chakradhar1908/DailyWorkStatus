Module modInternet
    Public Function URLEncode(ByVal sRawURL As String, Optional ByVal AllowAmpersand As Boolean = True) As String
        On Error GoTo Catch1
        Dim iLoop As Integer, sRtn As String, sTmp As String
        Dim sValidChars As String
        sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:/.?=_-$(){}~"
        If AllowAmpersand Then sValidChars = sValidChars & "&"

        If Len(sRawURL) > 0 Then
            For iLoop = 1 To Len(sRawURL) ' Loop through each char
                sTmp = Mid(sRawURL, iLoop, 1)

                If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                    ' If not ValidChar, convert to HEX and prefix with %
                    sTmp = Hex(Asc(sTmp))

                    If sTmp = "20" Then
                        sTmp = "+"
                    ElseIf Len(sTmp) = 1 Then
                        sTmp = "%0" & sTmp
                    Else
                        sTmp = "%" & sTmp
                    End If
                End If
                sRtn = sRtn & sTmp
            Next
            URLEncode = sRtn
        End If
Finally1:
        Exit Function
Catch1:
        URLEncode = ""
        Resume Finally1
    End Function

    Public Function QueryStringQueryL(ByVal QueryString As String, ByVal key As String) As Long
        QueryStringQueryL = Val(QueryStringQuery(QueryString, key))
    End Function

End Module

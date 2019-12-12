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

    Public Function QueryStringQueryL(ByVal QueryString As String, ByVal key As String) As Integer
        QueryStringQueryL = Val(QueryStringQuery(QueryString, key))
    End Function

    Public Function QueryStringQuery(ByVal QueryString As String, ByVal key As String) As String
        QueryStringQuery = UrlDecode(QueryStringParse(QueryString).Item(key))
    End Function

    Public Function QueryStringParse(ByVal QueryString As String) As clsHashTable
        Dim C As clsHashTable
        Dim S() As String, T, TT As String
        Dim F As Integer
        Dim K As String, V As String
        If Left(QueryString, 1) = "?" Then QueryString = Mid(QueryString, 2)
        S = Split(QueryString, "&")

        C = New clsHashTable

        For Each T In S
            TT = UrlDecode(T)
            F = InStr(TT, "=")
            If F > 0 Then
                K = Left(TT, F - 1)
                V = Mid(TT, F + 1)
                C.Add(K, V)
            End If
        Next

        QueryStringParse = C
    End Function

    Public Function UrlDecode(ByVal sText As String) As String
        Dim sTemp As String
        Dim sAns As String
        Dim sChar As String
        Dim lCtr As Integer
        Dim C1 As String, C2 As String, C3 As String

        For lCtr = 1 To Len(sText)

            sChar = Mid(sText, lCtr, 1)

            If sChar = "+" Then
                Mid(sText, lCtr, 1) = " "
            ElseIf sChar = "%" Then
                C1 = Mid(sText, lCtr + 1, 2)
                C2 = "&H" & C1
                C3 = Chr(Val(C2))
                sText = Replace(sText, "%" & C1, C3)
            End If
        Next
        UrlDecode = sText
    End Function

    Public Function ValidEmailAddress(ByVal Addr As String) As Boolean
        Dim TLD As String, X As String, N As Integer, O As Long
        If Addr = "" Then Exit Function
        N = InStr(Addr, "@")
        If N <= 1 Then Exit Function
        O = InStr(N + 1, Addr, ".")
        If O <= 0 Then Exit Function
        X = Addr
        TLD = ""
        Do While Right(X, 1) <> "."
            TLD = Right(X, 1) & TLD
            X = Mid(X, 1, Len(X) - 1)
        Loop
        If TLD = "" Then Exit Function
        ' for a complete list of TLDs, you can visit:
        ' http://www.iana.org/cctld/cctld-whois.htm
        ' but this isn't really needed
        '  If Not IsIn(TLD, TLD_List) Then Exit Function
        ValidEmailAddress = True
    End Function

End Module

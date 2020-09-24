Module modFuzzyStringMatching
    Public Function FuzzyStringMatch(ByVal StringA As String, ByVal StringB As String, Optional ByVal Threshhold As Double = 0.8) As Boolean
        FuzzyStringMatch = HotFuzz(StringA, StringB) > Threshhold
    End Function

    Private Function HotFuzz(ByVal S1 As String, ByVal S2 As String, Optional ByVal N As Boolean = True, Optional ByVal X As String = "", Optional ByVal W As Single = 2) As Single
        'Using Like operator for filtering, with added code to allow special characters in the input strings, including hyphen and the right bracket - passed in the 'x' parameter.
        'Use x & Chr(34) if you need to allow double quotes (") in the input strings
        'Allowing numbers in the input strings is optional (the 'n' parameter)
        'The 'w' parameter is the weight of "order" over "frequency" scores in the final score. Feel free to experiment, to get best matching results with your data.
        Dim I As Integer, D1 As Integer, D2 As Integer, Y As String, B As Boolean
        Dim C As String, A1 As String, A2 As String, K As Integer, P As Integer, F As Single, O As Single
        '
        '        ******* INPUT STRINGS CLEANSING *******
        '
        HotFuzz = 0
        B = False
        If N Then 'allow numbers in the input strings?
            Y = "[A-Z0-9"
        Else
            Y = "[A-Z"
        End If
        If Len(X) > 0 Then 'we want to allow some special characters in the input strings, i.e. space, punctuation etc
            If InStr(1, X, "-", 0) Then
                Y = Replace(X, "-", "") & "-" 'hyphen must be placed first or last inside a [..] group in a Like comparison
            End If
            If InStr(1, X, "]", 0) Then
                Y = Replace(X, "]", "") 'right bracket can't be part of a [..] group in a Like comparison - dedicated logic must be developed to treat this case
                B = True 'if we want to allow the right bracket in the input strings
            End If
        End If
        Y = Y & "]" 'closing the group
        S1 = UCase(S1) 'input strings are converted to uppercase
        D1 = Len(S1)
        A1 = ""
        For I = 1 To D1
            C = Mid(S1, I, 1)
            If C Like Y Then  'filter the allowable characters
                A1 = A1 & C 'a1 is what remains from s1 after filtering
            ElseIf B Then
                If C = "]" Then 'special treatment for the right bracket
                    A1 = A1 & C
                End If
            End If
        Next
        D1 = Len(A1)
        If D1 = 0 Then Exit Function
        S2 = UCase(S2)
        D2 = Len(S2)
        A2 = ""
        For I = 1 To D2
            C = Mid(S2, I, 1)
            If C Like Y Then
                A2 = A2 & C
            End If
        Next
        D2 = Len(A2)
        If D2 = 0 Then Exit Function
        K = D1
        If D2 < D1 Then 'to prevent doubling the code below s1 must be made the shortest string,
            'so we swap the variables
            K = D2
            D2 = D1
            D1 = K
            S1 = A2
            S2 = A1
            A1 = S1
            A2 = S2
        Else
            S1 = A1
            S2 = A2
        End If
        If K = 1 Then 'degenerate case, where the shortest string is just one character
            If InStr(1, S2, S1, 0) Then
                HotFuzz = 1 / D2
            Else
                HotFuzz = 0
            End If
        Else        '******* MAIN LOGIC HERE *******
            I = 1
            F = 0
            O = 0
            Do 'count the identical characters in s1 and s2 ("frequency analysis")
                P = InStr(1, S2, Mid(S1, I, 1), 0)
                'search the character at position i from s1 in s2
                If P > 0 Then   'found a matching character, at position p in s2
                    F = F + 1   'increment the frequency counter
                    Mid(S2, P, 1) = "~"
                    'replace the found character with one outside the allowable list
                    '(I used tilde here), to prevent re-finding
                    Do      'check the order of characters
                        If I >= K Then Exit Do 'no more characters to search
                        If Mid(S2, P + 1, 1) = Mid(S1, I + 1, 1) Then
                            'test if the next character is the same in the two strings
                            F = F + 1 'increment the frequency counter
                            O = O + 1 'increment the order counter
                            I = I + 1
                            P = P + 1
                        Else
                            Exit Do
                        End If
                    Loop
                End If
                If I >= K Then Exit Do
                I = I + 1
            Loop
            If O > 0 Then O = O + 1 'if we got at least one match, adjust the order counter because two characters are required to define "order"
            HotFuzz = (W * O + F) / (W + 1) / D2
        End If
    End Function
End Module

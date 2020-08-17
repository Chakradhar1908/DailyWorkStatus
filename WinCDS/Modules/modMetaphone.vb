Module modMetaphone
    Public Function Metaphone(ByVal strInput As Object) As String
        Dim B As String, C As String, D As String, E As String
        Dim Inp, OutP As String
        Dim Vowels As String, FrontV As String, Varson As String, Dbl As String
        Dim ExcPPair As String, NxtLtr As String
        Dim T As Integer, II As Integer, JJ As Integer, Lng As Integer, Lastchr As Integer
        Dim CurLtr, PrevLtr, NextLtr, NextLtr2, NextLtr3 As String
        Dim VowelAfter, VowelBefore, FrontVAfter, Silent, Hard As Integer
        Dim AlphaChr As String

        On Error Resume Next
        If IsNothing(strInput) Then strInput = ""
        Inp = CStr(UCase(strInput))

        Vowels = "AEIOU"
        FrontV = "EIY"
        Varson = "CSPTG"

        Dbl = "." 'Lets us allow certain letters to be doubled   excppair = "AGKPW"
        NxtLtr = "ENNNR"
        AlphaChr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        '--Remove non-alpha characters
        OutP = ""
        For T = 1 To Len(Inp)
            If InStr(AlphaChr, Mid(Inp, T, 1)) > 0 Then OutP = OutP + Mid(Inp, T, 1)
        Next

        Inp = OutP : OutP = ""

        If Len(Inp) = 0 Then Metaphone = "" : Exit Function

        '--Check rules at beginning of word
        If Len(Inp) > 1 Then
            B = Mid(Inp, 1, 1)
            C = Mid(Inp, 2, 1)
            II = InStr(ExcPPair, B)
            JJ = InStr(NxtLtr, C)
            If II = JJ And II > 0 Then
                Inp = Mid(Inp, 2, Len(Inp) - 1)
            End If
        End If

        If Mid(Inp, 1, 1) = "X" Then Mid(Inp, 1, 1) = "S"

        If Mid(Inp, 1, 2) = "WH" Then Inp = "W" + Mid(Inp, 3)

        If Right(Inp, 1) = "S" Then Inp = Left(Inp, Len(Inp) - 1)

        II = 0
        Do
            II = II + 1
            '--Main Loop!
            Silent = False
            Hard = False
            CurLtr = Mid(Inp, II, 1)
            VowelBefore = False
            PrevLtr = " "
            If II > 1 Then

                PrevLtr = Mid(Inp, II - 1, 1)
                If InStrC(PrevLtr, Vowels) > 0 Then VowelBefore = True

            End If

            If ((II = 1) And (InStrC(CurLtr, Vowels) > 0)) Then

                OutP = OutP + CurLtr
                GoTo ContinueMainLoop
            End If

            VowelAfter = False
            FrontVAfter = False
            NextLtr = " "
            If II < Len(Inp) Then

                NextLtr = Mid(Inp, II + 1, 1)
                If InStrC(NextLtr, Vowels) > 0 Then VowelAfter = True
                If InStrC(NextLtr, FrontV) > 0 Then FrontVAfter = True

            End If

            '--Skip double letters EXCEPT ones in variable double     If InStrC(curltr, dbl) = 0 Then
            If CurLtr = NextLtr Then
                GoTo ContinueMainLoop
            End If

            NextLtr2 = " "
            If Len(Inp) - II > 1 Then
                NextLtr2 = Mid(Inp, II + 2, 1)
            End If

            NextLtr3 = " "
            If (Len(Inp) - II) > 2 Then
                NextLtr3 = Mid(Inp, II + 3, 1)
            End If

            Select Case CurLtr
                Case "B"

                    Silent = False
                    If (II = Len(Inp)) And (PrevLtr = "M") Then Silent = True
                    If Not (Silent) Then OutP = OutP + CurLtr
                Case "C"
                    If Not ((II > 2) And (PrevLtr = "S") And FrontVAfter) Then
                        If ((II > 1) And (NextLtr = "I") And (NextLtr2 = "A")) Then
                            OutP = OutP + "X"
                        Else
                            If FrontVAfter Then
                                OutP = OutP + "S"
                            Else
                                If ((II > 2) And (PrevLtr = "S") And (NextLtr = "H")) Then
                                    OutP = OutP + "K"
                                Else
                                    If NextLtr = "H" Then
                                        If ((II = 1) And (InStrC(NextLtr2, Vowels) = 0)) Then
                                            OutP = OutP + "K"
                                        Else
                                            OutP = OutP + "X"
                                        End If
                                    Else
                                        If PrevLtr = "C" Then
                                            OutP = OutP + "C"
                                        Else
                                            OutP = OutP + "K"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case "D"
                    If ((NextLtr = "G") And (InStrC(NextLtr2, FrontV) > 0)) Then
                        OutP = OutP + "J"
                    Else
                        OutP = OutP + "T"
                    End If

                Case "G"
                    Silent = False
                    If ((II < Len(Inp)) And (NextLtr = "H") And (InStrC(NextLtr2, Vowels) = 0)) Then
                        Silent = True
                    End If
                    If ((II = Len(Inp) - 4) And (NextLtr = "N") And (NextLtr2 = "E") And (NextLtr3 = "D")) Then
                        Silent = True
                    ElseIf ((II = Len(Inp) - 2) And (NextLtr = "N")) Then
                        Silent = True
                    End If
                    If (PrevLtr = "D") And FrontVAfter Then Silent = True
                    If PrevLtr = "G" Then
                        Hard = True
                    End If

                    If Not (Silent) Then
                        If FrontVAfter And (Not (Hard)) Then
                            OutP = OutP + "J"
                        Else
                            OutP = OutP + "K"
                        End If
                    End If

                Case "H"
                    Silent = False
                    If InStrC(PrevLtr, Varson) > 0 Then Silent = True
                    If VowelBefore And (Not (VowelAfter)) Then Silent = True
                    If Not Silent Then OutP = OutP + CurLtr


                Case "F", "J", "L", "M", "N", "R" : OutP = OutP + CurLtr

                Case "K" : If PrevLtr <> "C" Then OutP = OutP + CurLtr

                Case "P" : If NextLtr = "H" Then OutP = OutP + "F" Else OutP = OutP + "P"

                Case "Q" : OutP = OutP + "K"

                Case "S"

                    If ((II > 2) And (NextLtr = "I") And ((NextLtr2 = "O") Or (NextLtr2 = "A"))) Then

                        OutP = OutP + "X"
                    End If
                    If (NextLtr = "H") Then
                        OutP = OutP + "X"
                    Else
                        OutP = OutP + "S"
                    End If

                Case "T"
                    If ((II > 0) And (NextLtr = "I") And ((NextLtr2 = "O") Or (NextLtr2 = "A"))) Then
                        OutP = OutP + "X"
                    End If
                    If NextLtr = "H" Then
                        If ((II > 1) Or (InStrC(NextLtr2, Vowels) > 0)) Then
                            OutP = OutP + "0"
                        Else
                            OutP = OutP + "T"
                        End If
                    ElseIf Not ((II < Len(Inp) - 3) And (NextLtr = "C") And (NextLtr2 = "H")) Then
                        OutP = OutP + "T"
                    End If


                Case "V" : OutP = OutP + "F"

                Case "W", "Y" : If (II < Len(Inp) - 1) And VowelAfter Then OutP = OutP + CurLtr

                Case "X" : OutP = OutP + "KS"

                Case "Z" : OutP = OutP + "S"

            End Select
ContinueMainLoop:
        Loop Until (II > Len(Inp))

        Metaphone = OutP
    End Function

    Private Function InStrC(ByVal SearchIn As String, ByVal SoughtCharacters As String) As Integer
        '--- Returns the position of the first character in SearchIn that is contained
        '--- in the string SoughtCharacters. Returns 0 if none found.
        Dim I As Integer

        On Error Resume Next
        SoughtCharacters = UCase(SoughtCharacters)
        SearchIn = UCase(SearchIn)
        For I = 1 To Len(SearchIn)
            If InStr(SoughtCharacters, Mid(SearchIn, I, 1)) > 0 Then
                InStrC = I
                Exit Function
            End If
        Next
        InStrC = 0
    End Function
End Module

Module modSoundEX
    Public Function SoundEx(ByVal Word As String) As String
        Dim Result As String
        Dim I As Integer, Acode As Integer
        Dim Dcode As Integer, oldCode As Integer

        If Word = "" Then Exit Function

        ' soundex is case-insensitive
        Word = UCase(Word)
        If Asc(Left(Word, 1)) - 64 < 1 Or Asc(Left(Word, 1)) - 64 > 26 Then Exit Function ' discard non-alphabetic chars (first char)

        ' the first letter is copied in the result
        SoundEx = Left(Word, 1)

        oldCode = Asc(Mid("01230120022455012623010202", Asc(Word) - 64))

        For I = 2 To Len(Word)
            Acode = Asc(Mid(Word, I, 1)) - 64
            ' discard non-alphabetic chars
            If Acode >= 1 And Acode <= 26 Then
                ' convert to a digit
                Dcode = Asc(Mid("01230120022455012623010202", Acode, 1))
                ' don't insert repeated digits
                If Dcode <> 48 And Dcode <> oldCode Then
                    SoundEx = SoundEx & Chr(Dcode)
                    If Len(SoundEx) = 4 Then Exit For
                End If
                oldCode = Dcode
            End If
        Next
    End Function
End Module

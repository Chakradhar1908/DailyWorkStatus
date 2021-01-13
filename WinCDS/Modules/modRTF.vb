Module modRTF
    Public Function RtfToHtml(ByVal RTF As RichTextBox) As String
        Const pL As String = "<p>"
        Const PC As String = "<p align='center'>"

        Dim FontFamily As String, FontSize As Double
        Dim B As Boolean, I As Boolean, U As Boolean, S As Boolean
        Dim P As Integer, Res As String, C As String

        FontFamily = ""
        FontSize = 0
        B = False
        I = False
        U = False
        S = False
        Res = ""

        'RTF.SelStart = 1
        RTF.SelectionStart = 1
        'RTF.SelLength = 1
        RTF.SelectionLength = 1

        If RTF.SelectionAlignment = HorizontalAlignment.Center Then
            Res = Res & PC
        Else
            Res = Res & pL
        End If

        For P = 1 To Len(RTF.Text)
            RTF.SelectionStart = P - 1
            RTF.SelectionLength = 1

            'If RTF.SelFontName <> FontFamily Then FontFamily = RTF.SelFontName : Res = Res & "<font face='" & FontFamily & "'>"
            If RTF.SelectionFont.Name <> FontFamily Then FontFamily = RTF.SelectionFont.Name : Res = Res & "<font face='" & FontFamily & "'>"
            'If RTF.SelFontSize <> FontSize Then FontSize = RTF.SelFontSize : Res = Res & "<font size='" & TranslateFontSize(FontSize) & "'>"
            If RTF.SelectionFont.Size <> FontSize Then FontSize = RTF.SelectionFont.Size : Res = Res & "<font size='" & TranslateFontSize(FontSize) & "'>"
            'If RTF.SelBold Xor B Then B = RTF.SelBold : Res = Res & "<" & IIf(B, "", "/") & "b>"
            If RTF.SelectionFont.Bold Xor B Then B = RTF.SelectionFont.Bold : Res = Res & "<" & IIf(B, "", "/") & "b>"
            'If RTF.SelItalic Xor I Then I = RTF.SelItalic : Res = Res & "<" & IIf(I, "", "/") & "i>"
            If RTF.SelectionFont.Italic Xor I Then I = RTF.SelectionFont.Italic : Res = Res & "<" & IIf(I, "", "/") & "i>"
            'If RTF.SelUnderline Xor U Then U = RTF.SelUnderline : Res = Res & "<" & IIf(U, "", "/") & "u>"
            If RTF.SelectionFont.Underline Xor U Then U = RTF.SelectionFont.Underline : Res = Res & "<" & IIf(U, "", "/") & "u>"
            'If RTF.SelStrikeThru Xor S Then S = RTF.SelStrikeThru : Res = Res & "<" & IIf(S, "", "/") & "del>"
            If RTF.SelectionFont.Strikeout Xor S Then S = RTF.SelectionFont.Strikeout : Res = Res & "<" & IIf(S, "", "/") & "del>"

            'C = RTF.SelText
            C = RTF.SelectedText
            Select Case C
                Case "<" : C = "&lt;"
                Case ">" : C = "&gt;"
                Case "&" : C = "&amp;"
            End Select

            RTF.SelectionLength = 2
            If InStr(RTF.SelectedText, vbLf) Then
                RTF.SelectionStart = P + 1
                If RTF.SelectionAlignment = HorizontalAlignment.Center Then
                    If C = vbLf Then
                        C = PC
                    Else
                        C = C & PC
                    End If
                Else
                    If C = vbLf Then
                        C = pL
                    Else
                        C = C & pL
                    End If
                End If
            End If

            Res = Res & C
        Next
        Res = Replace(Res, pL & pL, pL)
        Res = Replace(Res, PC & PC, PC)

        Res = "<span>" & Res & "</span>"

        RtfToHtml = Res
    End Function

    Public Function TranslateFontSize(ByVal Sz As Double) As String
        TranslateFontSize = "" & Math.Round(Sz / 4, 3) & "em"
    End Function
End Module

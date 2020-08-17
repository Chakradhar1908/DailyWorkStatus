Module modXML
    Public Function PullXMLTag(ByRef Src As String, ByVal StartTag As String, ByVal EndTag As String, Optional ByVal ReplaceTags As Boolean = False, Optional ByRef RemoveOrig As Boolean = False) As String
        Dim A As Integer, B As Integer
        On Error GoTo None
        A = InStr(Src, StartTag)
        B = InStr(A + 1, Src, EndTag)
        If A = 0 Then Exit Function
        PullXMLTag = Mid(Src, A + Len(StartTag), B - A - Len(StartTag))

        If ReplaceTags Then PullXMLTag = StartTag & PullXMLTag & EndTag : GoTo Done

        Do While IsIn(Left(PullXMLTag, 1), vbCr, vbLf, " ", vbTab)
            PullXMLTag = Mid(PullXMLTag, 2)
        Loop
        Do While IsIn(Right(PullXMLTag, 1), vbCr, vbLf, " ", vbTab)
            PullXMLTag = Left(PullXMLTag, Len(PullXMLTag) - 1)
        Loop
None:
Done:
        If RemoveOrig Then Src = Replace(Src, PullXMLTag, "")

    End Function

    Public Function ProtectXML(ByVal Str As String) As String
        Str = Replace(Str, "&", "&amp;")
        Str = Replace(Str, "<", "&lt;")
        Str = Replace(Str, ">", "&gt;")
        ProtectXML = Str
    End Function

    Public Function ProtectCDATA(ByVal Str As String, Optional ByVal AddCDATATags As Boolean = False) As String
        ProtectCDATA = Replace(Str, "]]>", "]]&gt;")
        If AddCDATATags Then ProtectCDATA = "<![CDATA[" & ProtectCDATA & "]]>"
    End Function

    Public Function XMLCurrency(ByVal Curr As String) As String
        XMLCurrency = CurrencyFormat(GetPrice(Curr), , , True)
    End Function

    Public Function HTTPImageType(ByVal FileName As String) As String
        Select Case GetFileExt(FileName)
            Case "png" : HTTPImageType = "image/png"
            Case "bmp" : HTTPImageType = "image/bmp"
            Case "jpg", "jpeg" : HTTPImageType = "image/jpeg"
            Case "gif" : HTTPImageType = "image/gif"
        End Select
    End Function

End Module

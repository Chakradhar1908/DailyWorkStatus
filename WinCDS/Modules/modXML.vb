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

End Module

Module modURLHelper
    '::::modURLHelper.bas
    ':::SUMMARY
    ': URL Helper Module
    ':::DESCRIPTION
    ': This module contains different functions that are used to extract URL and encode or decode a URL.
    Structure URLExtract
        Dim Scheme As String
        Dim Host As String
        Dim Port As Integer
        Dim URI As String
        Dim Query As String
    End Structure

    ' returns as type URL from a string
    Public Function ExtractUrl(ByVal strUrl As String) As URLExtract
        '::::ExtractUrl
        ':::SUMMARY
        ': Used to Extract URL.
        ':::DESCRIPTION
        ': This function is used to returns as type URL from a String.
        ':::PARAMETERS
        ': - strUrl - Indicates the URL string.
        ':::RETURN
        ': URLExtract
        Dim intPos1 As Integer
        Dim intPos2 As Integer

        Dim retURL As URLExtract

        '1 look for a scheme it ends with ://
        intPos1 = InStr(strUrl, "://")

        If intPos1 > 0 Then
            retURL.Scheme = Mid(strUrl, 1, intPos1 - 1)
            strUrl = Mid(strUrl, intPos1 + 3)
        End If

        '2 look for a port
        intPos1 = InStr(strUrl, ":")
        intPos2 = InStr(strUrl, "/")

        If intPos1 > 0 And intPos1 < intPos2 Then
            ' a port is specified
            retURL.Host = Mid(strUrl, 1, intPos1 - 1)

            If (IsNumeric(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))) Then
                retURL.Port = CInt(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))
            End If
        ElseIf intPos2 > 0 Then
            retURL.Host = Mid(strUrl, 1, intPos2 - 1)
        Else
            retURL.Host = strUrl
            retURL.URI = "/"

            ExtractUrl = retURL
            Exit Function
        End If

        strUrl = Mid(strUrl, intPos2)

        ' find a question mark ?
        intPos1 = InStr(strUrl, "?")

        If intPos1 > 0 Then
            retURL.URI = Mid(strUrl, 1, intPos1 - 1)
            retURL.Query = Mid(strUrl, intPos1 + 1)
        Else
            retURL.URI = strUrl
        End If

        ExtractUrl = retURL
    End Function

End Module

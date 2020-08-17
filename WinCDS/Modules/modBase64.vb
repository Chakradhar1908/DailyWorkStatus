Module modBase64
    Private pbBase64Byt(0 To 63) As Byte             ' base 64 encoder byte array
    Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    Public Function EncodeBase64String(ByVal str2Encode As String) As String
        ' ******************************************************************************
        ' Synopsis:     Base 64 encode a string
        ' Parameters:   str2Encode  - The input string
        ' Return:       encoded string
        '
        ' Description:
        ' Convert a string to a byte array and pass to EncodeBase64Byte function (above)
        ' for Base64 conversion. Convert byte array back to a string and return.
        ' ******************************************************************************
        Dim tmpByte() As Byte, iPtr as integer

        EncodeBase64String = Nothing
        For iPtr = 0 To 63
            pbBase64Byt(iPtr) = Asc(Mid(BASE64CHR, iPtr + 1, 1))
        Next

        If Len(str2Encode) Then
            'tmpByte = StrConv(str2Encode, vbFromUnicode)      ' convert string to byte array
            tmpByte = System.Text.Encoding.Default.GetBytes(str2Encode)

            tmpByte = EncodeBase64Byte(tmpByte)               ' pass to the byte array encoder
            'EncodeBase64String = StrConv(tmpByte, vbUnicode)  ' convert back to string & return
            Text.Encoding.Default.GetString(tmpByte)
        End If
    End Function
    Public Function DecodeBase64String(ByVal str2Decode As String) As String

        ' ******************************************************************************
        '
        ' Synopsis:     Decode a Base 64 string
        '
        ' Parameters:   str2Decode  - The base 64 encoded input string
        '
        ' Return:       decoded string
        '
        ' Description:
        ' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
        ' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
        ' ascii character equivalent. Stop converting at the end of the input string
        ' or when the first '=' (equal sign) is encountered.
        '
        ' ******************************************************************************

        Dim lPtr as integer
        Dim iValue As Integer
        Dim iLen As Integer
        Dim iCtr As Integer
        'Dim Bits(1 To 4) As Byte
        Dim Bits(0 To 3) As Byte
        Dim strDecode As String = ""

        DecodeBase64String = ""
        ' for each 4 character group....
        For lPtr = 1 To Len(str2Decode) Step 4
            iLen = 4
            For iCtr = 0 To 3
                ' retrive the base 64 value, 4 at a time
                iValue = InStr(1, BASE64CHR, Mid(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
                Select Case iValue
                ' A~Za~z0~9+/
                    Case 1 To 64 : Bits(iCtr + 1) = iValue - 1
                ' =
                    Case 65
                        iLen = iCtr
                        Exit For
                ' not found
                    Case 0 : Exit Function
                End Select
            Next

            ' convert the 4, 6 bit values into 3, 8 bit values
            Bits(0) = Bits(0) * &H4 + (Bits(1) And &H30) \ &H10
            Bits(1) = (Bits(1) And &HF) * &H10 + (Bits(2) And &H3C) \ &H4
            Bits(2) = (Bits(2) And &H3) * &H40 + Bits(3)

            ' add the three new characters to the output string
            For iCtr = 1 To iLen - 1
                strDecode = strDecode & Chr(Bits(iCtr))
            Next

        Next

        DecodeBase64String = strDecode

    End Function
    Private Function EncodeBase64Byte(ByRef InArray() As Byte, Optional ByVal AddCRLF As Boolean = False) As Byte()
        '******************************************************************************
        ' Synopsis:     Base 64 encode a byte array
        ' Parameters:   InArray  - The input byte array
        ' Return:       encoded byte array
        '
        ' Description:
        '   Convert a byte array to a Base 64 encoded byte array. Coerce 3 bytes into
        '   4 by converting 3, 8 bit bytes into 4, 6 bit values. Each 6 bit value
        '   (0 to 63) is then used as a pointer into a base64 byte array to derive a
        '   character.
        '******************************************************************************

        Dim lInPtr as integer         ' pointer into input array
        Dim lOutPtr as integer         ' pointer into output array
        Dim outArray() As Byte         ' output byte array buffer
        Dim lLen as integer         ' number of extra bytes past 3 byte boundry
        Dim iNewLine as integer         ' line counter

        ' if size of input array is not a multiple of 3,
        ' increase it to the next multiple of 3
        lLen = (UBound(InArray) - LBound(InArray) + 1) Mod 3
        If lLen Then
            lLen = 3 - lLen
            ReDim Preserve InArray(UBound(InArray) + lLen)
        End If

        ' create an output buffer
        ReDim outArray(UBound(InArray) * 2 + 100)

        ' step through the input array, 3 bytes at a time
        For lInPtr = 0 To UBound(InArray) Step 3

            ' add CrLf as required
            If iNewLine = 19 Then
                outArray(lOutPtr) = 13
                outArray(lOutPtr + 1) = 10
                lOutPtr = lOutPtr + 2
                iNewLine = 0
            End If

            ' convert 3 bytes into 4 base 64 encoded bytes
            outArray(lOutPtr) = pbBase64Byt((InArray(lInPtr) And &HFC) \ 4)
            outArray(lOutPtr + 1) = pbBase64Byt((InArray(lInPtr) And &H3) * &H10 + (InArray(lInPtr + 1) And &HF0) \ &H10)
            outArray(lOutPtr + 2) = pbBase64Byt((InArray(lInPtr + 1) And &HF) * 4 + (InArray(lInPtr + 2) And &HC0) \ &H40)
            outArray(lOutPtr + 3) = pbBase64Byt(InArray(lInPtr + 2) And &H3F)

            ' update pointers
            lOutPtr = lOutPtr + 4
            iNewLine = iNewLine + 1
        Next

        ' add terminator '=' as required
        Select Case lLen
            Case 1 : outArray(lOutPtr - 1) = 61
            Case 2 : outArray(lOutPtr - 1) = 61 : outArray(lOutPtr - 2) = 61
        End Select

        ' add CrLf if not already there
        If outArray(lOutPtr - 2) <> 13 And AddCRLF Then
            outArray(lOutPtr) = 13
            outArray(lOutPtr + 1) = 10
            lOutPtr = lOutPtr + 2
        End If

        ' resize output buffer and return
        ReDim Preserve outArray(lOutPtr - 1)
        EncodeBase64Byte = outArray

    End Function

End Module

Module modCostCode
    Public Function ConvertCostToCode(ByVal Plaintext As String) As String
        If IsNoOne Or IsDoubleR Then
            ConvertCostToCode = ConvertCostToCode_Alternate(Plaintext)
        Else
            ConvertCostToCode = ConvertCostToCode_Standard(Plaintext)
        End If
    End Function

    Public Function ConvertCostToCode_Alternate(ByVal Plaintext As String) As String
        ' Turn 345.21 to "xxx345x" and "48.89" to "xxxx49x"
        ' x is random digit
        ' Round and drop cents
        ' One character to the right of price
        ' 7-character length string

        Const TrimLength As Integer = 7
        Dim C As Decimal, F As String, I As Integer, N As Integer
        If Trim(Plaintext) Like "*#.##" Then
            C = GetPrice(Plaintext)
            C = Math.Round(C, 0)
#If True Then
            ConvertCostToCode_Alternate = CurrencyFormat(C, True, False, True) & RandomDigit()
            '    Do While Len(ConvertCostToCode_Alternate) < TrimLength
            For I = 1 To 3
                ConvertCostToCode_Alternate = RandomDigit() & ConvertCostToCode_Alternate
            Next
            '    Loop
#Else
    N = 0: For I = 1 To Len(Plaintext): N = N + Asc(Mid(Plaintext, I, 1)): Next
    ConvertCostToCode_Alternate = CurrencyFormat(C, True, False, True) & Right("" & N, 1)
    If Len(ConvertCostToCode_Alternate) < TrimLength Then ' prevent 1 or 2 in first digit before price
      F = Left("" & N, 1)
      If F = "1" Then F = "3"
      If F = "2" Then F = "5"
    End If
    I = 0
    Do While Len(ConvertCostToCode_Alternate) < TrimLength
      I = I + 1
      ConvertCostToCode_Alternate = Mid("" & Asc(Mid(Plaintext, I, 1)) * Asc(Mid(Plaintext, I, 1)), 2, 1) & ConvertCostToCode_Alternate
    Loop
#End If
        Else
            ConvertCostToCode_Alternate = ConvertCostToCode_Standard(Plaintext)
        End If
    End Function

    Public Function ConvertCostToCode_Standard(ByVal Plaintext As String) As String
        Dim I As Integer, CodeLetter As String
        ConvertCostToCode_Standard = ""
        For I = 1 To Len(Plaintext)
            ConvertCostToCode_Standard = ConvertCostToCode_Standard & GetCostCode(Mid(Plaintext, I, 1))
        Next
    End Function

    Public Function GetCostCode(ByVal strStyle As String) As String
        Select Case Left(strStyle, 1)
            Case "1" To "9", "0" : GetCostCode = QueryTicketCode(Left(strStyle, 1))
            Case Else : GetCostCode = Left(strStyle, 1)
        End Select
    End Function

End Module

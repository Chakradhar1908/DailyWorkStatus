Module modCSV
    Public Function CSVField(ByVal Line As String, ByVal FieldNo As Integer, Optional ByVal vDefault As String = "") As String
        Dim I As Integer, X As Integer

        CSVField = ""
        If FieldNo <= 0 Then Err.Raise(-1, , "CSV Fields start at 1, not zero") ' FieldNo = 1

        For I = 1 To FieldNo
            CSVField = ""
            If Line = "" Then GoTo Final
            If Left(Line, 1) = """" Then
                X = InStr(2, Line, """")
                Do While Mid(Line, X + 1, 1) = """"
                    X = InStr(X + 2, Line, """")
                Loop
                CSVField = Mid(Line, 2, X - 2)
                CSVField = Replace(CSVField, """""", """")
                X = InStr(X, Line, ",")
            Else
                X = InStr(Line, ",")
                If X = 0 Then CSVField = Line : GoTo Final
                CSVField = Mid(Line, 1, X - 1)
            End If
            If X = 0 Then
                Line = ""
            Else
                Line = Mid(Line, X + 1)
            End If
        Next

Final:
        If CSVField = "" Then CSVField = vDefault
    End Function
    'Public Function testcsv()
    '  Dim A As String, N as integer
    '  A = ""
    '  A = ProtectCSV("aaa")
    '  A = A & "," & ProtectCSV("bb""b")
    '  A = A & "," & ProtectCSV("cc,c")
    '  A = A & "," & ProtectCSV("dd"",d")
    '  A = A & "," & ProtectCSV("""e""")
    '  A = A & "," & ProtectCSV(","",f")
    '
    '  For N = 1 To 7
    '
    'If N = 7 Then Stop
    '    Debug.Print "" & N & ": " & CSVField(A, N)
    '  Next
    'End Function

    Public Function CSVLine(ParamArray Strs() As Object) As String
        Dim El As Object, NotFirst As Boolean
        For Each El In Strs
            CSVLine = CSVLine & IIf(NotFirst, ",", "") & ProtectCSV(El)
            NotFirst = True
        Next
        '  CSVLine = CSVLine & vbCrLf
    End Function

    Public Function CSVCurrency(ByVal Curr As String) As String
        CSVCurrency = CurrencyFormat(GetPrice(Curr), False, False, True)  ' Prevent $ and ,  --  This is the same as SQLCurrency
    End Function

    Public Function ProtectCSV(ByVal Str As String) As String
        Dim Protect As Boolean
        If InStr(Str, """") > 0 Then Protect = True
        If InStr(Str, ",") > 0 Then Protect = True

        If Protect Then Str = """" & Replace(Str, """", """""") & """"
        ProtectCSV = Str
    End Function
End Module

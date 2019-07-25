Module modCSV
    Public Function CSVField(ByVal Line As String, ByVal FieldNo as integer, Optional ByVal vDefault As String = "") As String
        Dim I as integer, X As Integer

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

End Module

Imports Scripting
Module JSON
    Private psErrors As String
    Public Function Parse(ByRef Str As String) As Object
        Dim Index As Integer

        Index = 1
        psErrors = ""
        On Error Resume Next
        skipChar(Str, Index)
        Select Case Mid(Str, Index, 1)
            Case "{"
                Parse = parseObject(Str, Index)
            Case "["
                Parse = parseArray(Str, Index)
            Case Else
                psErrors = "Invalid JSON"
        End Select
    End Function

    Private Function parseArray(ByRef Str As String, ByRef Index As Integer) As Collection
        parseArray = New Collection

        ' "["
        skipChar(Str, Index)
        If Mid(Str, Index, 1) <> "[" Then
            psErrors = psErrors & "Invalid Array at position " & Index & " : " + Mid(Str, Index, 20) & vbCrLf
            Exit Function
        End If

        Index = Index + 1

        Do
            skipChar(Str, Index)
            If "]" = Mid(Str, Index, 1) Then
                Index = Index + 1
                Exit Do
            ElseIf "," = Mid(Str, Index, 1) Then
                Index = Index + 1
                skipChar(Str, Index)
            ElseIf Index > Len(Str) Then
                psErrors = psErrors & "Missing ']': " & Right(Str, 20) & vbCrLf
                Exit Do
            End If

            ' add value
            On Error Resume Next
            parseArray.Add(parseValue(Str, Index))
            If Err.Number <> 0 Then
                psErrors = psErrors & Err.Description & ": " & Mid(Str, Index, 20) & vbCrLf
                Exit Do
            End If
        Loop
    End Function

    Private Function parseValue(ByRef Str As String, ByRef Index As Integer) As Object

        skipChar(Str, Index)

        Select Case Mid(Str, Index, 1)
            Case "{"
                parseValue = parseObject(Str, Index)
            Case "["
                parseValue = parseArray(Str, Index)
            Case """", "'"
                parseValue = parseString(Str, Index)
            Case "t", "f"
                parseValue = parseBoolean(Str, Index)
            Case "n"
                parseValue = parseNull(Str, Index)
            Case Else
                parseValue = parseNumber(Str, Index)
        End Select

    End Function

    Private Function parseNumber(ByRef Str As String, ByRef Index As Integer) As Object

        Dim Value As String
        Dim Charr As String

        skipChar(Str, Index)
        Do While Index > 0 And Index <= Len(Str)
            Charr = Mid(Str, Index, 1)
            If InStr("+-0123456789.eE", Charr) Then
                Value = Value & Charr
                Index = Index + 1
            Else
                parseNumber = CDec(Value)
                Exit Function
            End If
        Loop
    End Function

    Private Function parseNull(ByRef Str As String, ByRef Index As Integer) As Object

        skipChar(Str, Index)
        If Mid(Str, Index, 4) = "null" Then
            parseNull = Nothing
            Index = Index + 4
        Else
            psErrors = psErrors & "Invalid null value at position " & Index & " : " & Mid(Str, Index) & vbCrLf
        End If
    End Function

    Private Function parseBoolean(ByRef Str As String, ByRef Index As Integer) As Boolean

        skipChar(Str, Index)
        If Mid(Str, Index, 4) = "true" Then
            parseBoolean = True
            Index = Index + 4
        ElseIf Mid(Str, Index, 5) = "false" Then
            parseBoolean = False
            Index = Index + 5
        Else
            psErrors = psErrors & "Invalid Boolean at position " & Index & " : " & Mid(Str, Index) & vbCrLf
        End If
    End Function

    Private Function parseString(ByRef Str As String, ByRef Index As Integer) As String
        Dim Quote As String
        Dim Charr As String
        Dim Code As String

        Dim sb As New cStringBuilder

        skipChar(Str, Index)
        Quote = Mid(Str, Index, 1)
        Index = Index + 1

        Do While Index > 0 And Index <= Len(Str)
            Charr = Mid(Str, Index, 1)
            Select Case (Charr)
                Case "\"
                    Index = Index + 1
                    Charr = Mid(Str, Index, 1)
                    Select Case (Charr)
                        Case """", "\", "/", "'"
                            sb.Append(Charr)
                            Index = Index + 1
                        Case "b"
                            'sb.Append vbBack
                            sb.Append(VBA.Constants.vbBack)
                            Index = Index + 1
                        Case "f"
                            'sb.Append vbFormFeed
                            sb.Append(VBA.Constants.vbFormFeed)
                            Index = Index + 1
                        Case "n"
                            sb.Append(VBA.Constants.vbLf)
                            Index = Index + 1
                        Case "r"
                            sb.Append(VBA.Constants.vbCr)
                            Index = Index + 1
                        Case "t"
                            sb.Append(VBA.Constants.vbTab)
                            Index = Index + 1
                        Case "u"
                            Index = Index + 1
                            Code = Mid(Str, Index, 4)
                            sb.Append(ChrW(Val("&h" + Code)))
                            Index = Index + 4
                    End Select
                Case Quote
                    Index = Index + 1

                    parseString = sb.ToString
                    sb = Nothing

                    Exit Function

                Case Else
                    sb.Append(Charr)
                    Index = Index + 1
            End Select
        Loop

        parseString = sb.ToString
        sb = Nothing
    End Function

    Private Function parseObject(ByRef Str As String, ByRef Index As Integer) As Dictionary

        parseObject = New Dictionary
        Dim sKey As String

        ' "{"
        skipChar(Str, Index)
        If Mid(Str, Index, 1) <> "{" Then
            psErrors = psErrors & "Invalid Object at position " & Index & " : " & Mid(Str, Index) & vbCrLf
            Exit Function
        End If

        Index = Index + 1

        Do
            skipChar(Str, Index)
            If "}" = Mid(Str, Index, 1) Then
                Index = Index + 1
                Exit Do
            ElseIf "," = Mid(Str, Index, 1) Then
                Index = Index + 1
                skipChar(Str, Index)
            ElseIf Index > Len(Str) Then
                psErrors = psErrors & "Missing '}': " & Right(Str, 20) & vbCrLf
                Exit Do
            End If


            ' add key/value pair
            sKey = parseKey(Str, Index)
            On Error Resume Next

            parseObject.Add(sKey, parseValue(Str, Index))
            If Err.Number <> 0 Then
                psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
                Exit Do
            End If
        Loop
EH:
    End Function

    Private Function parseKey(ByRef Str As String, ByRef Index As Integer) As String
        Dim Dquote As Boolean
        Dim Squote As Boolean
        Dim Charr As String

        skipChar(Str, Index)
        Do While Index > 0 And Index <= Len(Str)
            Charr = Mid(Str, Index, 1)
            Select Case (Charr)
                Case """"
                    Dquote = Not Dquote
                    Index = Index + 1
                    If Not Dquote Then
                        skipChar(Str, Index)
                        If Mid(Str, Index, 1) <> ":" Then
                            psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
                            Exit Do
                        End If
                    End If
                Case "'"
                    Squote = Not Squote
                    Index = Index + 1
                    If Not Squote Then
                        skipChar(Str, Index)
                        If Mid(Str, Index, 1) <> ":" Then
                            psErrors = psErrors & "Invalid Key at position " & Index & " : " & parseKey & vbCrLf
                            Exit Do
                        End If
                    End If
                Case ":"
                    Index = Index + 1
                    If Not Dquote And Not Squote Then
                        Exit Do
                    Else
                        parseKey = parseKey & Charr
                    End If
                Case Else
                    If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Charr) Then
                    Else
                        parseKey = parseKey & Charr
                    End If
                    Index = Index + 1
            End Select
        Loop
    End Function

    Private Sub skipChar(ByRef Str As String, ByRef Index As Integer)
        Dim bComment As Boolean
        Dim bStartComment As Boolean
        Dim bLongComment As Boolean
        Do While Index > 0 And Index <= Len(Str)
            Select Case Mid(Str, Index, 1)
                Case vbCr, vbLf
                    If Not bLongComment Then
                        bStartComment = False
                        bComment = False
                    End If

                Case vbTab, " ", "(", ")"

                Case "/"
                    If Not bLongComment Then
                        If bStartComment Then
                            bStartComment = False
                            bComment = True
                        Else
                            bStartComment = True
                            bComment = False
                            bLongComment = False
                        End If
                    Else
                        If bStartComment Then
                            bLongComment = False
                            bStartComment = False
                            bComment = False
                        End If
                    End If

                Case "*"
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                        bLongComment = True
                    Else
                        bStartComment = True
                    End If

                Case Else
                    If Not bComment Then
                        Exit Do
                    End If
            End Select

            Index = Index + 1
        Loop
    End Sub
End Module

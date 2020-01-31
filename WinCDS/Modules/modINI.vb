Module modINI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize as integer, ByVal lpFileName As String) as integer
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) as Integer
    Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    Public Function ReadIniValue(ByVal INIPath As String, ByVal key As String, ByVal Variable As String, Optional ByVal vDefault As String = "") As String
        On Error Resume Next
        ReadIniValue = INIRead(key, Variable, INIPath)
        If ReadIniValue = "" Then ReadIniValue = vDefault
    End Function

    Public Function WriteIniValue(ByVal INIPath As String, ByVal PutKey As String, ByVal PutVariable As String, ByVal PutValue As String, Optional ByVal DeleteOnEmpty As Boolean = False)
        On Error Resume Next
        INIWrite(PutKey, PutVariable, PutValue, INIPath)
        WriteIniValue = INIRead(PutKey, PutVariable, INIPath)
    End Function

    Public Function INIRead(ByVal sSection As String, ByVal sKeyName As String, ByVal sINIFileName As String) As String
        On Error Resume Next
        Dim sRet As String
        Dim Length As Integer
        'sRet = String(255, Chr(0))
        'sRet = New String(Chr(0), 255)
        sRet = Space(255)
        Length = Len(sRet)
        'INIRead = Microsoft.VisualBasic.Left(sRet, GetPrivateProfileString(sSection, sKeyName, "", sRet, Len(sRet), sINIFileName))
        Length = GetPrivateProfileString(sSection, sKeyName, "", sRet, Length, sINIFileName)
        'INIRead = Microsoft.VisualBasic.Left(sRet, GetPrivateProfileString(sSection, sKeyName, "", sRet, Length, sINIFileName))
        INIRead = Left(sRet, Length)
    End Function

    Public Function INIWrite(ByVal sSection As String, ByVal sKeyName As String, ByVal sNewString As String, ByVal sINIFileName As String) As Boolean
        On Error Resume Next
        WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
        INIWrite = (Err.Number = 0)
    End Function

    Public Function INISections(ByVal FileName As String) As String()
        On Error Resume Next
        Dim strBuffer As String, intLen As Integer

        Do While (intLen = Len(strBuffer) - 2) Or (intLen = 0)
            If strBuffer = vbNullString Then
                strBuffer = Space(256)
            Else
                'strBuffer = String(Len(strBuffer) * 2, 0)
                strBuffer = New String("0", Len(strBuffer) * 2)
            End If

            intLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), FileName)
        Loop

        strBuffer = Left(strBuffer, intLen)
        INISections = Split(strBuffer, vbNullChar)
        ReDim Preserve INISections(UBound(INISections) - 1)
    End Function

    Public Function INISectionKeys(ByVal FileName As String, ByVal Section As String) As String()
        On Error Resume Next
        Dim strBuffer As String, intLen As Integer
        Dim I As Integer, N As Integer
        Dim RET() As String

        Do While (intLen = Len(strBuffer) - 2) Or (intLen = 0)
            If strBuffer = vbNullString Then
                strBuffer = Space(256)
            Else
                'strBuffer = String(Len(strBuffer) * 2, 0)
                strBuffer = New String("0", Len(strBuffer) * 2)
            End If

            intLen = GetPrivateProfileSection(Section, strBuffer, Len(strBuffer), FileName)
            If intLen = 0 Then Exit Function
        Loop

        strBuffer = Left(strBuffer, intLen)
        RET = Split(strBuffer, vbNullChar)
        ReDim Preserve RET(UBound(RET) - 1)
        For I = LBound(RET) To UBound(RET)
            N = InStr(RET(I), "=")
            If N > 0 Then
                RET(I) = Left(RET(I), N - 1)
            Else
                Debug.Print("modINI.INISectionKeys - No '=' character found in line.  Section=" & Section & ", Line=" & RET(I) & ", file=" & FileName)
            End If
        Next
        INISectionKeys = RET
    End Function

    Public Function INISectionAsHashTable(ByVal FileName As String, ByVal Section As String) As clsHashTable
        On Error Resume Next
        Dim I As Integer, L
        INISectionAsHashTable = New clsHashTable
        For Each L In INISectionKeys(FileName, Section)
            INISectionAsHashTable.Add(L, INIRead(Section, L, FileName))
        Next
    End Function

End Module

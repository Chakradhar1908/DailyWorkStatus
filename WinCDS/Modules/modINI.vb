Module modINI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize as integer, ByVal lpFileName As String) as integer
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) as integer
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
        'sRet = String(255, Chr(0))
        sRet = New String(Chr(0), 255)

        'INIRead = Left(sRet, GetPrivateProfileString(sSection, sKeyName, "", sRet, Len(sRet), sINIFileName))
    End Function
    Public Function INIWrite(ByVal sSection As String, ByVal sKeyName As String, ByVal sNewString As String, ByVal sINIFileName As String) As Boolean
        On Error Resume Next
        WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
        INIWrite = (Err.Number = 0)
    End Function

End Module

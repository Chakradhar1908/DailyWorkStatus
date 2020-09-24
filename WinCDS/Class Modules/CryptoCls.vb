Public Class CryptoCls
    Dim hCryptProv As Integer
    Dim hClientWriteKey As Integer
    Private Const CALG_MD5 As Integer = 32771
    Private Const HP_HASHVAL As Integer = 2
    Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Integer, ByVal Algid As Integer, ByVal hSessionKey As Integer, ByVal dwFlags As Integer, ByRef phHash As Integer) As Integer
    Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Integer, ByVal pbData As String, ByVal dwDataLen As Integer, ByVal dwFlags As Integer) As Integer
    Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Integer, ByVal dwParam As Integer, ByVal pbData As String, ByRef pdwDataLen As Integer, ByVal dwFlags As Integer) As Integer
    Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Integer) As Integer
    Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hSessionKey As Integer, ByVal hHash As Integer, ByVal Final As Integer, ByVal dwFlags As Integer, ByVal pbData As String, ByRef pdwDataLen As Integer, ByVal dwBufLen As Integer) As Integer

    Public Function RC4_Encrypt(ByVal Plaintext As String) As String
        'Encrypt with Client Write Key
        Dim lngLength As Integer
        Dim lngReturnValue As Integer

        lngLength = Len(Plaintext)
        lngReturnValue = CryptEncrypt(hClientWriteKey, 0, False, 0, Plaintext, lngLength, lngLength)

        RC4_Encrypt = Plaintext
    End Function

    Public Function MD5_Hash(ByVal TheString As String) As String
        'Digest a String using MD5
        Dim lngReturnValue As Integer
        Dim strHash As String
        Dim hHash As Integer
        Dim lngHashLen As Integer

        lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
        lngReturnValue = CryptHashData(hHash, TheString, Len(TheString), 0)
        lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0)
        'strHash = String(lngHashLen, vbNullChar)
        strHash = New String(vbNullChar, lngHashLen)
        lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0)

        If hHash <> 0 Then CryptDestroyHash(hHash)

        MD5_Hash = strHash
    End Function
End Class

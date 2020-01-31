Module SSLv2
    Public SecureSession As CryptoCls
    Public Layer As Integer
    Public InBuffer As String
    Public Processing As Boolean
    Public SeekLen As Integer

    'Encryption Keys
    Public MASTER_KEY As String
    Public CLIENT_READ_KEY As String
    Public CLIENT_WRITE_KEY As String

    'Server Attributes
    Public PUBLIC_KEY As String
    Public ENCODED_CERT As String
    Public CONNECTION_ID As String

    'Counters
    Public SEND_SEQUENCE_NUMBER As Double
    Public RECV_SEQUENCE_NUMBER As Double

    'Hand Shake Variables
    Public CLIENT_HELLO As String
    Public CHALLENGE_DATA As String

    Public Sub SSLSend(ByRef Socket As MSWinsockLib.Winsock, ByVal Plaintext As String)
        'Send Plaintext as an Encrypted SSL Record
        Dim SSLRecord As String
        Dim OtherPart As String
        Dim SendAnother As Boolean

        If Len(Plaintext) > 32751 Then
            SendAnother = True
            Plaintext = Mid(Plaintext, 1, 32751)
            OtherPart = Mid(Plaintext, 32752)
        Else
            SendAnother = False
        End If

        SSLRecord = AddMACData(Plaintext)
        SSLRecord = SecureSession.RC4_Encrypt(SSLRecord)
        SSLRecord = AddRecordHeader(SSLRecord)

        Socket.SendData(SSLRecord)

        If SendAnother = True Then
            SSLSend(Socket, OtherPart)
        End If
    End Sub

    Private Function AddMACData(ByVal Plaintext As String) As String
        'Prepend MAC Data to the Plaintext
        AddMACData = SecureSession.MD5_Hash(CLIENT_WRITE_KEY & Plaintext & SendSequence) & Plaintext
    End Function

    Private Function AddRecordHeader(ByVal RecordData As String) As String
        'Prepend SLL Record Header to the Data Record
        Dim FirstChar As String
        Dim LastChar As String
        Dim TheLen As Integer

        TheLen = Len(RecordData)

        FirstChar = Chr(128 + (TheLen \ 256))
        LastChar = Chr(TheLen Mod 256)

        AddRecordHeader = FirstChar & LastChar & RecordData
        IncrementSend()
    End Function

    Private Function SendSequence() As String
        'Convert Send Counter to a String
        Dim TempString As String
        Dim TempSequence As Double
        Dim TempByte As Double
        Dim I As Integer

        TempSequence = SEND_SEQUENCE_NUMBER

        For I = 1 To 4
            TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
            TempSequence = Int(TempSequence / 256)
            TempString = Chr(TempByte) & TempString
        Next

        SendSequence = TempString
    End Function

    Public Sub IncrementSend()
        'Increment Counter for Each Record Sent
        SEND_SEQUENCE_NUMBER = SEND_SEQUENCE_NUMBER + 1
        If SEND_SEQUENCE_NUMBER = 4294967296.0# Then SEND_SEQUENCE_NUMBER = 0
    End Sub
End Module

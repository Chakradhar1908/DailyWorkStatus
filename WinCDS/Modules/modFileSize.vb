Imports Microsoft.VisualBasic.Interaction
Module modFileSize
    Public Const FileSize_1KB = 1024
    Public Const FileSize_1MB = 1048576
    Public Const FileSize_1GB = 1073741824
    Public Const FileSize_1TB = 1099511627776.0#
    Public Const FileSize_1PB = 1.12589990684262E+15

    Public Function DescribeFileSize(ByVal Sz As Double, Optional ByVal Style As Integer = 0) As String
        '1 Bit = Binary Digit
        '8 Bits = 1 Byte
        '1024 Bytes = 1 Kilobyte
        '1024 Kilobytes = 1 Megabyte
        '1024 Megabytes = 1 Gigabyte
        '1024 Gigabytes = 1 Terabyte
        '1024 Terabytes = 1 Petabyte
        '1024 Petabytes = 1 Exabyte
        '1024 Exabytes = 1 Zettabyte
        '1024 Zettabytes = 1 Yottabyte
        '1024 Yottabytes = 1 Brontobyte
        '1024 Brontobytes = 1 Geopbyte
        Const tKB = 1024
        Const tMB = 1048576
        Const tGB = 1073741824
        Const tTB = 1099511627776.0#
        Const tPB = 1.12589990684262E+15
        Dim D As Double, S As String
        If Sz < tKB Then
            DescribeFileSize = "" & Sz
        ElseIf Sz < tMB Then
            D = 1 + (Sz / tKB)
            S = Format(D, "0.00")
            If Len(S) > 4 Then S = Format(D, "0.0")
            If Len(S) > 4 Then S = Fix(D)
            DescribeFileSize = S & Switch(Style = 0, " KB", True, "K")
        ElseIf Sz < tGB Then
            D = 1 + (Sz / tMB)
            S = Format(D, "0.00")
            If Len(S) > 4 Then S = Format(D, "0.0")
            If Len(S) > 4 Then S = Fix(D)
            DescribeFileSize = S & Switch(Style = 0, " MB", True, "M")
        ElseIf Sz < tTB Then
            D = 1 + (Sz / tGB)
            S = Format(D, "0.00")
            If Len(S) > 4 Then S = Format(D, "0.0")
            If Len(S) > 4 Then S = Fix(D)
            DescribeFileSize = S & Switch(Style = 0, " GB", True, "G")
        ElseIf Sz < tPB Then
            D = 1 + (Sz / tTB)
            S = Format(D, "0.00")
            If Len(S) > 4 Then S = Format(D, "0.0")
            If Len(S) > 4 Then S = Fix(D)
            DescribeFileSize = S & Switch(Style = 0, " TB", True, "T")
        Else
            D = 1 + (Sz / tPB)
            S = Format(D, "0.00")
            If Len(S) > 4 Then S = Format(D, "0.0")
            If Len(S) > 4 Then S = Fix(D)
            DescribeFileSize = S & Switch(Style = 0, " PB", True, "P")
        End If
    End Function
End Module

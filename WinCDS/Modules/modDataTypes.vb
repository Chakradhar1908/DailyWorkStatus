Module modDataTypes
    Structure SearchNew
        <VBFixedString(16)> Dim Style As String
        <VBFixedString(4)> Dim Code As String
        <VBFixedString(5)> Dim RN As String
    End Structure

    Public Structure EomFile
        <VBFixedString(24)> Dim LastName As String
        <VBFixedString(8)> Dim LeaseNo As String
        <VBFixedString(9)> Dim GrossSale As String
        <VBFixedString(9)> Dim TotDeposit As String
        <VBFixedString(9)> Dim Balance As String
        <VBFixedString(10)> Dim LastPay As String
        <VBFixedString(1)> Dim Status As String
        <VBFixedString(14)> Dim Salesman As String
    End Structure

End Module

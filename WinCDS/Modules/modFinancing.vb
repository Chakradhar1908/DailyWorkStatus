Module modFinancing
    Public Const ArNo_AddOnRecordSeparator As String = "-"
    Public Const ArNo_AddOnRecordToken As String = "AddOnRecord"
    Public Const ArNo_AddOnRecordIndicator As String = ArNo_AddOnRecordSeparator & ArNo_AddOnRecordToken & ArNo_AddOnRecordSeparator
    Public Const ArNo_AddOnRecordPattern_LIKE As String = "*" & ArNo_AddOnRecordToken & "*"
    Public Const ArNo_AddOnRecordPattern_SQL As String = "%" & ArNo_AddOnRecordToken & "%" ' MS Access requires % for wildcard

    Public Function ArNoIsAddOnRecord(ByVal ArNo As String) As Boolean
        ArNoIsAddOnRecord = ArNo Like ArNo_AddOnRecordPattern_LIKE
    End Function

    Public Function CalculateSIR(ByVal Balance As Decimal, ByVal TargetAPR As Double, ByVal Months As Integer) As Double
        Dim SIR As Double, FC As Decimal, Incr As Boolean, Delta As Double, APR As Double

        Dim Mx As Integer
        If Balance = 0 Then Exit Function
        SIR = 0.12
        FC = ((Balance * SIR) / 12 * Months)
        APR = CalculateAPR(Balance, FC, Months)
        Delta = 0.01
        If APR = TargetAPR Then   ' not likely..
            CalculateSIR = SIR
            Exit Function
        End If
        Incr = TargetAPR > APR

        Mx = 0
        Do While SIR <> APR And Delta > 0.00000001
            Mx = Mx + 1
            If TargetAPR > APR Then   ' target is higher than current
                If Incr Then            ' still going up
                    SIR = SIR + Delta
                Else                    ' we were going down, but went too far!
                    Delta = Delta / 10.0#
                    SIR = SIR + Delta
                    Incr = True           ' turn around and go back up (after changing delta)
                End If
            Else                      ' target is lower
                If Not Incr Then
                    SIR = SIR - Delta
                Else
                    Delta = Delta / 10.0#
                    SIR = SIR - Delta
                    Incr = False          ' turn around and go back down (after changing delta)
                End If
            End If
            FC = ((Balance * SIR) / 12 * Months)
            APR = CalculateAPR(Balance, FC, Months)
            If Mx > 1000000 Then Exit Do
        Loop
        If APR > TargetAPR Then
            SIR = SIR - 10 * Delta  'always keep below target, use 10* b/c it would have been shrunk
            FC = ((Balance * SIR) / 12 * Months)
            APR = CalculateAPR(Balance, FC, Months)
        End If
        CalculateSIR = SIR
        'Debug.Print "APR=" & APR
        'Debug.Print "SIR=" & SIR
    End Function

    Public Function CalculateAPR(ByVal Balance As Decimal, ByVal FinanceCharge As Decimal, ByVal Months As Integer, Optional ByVal DeferredMonths As Integer = 0) As Double
        Dim T As Double
        On Error GoTo BadVBARateFunction
        CalculateAPR = 1200 * Financial.Rate(Months, -(Balance + FinanceCharge) / Months, Balance)
        Exit Function
BadVBARateFunction:
        T = (3 * Balance * (Months + DeferredMonths + 1) + (FinanceCharge * (Months + DeferredMonths + 1)))
        If T <> 0 Then
            '  CalculateAPR = 100 * (6 * 12 * FinanceCharge) / (3 * Balance * (Months + DeferredMonths + 1) + (FinanceCharge * (Months + DeferredMonths + 1)))
            CalculateAPR = 100 * (6 * 12 * FinanceCharge) / T
        End If
    End Function

    Public Sub GetPreviousContractTerms(ByVal ArNo As String, Optional ByVal StoreNo As Long = 0, Optional ByRef Prev As Currency, Optional ByRef Sale As Currency, Optional ByRef Deposit As Currency, Optional ByRef DocFee As Currency, Optional ByRef tLife As Currency, Optional ByRef tAcc As Currency, Optional ByRef tProp As Currency, Optional ByRef tIUI As Currency, Optional ByRef tInt As Currency, Optional ByRef tIntST As Currency)
        Dim RS As Recordset
        Dim CL As Boolean, CA As Boolean, cP As Boolean, cU As Boolean ' , cI as boolean
        Dim IsSale As Boolean
        Dim F As String
  Set RS = GetRecordsetBySQL("SELECT * FROM [Transactions] WHERE ArNo='" & ArNo & "' ORDER BY [TransactionID]", , GetDatabaseAtLocation(StoreNo))
    
  Sale = 0
        Deposit = 0
        Prev = 0
        DocFee = 0
        tLife = 0
        tAcc = 0
        tProp = 0
        tIUI = 0
        tInt = 0
        tIntST = 0

        Do While Not RS.EOF
            F = IfNullThenNilString(RS("Type"))
            '    Debug.Print F
            If IsSale Or F Like "NewSal*" Then
                If F Like "NewSal*" Then
                    IsSale = True
                    Sale = IfNullThenZeroCurrency(RS("Charges"))
                    Deposit = IfNullThenZeroCurrency(RS("Credits"))
                    Prev = IfNullThenZeroCurrency(RS("Balance")) - IfNullThenZeroCurrency(RS("Charges")) + IfNullThenZeroCurrency(RS("Credits"))
                    DocFee = 0
                    tLife = 0
                    tAcc = 0
                    tProp = 0
                    tIUI = 0
                    tInt = 0
                    tIntST = 0
                Else
                    Select Case F
                        Case arPT_New
                        Case arPT_Doc : DocFee = IfNullThenZeroCurrency(RS("Charges"))
                        Case arPT_Lif : tLife = IfNullThenZeroCurrency(RS("Charges"))
                        Case arPT_Acc : tAcc = IfNullThenZeroCurrency(RS("Charges"))
                        Case arPT_Pro : tProp = IfNullThenZeroCurrency(RS("Charges"))
                        Case arPT_Int : tInt = IfNullThenZeroCurrency(RS("Charges"))
                        Case arPT_Tax : tIntST = IfNullThenZeroCurrency(RS("Charges"))
                    End Select
                End If
            End If
            RS.MoveNext
        Loop
  
  Set RS = Nothing
End Sub

End Module

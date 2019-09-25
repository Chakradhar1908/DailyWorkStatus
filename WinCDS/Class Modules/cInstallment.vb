Public Class cInstallment
    Public ArNo As String
    Public LastName As String
    Public Telephone As String
    Public MailIndex As Integer
    Public Financed As Decimal
    Public PerMonth As Decimal
    Public Months As Integer
    Public Rate As Double
    Public LateDueOn As Integer
    Public LateCharge As String
    Public DeliveryDate As Date
    Public FirstPayment As Date
    Public CashOpt As Integer
    Public TotPaid As Decimal
    Public Balance As Decimal
    Public LateChargeBal As Decimal
    Public Status As String
    Public INTEREST As Decimal
    Public Life As Decimal
    Public Accident As Decimal
    Public Prop As Decimal
    Public WriteOffDate As String
    Public SendNotice As String
    Public LastMetro426Status As String
    Public InterestSalesTax As Decimal
    Public APR As Double
    Public Period As String
    Public LifeType As Integer
    Public Satisfied As String
    Public SatisfiedDate As Date
    Public IUI As Decimal

    Public DataBase As String

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME As String = "InstallmentInfo"
    Private Const TABLE_INDEX As String = "ArNo"


    'Implements CDataAccess
    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose()
    End Sub

    Public Function Load(ByRef KeyVal As String, Optional ByRef KeyName As String = "") As Boolean
        ' Checks the database for a matching TransactionID.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        Load = False
        ' Search for the Style
        If KeyName = "" Then
            DataAccess.Records_OpenIndexAt(KeyVal)
        ElseIf Left(KeyName, 1) = "#" Then
            ' This allows searching by AutoNumber - specialized to query by number
            ' since Access is exceptionally picky about quotation marks.
            DataAccess.Records_OpenFieldIndexAtNumber(Mid(KeyName, 2), KeyVal)
        Else
            DataAccess.Records_OpenFieldIndexAt(KeyName, KeyVal)
        End If

        ' Move to the first record if we can, and return success.
        If DataAccess.Records_Available Then Load = True
    End Function

    Public Function GetPayoffRevolving(Optional ByRef AsOfDate As Date = Nothing) As Decimal
        'If AsOfDate = 0 Then AsOfDate = Date
        If IsNothing(AsOfDate) Then AsOfDate = Today
        If Not IsRevolvingCharge(ArNo) Then
            ' Payoff code is in ArCard.  If you want to use this for Installment accounts, transplant it here.
            Exit Function
        Else
            GetPayoffRevolving = Balance
            '    Dim H As New cHolding
            '    H.Load ArNo, "ArNo"
            '    Do Until H.DataAccess.Record_EOF
            '      GetPayoffRevolving = GetPayoffRevolving - H.GetInterestCredit(AsOfDate)
            '      H.DataAccess.Records_MoveNext
            '    Loop
            '    DisposeDA H
        End If
    End Function

    Public Function GetLastInterestDate() As Date
        GetLastInterestDate = 0
        Dim Trans As New cTransaction
        Trans.Load ArNo, "ArNo"
  Do Until Trans.DataAccess.Record_EOF
            If Trans.TransDate > GetLastInterestDate And Trans.Charges > 0 And Trans.TransType = arPT_Int Then
                GetLastInterestDate = Trans.TransDate
            End If
            Trans.DataAccess.Records_MoveNext()
        Loop
        DisposeDA Trans
End Function

    Public Function PaidInPeriod(ByRef TargetDate As Date) As Currency
        ' Returns the amount paid during the month of TargetDate, using LateDueOn as period boundary.
        PaidInPeriod = 0
        Dim StartDate As Date, EndDate As Date
        StartDate = DateAdd("d", -Day(TargetDate) + LateDueOn, TargetDate)
        If StartDate > TargetDate Then StartDate = DateAdd("m", -1, StartDate)
        EndDate = DateAdd("m", StartDate, 1)

        Dim Trans As New cTransaction
        Trans.Load ArNo, "ArNo"
  Do Until Trans.DataAccess.Record_EOF
            If Trans.TransDate > StartDate And Trans.TransDate <= EndDate Then
                PaidInPeriod = PaidInPeriod + Trans.Credits
            End If
            Trans.DataAccess.Records_MoveNext()
        Loop
        DisposeDA Trans
End Function

    Public Function RevolvingInterestDate(ByVal DTE As Date) As Date
        Dim X As Long
        X = LateDueOn
        Do Until IsDate(Month(DTE) & "/" & X & "/" & Year(DTE))
            X = X - 1
        Loop
        RevolvingInterestDate = DateValue(Month(DTE) & "/" & X & "/" & Year(DTE))
        If RevolvingInterestDate > DTE Then RevolvingInterestDate = DateAdd("m", -1, RevolvingInterestDate)
        RevolvingInterestDate = DateAdd("d", 1, RevolvingInterestDate)
    End Function

    Public Function AddInterest(ByRef NewInterest As Currency, Optional ByRef ChargeDate As Date = 0) As Boolean
        If NewInterest = 0 Then Exit Function
        If ChargeDate = 0 Then ChargeDate = Date

        ' Record the interest in the revolving account
        INTEREST = INTEREST + NewInterest
        Balance = Balance + NewInterest
        Financed = Financed + NewInterest

        ' add a transaction history
        Dim Trans As New cTransaction
        Trans.ArNo = ArNo
        Trans.LastName = LastName
        Trans.TransDate = ChargeDate
        Trans.MailIndex = MailIndex
        Trans.TransType = IIf(NewInterest > 0, arPT_Int, arPT_crInt)
        Trans.Charges = Abs(NewInterest)
        Trans.Credits = 0
        Trans.Balance = Balance
        Trans.Receipt = ""
        Trans.Save
        DisposeDA Trans
End Function

    Public Function Save(Optional ByRef ErrDesc As String = "") As Boolean
        ErrDesc = ""
        Save = True
        On Error GoTo NoSave
        ' This instructs the class (in one simple call) to save its data members to the database.
        If DataAccess.CurrentIndex <= 0 Then            ' If we're already using the current record,
            DataAccess.Records_OpenIndexAt ArNo 'there's no reason to re-open it.
        End If
        If DataAccess.Record_Count = 0 Then
            DataAccess.Records_Add()      ' Record not found.  This means we're adding a new one.
        End If

        DataAccess.Record_Update()      ' Then load our data into the recordset.
        DataAccess.Records_Update()     ' And finally, tell the class to save the recordset.
        Exit Function

NoSave:
        ErrDesc = Err.Description
        Err.Clear()
        Save = False
    End Function

    Public Function GetChargedInPeriod(ByVal AfterDate As Date, ByVal UpToIncluding As Date, Optional ByVal tType As String = arPT_Int) As Currency
        GetChargedInPeriod = 0
        Dim Trans As New cTransaction
        Trans.Load ArNo, "ArNo"
  Do Until Trans.DataAccess.Record_EOF
            If Trans.TransDate > AfterDate And Trans.TransDate <= UpToIncluding And Trans.Charges > 0 Then
                If Trans.TransType = tType Then
                    GetChargedInPeriod = GetChargedInPeriod + Trans.Charges
                End If
            End If
            Trans.DataAccess.Records_MoveNext()
        Loop
        DisposeDA Trans
End Function

End Class

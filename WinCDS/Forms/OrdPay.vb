Public Class OrdPay
    Dim Status As String             ' Used by cmdOK, needs to be saved between clicks..
    Dim OrgHoldingStatus As String   ' Original Holding Status.
    Dim PayMethod As String
    Dim PriorBal As Decimal
    Dim Deposit As Decimal

    Dim DeliveredAuditRecord As Integer, DeliveredPayment As Decimal

    ' Cash and Audit variables.. global until we rewrite those calls.
    Dim LeaseNo As String
    Dim Note As String
    Dim Money As Decimal
    Dim Account As String
    Dim Cashier As String
    Dim Name1 As String
    Dim TransDate As String
    Dim Written As Decimal
    Dim TaxCharged1 As Decimal
    Dim ArCashSls As Decimal
    Dim Controll As Decimal
    Dim UndSls As Decimal
    Dim DelSls As Decimal
    Dim TaxRec1 As Decimal
    Dim SalesTax1 As Decimal

    Dim Approval As String

    Dim FinanceArNo As String

    Public X As Integer                  ' BillOSale.GrossMargin checks this.
    Public Sale As Decimal           ' Called by ArPaySetup
    Public TotDeposit As Decimal     ' Called by ArPaySetup

    Private LockOn As Boolean         ' Used to simulate Modal state

End Class
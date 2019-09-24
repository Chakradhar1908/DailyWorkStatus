Public Class ArCard
    Private Const FRM_W as integer = 10525
    Private Const FRM_H_MIN as integer = 7290
    Private Const FRM_H_MAX as integer = 13000

    Private WithEvents mDBAccess As CDbAccessGeneral
    Private WithEvents mDBAccessTransactions As CDbAccessGeneral

    Public PayCount as integer
    Private PayLog() As PaymentRecord ' Array for tracking payments, for receipts

    Private altPayType As String    ' used for a couple specialty operations like payoffs
    Public ArNo As String
    Public Status As String
    Public MailRec as integer
    Public DocCredit As Decimal
    Public LifeCredit As Decimal
    Public AccidentCredit As Decimal
    Public PropertyCredit As Decimal
    Public IUICredit As Decimal
    Public InterestCredit As Decimal
    Public InterestTaxCredit As Decimal

    ' These need to be eliminated, to speed up the whole program.
    'GDM Form variable to indicate that we are processing a move button
    Public UsingMoveButton As Boolean
    Public f_strDirection As String 'GDM

    Private Const NON_ALERT_COLOR as integer = &H6666CC
    Private Const ALERT_COLOR as integer = &H6666CC ' '&HC0&

    Dim Mail As MailNew
    Dim Mail2 As MailNew2

    Dim mArNo As String
    Dim CustRec as integer
    Dim CashOpt as integer

    Public INTEREST As Decimal
    Public InterestTax As Decimal
    Public DocFee As Decimal
    Public Life As Decimal
    Public Accident As Decimal
    Public Prop As Decimal
    Public IUI As Decimal

    Dim Charges As Decimal
    Dim Credits As Decimal
    Dim Balance As Decimal
    Dim TotPaid As Decimal
    Dim Financed As Decimal
    Dim TransType As String
    Dim Payoff As String
    Dim PayoffSameAsCash As Boolean
    Dim StatusChg As String
    Dim Receipt As String
    Dim NewTypee As String

    Dim TransDate As String
    Dim LastPayDate As String
    Dim LastPay As String
    Dim LateChargeBal As String
    Dim Months As String
    Dim SendNotice As String
    Dim Counter as integer

    Dim Approval As String

    Dim OpenFormAs As String
    Dim InterestDebit As Decimal  ' For Bankruptcy transactions
    Dim InterestCreditRevolving As Decimal

    Dim DoRecordAccountClosed As Boolean

    Dim cmdApplyValue As Boolean                 ' Used to determine whether this button has been clicked.
    Dim cmdReceiptValue As Boolean               ' Future Languages do not support command button value property

    Public Sub ShowArCardForDisplayOnly(ByVal nArNo As String, Optional ByVal Modal As Boolean = True, Optional ByVal AllowClose As Boolean = False, Optional ByVal AllowContractChange As Boolean = False)
        Dim OldAR As String
        If nArNo <> "" Then LoadArNo nArNo

  cmdApply.Enabled = AllowContractChange
        'cmdCreditApp.Enabled = False
        'cmdDetail.Enabled = False
        cmdEdit.Enabled = False
        cmdExport.Enabled = False
        cmdFields.Enabled = False
        cmdMakeSameAsCash.Enabled = False
        cmdMoveFirst.Enabled = False
        cmdMoveLast.Enabled = False
        cmdMoveNext.Enabled = False
        cmdMovePrevious.Enabled = False
        cmdPayoff.Enabled = False
        cmdReceipt.Enabled = False
        cmdReprintContract.Enabled = False
        'Notes_Open.Enabled = False
        cmdPrint.Enabled = False
        'cmdPrintCard.Enabled = False
        'cmdPrintLabel.Enabled = False
        cmdCancel.Enabled = AllowClose

        fraPaymentOptions.Visible = False
        fraEditOptions.Visible = False

        Show IIf(Modal, 1, 0)

  If OldAR <> "" Then ArSelect = OldAR

        '  DisposeDA RS
    End Sub

End Class
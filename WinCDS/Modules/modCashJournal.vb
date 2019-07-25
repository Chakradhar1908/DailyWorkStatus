Module modCashJournal
    Public Sub AddNewCashJournalRecord(ByVal Account As String, ByVal Money As Decimal, ByVal LeaseNo As String, ByVal Note As String, ByVal TransDate As Date, Optional ByVal Cashier As String = "", Optional ByVal vTerminal As String = "")
        '::::AddNewCashJournalRecord
        ':::SUMMARY
        ': Used to Add New Cash Journal Record.
        ':::DESCRIPTION
        ': This function is used to add new records in CashJournal table, Cash.
        ':::PARAMETERS
        ': - Account - The account number handling this currency.
        ': - Money - The amount of the CASH transaction.
        ': - LeaseNo - Sale Number.Also often called “Lease Number” by program designer.Reason unclear.
        ': - Note - Indicates a memo field attached to this transaction.
        ': - TransDate - Indicates The date of this transaction.
        ': - Cashier - Auto-determined at time of transaction.If employee name is available, it will be filled in.  Otherwise, it will use the current COMPUTER NAME
        ': - vTerminal - Terminal ID for Terminal Tracking.
        'Dim NewCashRec As CashJournalNew, RS As ADODB.Recordset

        If Money = 0 Then Exit Sub
        If Trim(Account) = "" Then Exit Sub

        'NewCashRec.Account = Trim(Account)
        'NewCashRec.Cashier = Trim(IIf(Cashier = "", GetCashierName, Cashier))
        'NewCashRec.Terminal = Trim(IIf(vTerminal = "", Terminal, vTerminal))
        'NewCashRec.LeaseNo = Trim(LeaseNo)
        'NewCashRec.Money = Money
        'NewCashRec.Note = Trim(Note)
        'NewCashRec.TransDate = TransDate

        '      RS = getRecordsetByTableLabelIndex(CashJournal_TABLE, CashJournal_INDEX, "-1", True)
        '      CashJournalNew_RecordSet_Get NewCashRec, RS
        'SetRecordsetByTableLabelIndex RS, CashJournal_TABLE, CashJournal_INDEX, "-1"
        'DisposeDA(RS)
    End Sub

End Module

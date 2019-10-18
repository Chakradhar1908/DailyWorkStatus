Module modCashJournal
    Structure CashJournalNew
        Dim CashID As Integer
        Dim LeaseNo As String
        Dim Money As Decimal
        Dim Account As String
        Dim Note As String
        Dim TransDate As String
        Dim Cashier As String
        Dim Terminal As String
    End Structure
    Public Const CashJournal_FILE As String = "CASH2.exe"
    Public Const CashJournal_FILE_RecordSize As Integer = 56
    Public Const CashJournal_TABLE As String = "Cash"
    Public Const CashJournal_INDEX As String = "LeaseNo"

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
        Dim NewCashRec As CashJournalNew, RS As ADODB.Recordset

        If Money = 0 Then Exit Sub
        If Trim(Account) = "" Then Exit Sub

        NewCashRec.Account = Trim(Account)
        NewCashRec.Cashier = Trim(IIf(Cashier = "", GetCashierName, Cashier))
        NewCashRec.Terminal = Trim(IIf(vTerminal = "", Terminal, vTerminal))
        NewCashRec.LeaseNo = Trim(LeaseNo)
        NewCashRec.Money = Money
        NewCashRec.Note = Trim(Note)
        NewCashRec.TransDate = TransDate

        RS = getRecordsetByTableLabelIndex(CashJournal_TABLE, CashJournal_INDEX, "-1", True)
        CashJournalNew_RecordSet_Get(NewCashRec, RS)
        SetRecordsetByTableLabelIndex(RS, CashJournal_TABLE, CashJournal_INDEX, "-1")
        DisposeDA(RS)
    End Sub

    Public Sub CashJournalNew_RecordSet_Get(ByRef Cj As CashJournalNew, ByRef RS As ADODB.Recordset)
        '::::CashJournalNew_RecordSet_Get
        ':::SUMMARY
        ': Get Cash Journal New Record
        ':::DESCRIPTION
        ': This function is used to create the CashJournalNew recordset after trimming its field values.
        ':::PARAMETERS
        ': - Cj
        ': - RS
        On Error Resume Next
        '    RS("CashID") = .CashID
        RS("LeaseNo").Value = Trim(Cj.LeaseNo)
        RS("Money").Value = CStr(Cj.Money)
        RS("Account").Value = Trim(Cj.Account)
        RS("Note").Value = Trim(Cj.Note)
        RS("TransDate").Value = Trim(Cj.TransDate)
        RS("Cashier").Value = Trim(Cj.Cashier)
        RS("Terminal").Value = Trim(Cj.Terminal)
    End Sub
End Module

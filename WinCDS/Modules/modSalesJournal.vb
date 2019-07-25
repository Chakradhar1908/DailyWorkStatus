Module modSalesJournal
    Public Const SalesJournal_FILE As String = "AUDIT2" & ".exe"
    Public Const SalesJournal_FILE_RecordSize as integer = 112
    Public Const SalesJournal_TABLE As String = "Audit"
    Public Const SalesJournal_INDEX As String = "SaleNo"

    Structure SalesJournalNew
        Dim AuditID as integer
        Dim SaleNo As String
        Dim Name1 As String
        Dim TransDate As String
        Dim Written As Decimal
        Dim TaxCharged1 As Decimal
        Dim ArCashSls As Decimal
        Dim Control As Decimal
        Dim UndSls As Decimal
        Dim DelSls As Decimal
        Dim TaxRec1 As Decimal
        Dim TaxCode as integer
        Dim Salesman As String
        Dim NonTaxable As Decimal
        Dim Cashier As String
        Dim Terminal As String
    End Structure

    Public Function AddNewAuditRecord(
  ByVal SaleNo As String, ByVal Name As String, ByVal TransDate As Date,
  ByVal Written As Decimal, ByVal TaxCharged1 As Decimal, ByVal ArCashSls As Decimal,
  ByVal Control As Decimal, ByVal UndSls As Decimal, ByVal DelSls As Decimal,
  ByVal TaxRec1 As Decimal, ByVal TaxCode as integer,
  ByVal Salesman As String, Optional ByVal NonTaxable As Decimal = 0,
  Optional ByVal Cashier As String = "", Optional ByVal mTerminal As String = "") As Boolean
        '::::AddNewAuditRecord
        ':::SUMMARY
        ': Used to Add New Audit Record.
        ':::DESCRIPTION
        ': This function is used to update data in New Audit Record of SalesJournalNew Recordset.
        ':::PARAMETERS
        ':::RETURN
        ': Boolean - Returns True.
        Dim NewAudit As SalesJournalNew
        NewAudit.SaleNo = Trim(SaleNo)
        NewAudit.Name1 = Trim(Name)
        If Not IsDate(TransDate) Then
            NewAudit.TransDate = DateFormat(DateValue(Today))
        Else
            NewAudit.TransDate = DateFormat(DateValue(TransDate))
        End If
        NewAudit.Written = Written
        NewAudit.TaxCharged1 = TaxCharged1
        NewAudit.ArCashSls = ArCashSls
        NewAudit.Control = Control
        NewAudit.UndSls = UndSls
        NewAudit.DelSls = DelSls
        NewAudit.TaxRec1 = TaxRec1
        NewAudit.TaxCode = IIf(TaxCode = 0, 1, TaxCode)
        NewAudit.Salesman = Trim(Salesman)
        If NewAudit.Salesman = "" Then NewAudit.Salesman = "99" ' BFH20060519 - can't be zero length..?
        NewAudit.NonTaxable = NonTaxable
        NewAudit.Cashier = IIf(Cashier = "", GetCashierName, Cashier)
        NewAudit.Terminal = IIf(mTerminal = "", Terminal, mTerminal)

        SalesJournal_AddRecordNew(NewAudit)
        AddNewAuditRecord = True
    End Function
    Public Sub SalesJournal_AddRecordNew(ByRef Data As SalesJournalNew)
        '::::SalesJournal_AddRecordNew
        ':::SUMMARY
        ': Used to Add New Record to SalesJournalNew recordset.
        ':::DESCRIPTION
        ': This function is used to update inforrmation about salesman and cashier in SalesJournalNew Recordset after getting it,by using Index.
        ':::PARAMETERS
        ': - Data - Indicates the Data present in SalesJournal Table.
        Dim RS As ADODB.Recordset
        If Data.Salesman = "" Then Data.Salesman = "99" ' BFH20060519
        If Data.Cashier = "" Then Data.Cashier = GetCashierName
        RS = getRecordsetByTableLabelIndex(SalesJournal_TABLE, SalesJournal_INDEX, "-1", True, GetDatabaseAtLocation)
        SalesJournalNew_RecordSet_Get(Data, RS)
        SetRecordsetByTableLabelIndex(RS, SalesJournal_TABLE, SalesJournal_INDEX, "-1", GetDatabaseAtLocation)
        RS = Nothing
    End Sub
    Public Sub SalesJournalNew_RecordSet_Get(ByRef Sj As SalesJournalNew, ByRef RS As ADODB.Recordset)
        '::::SalesJournalNew_RecordSet_Get
        ':::SUMMARY
        ': Used to get the data from the SalesJournalNew Recordset.
        ':::DESCRIPTION
        ': This function is used to return the SalesJournalNew Recordset data.
        ': - Sj - Indicates the Sales Journal Recordset.
        ': - RS - Indicates the ADODB.Recordset.
        On Error Resume Next
        ' Don't set AuditID!
        RS("SaleNo").Value = Trim(Sj.SaleNo)
        RS("Name1").Value = Trim(Sj.Name1)
        RS("TransDate").Value = Trim(Sj.TransDate)
        RS("Written").Value = Sj.Written
        RS("TaxCharged1").Value = Sj.TaxCharged1
        RS("ArCashSls").Value = Sj.ArCashSls
        RS("Controll").Value = Sj.Control
        RS("UndSls").Value = Sj.UndSls
        RS("DelSls").Value = Sj.DelSls
        RS("TaxRec1").Value = Sj.TaxRec1
        RS("TaxCode").Value = IIf(Val(Sj.TaxCode) = 0, 1, Val(Sj.TaxCode))
        RS("Salesman").Value = Trim(Sj.Salesman)
        RS("NonTaxable").Value = Sj.NonTaxable
        RS("Cashier").Value = Sj.Cashier
        RS("Terminal").Value = Sj.Terminal
    End Sub

End Module

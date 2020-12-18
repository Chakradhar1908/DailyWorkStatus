Public Class clsServiceItemParts
    Public ServiceItemPartsId As Integer
    Public ServiceItemsId As Integer
    Public ServiceOrderNumber As Integer
    Public MarginNo As Integer
    Public Style As String
    Public Desc As String
    Public Vendor As String
    Public VendorNo As String
    Public InvoiceNo As String
    Public InvoiceDate As Date
    Public PartsReceived As Boolean
    Public FactoryPartsOrderNo As String
    Public Notes As String

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess
    Private Const TABLE_NAME = "ServiceItemParts"
    Private Const TABLE_INDEX = "ServiceItemPartsID"

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseAtLocation()
        mDataAccess.Table = TABLE_NAME
        mDataAccess.Index = TABLE_INDEX
    End Sub

    Public Function Save() As Boolean
        On Error GoTo NoSave
        ' This instructs the class (in one simple call) to save its data members to the database.
        If DataAccess.Record_Count = 0 Then
            ' Record not found.  This means we're adding a new one.
            DataAccess.Records_Add()
            cDataAccess_SetRecordSet(DataAccess.RS)
        End If
        ' Then load our data into the recordset.
        DataAccess.Record_Update()
        cDataAccess_SetRecordSet(DataAccess.RS)
        ' And finally, tell the class to save the recordset.
        DataAccess.Records_Update()
        Exit Function

NoSave:
        Err.Clear()
        Save = False
    End Function

    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose()
    End Sub

    Public Function cDataAccess_SuperClass() As CDataAccess
        cDataAccess_SuperClass = mDataAccess
    End Function

    Public Sub cDataAccess_SetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        RS("ServiceItemsId").Value = ServiceItemsId
        RS("ServiceOrderNo").Value = ServiceOrderNumber
        RS("MarginNo").Value = IfNullThenZero(MarginNo)
        RS("Style").Value = IfNullThenNilString(Trim(Style))
        RS("Desc").Value = IfNullThenNilString(Trim(Desc))
        RS("Vendor").Value = IfNullThenNilString(Trim(Vendor))
        RS("VendorNo").Value = IfNullThenNilString(Trim(VendorNo))
        RS("InvoiceNo").Value = IfNullThenNilString(Trim(InvoiceNo))
        RS("InvoiceDate").Value = InvoiceDate
        RS("PartsReceived").Value = PartsReceived
        RS("FactoryPartsOrderNo").Value = IfNullThenNilString(Trim(FactoryPartsOrderNo))
        RS("Notes").Value = IfNullThenNilString(Trim(Notes))
    End Sub

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        ServiceItemPartsId = RS("ServiceItemsPartsId").Value
        ServiceItemsId = RS("ServiceItemsId").Value
        ServiceOrderNumber = RS("ServiceOrderNumber").Value
        MarginNo = IfNullThenZero(RS("MarginNo").Value)
        Style = IfNullThenNilString(Trim(RS("Style").Value))
        Desc = IfNullThenNilString(Trim(RS("Desc").Value))
        Vendor = IfNullThenNilString(Trim(RS("Vendor").Value))
        VendorNo = IfNullThenNilString(Trim(RS("VendorNo").Value))
        InvoiceNo = IfNullThenNilString(Trim(RS("InvoiceNo").Value))
        InvoiceDate = RS("InvoiceDate").Value
        PartsReceived = RS("PartsReceived").Value
        FactoryPartsOrderNo = IfNullThenNilString(Trim(RS("FactoryPartsOrderNo").Value))
        Notes = IfNullThenNilString(Trim(RS("Notes").Value))
    End Sub

    Public Function Load(ByVal KeyVal As String, Optional ByVal KeyName As String = "") As Boolean
        ' Checks the database for a matching LeaseNo.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        ' Search for the Style
        If KeyName = "" Then
            DataAccess.Records_OpenIndexAt(KeyVal)
        ElseIf Left(KeyName, 1) = "#" Then
            ' This allows searching by AutoNumber - specialized to query by number
            ' since Access is exceptionally picky about quotation marks.
            DataAccess.Records_OpenFieldIndexAtNumber(Mid(KeyName, 2), KeyVal)
        ElseIf Left(KeyName, 1) = "@" Then
            DataAccess.Records_OpenFieldIndexAtDate(Mid(KeyName, 2), KeyVal)
        Else
            DataAccess.Records_OpenFieldIndexAt(KeyName, KeyVal)
        End If

        ' Move to the first record if we can, and return success.
        If DataAccess.Records_Available Then
            cDataAccess_GetRecordSet(DataAccess.RS)
            Load = True
        End If
    End Function

    Private Sub mDataAccess_RecordUpdated()
        ServiceItemPartsId = mDataAccess.Value("ServiceItemPartsID")
    End Sub

End Class

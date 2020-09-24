Public Class clsServiceOrder
    Public ServiceOrderNo As Integer
    Public LastName As String
    Public Telphone As String
    Public MailIndex As Integer
    Public SaleNo As String
    Public ServiceOnDate As String
    Public DateOfClaim As Date
    Public Status As String
    Public QuickCheck As String
    Public Item As String
    Public Complaint As String
    Public StoreAction As String
    Public SOType As String
    Public Mfg As String
    Public InvoiceNo As String
    Public Detail As String
    Public StopStart As String
    Public StopEnd As String

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME = "Service"
    Private Const TABLE_INDEX = "ServiceOrderID"

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
        If DataAccess.Records_Available Then Load = True
    End Function

    Public Function Save() As Boolean
        On Error GoTo NoSave
        ' This instructs the class (in one simple call) to save its data members to the database.
        If DataAccess.Record_Count = 0 Then
            ' Record not found.  This means we're adding a new one.
            DataAccess.Records_Add()
        End If
        ' Then load our data into the recordset.
        DataAccess.Record_Update()
        ' And finally, tell the class to save the recordset.
        DataAccess.Records_Update()
        Exit Function

NoSave:
        Err.Clear()
        Save = False
    End Function

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        ServiceOrderNo = RS("ServiceOrderNo").Value
        LastName = IfNullThenNilString(Trim(RS("LastName").Value))
        Telphone = IfNullThenNilString(Trim(RS("Telphone").Value))
        MailIndex = RS("MailIndex").Value
        SaleNo = IfNullThenNilString(Trim(RS("SaleNo").Value))
        ServiceOnDate = IfNullThenNilString(Trim(RS("ServiceOnDate").Value))
        DateOfClaim = RS("DateOfClaim").Value
        Status = IfNullThenNilString(Trim(RS("Status").Value))
        QuickCheck = IfNullThenNilString(Trim(RS("QuickCheck").Value))
        Item = IfNullThenNilString(Trim(RS("Item").Value))
        Complaint = IfNullThenNilString(Trim(RS("Complaint").Value))
        StoreAction = IfNullThenNilString(Trim(RS("StoreAction").Value))
        SOType = IfNullThenNilString(Trim(RS("Type").Value))
        Mfg = IfNullThenNilString(Trim(RS("Mfg").Value))
        InvoiceNo = IfNullThenNilString(Trim(RS("InvoiceNo").Value))
        Detail = IfNullThenNilString(Trim(RS("Detail").Value))

        If IsDate(IfNullThenNilString(RS("StopStart").Value)) Then
            StopStart = Trim(Format(TimeValue(RS("StopStart").Value), "h:mm ampm"))
        Else
            StopStart = ""
        End If
        If IsDate(IfNullThenNilString(RS("StopEnd").Value)) Then
            StopEnd = Trim(Format(TimeValue(RS("StopEnd").Value), "h:mm ampm"))
        Else
            StopEnd = ""
        End If
    End Sub
End Class

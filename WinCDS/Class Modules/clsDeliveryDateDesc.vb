Public Class clsDeliveryDateDesc
    ' Field declarations.
    Public DeliveryDate As Date
    Public DeliveryDesc As String

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME As String = "DeliveryDateDescriptions"
    Private Const TABLE_INDEX As String = "DeliveryDate"

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseAtLocation(IIf(StoreSettings.bOneCalendar, 1, 0))
        mDataAccess.Table = TABLE_NAME
        mDataAccess.Index = TABLE_INDEX
    End Sub

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Function cDataAccess_SuperClass() As CDataAccess
        cDataAccess_SuperClass = mDataAccess
    End Function

    Private Sub cDataAccess_SetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        RS("DeliveryDate").Value = DeliveryDate
        RS("DeliveryCaption").Value = IfNullThenNilString(Trim(DeliveryDesc))
    End Sub

    Private Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        DeliveryDate = RS("DeliveryDate").Value
        DeliveryDesc = IfNullThenNilString(Trim(RS("DeliveryCaption").Value))
    End Sub

    Public Function Save() As Boolean
        On Error GoTo NoSave
        ' This instructs the class (in one simple call) to save its data members to the database.
        If DataAccess.Record_Count = 0 Then
            ' Record not found.  This means we're adding a new one.
            DataAccess.Records_Add()
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
End Class

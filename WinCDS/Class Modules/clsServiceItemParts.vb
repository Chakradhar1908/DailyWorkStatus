Public Class clsServiceItemParts
    Public ServiceItemPartsId as integer
    Public ServiceItemsId as integer
    Public ServiceOrderNumber as integer
    Public MarginNo as integer
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

End Class

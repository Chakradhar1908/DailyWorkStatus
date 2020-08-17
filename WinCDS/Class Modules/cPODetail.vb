Public Class cPODetail
    Private Structure PODetailOld
        <VBFixedString(6)> Dim PoNo As String
        <VBFixedString(8)> Dim SaleNo As String
        <VBFixedString(10)> Dim PoDate As String
        <VBFixedString(24)> Dim Name As String
        <VBFixedString(16)> Dim Vendor As String
        <VBFixedString(5)> Dim Quantity As String
        <VBFixedString(16)> Dim Style As String
        <VBFixedString(47)> Dim Desc As String
        <VBFixedString(8)> Dim Cost As String
        <VBFixedString(1)> Dim Location As String
        <VBFixedString(1)> Dim SoldTo As String
        <VBFixedString(1)> Dim ShipTo As String
        <VBFixedString(1)> Dim Note1 As String
        <VBFixedString(1)> Dim Note2 As String
        <VBFixedString(1)> Dim Note3 As String
        <VBFixedString(1)> Dim Note4 As String
        <VBFixedString(39)> Dim PoNotes As String
        <VBFixedString(12)> Dim AckInv As String
        <VBFixedString(1)> Dim PrintPo As String
        <VBFixedString(1)> Dim Posted As String
        <VBFixedString(1)> Dim wCost As String
        <VBFixedString(6)> Dim RN As String
        <VBFixedString(7)> Dim Detail As String
        <VBFixedString(7)> Dim MarginLine As String
    End Structure

    Public Poid as integer
    Public PoNo as integer
    Public SaleNo As String
    Public PoDate As String
    Public Name As String
    Public Vendor As String
    Public InitialQuantity As Single
    Public Quantity As Single
    Public Style As String
    Public Desc As String
    Public Cost As Decimal
    Public Location as integer
    Public SoldTo as integer
    Public ShipTo as integer
    Public Note1 as integer
    Public Note2 as integer
    Public Note3 as integer
    Public Note4 as integer
    Public PoNotes As String
    Public AckInv As String
    Public PrintPo As String
    Public Posted As String
    Public wCost as integer
    Public RN as integer
    Public Detail as integer
    Public MarginLine as integer
    Public ShiptoName As String
    Public ShipToAddress As String
    Public ShipToCity As String
    Public ShipToTele As String
    Public SpecialNote As String
    Public DueDate As String
    Public PoRecDate As String
    Public Blank As String

    Public ForceUpdate As Boolean
    Public CancelledUpdate As Boolean

    Private mDataConvert As cDataConvert
    'Implements cDataConvert

    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const TABLE_NAME = "PO"
    Private Const TABLE_INDEX = "PoNo"
    Private Const FILE_Name = "PO.exe"
    Private Const FILE_RecordSize = 221 ' 221
    Private Const FILE_Index = 10

    Public Function Save() As Boolean
        Save = True
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

    Public Sub New()
        CDataConvert_Init
        CDataAccess_Init
    End Sub

    Public Sub CDataConvert_Init()
        mDataConvert = New cDataConvert
        With mDataConvert '@NO-LINT-WITH
            .SubClass = Me.mDataConvert
            .DataBase = GetDatabaseInventory()
            .Table = TABLE_NAME
            .Index = TABLE_INDEX
        End With
        '  ConvertSkip = True
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseInventory()
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
        Else
            DataAccess.Records_OpenFieldIndexAt(KeyName, KeyVal)
        End If

        ' Move to the first record if we can, and return success.
        If DataAccess.Records_Available Then Load = True
    End Function

End Class

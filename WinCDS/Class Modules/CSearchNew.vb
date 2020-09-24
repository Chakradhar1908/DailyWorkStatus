Public Class CSearchNew
    Public Style As String
    Public Dept As String
    Public Mfg As String
    Public RN As String

    Private mDataConvert As cDataConvert
    'Implements cDataConvert

    Private mDataAccess As CDataAccess
    'Implements CDataAccess

    Private Const FILE_Name = "Search2" & ".exe"
    Private Const FILE_RecordSize = 25
    Private Const FILE_Index = 1

    Private Const TABLE_NAME = "Search"
    Private Const TABLE_INDEX = "Style"

    Public Function Load(ByVal KeyVal As String, Optional ByVal KeyName As String = "") As Boolean
        ' Checks the database for a matching Style ID.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        If KeyName = "" Then
            DataAccess.Records_OpenIndexAt(KeyVal)
        ElseIf Left(KeyName, 1) = "#" Then
            ' This allows searching by AutoNumber - specialized to query by number
            ' since Access is exceptionally picky about quotation marks.
            DataAccess.Records_OpenFieldIndexAtNumber(Mid(KeyName, 2), KeyVal)
        Else
            DataAccess.Records_OpenFieldIndexAt(KeyName, KeyVal)
        End If

        If DataAccess.Records_Available Then
            cDataAccess_GetRecordSet(DataAccess.RS)
            Load = True
        End If
    End Function

    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose()
    End Sub

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
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        mDataAccess.SubClass = Me.mDataAccess
        mDataAccess.DataBase = GetDatabaseInventory()
        mDataAccess.Table = TABLE_NAME
        mDataAccess.Index = TABLE_INDEX
    End Sub

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        Style = IfNullThenNilString(RS("Style").Value)
        Dept = IfNullThenNilString(RS("Dept").Value)
        Mfg = IfNullThenNilString(RS("Mfg").Value)
        RN = IfNullThenNilString(RS("Rn").Value)
    End Sub
End Class

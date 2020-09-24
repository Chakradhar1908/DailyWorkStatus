Public Class CInventoryDetail
    Private WithEvents mDataAccess As CDataAccess
    Public Style As String
    Public Lease1 As String
    Public Name As String
    Public Misc As String
    Public DDate1 As String
    Public Trans As String
    Public AmtS1 As Single
    Public Ns1 As Single
    Public SO1 As Single
    Public LAW As Single
    Public ItemCost As Decimal  ' BFH20060124
    'Private Loc(1 To Setup_MaxStores_DB)
    Private Loc(0 To Setup_MaxStores_DB - 1)
    Public DetailID As Integer  ' MJK Autonumber
    Public MarginRn As Integer
    Public InvRn As Integer
    Public Store As Integer
    Private mDataConvert As cDataConvert
    'Implements cDataConvert
    Private Const FILE_Name As String = "DETAIL" & ".exe"
    Private Const FILE_RecordSize As Integer = 120
    Private Const FILE_Index As Integer = 6
    Private Const TABLE_NAME As String = "Detail"
    Private Const TABLE_INDEX As String = "SaleNo"
    Public Notes As String      ' BFH20080212

    Public Sub New()
        CDataConvert_Init()
        CDataAccess_Init()
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

    'Implements CDataAccess
    Public Function Load(ByVal KeyVal As String, Optional ByRef KeyName As String = "") As Boolean
        ' Checks the database for a matching LeaseNo.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        ' Search for the Style
        Load = False
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
        mDataAccess_RecordUpdated()
        Exit Function

NoSave:
        Err.Clear()
        Save = False
    End Function

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose()
    End Sub

    Public Sub SetLocationQuantity(ByVal Location As Integer, ByVal Quantity As Double)
        If Not LocationValid(Location) Then Exit Sub
        Loc(Location - 1) = Quantity
    End Sub

    Private Function LocationValid(ByVal Loc As Integer) As Boolean
        If Loc <= 0 Or Loc > Setup_MaxStores_DB - 1 Then
            LocationValid = False
            Exit Function
        End If
        LocationValid = True
    End Function

    Public Sub cDataAccess_SetRecordSet(RS As ADODB.Recordset)
        Dim I As Integer
        On Error Resume Next
        ' mDataAccess.Count + 1
        'rs("DetailID") = DetailID  ' Don't set the autonumber field.  This causes problems.
        RS("Style").Value = Left(IfNullThenNilString(Trim(Style)), Setup_2Data_StyleMaxLen)
        RS("SaleNo").Value = IfNullThenNilString(Trim(Lease1))
        RS("Name").Value = IfNullThenNilString(Trim(Name))
        RS("Misc").Value = IfNullThenNilString(Trim(Misc))
        RS("Ddate1").Value = DDate1
        RS("Trans").Value = IfNullThenNilString(Trans)
        RS("AmtSold").Value = AmtS1
        RS("NewStock").Value = Ns1
        RS("SpecOrd").Value = SO1
        RS("LAW").Value = LAW
        For I = 1 To Setup_MaxStores_DB
            RS("Loc" & I).Value = GetLocationQuantity(I)
        Next
        RS("Store").Value = Store
        RS("InvRn").Value = InvRn
        RS("MarginRn").Value = MarginRn
        RS("ItemCost").Value = ItemCost
        RS("Notes").Value = Notes
    End Sub

    Public Function GetLocationQuantity(ByVal Location As Integer) As Double
        If Not LocationValid(Location) Then Exit Function
        GetLocationQuantity = Loc(Location)
    End Function

    Private Sub mDataAccess_RecordUpdated()
        DetailID = mDataAccess.Value("DetailID")
    End Sub

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        Dim I As Integer
        On Error Resume Next
        DetailID = RS("DetailID").Value
        Style = IfNullThenNilString(RS("Style").Value)
        Lease1 = IfNullThenNilString(RS("SaleNo").Value)
        Name = IfNullThenNilString(RS("Name").Value)
        Misc = IfNullThenNilString(Trim(RS("Misc").Value))
        DDate1 = RS("Ddate1").Value
        If IsNothing(RS("Ddate1").Value) Then DDate1 = ""
        Trans = IfNullThenNilString(RS("Trans").Value)
        AmtS1 = RS("AmtSold").Value
        Ns1 = RS("NewStock").Value
        SO1 = RS("SpecOrd").Value
        LAW = RS("LAW").Value

        For I = 1 To Setup_MaxStores_DB
            SetLocationQuantity(I, RS("Loc" & I).Value)
        Next

        Store = RS("Store").Value
        InvRn = IfNullThenZero(RS("InvRn").Value)
        MarginRn = IfNullThenZero(RS("MarginRn").Value)
        ItemCost = IfNullThenZeroCurrency(RS("ItemCost").Value)
        Notes = IfNullThenNilString(RS("Notes").Value)
    End Sub

    Public Function GetFirstLocationWithPositiveQuantity() As Integer
        Dim I As Integer
        For I = 1 To Setup_MaxStores
            If GetLocationQuantity(I) Then GetFirstLocationWithPositiveQuantity = I : Exit Function
        Next
    End Function
End Class

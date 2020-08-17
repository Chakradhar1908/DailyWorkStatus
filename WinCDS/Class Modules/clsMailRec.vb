Public Class clsMailRec
    Public Index as integer
    Public Last As String
    Public First As String
    Public Address As String
    Public AddAddress As String
    Public City As String
    Public Zip As String
    Public Tele As String
    Public Tele2 As String
    Public PhoneLabel1 As String
    Public PhoneLabel2 As String
    Public Special As String
    Public MailType As String
    Public CustType As String
    Public TypeIndex as integer
    Public Blank As String
    Public Email As String
    Public Business As Boolean
    Public CreditCard As String
    Public ExpDate As String
    Public TaxZone as integer
    Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess
    Private Const TABLE_NAME As String = "Mail"
    Private Const TABLE_INDEX = "Index"

    Public Sub New()
        CDataAccess_Init
    End Sub

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

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
        Load = False
        If KeyName = "" Then
            DataAccess.Records_OpenFieldIndexAtNumber(TABLE_INDEX, KeyVal)
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

    Public Function ShipTo(Optional ByVal StoreNo As Integer = 0) As clsMailShipTo
        '  If StoreNo = 0 Then StoreNo = StoresSld
        '  modMail.Mail2_GetAtIndex CStr(Index), ShipTo, StoreNo
        ShipTo = New clsMailShipTo
        ShipTo.DataAccess.DataBase = DataAccess.DataBase ' default to same database...
        If StoreNo <> 0 Then ShipTo.DataAccess.DataBase = GetDatabaseAtLocation(StoreNo)
        ShipTo.Load(Index)
    End Function

    Private Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        Index = RS("Index").Value
        Last = IfNullThenNilString(Trim(RS("Last").Value))
        First = IfNullThenNilString(Trim(RS("First").Value))
        Address = IfNullThenNilString(Trim(RS("Address").Value))
        AddAddress = IfNullThenNilString(Trim(RS("AddAddress").Value))
        City = IfNullThenNilString(Trim(RS("City").Value))
        Zip = IfNullThenNilString(Trim(RS("Zip").Value))
        Tele = CleanAni(IfNullThenNilString(RS("Tele").Value))
        Tele2 = CleanAni(IfNullThenNilString(RS("Tele2").Value))
        PhoneLabel1 = IfNullThenNilString(Trim(RS("PhoneLabel1").Value))
        PhoneLabel2 = IfNullThenNilString(Trim(RS("PhoneLabel2").Value))
        Special = IfNullThenNilString(Trim(RS("Special").Value))
        MailType = IfNullThenNilString(Trim(RS("Type").Value))
        TypeIndex = IfNullThenZero(RS("TypeIndex").Value)
        CustType = IfNullThenNilString(Trim(RS("CustType").Value))
        Email = IfNullThenNilString(Trim(RS("Email").Value))
        Blank = IfNullThenNilString(Trim(RS("Blank").Value))
        Business = IfNullThenBoolean(RS("Business").Value)
        CreditCard = IfNullThenNilString(Trim(RS("CreditCard").Value))
        ExpDate = IfNullThenNilString(Trim(RS("ExpDate").Value))
        TaxZone = IfNullThenZero(RS("TaxZone").Value)

        If PhoneLabel1 = "" Then PhoneLabel1 = "Telephone"
        If PhoneLabel2 = "" Then PhoneLabel2 = "Telephone 2"
    End Sub

    Private Sub cDataAccess_SetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        If Index <= 0 Then
            ' Poor man's autonumber.. MailIndex is too tied in to everything else to replace right now.
            Index = MailTableRecordMax("index") + 1
        Else
            RS("Index").Value = Index
        End If

        RS("Last").Value = IfNullThenNilString(Trim(Last))
        RS("First").Value = IfNullThenNilString(Trim(First))
        RS("Address").Value = IfNullThenNilString(Trim(Address))
        RS("AddAddress").Value = IfNullThenNilString(Trim(AddAddress))
        RS("City").Value = IfNullThenNilString(Trim(City))
        RS("Zip").Value = IfNullThenNilString(Trim(Zip))
        RS("Tele").Value = CleanAni(IfNullThenNilString(Tele))
        RS("Tele2").Value = CleanAni(IfNullThenNilString(Tele2))
        RS("PhoneLabel1").Value = IfNullThenNilString(Trim(PhoneLabel1))
        RS("PhoneLabel2").Value = IfNullThenNilString(Trim(PhoneLabel2))
        RS("Special").Value = IfNullThenNilString(Trim(Special))
        RS("Type").Value = IfNullThenNilString(Trim(MailType))
        RS("TypeIndex").Value = TypeIndex
        RS("CustType").Value = IfNullThenNilString(Trim(CustType))
        RS("Email").Value = IfNullThenNilString(Trim(Email))
        RS("Blank").Value = IfNullThenNilString(Trim(Blank))
        RS("Business").Value = Business
        RS("CreditCard").Value = IfNullThenNilString(Trim(CreditCard))
        RS("ExpDate").Value = IfNullThenNilString(Trim(ExpDate))
        RS("TaxZone").Value = TaxZone
    End Sub

End Class

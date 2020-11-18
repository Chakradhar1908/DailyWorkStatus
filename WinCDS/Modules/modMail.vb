Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module modMail
    Dim printer As New Printer
    Public MailRec As Integer
    ' Use this for Access Version
    Public Structure MailNew
        Dim Index As String
        Dim Last As String
        Dim First As String
        Dim Address As String
        Dim AddAddress As String
        Dim City As String
        Dim Zip As String
        Dim Tele As String
        Dim Tele2 As String
        Dim PhoneLabel1 As String
        Dim PhoneLabel2 As String
        Dim Special As String
        Dim Type As String
        Dim CustType As String
        Dim Blank As String
        Dim Email As String
        Dim Business As Boolean
        Dim CreditCard As String
        Dim ExpDate As String
        Dim TaxZone As Integer
    End Structure

    Public Structure MailNew2
        Dim Index As String
        Dim ShipToLast As String
        Dim ShipToFirst As String
        Dim Address2 As String
        Dim City2 As String
        Dim Zip2 As String
        Dim Tele3 As String
        Dim PhoneLabel3 As String
        Dim Blank As String
    End Structure

    Public Function LoadCashAndCarryMail() As clsMailRec
        '::::LoadCashAndCarryMail
        ':::SUMMARY
        ': Loads Cash and Carry from mail table.
        ':::DESCRIPTION
        ': This function is called when we want to load cash and carry  from mail table.
        ':::PARAMETERS
        ':::RETURN
        ': clsMailRec

        LoadCashAndCarryMail = New clsMailRec
        LoadCashAndCarryMail.Last = "CASH & CARRY"
        LoadCashAndCarryMail.First = ""
        LoadCashAndCarryMail.Address = ""
        LoadCashAndCarryMail.AddAddress = ""
        LoadCashAndCarryMail.City = ""
        LoadCashAndCarryMail.Zip = ""
        LoadCashAndCarryMail.Tele = ""
        LoadCashAndCarryMail.Tele2 = ""
        LoadCashAndCarryMail.PhoneLabel1 = ""
        LoadCashAndCarryMail.PhoneLabel2 = ""
        LoadCashAndCarryMail.Index = 0
        LoadCashAndCarryMail.Special = ""
        '    .Type = ""
        LoadCashAndCarryMail.CustType = ""
        LoadCashAndCarryMail.CreditCard = ""
        LoadCashAndCarryMail.ExpDate = ""
        LoadCashAndCarryMail.TaxZone = 0
    End Function

    Public Sub PrintDYMOMailingLabel(ByVal LastName As String, Optional ByVal FirstName As String = "", Optional ByVal Address1 As String = "", Optional ByVal Address2 As String = "", Optional ByVal City As String = "", Optional ByVal Zip As String = "", Optional ByVal LabelType As Integer = 30323)
        '::::PrintDYMOMailingLabel
        ':::SUMMARY
        ': Used to print  required information using Dymo printer.
        ':
        ':::DESCRIPTION
        ': This function is called , When we want to print any required information.
        ': We can select Dymo printer using available label options from DYMO Label printer available in Components under Store Setup under File menu.
        ':
        ':::PARAMETERS
        ': - LastName - Indicates the lastname given by user.
        ': - FirstName - Indicates the firstname given by user.
        ': - Address1 - Indicates the Address given by user.
        ': - Address2 - Indicates the Address given by user.
        ': - City - Indicates the nae of city given by user.
        ': - Zip - Indicates the Zip code given by user.
        ': - LabelType -Indicates the tyoe of label and is 30323.
        ':::RETURN

        Dim OriginalPrint As String
        OriginalPrint = printer.DeviceName

        If Not SetDymoPrinter(LabelType) Then
            MessageBox.Show("Printing address mailing labels requires a DYMO printer." & vbCrLf & "A DYMO label printer could not be detected on your computer.", "Dymo Printer Required!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        printer.Orientation = vbPRORLandscape
        printer.FontSize = 12
        printer.CurrentX = 0
        printer.CurrentY = 0
        If LabelType = 30232 Then printer.Print()
        printer.Print(Trim(FirstName), "  ", LastName)
        printer.Print(Address1)
        If Trim(Address2) <> "" Then printer.Print(Address2)
        printer.Print(Trim(City), "  ", Zip)
        printer.EndDoc()

        printer.Orientation = vbPRORPortrait
        SetPrinter(OriginalPrint)
    End Sub

    Public Function MailTableRecordMax(Optional ByVal Field As String = "Index") As Integer
        '::::MailTableRecordMax
        ':::SUMMARY
        ': Gets Maximum index value.
        ':::DESCRIPTION
        ': By calling this function , we can get Maximum index value from mail table .
        ':::PARAMETERS
        ': - Field - Indicates the Index as a string.
        ':::RETURN
        ': Long - Returns the Maximum index long value.
        ':::SEE ALSO
        ':MailTableRecordCount
        Dim RS As ADODB.Recordset
        Dim SQL As String
        SQL = "SELECT Max(CLng(" & Field & ")) AS GetMax FROM Mail;"

        Dim cDBa As CDbAccessGeneral
        cDBa = DbAccessGeneral(GetDatabaseAtLocation, SQL)
        RS = cDBa.getRecordset(Always:=False)
        MailTableRecordMax = IfNullThenZero(RS("GetMax").Value)
        cDBa.dbClose()
    End Function

    Public Sub SetMailRecordsetFromMailNew(ByRef RS As ADODB.Recordset, ByRef tMailNew As MailNew)
        '::::SetMailRecordsetFromMailNew
        ':::SUMMARY
        ': Sets the MailRecorder from the MailNew data structure.
        ':::DESCRIPTION
        ': This function is useful to  set the MailRecordset from MailNew data structure.
        ':::PARAMETERS
        ': - RS - Indicates the Recordset.
        ': -tMailNew - Indicates the data structure MailNew.
        ':::RETURN
        RS("Index").Value = Trim(tMailNew.Index)
        RS("Last").Value = Trim(tMailNew.Last)
        RS("First").Value = Trim(tMailNew.First)
        RS("address").Value = Trim(tMailNew.Address)
        RS("City").Value = Trim(tMailNew.City)
        RS("zip").Value = Trim(tMailNew.Zip)
        RS("Tele").Value = CleanAni(Trim(tMailNew.Tele))
        RS("Tele2").Value = CleanAni(Trim(tMailNew.Tele2))
        RS("PhoneLabel1").Value = Trim(tMailNew.PhoneLabel1)
        RS("PhoneLabel2").Value = Trim(tMailNew.PhoneLabel2)
        RS("Special").Value = Trim(tMailNew.Special)
        RS("Type").Value = Trim(tMailNew.Type)
        '    If .CustType = "7" Then .CustType = "1"  'change old "no mail"
        If tMailNew.CustType = "-" Then tMailNew.CustType = "0"
        RS("CustType").Value = Trim(tMailNew.CustType)
        RS("Blank").Value = Trim(tMailNew.Blank)

        ' Use this for Access Version
        RS("Addaddress").Value = Trim(tMailNew.AddAddress)
        RS("Email").Value = Trim(tMailNew.Email)
        RS("Business").Value = tMailNew.Business

        RS("TaxZone").Value = tMailNew.TaxZone
        RS("CreditCard").Value = Val(tMailNew.CreditCard)
        RS("ExpDate").Value = tMailNew.ExpDate
    End Sub

    Public Sub CopyMailRecordsetToMailNew2(ByRef RS As ADODB.Recordset, ByRef tMailNew2 As MailNew2)
        '::::CopyMailRecordsetToMailNew2
        ':::SUMMARY
        ': Copies the MailRecordset to the MailNew2 data structure.
        ':
        ':::DESCRIPTION
        ': This function is used  to copy the MailRecordset to date structure MailNew2
        ': Useful to save ShipTo.
        ':
        ':::PARAMETERS
        ': - Recordset - Selection of records from table.
        ': - tMailNew2 - Indicates the MailNew2 data structure.

        Dim Loaded As Boolean
        If Not (RS Is Nothing) Then
            If Not RS.EOF Then
                Loaded = True
                tMailNew2.Index = RS.Fields("Index").Value
                tMailNew2.Index = IfNullThenNilString(RS("Index").Value)
                tMailNew2.Address2 = IfNullThenNilString(RS("Address").Value)
                tMailNew2.City2 = IfNullThenNilString(RS("City").Value)
                tMailNew2.Zip2 = IfNullThenNilString(RS("Zip").Value)
                tMailNew2.Tele3 = CleanAni(IfNullThenNilString(RS("Tele").Value))
                tMailNew2.PhoneLabel3 = Trim(IfNullThenNilString(RS("PhoneLabel3").Value))

                ' Use this for Access Version
                tMailNew2.ShipToLast = IfNullThenNilString(RS("Last").Value)
                tMailNew2.ShipToFirst = IfNullThenNilString(RS("First").Value)
                tMailNew2.Blank = IfNullThenNilString(RS("Blank").Value)
            End If
        End If

        If Not Loaded Then
            tMailNew2.Index = ""
            tMailNew2.Address2 = ""
            tMailNew2.City2 = ""
            tMailNew2.Zip2 = ""
            tMailNew2.Tele3 = ""
            tMailNew2.PhoneLabel3 = ""

            ' Use this for Access Version
            tMailNew2.ShipToLast = ""
            tMailNew2.ShipToFirst = ""
            tMailNew2.Blank = ""
        End If
    End Sub

    Public Sub SetMailRecordsetFromMailNew2(ByRef RS As ADODB.Recordset, ByRef tMailNew2 As MailNew2)
        '::::SetMailRecordsetFromMailNew2
        ':::SUMMARY
        ': Set the MailRecordset from the MailNew2 data structure.
        ':::DESCRIPTION
        ': By calling this function, we can set the mailrecordset from the MailNew2 data structure.
        ':::PARAMETERES
        ': - RS -Indicates Recordset.
        ': - tMailNew2 - Indicates the data structure.

        RS("Index").Value = Trim(tMailNew2.Index)
        RS("Address").Value = Trim(tMailNew2.Address2)
        RS("City").Value = Trim(tMailNew2.City2)
        RS("Zip").Value = Trim(tMailNew2.Zip2)
        RS("Tele").Value = Trim(tMailNew2.Tele3)
        RS("PhoneLabel3").Value = Trim(tMailNew2.PhoneLabel3)

        ' Use this for Access Version
        RS("Last").Value = Trim(tMailNew2.ShipToLast)
        RS("First").Value = Trim(tMailNew2.ShipToFirst)
        RS("Blank").Value = Trim(tMailNew2.Blank)
    End Sub

    Public Sub SetMailRecordset(ByRef RS As ADODB.Recordset, Optional ByVal StoreNum As Integer = 0)
        '::::SetMailRecordset
        ':::SUMMARY
        ': This function is used to save the recordset.
        ':::DESCRIPTION
        ': By calling this function , we can save the recordset and also update the database.
        ':::PARAMETERS
        ': - Recordset - Selection of records from table.
        ': - StoreNum - Indicates the storenumber.
        ':::RETURN
        Dim F As String
        F = GetDatabaseAtLocation(StoreNum)
        Dim cDBa As CDbAccessGeneral
        cDBa = DbAccessGeneral(SQL:=getMailByIndex("-1"), File:=F)
        cDBa.UpdateRecordSet(RS)   ' This must be called to update the database
        cDBa.dbClose()   ' used to close recordset
    End Sub

    Public Function GetMailRecordset(Optional ByVal Index As String = "-1", Optional ByVal StoreNum As Integer = 0) As ADODB.Recordset
        '::::GetMailRecordset
        ':::SUMMARY
        ': Gets mail recordset by index
        ':::DESCRIPTION
        ': We can get recordset from Mail table by Index
        ':::PARAMETERS
        ': - Index - Indicates the index value.
        ': - StoreNum - Indicates the store number.
        ':::RETURN
        ': Recordset - Returns the Mail Recordset.
        ':::SEE ALSO
        ':SetMailRecordset
        Dim F As String
        F = GetDatabaseAtLocation(StoreNum)
        Dim cDBa As CDbAccessGeneral
        cDBa = DbAccessGeneral(SQL:=getMailByIndex(Index), File:=F)
        GetMailRecordset = cDBa.getRecordset(False)   ' if 'SetNew:=False' by default
        cDBa.dbClose()
    End Function

    Public Function GetMailRecordsetByTele(Optional ByVal Index As String = "-1") As ADODB.Recordset
        '::::GetMailRecordsetByTele
        ':::SUMMARY
        ': Gets recordset from Mail table.
        ':::DESCRIPTION
        ':  by calling this function, We can get recordset from Mail table with tele value.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ':::RETURN
        ': ADODB.Recordset
        Dim cDBa As CDbAccessGeneral
        cDBa = DbAccessGeneral(SQL:=getMailByTele(Index), File:=GetDatabaseAtLocation)
        GetMailRecordsetByTele = cDBa.getRecordset(Always:=False)  ' if 'SetNew:=False' by default
        cDBa.dbClose()
    End Function

    Public Function getMailRecordsetByServiceCall(Optional ByVal Index As String = "-1") As ADODB.Recordset

        '::::getMailRecordsetByServiceCall
        ':::SUMMARY
        ': Gets recordset from Mail table.
        ':::DESCRIPTION
        ': By calling this function,we can display recordset from mail table using ServiceCall.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ':::RETURN
        ': ADODB.Recordset


        If Not IsNumeric(Index) Then Index = -1
        Dim cDBa As CDbAccessGeneral
        cDBa = DbAccessGeneral(SQL:=getMailByServiceCall(Index), File:=GetDatabaseAtLocation)
        getMailRecordsetByServiceCall = cDBa.getRecordset(Always:=False)  ' if 'SetNew:=False' by default
        cDBa.dbClose()
    End Function

    Public Sub CopyMailRecordsetToMailNew(ByRef RS As ADODB.Recordset, ByRef tMailNew As MailNew)
        '::::CopyMailRecordsetToMailNew
        ':::SUMMARY
        ': Copy the MailRecordset to the MailNew data structure.
        ':::DESCRIPTION
        ': This function is useful to avoid errors and to Copy the Mail Recordset to the MailNew data structure.
        ': PARAMETERS
        ': - RS - Indicates the Recordset.
        ': -tMailNew - Indicates the data structure MailNew.
        ':::RETURN


        On Error Resume Next
        tMailNew.Index = RS("Index").Value
        MailRec = tMailNew.Index
        tMailNew.Last = IfNullThenNilString(RS("Last").Value)
        tMailNew.First = IfNullThenNilString(RS("First").Value)
        tMailNew.Address = IfNullThenNilString(RS("address").Value)
        tMailNew.AddAddress = IfNullThenNilString(RS("addaddress").Value)
        tMailNew.City = IfNullThenNilString(RS("City").Value)
        tMailNew.Zip = IfNullThenNilString(RS("Zip").Value)
        tMailNew.Tele = IfNullThenNilString(DressAni(CleanAni(RS("Tele").Value)))
        tMailNew.Tele2 = IfNullThenNilString(DressAni(CleanAni(RS("Tele2").Value)))
        tMailNew.PhoneLabel1 = Trim(IfNullThenNilString(RS("PhoneLabel1").Value))
        tMailNew.PhoneLabel2 = Trim(IfNullThenNilString(RS("PhoneLabel2").Value))
        tMailNew.Special = IfNullThenNilString(RS("Special").Value)
        tMailNew.Type = IfNullThenNilString(RS("Type").Value)
        tMailNew.Zip = IfNullThenNilString(RS("Zip").Value)
        tMailNew.CustType = IfNullThenNilString(RS("CustType").Value)
        tMailNew.Blank = IfNullThenNilString(RS("Blank").Value)

        ' Use this for Access Version
        tMailNew.AddAddress = IfNullThenNilString(RS("AddAddress").Value)
        tMailNew.Email = IfNullThenNilString(RS("Email").Value)
        tMailNew.Business = RS("Business").Value

        tMailNew.TaxZone = RS("TaxZone").Value
        tMailNew.CreditCard = RS("CreditCard").Value
        tMailNew.ExpDate = RS("ExpDate").Value
    End Sub

    Public Function GetGrossSales(ByVal Index As Integer) As Decimal
        '::::GetGrossSales
        ':::SUMMARY
        ': Gets Gross sales.
        ':::DESCRIPTION
        ': By calling this function, we can get gross  sale from table GrossMargin.
        ': This function is useful to avoid errors and get GrossSales Currency.
        ':::PARAMETERS
        ': - Index - Indicates the Index number.
        ':::RETURN
        ': Currency - Returns the GrossSales currency.
        Dim S As String, R As ADODB.Recordset
        On Error Resume Next
        S = ""
        S = S & "SELECT Sum(SellPrice) as Tot FROM GrossMargin "
        S = S & "WHERE MailIndex=" & Index & " "
        S = S & "AND Left(Status, 1) <> 'x' "
        S = S & "AND NOT Style IN ('SUB','PAYMENT','TAX1','TAX2','--- Adj ---')"
        S = S & "AND left(Status, 2)<>'VD'"
        S = S & "AND Trim(Status)<>'VOID'"

        R = GetRecordsetBySQL(S, , GetDatabaseAtLocation)
        GetGrossSales = R("Tot").Value
        R.Close()
        R = Nothing
    End Function

    Public Sub Mail_GetAtIndex(ByVal Index As String, ByRef Mail As MailNew, Optional ByVal StoreNum As Integer = 0)
        '::::Mail_GetAtIndex
        ':::SUMMARY
        ': Gets Mail Recordset through Sql values.
        ':::DESCRIPTION
        ': By calling this function, we can get Mail Recordset through sql values and used to handle errors.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ': - Mail - Represents as MailNew, which indicates data structure.
        ': - StoreNum - Indicates the Store number.
        ':::RETURN
        ':::SEE ALSO
        ':Mail2_GetAtIndex
        Dim RS As ADODB.Recordset, F As String

        F = GetDatabaseAtLocation(StoreNum)
        RS = getRecordsetByTableLabelIndexNumber("Mail", "Index", Index, File:=F)
        If (RS.RecordCount <> 0) Then CopyMailRecordsetToMailNew(RS, Mail)
        'GetCust2 Index

        Exit Sub
HandleErr:
    End Sub

    Public Sub Mail2_GetAtIndex(ByVal Index As String, ByRef Mail2 As MailNew2, Optional ByVal StoreNum As Integer = 0)
        '::::Mail2_GetAtIndex
        ':::SUMMARY
        ': Gets Mail2 recordset through Sql values.
        ':::DESCRIPTION
        ': By calling this function, we can get Mail2 recordset through Sql values.
        ':This function is also used to handle errors.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ': - Mail2 - Represents as MailNew2,which indicates data structure.
        ': - StoreNum - Indicates the Store number.
        ':::RETURN
        On Error GoTo HandleErr
        Dim RS2 As ADODB.Recordset, F As String
        ' first things first...
        Mail2.Address2 = ""
        Mail2.Blank = ""
        Mail2.City2 = ""
        Mail2.Index = 0
        Mail2.ShipToFirst = ""
        Mail2.ShipToLast = ""
        Mail2.Tele3 = ""
        Mail2.PhoneLabel3 = ""
        Mail2.Zip2 = ""

        F = GetDatabaseAtLocation(StoreNum)
        RS2 = getRecordsetByTableLabelIndexNumber("MailShipTo", "Index", Index, File:=F)
        If (RS2.RecordCount <> 0) Then
            CopyMailRecordsetToMailNew2(RS2, Mail2)
            Exit Sub
        Else
            Mail2.Address2 = "" : Mail2.City2 = "" : Mail2.Zip2 = "" : Mail2.Tele3 = "" : Mail2.Blank = ""
        End If

        Exit Sub
HandleErr:
    End Sub

    Public Function getMailByIndex(ByVal Index As String) As String  'by index
        '::::getMailByIndex
        ':::SUMMARY
        ': Gets Mail with Index.
        ':::DESCRIPTION
        ': This function is useful to get Mail with Index.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ':::RETURN
        ': STRING - Returns the Mail as a string.
        ':::SEE ALSO
        ':getMailByServiceCall
        ':GetMailRecordsetByTele
        getMailByIndex = "SELECT Mail.* FROM Mail WHERE Mail.Index=" & Index
    End Function

    Private Function getMailByTele(ByVal I As String) As String
        getMailByTele = "SELECT mail.*" & " FROM mail" & " WHERE mail.Tele=""" & ProtectSQL(I) & """"
    End Function

    Public Function getMailByServiceCall(ByVal Index As String) As String
        '::::getMailByServiceCall
        ':::SUMMARY
        ': Gets Mail with Service Call.
        ':::DESCRIPTION
        ': This function is helpfule to get Mail with Service Call as a string.
        ':::PARAMETERS
        ': - Index - Returns the Mail as a string.
        ':::RETURN
        ': STRING - Returns the Mail as a string.
        getMailByServiceCall = "SELECT Mail.* FROM Mail INNER JOIN Service ON Mail.Index=Service.MailIndex WHERE Service.ServiceOrderNo=" & Index
    End Function

    Public Sub GetMailNew2ByIndex(ByVal Index As Integer, ByRef tMailNew2 As MailNew2, Optional ByVal StoreNo As Integer = 0)
        '::::GetMailNew2ByIndex
        ':::SUMMARY
        ': Gets  data structure MailNew2.
        ':::DESCRIPTION
        ': By calling this function, we can get MailNew2 with index number.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ': - tMailNew2 - Indicates the MailNew2 data structure.
        ': - StoreNo - Indicates the storenumber.
        ':::RETURN
        ':::SEE ALSO
        ':GetMailNewByIndex
        Dim RS As ADODB.Recordset
        If StoreNo = 0 Then StoreNo = StoresSld
        RS = GetRecordsetBySQL("SELECT * FROM MailShipTo WHERE Index=" & Index, , GetDatabaseAtLocation(StoreNo))
        CopyMailRecordsetToMailNew2(RS, tMailNew2)
        DisposeDA(RS)
    End Sub

    Public Function LoadMailRecord(ByVal Index As Integer, Optional ByVal StoreNo As Integer = 0) As clsMailRec
        '::::LoadMailRecord
        ':::SUMMARY
        ': Loads records from Mail table.
        ':::DESCRIPTION
        ': This function is called when we want to load records from Mail table by testing below scenarios.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ': - StoreNo - Indicates the Store number.
        ':::RETURN
        ': clsMailRec
        If StoreNo = 0 Then StoreNo = StoresSld

        If Index = 0 Then
            LoadMailRecord = LoadCashAndCarryMail()
            Exit Function
        End If

        Dim Ltmp As clsMailRec
        Ltmp = New clsMailRec
        Ltmp.DataAccess.DataBase = GetDatabaseAtLocation(StoreNo)
        If Not Ltmp.Load(CStr(Index), "#Index") Then
            Ltmp = Nothing
            LoadMailRecord = LoadCashAndCarryMail()
            Exit Function
        End If
        LoadMailRecord = Ltmp
    End Function

    Public Sub GetMailNewByIndex(ByVal Index As Integer, ByRef tMailNew As MailNew, Optional ByVal StoreNo As Integer = 0)
        '::::GetMailNewByIndex
        ':::SUMMARY
        ': Gets a data structure MailNew
        ':::DESCRIPTION
        ': Gets a data structure MailNew with Index.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ': - tMailNew - Indicates the MailNew data structure.
        ': - StoreNo - Indicates the storenumber.
        ':::RETURN

        Dim RS As ADODB.Recordset
        If StoreNo = 0 Then StoreNo = StoresSld
        RS = GetRecordsetBySQL("SELECT * FROM [Mail] WHERE Index=" & Index, , GetDatabaseAtLocation(StoreNo))
        CopyMailRecordsetToMailNew(RS, tMailNew)
        DisposeDA(RS)
    End Sub

    Public Function GetMailByLeaseNo(ByVal LeaseNo As String) As clsMailRec
        '::::GetMailByLeaseNo
        ':::SUMMARY
        ': Gets Mail with Lease number.
        ':::DESCRIPTION
        ': This function is useful to get Mail with Lease number.
        ':::PARAMETERS
        ': - LeaseNo - Indicates the Lease number.
        ':::RETURN
        ': clsMailRec
        Dim R As ADODB.Recordset, I As Integer
        R = GetRecordsetBySQL("SELECT * FROM GrossMargin WHERE SaleNo='" & LeaseNo & "'", , GetDatabaseAtLocation)
        If R.EOF Then Exit Function
        I = R("MailIndex").Value
        GetMailByLeaseNo = New clsMailRec
        GetMailByLeaseNo.Load(I, "#index")
    End Function

    Public Function GetMailLastNameByIndex(ByVal Index As Integer, Optional ByVal StoreNo As Integer = 0, Optional ByVal ShipToIfAvailable As Boolean = False) As String
        '::::GetMailLastByIndex
        ':::SUMMARY
        ': Gets a MailLastName with Index.
        ':::DESCRIPTION
        ': This function is useful to get MailLastName with Index.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ': - StoreNo - Indicates the storenumber.
        ': - ShipToIfAVailable - Indicates Boolean value.
        ':::RETURN
        ': STRING - Returns Mail City string.
        Dim X As MailNew, Y As MailNew2
        On Error Resume Next
        If ShipToIfAvailable Then GetMailNew2ByIndex(Index, Y, StoreNo)
        If Val(Y.Index) <> 0 Then GetMailLastNameByIndex = Y.ShipToLast
        If GetMailLastNameByIndex = "" Then
            GetMailNewByIndex(Index, X, StoreNo)
            GetMailLastNameByIndex = X.Last
        End If
    End Function

    Public Function GetMailCityByIndex(ByVal Index As Integer, Optional ByVal StoreNo As Integer = 0, Optional ByVal ShipToIfAvailable As Boolean = False) As String
        '::::GetMailCityByIndex
        ':::SUMMARY
        ': Gets a Mailcity.
        ':::DESCRIPTION
        ': This function is used to get Mail City with Index value.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ': - StoreNo - Indicates the storenumber.
        ': - ShipToIfAvailable - Indicates Boolean value.
        ':::RETURN
        ': STRING - Returns Mail City string.

        Dim X As MailNew, Y As MailNew2
        On Error Resume Next
        If ShipToIfAvailable Then GetMailNew2ByIndex(Index, Y, StoreNo)
        If Val(Y.Index) <> 0 Then GetMailCityByIndex = Y.City2
        If GetMailCityByIndex = "" Then
            GetMailNewByIndex(Index, X, StoreNo)
            GetMailCityByIndex = X.City
        End If
    End Function

    Public Sub Mail2_SetAtIndex(ByVal Index As String, ByRef Mail2 As MailNew2, Optional ByVal StoreNum As Integer = 0)
        '::::Mail2_SetAtIndex
        ':::SUMMARY
        ': Sets Mail2 recordset through Sql values.
        ':::DESCRIPTION
        ': By calling this function, we can set required information in Mail2 recordset through Sql values.
        ':::PARAMETERS
        ': - Index - Indicates the Index value.
        ': - Mail2 - Represents as MailNew2,which indicates data structure.
        ': - StoreNum - Indicates the Store number.
        ':::RETURN
        Dim dB As String, S As String
        dB = GetDatabaseAtLocation(StoreNum)
        ExecuteRecordsetBySQL("DELETE * FROM MailShipTo WHERE Index=" & Index, , dB)
        S = ""
        S = S & "INSERT INTO [MailShipTo] "
        S = S & "([Index],[Last],[First],[Address],[City],[Zip],[Tele],[PhoneLabel3],[Blank]) "
        S = S & "VALUES ("
        S = S & Mail2.Index & ", "
        S = S & "'" & ProtectSQL(Mail2.ShipToFirst) & "', "
        S = S & "'" & ProtectSQL(Mail2.ShipToLast) & "', "
        S = S & "'" & ProtectSQL(Mail2.Address2) & "', "
        S = S & "'" & ProtectSQL(Mail2.City2) & "', "
        S = S & "'" & ProtectSQL(Mail2.Zip2) & "', "
        S = S & "'" & ProtectSQL(Mail2.Tele3) & "', "
        S = S & "'" & ProtectSQL(Mail2.PhoneLabel3) & "', "
        S = S & "''"
        S = S & ")"
        ExecuteRecordsetBySQL(S, , dB)
    End Sub

End Module

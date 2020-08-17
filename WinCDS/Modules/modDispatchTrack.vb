Module modDispatchTrack

    Public ReadOnly Property DDTLicensed(Optional ByVal CheckTo As String = "#") As Boolean
        Get
            If CheckTo = "#" Then CheckTo = StoreSettings.DispatchTrackLicense
            DDTLicensed = IsIn(CheckTo, LICENSE_DISPATCHTRACK, "TEST")
        End Get
    End Property

    Public Function DDT_Header() As String
        ' Required Columns:
        '   Order Number,Ship Name,Ship Address1,Ship Address2,Ship City,Ship State,Ship Zip,Order Detail,Delivery Date,Delivery Type
        '
        ' Available Columns:
        '   Order Number,Ship Name,Ship Address1,Ship Address2,Ship City,Ship State,Ship Zip,Description,Quantity,Delivery Date,Delivery Type,Customer Code,Bill Name,Bill Address1,Bill Address2,Bill City,Bill State,Bill Zip, Phone1, Phone2, Phone3,Email,Delivery Quantity,Amount,Delivery Charges,Taxes,Service Time,Cube,Truck,Account,Request Start Time,Request End Time,Order Detail,Number,Comment1,Comment2,Comment3,COD,Model,Inventory,Route Label,Store Code
        '
        ' Our Additional Fields:
        '   Salesman
        '
        ' Our Columns:
        '   Order Number,Ship Name,Ship Address1,Ship Address2,Ship City,Ship State,Ship Zip,Description,Quantity,Delivery Date,Delivery Type,Customer Code,
        '   Bill Name,Bill Address1,Bill Address2, Bill City,Bill State,Bill Zip, Phone1, Phone2, Phone3,Email,
        '   Delivery Quantity,Amount,Cube,Request Start Time,Request End Time,Order Detail, Model, Store Code
        '   Salesman

        DDT_Header = ""
        DDT_Header = DDT_Header & "Order Number,Ship Name,Ship Address1,Ship Address2,Ship City,Ship State,Ship Zip,Description,Quantity,Delivery Date,Delivery Type,Customer Code,"
        DDT_Header = DDT_Header & "Bill Name,Bill Address1,Bill Address2,Bill City,Bill State,Bill Zip, Phone1, Phone2, Phone3,Email,"
        DDT_Header = DDT_Header & "Delivery Quantity,Amount,Cube,Request Start Time,Request End Time,Order Detail,Model,Store Code"
        DDT_Header = DDT_Header & ",Salesman"
    End Function

    Private ReadOnly Property Scode() As String ' store code
        Get
            Scode = StoresSld
        End Get
    End Property

    Public Function DDT_ExportServiceOrder(ByVal SON As String, Optional ByVal StoreNum As Integer = 0) As String
        Dim V As clsServiceOrder
        Dim M As clsMailRec, S As clsMailShipTo
        Dim Cubes As Double, CCode As String
        Dim Desc As String

        If StoreNum = 0 Then StoreNum = Scode

        V = New clsServiceOrder

        If Not V.Load(SON, "#ServiceOrderNo") Then
            DisposeDA(V)
            Exit Function
        End If
        CCode = V.MailIndex
        M = LoadMailRecord(V.MailIndex)
        S = M.ShipTo()

        Desc = "SERVICE ORDER"

        '  Cubes = GetCubesByStyle(v.st) * C.Quantity
        DDT_ExportServiceOrder = DDT_Line("SO" & V.ServiceOrderNo, M.First & " " & M.Last, M.Address, M.AddAddress, GetWinCDSCity(M.City), GetWinCDSState(M.City), M.Zip, Desc, 1, V.ServiceOnDate, DType("service order"), CCode,
                   S.First & " " & S.Last, S.Address, "", GetWinCDSCity(S.City), GetWinCDSState(S.Zip), S.Zip, M.Tele, M.Tele2, S.Tele, M.Email,
                   1, "0.00", Cubes, "", "", "", "", StoreNum, "")

        DisposeDA(V, M, S)
    End Function

    Public Function DDT_ExportTransfer(ByVal DID As String, Optional ByVal StoreNum As Integer = 0) As String
        Dim V As CInventoryDetail
        Dim Cubes As Double
        Dim ToStore As Integer
        Dim Desc As String

        If StoreNum = 0 Then StoreNum = Scode


        V = New CInventoryDetail

        If Not V.Load(DID, "#DetailID") Then
            DisposeDA(V)
            Exit Function
        End If

        ToStore = V.GetFirstLocationWithPositiveQuantity
        Cubes = GetCubesByStyle(V.Style) * V.AmtS1
        Desc = GetDescByStyle(V.Style)

        '  If V.GetLocationQuantity(StoreNum) > 0 Then Exit Function
        DDT_ExportTransfer = DDT_Line("TR" & V.Misc, StoreSettings(ToStore).Name, StoreSettings(ToStore).Address, StoreSettings(ToStore).Name, GetWinCDSCity(StoreSettings(ToStore).City), GetWinCDSState(StoreSettings(ToStore).City), GetWinCDSZip(StoreSettings(ToStore).City), Desc, -V.GetLocationQuantity(StoreNum), V.DDate1, DType("transfer"), "",
                 StoreSettings(ToStore).Name, StoreSettings(ToStore).Address, "", GetWinCDSCity(StoreSettings(ToStore).City), GetWinCDSState(StoreSettings(ToStore).City), GetWinCDSZip(StoreSettings(ToStore).City), StoreSettings(ToStore).Phone, "", "", StoreSettings(ToStore).Email,
                 -V.GetLocationQuantity(StoreNum), 0#, Cubes, "", "", "", V.Style, StoreNum, "")

        DisposeDA(V)
    End Function

    Public Function DDT_ExportMarginLine(ByVal MarginLine As String, Optional ByVal StoreNum As Integer = 0) As String
        Dim C As CGrossMargin
        C = New CGrossMargin
        C.DataAccess.DataBase = GetDatabaseAtLocation(StoreNum)
        If C.Load(MarginLine, "#MarginLine") Then
            If C.PorD <> "P" Then
                DDT_ExportMarginLine = DDT_MarginToCSV(C, StoreNum)
            End If
        End If
        DisposeDA(C)
    End Function

    Public Function DDT_DoUploadData(ByVal dat As String, ByVal vData As Collection) As Boolean
        Dim X As String, S As String, Res As String
        X = DDT_GenerateXMLFromCalendarData(dat, vData)

        If IsDevelopment() Then
            WriteFile(DevOutputFolder() & DateStampFile("DDT-$.xml"), X, True)
            If MessageBox.Show("DEVELOPER MODE:" & vbCrLf & "File Written to Desktop." & vbCrLf2 & "Abort Upload?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Yes Then Exit Function
        End If

        'code=[CODE]
        'api_key=[KEY]
        'data=[<data-file>]
        Dim T As Integer

        Do
            T = Len(X)
            X = Replace(X, vbLf & " ", vbLf)
        Loop While Len(X) <> T

        S = "code=" & DDT_SERVICE_CODE & "&api_key=" & DDT_SERVICE_API & "&data=" & URLEncode(X, False)


        ProgressForm(0, 1, "Uploading...")
        LogFile("DDT", S)
        Res = INETPOST(DDT_Import_URL, S)

        ProgressForm()
        MessageBox.Show(Res, "DDT Result", MessageBoxButtons.OK)
    End Function

    Private Function DDT_Line(LeaseNo, Name, Add, Add2, City, ST, Zip, Desc, Qty, DelDate, DType, CustCode,
    Optional ShipName = "", Optional ShipAdd = "", Optional ShipAdd2 = "", Optional ShipCity = "", Optional ShipST = "", Optional ShipZIP = "", Optional pH1 = "", Optional Ph2 = "", Optional Ph3 = "", Optional Email = "",
    Optional DelQty = "", Optional Amount = "", Optional Cubes = "", Optional StopStart = "", Optional StopEnd = "", Optional OrderDetail = "", Optional Model = "", Optional StoreCode = "" _
    , Optional Salesman = "") As String

        If Trim(ShipName) = "" Then ShipName = Name
        If Trim(ShipAdd) = "" Then
            ShipAdd = Add
            If Trim(ShipAdd2) = "" Then ShipAdd2 = Add2
        End If
        If Trim(ShipCity) = "" Then ShipCity = City
        If Trim(ShipST) = "" Then ShipST = ST
        If Trim(ShipZIP) = "" Then ShipZIP = Zip

        If ST = "" Then ST = GetWinCDSState(StoreSettings.City) ' Occasionally, only a city is supplied.  Assume current state.

        ' These are all 'Required' fields, and will cause validation to fail if not supplied.  We supply them if missing.
        If Name = "" Then Name = "UNKNOWN"
        If Add = "" Then Add = "UNKNOWN"
        If Add2 = "" Then Add2 = "UNKNOWN"
        If ST = "" Then ST = "UNKNOWN"
        If Zip = "" Then Zip = "UNKNOWN"
        If Desc = "" Then Desc = "UNAVAILABLE"
        If Qty = "" Then Qty = "1"
        If DelDate = "" Then DelDate = Today
        If DType = "" Then DType = "1"
        If OrderDetail = "" Then OrderDetail = "NONE"
        '  If Model = "" Then Model = ""
        If Salesman = "" Then Salesman = "UNKNOWN"
        Salesman = TranslateSalesmen(Salesman)

        DDT_Line = CSVLine(LeaseNo, Name, Add, Add2, City, ST, Zip, Desc, Qty, DDT_Date(DelDate), DType, CustCode,
          ShipName, ShipAdd, ShipAdd2, ShipCity, ShipST, ShipZIP, pH1, Ph2, Ph3, Email,
          DelQty, CSVCurrency(Amount), FormatQuantity(Val(Cubes)), StopStart, StopEnd, OrderDetail, Model, StoreCode, Salesman)

        If LeaseNo = "" Then ErrMsg("Missing Order Number on DDT Line: " & DDT_Line) : DDT_Line = "" : Exit Function

        'LeaseNo , Name, Add, Add2, City, St, Zip, Desc, Qty, DDT_Date(DelDate), DType
    End Function

    Private ReadOnly Property DType(Optional ByVal AltType As String = "") As String
        Get
            '  [Optional: example: deliveery, pickup, credit memos, even exchange, service, installation, etc]
            DType = AltType
            If UCase(DType) = "P" Then DType = "pickup"
            If UCase(DType) = "D" Then DType = "delivery"
            If DType = "" Then DType = "delivery"
        End Get
    End Property
    ' Deliverys through Dispatch Track (DDT)

    Public Function DDT_MarginToCSV(ByRef C As CGrossMargin, Optional ByVal StoreNum As Integer = 0) As String
        Dim R As String, L As String
        Dim M As clsMailRec, S As clsMailShipTo
        Dim Cubes As Double, CCode As String, TDesc As String
        Dim Price As String

        If StoreNum = 0 Then StoreNum = Scode


        Do Until C.DataAccess.Record_EOF
            '    If IsItem(C.Style) Then
            CCode = C.Index
            M = LoadMailRecord(C.Index, StoreNum)
            S = M.ShipTo()

            Cubes = GetCubesByStyle(C.Style) * C.Quantity
            TDesc = IfNullThenNilString(C.Vendor) & " (" & IfNullThenNilString(C.VendorNo) & "): " & C.Desc
            Price = HoldNew_GetBalance(C.SaleNo, StoreNum)

            L = DDT_Line(C.SaleNo, M.First & " " & M.Last, M.Address, M.AddAddress, GetWinCDSCity(M.City), GetWinCDSState(M.City), M.Zip, TDesc, C.Quantity, C.DDelDat, DType(C.PorD), CCode,
                   S.First & " " & S.Last, S.Address, "", GetWinCDSCity(S.City), GetWinCDSState(S.Zip), S.Zip, M.Tele, M.Tele2, S.Tele, M.Email,
                   C.Quantity, Price, Cubes, C.StopStart, C.StopEnd, M.Special, C.Style, StoreNum, C.Salesman)
            DisposeDA(M, S)

            R = R & IIf(R = "", "", vbCrLf) & L
            '    End If
            C.DataAccess.Records_MoveNext()
        Loop

        DDT_MarginToCSV = R
    End Function

    Public Function DDT_GenerateXMLFromCalendarData(ByVal dat As String, ByVal vData As Collection) As String
        Dim CD As ADODB.Recordset
        Dim C As String
        Dim A As String, M As String, N As String
        Dim LastSale As String
        Dim IA As Integer, iB As Integer, II As Integer
        Dim TransferList As String
        IA = vData("IA")
        iB = vData("IB")

        M = ""
        N = vbCrLf

        A = ""
        A = A & M & DDT_XML_Header()

        For II = IA To iB
            CD = vData("L" & II)

            'If IsDevelopment And II = 3 Then Stop
            Do Until CD.EOF
                C = ""
                If Left(CD(2).Value, 2) = "SO" Then
                    C = DDT_XML_ServiceOrder(dat, "Service", CD(2).Value, II, Nothing)
                ElseIf Left(CD(2).Value, 2) = "TR" Then
                    TransferList = TransferList & IIf(TransferList = "", "", ",") & CD("record").Value
                    '          C = DDT_XML_ServiceOrder(dat, "Transfer", CD("record"), II, CD)
                    '        C = DDT_XML_ServiceOrder(dat, CD("Record"), "Transfer", CD(2).Value, StoreNum)
                Else
                    Dim G As CGrossMargin
                    G = New CGrossMargin
                    G.DataAccess.DataBase = GetDatabaseAtLocation(StoreNum)
                    If G.Load(CD("record").Value, "#MarginLine") Then
                        C = DDT_XML_ServiceOrder(dat, "Sale", CD("record").Value, II, CD)
                    End If
                    '          C = DDT_XML_ServiceOrder(dat, "Sale", CD("record"), II, CD)
                End If

NoRecord:
                If C <> "" Then A = A & N & C

                CD.MoveNext()
            Loop

            If TransferList <> "" Then
                CD.MoveFirst()
                C = DDT_XML_ServiceOrder(dat, "Transfer", TransferList, II, CD)
                If C <> "" Then A = A & N & C
            End If


        Next

        A = A & N & DDT_XML_Footer()

        DDT_GenerateXMLFromCalendarData = A
    End Function

    Private ReadOnly Property DDT_SERVICE_CODE() As String
        Get
            DDT_SERVICE_CODE = StoreSettings.DispatchTrackServiceCode
        End Get
    End Property

    Private ReadOnly Property DDT_SERVICE_API() As String
        Get
            DDT_SERVICE_API = StoreSettings.DispatchTrackServiceAPI
        End Get
    End Property

    Private ReadOnly Property DDT_Import_URL() As String
        Get
            ' https://[ServerName]/orders/api/import
            DDT_Import_URL = "https://" & DDT_SERVICE_URL & "/orders/api/import"
        End Get
    End Property

    Private ReadOnly Property DDT_SERVICE_URL() As String
        Get
            DDT_SERVICE_URL = StoreSettings.DispatchTrackServiceURL
            DDT_SERVICE_URL = Replace(DDT_SERVICE_URL, "http://", "")
            DDT_SERVICE_URL = Replace(DDT_SERVICE_URL, "https://", "")
            DDT_SERVICE_URL = Replace(DDT_SERVICE_URL, "/", "")
        End Get
    End Property

    Private Function DDT_Date(Optional ByVal DateIn As String = "") As String
        If Not IsDate(DateIn) Then DateIn = Today
        DDT_Date = DateFormat(DateValue(DateIn))
    End Function

    Private Function DDT_XML_Header() As String
        Dim A As String, M As String, N As String
        M = ""
        N = vbCrLf

        A = ""
        A = ""
        A = A & M & "<?xml version='1.0' encoding='UTF-8'?>"
        A = A & N & "<service_orders>"

        DDT_XML_Header = A
    End Function

    Private Function DDT_XML_ServiceOrder(ByVal dat As String, ByVal Typ As String, ByVal Record As String, Optional ByVal StoreNum As Integer = 0, Optional CD As ADODB.Recordset = Nothing)
        Const ACCOUNT_ID = ""

        If StoreNum = 0 Then StoreNum = StoresSld

        Dim SaleN As Integer
        Dim SNo As String
        Dim AcctNo As String

        Dim A As String, M As String, N As String, C As String
        M = ""
        N = vbCrLf

        Dim T As clsMailRec, S As clsMailShipTo
        Dim Indx As Integer, Last As String, Frst As String, Addr As String, Add2 As String, City As String, Stat As String, Zipc As String, Desc As String, Emai As String
        Dim Pho1 As String, Pho2 As String, Pho3 As String, Emal As String
        Dim Spec As String
        Dim ShipLast As String, ShipFrst As String, ShipAddr As String, ShipAdd2 As String, ShipCity As String, ShipSTat As String, ShipZIPc As String
        Dim Style As String, Quant As Double, SellPrice As Decimal, DDelDat As String, SellDat As String
        Dim StopStart As String, StopEnd As String

        Dim Salesman As String

        Dim Cubes As Double

        Dim Ty As String

        If Typ = "Sale" Then
            Dim CG As CGrossMargin
            CG = New CGrossMargin
            CG.DataAccess.DataBase = GetDatabaseAtLocation(StoreNum)
            CG.Load(Record, "#MarginLine")
            T = CG.LoadMailRecord
            S = T.ShipTo()


            AcctNo = CG.SaleNo
            SellDat = CG.SellDte
            DDelDat = CG.DDelDat

            StopStart = CG.StopStart
            StopEnd = CG.StopStart

            Salesman = TranslateSalesmen(CG.Salesman)
            Ty = ""
        ElseIf Typ = "Service" Then
            Dim V As clsServiceOrder
            V = New clsServiceOrder
            V.DataAccess.DataBase = GetDatabaseAtLocation(StoreNum)
            If Not V.Load(Mid(Record, 3), "#ServiceOrderNo") Then Exit Function

            AcctNo = Record
            Style = ""
            Desc = "Service Order #" & Record
            Quant = 1
            T = LoadMailRecord(V.MailIndex, StoreNum)
            S = T.ShipTo()

            SellDat = dat
            DDelDat = dat

            StopStart = ""
            StopEnd = ""

            Salesman = ""
            DisposeDA(V)

            Ty = "SO"
        ElseIf Typ = "Transfer" Then
            '    Dim W As CInventoryDetail
            '    Set W = New CInventoryDetail
            '    If Not W.Load(Record, "#DetailID") Then Exit Function
            AcctNo = "Transfers" ' W.Misc ' Transfer No
            SellDat = dat
            DDelDat = dat

            StopStart = ""
            StopEnd = ""

            Salesman = ""
            '    DisposeDA W

            Ty = "TR"
        End If

        If Not (T Is Nothing) Then
            Indx = T.Index

            ShipLast = IIf(S.Last <> "", S.Last, T.Last)
            ShipFrst = IIf(S.First <> "", S.First, T.First)
            ShipAddr = IIf(S.Address <> "", S.Address, T.Address)
            ShipAdd2 = IIf(S.Address <> "", "", T.AddAddress)
            ShipCity = IIf(S.City <> "", S.City, T.City)
            ShipZIPc = IIf(S.Zip <> "", S.Zip, T.Zip)

            ' These are all 'Required' fields, and will cause validation to fail if not supplied.  We supply them if missing.
            Last = IIf(ShipLast <> "", ShipLast, IIf(T.Last <> "", T.Last, "UNKNOWN"))
            Frst = IIf(ShipFrst <> "", ShipFrst, IIf(T.First <> "", T.First, "UNKNOWN"))
            Addr = IIf(ShipAddr <> "", ShipAddr, IIf(T.Address <> "", T.Address, "UNKNOWN"))
            Add2 = IIf(ShipAdd2 <> "", ShipAdd2, IIf(T.AddAddress <> "", T.AddAddress, "UNKNOWN"))
            City = IIf(ShipCity <> "", ShipCity, IIf(GetWinCDSCity(T.City) <> "", GetWinCDSCity(T.City), "UNKNOWN"))
            Stat = IIf(ShipSTat <> "", ShipSTat, IIf(GetWinCDSState(T.City) <> "", GetWinCDSState(T.City), "UNKNOWN"))
            Zipc = IIf(ShipZIPc <> "", ShipZIPc, IIf(T.Zip <> "", T.Zip, "UNKNOWN"))

            Emai = T.Email

            Pho1 = T.Tele
            Pho2 = T.Tele2
            Pho3 = S.Tele

            Spec = T.Special
        Else
            Indx = 0

            Last = StoreSettings.Name
            Frst = ""
            Addr = StoreSettings.Address
            Add2 = ""
            City = GetWinCDSCity(StoreSettings.City)
            Stat = GetWinCDSState(StoreSettings.City)
            Zipc = GetWinCDSZip(StoreSettings.City)

            Emai = StoreSettings.Email
            Pho1 = StoreSettings.Phone
        End If

        Cubes = GetCubesByStyle(Style) * Quant

        A = ""
        A = A & M & "  <service_order>"
        A = A & N & "    <number>" & Ty & ProtectXML(AcctNo) & "</number>"
        A = A & N & "    <account>" & ACCOUNT_ID & "</account>"
        A = A & N & "    <service_type>" & IIf(Typ = "Service", "Service", "Delivery") & "</service_type>"
        A = A & N & "    <description>" & ProtectCDATA("", True) & "</description>"
        A = A & N & "    <customer>"
        A = A & N & "      <customer_id>" & ProtectXML(Indx) & "</customer_id>"
        A = A & N & "      <first_name>" & ProtectXML(Frst) & "</first_name>"
        A = A & N & "      <last_name>" & ProtectXML(Last) & "</last_name>"
        A = A & N & "      <email>" & ProtectXML(Emai) & "</email>"
        A = A & N & "      <phone1>" & ProtectXML(Pho1) & "</phone1>"
        A = A & N & "      <phone2>" & ProtectXML(Pho2) & "</phone2>"
        A = A & N & "      <phone3>" & ProtectXML(Pho3) & "</phone3>"
        A = A & N & "      <address1>" & ProtectXML(Addr) & "</address1>"
        A = A & N & "      <address2>" & ProtectXML(Add2) & "</address2>"
        A = A & N & "      <city>" & ProtectXML(City) & "</city>"
        A = A & N & "      <state>" & ProtectXML(Stat) & "</state>"
        A = A & N & "      <zip>" & ProtectXML(Zipc) & "</zip>"
        '  A = A & N & "      <latitude>{Customer Geographic latitude (optional)}</latitude>"
        '  A = A & N & "      <longitude>{Customer Geographic longitude (optional)}</longitude>"
        A = A & N & "    </customer>"

        If Spec <> "" Then
            A = A & N & "    <notes count='1'>" ' modify if > 1
            A = A & N & "      <note created_at='" & DDT_XMLDateTime() & "' author='{USER/SERVICE_UNIT_LOGIN}'>" & ProtectCDATA(Spec, True) & "</note>"
            '  A = A & N & "      <note created_at='{DATE_TIMESTAMP}' author='{USER/SERVICE_UNIT_LOGIN}'>"
            '  A = A & N & "        <![CDATA[]]>"
            '  A = A & N & "      </note>"
            A = A & N & "    </notes>"
        End If
        A = A & N & "    <items>"


        SaleN = 1
        If Typ = "Sale" Then
            SNo = CD(2).Value
            C = DDT_XML_MargineLine(SaleN, Record, StoreNum)
            If C <> "" Then
                A = A & N & C
                SaleN = SaleN + 1
            End If
AnotherSaleItem:
            CD.MoveNext()

            If CD.EOF Then
                CD.MovePrevious()
                GoTo NoMoreSaleItems
            ElseIf CD(2).Value <> SNo Then
                CD.MovePrevious()
                GoTo NoMoreSaleItems
            Else
                C = DDT_XML_MargineLine(SaleN, CD("Record").Value, StoreNum)
                If C <> "" Then
                    A = A & N & C
                    SaleN = SaleN + 1
                End If
                GoTo AnotherSaleItem
            End If
NoMoreSaleItems:
        ElseIf Typ = "Transfer" Then
            A = A & N & DDT_XML_DetailID(SaleN, CD("Record").Value, StoreNum)
AnotherTransferItem:
            CD.MoveNext()

            If CD.EOF Then
                '      CD.MovePrevious
                GoTo NoMoreTransferItems
                '    ElseIf Left(CD(2).Value, 2) <> "TR" Then
                '      CD.MovePrevious
                '      GoTo NoMoreTransferItems
            ElseIf Left(CD(2).Value, 2) = "TR" Then
                SaleN = SaleN + 1
                A = A & N & DDT_XML_DetailID(SaleN, CD("Record").Value, StoreNum)
                GoTo AnotherTransferItem
            End If
NoMoreTransferItems:
        Else ' Service
            A = A & N & DDT_XML_ITEM(1, Style, "0", Desc, Quant, StoreNum, Cubes, 0)
        End If
        A = A & N & "    </items>"
        '  A = A & N & "    <pre_reqs>100123,100124,100454</pre_reqs>"
        '  A = A & N & "    <!-- Comma Separated Uniq Order Numbers -->"
        A = A & N & "    <amount>" & SQLCurrency(0) & "</amount>"
        '  A = A & N & "    <cod_amount>{COD Amount on delivery/service (optional)}</cod_amount>"
        '  A = A & N & "    <service_unit>{name of the resource/route (optional)}</service_unit>"
        A = A & N & "    <delivery_date>" & DDT_XMLDate(SellDat) & "</delivery_date>"
        A = A & N & "    <request_delivery_date>" & DDT_XMLDateTime(DDelDat) & "</request_delivery_date>"
        '  A = A & N & "  <driver_id>{driver ID if pre-assigned (optional)}</driver_id>"
        '  A = A & N & "  <truck_id>{vehicle ID if pre-assigned (optional)}</truck_id>"
        '  A = A & N & "  <origin>{Warehouse (optional)}</origin>"
        '  A = A & N & "  <stop_number>{Manifest Stop number (optional)}</stop_number>"
        '  A = A & N & "  <stop_time>{Manifest Stop time (optional)}</stop_time>"
        '  A = A & N & "  <service_time>{Total time at the stop (optional)}</service_time>"
        A = A & N & "  <request_time_window_start>" & DDT_XMLDateTime(DDT_Date() & " " & StopStart) & "</request_time_window_start>"
        A = A & N & "  <request_time_window_end>" & DDT_XMLDateTime(DDT_Date() & " " & StopEnd) & "</request_time_window_end>"
        '  A = A & N & "  <delivery_time_window_start>{Window request Start Time}</delivery_time_window_start>"
        '  A = A & N & "  <delivery_time_window_end>{Window request End Time}</delivery_time_window_end>"
        A = A & N & "<extra>"
        '  A = A & N & "  <custom_field_1>value1</custom_field_1>"
        '  A = A & N & "  <custom_field_2>value2</custom_field_2>"
        '  A = A & N & "  <custom_field_3>value3</custom_field_3>"
        A = A & N & "  <salesman>" & Salesman & "</salesman>"
        A = A & N & "</extra>"
        A = A & N & "</service_order>"

        DDT_XML_ServiceOrder = A
    End Function

    Private Function DDT_XML_Footer() As String
        Dim A As String, M As String, N As String
        M = ""
        N = vbCrLf

        A = ""
        A = A & N & "</service_orders>"

        DDT_XML_Footer = A
    End Function

    Private Function DDT_XMLDateTime(Optional ByVal DateIn As String = "") As String
        If Not IsDate(DateIn) Then DateIn = Now
        DDT_XMLDateTime = Format(DateValue(DateIn), "yyyy-mm-dd") & " " & Format(TimeValue(DateIn), "HH:mm:ss") & " " & GetCurrentTimeZoneOffset()
    End Function

    Private Function DDT_XML_MargineLine(ByVal Idx As Integer, ByVal ML As Integer, ByVal StoreNum As Integer) As String
        Dim C As CGrossMargin
        C = New CGrossMargin
        C.DataAccess.DataBase = GetDatabaseAtLocation(StoreNum)
        C.Load(ML, "#MarginLine")

        Dim Price As Decimal
        Price = GetPrice(HoldNew_GetBalance(C.SaleNo, StoreNum))

        If C.PorD <> "P" And IsDeliverable(C.Status, C.Style, True) Then
            DDT_XML_MargineLine = DDT_XML_ITEM(Idx, C.Style, "0", IfNullThenNilString(C.Vendor) & " (" & IfNullThenNilString(C.VendorNo) & "): " & C.Desc, C.Quantity, StoreNum, GetCubesByStyle(C.Style) * C.Quantity, Price)
        End If
        DisposeDA(C)
    End Function

    Private Function DDT_XML_DetailID(ByVal Idx As Integer, ByVal DID As Integer, ByVal StoreNum As Integer) As String
        Dim W As CInventoryDetail
        W = New CInventoryDetail
        If Not W.Load(DID, "#DetailID") Then Exit Function
        DDT_XML_DetailID = DDT_XML_ITEM(Idx, W.Style, GetDescByStyle(W.Style), "Transfer #" & W.Misc, 0, StoreNum, 0, 0)
        DisposeDA(W)
    End Function

    Private Function DDT_XML_ITEM(ByVal Idx As Integer, ByVal itemID As String, ByVal SerNo As String, ByVal Desc As String, ByVal Qty As Double, ByVal Loc As Integer, ByVal Cube As Double, ByVal Price As Decimal) As String
        Dim A As String, N As String

        A = ""
        A = A & N & "      <item>"
        N = vbCrLf
        A = A & N & "        <sale_sequence>" & Idx & "</sale_sequence>"
        A = A & N & "        <item_id>" & ProtectXML(itemID) & "</item_id>"
        A = A & N & "        <serial_number>" & ProtectXML("0") & "</serial_number>"
        A = A & N & "        <description>" & ProtectCDATA(Desc, True) & "</description>"
        A = A & N & "        <quantity>" & FormatQuantity(Qty) & "</quantity>"
        A = A & N & "        <location>" & Loc & "</location>"
        A = A & N & "        <cube>" & FormatQuantity(Cube, 2, False) & "</cube>"
        '  A = A & N & "        <setup_time>{Item Setup Time(optional)}</setup_time>"
        '  A = A & N & "        <weight>{Item Weight (optional)}</weight>"
        A = A & N & "        <price>" & XMLCurrency(Price) & "</price>"
        '  A = A & N & "        <countable>{true/false – default (true) – (optional)}</countable>"
        A = A & N & "      </item>"

        DDT_XML_ITEM = A
    End Function

    Private Function DDT_XMLDate(Optional ByVal DateIn As String = "") As String
        If Not IsDate(DateIn) Then DateIn = Today
        DDT_XMLDate = Format(DateValue(DateIn), "yyyy-mm-dd")
    End Function

End Module

Imports Microsoft.VisualBasic.Interaction
Public Class clsSaleItem
    Public TransID As String
    Public Style As String
    Public Quantity As Double
    Public DisplayPrice As Decimal
    Public Price As Decimal
    Public NonTaxable As Boolean
    Public Desc As String

    Public Status As String
    Public Location As String

    Public Vendor As String
    Public VendorNo As String

    Public Cost As Decimal
    Public Landed As Decimal
    Public Freight As Decimal

    Public Extra1 As String ' BFH20060602
    Public Extra2 As String
    Public Balance As Decimal    ' BFH20110629 - for partial approval support, mc/visa gift card balance, will show a NOTE line on receipt

    Public Sub LoadPricing(Optional ByVal vStyle As String = "")
        Dim InvData As CInvRec
        InvData = New CInvRec
        With InvData
            If .Load(Style, "Style") Then
                Cost = .Cost * Quantity
                Freight = IIf(.FreightType = 0, .Freight, .Freight * .Cost) * Quantity
                Landed = .Landed * Quantity
            End If
        End With
        DisposeDA(InvData)
    End Sub

    Public Sub AddItemGrossMargin(ByVal Sale As sSale)
        Dim A As String, B As String
        Dim cGM As CGrossMargin, CI As CInvRec, Found As Boolean
        Dim DetailNo as integer, MarginNo as integer

        On Error GoTo HandleErr
        cGM = New CGrossMargin

        CI = New CInvRec
        Found = CI.Load(Style, "Style")

        If Found Then DetailNo = Me.CreateDetail(Sale, CI)

        With cGM
            .SaleNo = Sale.SaleNo

            .Quantity = Trim(Quantity)
            .Style = Trim(Style)
            .Vendor = Trim(Vendor)          ' SS items might have this set
            .VendorNo = Trim(VendorNo)
            .Desc = Trim(Desc)

            .PorD = Switch(Sale.PorD = "P", "P", Sale.PorD = "D", "D", True, "") ' only P, D, or blank

            If Not IsIn(Status, "SS", "SSLAW") Then
                If Found Then
                    .VendorNo = CI.VendorNo
                    .DeptNo = CI.DeptNo
                Else
                    .VendorNo = ""
                    .DeptNo = ""
                End If
            End If
            If .VendorNo = "" Then .VendorNo = GetVendorNoFromName(.Vendor)

            If Found Then
                .Cost = GetItemCost(.Style, Sale.Store, , .Quantity)
                .ItemFreight = (CI.Landed - CI.Cost) * .Quantity
                .Spiff = CI.Spiff * .Quantity
            Else
                .Cost = 0
                .ItemFreight = 0
                .Spiff = 0
            End If

            .SellPrice = Price
            .Salesman = Sale.SalesCode
            .SalesSplit = Sale.SalesSplit

            .Status = Status
            .Location = Location
            .SellDte = Sale.SaleDate

            If .Status = "DELTW" Then
                .DDelDat = Sale.SaleDate
            ElseIf Sale.PorD <> "" And Sale.SaleDate <> "" Then
                .DDelDat = Sale.DelDate
                .StopStart = Sale.StopStart
                .StopEnd = Sale.StopEnd
            End If

            .Store = Sale.Store
            .Name = Sale.Name
            .ShipDte = ""
            If OrderMode("B") Then .DDelDat = InvDel.TransDate   'deliver sales

            .Index = Sale.MailIndex

            ' sets SS in data base for Reports
            .SS = IIf(.Status = "SS" Or .Status = "FND" Or .Status = "SSLAW", .Status, "")

            .Detail = 0
            .GM = 0

            If Found Then
                .RN = CI.RN
                .GM = CalculateGM(.SellPrice, .ItemFreight + .Cost, 0)
            Else
                .RN = 0
                .GM = 0
            End If
            .Detail = DetailNo

            .Phone = CleanAni(Sale.Tele)
            .TransID = TransID

            On Error GoTo SaveError
            '.DataAccess.Records_AddAndClose()
            .DataAccess.Records_AddAndClose1()
            cGM.cDataAccess_SetRecordSet(.DataAccess.RS)
            .DataAccess.Records_AddAndClose2()
            'MarginNo = .MarginLine       -------> This line replaced with the below block using GetRecordSetBySQL.
            Dim rsMax As New ADODB.Recordset
            rsMax = GetRecordsetBySQL("Select max(MarginLine) from GrossMargin", True, GetDatabaseAtLocation)
            If Not rsMax.EOF And Not rsMax.BOF Then
                MarginNo = rsMax(0).Value
            End If

            On Error GoTo HandleErr

            If .Detail <> 0 Then SetDetailMarginLine(DetailNo, MarginNo)

            If .Status = "SO" Or .Status = "SS" Then MakePo(Sale, cGM)
        End With

        DisposeDA(cGM, CI)

        Exit Sub

HandleErr:
        '  WriteError Err.Number, Err.Description
        '  If Err.Number = 75 Then Resume  ' We don't write to files anymore...
        ErrMsg("Error when saving item [" & Style & "]." & vbCrLf & "clsSaleItem.AddItemGrossMargin - SaleNo=" & Sale.ToString & vbCrLf & "[" & Err.Description & "] " & Err.Description, "Error Creating Sale Item")
        Debug.Print("ERROR in sSale.AddItemGrossMargin [" & Err.Number & "]: " & Err.Description)
        Err.Clear()
        Resume Next

SaveError:
        MsgBox("Save error in sSale.AddItemGrossMargin [" & Err.Number & "]" & vbCrLf & Err.Description, vbCritical, "Error")
        Err.Clear()
        Resume Next
    End Sub

    Public Function MakePo(ByVal Sale As sSale, ByVal Margin As CGrossMargin) As String
        ' This doesn't seem to update InvData's OnOrder fields.
        On Error GoTo HandleErr
        Dim PO As New cPODetail
        With PO
            .PoNo = Sale.PoNo(Me)
            .SaleNo = Sale.SaleNo
            ' bfh20050825
            ' this shouldn't be the original sale date, but rather the current date (which is when the po was made!)
            .PoDate = Sale.SaleDate
            .Name = Sale.Name
            .Vendor = Vendor
            .InitialQuantity = Quantity
            .Quantity = Quantity
            .Style = Style
            .Desc = Trim(Desc)

            'If Trim(Margin.Status) = "SS" Then Margin.Cost = 0: Margin.ItemFreight = 0  ' This doesn't matter, Margin has already been saved.
            .Cost = Format(Margin.Cost, "Currency")

            .Location = Margin.Store  'location sold from 04-01-2002
            .SoldTo = "1"
            'BFH20060512 - Added b/c F1 wanted to make a SO sale in Store 2 or 3, selecting loc 1, and have shipto show loc 1
            If .Location <> Margin.Location And IsFurnOne() Then
                .ShipTo = "3"
                .ShiptoName = StoreSettings(Margin.Location).Name
                .ShipToAddress = StoreSettings(Margin.Location).Address
                .ShipToCity = StoreSettings(Margin.Location).City
                .ShipToTele = StoreSettings(Margin.Location).Phone
            Else
                .ShipTo = "2"  '04-01-2002 SHOULD BE DEFAULT LOCATION
            End If

            If StoreSettings.bPOSpecialInstr Then
                .Note1 = "1"
                .Note2 = "1"
            Else
                .Note1 = "0"
                .Note2 = "0"
            End If

            .Note3 = "0"
            .Note4 = "0"
            .PoNotes = ""
            .AckInv = ""
            .Posted = ""
            .PrintPo = ""
            .wCost = "1" ' Print w/Cost
            If StoreSettings.bPrintPoNoCost Then .wCost = "0"
            .RN = Margin.RN               'added margin.. 11-07-01
            .Detail = Margin.Detail       'added margin.. 11-07-01

            ' .MarginLine will be empty for Stock orders.
            If OrderMode("A") Then   'changed 11-07-01
                .MarginLine = Margin.MarginLine
            Else
                .MarginLine = MailCheck.MarginNo
            End If
        End With
        PO.Save()
        MakePo = PO.PoNo

        If IsDoddsLtd Then
            If PO.Location = 0 Then
                MsgBox("PO Loc = 0 on create sale.")
            End If
        End If

        DisposeDA(PO)

        Exit Function

HandleErr:
        Resume Next
    End Function

    Public Function CreateDetail(ByVal Sale As sSale, ByRef CI As CInvRec) as integer
        ' prevents a SS or SSlaw, FND from going into detail or inventory data base
        If IsIn(Status, "", "SS", "SSLAW", "FND") Then Exit Function

        Dim InvDetail As CInventoryDetail
        On Error GoTo ErrHandler

        InvDetail = CreateDetailRecord(CI, Sale.SaleNo, Sale.Name,
      Quantity, Status, Val(Location), DateFormat(Sale.SaleDate), Sale.PoNo(Me))
        CreateDetail = InvDetail.DetailID
        DisposeDA(InvDetail)

        Exit Function

ErrHandler:
        MsgBox("Error in clsSaleItem.CreateDetail: " & Err.Description)
        Err.Clear()
        Resume Next
    End Function

    Public Sub LoadVendor(Optional ByVal vStyle As String = "", Optional ByRef VendorNo As String = "", Optional ByRef DeptNo As String = "")
        Vendor = GetVendorByStyle(Style, VendorNo, DeptNo)
    End Sub

    Private Function SetDetailMarginLine(ByVal DetailNo As Integer, ByVal MarginNo As Integer) As Boolean
        Dim InvDetail As CInventoryDetail
        InvDetail = New CInventoryDetail
        If InvDetail.Load(DetailNo, "#DetailID") Then
            InvDetail.cDataAccess_GetRecordSet(InvDetail.DataAccess.RS)
            InvDetail.MarginRn = MarginNo
            InvDetail.Save()
        End If
        DisposeDA(InvDetail)
    End Function

End Class

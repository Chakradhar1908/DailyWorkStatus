Imports Microsoft.VisualBasic.Compatibility.VB6
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class CGrossMargin
    Dim printer As New Printer
    Public Structure GrossMargin
        <VBFixedString(8)> Dim SaleNo As String
        <VBFixedString(5)> Dim Quantity As String
        <VBFixedString(16)> Dim Style As String
        <VBFixedString(16)> Dim Vendor As String
        <VBFixedString(46)> Dim Desc As String
        <VBFixedString(1)> Dim PorD As String
        <VBFixedString(4)> Dim Code As String
        <VBFixedString(1)> Dim Commission As String
        <VBFixedString(8)> Dim Cost As String
        <VBFixedString(6)> Dim ItemFreight As String
        <VBFixedString(8)> Dim SellPrice As String
        <VBFixedString(12)> Dim Salesman As String
        <VBFixedString(5)> Dim Status As String
        <VBFixedString(2)> Dim Location As String
        <VBFixedString(10)> Dim SellDte As String
        <VBFixedString(10)> Dim DDelDat As String
        <VBFixedString(5)> Dim RN As String
        <VBFixedString(2)> Dim Store As String
        <VBFixedString(12)> Dim Name As String
        <VBFixedString(10)> Dim ShipDte As String
        <VBFixedString(5)> Dim GM As String
        <VBFixedString(12)> Dim Phone As String
        <VBFixedString(6)> Dim Index As String
        <VBFixedString(6)> Dim Detail As String
        'access
        'CommPd As String * 12
    End Structure

    Public SaleNo As String
    Public Quantity As Double
    Public Style As String
    Public Vendor As String
    Public Desc As String
    Public PorD As String
    Public DeptNo As String
    Public VendorNo As String
    Public Commission As String
    Public Cost As Decimal
    Public ItemFreight As Decimal
    Public SellPrice As Decimal
    Public Salesman As String
    Public Status As String
    Public Location As Integer
    Public SellDte As String
    Public DDelDat As String
    Public RN As String
    Public Store As Integer
    Public Name As String
    Public ShipDte As String
    Public GM As String
    Public Phone As String
    Public Index As String
    Public Detail As Integer
    Public SS As String
    Public DelPrint As String
    Public PullPrint As String
    Public CommPd As Date
    Public MarginLine As String
    Public Spiff As Decimal
    Public SalesSplit As String
    Public StopStart As String
    Public StopEnd As String
    Public IsPackage As Boolean
    Public PackSell As Decimal
    Public PackSellGM As Double
    Public PackSaleGM As Double
    Public TransID As String

    Private mDataConvert As cDataConvert
    'Implements cDataConvert
    'Private mDataConvert As cDataConvert
    'Private WithEvents mDataAccess As CDataAccess
    'Implements CDataAccess
    Private WithEvents mDataAccess As CDataAccess

    Private Const FILE_Name = "Margin.exe"
    Private Const FILE_RecordSize = 216
    Private Const FILE_Index = 10
    Private Const TABLE_NAME = "GrossMargin"
    Private Const TABLE_INDEX = "SaleNo"

    Public Sub New()
        CDataConvert_Init()
        CDataAccess_Init()
    End Sub

    Public Sub CDataConvert_Init()
        mDataConvert = New cDataConvert
        With mDataConvert
            .SubClass = Me.mDataConvert
            .DataBase = GetDatabaseAtLocation()
            .Table = TABLE_NAME
            .Index = TABLE_INDEX
        End With
    End Sub

    Public Sub CDataAccess_Init()
        mDataAccess = New CDataAccess
        With mDataAccess
            .SubClass = Me.mDataAccess
            .DataBase = GetDatabaseAtLocation()
            .Table = TABLE_NAME
            .Index = TABLE_INDEX
        End With
    End Sub

    ' These are needed for the practice.frm WHY?, to know about friend.count
    Public Function Count() As Integer
        Count = mDataConvert.Count
    End Function

    Public Sub ConvertData()
        mDataConvert.ConvertData()
    End Sub

    Public Function CDataConvert_SuperClass() As cDataConvert
        CDataConvert_SuperClass = mDataConvert
    End Function

    Private Sub cDataConvert_FileOpen()
        'Open(NewOrderFolder() & FILE_Name) For Random Shared As FILE_Index Len = FILE_RecordSize
        FileOpen(FILE_Index, NewOrderFolder() & FILE_Name, OpenMode.Random,, OpenShare.Shared, FILE_RecordSize)
    End Sub
    '    Private Property Get cDataConvert_FileRecords() as integer
    '  cDataConvert_FileRecords = LOF(FILE_Index) / FILE_RecordSize
    'End Property
    Private ReadOnly Property cDataConvert_FileRecords() As Integer
        Get
            cDataConvert_FileRecords = LOF(FILE_Index) / FILE_RecordSize
        End Get
    End Property

    Private Sub cDataConvert_SetRecordSet(Index As Integer, RS As ADODB.Recordset)
        Dim tDataStruct As New GrossMargin
        'Get( #FILE_Index, Index + 1, tDataStruct)

        FileGet(FILE_Index, tDataStruct, Index + 1)

        On Error Resume Next
        With tDataStruct
            RS("MarginLine").Value = Index + 1
            RS("SaleNo").Value = Trim(.SaleNo)
            RS("Quantity").Value = Val(.Quantity)
            RS("Style").Value = Trim(.Style)
            RS("Vendor").Value = Trim(.Vendor)
            RS("Desc").Value = Trim(.Desc)
            RS("PorD").Value = Trim(.PorD)
            RS("DeptNo").Value = Left(.Code, 1)
            RS("VendorNo").Value = Mid(.Code, 2)
            'rs("Code") = .Code
            RS("Commission").Value = .Commission
            RS("Cost").Value = GetPrice(.Cost)
            RS("ItemFreight").Value = GetPrice(.ItemFreight)
            RS("SellPrice").Value = GetPrice(.SellPrice)
            RS("Salesman").Value = .Salesman
            RS("Status").Value = Left(Trim(.Status), 5)
            RS("Location").Value = GetPrice(.Location)
            RS("SellDate").Value = .SellDte
            RS("DelDate").Value = .DDelDat
            RS("Rn").Value = .RN
            RS("Store").Value = GetPrice(.Store)
            RS("Name").Value = Trim(.Name)
            RS("ShipDate").Value = .ShipDte
            RS("GM").Value = .GM
            RS("Tele").Value = CleanAni(.Phone)
            RS("MailIndex").Value = .Index
            RS("Detail").Value = GetPrice(.Detail)
        End With
    End Sub

    Private Sub cDataConvert_FileClose()
        FileClose(FILE_Index)
    End Sub

    Private Sub cDataConvert_ConvertExceptions(RS As ADODB.Recordset)
        ' This gets passed a filtered recordset containing things that failed to insert.
        ' We're going to assume they're all data errors.  In this case, duplicates.

        ' Print a list of all the GM data, so the store can fix whatever's wrong.
        Dim DataArr() As Object, FieldNameArr() As String, FieldAlignArr() As Integer, FieldWidthArr() As Integer, I As Integer
        DataArr = RS.GetRows()
        ReDim FieldNameArr(RS.Fields.Count - 1)
        ReDim FieldAlignArr(RS.Fields.Count - 1)
        ReDim FieldWidthArr(RS.Fields.Count - 1)
        For I = 0 To RS.Fields.Count - 1
            Select Case RS.Fields(I).Name
                Case Else
                    FieldNameArr(I) = RS.Fields(I).Name
                    FieldAlignArr(I) = ContentAlignment.MiddleLeft
            End Select
        Next
        '  --> The below lines are commented because they are for printer to print hardcoded printing code.

        '        PrintArray(FieldNameArr, FieldAlignArr, DataArr, FieldWidthArr, "Store " & StoresSld & " GrossMargin records which could not be imported", "Failed GrossMargin import report." & vbCrLf & RS.RecordCount & " records affected.")
        'REPRINT:
        '        On Error GoTo PrintError
        '        Printer.EndDoc()
        '        MsgBox(RS.RecordCount & " GrossMargin records failed to import, and require manual correction." & vbCrLf & "All information pertaining to the items in question has been sent to your printer.", vbInformation, "Error")
        '        Exit Sub
        'PrintError:
        '        If MsgBox(RS.RecordCount & " GrossMargin records failed to import, and require manual correction." & vbCrLf & "The report containing all information pertaining to the items in question failed to print.  Try again?", vbCritical + vbYesNo, "Error") = vbYes Then
        '            Resume REPRINT
        '        End If
    End Sub

    Public Function DataAccess() As CDataAccess
        DataAccess = mDataAccess
    End Function

    Public Sub Dispose()
        On Error Resume Next
        mDataAccess.Dispose()
    End Sub

    Public Function cDataAccess_SuperClass() As CDataAccess
        cDataAccess_SuperClass = mDataAccess
    End Function

    Public Sub cDataAccess_SetRecordSet(RS As ADODB.Recordset)
        Dim ST As String ' shortened style..
        On Error Resume Next
        ST = Trim(Left(Style, Setup_2Data_StyleMaxLen))

        ' MarginLine can't be updated.
        If MarginLine <> 0 And MarginLine <> RS("MarginLine").Value Then
            'If IsNothing(MarginLine) Then
            '    RS("MarginLine").Value = -1
            'Else
            '    RS("MarginLine").Value = MarginLine
            'End If
            'Dim rsMax As New ADODB.Recordset
            'rsMax = GetRecordsetBySQL("Select max(MarginLine) from GrossMargin", True, GetDatabaseAtLocation)
            'If Not rsMax.EOF And Not rsMax.BOF Then
            '    RS("MarginLine").Value = rsMax(0).Value + 1
            'End If
        End If

        RS("SaleNo").Value = Trim(SaleNo)              ' SHOULD BE USED INSTEAD OF ID
        RS("Style").Value = ST

        RS("Vendor").Value = Trim(Left(Vendor, Setup_2Data_ManufMaxLen))
        RS("Desc").Value = Trim(Left(Desc, Setup_2Data_DescMaxLen))
        RS("Quantity").Value = Quantity
        RS("PorD").Value = Trim(PorD)
        RS("DeptNo").Value = DeptNo
        RS("VendorNo").Value = Format(VendorNo, "000")
        RS("Commission").Value = Commission
        RS("Cost").Value = Cost
        RS("ItemFreight").Value = ItemFreight
        RS("SellPrice").Value = SellPrice
        RS("Salesman").Value = Trim(Salesman)
        RS("Status").Value = Left(Trim(Status), 10)

        ' This will prevent SUB/TAX/etc from having locations and showing up in the reports..
        ' BFH20151210 - Jerry wanted Del/Lab/Stain in the reports...
        If Not IsItem(ST) And Not IsNote(ST) And Not IsDLS(ST) Then Location = 0
        RS("Location").Value = Location

        RS("SellDate").Value = SellDte
        If IsDate(DDelDat) Then
            RS("DelDate").Value = DDelDat
        Else
            RS("DelDate").Value = DBNull.Value
        End If
        RS("Rn").Value = RN
        RS("Store").Value = Store
        RS("Name").Value = Trim(Name)
        RS("ShipDate").Value = ShipDte
        RS("GM").Value = GM
        RS("Tele").Value = CleanAni(Phone)
        RS("MailIndex").Value = Index
        RS("Detail").Value = Detail
        'Access
        RS("SS").Value = SS
        RS("DelPrint").Value = DelPrint
        RS("PullPrint").Value = PullPrint
        RS("CommPd").Value = CommPd
        If IsNothing(CommPd) Then
            RS("CommPd").Value = DBNull.Value
        End If
        RS("Spiff").Value = Spiff
        RS("SalesSplit").Value = Trim(SalesSplit)

        If IsDate(StopStart) Then
            RS("StopStart").Value = Format(TimeValue(StopStart), "h:mm ampm")
        Else
            RS("StopStart").Value = ""
        End If
        If IsDate(StopEnd) Then
            RS("StopEnd").Value = Format(TimeValue(StopEnd), "h:mm ampm")
        Else
            RS("StopEnd").Value = ""
        End If

        RS("IsPackage").Value = IIf(IsPackage, 1, 0)
        RS("PackSell").Value = PackSell
        RS("PackSellGM").Value = Math.Round(PackSellGM, 2)
        RS("PackSaleGM").Value = Math.Round(PackSaleGM, 2)
        RS("TransID").Value = Trim(TransID)
    End Sub

    Public Sub cDataAccess_GetRecordSet(RS As ADODB.Recordset)
        On Error Resume Next
        MarginLine = RS("MarginLine").Value
        SaleNo = Trim(IfNullThenNilString(RS("SaleNo").Value))
        'If MarginLine = "53664" And SaleNo = "25775" Then Stop
        Quantity = RS("Quantity").Value
        Style = Trim(IfNullThenNilString(RS("Style").Value))
        Vendor = Trim(IfNullThenNilString(RS("Vendor").Value))
        Desc = Trim(IfNullThenNilString(RS("Desc").Value))
        PorD = Trim(IfNullThenNilString(RS("PorD").Value))
        DeptNo = RS("DeptNo").Value
        VendorNo = Format(RS("VendorNo").Value, "000")
        'Code = rs("Code")
        Commission = RS("Commission").Value
        Cost = GetPrice(RS("Cost").Value)
        ItemFreight = RS("ItemFreight").Value
        SellPrice = RS("SellPrice").Value
        Salesman = Trim(RS("Salesman").Value)
        Status = Left(GetField_BlankDefault(RS, "Status"), 10)  ' USE THIS TO convert a null to blank
        Location = RS("Location").Value
        SellDte = RS("SellDate").Value
        DDelDat = IfNullThenNilString(RS("DelDate").Value)
        RN = IfNullThenZero(RS("Rn").Value)
        Store = IfNullThenZero(RS("Store").Value)
        Name = Trim(IfNullThenNilString(RS("Name").Value))
        ShipDte = RS("ShipDate").Value
        GM = IfNullThenNilString(RS("GM").Value)
        Phone = CleanAni(IfNullThenNilString(RS("Tele").Value))
        Index = RS("MailIndex").Value
        Detail = RS("Detail").Value
        'Access
        SS = IfNullThenNilString(RS("SS").Value)
        DelPrint = IfNullThenNilString(RS("DelPrint").Value)
        PullPrint = IfNullThenNilString(RS("PullPrint").Value)
        CommPd = RS("CommPd").Value
        If IsNothing(RS("CommPd").Value) Then
            CommPd = Nothing    ' MJK 20131026
        End If
        If IsNothing(CommPd) Then CommPd = Nothing          ' This doesn't work; CommPd can't be null.
        Spiff = IfNullThenZeroCurrency(RS("Spiff").Value)
        SalesSplit = IfNullThenNilString(RS("SalesSplit").Value)

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

        IsPackage = Val(IfNullThenNilString(RS("IsPackage").Value)) <> 0
        PackSell = IfNullThenZeroCurrency(RS("PackSell").Value)
        PackSellGM = Math.Round(IfNullThenZeroDouble(RS("PackSellGM").Value), 2)
        PackSaleGM = Math.Round(IfNullThenZeroDouble(RS("PackSaleGM").Value), 2)
        TransID = IfNullThenNilString(RS("TransID").Value)
    End Sub

    Public Sub mDataAccess_RecordUpdated()
        MarginLine = mDataAccess.Value("MarginLine")
    End Sub

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
        Dim rsMax As New ADODB.Recordset
        rsMax = GetRecordsetBySQL("Select max(MarginLine) from GrossMargin", True, GetDatabaseAtLocation)
        If Not rsMax.EOF And Not rsMax.BOF Then
            MarginLine = rsMax(0).Value
        End If
        Save = True
        Exit Function

NoSave:
        Err.Clear()
        Save = False
    End Function

    Public Function Load(ByVal KeyVal As String, Optional ByRef KeyName As String = "") As Boolean
        ' Checks the database for a matching record.
        ' Returns True if the load was successful, false otherwise.
        ' If a record was found, also loads the data into this object.

        Load = False
        ' Search for the record
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
        If DataAccess.Records_Available Then
            cDataAccess_GetRecordSet(DataAccess.RS)
            Load = True
        End If
    End Function

    Public Function Void(ByRef VoidDate As Date) As Boolean
        Void = False
        If MarginLine = 0 Then Exit Function ' Can't void unsaved records.
        If Status Like "VD*" Or Status = "VOID" Then
            ' Already voided.
            Void = True
            Exit Function
        End If

        ' Perform the void..

        ' Return the item to stock.
        ReturnToStock(VoidDate, True)  ' We don't care if it fails.  This depends on Rn, Status, SellDte.

        ' Update the void indicators.
        If Status Like "x*" Then
            ' Don't change status..
        ElseIf Status = "" Then
            Status = "VOID"
        Else
            Status = "VD" & Left(Status, 3)
        End If

        ShipDte = VoidDate  ' Ship Date, not Sell Date!
        If Status <> "VDDEL" And Not IsDelivered(Status) Then
            ' Clear pickup/delivery schedule.
            DDelDat = ""
            PorD = ""
        End If

        Save()    ' It is very important that this function cause a save.

        Void = True
    End Function

    ' bfh20050721
    ' AllowAll option currently only used to onscreenreport..
    ' added to make it process SS and SO not received so we could do adjustments
    ' on delivered sales...
    ' often, in practice, they have been delivered without being receieved, so we
    ' wanted them all to go through this (because of the printing the hard copy)
    Public Function ReturnToStock(ByRef ReturnDate As Date, Optional ByRef AllowAll As Boolean = False) As Boolean
        Dim tS As String

        ReturnToStock = False
        tS = Trim(Status)
        If tS = "" Or Left(tS, 1) = "x" Then Exit Function ' Invalid or already voided
        If Not AllowAll And GetPrice(RN) = 0 Then Exit Function                       'S/S & S/O not received

        Dim InvData As CInvRec
        InvData = New CInvRec
        If InvData.Load(RN, "#Rn") Then
            If Quantity > 0 Then
                AddItemCost(Style, Location, Cost / Quantity, 0, ReturnDate, Quantity)
            End If
            InvData.ItemsSold(-Quantity, CDate(SellDte))    'Update the item's sales history.

            If IsIn(tS, "SO", "PO") Then
                If tS = "PO" Then InvData.PoSold = InvData.PoSold - Quantity
            ElseIf tS <> "LAW" Then 'Everything except LAW goes through here
                InvData.Available = InvData.Available + Quantity
                If IsDelivered(tS) Then
                    If IsIn(tS, "DELSO", "DELSOREC", "DELSOR") Then
                        Dim M As String
                        If Desc <> InvData.Desc Then
                            M = "The descriptions for the Delivered Special Order (DELSO) Item " & Style & vbCrLf
                            M = M & "does not match the main inventory description:" & vbCrLf
                            M = M & "Item Desc:" & vbTab & Desc & vbCrLf
                            M = M & "Original Desc: " & vbTab & InvData.Desc & vbCrLf2
                            M = M & "Was this piece the same as the standard pieces?"
                            If MsgBox(M, vbQuestion + vbYesNo, "Descriptions are Different") = vbYes Then
                                InvData.OnHand = InvData.OnHand + Quantity
                            Else
                                MsgBox("Item not added to " & Style & vbCrLf & "Returned item must be manually entered into system.", vbInformation, "Quantity Not Updated")
                                PrintHardCopyOfDeliveredReturn()
                            End If
                        Else
                            InvData.OnHand = InvData.OnHand + Quantity
                        End If
                    ElseIf IsIn(tS, "DELFND", "DELFN") Then
                        ' THIS ONE WILL ONLY BE HIT IF THE FND ITEM IS ACTUALLY A REAL STYLE NUMBER...  SEE BELOW FOR USUAL CASE
                        MsgBox("Returned FND items are not automatically added back to stock." & vbCrLf & "It must be manually entered into the system to properly reflect quantity changes.")
                        PrintHardCopyOfDeliveredReturn()
                    ElseIf IsIn(tS, "DELSS", "DELSSR", "DELSSREC") Then
                        InvData.OnHand = InvData.OnHand + Quantity
                        If MsgBox("The item " & Style & " was a Special Special piece." & vbCrLf & "We recorded a cost of " & InvData.Cost & " for this style." & vbCrLf & "Is this correct?", vbQuestion + vbYesNo, "Verify Special Special Pricing") = vbNo Then
                            MsgBox("Please manually update the pricing for this item.")
                            PrintHardCopyOfDeliveredReturn()
                        End If
                    Else      ' item isn't a 'DELSO', 'DELFND'/'DELFN', 'DELSS'
                        InvData.OnHand = InvData.OnHand + Quantity
                    End If
                End If

                InvData.AddLocationQuantity(Location, Quantity)
            End If
            InvData.Save()

            If Detail > 0 Then
                Dim InvDetail As CInventoryDetail
                InvDetail = New CInventoryDetail
                If Not InvDetail.Load(CStr(Detail), "#DetailID") Then
                    MsgBox("Error returning " & Style & " to stock: Can't load Detail record #" & Detail & ".", vbCritical, "Error!")
                Else
                    ' Deletes the line entry
                    InvDetail.Trans = "VD"
                    InvDetail.Save()

                    If Trim(Status) = "POREC" Then
                        'add back quantity to detail
                        ' This creates a second detail line!
                        InvDetail.DataAccess.Records_Add()
                        InvDetail.Style = Style
                        InvDetail.DDate1 = ReturnDate
                        InvDetail.Misc = SaleNo & " VD"
                        InvDetail.Trans = "IN"
                        InvDetail.AmtS1 = 0
                        InvDetail.Ns1 = Quantity
                        InvDetail.SO1 = 0
                        InvDetail.LAW = 0
                        InvDetail.SetLocationQuantity(Location, Quantity)
                        InvDetail.ItemCost = Cost
                        InvDetail.Save()  ' Handles DetailID automatically.  We don't care what the new detail record number is.
                    End If
                End If
                DisposeDA(InvDetail)
            End If

            ' Print a stock tag?
            '  This gets called last so it won't matter if something breaks in the middle.
            '  It should only prompt if an item was actually returned to stock.
            If Location >= 1 Then
                If IsIn(Trim(Status), "ST", "SOREC", "POREC") Or IsDelivered(Status) Then
                    If MsgBox("Make New Ticket For Style No: " & InvData.Style, vbYesNo + vbQuestion) = vbYes Then
                        With InvData
                            SelectPrinter.PrintTags(.Style, .Desc, .Landed, .List, .OnSale, .DeptNo, .DeptNo & .VendorNo, .Vendor, .Available, .Comments)
                        End With
                        If SelectPrinter.SmallTags Then ' small tag was printed
                            Printer.EndDoc()
                            SelectPrinter.SmallTags = False
                        End If
                    End If
                End If
            End If
        Else
            ' Special-Special items can't be loaded, of course.
            If IsIn(tS, "SSREC", "DELSS", "DELSSREC", "DELSSR") Then
                MsgBox("You must manually add this Special Special received item into stock!", vbExclamation)
                PrintHardCopyOfDeliveredReturn()
            ElseIf IsIn(tS, "DELFN", "DELFND") Then
                MsgBox("Returned FND items are not automatically added back to stock." & vbCrLf & "It must be manually entered into the system to properly reflect quantity changes." & vbCrLf & "Printing reminder notice...", vbInformation, "Quantity not updated")
                PrintHardCopyOfDeliveredReturn()
            ElseIf IsIn(tS, "DELSO", "DELSOR", "DELSOREC") Then
                ' it's a little odd for this one to hit...  SOs should be already in the 2data table (and
                ' hence have a valid rn number which this is the failure to load clause of)...
                ' but, if we can't find the item in the 2data table, we certainly can't update it..
                ' we display a message and print the reminder notice
                MsgBox("This item could not be updated automatically because the RN number could not be located." & vbCrLf & "Its quantity must be updated manually to preserve an accurate inventory." & vbCrLf & "Printing reminder notice...", vbInformation, "Quantity not updated")
                PrintHardCopyOfDeliveredReturn()
            End If
        End If
        DisposeDA(InvData)

        Save()    ' This function needs to save, or 2Data/Detail will be corrupted.
    End Function

    Public Sub PrintHardCopyOfDeliveredReturn()
        ' --> Printer code. Not required. Reports are developed using reporting software.

        '        Dim Letter As String

        '        On Error GoTo PrinterError

        '        OutputToPrinter = True
        '        OutputObject = printer
        '        printer.FontSize = 14
        '        Printer.FontName = "Courier New"

        '        Printer.FontBold = True
        '        PrintCentered(StoreSettings.Name, , True)
        '        PrintCentered "Location " & StoresSld
        '  Printer.FontBold = False

        '        Printer.Print()
        '        Printer.Print()
        '        PrintAligned "Date:            " & Date
        '  PrintAligned "Sale No.:        " & SaleNo
        '  PrintAligned "Sell Date:       " & SellDte
        '  PrintAligned "Sale Name:       " & Name
        '  PrintAligned "Style:           " & Style
        '  PrintAligned "Desc:            " & Desc
        '  PrintAligned "Quantity:        " & Quantity
        '  PrintAligned "Original Status: " & Status
        '  Printer.Print()

        '        Select Case Status
        '            Case "DELFND", "DELFN", "DELFD", "DELF"
        '                PrintAligned "The actual style number for this item, '" & Status & "',"
        '      PrintAligned "was never located."
        '      PrintAligned "Either find the correct style number for this item and add it back"
        '      PrintAligned "to that item, create a new style number for this item, or do not"
        '      PrintAligned "track this item."
        '    Case "DELSO", "DELSOR", "DELSOREC"
        '                PrintAligned "The quantity of this item could not be automatically updated."
        '      PrintAligned "Either the item was not found in the inventory or the"
        '      PrintAligned "special order description did not match the inventory record"
        '      PrintAligned "description (usually indicating a different fabric or other"
        '      PrintAligned "customization)."
        '    Case "DELSS", "DELSSR", "DELSSREC"
        '                PrintAligned "This item was a Special/Special."
        '      PrintAligned "The style number is not automatically entered for this type of item"
        '      PrintAligned "so there is not inventory record to update to track the quantity"
        '      PrintAligned "of this item.  You can either create a new style for this piece,"
        '      PrintAligned "add it to an appropriate existing piece, or not track this item."
        '  End Select
        '        Printer.EndDoc()
        '        Exit Sub

        'PrinterError:
    End Sub

    Public Function WrittenInPeriod(ByRef StartDate As Date, ByRef EndDate As Date) As Boolean
        WrittenInPeriod = DateInRange(SellDte, StartDate, EndDate) And Trim(Style) <> "" 'And Trim(Style) <> "NOTES" And Trim(Style) <> "STAIN"
    End Function

    Public Function DeliveredInPeriod(ByRef StartDate As Date, ByRef EndDate As Date) As Boolean
        DeliveredInPeriod = DateInRange(DDelDat, StartDate, EndDate) And (IsDelivered(Status) Or Status = "VDDEL")
    End Function

    Public Function VoidedInPeriod(ByRef StartDate As Date, ByRef EndDate As Date) As Boolean
        VoidedInPeriod = DateInRange(ShipDte, StartDate, EndDate) And IsVoid(Status)
    End Function

    Public Function LoadMailRecord() As clsMailRec
        LoadMailRecord = modMail.LoadMailRecord(Index, GetStoreNumber(DataAccess.DataBase))
    End Function

    Public Function LoadHoldingRecord() As cHolding
        LoadHoldingRecord = New cHolding
        LoadHoldingRecord.DataAccess.DataBase = DataAccess.DataBase
        If Not LoadHoldingRecord.Load(SaleNo, "LeaseNo") Then DisposeDA(LoadHoldingRecord)
    End Function

    Public Function LoadInventoryRecord() As CInvRec
        LoadInventoryRecord = Nothing
        If Not IsItem(Style) Then Exit Function
        LoadInventoryRecord = New CInvRec
        If Not LoadInventoryRecord.Load(Style, "Style") Then
            DisposeDA(LoadInventoryRecord)
        End If
    End Function

    Public Function LoadInstallmentTransactionRecords() As cTransaction
        LoadInstallmentTransactionRecords = New cTransaction
        If Not LoadInstallmentTransactionRecords.Load("NewSale " & SaleNo, "Type") Then DisposeDA(LoadInstallmentTransactionRecords)
    End Function

    Public Function LoadInstallmentRecord() As cInstallment
        Dim T As cTransaction, ArNo As String

        LoadInstallmentRecord = Nothing
        T = LoadInstallmentTransactionRecords()
        If T Is Nothing Then DisposeDA(T) : Exit Function

        ArNo = T.ArNo
        DisposeDA(T)

        LoadInstallmentRecord = New cInstallment
        If Not LoadInstallmentRecord.Load(ArNo, "ArNo") Then DisposeDA(LoadInstallmentRecord)
    End Function
End Class

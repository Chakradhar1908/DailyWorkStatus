Imports Microsoft.VisualBasic.Interaction
Imports VBRUN

Public Class InvPoPrint
    Public DebugPrintPO As String

    Dim LastPO As Integer

    Dim StoreLoc As Integer
    Public StoreName As String
    Public StoreAddress As String
    Public StoreCity As String
    Public StorePhone As String
    Public StoreShipTo As String
    Public StoreShipAdd As String
    Public StoreShipCity As String
    Public StoreShipPhone As String

    Dim ShowCost As Boolean
    Dim PrintedCost As Decimal, TotCost As Decimal, Cost As Decimal
    Dim Counter As Integer, II As Integer, I As Integer, Y As Integer, YY As Integer ' Counters, Fax position, etc...
    Dim TotCubes As Double

    Dim TName As String             ' Fax vendor name+address.. tricky to eliminate.
    Dim tAddress As String, tAddress2 As String, tAddress3 As String
    Dim tZip As String, tPhone As String, tFax As String

    Dim REPRINT As Boolean, EPrintReport As Boolean  ' Passed from EditPO

    ' Variables to kill!
    'Dim D(800, 5) As String * 16  '@NO-LINT 'tried to rem out and no help
    Dim D(800, 5) As String  '@NO-LINT 'tried to rem out and no help

    Private Structure FontObj
        Dim OpCode As Integer
        Dim RecLen As Integer
        <VBFixedString(20)> Dim Name As String
        Dim Size As Integer
        Dim Bold As Integer
        Dim Italic As Integer
        Dim Underline As Integer
        Dim Color As Integer
    End Structure

    Private Structure TextObj
        Dim OpCode As Integer
        Dim RecLen As Integer
        Dim X1 As Integer
        Dim Y1 As Integer
        Dim X2 As Integer
        Dim Y2 As Integer
        Dim Color As Integer
        Dim Flags As Integer
    End Structure

    Private Structure LineObj
        Dim OpCode As Integer
        Dim RecLen As Integer
        Dim X1 As Integer
        Dim Y1 As Integer
        Dim X2 As Integer
        Dim Y2 As Integer
        Dim Width As Integer
        Dim Style As Integer
        Dim Color As Integer
    End Structure

    Public Sub ReprintPO(ByVal FromPO As String, ByVal ToPO As String, Optional ByVal PrintReport As Boolean = False, Optional ByVal OnlyUnprinted As Boolean = False)
        ' Only called by EditPO.cmdPrint_Click.
        EPrintReport = PrintReport
        Dim PO As New cPODetail
        PO.DataAccess.Records_OpenSQL("SELECT * FROM [" & PODetail_TABLE & "] WHERE PoNo>=" & Val(FromPO) & " AND PoNo<=" & Val(ToPO) & " ORDER BY PoNo, POID")

        Do While PO.DataAccess.Records_Available
            If PO.PrintPo = "V" Then GoTo SkipIt
            If PO.PrintPo = "X" And OnlyUnprinted Then
                If IsFormLoaded("frmEmail") Then frmEmail.LOG("PO #" & PO.PoNo & " Already Printed.")
                GoTo SkipIt
            End If
            DebugPrintPO = "InvPrintPO.ReprintPO"
            LastPO = PO.PoNo
            REPRINT = True
            On Error Resume Next
            txtDate.Value = Today
            txtDate.Value = PO.PoDate  'bfh20050502 - wasn't putting the right date on the po
            On Error GoTo 0
            PrintPo(PO)
SkipIt:
        Loop
        DisposeDA(PO)
        REPRINT = False
        EPrintReport = False
    End Sub

    Private Sub PrintPo(ByRef PO As cPODetail)
        Dim PrintWithCost As Boolean
        Dim Fax As Boolean, FH As Integer, FFile As String         ' for faxes..
        Dim Email As Boolean, Mail As String                    ' for emails..
        On Error GoTo HandleErr
        DebugPrintPO = "InvPrintPO.PrintPo-1"

        If Trim(PO.PrintPo) = "V" Then Exit Sub  'Won't print voided POs

        PrintWithCost = (PO.wCost = "1" And Not StoreSettings.bPrintPoNoCost)
        Fax = InvenMode("FPO")
        Email = InvenMode("EPO")

        '  PrintWithCost = True
        '  Email = True

        TotCost = 0
        TotCubes = 0
        Counter = 0

        DebugPrintPO = "InvPrintPO.PrintPo-2"
        GetLocation(PO)

        If Fax Then
            '    FH = FreeFile
            '    FFile = FaxPo.FaxFileName(PO.PoNo)
            '    Open FFile For Binary Access Write As #FH
            '
            '    FaxHeading PO, PrintWithCost, FH
            '    YY = 7450 '1st Line
            '    FaxLineItems PO, PrintWithCost, FH
        ElseIf Email Then
            Mail = ""
            EmailHeading(PO, PrintWithCost, Mail)
            EmailLineItems(PO, PrintWithCost, Mail)
        Else
            Heading(PO, PrintWithCost)
            LineItems(PO, PrintWithCost)
        End If

        DebugPrintPO = "InvPrintPO.PrintPo-3"

        Do While PO.DataAccess.Records_Available
            If (Trim(PO.PoNo) <> LastPO) Then Exit Do

            If Fax Then
                '      FaxLineItems PO, PrintWithCost, FH
            ElseIf Email Then
                EmailLineItems(PO, PrintWithCost, Mail)
                If SaveEmailToDesktop Then WriteFile(UIOutputFolder() & "mail.html", Mail)
            Else
                LineItems(PO, PrintWithCost)
            End If
        Loop
        If Not PO.DataAccess.Record_BOF Then PO.DataAccess.Records_MovePrevious()

        If Fax Then
            '    FaxTotalPo PO, PrintWithCost, FH
            '    Close #FH
            '    MakeFaxlist
        ElseIf Email Then
            EmailTotalPo(PO, PrintWithCost, Mail)
            Mail = Mail & "" ' break line
            If SaveEmailToDesktop Then WriteFile(UIOutputFolder() & "mail.html", Mail, True)
            frmEmail.EmailPO(PO, Mail)
        Else
            TotalPo(PO, PrintWithCost)
        End If

        DebugPrintPO = "InvPrintPO.PrintPo-4"
        If InvPo.PoNo = PO.PoNo Then InvPo.PoNo = 0  ' This PO is no longer able to be added on to.
        Cost = 0
        TotCost = 0
        '  If Not PO.DataAccess.Record_EOF Then PO.DataAccess.Records_MovePrevious
        Printer.EndDoc()
        Exit Sub

HandleErr:
        Dim M As String
        M = "Error while " & Switch(Fax, "creating PO faxes", Email, "emailing POs", True, "printing POs") & "."
        M = M & vbCrLf & "Error #: " & Err.Number & ": " & Err.Description
        If DebugPrintPO <> "" Then M = M & vbCrLf2 & "DEBUG: " & DebugPrintPO
        M = M & vbCrLf & "OutputObject=" & DescribeOutputObject()

        MessageBox.Show(M, "Error Printing POs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Resume Next
    End Sub

    Private Sub GetLocation(ByRef PO As cPODetail)
        If Val(PO.Location) = 0 Then PO.Location = 1
        StoreLoc = Val(PO.Location)
        If StoreLoc <= 0 Then StoreLoc = 1
        StoreName = StoreSettings(StoreLoc).Name
        StoreAddress = StoreSettings(StoreLoc).Address
        StoreCity = StoreSettings(StoreLoc).City
        StorePhone = StoreSettings(StoreLoc).Phone

        StoreShipTo = StoreSettings(StoreLoc).StoreShipToName
        StoreShipAdd = StoreSettings(StoreLoc).StoreShipToAddr
        StoreShipCity = StoreSettings(StoreLoc).StoreShipToCity
        StoreShipPhone = StoreSettings(StoreLoc).StoreShipToTele
        OpenApDatabase(PO.Location, True)
    End Sub

    Private Sub EmailHeading(ByRef PO As cPODetail, ByVal PrintWithCost As Boolean, ByRef M As String)
        Dim N As String
        Dim A As String, B As String, C As String, D As String
        Const STRIKE As String = "" ' This didn't look right.. maybe something else..  "text-decoration: line-through;"
        Const Blnk As String = "_____"
        Const FILL As String = "__X__"
        Dim tST As Boolean, TstX As String


        VendorAddress(PO.Vendor, False)
        N = vbCrLf
        M = M & N & "<table class='PO' cell-padding='0' cell-spacing='0' border='0' style='font-family:arial;width:90%;'>"
        M = M & N & "  <tr><td>"
        M = M & N & "    <table width='100%'>"
        M = M & N & "      <tr>"
        M = M & N & "        <td style='border: black 2px solid;padding: 10px 10px 10px 10px;'>" ' SOLD TO

        Select Case PO.SoldTo
            Case 2
                A = StoreShipTo
                B = StoreShipAdd
                C = StoreShipCity
                D = StoreShipPhone
            Case Else
                A = StoreName
                B = StoreAddress
                C = StoreCity
                D = StorePhone
        End Select
        ' Sold to
        M = M & N & "<span class='SOLDTO' style='font-weight:bold;'>"
        M = M & N & "<span class='HEADTITLE' style='font-weight:normal'>SOLD TO:</span><br>"
        M = M & N & "<span class='ADDRINFO' style='text-indent:10px'>" & A & "</span><br>"
        M = M & N & "<span class='ADDRINFO' style='text-indent:10px'>" & B & "</span><br>"
        M = M & N & "<span class='ADDRINFO' style='text-indent:10px'>" & C & "</span><br>"
        M = M & N & "<span class='ADDRINFO' style='text-indent:10px'>" & D & "</span><br>"
        M = M & N & "</span>"    ' end SOLDTO

        M = M & N & "        </td>"
        M = M & N & "        <td>" ' PO Title

        ' Title
        M = M & N & "<span class='POTITLE' style='display:inline-block;float:left;clear:left;padding-top:30px;font-weight:bold;width:100%;height:auto;text-align:center;'>"
        M = M & N & "Purchase Order (Location " & PO.Location & ")"
        M = M & N & "</span>" ' end POTITLE

        M = M & N & "        </td>"
        M = M & N & "        <td style='border: black 2px solid;'>" ' SHIP TO

        Select Case PO.ShipTo
            Case 2
                A = StoreShipTo
                B = StoreShipAdd
                C = StoreShipCity
                D = StoreShipPhone
            Case 3
                A = PO.ShiptoName
                B = PO.ShipToAddress
                C = PO.ShipToCity
                D = PO.ShipToTele
            Case Else
                A = StoreName
                B = StoreAddress
                C = StoreCity
                D = StorePhone
        End Select
        ' Ship to
        M = M & N & "<span class='SHIPTO' style='font-weight:bold;padding: 10px 10px 10px 10px;'>"
        M = M & N & "<span class='HEADTITLE' style='font-weight:normal'>SHIP TO:</span><br>"
        M = M & N & "<span class='ADDRINFO' style='text-indent:10px'>" & A & "</span><br>"
        M = M & N & "<span class='ADDRINFO' style='text-indent:10px'>" & B & "</span><br>"
        M = M & N & "<span class='ADDRINFO' style='text-indent:10px'>" & C & "</span><br>"
        M = M & N & "<span class='ADDRINFO' style='text-indent:10px'>" & D & "</span><br>"
        M = M & N & "</span>"    ' end SHIPTO

        M = M & N & "      </tr>"
        M = M & N & "    </table>"
        M = M & N & "  </td></tr>"

        M = M & N & "  <tr><td>" ' Vendor Address

        'Vendor Name
        M = M & N & "<span class='VENDOR' STYLE='display:inline-block;float:left;clear:left;padding-top:30px;font-weight:bold;'>"
        M = M & N & "<span class='ADDRINFO'>" & PO.Vendor & "</span><br>"
        M = M & N & "<span class='ADDRINFO'>" & tAddress & "</span><br>"
        M = M & N & "<span class='ADDRINFO'>" & tAddress2 & " " & tZip & "</span><br>"
        M = M & N & "<span class='ADDRINFO'>" & tPhone & "   Fax: " & tFax & "</span><br>"
        M = M & N & "</span>"    ' end VENDOR

        M = M & N & "  </td></tr>"

        M = M & N & "  <tr><td>" ' Spec Instr

        ' Special Instructions
        If StoreSettings.bPOSpecialInstr Then
            M = M & N & "<span class='SPIN' style='display:inline-block;float:left;clear:left;padding-top:30px;font-weight:bold;'>"
            M = M & N & "<span class='HEADTITLE' style='font-weight:normal'>**** SPECIAL INSTRUCTIONS ****</span><br>"
            If IsParkPlace Then
                tST = PO.Note1 = "1"
                M = M & N & "<span class='SPINLN' style='text-indent:10px;" & IIf(tST, "", STRIKE) & "'>" & IIf(tST, FILL, Blnk) & " If order is less than 90 lbs., Hold and SHIP with other goods.</span><br>"
            Else
                M = M & N & "<span class='SPINLN' style='text-indent:10px;" & IIf(tST, "", STRIKE) & "'>" & IIf(tST, FILL, Blnk) & " " & StoreSettings.PoSpecInstr1 & "</span><br>"
            End If

            tST = PO.Note2 = "1"
            M = M & N & "<span class='SPINLN' style='text-indent:10px;" & IIf(tST, "", STRIKE) & "'>" & IIf(tST, FILL, Blnk) & " " & StoreSettings.PoSpecInstr2 & "</span><br>"
            tST = PO.Note3 = "1"
            M = M & N & "<span class='SPINLN' style='text-indent:10px;" & IIf(tST, "", STRIKE) & "'>" & IIf(tST, FILL, Blnk) & " " & StoreSettings.PoSpecInstr3 & "</span><br>"
            Dim FF As String
            FF = PO.PoNotes
            If Len(FF) = 0 Then FF = StoreSettings.PoSpecInstr4
            tST = PO.Note4 = "1"
            M = M & N & "<span class='SPINLN' style='text-indent:10px;" & IIf(tST, "", STRIKE) & "'>" & IIf(tST, FILL, Blnk) & " " & IIf(Trim(FF) <> "", "<u>" & FF & "</u>", "_______________________________") & "</span><br>"
            M = M & N & "</span>"    ' end SPIN
        End If

        M = M & N & "  </td></tr>"

        M = M & N & "  <tr><td>" ' Final Instr

        ' Final Instructions
        M = M & N & "<span class='FINALHEAD' style='display:inline-block;float:left;clear:left;padding-top:30px;text-align:center;'>"
        M = M & N & "<span class='FINSTR' style='text-align:center'>Please put our PO NUMBER, ORDER NUMBER & TAG NAME on all correspondence!</span><br>"
        M = M & N & "<br>" ' extra line break here..
        ' no line breaks after these...
        M = M & N & "<span class='FINFO' style='font-weight:bold;font-size:28px;'>PO Number: " & PO.PoNo & "</span>"
        M = M & N & "&nbsp;&nbsp;&nbsp;&nbsp;"
        M = M & N & "<span class='FINFO' style='font-weight:bold;'>Order Number: " & PO.SaleNo & "</span>"
        M = M & N & "&nbsp;&nbsp;&nbsp;&nbsp;"
        M = M & N & "<span class='FINFO' style='font-weight:bold;'>Date: " & txtDate.Value & "</span>"
        M = M & N & "&nbsp;&nbsp;&nbsp;&nbsp;"
        M = M & N & "<span class='FINFO' style='font-weight:bold;'>Tag: " & PO.Name & "</span>"
        M = M & N & "<br>" ' finally finish off these 4 elements w/ a line break
        M = M & N & "</span>"    ' end FINALHEAD

        M = M & N & "  </td></tr>"

        M = M & N & "  <tr><td>" ' PO BODY

        M = M & N & "<table cellpadding=0 cellspacing=0 border=1 width='100%'>"
        M = M & N & "<thead>"
        M = M & N & "<tr>"
        M = M & N & "<td><span class='POH' style='font-weight:bold;' width='10%'>QUAN</span></td>"
        M = M & N & "<td><span class='POH' style='font-weight:bold;' width='20%'>STYLE NO.</span></td>"
        M = M & N & "<td><span class='POH' style='font-weight:bold;' width='15%'>CUBES</span></td>"
        M = M & N & "<td><span class='POH' style='font-weight:bold;' width='35%'>DESCRIPTION</span></td>"
        If PrintWithCost Then
            M = M & N & "<td><span class='POH' style='font-weight:bold;' width='20%'>COST</span></td>"
        End If

        M = M & N & "</tr>"
        M = M & N & "</thead>"
        M = M & N & "<br/><br/><br/>"

    End Sub

    Private Sub EmailLineItems(ByRef PO As cPODetail, ByVal PrintWithCost As Boolean, ByRef M As String)
        Dim N As String
        Dim tD As String, FS As Integer, FSS As String
        FS = 11
        FSS = "font-size:" & FS & ";"
        If Not REPRINT And PO.PrintPo = "v" Then Exit Sub

        PrintedCost = IIf(PrintWithCost, PO.Cost, "0")

        N = vbCrLf
        ' goes on line items
        M = M & N & "<tr>"
        M = M & N & "<td><span class='LIQUAN' style='text-align:right;" & FSS & "'>" & PO.Quantity & "</span>"
        M = M & N & "<td><span class='LISTYL' style='text-align:left;" & FSS & "'>" & PO.Style & "</span>"

        Dim R As CInvRec
        R = New CInvRec
        R.Load(PO.Style, "Style")
        TotCubes = TotCubes + (R.Cubes * PO.Quantity)
        M = M & N & "<td><span class='LICUBE' style='text-align:left;" & FSS & "'>" & CurrencyFormat(R.Cubes * PO.Quantity) & "</span>"
        DisposeDA(R)

        tD = PO.Desc
        If PO.PrintPo = "v" Then tD = "   <b><s>VOID VOID VOID</s></b>"

        M = M & N & "<td><span class='LIDESC' style='text-align:left;" & FSS & "'>" & tD & "</span>"

        If PrintWithCost And Trim(PO.PrintPo) <> "V" And PO.PrintPo <> "v" Then
            M = M & N & "<td><span class='LICOST' style='text-align:right;" & FSS & "'>" & PriceFormatFunc(PrintedCost) & "</span>"
            Cost = PO.Cost
            TotCost = TotCost + Cost
        End If

        M = M & N & "</tr>"

        If Not EPrintReport Then
            If PO.PrintPo <> "v" Then
                PO.PrintPo = "X" 'only for first print
                PO.Save()
            End If
        End If

        DisposeDA(R)
    End Sub

    Private Sub Heading(ByRef PO As cPODetail, Optional ByRef PrintWithCost As Boolean = False)
        OutputObject.FontName = "Arial"
        OutputObject.CurrentX = 0
        OutputObject.CurrentY = 200
        OutputObject.FontSize = 13
        OutputObject.FontBold = True

        DebugPrintPO = "Header-x"
        PrintCentered("Purchase Order (Location " & PO.Location & ")")
        DebugPrintPO = "Header-a"
        OutputObject.FontBold = False
        DebugPrintPO = "Header-1"
        Printer.CurrentX = 9500
        DebugPrintPO = "Header-2"
        Printer.CurrentY = 200
        DebugPrintPO = "Header-3"
        Printer.FontSize = 9
        'If OutputObject.page > 1 Then Printer.Print " Page: "; OutputObject.page
        DebugPrintPO = "Header-4"
        Printer.Print(" Page: ", OutputObject.Page)
        DebugPrintPO = "Header-5"
        OutputObject.FontSize = 13

        DebugPrintPO = "Header-6"
        Printer.CurrentY = 700
        DebugPrintPO = "Header-7"
        Printer.Print(TAB(8), "SOLD TO:", TAB(60))

        DebugPrintPO = "Header-8"
        OutputObject.FontBold = True
        DebugPrintPO = "Header-9"
        OutputObject.Print("SHIP TO:")
        DebugPrintPO = "Header-10"
        OutputObject.FontBold = False
        DebugPrintPO = "Header-11"
        Printer.CurrentY = 1200

        DebugPrintPO = "Printing Addresses"
        Addresses(PO)
        ' when you need a different Bill To number
        If IsSleepCity() Then
            StoreName = "Sleep City Warehouse"
            StoreAddress = " 180 Professional Center Dr."
            StoreCity = "Rohnert Park, Ca 94928-2144"
            StorePhone = "(707) 584-1382 Fax 584-3653"
        ElseIf IsRoughingItInStyle() Then
            StoreName = "Roughing It In Style"
            StoreAddress = "5262 Verona Rd. "
            StoreCity = "Madison, WI 53711"
            StorePhone = "(608)274-5559  Fax 274-5558"
        ElseIf IsBilliardsAndBarstools() Then
            StoreName = "Billiards & Barstools"
            StoreAddress = "3333 Country Club Dr. "
            StoreCity = "Glendale, CA  91208"
            StorePhone = "(818) 957-3514  Fax 957-0139 "
        End If

        'BILL TO                                       SHIP TO
        OutputObject.Print(TAB(10), StoreName, TAB(65), StoreShipTo)
        OutputObject.Print(TAB(10), StoreAddress, TAB(65), StoreShipAdd)
        OutputObject.Print(TAB(10), StoreCity, TAB(65), StoreShipCity)
        OutputObject.Print(TAB(10), StorePhone, TAB(65), StoreShipPhone)
        OutputObject.Print(vbCrLf)

        DebugPrintPO = "Vendor Address"
        VendorAddress(PO.Vendor)

        OutputObject.Print(vbCrLf)
        OutputObject.Print(TAB(10)) '"    FAX NO: "
        OutputObject.Print(TAB(10)) '; "CONTROL NO: "

        DebugPrintPO = "Spec Instr"
        If StoreSettings.bPOSpecialInstr Then
            OutputObject.FontSize = 12
            OutputObject.CurrentX = 5900
            OutputObject.CurrentY = 3000
            OutputObject.FontBold = True
            OutputObject.Print(" **** SPECIAL INSTRUCTIONS ****")
            OutputObject.FontBold = False
            OutputObject.CurrentY = 3300

            PrintSpecInstr(1, PO.Note1 = "1") ' bfh20051115 - fixed bug in this.. printed 1st line multiple times
            PrintSpecInstr(2, PO.Note2 = "1")
            PrintSpecInstr(3, PO.Note3 = "1")
            PrintSpecInstr(4, PO.Note4 = "1", PO.PoNotes)

            '      If PO.Note1 = "1" Then
            '        OutputObject.CurrentY = 3300
            '        OutputObject.CurrentX = 7000
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print " X "
            '        OutputObject.FontUnderline = False
            '
            '        OutputObject.CurrentX = 7500
            '        OutputObject.CurrentY = 3300
            '      Else
            '        OutputObject.CurrentY = 3300
            '        OutputObject.CurrentX = 7000
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print "   "
            '        OutputObject.FontUnderline = False
            '
            '        OutputObject.CurrentX = 7000
            '        OutputObject.CurrentY = 3300
            '      End If
            '
            '      Dim TStr As String, TStrs() As String, TI as integer
            '      OutputObject.CurrentX = 7500
            '      If isparkplace Then
            '        TStr = "If order is less than 90 lbs., HOLD and SHIP with other goods."
            '      Else
            '        TStr = frmSetup OutputObject.GetPOSpecInstr(1)
            '      End If
            '
            '      TStrs = SplitLongText(TStr, 35)
            '      For TI = LBound(TStrs) To UBound(TStrs)
            '        If I = 2 Then Exit For
            '        OutputObject.CurrentX = 7500
            '        OutputObject.Print TStrs(I)
            '      Next
            '
            '
            '      If PO.Note2 = "1" Then
            '        OutputObject.CurrentY = 3900
            '        OutputObject.CurrentX = 7000
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print " X "
            '        OutputObject.FontUnderline = False
            '        OutputObject.CurrentX = 7000
            '        OutputObject.CurrentY = 3900
            '      Else
            '        OutputObject.CurrentY = 3900
            '        OutputObject.CurrentX = 7000
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print "   "
            '        OutputObject.FontUnderline = False
            '        OutputObject.CurrentX = 7000
            '        OutputObject.CurrentY = 3900
            '      End If
            '      OutputObject.CurrentX = 7500
            '      OutputObject.Print "Sold orders:  Ship Complete Only"
            '
            '      If PO.Note3 = "1" Then
            '        OutputObject.CurrentY = 4200
            '        OutputObject.CurrentX = 7000
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print " X "
            '        OutputObject.FontUnderline = False
            '        OutputObject.CurrentX = 7000
            '        OutputObject.CurrentY = 4200
            '      Else
            '        OutputObject.CurrentY = 4200
            '        OutputObject.CurrentX = 7000
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print "   "
            '        OutputObject.FontUnderline = False
            '        OutputObject.CurrentX = 7000
            '        OutputObject.CurrentY = 4200
            '      End If
            '      OutputObject.CurrentX = 7500
            '      OutputObject.Print "Ship UPS, PP or With Other Goods"
            '
            '      If PO.Note4 = "1" Then
            '        OutputObject.CurrentY = 4500
            '        OutputObject.CurrentX = 7000
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print " X "
            '        OutputObject.FontUnderline = False
            '        OutputObject.CurrentX = 7400
            '        OutputObject.CurrentY = 4500
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print PO.PoNotes
            '        OutputObject.FontUnderline = False
            '      Else
            '        OutputObject.CurrentY = 4500
            '        OutputObject.CurrentX = 7000
            '        OutputObject.FontUnderline = True
            '        OutputObject.Print "   "
            '        OutputObject.FontUnderline = False
            '        OutputObject.CurrentY = 4500
            '        OutputObject.CurrentX = 7450
            '        OutputObject.Print "_______________________________"
            '      End If
        End If

        DebugPrintPO = "Please put..."
        OutputObject.CurrentX = 0
        OutputObject.CurrentY = 5000
        OutputObject.FontSize = 12
        'OutputObject.Print
        OutputObject.Print(vbCrLf)
        OutputObject.Print(TAB(10), "Please put our PO NUMBER, ORDER NUMBER & TAG NAME on all correspondence!")
        OutputObject.Print

        OutputObject.FontBold = True
        OutputObject.Print(TAB(1), "PO Number: ", PO.PoNo, TAB(24), "Order Number: ", PO.SaleNo, TAB(50), "Date: ", DateFormat(txtDate.Value), TAB(72), "TAG: ", PO.Name)

        DebugPrintPO = "Column Headers"
        OutputObject.Print
        OutputObject.FontUnderline = True
        'OutputObject.Print ; "QUAN.   STYLE NO.                   DESCRIPTION                                                                                    Cost "
        PrintTo(OutputObject, "QUAN.", 10, AlignConstants.vbAlignRight, False)
        PrintTo(OutputObject, "STYLE NO.", 12, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "DESCRIPTION", 42, AlignConstants.vbAlignLeft, False)
        PrintTo(OutputObject, "CUBES", 125, AlignConstants.vbAlignRight, False)
        If PrintWithCost Then
            PrintTo(OutputObject, "COST", 143, AlignConstants.vbAlignRight, False)
        End If
        DebugPrintPO = "Headerline"
        PrintTo(OutputObject, New String("_", (Printer.ScaleWidth - 200) / Printer.TextWidth("_")), 2, AlignConstants.vbAlignLeft, True)
        '    DrawLine 100, Printer.CurrentY, Printer.Width - 100, Printer.CurrentY

        OutputObject.FontUnderline = False
        OutputObject.FontBold = False
    End Sub

    Private Sub LineItems(ByRef PO As cPODetail, ByVal PrintWithCost As Boolean)
        Dim VoidLine As Boolean
        DebugPrintPO = "Line Items"
        If PO.wCost <> "1" Or StoreSettings.bPrintPoNoCost Then
            PrintedCost = 0
        Else
            PrintedCost = PO.Cost
        End If


        DebugPrintPO = "Voidline"
        VoidLine = False
        If Trim(PO.PrintPo) <> "V" Then   'voids
            Printer.FontSize = 11
            PrintTo(OutputObject, PO.InitialQuantity, 8, AlignConstants.vbAlignRight, False) ' was .Quantity
            If Trim(PO.PrintPo) = "v" Then
                OutputObject.FontBold = True
                OutputObject.Font.Strikethrough = True
                PrintTo(OutputObject, PO.Style, 12, AlignConstants.vbAlignLeft, False)
                OutputObject.Font.Strikethrough = False
                OutputObject.FontBold = False
                VoidLine = True
                PrintTo(OutputObject, "VOID", 42, AlignConstants.vbAlignLeft, False)
            Else
                PrintTo(OutputObject, PO.Style, 12, AlignConstants.vbAlignLeft, False)
            End If

            PrintTo(OutputObject, Microsoft.VisualBasic.Left(PO.Desc, 38), 42, AlignConstants.vbAlignLeft, False)

            DebugPrintPO = "InvRec Load"
            Dim R As CInvRec
            R = New CInvRec
            R.Load(PO.Style, "Style")
            TotCubes = TotCubes + (R.Cubes * PO.Quantity)
            PrintTo(OutputObject, CurrencyFormat(R.Cubes * PO.Quantity), 125, AlignConstants.vbAlignRight, False)
            DisposeDA(R)

            DebugPrintPO = "P Cost"
            'allow overwrite
            If PrintWithCost And Not VoidLine Then
                ' Only print Cost if the PO allows it.
                PrintTo(OutputObject, PriceFormatFunc(PrintedCost), 143, AlignConstants.vbAlignRight, True)
            Else
                ' Otherwise print a line terminator, or we'll have overprinting problems.
                OutputObject.Print
            End If

            DebugPrintPO = "P Desc"
            If Len(PO.Desc) > 38 Then
                Counter = Counter + 1
                PrintTo(OutputObject, Mid(PO.Desc, 39, 38), 50, AlignConstants.vbAlignLeft, True)
                If Len(PO.Desc) > 76 Then
                    Counter = Counter + 1
                    PrintTo(OutputObject, Mid(PO.Desc, 77, 38), 50, AlignConstants.vbAlignLeft, True)
                End If
                If Len(PO.Desc) > 114 Then
                    Counter = Counter + 1
                    PrintTo(OutputObject, Mid(PO.Desc, 115, 38), 50, AlignConstants.vbAlignLeft, True)
                End If
            End If

            DebugPrintPO = "P New Page"
            Printer.FontSize = 12
            Counter = Counter + 1
            If Counter >= 28 Then
                Printer.NewPage()
                Heading(PO)
                Counter = 0
            End If
        End If
        If Not EPrintReport Then
            'only for first print
            If PO.PrintPo <> "v" Then
                PO.PrintPo = "X"
                PO.Save()
            End If
        End If
        If Trim(PO.PrintPo) <> "V" And Trim(PO.PrintPo) <> "v" Then   'void entry
            Cost = PO.Cost
            TotCost = TotCost + Cost
        End If
    End Sub

    Private ReadOnly Property SaveEmailToDesktop() As Boolean
        Get
            SaveEmailToDesktop = IsDevelopment()
            '  SaveEmailToDesktop = False
        End Get
    End Property

    Private Sub EmailTotalPo(ByRef PO As cPODetail, ByVal PrintWithCost As Boolean, ByRef M As String)
        Dim N As String
        Dim T As String

        N = vbCrLf
        If PrintWithCost Then ' PrintedCost <> "0" Then  ' Caused some totals to not be printed.
            '    m = m & n & "<tr>"
            '    m = m & n & "<td></td><td></td><td></td>"
            '    m = m & n & "<td class='SUMLINE' style='font-weight:bold;'>___________</span>"
            '    m = m & n & "</tr>"
            M = M & N & "<tr>"
            M = M & N & "<td></td>"
            M = M & N & "<td></td>"
            M = M & N & "<td class='TOTCUBE' style='font-weight:bold;'>" & CurrencyFormat(TotCubes) & "</span>"
            M = M & N & "<td></td>"
            M = M & N & "<td class='TOTCOST' style='font-weight:bold;'>" & PriceFormatFunc(TotCost) & "</span>"
            M = M & N & "</tr>"
        End If
        M = M & N & "</table>"
        M = M & N & "<br>"

        If Trim(PO.SpecialNote) <> "" Then
            M = M & N & "<span class='SPECNOTE' style='display:inline-block;float:left;clear:left;padding-top:30px;font-weight:bold;'>"
            M = M & N & "SPECIAL:<br>" & PO.SpecialNote
            M = M & "</span><br><br><br>"
        End If

        If REPRINT Then
            M = M & N & "<span class='REPRINT' class='display:inline-block;float:left;clear:left;padding-top:30px;'>REPRINT!   REPRINT!   REPRINT!</span>"
        End If

        M = M & N & "<span class='AUTHBY' style='display:inline-block;float:left;clear:left;padding-top:30px;'>"
        M = M & N & "Authorized By: ________________________________"
        M = M & N & "</span>"

        M = M & N & "  </td></tr>"
        M = M & N & "</table>"

        M = M & N & "</span>" ' end POBODY
    End Sub

    Private Sub TotalPo(ByRef PO As cPODetail, ByVal PrintWithCost As Boolean)
        DebugPrintPO = "P TotalPO"
        If PrintWithCost Then ' PrintedCost <> "0" Then  ' Caused some totals to not be printed.
            PrintTo(OutputObject, New String("_", 8), 125, AlignConstants.vbAlignRight, False)
            OutputObject.Print(TAB(97), New String("_", 8))

            OutputObject.Print(TAB(60), "TOTAL:")
            PrintTo(OutputObject, CurrencyFormat(TotCubes), 125, AlignConstants.vbAlignRight, False)
            PrintTo(OutputObject, PriceFormatFunc(TotCost), 143, AlignConstants.vbAlignRight, True)
        End If

        DebugPrintPO = "SPECIAL"
        If Trim(PO.SpecialNote) <> "" Then
            OutputObject.Print
            OutputObject.Print
            OutputObject.CurrentX = 0
            OutputObject.FontSize = 14
            OutputObject.FontBold = True
            OutputObject.Print("SPECIAL:")
            OutputObject.FontSize = 12

            Dim SpInstLines() As String, Sp As String, I As Integer
            Sp = PO.SpecialNote
            Sp = WrapLongTextByPrintWidth(OutputObject, Sp, OutputObject.ScaleWidth, vbCrLf)
            SpInstLines = Split(Sp, vbCrLf)
            For I = LBound(SpInstLines) To UBound(SpInstLines)
                If I - LBound(SpInstLines) > 11 Then Exit For
                OutputObject.CurrentX = 400 ': OutputObject.CurrentY = OutputObject.CurrentY - 300
                Printer.Print(SpInstLines(I)) ' lblSpecial
            Next

            '      OutputObject.Print PO.SpecialNote  'special box
            OutputObject.FontBold = False
        End If

        DebugPrintPO = "REPRINT"
        If REPRINT Then
            OutputObject.CurrentX = 500
            OutputObject.CurrentY = 14500
            OutputObject.Print("REPRINT!   REPRINT!   REPRINT!")
        End If

        DebugPrintPO = "AUTH LINE"
        OutputObject.CurrentX = 5000
        OutputObject.CurrentY = 14500
        OutputObject.Print("Authorized By: ________________________________")
        Printer.EndDoc() 'Must be Printer
    End Sub

    Private Sub VendorAddress(ByVal Vendor As String, Optional ByVal DoOutput As Boolean = True)
        'Go to AP to get vendor Physical address

        TName = ""
        tAddress = ""
        tAddress2 = ""
        tZip = ""
        tPhone = ""
        tFax = ""

        If UseQB() Then
            QBGetVendorName(Vendor, TName, tAddress, tAddress2, tAddress3, tZip, tPhone, tFax)
        Else
            GetVendorName(Vendor, TName, tAddress, tAddress2, tAddress3, tZip, tPhone, tFax)
        End If

        If DoOutput Then
            If Trim(TName) = "" Then TName = Vendor
            OutputObject.Print(TAB(10), TName)
            OutputObject.Print(TAB(10), tAddress)
            OutputObject.Print(TAB(10), Trim(tAddress2 & " " & tZip))
            OutputObject.Print(TAB(10), PhoneAndFax(tPhone, tFax))
            OutputObject.Print(TAB(10), tAddress3)
        End If
    End Sub

    Private Sub Addresses(ByRef PO As cPODetail)
        Select Case PO.SoldTo
            Case "1"
                StoreName = StoreSettings(StoreLoc).Name
                StoreAddress = StoreSettings(StoreLoc).Address
                StoreCity = StoreSettings(StoreLoc).City
                StorePhone = StoreSettings(StoreLoc).Phone
            Case "2"
                StoreName = StoreSettings(StoreLoc).StoreShipToName
                StoreAddress = StoreSettings(StoreLoc).StoreShipToAddr
                StoreCity = StoreSettings(StoreLoc).StoreShipToCity
                StorePhone = StoreSettings(StoreLoc).StoreShipToTele
        End Select

        Select Case PO.ShipTo
            Case "2"
                StoreShipTo = StoreSettings(StoreLoc).StoreShipToName
                StoreShipAdd = StoreSettings(StoreLoc).StoreShipToAddr
                StoreShipCity = StoreSettings(StoreLoc).StoreShipToCity
                StoreShipPhone = StoreSettings(StoreLoc).StoreShipToTele
            Case "1"
                StoreShipTo = StoreSettings(StoreLoc).Name
                StoreShipAdd = StoreSettings(StoreLoc).Address
                StoreShipCity = StoreSettings(StoreLoc).City
                StoreShipPhone = StoreSettings(StoreLoc).Phone
            Case "3"
                StoreShipTo = PO.ShiptoName
                StoreShipAdd = PO.ShipToAddress
                StoreShipCity = PO.ShipToCity
                StoreShipPhone = PO.ShipToTele
        End Select
    End Sub

    Private Function PrintSpecInstr(ByRef Line As Integer, ByRef OnOff As Boolean, Optional ByRef Alt As String = "") As Integer
        Dim Y As Integer
        Dim tStr As String, TStrs As Object, TI As Integer

        Y = OutputObject.CurrentY

        OutputObject.CurrentX = 7000
        OutputObject.FontUnderline = True
        If OnOff Then OutputObject.Print(" X ") Else OutputObject.Print("   ")
        OutputObject.FontUnderline = False
        OutputObject.CurrentY = Y

        Select Case Line
            Case 1
                tStr = StoreSettings.PoSpecInstr1
                If IsParkPlace Then tStr = "If order is less than 90 lbs., HOLD and SHIP with other goods."
            Case 2 : tStr = StoreSettings.PoSpecInstr2
            Case 3 : tStr = StoreSettings.PoSpecInstr3
            Case 4
                tStr = StoreSettings.PoSpecInstr4
                If Len(Trim(Alt)) > 0 Then tStr = Alt
                If Len(Trim(tStr)) = 0 Then tStr = New String("_", 30)
        End Select
        TStrs = SplitLongText(tStr, 35)

        If Not IsNothing(TStrs) Then
            For TI = LBound(TStrs) To UBound(TStrs)
                If TI = 2 Then Exit For
                OutputObject.CurrentX = 7500
                OutputObject.Print(TStrs(TI))
            Next
        End If
        PrintSpecInstr = OutputObject.CurrentY
    End Function

    Public Sub OpenPOReport()
        Inven = "POR"
        'InvPoPrint.HelpContextID = 57900
        ' We want Print/Preview/Cancel buttons, and no date.
        Text = IIf(ReportsMode("Ashley"), "Ashley Open Po Report", "Open PO Report")
        lblLabel.Text = "Sort Order:"
        txtDate.Visible = False
        cboSortOrder.Visible = True
        cboSortOrder.SelectedIndex = 0
        cmdPrint.Visible = True
        cmdPrintPreview.Visible = True
        Show()
    End Sub

End Class

Module modGrossMargin
    Private Const SoldTagFlag As String = "tg"
    Private PrepareMLForPackages_SaleNo As String, PrepareMLForPackages_cGM As CGrossMargin
    Public Function DescHasSoldTagPrinted(ByVal Desc As String) As Boolean
        DescHasSoldTagPrinted = Left(Desc, Len(SoldTagFlag)) = SoldTagFlag
    End Function
    Public Function PrintSoldTags(ByVal Style As String, Optional ByVal LastName As String = "", Optional ByVal SaleNo As String = "", Optional ByRef Q as integer = 1) As Boolean
        Dim UnloadAfter As Boolean
        If Not IsFormLoaded("SelectPrinter") Then UnloadAfter = True

        PrintSoldTags = SelectPrinter.PrintSoldTags(Style, LastName, SaleNo, Q)

        If UnloadAfter Then
            'Unload SelectPrinter
            SelectPrinter.Close()
        End If
    End Function
    Public Function DescSetSoldTagPrinted(ByVal Desc As String, Optional ByVal SaleNo As String = "", Optional ByVal StyleNo As String = "", Optional ByVal StoreNo As Integer = 0) As String
        Dim S As String

        If DescHasSoldTagPrinted(Desc) Then
            DescSetSoldTagPrinted = Desc
            Exit Function
        End If

        DescSetSoldTagPrinted = SoldTagFlag & Desc
        If SaleNo = "" Or StyleNo = "" Then Exit Function

        S = ""
        S = S & "UPDATE [GrossMargin] "
        S = S & "SET [Desc] = """ & ProtectSQL(DescSetSoldTagPrinted) & """ "
        S = S & "WHERE 1=1 "
        ' We don't have margin line on BOS
        S = S & "AND [SaleNo]=""" & ProtectSQL(SaleNo) & """ "
        S = S & "AND [Style]=""" & ProtectSQL(StyleNo) & """ "
        ' This should make it safe... Since we're updating the Desc, duplicates are changing as we go along...
        S = S & "AND [Desc]=""" & ProtectSQL(Desc) & """ "

        ExecuteRecordsetBySQL(S, , GetDatabaseAtLocation(StoreNo))
    End Function

    Public Sub SalePackageUpdate(ByVal SaleNo As String, Optional ByVal StoreNo as integer = 0, Optional ByVal TempTable As Boolean = False, Optional ByVal AllowCache As Boolean = True)
        Dim RS As ADODB.Recordset, ML as integer, S As String
        Dim IsPackage As Boolean, GM As Double, SellPrice As Decimal
        Dim TotLanded As Decimal, TotSellPr As Decimal
        Dim sTable As String, Sty As String

        If Trim(SaleNo) = "" Then Exit Sub

        sTable = "GrossMargin" & IIf(TempTable, "Tmp", "")

        RS = GetRecordsetBySQL("SELECT Style,MarginLine,Cost,ItemFreight,SellPrice,GM FROM " & sTable & " WHERE SaleNo='" & ProtectSQL(SaleNo) & "' ORDER BY MarginLine", , GetDatabaseAtLocation(StoreNo))
        Do While Not RS.EOF
            ML = IfNullThenZero(RS("MarginLine").Value)
            'Debug.Print "SalePackageUpdate ML=" & ML
            'If ML = 29751 Then Stop
            If ML <> 0 Then
                Sty = IfNullThenNilString(RS("Style").Value)
                IsPackage = PrepareMLForPackages(StoreNo, ML, SaleNo, GM, SellPrice, AllowCache)
                If IsItem(Sty) Or IsIn(Sty, "NOTES", "STAIN", "DEL", "LAB") Then
                    If Not IsItem(Sty) Then
                        SellPrice = IfNullThenZeroCurrency(RS("SellPrice").Value)
                        GM = 100 'IfNullThenZeroDouble(RS("GM"))
                    End If
                    TotLanded = TotLanded + IfNullThenZeroCurrency(RS("Cost").Value) + IfNullThenZeroCurrency(RS("ItemFreight").Value)
                    TotSellPr = TotSellPr + IfNullThenZeroCurrency(RS("SellPrice").Value)

                    If IsPackage Then
                        S = "UPDATE " & sTable & " SET IsPackage=1, PackSellGM=" & FormatGM(GM) & ", PackSell=" & SQLCurrency(SellPrice) & " WHERE MarginLine=" & ML
                    Else
                        S = "UPDATE " & sTable & " SET IsPackage=0, PackSellGM=" & FormatGM(GM) & ", PackSell=" & SQLCurrency(SellPrice) & " WHERE MarginLine=" & ML
                    End If

                    'Debug.Print "SalePackageUpdate S=" & S
                    ExecuteRecordsetBySQL(S, , GetDatabaseAtLocation(StoreNo))
                End If
            End If

            RS.MoveNext()
        Loop
        ClearPrepareMLForPackagesCache()
        DisposeDA(RS)

        GM = CalculateGM(TotSellPr, TotLanded)
        ExecuteRecordsetBySQL("UPDATE " & sTable & " SET PackSaleGM=" & FormatGM(GM) & " WHERE SaleNo='" & ProtectSQL(SaleNo) & "'", , GetDatabaseAtLocation(StoreNo))
    End Sub

    Public Sub ClearPrepareMLForPackagesCache()
        DisposeDA(PrepareMLForPackages_cGM)
        PrepareMLForPackages_SaleNo = ""
    End Sub

    Public Function WantCheckDisposal() As Boolean
        If IsSleepCity Then WantCheckDisposal = True
        If IsBarrs Then WantCheckDisposal = True
    End Function

    Public Function DisposalDepartment() As Integer
        DisposalDepartment = -1
        If Not WantCheckDisposal() Then Exit Function
        If IsSleepCity Then DisposalDepartment = 10
        If IsBarrs Then DisposalDepartment = 10
    End Function

    Public Function PrepareMLForPackages(ByVal S as integer, ByVal ML as integer, ByVal SaleNo As String, Optional ByRef GM As Double = 0, Optional ByRef SellPrice As Decimal = 0, Optional ByVal Cache As Boolean = False) As Boolean
        Dim G As CGrossMargin, I as integer, Pkg As Boolean, Cnt as integer, TotLand As Decimal, TotCost As Decimal, TotSell As Decimal, PKGM As Double
        If Cache And SaleNo = PrepareMLForPackages_SaleNo And Not PrepareMLForPackages_cGM Is Nothing Then
            G = PrepareMLForPackages_cGM
            G.DataAccess.RS.MoveFirst()
            G.DataAccess.CurrentIndex = -1
        Else
            ClearPrepareMLForPackagesCache()
            G = New CGrossMargin
            G.DataAccess.DataBase = GetDatabaseAtLocation(S)
            G.DataAccess.Records_OpenSQL("SELECT * FROM GrossMargin WHERE SaleNo='" & SaleNo & "' ORDER BY MarginLine")
        End If

        Pkg = False
        Cnt = 0

        'If isdevelopment and Trim(SaleNo) = "11344" Then Stop
        'If IsDevelopment And ML = 40171 Then Stop

        GM = 0
        SellPrice = 0
        PrepareMLForPackages = False

        Do While G.DataAccess.Records_Available
            G.cDataAccess_GetRecordSet(G.DataAccess.RS)
            If IsItem(G.Style) Then
                If G.MarginLine = ML Then
                    GM = G.GM
                    If Val(GM) = 0 And G.SellPrice <> 0 And G.Cost <> 0 Then GM = CalculateGM(GetPrice(G.SellPrice), GetPrice(G.Cost) + GetPrice(G.ItemFreight), , 0)
                    SellPrice = G.SellPrice
                    If G.SellPrice <> 0 And Not Pkg Then GoTo FoundNotPackage
                End If

                If G.SellPrice = 0 Then
                    Pkg = True
                    Cnt = Cnt + 1
                    TotLand = TotLand + G.Cost + G.ItemFreight
                    TotCost = TotCost + G.Cost
                ElseIf Pkg Then ' end of package
                    Cnt = Cnt + 1
                    TotCost = TotCost + G.Cost
                    TotLand = TotLand + G.Cost + G.ItemFreight

                    Pkg = False
                    TotSell = G.SellPrice
                    GM = CalculateGM(TotSell, TotLand, , 2)
                    For I = 1 To Cnt - 1
                        G.DataAccess.Records_MovePrevious()
                    Next
                    For I = 1 To Cnt
                        If G.MarginLine = ML Then
                            If TotCost <> 0 Then SellPrice = Math.Round(G.Cost / TotCost * TotSell, 2)
                            PrepareMLForPackages = True
                            DisposeDA(PrepareMLForPackages_cGM) : PrepareMLForPackages_SaleNo = ""
                            DisposeDA(G)
                            Exit Function
                        End If
                        If I <> Cnt Then G.DataAccess.Records_MoveNext()
                    Next
                    If G.MarginLine = ML Then
                        SellPrice = Math.Round(G.Cost / TotCost * TotSell, 2)
                        PrepareMLForPackages = True
                        DisposeDA(PrepareMLForPackages_cGM) : PrepareMLForPackages_SaleNo = ""
                        DisposeDA(G)
                        Exit Function
                    End If

                    Cnt = 0
                    TotLand = 0
                    TotCost = 0
                End If
            Else
                Pkg = False    ' SUB and other stuff break packages... dangling items do not get updated.
            End If
        Loop

FoundNotPackage:

        If Cache Then
            PrepareMLForPackages_SaleNo = SaleNo
            PrepareMLForPackages_cGM = G
            G = Nothing ' don't dispose, that is done globally.  Simply remove this pointer.
        Else
            ClearPrepareMLForPackagesCache()
            DisposeDA(G)
        End If
    End Function

    Public Function SaleToHTML(ByVal SaleNo As String, Optional ByVal StoreNo as integer = 0, Optional ByRef GetEmailAddr As String = "", Optional ByRef GetEmailName As String = "", Optional ByRef CustomerCopy As Boolean = True) As String
        Dim S As String, G As CGrossMargin, H As cHolding, C As clsMailRec, C2 As MailNew2
        Dim P As Decimal, N as integer, Alt As Boolean
        Dim Terms As String
        On Error Resume Next

        If StoreNo = 0 Then StoreNo = StoresSld


        H = New cHolding
        If Not H.Load(SaleNo, "LeaseNo") Then
            H = Nothing
            Exit Function
        End If

        G = New CGrossMargin
        G.DataAccess.DataBase = GetDatabaseAtLocation(StoreNo)
        G.DataAccess.Records_OpenSQL("SELECT * FROM [GrossMargin] WHERE SaleNo='" & SaleNo & "' ORDER BY [MarginLine]")
        G.DataAccess.Records_Available()

        C = New clsMailRec
        If Val(G.Index) <> 0 Then
            C.Load(G.Index, "#Index")
            GetEmailAddr = C.Email
            GetEmailName = C.First & " " & C.Last
            '    DisposeDA C

            Mail2_GetAtIndex(G.Index, C2, StoreNo)
        End If

        S = ""

        ' templates
        S = S & "" & vbCrLf
        '  S = S & "  <></>" & vbCrLf

        '------- Page Setup
        S = S & "<html>" & vbCrLf
        S = S & " <head>" & vbCrLf
        'S = S & "  <title>Order #" & SaleNo & " - " & StoreSettings(StoreNum).Name & "</title>" & vbCrLf
        S = S & "" & vbCrLf
        S = S & " </head>" & vbCrLf
        S = S & " <body bgcolor=#FFFFFF text=#000000 vlink=#FF0000 link=#FFFF00 hspace=0 vspace=0>" & vbCrLf

        S = S & " <table border=0 bgcolor=#FFFF99 width='700'><tr><td>" & vbCrLf ' background table

        S = S & "  <table border=0 cellspacing=0 cellpadding=0 width='100%'>" & vbCrLf ' alignment table
        S = S & "   <tr><td colspan=3 align=center>" & vbCrLf

        '------- Store Information
        S = S & "      <table width='100%'><tr>" & vbCrLf
        S = S & "      <td width='50%'>" & vbCrLf
        S = S & "        <table border=0 width=100% height='100%'>" & vbCrLf
        'S = S & "          <tr><td align=center><font size=+3><b>" & StoreSettings(StoreNum).Name & "</b></font></td></tr>" & vbCrLf
        'S = S & "          <tr><td align=center>" & StoreSettings(StoreNum).Address & "</td></tr>" & vbCrLf
        'S = S & "          <tr><td align=center>" & StoreSettings(StoreNum).City & "</td></tr>" & vbCrLf
        'S = S & "          <tr><td align=center>" & DressAni(StoreSettings(StoreNum).Phone) & "</td></tr>" & vbCrLf
        'S = S & "        </table>" & vbCrLf
        S = S & "      </td>" & vbCrLf
        S = S & "      </tr></table><br>" & vbCrLf

        S = S & "   </tr></td>" & vbCrLf
        S = S & "   <tr><td colspan=3 align=center>" & vbCrLf
        '------- Page Header
        S = S & "    <table border=0 cellspacing=0 cellpadding=0 width='100%'><tr height='100%'>" & vbCrLf
        S = S & "      <td>" & vbCrLf
        S = S & "        <table height='100%'>" & vbCrLf
        S = S & "          <tr>" & vbCrLf
        S = S & "            <td><b>Sale Date</b></td>" & vbCrLf
        S = S & "            <td>" & Format(G.SellDte, "mm/dd/yyyy") & "</td>" & vbCrLf
        S = S & "          </tr>" & vbCrLf
        S = S & "          <tr>" & vbCrLf
        S = S & "            <td><b>Delivery / Pickup?</b></td>" & vbCrLf
        S = S & "            <td>" & G.PorD & "</td>" & vbCrLf
        S = S & "          </tr>" & vbCrLf
        S = S & "          <tr>" & vbCrLf
        S = S & "            <td><b>Delivery Date</b></td>" & vbCrLf
        S = S & "            <td>" & Format(G.DDelDat, "mm/dd/yyyy") & "</td>" & vbCrLf
        S = S & "          </tr>" & vbCrLf
        S = S & "        </table>" & vbCrLf
        S = S & "      </td>" & vbCrLf

        S = S & "      <td>" & vbCrLf
        S = S & "        <table height='100%'>" & vbCrLf
        S = S & "          <tr>" & vbCrLf
        S = S & "            <td>Sale No:</td>" & vbCrLf
        S = S & "            <td><font size=+2>" & G.SaleNo & "</td>" & vbCrLf
        S = S & "          </tr>" & vbCrLf
        '  S = S & "          <tr>" & vbCrLf
        '  S = S & "            <td>Status:</td>" & vbCrLf
        '  S = S & "            <td>" & DescribeHoldingStatus(H.Status) & "</td>" & vbCrLf
        '  S = S & "          </tr>" & vbCrLf
        S = S & "        </table>" & vbCrLf
        S = S & "      </td>" & vbCrLf

        S = S & "    </tr></table>" & vbCrLf

        S = S & "   </td></tr>" & vbCrLf ' alignment table
        S = S & "   <tr height=100%><td width='49%'>" & vbCrLf
        '------- Customer Info
        S = S & "      <table width='100%' height='100%' valign=top bgcolor=#AA0000 border=0 cellpadding=0 cellspacing=1><tr><td>" & vbCrLf
        S = S & "       <table width='100%' height='100%' bgcolor=#FFFFFF border=0 cellpadding=0 cellspacing=0>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td><b><font size=-2>First Name</font></b></td>" & vbCrLf
        S = S & "           <td><b><font size=-2>Last Name</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(C.First) & "</td>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(C.Last) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td colspan=2><b><font size=-2>Address</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td colspan=2>" & HTMLBS(C.Address) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td colspan=2><b><font size=-2>Additional Address</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td colspan=2>" & HTMLBS(C.AddAddress) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td><b><font size=-2>City / State</font></b></td>" & vbCrLf
        S = S & "           <td><b><font size=-2>Zip</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(C.City) & "</td>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(C.Zip) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td><b><font size=-2>Telephone 1</font></b></td>" & vbCrLf
        S = S & "           <td><b><font size=-2>Telephone 2</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(DressAni(C.Tele)) & "</td>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(DressAni(C.Tele2)) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "       </table>" & vbCrLf
        S = S & "      </td></tr></table>" & vbCrLf

        S = S & "   </td><td width='2%'>&nbsp;" & vbCrLf
        S = S & "   </td><td width='49%'>" & vbCrLf ' alignment table
        '------- Ship To
        S = S & "      <table width='100%' height='100%' align=top bgcolor=#AA0000 border=0 cellpadding=0 cellspacing=1><tr><td>" & vbCrLf
        S = S & "       <table width='100%' height='100%' bgcolor=#FFFFFF border=0 cellpadding=0 cellspacing=0>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td colspan=2 align=center><b>Ship To Address:</b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td><b><font size=-2>First</font></b></td>" & vbCrLf
        S = S & "           <td><b><font size=-2>Last/Company</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(C2.ShipToFirst) & "</td>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(C2.ShipToLast) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td colspan=2><b><font size=-2>Address</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td colspan=2>" & HTMLBS(C2.Address2) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td><b><font size=-2>City / State</font></b></td>" & vbCrLf
        S = S & "           <td><b><font size=-2>Zip</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(C2.City2) & "</td>" & vbCrLf
        'S = S & "           <td>" & HTMLBS(C2.Zip2) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td colspan=2><b><font size=-2>Telephone 3</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td colspan=2>" & HTMLBS(DressAni(C.Tele)) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "         <tr><td colspan=2></td></tr>" & vbCrLf

        S = S & "         <tr>" & vbCrLf
        S = S & "           <td colspan=2><b><font size=-2>Sales Staff</font></b></td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf
        S = S & "         <tr>" & vbCrLf
        'S = S & "           <td colspan=2>" & HTMLBS(TranslateSalesmen(G.Salesman)) & "</td>" & vbCrLf
        S = S & "         </tr>" & vbCrLf

        S = S & "       </table>" & vbCrLf
        S = S & "      </td></tr></table>" & vbCrLf

        S = S & "   </td></tr>" & vbCrLf ' alignment table
        S = S & "   <tr><td colspan=3>" & vbCrLf
        '------- Special Instr

        S = S & "<b>Special Instructions:</b><br>"
        S = S & C.Special

        S = S & "   </td></tr>" & vbCrLf ' alignment table
        S = S & "  </table>" & vbCrLf
        S = S & "<br>" & vbCrLf
        '------- The actual sale w/ all the items...

        S = S & "  <table width='100%' cellspacing=0 cellpadding=2 border=1 bgcolor=#FFFFF>" & vbCrLf
        S = S & "   <thead bgcolor=#AAAAFF>"
        S = S & "    <tr>" & vbCrLf
        If CustomerCopy Then
            S = S & "      <td width='15%'>&nbsp;</td>" & vbCrLf
            S = S & "      <td width='15%'>&nbsp;</td>" & vbCrLf
        Else
            S = S & "      <td width='15%'>Style</td>" & vbCrLf
            S = S & "      <td width='15%'>Vendor</td>" & vbCrLf
        End If
        S = S & "      <td width='4%' >Loc</td>" & vbCrLf
        S = S & "      <td width='11%'>Status</td>" & vbCrLf
        S = S & "      <td width='7%' >Quantity </td>" & vbCrLf
        S = S & "      <td width='33%'>Description</td>" & vbCrLf
        S = S & "      <td width='15%' align=right>Price</td>" & vbCrLf
        S = S & "    </tr>" & vbCrLf
        S = S & "   </thead>" & vbCrLf
        S = S & "   <tbody>" & vbCrLf
        Do ' we already did the first records_available above to get mailindex
            S = S & "    <tr bgcolor=" & IIf(Alt, "#FFDDAA", "#FFFFFF") & ">" & vbCrLf
            Alt = Not Alt
            If CustomerCopy Then
                S = S & "      <td>&nbsp;</td>" & vbCrLf
                S = S & "      <td>&nbsp;</td>" & vbCrLf
            Else
                'S = S & "      <td>" & HTMLBS(G.Style) & "</td>" & vbCrLf
                'S = S & "      <td>" & HTMLBS(G.Vendor) & "</td>" & vbCrLf
            End If
            N = Val(G.Location)
            S = S & "      <td>" & IIf(N = 0, "&nbsp;", N) & "</td>" & vbCrLf
            'S = S & "      <td>" & HTMLBS(G.Status) & "</td>" & vbCrLf
            N = Val(G.Quantity)
            S = S & "      <td>" & IIf(N = 0, "&nbsp;", N) & "</td>" & vbCrLf
            'S = S & "      <td>" & HTMLBS(G.Desc) & "</td>" & vbCrLf
            P = GetPrice(G.SellPrice)
            S = S & "      <td align=right>" & IIf(P = 0, "&nbsp;", CurrencyFormat(P)) & "</td>" & vbCrLf
            S = S & "    </tr>" & vbCrLf
        Loop While G.DataAccess.Records_Available
        S = S & "    <tr bgcolor=#DDDDDD>" & vbCrLf
        S = S & "      <td colspan=6 align=right><b>Balance Due:</b></nbsp>" & vbCrLf
        S = S & "      <td align=right>" & CurrencyFormat(H.Sale - H.Deposit) & "</td>" & vbCrLf
        S = S & "    </tr>" & vbCrLf
        S = S & "   </tbody>" & vbCrLf
        S = S & "  </table>" & vbCrLf

        S = S & "<br>" & vbCrLf
        S = S & "  <table width='100%' align=top bgcolor=#AA0000 border=0 cellpadding=0 cellspacing=1><tr><td>" & vbCrLf
        S = S & "   <table width='100%' height='100%' bgcolor=#FFFFFF border=0 cellpadding=0 cellspacing=0>" & vbCrLf
        S = S & "     <tr><td>" & vbCrLf
        '      MainMenu.rtb.LoadFile CustomerTermsMessageFile
        'Terms = MainMenu.rtb.Text
        '      MainMenu.rtb.Text = ""
        Terms = Replace(Terms, vbCr, "<br>")
        Terms = Replace(Terms, vbLf, "")

        'S = S & HTMLBS(Terms) & vbCrLf
        S = S & "     </td></tr>" & vbCrLf
        S = S & "   </table>" & vbCrLf
        S = S & "  </td></tr></table>" & vbCrLf


        '------- Page CleanUp
        S = S & " </td></tr></table>" & vbCrLf ' end background table

        S = S & " </body>" & vbCrLf
        S = S & "</html>" & vbCrLf

        SaleToHTML = S

        DisposeDA(H, G, C)
    End Function
End Module

Imports VBRUN
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class Calendar
    Private Const MAX_DAY_COLUMNS as integer = 31
    ' This form is loaded modal by InvOrdStatus and BillOSale.
    Private mLoadedByForm As Boolean
    Private mGridMoving As Boolean
    Private AllowMap As Boolean
    Private AllowInstr As Boolean
    Private DDTRS() As ADODB.Recordset

    Public Property LoadedByForm() As Boolean
        Get
            LoadedByForm = mLoadedByForm
        End Get
        Set(value As Boolean)
            mLoadedByForm = value
            If value Then
                cmdMenu.Text = "Back"
            Else
                cmdMenu.Text = "Menu"
            End If
        End Set
    End Property

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        ' Update the column label..
        Dim OldRow as integer

        If mGridMoving Then Exit Sub
        If grid.Rows < 2 Then Exit Sub
        mGridMoving = True
        OldRow = grid.Row
        grid.Row = 1
        grid.Text = txtDayLabel.Text
        UpdateGridHeight()
        grid.Row = OldRow
        mGridMoving = False

        Dim cDelDesc As clsDeliveryDateDesc
        cDelDesc = New clsDeliveryDateDesc

        If Not cDelDesc.Load(DateAdd("d", grid.Col, Today), "@DeliveryDate") Then
            cDelDesc.DeliveryDate = DateAdd("d", grid.Col, Today)
        End If
        cDelDesc.DeliveryDesc = txtDayLabel.Text

        cDelDesc.Save()

        DisposeDA(cDelDesc)
    End Sub

    Private Sub UpdateGridHeight()
        ' I wish there were a more automatic way to do this..

        If Not mGridMoving Then Exit Sub
        Dim TR as integer, TW as integer, OldCol as integer, mGridDescHeight as integer
        OldCol = grid.Col
        grid.Row = 1
        For TR = 0 To grid.Cols - 1
            grid.Col = TR
            TW = Math.Round(Printer.TextWidth(grid.Text) / 1500 + 0.4999) ' # of cell width units this text requires, fudged for word wrap.
            If TW > mGridDescHeight Then
                mGridDescHeight = TW
            End If
        Next
        'grid.RowHeight(1) = 240 * mGridDescHeight
        grid.set_RowHeight(1, 240 * mGridDescHeight)
        grid.Col = OldCol
    End Sub

    Private Sub cmdInstr_Click(sender As Object, e As EventArgs) Handles cmdInstr.Click
        'Load CalendarInstr
        CalendarInstr.DeliveryDay = DateAdd("d", grid.Col, Today)
        'CalendarInstr.Show vbModal
        CalendarInstr.ShowDialog()
    End Sub

    Public Function GetDeliveryCalendarData(ByVal StartDate As Date, ByVal DayCount as integer, Optional ByVal Store as integer = 0, Optional ByVal WithRecord As Boolean = False, Optional ByVal AllowSLDN As Boolean = False) As ADODB.Recordset
        Dim SQL As String, EndDate As Date

        If Store = 0 Then Store = StoresSld
        StartDate = DateFormat(StartDate)
        EndDate = DateFormat(DateAdd("d", DayCount - 1, StartDate))
        '  sql = " SELECT DISTINCT (DateDiff('d',#" & StartDate & "#,[DelDate])) AS [Index], DelDate, SaleNo, Name, iif(PorD='P', 'P', '') as PD" _
        '      & " FROM GrossMargin" _
        '      & " WHERE DelDate BETWEEN #" & StartDate & "# AND #" & EndDate & "#" _
        '      & "     And Trim(Status) Not in ('xLAW', 'VDLAW', 'xST')" _
        '      & "     And Left(Trim(Status),3)<>'DEL'" _
        '      & "     And Trim(Style) Not In (" & NonItemStyleString & ")" _
        '      & "     And Trim(Name) Not In ('', 'CASH & CARRY')" _
        '      & " ORDER BY DelDate, SaleNo, Name"

        ' Also get Service Calls.
        SQL = ""
        SQL = SQL & " SELECT DISTINCT (DateDiff('d',#" & StartDate & "#,[DelDate])) AS [Index], DelDate, SaleNo, Name, MailIndex, iif(PorD='P', 'P', '') as PD, Tele, StopStart, StopEnd" & IIf(WithRecord, ", MarginLine as [Record]", "")
        SQL = SQL & " FROM GrossMargin"
        SQL = SQL & " WHERE DelDate BETWEEN #" & StartDate & "# AND #" & EndDate & "#"
        SQL = SQL & "     And Trim(Status) Not in ('xLAW', 'VDLAW', 'xST')"
        SQL = SQL & "     And Left(Trim(Status),3)<>'DEL'"
        SQL = SQL & "     And Trim(Style) Not In (" & NonItemStyleString(AllowSLDN, AllowSLDN) & ")"
        'BFH20150225 - BFMyer was getting empty names for some sales, and so they weren't showing up in the Delivery Calendar
        '  SQL = SQL & "     And Trim(Name) Not In ('', 'CASH & CARRY')"
        SQL = SQL & "     And Trim(Name) Not In ('CASH & CARRY')"
        SQL = SQL & " UNION "
        SQL = SQL & " SELECT (DateDiff('d',#" & StartDate & "#,[ServiceOnDate])) AS [Index], ServiceOnDate, ""SO"" & ServiceOrderNo, LastName, MailIndex, 'S' as PD, Telephone, StopStart, StopEnd" & IIf(WithRecord, ", ServiceOrderNo as [Record]", "")
        SQL = SQL & " FROM Service"
        SQL = SQL & " WHERE [Status]='Open' "
        SQL = SQL & " AND IsDate(ServiceOnDate) and  DateDiff('d',#" & StartDate & "#,[ServiceOnDate]) between 0 and " & (DayCount - 1)
        SQL = SQL & " UNION "
        SQL = SQL & " SELECT (DateDiff('d',#" & StartDate & "#,[Ddate1])) AS [Index], [Ddate1], 'TR' & Misc, Name, 0, 'T' as PD, '', '', ''" & IIf(WithRecord, ", DetailID as [Record]", "")
        SQL = SQL & " FROM [Detail]"
        SQL = SQL & " IN '" & GetDatabaseInventory() & "'" ' pulled from a different database using the IN <database-path> clause
        SQL = SQL & " WHERE [Trans]='TP'"
        SQL = SQL & " AND IsDate(Ddate1) and  DateDiff('d',#" & StartDate & "#,[Ddate1]) between 0 and " & (DayCount - 1)
        SQL = SQL & " ORDER BY 1, 2, 5, 3, 4"   '"DelDate, PorD, SaleNo, Name"  ' Union requires sorting by position.

        GetDeliveryCalendarData = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(Store))
    End Function

    Private Sub cmdManifest_Click(sender As Object, e As EventArgs) Handles cmdManifest.Click
        PrintDayDeliveryManifest
    End Sub

    Private Sub PrintDayDeliveryManifest()
        Dim Whenn As Date, CD As ADODB.Recordset, RD As ADODB.Recordset, X As String, Y as integer
        Whenn = DateAdd("d", grid.Col, Today)

        OutputToPrinter = True
        OutputObject = Printer

        PrintManifestHeader(Whenn)
        CD = GetDeliveryCalendarData(Whenn, 1)
        Do While Not CD.EOF
            Y = Printer.CurrentY

            X = IfNullThenNilString(CD("PD").Value)
            If X = "" Then X = "D"
            '    If X = "D" Then
            PrintAligned(X, , 100, Y)
            PrintAligned(CD("SaleNo").Value, , 500, Y)
            PrintAligned(CD("Name").Value, , 1600, Y)

            Dim R As clsMailRec
            R = GetMailByLeaseNo(CD("SaleNo").Value)
            If Not (R Is Nothing) Then
                PrintAligned(R.City, , 3600, Y)
            End If
            DisposeDA(R)

            PrintAligned(DressAni(CleanAni(IfNullThenNilString(CD("Tele").Value))), , 6000, Y)
            RD = GetRecordsetBySQL("SELECT Sale - Deposit AS BalDue FROM Holding WHERE LeaseNo='" & CD("SaleNo").Value & "'", , GetDatabaseAtLocation())
            If Not RD.EOF Then
                PrintAligned(FormatCurrency(RD("BalDue").Value), , 7800, Y)
            End If
            RD = Nothing
            PrintAligned("_______", , 9100, Y)
            PrintAligned("_______", , 10100, Y)
            PrintAligned("")
            '    End If

            If Printer.CurrentY > Printer.ScaleHeight - 500 Then
                Printer.NewPage()
                PrintManifestHeader(Whenn)
            End If

            CD.MoveNext
        Loop
        CD = Nothing
        Printer.EndDoc()
    End Sub

    Private Sub PrintManifestHeader(ByVal Whenn As Date)
        Dim Y as integer
        Printer.FontSize = 20
        Printer.FontBold = True
        PrintAligned(StoreSettings.Name, AlignmentConstants.vbCenter)
        PrintAligned(StoreSettings.Address, AlignmentConstants.vbCenter)
        Printer.FontSize = 14
        PrintAligned("")
        PrintAligned("Delivery Manifest for " & Whenn, AlignmentConstants.vbCenter)
        PrintAligned("")
        Y = Printer.CurrentY
        Printer.FontSize = 10
        Printer.FontBold = False
        PrintAligned("P/D", , 100, Y, True)
        PrintAligned("SaleNo", , 500, Y, True)
        PrintAligned("Name", , 1600, Y, True)
        PrintAligned("City", , 3600, Y, True)
        PrintAligned("Tele", , 6000, Y, True)
        PrintAligned("BalDue", , 7800, Y, True)
        PrintAligned("Complete", , 9100, Y, True)
        PrintAligned("Partial", , 10100, Y, True)
        'Printer.Line(100, Printer.CurrentY) - (Printer.ScaleWidth - 100, Printer.CurrentY)
        Printer.Line(100, Printer.CurrentY, Printer.ScaleWidth - 100, Printer.CurrentY)
    End Sub

    Private Sub cmdMap_Click(sender As Object, e As EventArgs) Handles cmdMap.Click
        Dim I as integer, DCount as integer, Ty As String

        ActiveLog("Calendar::OpenMap", 5)
        If Not AllowMap Then Exit Sub ' They have to have mapping components
        ActiveLog("Calendar::OpenMap - Allowed", 6)
        If grid.Col < 0 Then          ' They have to select a day first.
            MessageBox.Show("Please select a date first.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        For I = 2 To grid.Rows - 1
            'Ty = Left(grid.TextMatrix(I, grid.Col), 1)
            Ty = Microsoft.VisualBasic.Left(grid.get_TextMatrix(I, grid.Col), 1)
            'If Ty = "L" Then Ty = Mid(grid.TextMatrix(I, grid.Col), 5, 1)
            If Ty = "L" Then Ty = Mid(grid.get_TextMatrix(I, grid.Col), 5, 1)
            If IsIn(Ty, "D", "S") Then DCount = DCount + 1
        Next
        If DCount = 0 Then
            MessageBox.Show("There are no deliveries scheduled for this day.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Hide()
        ProgressForm(0, 1, "Loading MapPoint...")
        ActiveLog("Calendar::OpenMap - Loading Map Form...", 4)
        'Load frmDeliveryMap
        ActiveLog("Calendar::OpenMap - Loaded Map Form...", 3)
        ProgressForm()

        ActiveLog("Calendar::OpenMap - Showing Map Form...", 4)
        frmDeliveryMap.Show
        ActiveLog("Calendar::OpenMap - Showed Map Form...", 3)
        frmDeliveryMap.CreateRoute(IIf(chkMultiple.Checked = True, 0, StoresSld), DateValue(Mid(grid.get_TextMatrix(0, grid.Col), 6))) '  DateAdd("d", grid.Col, Date)
    End Sub

    Private Sub cmdMenu_Click(sender As Object, e As EventArgs) Handles cmdMenu.Click
        Me.Close()
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim T As String, D As Date

        D = Now
        If MessageBox.Show("Print calendar for current range?", "Print Calendar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            T = InputBox("Enter Start of Period:", "Print Calendar", Today)
            If Not IsDate(T) Then
                MessageBox.Show("Invalid date.", "Print Calendar", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            D = DateValue(T)
        End If
        PrintCalendar(D, MAX_DAY_COLUMNS)
    End Sub

    Private Sub PrintCalendar(ByVal StartDate As Date, ByVal DayCount as integer)
        Dim Y as integer, LastY as integer
        Dim NumPages as integer, PageColCount as integer
        Dim ColCount as integer, ColRepeats as integer, ColRows as integer, newCol As Boolean
        Dim MaxItemsPerColumn as integer, MaxPageColCount as integer

        MaxItemsPerColumn = 50
        MaxPageColCount = 7
        NumPages = 1
        LastY = -1
        ' Data initialization and validation
        StartDate = DateFormat(StartDate)

        'Printer initialization
        'Printer.Font = "Arial"
        Printer.Font = New Font("Arial", FontStyle.Regular)
        Printer.Orientation = 2

        ' Print deliveries.  Automatically add date headers for each new column, and pages as necessary.
        Dim CD As ADODB.Recordset
        CD = GetDeliveryCalendarData(StartDate, DayCount)

        Do Until CD.EOF
            ' Find out if we need to write a header.
            newCol = False
            Y = CD("Index").Value
            If LastY <> Y Then
                ' Date has changed, so at least one new column is needed.
                newCol = True
                ColRepeats = 1
            Else
                ColRows = ColRows + 1
                If ColRows >= MaxItemsPerColumn Then
                    newCol = True
                    ColRepeats = ColRepeats + 1
                End If
            End If

            If newCol Then
                ' Print date headers from the last date to the current.
                ' This prints empty columns for days with no deliveries.
                Dim I as integer
                For I = LastY + 1 To Y
                    ' Do we need a new page?
                    ColCount = ColCount + 1
                    PageColCount = PageColCount + 1
                    If PageColCount > MaxPageColCount Then
                        NumPages = NumPages + 1
                        PageColCount = 1
                    End If
                    If PageColCount = 1 Then
                        If NumPages > 1 Then Printer.NewPage()
                        PrintDeliveryCalendarHeader(DayCount, NumPages)
                    End If
                    PrintDeliveryCalendarDateHeader(StartDate, I, PageColCount, ColRepeats)
                Next
                ColRows = 1
                LastY = Y
            End If

            ' Print delivery information.  Each line automatically adjusts Y position.
            '      If PageColCount = 1 Then Printer.CurrentX = 50
            '      Printer.Print Tab((PageColCount - 1) * 30 + 5); Left(CD("Name"), 18); Tab((PageColCount - 1) * 30 + 24); CD("SaleNo")
            Printer.Print(TAB(PrinterPosition(PageColCount, 5)), Microsoft.VisualBasic.Left(CD("Name").Value, 18), TAB(PrinterPosition(PageColCount, 24)), CD("SaleNo"))
            CD.MoveNext()
        Loop


        'Y*30+5, Y*30+24

        '      If Y = 1 Then
        '          Printer.CurrentX = 50
        '          Printer.Print Left(.rs("Last"), 18); Tab(24); SaleNo(Ix, Y)
        '      ElseIf Y = 2 Then
        '          Printer.Print Tab(35); Left(.rs("Last"), 18); Tab(54); SaleNo(Ix, Y)
        '      ElseIf Y = 3 Then
        '          Printer.Print Tab(65); Left(.rs("Last"), 18); Tab(84); SaleNo(Ix, Y)
        '      ElseIf Y = 4 Then
        '          Printer.Print Tab(95); Left(.rs("Last"), 18); Tab(114); SaleNo(Ix, Y)
        '      ElseIf Y = 5 Then
        '          Printer.Print Tab(125); Left(.rs("Last"), 18); Tab(144); SaleNo(Ix, Y)
        '      ElseIf Y = 6 Then
        '          Printer.Print Tab(154); Left(.rs("Last"), 18); Tab(173); SaleNo(Ix, Y)
        '      ElseIf Y = 7 Then
        '          Printer.Print Tab(183); Left(.rs("Last"), 18); Tab(202); SaleNo(Ix, Y)
        '      End If
        '    Next
        '  Next

        Printer.EndDoc()
        'MousePointer = 0
        Me.Cursor = Cursors.Default
        Printer.Orientation = 1
        Exit Sub

HandleErr:
        Resume Next
    End Sub

    Private Sub PrintDeliveryCalendarHeader(ByVal DayCount as integer, ByVal NumPages as integer)
        ' Set the printer font and position for document title
        Printer.FontSize = 16
        Printer.FontBold = True
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        PrintCentered("Delivery Calendar - Next " & DayCount & " Days" & IIf(NumPages > 1, " - Page " & NumPages, ""))

        ' Set the printer font and position for store information
        Printer.CurrentY = 350
        Printer.FontSize = 8
        PrintCentered(StoreSettings.Name & "  " & "  " & StoreSettings.Address & "  " & "  " & StoreSettings.City)
    End Sub

    Private Sub PrintDeliveryCalendarDateHeader(ByVal StartDate As Date, ByVal Offset as integer, ByVal ColNum as integer, ByVal ColRepeats as integer)
        Dim Rm as integer
        ' Set the printer font and position for day and date labels
        Printer.FontBold = True
        Printer.FontSize = 12
        Printer.CurrentX = 0
        Printer.CurrentY = 700

        Rm = 800 + (ColNum - 1) * 2150
        ' Set CurrentX instead of using Tabs!  This will allow large, small, bold, and normal text to line up.
        Printer.CurrentX = Rm + 220
        Printer.Print(UCase(Format(DateAdd("d", Offset, StartDate), "DDD." & IIf(ColRepeats > 1, " (" & ColRepeats & ")", ""))))
        Printer.CurrentX = Rm
        Printer.Print(Format(DateAdd("d", Offset, StartDate), "m/d/yyyy"))

        ' Old code, kept for spacing information.  Prints 7 day and date labels.
        'Printer.Print "   "; DayLabel(0).Caption; Tab(26); DayLabel(1).Caption; Tab(44); DayLabel(2).Caption; Tab(63); DayLabel(3).Caption; Tab(82); DayLabel(4).Caption; Tab(100); DayLabel(5).Caption; Tab(119); DayLabel(6).Caption
        '  Printer.Print " "; dateformat(Now); Tab(24); DateAdd("D", 1, dateformat(Now)); Tab(42); DateAdd("D", 2, dateformat(Now)); Tab(61); DateAdd("D", 3, dateformat(Now)); Tab(80); DateAdd("D", 4, dateformat(Now)); _
        '  Tab(98); DateAdd("D", 5, dateformat(Now)); Tab(117); DateAdd("D", 6, dateformat(Now))

        ' And set up the printer object for the following delivery data.
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = 20
        Printer.CurrentY = 1300
    End Sub

    Private Function PrinterPosition(ByVal ColumnNumber As Object, ByVal AdditionalIndent As Object) as integer
        PrinterPosition = (ColumnNumber - 1) * 30 + AdditionalIndent
    End Function

    Private Sub Calendar_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Unload Me  ' This shouldn't be necessary, the form's already being unloaded when it gets here.
        If Not LoadedByForm Then
            MainMenu.Show()
            'Unload BillOSale
            BillOSale.Close()
        End If
        LoadedByForm = False
    End Sub

    Private Function CheckAllowMap() As Boolean
        Dim C As String, cX As String, L As String, Lx As String
        Dim MI(0 To 45) As String, I as integer
        Dim N As Object, M As Object

        C = LocalRoot & "Program Files\"
        cX = LocalRoot & "Program Files (x86)\"
        L = LocalProgramFilesFolder()
        Lx = Microsoft.VisualBasic.Left(L, Len(L) - 1) & " (x86)\"

        MI(0) = "Microsoft MapPoint\MappointControl.ocx"
        MI(1) = "Microsoft MapPoint 2002\MappointControl.ocx"
        For I = 2 To UBound(MI)
            MI(I) = Replace(MI(1), "2002", "" & (2001 + I))
        Next

        Dim A() As Object = {C, cX, L, Lx}
        For Each M In MI
            'For Each N In Array(C, cX, L, Lx)
            For Each N In A
                    If FileExists(N & M) Then CheckAllowMap = True : Exit Function
                    Next
                Next
    End Function

    Private Sub Calendar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim AllStores As Boolean

        SetButtonImage(cmdMenu, "menu")
        SetButtonImage(cmdManifest, "calendar")
        SetButtonImage(cmdPrint, "print")
        SetButtonImage(cmdInstr, "zoom")
        SetButtonImage(cmdMap, "map")
        SetButtonImage(cmdDDT, "south")

        AllowInstr = True
        AllowMap = CheckAllowMap()  ' Poor man's component-installation check.
        cmdMap.Enabled = AllowMap
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor

        AllStores = (StoreSettings.bOneCalendar)
        LoadGrid(Now(), MAX_DAY_COLUMNS, AllStores)
        'MousePointer = vbNormal
        Me.Cursor = Cursors.Default

        cmdDDT.Enabled = DDTLicensed

        If Width > Screen.PrimaryScreen.Bounds.Width Then
            ' Form is too wide for some customers!
            Dim GridBorder as integer
            GridBorder = Width - grid.Width
            'Width = Screen.Width
            Width = Screen.PrimaryScreen.Bounds.Width
            grid.Width = Width - GridBorder
            ' cmdApply is the Apply button.
            ' txtDayLabel and lblDayLabel are the caption entry boxes..
            ' fraButtons contains the Back/Print buttons.
            If cmdApply.Left + cmdApply.Width > Width Then
                Dim Diff as integer, Diff2 as integer
                Diff = cmdApply.Left - fraButtons.Left
                Diff2 = cmdApply.Left - txtDayLabel.Left

                cmdApply.Left = Width - cmdApply.Width - 60
                txtDayLabel.Left = cmdApply.Left - Diff2
                lblDayLabel.Left = txtDayLabel.Left
                fraButtons.Left = cmdApply.Left - Diff
            End If
        End If

        If True And (LicensedNoOfStores() > 1) Then
            chkMultiple.Visible = True
            chkMultiple.Checked = IIf(StoreSettings.bOneCalendar, 1, 0)
        Else
            chkMultiple.Visible = False
            chkMultiple.Checked = 0
        End If
    End Sub

    Private Sub LoadGrid(ByVal StartDate As Date, ByVal DayCount as integer, Optional ByVal AllStores As Boolean = False)
        Dim cDelDesc As clsDeliveryDateDesc
        Dim RS As ADODB.Recordset
        Dim I as integer, Rows() as integer
        Dim XIndex as integer
        Dim SaleNm As String, SaleNo As String, SalePD As String, SaleStr as integer
        Dim SaleTS As String, SaleCY As String
        Dim X as integer, StartStore as integer, EndStore as integer
        Dim Transfers As New Collection
        Dim DoSkip As Boolean

        cDelDesc = New clsDeliveryDateDesc

        If AllStores Then
            StartStore = 1
            EndStore = LicensedNoOfStores()
        Else
            StartStore = StoresSld
            EndStore = StartStore
        End If

        ReDim Rows(DayCount)

        ' Data initialization and validation
        StartDate = DateFormat(StartDate)

        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor

        grid.Clear()
        grid.Rows = 2
        grid.Cols = DayCount
        'grid.Font.Name = "Courier New"
        grid.Font = New Font("Courier New", grid.Font.Size, FontStyle.Regular)
        grid.Row = 0
        For I = 0 To DayCount - 1
            grid.Col = I
            grid.Row = 0
            grid.Text = Format(DateAdd("D", I, StartDate), "ddd, m/d/yyyy")
            If IsFurnitureStoreOfKansas Then
                'grid.ColWidth(I) = 5220
                grid.set_ColWidth(I, 5220)
            ElseIf StoreSettings.bUseTimeWindows Then
                'grid.ColWidth(I) = 5220 '3900 ' 2565
                grid.set_ColWidth(I, 5220)
            Else
                'grid.ColWidth(I) = IIf(AllStores, 2800, 2300)
                grid.set_ColWidth(I, IIf(AllStores, 2800, 2300))
            End If
            grid.Row = 1
            If cDelDesc.Load(DateAdd("D", I, StartDate), "@DeliveryDate") Then
                grid.Text = cDelDesc.DeliveryDesc
            End If
        Next

        'MousePointer = 0
        Me.Cursor = Cursors.Default
        DisposeDA(cDelDesc)
        cDelDesc = Nothing

        For X = StartStore To EndStore
            ProgressForm(0, 1, "Generating Calendar Data for Store " & X & "...", , , , ProgressBarStyle.prgIndefinite)
            RS = GetDeliveryCalendarData(StartDate, DayCount, X) ' cTable.DataAccess()
            ProgressForm(0, RS.RecordCount, "Loading Calendar Data To Form...")
            Do Until RS.EOF
                ProgressForm(RS.AbsolutePosition)
                'If IsDevelopment And rs.Fields(2).Value = "153319" Then Stop
                XIndex = RS("Index").Value
                SaleNm = GetMailLastNameByIndex(IfNullThenZero(RS("MailIndex").Value), X, True) ' RS("Name")
                SaleNo = RS("SaleNo").Value
                SaleCY = GetMailCityByIndex(IfNullThenZero(RS("MailIndex").Value), X, True)
                SalePD = IfNullThenNilString(RS("PD").Value)
                If Trim(IfNullThenNilString(RS("StopStart").Value)) <> "" Or Trim(IfNullThenNilString(RS("StopEnd").Value)) <> "" Then
                    SaleTS = "(" & Trim(RS("StopStart").Value) & "-" & Trim(RS("StopEnd").Value) & ")"
                Else
                    SaleTS = ""
                End If
                SaleStr = X
                If Not IsIn(SalePD, "P", "S", "D", "T") Then SalePD = "D"  ' Default to Delivery.

                If SalePD = "T" Then
                    On Error GoTo Skip
                    Transfers.Add("1", SaleNo)
                    On Error GoTo 0
                    If DoSkip Then GoTo SkipIt
                End If

                Rows(XIndex) = Rows(XIndex) + 1

                If (Rows(XIndex) + 1 >= grid.Rows) Then grid.Rows = Rows(XIndex) + 2
                If XIndex <= grid.Cols Then
                    grid.Col = XIndex : grid.Row = Rows(XIndex) + 1
                    grid.Text =
                            IIf(AllStores, AlignString("L" & SaleStr, 3, AlignConstants.vbAlignLeft, True) & " ", "") &
                            AlignString(SalePD, 1, AlignConstants.vbAlignLeft, True) & " " &
                            AlignString(SaleNm, 12, AlignConstants.vbAlignLeft, True) & " " &
                            AlignString(SaleNo, 6, AlignConstants.vbAlignLeft, False) & " " &
                            IIf(ShowCity, AlignString(SaleCY, 13, AlignConstants.vbAlignRight, True) & " ", "") &
                            IIf(StoreSettings.bUseTimeWindows, SaleTS, "")
                End If

SkipIt:
                DoSkip = False
                RS.MoveNext()
            Loop
            ProgressForm()
            DisposeDA(RS)
        Next

        mGridMoving = True
        UpdateGridHeight()
        mGridMoving = False
        lblDayLabel.Text = "Click column heading (date) to enter delivery zone:"
        Exit Sub

Skip:
        Err.Clear()
        DoSkip = True
        Resume Next
    End Sub

    Public Function ShowCity() As Boolean
        ShowCity = IsFurnitureStoreOfKansas
        If IsDevelopment() Then ShowCity = True
    End Function

    Public Sub TestPrinterLocations()
        Printer.FontSize = 8
        Printer.FontBold = False
        Dim I as integer
        For I = 1 To 7
            ' At size 8, an offset of 1 is 72 units.
            Printer.Print(TAB(PrinterPosition(I, 1)))
            ' Debug.Print "I="; I, "X="; Printer.CurrentX
        Next
        Printer.KillDoc()
    End Sub

    Private Function GetSaleNoFromText(ByVal Text As String, Optional ByRef StoreNo as integer = 0) As String
        If Microsoft.VisualBasic.Left(Text, 1) = "L" Then
            GetSaleNoFromText = Trim(Mid(Text, 20))
            StoreNo = Val(Mid(Text, 2, 2))
        Else
            GetSaleNoFromText = Trim(Mid(Text, 16, 6))
            StoreNo = StoresSld
        End If
    End Function

    Private Sub grid_DblClick(sender As Object, e As EventArgs) Handles grid.DblClick
        If mLoadedByForm Then Exit Sub  ' This doesn't work if we're loaded modally.
        If grid.Row < 2 Then Exit Sub
        Dim tSaleNo As String, tStore as integer
        tSaleNo = GetSaleNoFromText(grid.Text, tStore)
        If Microsoft.VisualBasic.Right(tSaleNo, 2) = " T" Then tSaleNo = Trim(Mid(tSaleNo, 1, Len(tSaleNo) - 2))
        If InStr(tSaleNo, "(") <> 0 Then
            tSaleNo = Microsoft.VisualBasic.Left(tSaleNo, InStr(tSaleNo, "(") - 1)
        End If
        If tSaleNo = "" Or tStore < 0 Then Exit Sub

        If Microsoft.VisualBasic.Left(tSaleNo, 2) = "SO" Then
            Service.QuickShowServiceCall(tSaleNo, tStore, True)
            Exit Sub ' We can't zoom to a service order yet.
        End If

        If Microsoft.VisualBasic.Left(tSaleNo, 2) = "TR" Then
            frmTransferSetup.QuickShowTransfer(Mid(tSaleNo, 3))
            Exit Sub
        End If

        If StoresSld <> tStore Then
            BillOSale.QuickShowSaleTicket(tSaleNo, tStore, True)
            Exit Sub
        End If

        ' Load up the sale in an editproof and movement-proof BillOSale.
        'Unload BillOSale
        BillOSale.Close()
        Order = "E"
        MailCheck.optSaleNo.Checked = True
        MailCheck.InputBox.Text = tSaleNo
        'MailCheck.cmdOK.Value = True
        MailCheck.cmdOK.PerformClick()
        BillOSale.Show()
        BillOSale.cmdApplyBillOSale.Enabled = False
        BillOSale.cmdCancel.Enabled = False

        BillOSale.BillOSale2_Show()
        BillOSale.cmdClear.Enabled = False
        BillOSale.cmdNextSale.Enabled = False
        BillOSale.cmdProcessSale.Enabled = False
        BillOSale.ScanDn.Enabled = False
        BillOSale.ScanUp.Enabled = False
        BillOSale.UGridIO1.GetDBGrid.AllowUpdate = False
        BillOSale.cmdMainMenu.Text = "Back"
        modProgramState.Order = ""
    End Sub

    Private Sub grid_EnterCell(sender As Object, e As EventArgs) Handles grid.EnterCell
        Dim MapEnabled As Boolean, InstrEnabled As Boolean, ManiEnabled As Boolean
        Dim I as integer, K As String

        MapEnabled = False
        InstrEnabled = False
        ManiEnabled = False

        For I = 2 To grid.Rows - 1
            K = Microsoft.VisualBasic.Left(grid.get_TextMatrix(I, grid.Col), 1)
            If K = "L" Then K = Mid(grid.get_TextMatrix(I, grid.Col), 5, 1)
            If Len(K) > 0 Then ManiEnabled = True
            If IsIn(K, "D", "S") Then ' Delivery or Service call (not P for pickup)
                MapEnabled = True
                InstrEnabled = True
            End If
        Next

        If Not AllowMap Then MapEnabled = False
        If Not AllowInstr Then InstrEnabled = False

        cmdMap.Enabled = MapEnabled
        cmdInstr.Enabled = InstrEnabled
        cmdManifest.Enabled = ManiEnabled
    End Sub

    Private Sub grid_RowColChange(sender As Object, e As EventArgs) Handles grid.RowColChange
        Dim OldRow as integer

        If mGridMoving Then Exit Sub
        If grid.Rows < 2 Then Exit Sub
        mGridMoving = True
        OldRow = grid.Row
        grid.Row = 1
        txtDayLabel.Text = grid.Text
        grid.Row = 0
        lblDayLabel.Text = "Enter delivery zone for " & grid.Text & "."

        If False Then   ' for now, too slow..
            Dim I as integer, C As Double, L As String, LC as integer
            For I = 2 To grid.Rows - 1
                L = GetSaleNoFromText(grid.get_TextMatrix(I, grid.Col), LC)
                C = C + GetCubesOnSale(L, Mid(grid.Text, 6), LC)
            Next
            lblCubes.Text = "Total Cubes: " & Format(C, "0.00")
        End If

        grid.Row = OldRow
        mGridMoving = False
    End Sub

    Private Sub lblDayLabel_DoubleClick(sender As Object, e As EventArgs) Handles lblDayLabel.DoubleClick
        LoadGrid(Now(), MAX_DAY_COLUMNS, StoreSettings.bOneCalendar)
    End Sub

    Private Sub txtDayLabel_Enter(sender As Object, e As EventArgs) Handles txtDayLabel.Enter
        SelectContents(txtDayLabel)
    End Sub

    Private Sub txtDayLabel_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDayLabel.KeyPress
        'If KeyAscii = 13 Then cmdApply.Value = True
        If e.KeyChar = Convert.ToChar(13) Then
            cmdApply_Click(cmdApply, New EventArgs)
        End If
    End Sub

    Private Sub cmdDDT_Click(sender As Object, e As EventArgs) Handles cmdDDT.Click
        Dim CD As ADODB.Recordset
        Dim D As String

        If Not DDTLicensed() Then Exit Sub

        grid.Row = 0
        D = Mid(grid.Text, 5)
        If Not IsDate(D) Then
            MessageBox.Show("Click on a day column.")
            Exit Sub
        End If

        Dim F As String
        If StoreSettings.DispatchTrackServiceCode = "" Then
            F = "Generate CSV"
        Else
            F = SelectOptionX("Select Export Option:", frmSelectOption.ESelOpts.SelOpt_ToItem + frmSelectOption.ESelOpts.SelOpt_List, "Select", "", "Generate CSV", "Upload to DispatchTrack")
        End If

        Select Case F
            Case "Generate CSV" : DDT_GenerateCSV(D)
            Case "Upload to DispatchTrack" : DDT_UploadData(D)
            Case Else : Exit Sub
        End Select
    End Sub

    Public Sub DDT_GenerateCSV(ByVal D As String)
        Dim CD As ADODB.Recordset
        Dim T As String, C As String
        Dim II as integer, IA as integer, iB as integer
        Dim Ty As String

        T = DDT_Header()

        If chkMultiple.Checked = True Then
            'IA = 1
            IA = 0
            'iB = LicensedNoOfStores()
            iB = LicensedNoOfStores() - 1
            ReDim DDTRS(0 To iB)
        Else
            IA = StoresSld
            iB = StoresSld
            ReDim DDTRS(StoresSld)
        End If

        ProgressForm(0, 1, "Generating Calendar Data...")
        'ReDim DDTRS(IA To iB)

        For II = IA To iB
            DDTRS(II) = GetDeliveryCalendarData(DateValue(D), 1, II, True, True)
        Next

        ProgressForm(0, iB, "Generating CSV...")
        For II = IA To iB
            ProgressForm(II)
            'If IsDevelopment And II = 3 Then Stop

            CD = DDTRS(II)
            Do Until CD.EOF
                Ty = CD(2).Value
                C = ""

                If Microsoft.VisualBasic.Left(Ty, 2) = "SO" Then
                    C = DDT_ExportServiceOrder(CD("Record").Value, II)
                ElseIf microsoft.VisualBasic.Left(Ty, 2) = "TR" Then
                    C = DDT_ExportTransfer(CD("Record").Value, II)
                Else
                    'If IsDevelopment And "" & CD("record").Value = "163141" Then Stop
                    'If IsDevelopment And II = 3 Then Stop
                    Dim G As CGrossMargin
                    G = New CGrossMargin
                    G.DataAccess.DataBase = GetDatabaseAtLocation(II)
                    If G.Load(CD("Record").Value, "#MarginLine") Then
                        'If IsDevelopment And G.SaleNo = "163141" Then Stop
                        If IsDeliverable(G.Status, G.Style, True) Then
                            C = DDT_ExportMarginLine(CD("Record").Value, II)
                        End If
                    End If
                    DisposeDA(G)
                End If

                If C <> "" Then
                    T = T & IIf(Microsoft.VisualBasic.Right(T, 2) = vbCrLf, "", vbCrLf) & C & vbCrLf
                End If

                CD.MoveNext()
            Loop

        Next

        ProgressForm(0, 1, "Saving file...")
        WriteFile(UIOutputFolder & "DDT.csv", T, True)

        ProgressForm()
    End Sub

    Public Sub DDT_UploadData(ByVal D As String)
        Dim CD As ADODB.Recordset
        Dim II as integer, IA as integer, iB as integer
        Dim StoreData As Collection

        If chkMultiple.Checked = True Then
            IA = 1
            iB = LicensedNoOfStores()
        Else
            IA = StoresSld
            iB = StoresSld
        End If

        StoreData = New Collection
        StoreData.Add(IA, "IA")
        StoreData.Add(iB, "IB")
        For II = IA To iB
            ProgressForm(0, 1, "Generating Calendar Data for Store " & II & "...")
            CD = GetDeliveryCalendarData(DateValue(D), 1, II, True, True)
            StoreData.Add(CD, "L" & II)
        Next
        ProgressForm()

        ProgressForm(0, 1, "Generating XML File...")
        DDT_DoUploadData(D, StoreData)
        ProgressForm()
    End Sub

End Class
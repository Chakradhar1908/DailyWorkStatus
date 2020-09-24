Imports VBRUN
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class InvCkStyle
    Public StyleCkIt As String
    Public StyleCkIt2 As String
    Public RN As Integer
    Public OnSale As Object

    Dim Search As SearchNew
    Dim LoadStyle As Integer
    Dim SearchStyle As String
    Dim Counter As Object
    Dim ListSearch() As String    ' List of styles.
    Dim ListDesc() As String      ' Descriptions..
    Dim ListQuan() As String      ' Quantities (for "In Stock")
    Dim ListCost() As String      ' OnSale prices..
    Dim RecordNo() As String      ' List of record numbers.

    Private mNewStyle As Boolean
    Private mCanceled As Boolean
    Private FoundRecord As Boolean
    Private DontUpdate As Boolean

    Public Owned As Boolean

    Dim KitSyleNo As String
    Dim ClickedList As Boolean, ClickedStyle As String
    Dim Quan As String
    Dim Vendor As String
    Dim Desc As String
    Dim StyleNo As String
    Dim Row As Integer
    Dim IsByVendor As String

    Public KitStatus As String, KitQuantity As Double

    Public MailIndex As Integer         ' These properties influence the search.
    Public LimitToStock As Boolean

    'Public ParentForm As String      ' For tracking down rogue forms, temporary!

    Dim WithEvents mDBInvKit As CDbAccessGeneral
    Dim WithEvents mDBAccess As CDbAccessGeneral

    Public Event OKClicked(ByRef Override As Boolean, ByVal Picked As String, ByVal IsNew As Boolean)
    Public Event CancelClicked(ByRef Override As Boolean)

    'Private Const FRMW_1 As Integer = 3150   ' 2925    '3345
    Private Const FRMW_1 As Integer = 440   ' 2925    '3345
    'Private Const FRMW_1b As Integer = 6000
    Private Const FRMW_1b As Integer = 450
    'Private Const FRMW_2 As Integer = 7800   ' 6500
    Private Const FRMW_2 As Integer = 600   ' 6500
    'Private Const FRMW_3 As Integer = 13950  ' 8500
    Private Const FRMW_3 As Integer = 1000  ' 8500
    'Private Const Spacing As Integer = 2
    Private Const Spacing As Integer = 2
    Private Const DescLen As Integer = 50
    Private Const QuanLen As Integer = 3
    Private Const QuanLenData As Integer = 35
    'Private Const QuanLen As Integer = 2
    Private Const CostLen As Integer = 12
    Private Const CostLenData As Integer = 30
    'Private Const CostLen As Integer = 2
    Private PopUpStyle As String
    Private OpenedMe As String

    Private Sub InvCkStyle_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '  Debug.Print "InvCkStyle loaded."
        Style.Text = ""
        RN = 0
        Me.Width = FRMW_1
        lstStyles.Visible = False
        LoadStyle = 0
        SearchStyle = ""
        ClickedList = False
        ClickedStyle = ""

        Style.MaxLength = Setup_2Data_StyleMaxLen

        UpdateSearchBox()

        If IsAuthenTeak() Then
            optSearchByDesc.Checked = True
            Width = FRMW_3
            'If Width < Screen.Width Then Left = (Screen.Width - Width) / 2
        Else
            optSearchByStyle.Checked = True
        End If
        'OpenedMe = "InvKitStock"

        '  Counter = 0
        If OrderMode("A") Then
            BillOSale.cmdProcessSale.Enabled = False
        ElseIf ReportsMode("ET") Then
            optSearchByDesc.Visible = False
            optSearchByVendor.Visible = False
            optKitVendors.Top = optSearchByVendor.Top 'Use it's placement
        ElseIf Not (ReportsMode("ET") Or OpenedMe = "InvKitStock") Then
            optKitVendors.Visible = False
            'Robert Koernke -- Tweeking Email 5/10/2017 - Simple form adjustment
            'cmdApply.Top = cmdApply.Top - 25    ----> NOTE: These five lines of code is to move the buttons a bit above. Cause optkitvendors visible is false, so to cover the blank space. But it is not required. Original design is enough. So commented the lines.
            'cmdDesc.Top = cmdDesc.Top - 25
            'cmdCancel.Top = cmdCancel.Top - 25
            'cmdBarcode.Top = cmdBarcode.Top - 25
            'Height = Height - 30
        End If
    End Sub

    Public Sub New()
        'This constructor is for form initialize event of vb6.0. In vb.net, form Initialize event is not available.
        'For form activate event, use activated event in vb.net

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        If OrderMode("A") Then
            BillOSale.cmdProcessSale.Enabled = False
        End If
    End Sub

    'NOTE: Form activate event is not available in vb.net. In this event, this code is not required. Code in this event is to do
    'the U.I design of the form only. So not required. So commented it.
    'Private Sub Form_Activate()
    '    SetCustomFrame Me, ncBasicDialog
    'On Error Resume Next
    '    '  Style.SetFocus
    'End Sub

    Public Property NewStyle As Boolean
        Get
            NewStyle = mNewStyle
        End Get
        Set(value As Boolean)
            mNewStyle = value
        End Set
    End Property

    Private Sub Style_TextChanged(sender As Object, e As EventArgs) Handles Style.TextChanged
        IsByVendor = ""
        If Microsoft.VisualBasic.Left(Style.Text, 3) = "KIT" Then HandleStyleChange() 'makes this work right
        StartTypeTimer()
        'HandleStyleChange()  '-> NOTE: Added this line temporarily. Remove the above line comment and remove this line afte complete project coding.
    End Sub

    Private Sub HandleStyleChange()
        Dim optFlag As Boolean
        optFlag = optKitVendors.Visible
        If OpenedMe = "KitStock" And Not optFlag Then
            optFlag = True
        End If
        If optSearchByStyle.Checked = True Then
            If Len(Style.Text) > Setup_2Data_StyleMaxLen Then                          '###STYLELENGTH16
                MessageBox.Show(" Maximum of " & Setup_2Data_StyleMaxLen & " Characters Allowed!", "WinCDS")
            ElseIf Microsoft.VisualBasic.Left(Style.Text, 4) = KIT_PFX And Len(Style.Text) = 4 And optFlag Then
                Width = FRMW_3
                GetKitNames()
            ElseIf Microsoft.VisualBasic.Left(Style.Text, 4) = KIT_PFX And Len(Style.Text) > 4 And optFlag Then
                Width = FRMW_3
                If Not lstStyles.Visible Then GetKitNames() : lstStyles.Visible = True
                LoadStyle = 2
                KitSyleNo = Style.Text
                FindKits2()
            ElseIf Microsoft.VisualBasic.Left(Style.Text, 4) = KIT_PFX And Len(Style.Text) >= 4 And Not optFlag Then
                MessageBox.Show("""KIT-"" is only for the 'Package Ticket builder' section, not from the inventory screen.", "WinCDS", MessageBoxButtons.OK)
                Style.Text = ""
            Else
                GetStyle()
            End If
        ElseIf optSearchByVendor.Checked = True Then
            ' Search for a vendor whose name matches Style.
            LoadMfgName(Style.Text)
        ElseIf optKitVendors.Checked = True Then
            'do nothing
        Else ' desc
            GetMatchingDescs()
        End If
    End Sub

    Private Sub StartTypeTimer()
        If tmrType.Tag = "NO" Then Exit Sub
        tmrType.Enabled = False
        tmrType.Interval = 175
        tmrType.Enabled = True
    End Sub

    Private Sub GetKitNames()
        'puts kits in help window
        KitSyleNo = Microsoft.VisualBasic.Left(Style.Text, 4)
        mDBAccess_Init(Microsoft.VisualBasic.Left(KitSyleNo, 4))
        mDBAccess.GetRecord()   ' this gets the record
        mDBAccess.dbClose()
        mDBAccess = Nothing

        mDBInvKit_Init()
        mDBInvKit_SqlSet(Microsoft.VisualBasic.Left(KitSyleNo, 4))
        mDBInvKit.GetRecord()
        mDBInvKit.dbClose()
        mDBInvKit = Nothing
    End Sub

    Public Sub FindKits2()
        ' this selects from loaded list
        lstStyles.Items.Clear()
        Dim X As Integer
        For X = 1 To Counter
            If Trim(Style.Text) = Microsoft.VisualBasic.Left(ListSearch(X), Len(Trim(Style.Text))) Then
                'lstStyles.AddItem(ListSearch(X))
                lstStyles.Items.Add(ListSearch(X))
            End If
        Next
    End Sub

    Private Sub GetStyle() 'Search main file
        If Me.Width < FRMW_1b Then Me.Width = FRMW_1b
        lstStyles.Visible = True
        '  lstStyles.Width = ScaleWidth - 120 - lstStyles.Left
        lstStyles.Items.Clear()

        If Len(Style.Text) <= 1 Then Exit Sub 'hold 2 style numbers before searching

        If LoadStyle = 1 And Microsoft.VisualBasic.Left(Style.Text, Len(SearchStyle)) = SearchStyle Then
            GetStyle2()
            Exit Sub
        End If

        On Error GoTo HandleErr
        Counter = 0
        SearchStyle = Trim(Style.Text)

        ' Search text is in Style.Text
        ' Populate results in lstStyles.
        Dim SearchObj As New CSearchNew
        Dim Query As String
        'BFH20050609 - Removed join to [2Data] because it was unecessary...  it also made BFMyers really slow on 'Workstation'
        Query = "SELECT Search.Style, Search.RN, 0 as Available, 0 as OnSale from [Search] WHERE left(Style, " & Len(SearchStyle) & ") = """ & ProtectSQL(SearchStyle) & """"
        '  Query = "SELECT Search.Style, Search.RN from [2Data] inner join Search on [2Data].Rn=Search.Rn WHERE left([2Data].Style, " & Len(SearchStyle) & ") = """ & ProtectSQL(SearchStyle) & """"
        If LimitToStock Then
            Query = "SELECT DISTINCT Search.Style, Search.Rn, 0 as Available, 0 as OnSale FROM Search inner join Detail on Search.Rn=Detail.InvRn WHERE left(Search.Style, " & Len(SearchStyle) & ") = """ & ProtectSQL(SearchStyle) & """ AND (Detail.SaleNo = '' OR Detail.SaleNo is null)"
        ElseIf chkStkOnly.Checked = True Then
            Query = "SELECT [2Data].Style, [2Data].Rn, [2Data].Available, [2Data].OnSale FROM [2Data] inner join Search on Search.Rn=[2Data].Rn WHERE Available>0 "
        End If
        '    If MailIndex > 0 Then query = query & " and [2Data].Rn in (select Rn from [" & GetDatabaseAtLocation() & "].GrossMargin where MailIndex=" & MailIndex & ")"  ' This would be much easier with merged databases.
        Query = Query & " ORDER BY Search.Style;"

        ' The following method is returning 0 records in the program, but many in the database.
        ' I'm leaving the SQL in for now.
        SearchObj.DataAccess.Records_OpenSQL(Query)

        ' This counter check is an awful hack, but what to do about it?
        ' If there were only one linked array, I'd put it in ItemData.
        ' Since there are two, we should just make them variable..
        ReDim ListSearch(SearchObj.DataAccess.Record_Count)
        ReDim ListQuan(SearchObj.DataAccess.Record_Count)
        ReDim ListCost(SearchObj.DataAccess.Record_Count)
        ReDim RecordNo(SearchObj.DataAccess.Record_Count)
        Do While SearchObj.DataAccess.Records_Available
            SearchObj.cDataAccess_GetRecordSet(SearchObj.DataAccess.RS)
            Counter = Counter + 1
            ListSearch(Counter) = SearchObj.Style
            On Error Resume Next
            ListQuan(Counter) = SearchObj.DataAccess.RS.Fields("Available").Value
            ListCost(Counter) = FormatCurrency(IfNullThenZeroCurrency(SearchObj.DataAccess.RS.Fields("OnSale").Value))
            On Error GoTo 0
            RecordNo(Counter) = Val(SearchObj.RN)
            '      lstStyles.AddItem searchobj.Style
        Loop
        DisposeDA(SearchObj)

        GetStyle2()                   ' This loads the listbox

        LoadStyle = 0
        mNewStyle = (Counter = 0)
        Exit Sub
HandleErr:
        If Err.Number = 13 Then Resume Next
        MessageBox.Show("ERROR in GetStyle: " & Err.Description & ", " & Err.Source, "WinCDS")
    End Sub

    Private Sub UpdateSearchBox()
        On Error Resume Next
        Style.Select()

        chkStkOnly.Visible = True ' False
        chkStkOnly.Checked = False
        If optSearchByDesc.Checked = True Then
            Me.Width = FRMW_3
            fraSearch.Text = "De&scription:"
            '    chkStkOnly.Visible = True
            '    chkStkOnly.Value = 0
            chkStkOnly.Top = optSearchByDesc.Top
        End If
        If optSearchByStyle.Checked = True Then
            fraSearch.Text = "&Style Number:"
            Me.Width = IIf(chkStkOnly.Checked = True, FRMW_2, FRMW_1b)
            InvCkStyle_Resize(Me, New EventArgs)
            chkStkOnly.Top = optSearchByStyle.Top
        End If
        If optSearchByVendor.Checked = True Then
            fraSearch.Text = "Vendor'&s Name:"
            Me.Width = IIf(chkStkOnly.Checked = True, FRMW_2, FRMW_1b)
            chkStkOnly.Top = optSearchByVendor.Top
        End If
        If optSearchByVendor.Checked = True Then
            LoadMfgName()
        End If
        If optKitVendors.Checked = True Then
            'LoadKitVendors
        End If
        'Style_Change
        'AddHandler Style.TextChanged, AddressOf Style_TextChanged
        Style_TextChanged(Style, New EventArgs)

    End Sub

    Private Sub LoadMfgName(Optional ByVal VenName As String = "")
        ' Set lstStyles to a list of vendors.
        Width = IIf(chkStkOnly.Checked = True, FRMW_2, FRMW_1b)
        '  lstStyles.Width = 3400

        lstStyles.Visible = True
        LoadMfgNamesIntoListBox(lstStyles, VenName, True)
    End Sub

    Private Sub GetMatchingDescs()
        Dim RS As ADODB.Recordset
        Dim SQL As String, F As String, C As Integer
        lstStyles.Visible = True
        If Width < FRMW_3 Then Width = FRMW_3
        '  lstStyles.Width = 5400
        lstStyles.Items.Clear()

        F = ProtectSQL(Style.Text)
        If Len(Style.Text) <= 1 Then Exit Sub
        SQL = ""
        SQL = SQL & "SELECT [2Data].Desc, [2Data].Style, [2Data].Rn, [2Data].Available, [2Data].OnSale "
        SQL = SQL & "FROM [2Data] "
        SQL = SQL & "WHERE Style IN (SELECT [Search].Style FROM Search) AND "
        SQL = SQL & "(Desc LIKE ""%" & (F) & "%"" OR Desc LIKE """ & (F) & "%"" OR Desc LIKE ""%" & (F) & """ OR Desc = """ & (F) & """) "

        If chkStkOnly.Checked = False Then
            SQL = SQL & "ORDER BY Style"
        Else
            SQL = SQL & "ORDER BY Available DESC"
        End If

        RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory, True)
        C = RS.RecordCount
        If C > 0 Then
            Counter = 0
            ReDim ListSearch(C)
            ReDim ListQuan(C)
            ReDim ListCost(C)
            ReDim ListDesc(C)
            ReDim RecordNo(C)
            Do Until RS.EOF
                If chkStkOnly.Checked = False Or (chkStkOnly.Checked = True And (IfNullThenZero(RS("Available").Value) > 0)) Then
                    Counter = Counter + 1
                    ListSearch(Counter) = RS("Style").Value
                    On Error Resume Next
                    ListQuan(Counter) = RS("Available").Value
                    ListQuan(Counter) = FormatCurrency(IfNullThenZeroCurrency(RS("OnSale").Value))
                    On Error GoTo 0
                    If chkStkOnly.Checked = False Then
                        'ListDesc(Counter) = ArrangeString(RS("style").Value, StyleLen) & Space(Spacing) & ArrangeString(RS("Desc").Value, DescLen) & Space(Spacing) & AlignString(FormatCurrency(RS("OnSale").Value), CostLen, AlignConstants.vbAlignRight)
                        ListDesc(Counter) = ArrangeString(RS("style").Value, StyleLen) & Chr(9) & Chr(9) & ArrangeString(RS("Desc").Value, DescLen) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & AlignString(FormatCurrency(RS("OnSale").Value), CostLen, AlignConstants.vbAlignRight)
                    Else
                        'ListDesc(Counter) = ArrangeString(RS("style").Value, StyleLen) & Space(Spacing) & ArrangeString(RS("Available").Value, QuanLen, AlignConstants.vbAlignRight) & Space(Spacing) & ArrangeString(RS("Desc").Value, DescLen) & Space(Spacing) & AlignString(FormatCurrency(RS("OnSale").Value), CostLen, AlignConstants.vbAlignRight)
                        ListDesc(Counter) = ArrangeString(RS("style").Value, StyleLen) & Chr(9) & Chr(9) & ArrangeString(RS("Available").Value, QuanLen, AlignConstants.vbAlignRight) & Chr(9) & ArrangeString(RS("Desc").Value, DescLen) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & AlignString(FormatCurrency(RS("OnSale").Value), CostLen, AlignConstants.vbAlignRight)
                    End If
                    RecordNo(Counter) = RS("RN").Value
                    End If
                    RS.MoveNext()
            Loop
            GetStyle2()
        End If

    End Sub

    Private Sub mDBAccess_Init(ByVal Tid As String)
        Tid = KitSyleNo
        mDBInvKit_Init()
        mDBAccess = New CDbAccessGeneral
        mDBAccess.dbOpen(GetDatabaseAtLocation(1))  ' Kits are only at location 1.
        If Microsoft.VisualBasic.Left(KitSyleNo, 4) = KIT_PFX And Len(KitSyleNo) = 4 Then
            mDBAccess.SQL =
      "SELECT InvKit.*" _
      & " From InvKit" _
      & " WHERE (((Left(InvKit.KitStyleNo,4))  =""" & ProtectSQL(Trim(Tid)) & """)) ORDER BY KitStyleNo"
        Else
            mDBAccess.SQL =
      "SELECT InvKit.*" _
      & " From InvKit" _
      & " WHERE (((InvKit.KitStyleNo)  =""" & ProtectSQL(Trim(Tid)) & """)) ORDER BY KitStyleNo"
        End If
    End Sub

    Private Sub mDBInvKit_Init()
        mDBInvKit = New CDbAccessGeneral
        mDBInvKit.Dispose()
        mDBInvKit.dbOpen(GetDatabaseAtLocation(1))  ' Kits are always in DB1.
    End Sub

    Public Sub mDBInvKit_SqlSet(ByVal Tid As Object)
        If Microsoft.VisualBasic.Left(KitSyleNo, 4) = KIT_PFX And Len(KitSyleNo) = 4 Then
            mDBInvKit.SQL =
      "SELECT InvKit.*" _
      & " From InvKit" _
      & " WHERE (((left(InvKit.KitStyleNo,4))  =""" & ProtectSQL(Trim(Tid)) & """))" _
      & " ORDER BY InvKit.KitStyleNo"
        Else
            mDBInvKit.SQL =
        "SELECT InvKit.*" _
        & " From InvKit" _
        & " WHERE (((InvKit.KitStyleNo)  =""" & ProtectSQL(Trim(Tid)) & """))" _
        & " ORDER BY InvKit.KitStyleNo"
        End If
    End Sub

    Private Sub GetStyle2()
        Dim Q As String
        On Error GoTo FailureInGetStyle2
        'After first search, search list only
        lstStyles.Items.Clear()

        If Len(Style.Text) = 0 Then
            ' erase complete style reset
            LoadStyle = 0
            Erase ListSearch
            Counter = 0
            Exit Sub
        End If

        Dim X As Integer
        'LockWindowUpdate(lstStyles.hwnd)  Replacement for hwnd property is Handle in vb.net
        LockWindowUpdate(lstStyles.Handle)  'Locks the listbox until all the data will be loaded in to it.

        '---------------------------
        'Finding maximum lenth of style no.  This code block is not there in vb6.0. It is added here because vb6.0 code is not enough for this requirement here.
        Dim MaxLengthStyleNo As Integer
        For X = 1 To Counter
            If Trim(Style.Text) = Microsoft.VisualBasic.Left(ListSearch(X), Len(Trim(Style.Text))) Then
                If X = 1 Then
                    MaxLengthStyleNo = Len(ListSearch(X))
                Else
                    If Len(ListSearch(X)) > MaxLengthStyleNo Then
                        MaxLengthStyleNo = Len(ListSearch(X))
                    End If
                End If
            End If
        Next
        '-----------------------------
        Dim Stylenumber As String, NotMax As Boolean
        For X = 1 To Counter
            If optSearchByDesc.Checked = True Then
                'lstStyles.AddItem(ListDesc(X))
                lstStyles.Items.Add(ListDesc(X))
            ElseIf Trim(Style.Text) = Microsoft.VisualBasic.Left(ListSearch(X), Len(Trim(Style.Text))) Then
                If chkStkOnly.Checked = True Then
                    Stylenumber = Trim(ListSearch(X))
                    'Stylenumber = Stylenumber & Space(MaxLengthStyleNo - Len(Stylenumber))
                    If Len(Stylenumber) < MaxLengthStyleNo Then
                        Stylenumber = Stylenumber & Space((MaxLengthStyleNo - Len(Stylenumber)) + 1)
                        NotMax = True
                    Else
                        '   Stylenumber = Stylenumber & Space(MaxLengthStyleNo - Len(Stylenumber))
                        NotMax = False
                    End If
                    'Stylenumber = Stylenumber & New String("t"c, (MaxLengthStyleNo - Len(Stylenumber)) + 1)
                    'Q = Stylenumber & Space(5) & Len(Stylenumber)
                    'Q = ArrangeString(ListSearch(X), StyleLen) & Space(Spacing) & AlignString(ListQuan(X), QuanLen) '& Space(Spacing) & AlignString(ListCost(X), CostLen)

                    If NotMax = True Then
                        Q = ArrangeString(Stylenumber, StyleLen) & AlignString(ListQuan(X), QuanLenData) & AlignString(ListCost(X), CostLenData)
                        'Q = ListSearch(X) & ListQuan(X)
                    Else
                        Q = ArrangeString(Stylenumber, (StyleLen - 1)) & AlignString(ListQuan(X), QuanLenData) & AlignString(ListCost(X), CostLenData)
                    End If
                Else
                    Q = ArrangeString(ListSearch(X), StyleLen)
                End If
                lstStyles.Items.Add(Q)
            End If
        Next
        'LockWindowUpdate 0
        LockWindowUpdate(IntPtr.Zero)   'Release the lock after data loading completed.
        Exit Sub

FailureInGetStyle2:
        MsgBox("Error in InvCkStyle.GetStyle2" & vbCrLf & "[" & Err.Number & "] " & Err.Description)
    End Sub

    Private ReadOnly Property StyleLen() As Integer
        Get
            StyleLen = Setup_2Data_StyleMaxLen
        End Get
    End Property

    Public WriteOnly Property CallingForm As String
        Set(value As String)
            OpenedMe = value
        End Set
    End Property

    Public Property Canceled As Boolean
        Get
            Canceled = mCanceled
        End Get
        Set(value As Boolean)
            mCanceled = value
        End Set
    End Property

    Public Sub GetSpeechInputMode(ByRef Result As Boolean, ByVal SIType As String, ByVal CtrlName As String)
        If SIType = "spell" And CtrlName = "Style" Then Result = False
    End Sub

    Private Sub cmdBarcode_Click(sender As Object, e As EventArgs) Handles cmdBarcode.Click
        'MousePointer = vbHourglass
        Me.Cursor = Cursors.WaitCursor
        Style.Text = GetNextBarcode(Me)
        HandleStyleChange()
        'MousePointer = vbDefault
        Me.Cursor = Cursors.Default
        If Style.Text <> "" Then cmdApply.PerformClick()
    End Sub

    Private Sub FinishSelect()
        Dim Override As Boolean, ST As String

        PreviewItemByStyle()
        ST = IIf(ClickedList, ClickedStyle, Style.Text)

        If Microsoft.VisualBasic.Left(ST, 4) = KIT_PFX And OrderMode("A") Then 'New sales only "Kits"
            'Unload Me
            Me.Close()
            Exit Sub
        End If

        mCanceled = False
        mNewStyle = Not GetStyleFound(ST)
        Override = False

        RaiseEvent OKClicked(Override, ST, mNewStyle)
        If Override Then Exit Sub

        If ST = "" Then
            If Not ReportsMode("CS") Then
                'InvCkStyle.Show
                If Visible Then Style.Select()
            End If
            Exit Sub
        End If

        If ReportsMode("ET") Then  'Edit Packages (kits)..
            'edit packages
            If PackagePrice.FindKits(ST) Then
                'Unload Me
                Me.Close()
            End If
            Exit Sub
        End If

        If Not OrderMode("F", "Credit") Then
            If Not InvenMode("A") And NewStyle Then
                ' Error?
                MessageBox.Show("Invalid Style Number.  Please try again.", "WinCDS")
                Exit Sub
            Else
                If InvenMode("A") Then
                    InvenA.Style.Text = Style.Text  'ST
                    StyleCkIt = Style.Text  'ST
                Else
                    InvenA.Style.Text = ST
                    StyleCkIt = ST
                End If
            End If
        End If

        If OrderMode("F") Then 'Stock Preview
            Top = 6200
            OrdPreview.LoadItemByRN(RN)

            On Error Resume Next
            OrdPreview.Show()
            'Unload Me
            Me.Close()
            Exit Sub
        End If

        If Not OrderMode("Credit") Then
            If Not InvenA.CheckBal Then
                ' The selection failed, so leave this object visible?
                ' This opens us up to a lot of inconsistencies.
                Exit Sub
            End If
        End If

        If False Then
            '  ElseIf InvenMode("B") Then
            '    InvenA.Show
            '    InvenA.Caption = "Changing Price Structure"
        ElseIf InvenMode("D") Then
            InvenA.Text = "Processing Factory Shipments"
            InvStkRec.Text = "Processing Factory Shipments"
            InvStkRec.lblInvoice.Text = " Inv / Ack Number"
        ElseIf InvenMode("T") Then
            InvenA.Text = "Processing Store Transfers"
            InvStkRec.lblInvoice.Text = "Transfer No:"
            InvStkRec.Text = "Processing Store Transfers"
        ElseIf InvenMode("H") Then
            InvenA.Text = "Inventory Maintenance"
            InvenA.Show()  ' This isn't in the right spot, but it patches..
        ElseIf InvenMode("A") Then      'Returns from finding record
            'Unload InvStkRec
            InvStkRec.Close()
        End If

        If InvenMode("B", "D", "H", "E") Then
            If InvenA.InitialStyle = "" Then
                InvADefault.Show()
            End If
        End If

        If InvenMode("D") Then
            If InvenA.InitialStyle <> "" Then
                InvStkRec.Show()
            End If
        End If

        StyleCkIt = ST

        If InvenMode("A") And Trim(ST) = Trim(Search.Style) Then
            ' Do nothing.
        Else
            If OrderMode("Credit") Or InvenMode("T") Then
                Hide()
            Else
                'Unload Me
                Me.Close()
            End If
        End If

        Exit Sub

HandleErr:
        If Err.Number = 13 Then Resume Next
        MessageBox.Show("ERROR in FinishSelect: " & Err.Description & ", " & Err.Source, "WinCDS")
    End Sub

    Private Function GetStyleFound(Optional ByVal ST As String = "_") As Boolean
        GetStyleFound = False
        On Error GoTo HandleErr
        Counter = 0

        If ST = "_" Then ST = Style.Text

        If Microsoft.VisualBasic.Left(ST, 4) = KIT_PFX Or optKitVendors.Checked = True Then Exit Function  'added 02-08-2003

        Dim SearchObj As New CSearchNew
        If SearchObj.Load(Trim(ST)) Then
            RN = SearchObj.RN
            GetStyleFound = True
        Else
            GetStyleFound = False
        End If
        SearchObj.Dispose()
        SearchObj = Nothing

        Exit Function
HandleErr:
        If Err.Number = 13 Then Resume Next
        MessageBox.Show("ERROR in GetStyleFound: " & Err.Description & ", " & Err.Source, "WinCDS")
    End Function

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Dim Override As Boolean
        mCanceled = True
        Override = False
        RaiseEvent CancelClicked(Override)
        If (Override = True) Then Exit Sub
        Counter = 0

        If OrderMode("F") Then
            If Not Owned Then modProgramState.Order = ""
            'Unload Me
            Me.Close()
            If Not Owned Then MainMenu.Show()
            Exit Sub
        End If

        If Order <> "" Then
            modProgramState.Inven = ""
            'Unload Me
            Me.Close()
            Exit Sub
        End If

        If Inven <> "" Then
            modProgramState.Order = ""
            modProgramState.Inven = ""
            lstStyles.Items.Clear()
            'Unload InvDefault
            InvDefault.Close()
            'Unload Me
            Me.Close()
            'Unload InvenA
            InvenA.Close()
            'Unload OrdPreview
            OrdPreview.Close()
            'Unload PackagePrice
            PackagePrice.Close()
            'Unload InvKitStock
            InvKitStock.Close()
            MainMenu.Show()
            Exit Sub
        End If

        'Unload PackagePrice
        PackagePrice.Close()
        'Unload InvKitStock  ' This is awful!  InvKitStock should be creating an instance.
        InvKitStock.Close()
        'Unload Me
        Me.Close()
        MainMenu.Show()
        modProgramState.Reports = ""
    End Sub

    Private Sub InvCkStyle_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'If KeyCode = vbKeyF2 Then cmdBarcode_Click()
        If e.KeyCode = Keys.F2 Then cmdBarcode_Click(cmdBarcode, New EventArgs)
    End Sub

    Private Sub InvCkStyle_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        Dim X As Integer
        If Me.Left + Me.Width > Screen.PrimaryScreen.Bounds.Width Then Left = Screen.PrimaryScreen.Bounds.Width - Me.Width
        'X = ScaleWidth - lstStyles.Left - 120
        'X = Me.ClientSize.Width - lstStyles.Left - 120
        X = Me.ClientSize.Width - lstStyles.Left - 10
        If X > 0 Then lstStyles.Width = X
        'If ScaleHeight > lstStyles.Top Then lstStyles.Height = ScaleHeight - 2 * lstStyles.Top
        If Me.ClientSize.Height > lstStyles.Top Then lstStyles.Height = Me.ClientSize.Height - 2 * lstStyles.Top
        'On Error Resume Next
        'If Left + Width > Screen.Width Then Left = (Screen.Width - Width) / 2
        If Me.Left + Me.Width > Screen.PrimaryScreen.Bounds.Width Then Me.Left = (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2
        DoCaptions()

        CenterForm(Me)
    End Sub

    Private Sub DoCaptions()
        lblCaptions.Width = lstStyles.Width
        Select Case Me.Width
            Case FRMW_1 : lblCaptions.Text = ""
            Case FRMW_1b
                If optSearchByVendor.Checked = True Then
                    lblCaptions.Text = "Vendor"
                ElseIf optKitVendors.Checked = True Then
                    lblCaptions.Text = "Kit Vendors"
                Else
                    lblCaptions.Text = "Style"
                End If
          'lblCaptions = IIf(optSearchByVendor, "Vendor", "Style")
            Case FRMW_2
                lblCaptions.Text = ArrangeString("Style", StyleLen) & Space(Spacing) & AlignString("Qty", QuanLen) & Space(Spacing) & AlignString("On Sale", CostLen)

                'Q = ArrangeString(ListSearch(X), StyleLen) & Space(Spacing) & AlignString(ListQuan(X), QuanLen) & Space(Spacing) & AlignString(ListCost(X), CostLen)
            Case FRMW_3
                If optSearchByDesc.Checked = True Then
                    If chkStkOnly.Checked = False Then
                        lblCaptions.Text = ArrangeString("Style", StyleLen) & Space(Spacing) & ArrangeString("Description", DescLen) & Space(Spacing) & AlignString("On Sale", CostLen)
                    Else
                        lblCaptions.Text = ArrangeString("Style", StyleLen) & Space(Spacing) & AlignString("Qty", QuanLen) & Space(Spacing) & ArrangeString("Description", DescLen) & Space(Spacing) & AlignString("On Sale", CostLen)
                    End If
                Else
                    If chkStkOnly.Checked = False Then
                        lblCaptions.Text = ArrangeString("Style", StyleLen) & Space(Spacing) & ArrangeString("Description", DescLen)
                        'lblCaptions.Text = ArrangeString("Style", StyleLen) & Space(1) & ArrangeString("Description", DescLen)
                    Else
                        'lblCaptions.Text = ArrangeString("Style", StyleLen) & Space(Spacing) & AlignString("Qty", QuanLen) & Space(Spacing) & ArrangeString("Description", DescLen) & Space(Spacing) & AlignString("On Sale", CostLen)
                        lblCaptions.Text = ArrangeString("Style", StyleLen - 3) & AlignString("Qty", QuanLen) & Space(Spacing) & ArrangeString("Description", DescLen) & Space(Spacing) & AlignString("On Sale", CostLen)
                    End If
                End If
            Case Else : lblCaptions.Text = "Style"
        End Select
    End Sub

    Private Sub lstStyles_Enter(sender As Object, e As EventArgs) Handles lstStyles.Enter
        On Error Resume Next
        If lstStyles.SelectedIndex = -1 Then lstStyles.SelectedIndex = 0
        'lstStyles.Selected(lstStyles.ListIndex) = True
        lstStyles.SetSelected(lstStyles.SelectedIndex, True)
    End Sub

    Private Sub lstStyles_KeyPress(sender As Object, e As KeyPressEventArgs) Handles lstStyles.KeyPress
        'If KeyAscii = vbKeyReturn Then
        '    lstStyles_DblClick
        'End If

        If e.KeyChar = Convert.ToChar(13) Then
            'AddHandler lstStyles.DoubleClick, AddressOf lstStyles_DoubleClick
            lstStyles_DoubleClick(lstStyles, New EventArgs)
        End If
    End Sub

    Private Sub lstStyles_DoubleClick(sender As Object, e As EventArgs) Handles lstStyles.DoubleClick
        If Style.Visible = False Then
            If lstStyles.Items.Count > 0 Then
                'Style.Text = Trim(Left(lstStyles.List(lstStyles.ListIndex), 16)) '###STYLELENGTH16
                'Style.Text = Trim(Microsoft.VisualBasic.Left(lstStyles.GetItemText(lstStyles.SelectedIndex), 16)) '###STYLELENGTH16
                Style.Text = Trim(Microsoft.VisualBasic.Left(lstStyles.SelectedItem.ToString, 16)) '###STYLELENGTH16
            End If
        End If

        DoSelect()
    End Sub

    Private Sub LoadKitRecord()
        mDBInvKit_Init()
        'If Not optKitVendors Then
        mDBInvKit_SqlSet(Trim(KitSyleNo))
        'Else
        ' mDBInvKit_SqlSet2 KIT_PFX, ClickedStyle
        'End If
        mDBInvKit.GetRecord()
        DisposeDA(mDBInvKit)
    End Sub

    Private Function GetKitStatus() As Boolean
        Dim OW As Integer
        OW = Width
        Width = FRMW_1

        If IsUFO() Or IsFurnOne() Then ' Sets kit status
            KitStatus = "LAW"
        ElseIf IsLapeer() Or IsPuritan() Or IsRockyMountain() Then
            KitStatus = "SO"
        Else
            KitStatus = "ST"
        End If

        GetKitStatus = SelectKitStatusAndQuantity(KitStatus, KitQuantity, Style.Text)

        Width = OW
    End Function

    Public Sub DoSelect()
        On Error GoTo HandleErr
        Dim IsKit As Boolean

        If optKitVendors.Checked = True And Width = FRMW_3 And Microsoft.VisualBasic.Left(Style.Text, 4) <> KIT_PFX Then
            Style.Text = KIT_PFX
        End If
        IsKit = Microsoft.VisualBasic.Left(Style.Text, 4) = KIT_PFX

        If Not optSearchByDesc.Checked = True Then
            ClickedList = True
            'ClickedStyle = Trim(Left(lstStyles.List(lstStyles.ListIndex), 16)) '###STYLELENGTH16
            ClickedStyle = Trim(Microsoft.VisualBasic.Left(lstStyles.GetItemText(lstStyles.SelectedItem), 16)) '###STYLELENGTH16
        End If

        If optKitVendors.Checked = True And Width <> FRMW_3 Then
            'Here do the work to open up the form - Robert
            Width = FRMW_3
            KitSyleNo = KIT_PFX
            mDBAccess_Init(KIT_PFX)
            mDBAccess.GetRecord()   ' this gets the record
            mDBAccess.dbClose()
            mDBAccess = Nothing

            mDBInvKit_Init()
            mDBInvKit_SqlSet2(KIT_PFX, ClickedStyle)
            mDBInvKit.GetRecord()
            mDBInvKit.dbClose()
            mDBInvKit = Nothing

        ElseIf optKitVendors.Checked = True Then
            tmrType.Tag = "NO" 'Made more changes here 5/16/2017 Robert
            'Style.Text = Trim(Microsoft.VisualBasic.Left(lstStyles.GetItemText(lstStyles.SelectedIndex), 16)) '###STYLELENGTH16
            Style.Text = Trim(Microsoft.VisualBasic.Left(lstStyles.SelectedItem, 16))
            tmrType.Tag = ""
            DoApply()
            Exit Sub
        End If

        If optSearchByStyle.Checked = True Or optSearchByDesc.Checked = True Or (optKitVendors.Checked = True And IsKit) Then
            If optSearchByDesc.Checked = True Then Width = FRMW_3
            '    lstStyles.Width = 5400

            If Not IsKit Then
                Style.Text = ListSearch(lstStyles.SelectedIndex + 1)
            Else
                tmrType.Tag = "NO"
                Style.Text = Trim(Microsoft.VisualBasic.Left(lstStyles.GetItemText(lstStyles.SelectedIndex), 16)) '###STYLELENGTH16
                tmrType.Tag = ""
                HandleStyleChange()
            End If
            '    Style.Text = Trim(Left(lstStyles.List(lstStyles.ListIndex), 16)) '###STYLELENGTH16

            If Not IsKit And ReportsMode("CS") Then
                MessageBox.Show("You may only select kits from this list.", "Warning")
                Exit Sub
            End If

            If IsKit And OrderMode("A") Or ReportsMode("CS") Then
                'Set mDBInvKit.db = mDBAccess.db

                If OrderMode("A", "Credit") Then    'BFH20100917 get kit status for New Sales and Adjustments
                    If Not GetKitStatus() Then Exit Sub
                End If

                LoadKitRecord()
                'Unload Me
                Me.Close()
                Exit Sub
            End If

            If IsKit And OrderMode("Credit") Then
                If Not GetKitStatus() Then Exit Sub
            End If

            FinishSelect()

            ClickedList = False
            ClickedStyle = ""

            If OrderMode("Credit") Then Exit Sub  'added 04-22-01
        ElseIf optSearchByVendor.Checked = True Then
            ' The form may already be unloaded at this point, and the next line causes it to reload!
            'IsByVendor = lstStyles.itemData(lstStyles.ListIndex)
            IsByVendor = CType(lstStyles.SelectedItem, ItemDataClass).ItemData
            'VendorItemsSearch lstStyles.itemData(lstStyles.ListIndex)
            VendorItemsSearch(CType(lstStyles.SelectedItem, ItemDataClass).ItemData)
            Exit Sub
        End If
        Exit Sub

HandleErr:
        Select Case Err().Number
            Case 5 : Resume Next ' Style.SetFocus             'Click on a bad spot  **invalid procedure call
            Case 13 : Style.Select()
        End Select
    End Sub

    Public Sub mDBInvKit_SqlSet2(ByVal Tid As String, ByVal KitSKU As String)
        If (Microsoft.VisualBasic.Left(KitSyleNo, 4) = KIT_PFX And Len(KitSyleNo) = 4) Or optKitVendors.Checked = True Then
            mDBInvKit.SQL =
        "SELECT InvKit.*" _
        & " From InvKit" _
        & " WHERE (((left(InvKit.KitStyleNo,4))  =""" & ProtectSQL(Trim(Tid)) & """)) and KitSKU = """ & ProtectSQL(Trim(KitSKU)) & """" _
        & " ORDER BY InvKit.KitStyleNo"
        End If
    End Sub

    Public Sub DoApply()
        StyleCkIt = Style.Text   ' BFH20090701
        cmdCancel.Enabled = False
        cmdDesc.Enabled = False

        If Microsoft.VisualBasic.Left(Style.Text, 4) = KIT_PFX Then
            If frmBarcode.FromBarCodeReader = True Then
                mDBAccess_Init(KitSyleNo) 'needed for Mini-Scanner
                mDBAccess.GetRecord()
                mDBAccess.dbClose()
                mDBAccess = Nothing
                frmBarcode.FromBarCodeReader = False
            End If

            'BFH20100917 get kit status for New Sales and Adjustments
            If OrderMode("A", "Credit") Then
                If Not GetKitStatus() Then Exit Sub
            End If

            LoadKitRecord()
        End If

        ' Hacky protection from adding accidental SS records to customer adjustments.
        If Trim(Style.Text) = "LABOR" Then Style.Text = "LAB"
        If Trim(Style.Text) = "DELIVERY" Then Style.Text = "DEL"

        FinishSelect()                    ' Unloads form..
        '  If False Then
        If Not IsFormLoaded("BillOSale") Then
            cmdCancel.Enabled = True        ' Reloads form, clears RN
            cmdDesc.Enabled = True
        End If
    End Sub

    Private Sub VendorItemsSearch(ByVal VendorNo As String)
        ' A vendor has been selected.  The vendor's ID number is passed in here.
        ' Populate lstStyles with all matching items, and switch back to Style Search.
        Dim DataObj As CInvRec, C As Integer
        Dim SQL As String
        DataObj = New CInvRec
        On Error Resume Next
        lstStyles.Items.Clear()
        Width = FRMW_3
        '  lstStyles.Width = 5400

        If Len(VendorNo) < 3 Then VendorNo = New String("0"c, 3 - Len(VendorNo)) & VendorNo
        If chkStkOnly.Checked = True Then
            SQL = "SELECT [2Data].* FROM [2Data] INNER JOIN Search on [2Data].Rn=Search.Rn WHERE Available>0 AND [2Data].VendorNo=""" & ProtectSQL(VendorNo) & """ ORDER BY [2Data].Style"
        Else
            SQL = "SELECT [2Data].* FROM [2Data] INNER JOIN Search on [2Data].Rn=Search.Rn WHERE [2Data].VendorNo=""" & ProtectSQL(VendorNo) & """ ORDER BY [2Data].Style"
        End If
        DataObj.DataAccess.Records_OpenSQL(SQL)


        C = DataObj.DataAccess.Record_Count
        ReDim ListSearch(C)
        ReDim ListDesc(C)
        ReDim RecordNo(C)
        Counter = 0
        '----------------
        'for desc length
        Dim rstDesc As ADODB.Recordset
        Dim cnt As Integer = 1
        Dim DescLength As Integer
        Dim BlankSpace As Integer
        rstDesc = DataObj.DataAccess.RS.Clone(ADODB.LockTypeEnum.adLockReadOnly)
        If rstDesc.BOF = False And rstDesc.EOF = False Then
            Do While Not rstDesc.EOF
                If cnt = 1 Then
                    DescLength = Len(rstDesc("desc").Value)
                    cnt = 2
                Else
                    If Len(rstDesc("desc").Value) > DescLength Then
                        DescLength = Len(rstDesc("desc").Value)
                    End If
                End If
                rstDesc.MoveNext()
            Loop
        End If
        rstDesc.Close()
        rstDesc = Nothing


        'LockWindowUpdate(lstStyles.hwnd)
        LockWindowUpdate(lstStyles.Handle)
        Do While DataObj.DataAccess.Records_Available
            DataObj.cDataAccess_GetRecordSet(DataObj.DataAccess.RS)
            ' Say the maximum reasonable description is about 2500..
            ' We want a number of tabs equal to (2500-textwidth(x))/textwidth(vbtab)
            Dim Str As String, Spaces As Integer
            'str = AlignString(DataObj.Style, 16, vbAlignLeft, False)  '###STYLELENGTH16
            Str = DataObj.Style
            Spaces = Int((1500 - Printer.TextWidth(Str)) / Printer.TextWidth(" "))
            If Spaces < 0 Then Spaces = 0
            If Len(Str) + Spaces < 16 Then Spaces = 16 - Len(Str) '###STYLELENGTH16
            '      Debug.Print "Adding " & spaces & " spaces to " & str & "(" & TextWidth(str) & ")."
            '      If Left(str, 2) = "NI" Then Stop

            If chkStkOnly.Checked = True Then
                BlankSpace = DescLength - Len(DataObj.Desc)

                'lstStyles.AddItem ArrangeString(DataObj.Style, StyleLen) & Space(Spacing) & AlignString(DataObj.Available, QuanLen, vbAlignRight) & Space(Spacing) & ArrangeString(DataObj.Desc, DescLen) & Space(Spacing) & AlignString(FormatCurrency(DataObj.OnSale), CostLen)
                'Dim Lststyle As String, LstAvailable As String
                'Lststyle = ArrangeString(DataObj.Style, StyleLen)
                'Lststyle = String.Format("{0,-20}", DataObj.Style)
                'LstAvailable = AlignString(DataObj.Available, QuanLen)
                'LstAvailable = String.Format("{0,20}", DataObj.Available)
                'lstStyles.Items.Add(Lststyle & LstAvailable)
                'lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Space(30) & AlignString(DataObj.Available, QuanLen, AlignConstants.vbAlignRight) & Space(10) & ArrangeString(DataObj.Desc, DescLen) & Space(80) & AlignString(FormatCurrency(DataObj.OnSale), CostLen))
                'lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Space(30) & ArrangeString(DataObj.Available, QuanLen, AlignConstants.vbAlignRight) & Space(10) & ArrangeString(DataObj.Desc, DescLen) & Space(80) & ArrangeString(DataObj.OnSale, CostLen))
                'lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Chr(9) & ArrangeString(DataObj.Available, QuanLen) & Chr(9) & ArrangeString(DataObj.Desc, DescLen) & Chr(9) & AlignString(FormatCurrency(DataObj.OnSale), CostLen))
                'lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Chr(9) & Chr(9) & AlignString(DataObj.Available, QuanLen, AlignConstants.vbAlignRight) & Chr(9) & ArrangeString(DataObj.Desc, DescLen) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & AlignString(FormatCurrency(DataObj.OnSale), CostLen))
                'lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Chr(9) & AlignString(DataObj.Available, QuanLen, AlignConstants.vbAlignRight) & Chr(9) & ArrangeString(DataObj.Desc & New String("          ", DescLength - Len(DataObj.Desc)), DescLen) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & AlignString(FormatCurrency(DataObj.OnSale), CostLen))
                lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Chr(9) & AlignString(DataObj.Available, QuanLen, AlignConstants.vbAlignRight) & Chr(9) & ArrangeString(DataObj.Desc & Space(BlankSpace), DescLen) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & AlignString(FormatCurrency(DataObj.OnSale), CostLen))

                'If Len(DataObj.Desc) = DescLength Then
                '    lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Chr(9) & AlignString(DataObj.Available, QuanLen, AlignConstants.vbAlignRight) & Chr(9) & ArrangeString(DataObj.Desc, DescLen) & Chr(9) & AlignString(FormatCurrency(DataObj.OnSale), CostLen))
                'Else
                '    lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Chr(9) & AlignString(DataObj.Available, QuanLen, AlignConstants.vbAlignRight) & Chr(9) & ArrangeString(DataObj.Desc & Space(BlankSpace), DescLen) & Chr(9) & Chr(9) & AlignString(FormatCurrency(DataObj.OnSale), CostLen))
                'End If
            Else
                    'lstStyles.AddItem ArrangeString(DataObj.Style, StyleLen) & Space(Spacing) & ArrangeString(DataObj.Desc, DescLen)
                    'lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Space(Spacing) & ArrangeString(DataObj.Desc, DescLen))
                    'lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Space(30) & ArrangeString(DataObj.Desc, DescLen))
                    lstStyles.Items.Add(ArrangeString(DataObj.Style, StyleLen) & Chr(9) & ArrangeString(DataObj.Desc, DescLen))
            End If
            Counter = Counter + 1
            ListSearch(Counter) = DataObj.Style
            RecordNo(Counter) = DataObj.RN
        Loop
        'LockWindowUpdate 0
        LockWindowUpdate(IntPtr.Zero)

        DontUpdate = True
        optSearchByVendor.Checked = False
        optSearchByStyle.Checked = True
        chkStkOnly.Top = optSearchByStyle.Top
        DontUpdate = False

        DisposeDA(DataObj)
    End Sub

    Private Sub cmdDesc_Click(sender As Object, e As EventArgs) Handles cmdDesc.Click
        'Dim s As String = New String("", 10)
        'MessageBox.Show(s)
        'MessageBox.Show(Len(s))
        'Exit Sub
        Me.Width = FRMW_3
        '  lstStyles.Width = 5400
        PrintDesc()
    End Sub

    Private Sub PrintDesc()
        ' If Len(Style)=1, populate lstStyles with 1-Counter items.
        ' Else, populate lstStyles with 1-Counter matching items.
        ' But what is Counter, and what do we do with RecordNo(X)?
        ' Counter is a module-wide count of elements.
        ' RecordNo is an array pointing to file positions!
        'Dim ArStyle(), ArDesc(), ArSale() As String
        'Dim i As Integer

        lstStyles.Items.Clear()
        Dim Query As String
        Dim DataObj As CInvRec
        DataObj = New CInvRec
        Query = "SELECT * FROM [2DATA] WHERE left(style, " & Len(Trim(Style.Text)) & ")=""" & ProtectSQL(Trim(Style.Text)) & """ and Rn in (SELECT RN from Search) ORDER BY Style"

        DataObj.DataAccess.Records_OpenSQL(Query)
        '.DataAccess.Records_OpenIndexLike Trim(Style.Text)
        'LockWindowUpdate lstStyles.hwnd
        'LockWindowUpdate(lstStyles.Handle)      -------> LockWindowUpdate is for drag/drop requirement. Here it is not required to load data in to listbox.
        Do While DataObj.DataAccess.Records_Available
            DataObj.cDataAccess_GetRecordSet(DataObj.DataAccess.RS)
            'ReDim Preserve ArStyle(i)
            'ReDim Preserve ArDesc(i)
            'ReDim Preserve ArSale(i)
            'ArStyle(i) = DataObj.Style
            'ArDesc(i) = DataObj.Desc
            'ArSale(i) = DataObj.OnSale

            'lstStyles.Items.Add(DataObj.Style & New String("", 16 - Len(DataObj.Style)) & DataObj.Desc & New String("", 10 - Len(FormatCurrency(DataObj.OnSale))) & FormatCurrency(DataObj.OnSale))
            lstStyles.Items.Add(DataObj.Style & Chr(9) & Chr(9) & DataObj.Desc & Chr(9) & FormatCurrency(DataObj.OnSale))
            'lstStyles.AddItem DataObj.Style & DataObj.Desc

            'i = i + 1
        Loop
        'lstStyles.Items.Add(ArStyle)
        'LockWindowUpdate 0
        'LockWindowUpdate(IntPtr.Zero)
        DisposeDA(DataObj)

        Exit Sub
HandleErr:
        If Err.Number = 13 Then Resume Next
        MessageBox.Show("ERROR in PrintDesc: " & Err.Description & ", " & Err.Source, "WinCDS")
    End Sub

    Private Sub mDBInvKit_GetRecordEvent(RS As ADODB.Recordset) Handles mDBInvKit.GetRecordEvent
        Dim I As Integer
        Dim X As Integer, isCS As Boolean
        Dim KitLandedTotal As Decimal, KitOnSaleTotal As Decimal, ItemLanded As Decimal

        X = BillOSale.X
        FoundRecord = True
        isCS = ReportsMode("CS")

        'loads the styles that begin with kit-
        Row = 0
        Counter = 0
        lstStyles.Items.Clear()

        If Microsoft.VisualBasic.Left(KitSyleNo, 4) = KIT_PFX And Len(KitSyleNo) = 4 Then
            ReDim ListSearch(RS.RecordCount)
            ReDim RecordNo(RS.RecordCount)
            LockWindowUpdate(lstStyles.Handle)
            Do While Counter < 3000 And Not RS.EOF
                KitSyleNo = IfNullThenNilString(RS("KitStyleNo").Value)
                'pad kit with 16 blanks
                KitSyleNo = KitSyleNo & New String(" ", 16 - Len(KitSyleNo)) '###STYLELENGTH16
                Counter = Counter + 1
                ListSearch(Counter) = KitSyleNo & Space(Spacing) & IfNullThenNilString(RS("Heading").Value)
                'lstStyles.Items.Add(KitSyleNo & Space(Spacing) & IfNullThenNilString(RS("Heading").Value))
                lstStyles.Items.Add(KitSyleNo & Chr(9) & Chr(9) & IfNullThenNilString(RS("Heading").Value))
                RS.MoveNext()
            Loop
            'LockWindowUpdate 0
            LockWindowUpdate(IntPtr.Zero)
            Exit Sub
        End If

        If ReportsMode("ET") Or modProgramState.Inven <> "" And modProgramState.Order <> "" Then Exit Sub
        'If cmdApply.Value = False Then Exit Sub

        If isCS Then
            ' Kit Stock Lookup
            InvKitStock.Style = RS("KitStyleNo").Value
            InvKitStock.Desc = RS("Heading").Value
            InvKitStock.Landed = RS("Landed").Value
            InvKitStock.List = RS("List").Value
            InvKitStock.PackPrice = RS("PackPrice").Value
            InvKitStock.Comments = RS("MemoArea").Value
        End If

        BillOSale.KitLines = 0

        If IsFormLoaded("frmKitLevels") Then
            KitLandedTotal = frmKitLevels.KitCost("Landed")
            KitOnSaleTotal = frmKitLevels.KitCost("OnSale")
        End If

        If frmKitLevels.IsfrmKitLevelsHide = True Then
            KitLandedTotal = frmKitLevels.KitCost("Landed")
            KitOnSaleTotal = frmKitLevels.KitCost("OnSale")
        End If


        For I = 1 To Setup_MaxKitItems
            StyleNo = Trim(IfNullThenNilString(RS("Item" & I).Value))
            If StyleNo = "" Then GoTo ExitHere
            BillOSale.KitLines = I
            RN = IfNullThenZero(RS("Item" & I & "Rec").Value)
            If isCS Then
                Quan = IfNullThenNilString(RS("Quan" & I).Value) * IIf(KitQuantity = 0, 1, KitQuantity) ' kit quantity is 0 for kit stock lookup
            Else
                If Not OrderMode("Credit") Then
                    X = BillOSale.X
                    BillOSale.SetQuan(X, frmKitLevels.ItemQuantityByStyle(IfNullThenNilString(RS("Item" & I).Value)))
                    If StoreSettings.bShowPackageItemPrices Then
                        BillOSale.SetPrice(X, RS("PackPrice").Value * frmKitLevels.KitCost(Style:=IfNullThenNilString(RS("Item" & I).Value)) / KitLandedTotal)
                    End If
                End If
            End If
            GetRecord   ' Loads BoS2 form with data
            If Not isCS And Not OrderMode("Credit") Then
                BillOSale.SetLoc(X, frmKitLevels.ItemLocByStyle(IfNullThenNilString(RS("Item" & I).Value)))
            End If
        Next

ExitHere:
        'Unload frmKitLevels
        frmKitLevels.Close()
        If OrderMode("A") Then
            If StoreSettings.bShowPackageItemPrices Then
                X = BillOSale.X
                BillOSale.SetStyle(X, "NOTES")
                BillOSale.SetDesc(X, "KIT SOLD - " & KitSyleNo & " - Total Price: " & FormatCurrency(RS("PackPrice").Value))
                BillOSale.SetPrice(X, 0)
                BillOSale.SetLoc(X, QuerySaleLocation()) 'StoresSld
                BillOSale.QuanEnabled = True
                BillOSale.PriceFocus()
                BillOSale.StyleAddEnd(False, X - BillOSale.NewStyleLine + 1) ' Hope this math is right...
            Else
                X = BillOSale.X
                BillOSale.SetDesc(X, IfNullThenNilString(RS("Heading").Value))
                BillOSale.SetPrice(X, IfNullThenZeroCurrency(RS("PackPrice").Value * KitQuantity))
                BillOSale.SetLoc(X, QuerySaleLocation()) 'StoresSld
                BillOSale.SetStyle(X, KitSyleNo)
                BillOSale.QuanEnabled = True
                BillOSale.PriceFocus()
                BillOSale.StyleAddEnd(False, X - BillOSale.NewStyleLine + 1) ' Hope this math is right...
            End If
            'Unload Me
            'FinishSelect
            Exit Sub
        Else
            'kit stock
            InvKitStock.fra.Text = " " & Style.Text & " "
            'Unload Me
            Me.Close()
        End If
    End Sub

    Private Sub GetRecord()
        ' Pull RN# "Rn" from 2Data.
        ' Show the resulting data in the form.

        Dim X As Integer
        X = BillOSale.X

        Dim DataObj As CInvRec
        DataObj = New CInvRec
        If Not DataObj.Load(CStr(RN), "#Rn") Then
            DisposeDA(DataObj)
            Exit Sub
        Else
            ' Show the data on the bos2 form.
            Dim OnOrder As String
            OnOrder = DataObj.QueryTotalOnOrder

            If OrderMode("A") Then
                BillOSale.SetStyle(X, StyleNo)
                BillOSale.SetMfg(X, DataObj.Vendor)
                BillOSale.SetMfgNo(X, DataObj.VendorNo)

                BillOSale.SetStatus(X, frmKitLevels.ItemStatusByStyle(StyleNo)) '   KitStatus '"ST"

                ' BFH20111107 - this is now part of getting kit status form...  starts at these but lets them change
                '        If IsUFO() Or IsFurnOne() Then ' Sets kit status
                '          BillOSale.SetStatus x, "LAW"
                '        ElseIf IsLapeer() Or IsPuritan() Or IsRockyMountain() Then
                '          BillOSale.SetStatus x, "SO"
                '        End If
                '
                ' 'tg' should be prepended only when status is 'LAW'
                ' the available check should be performed not 'vs -1' but 'vs -qty'
                If DataObj.Available - BillOSale.QueryQuan(X) <= 0 And IsIn(Trim(BillOSale.QueryStatus(X)), "ST", "DELTW") Then
                    BillOSale.SetDesc(X, "tg " & DataObj.Desc)
                Else
                    BillOSale.SetDesc(X, DataObj.Desc)
                    'BillOSale.Desc = DataObj.Desc  ' Set it to itself, how productive! :)
                End If



                BillOSale.SetLoc(X, QuerySaleLocation())  'StoresSld

                BillOSale.X = BillOSale.X + 1
                X = BillOSale.X  ' Added MJK20030714 to resynchronize kits.
            Else
                'kits Stock look up
                Dim IIo As Integer
                InvKitStock.UGridIO1.SetValueDisplay(Row, 0, IfNullThenNilString(Quan))
                InvKitStock.UGridIO1.SetValueDisplay(Row, 1, IfNullThenNilString(DataObj.Style))
                InvKitStock.UGridIO1.SetValueDisplay(Row, 2, IfNullThenNilString(DataObj.Vendor))
                InvKitStock.UGridIO1.SetValueDisplay(Row, 3, IfNullThenNilString(DataObj.Desc))
                InvKitStock.UGridIO1.SetValueDisplay(Row, 4, IfNullThenNilString(OnOrder))  ' Local variable, missing DataObj. intentional!
                For IIo = 1 To Setup_MaxStores
                    InvKitStock.UGridIO1.SetValueDisplay(Row, 4 + IIo, DataObj.QueryStock(IIo))
                Next
                Row = Row + 1
            End If

        End If

        DisposeDA(DataObj)
        Exit Sub

HandleErr:
        If Err.Number = 13 Then Resume Next
        MessageBox.Show("ERROR in GetRecord: " & Err.Description & ", " & Err.Source, "WinCDS")
    End Sub

    Private Sub mDBInvKit_GetRecordNotFound() Handles mDBInvKit.GetRecordNotFound
        'I believe that this lets you overwrite an entry
        'MsgBox (" Style Not Found! "), vbInformation
        FoundRecord = False
    End Sub

    Private Sub MouseMoveEvent(sender As Object, e As MouseEventArgs) Handles Style.MouseMove, cmdApply.MouseMove, cmdCancel.MouseMove, cmdDesc.MouseMove, cmdBarcode.MouseMove, chkStkOnly.MouseMove, fraSearch.MouseMove
        HidePreview
    End Sub

    Private Sub HidePreview()
        PreviewItemByStyle()
    End Sub

    Private Sub tmrItemPreview_Tick(sender As Object, e As EventArgs) Handles tmrItemPreview.Tick
        tmrItemPreview.Enabled = False
        lstStyles_MouseMove_PopUp
    End Sub

    Private Sub lstStyles_MouseMove_PopUp()
        On Error Resume Next
        If Not Visible Then Exit Sub
        PreviewItemByStyle(PopUpStyle, Me)
        '  lstStyles.ListIndex = yPos                                        ' Move to the list index
        '  lstStyles.Caption = lstStyles.List(yPos)                          ' Show listbox index value in label
    End Sub

    Private Sub lstStyles_MouseMove(sender As Object, e As MouseEventArgs) Handles lstStyles.MouseMove
        Dim yPos As Integer, tStyle As String

        ' lstStyles.TopIndex
        'FontName = lstStyles.FontName
        'FontSize = lstStyles.FontSize
        'Me.Font = New Font(Me.Font.Name, lstStyles.Font.Name)
        'Me.Font = New Font(Me.Font.Size, lstStyles.Font.Size)
        yPos = Int(e.Y / Printer.TextHeight("A")) + lstStyles.TopIndex
        '  yPos = (Y \ TextHeight("A")) + GetScrollPos(lstStyles.hwnd, &H1)   ' Get Item Position in the list (API call)
        If yPos < lstStyles.Items.Count Then
            tStyle = Trim(Microsoft.VisualBasic.Left(lstStyles.GetItemText(yPos), Setup_2Data_StyleMaxLen))
            If tStyle = "" Then PopUpStyle = "" : PreviewItemByStyle() : 
            Exit Sub
            If tStyle <> "" And tStyle = PopUpStyle Then Exit Sub
            If GetRNByStyle(tStyle) = 0 Then Exit Sub
            PopUpStyle = tStyle

            tmrItemPreview.Enabled = False
            'tmrItemPreview.Interval = 100
            'tmrItemPreview.Enabled = True  NOTE: REMOVE THESE TWO LINE COMMENTS AFTER COMPLETE CODING.
        End If
    End Sub

    Private Sub cmdApply_Click(sender As Object, e As EventArgs) Handles cmdApply.Click
        If Style.Visible = False Then
            If lstStyles.Items.Count > 0 Then
                Style.Text = Trim(Microsoft.VisualBasic.Left(lstStyles.GetItemText(lstStyles.SelectedIndex), 16)) '###STYLELENGTH16
            End If
        End If

        If Trim(Style.Text) = "" Then
            'Below line is commented. It is custom msgbox. find full details later.
            'MsgBox("Please enter a Style Number.", vbExclamation, "No Style Number")
            MessageBox.Show("Please enter a Style Number.", "No Style Number", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If Microsoft.VisualBasic.Left(Style.Text, 4) <> KIT_PFX And ReportsMode("CS") And Not Style.Visible = False Then
            cmdCancel.Enabled = True
            cmdDesc.Enabled = True
            'MsgBox("You may only select kits from this list.", vbExclamation, "Warning")
            MessageBox.Show("You may only select kits from this list.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        ElseIf (microsoft.VisualBasic.Left(Style.Text, 4) <> KIT_PFX And ReportsMode("CS")) Or optKitVendors.Checked = True Then 'Added by Robert 5/16/2017
            DoSelect()
            Exit Sub
        End If
        DoApply()

    End Sub

    Private Sub chkStkOnly_Click(sender As Object, e As EventArgs) Handles chkStkOnly.Click
        If Width = FRMW_2 Or Width = FRMW_1b Then Width = IIf(chkStkOnly.Checked = True, FRMW_2, FRMW_1b)
        DoCaptions()

        If optSearchByStyle.Checked = True And IsByVendor <> "" Then
            VendorItemsSearch(IsByVendor)
        Else
            'Style_Change
            'AddHandler Style.TextChanged, AddressOf Style_TextChanged
            Style_TextChanged(Style, New EventArgs)
        End If
    End Sub

    Private Sub optSearchByDesc_Click(sender As Object, e As EventArgs) Handles optSearchByDesc.Click
        UpdateSearchBox()
    End Sub

    Private Sub optSearchByStyle_Click(sender As Object, e As EventArgs) Handles optSearchByStyle.Click
        fraSearch.Visible = True
        If Not DontUpdate Then UpdateSearchBox()
    End Sub

    Private Sub optSearchByVendor_Click(sender As Object, e As EventArgs) Handles optSearchByVendor.Click
        DoCaptions()
        UpdateSearchBox()
    End Sub

    Private Sub optKitVendors_Click(sender As Object, e As EventArgs) Handles optKitVendors.Click
        Dim KitVendor As String

        fraSearch.Visible = False ' Let's take out the frame

        KitVendor = Style.Text
        'KitSyleNo = Left(Style, 4)
        'mDBAccess_Init Left(KitSyleNo, 4)

        mDBInvKit_Init()
        'SELECT DISTINCT kitSKU
        '         From InvKit
        '       ORDER BY KITSKU

        mDBAccess = New CDbAccessGeneral
        mDBAccess.dbOpen(GetDatabaseAtLocation(1))  ' Kits are only at location 1.
        mDBAccess.SQL =
      "SELECT DISTINCT KitSKU" _
      & " From InvKit" _
      & " WHERE """ & ProtectSQL(Trim(KitVendor)) & """ = """" or """ & ProtectSQL(Trim(KitVendor)) & "*"" like KitSKU ORDER BY KitSKU"
        Dim RS As ADODB.Recordset : RS = mDBAccess.getRecordset()

        If RS.RecordCount <> 1 Then ' did it find any records
        End If
        lstStyles.items.Clear
        Do While Not RS.EOF  ' iterate through all records
            On Error Resume Next
            'lstStyles.AddItem CStr(RS("KitSKU"))
            lstStyles.Items.Add(CStr(RS("KitSKU").Value))
            RS.MoveNext()
        Loop
        mDBAccess.dbClose()
        mDBAccess = Nothing

        'Adjust look and feel of the page

        DoCaptions()
        'UpdateSearchBox
    End Sub

    Private Sub Style_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Style.KeyPress
        'If KeyAscii = vbKeyReturn Then
        '    cmdApply.Value = True
        '    Exit Sub
        'End If
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))

        If e.KeyChar = Convert.ToChar(13) Then
            cmdApply.PerformClick()
            Exit Sub
        End If
        e.KeyChar = UCase(e.KeyChar)
    End Sub

    Private Sub tmrType_Tick(sender As Object, e As EventArgs) Handles tmrType.Tick
        tmrType.Enabled = False
        HandleStyleChange()
    End Sub

    Private Sub Combo1_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Sub

    Private Sub InvCkStyle_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' This event is replacement for form unload and query unload event of vb6.0
        'If UnloadMode = vbFormControlMenu Then Cancel = True 'cmdCancel.value = True  ' Hangs program when we do this the right way.
        If e.CloseReason = CloseReason.FormOwnerClosing Then
            e.Cancel = True
        End If
        PreviewItemByStyle("") ' Remove the preview on form close
    End Sub
End Class
Imports VBA
Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class InvenA
    Public Balance(40) As Integer
    Public OnOrder(40) As Integer
    Public RN As Integer
    Public InitialStyle As String  ' Set in GetRec, used in cmdOk_Click
    Public Freight As Double, GM As Double, Mu As Double  ' Expanded values for the percentage boxes.
    Dim LastVendor As String, LastStyle As String, LastVendorNo As String
    '###STORECOUNT32
    Private Const StoreCount As Integer = 32 ' how many text boxes are represented on this screen, regardless of the max # of stores
    Dim ItemList As ADODB.Recordset

    Public Function CheckBal() As Boolean
        ' Called by InvenA.Form_Load, InvCkStyle.FinishSelect
        '    StyleCkIt2 = ""
        On Error GoTo HandleErr

        ' Search for product by Style.
        ' If anything is found,
        '   If invenmode("A") give error message.
        '   Otherwise, set Rn and call GetRec.
        ' Else what?

        ' First off, we encapsulate loading by Style.
        Dim SearchObj As New CSearchNew
        If SearchObj.Load(Style.Text) Then
            ' Record was found.
            If InvenMode("A") Then
                MsgBox("This Style Number is already in use.  Add a -1, -2, -3 ect. to the end of the style number.", vbExclamation)
                Me.Style.Text = ""
                'Unload Me
                Exit Function
            Else
                RN = Val(SearchObj.RN)
                GetRec(CStr(RN))
            End If
        Else
            If InvenMode("A") Then
                ' The style isn't in Search, find out if it's in 2Data.
                Dim InvRec As CInvRec
                InvRec = New CInvRec
                If InvRec.Load(Style.Text, "Style") Then
                    ' Style existed at one point.
                    Dim C As VbMsgBoxResult
                    C = MsgBox("This style number has been deactivated." & vbCrLf & "You can restore it using the File menu." & vbCrLf2 & "Would you like to restore this item now?", vbExclamation + vbYesNo + vbDefaultButton2, "Item Exists, but is Deleted")
                    If C = vbYes Then DoInvRestore(Style.Text)
                    DisposeDA(InvRec)
                    '            Unload Me
                    Exit Function
                Else
                    If ValidBarcode(Style.Text) Then
                        ' Style has never existed, it's ok to create.
                    Else
                        If Style.Text = "" Then
                            MsgBox("Blank styles are not accepted.", vbExclamation)
                        Else
                            MsgBox("Invalid style Number: " & Style.Text & vbCrLf & "You may only use these characters: A-Z, 0-9, Space, $ % + - . /", vbExclamation, "Invalid Style")
                        End If
                        Style.Text = ""
                        'Unload Me
                        Me.Close()
                        Exit Function
                    End If
                End If
                DisposeDA(InvRec)
            End If
        End If

        ' What does this comment mean?
        'B / D / E / H / I / J

        If Trim(SearchObj.Style) = Trim(InvCkStyle.Style.Text) Then
            If InvenMode("D") Then Me.Text = "Processing Factory Shipments"
            If InvenMode("E") Then Me.Text = "View Any Item"
            'Unload InvCkStyle
            InvCkStyle.Close()
        End If

        If InvenMode("D") Then
            'Load InvStkRec
            'InvStkRec.HelpContextID = HelpContextID
            InvStkRec.Show()
        End If
        CheckBal = True
        Exit Function

HandleErr:
        Resume Next
    End Function

    Public Sub GetRec(ByVal SearchRn As String)
        Dim InvData As CInvRec
        Dim I As Integer, BadPic As Boolean

        If Trim(SearchRn) <> "" Then RN = SearchRn
        If Trim(RN) = "" Then
            MsgBox("Invalid Search Rn in InvenA.GetRec.", vbCritical, "Error")
            Exit Sub
        End If

        On Error GoTo HandleErr

        ClearData(False)

        ' Get InvData (by Rn) from class module.
        ' Maybe: InvData.Load(Style)
        InvData = New CInvRec
        If InvData.Load(CStr(RN), "#Rn") Then
            '    StyleCkIt2 = InvData.Style

            InitialStyle = InvData.Style  ' Save the original Style, in case it's updated.

            Style.Text = InvData.Style
            LastStyle = Style.Text
            LoadMfgName()
            Mfg.Text = InvData.Vendor
            LastVendor = Mfg.Text
            LastVendorNo = InvData.VendorNo
            RDate.Text = DateFormat(InvData.RDate)
            If IsDate(InvData.RDate) Then RDate2.Value = InvData.RDate

            DeptNo.Text = InvData.DeptNo
            VendorNo.Text = InvData.VendorNo
            SelectManufacturer(InvData.VendorNo, InvData.Vendor)

            LoadDeptNames()
            SelectDeptName(InvData.DeptNo)
            cboDeptName.Visible = True
            DeptNo.Visible = False

            If InvData.FreightType = 0 Then
                optPercent.Checked = True
            Else  ' dollar amount
                optAmtFt.Checked = True
            End If

            Desc.Text = InvData.Desc

            Ms.Text = InvData.MinStk
            Freight = InvData.Freight
            GM = InvData.GM
            Mu = InvData.MarkUp
            txtFreight.Text = Format(Freight, "0.00")

            txtGM.Text = Format(GM, "0.00")
            txtMU.Text = Format(Mu, "0.00")

            Cost.Text = CurrencyFormat(InvData.Cost)
            Spiff.Text = CurrencyFormat(InvData.Spiff)
            Cubes.Text = CurrencyFormat(InvData.Cubes)
            Landed.Text = CurrencyFormat(InvData.Landed)

            OnSale.Text = CurrencyFormat(InvData.OnSale)
            List.Text = CurrencyFormat(InvData.List)

            Comments.Text = InvData.Comments
            SKU.Text = InvData.SKU

            ' Allocation can become inaccurate, but setting Ba# fixes it.
            Ns.Text = InvData.Available
            Allocation.Text = InvData.OnHand

            For I = 1 To StoreCount
                Balance(I) = InvData.QueryStock(I)
                OnOrder(I) = InvData.QueryOnOrder(I)
            Next

            Write1.Text = ZeroToEmptyString(InvData.Sales1)
            Write2.Text = ZeroToEmptyString(InvData.Sales2)
            Write3.Text = ZeroToEmptyString(InvData.Sales3)
            Write4.Text = ZeroToEmptyString(InvData.Sales4)

            Pwrite1.Text = ZeroToEmptyString(InvData.Psales1)
            Pwrite2.Text = ZeroToEmptyString(InvData.Psales2)
            Pwrite3.Text = ZeroToEmptyString(InvData.Psales3)
            Pwrite4.Text = ZeroToEmptyString(InvData.Psales4)

            PoSold.Text = InvData.PoSold
            Dim GG As Double, TN As Double
            Dim G As Double, AI As Double, GD As Decimal, ID As Decimal, Sls As Decimal, Cgs As Decimal, Env As Double
            GG = CalculateGMROI(InvData.Style, YearStart, Today, False, AI, GD, ID, Sls, Cgs, Env)
            If Env = 0 Then TN = 0 Else TN = Cgs / Env
            GMROI.Text = CurrencyFormat(GG)
            If Env <> 0 Then
                txtTurns.Text = Format(Cgs / Env, "0.00")
            Else
                txtTurns.Text = Format(0, "0.00")
            End If


            lblImageFileName.Visible = False
            txtImageFileName.Visible = False

            If ItemList Is Nothing Then
                ItemList = GetRecordsetBySQL("SELECT Style, VendorNo, Rn FROM [2Data] WHERE [2Data].Rn IN (SELECT Rn FROM Search) ORDER BY VendorNo, Style", , GetDatabaseInventory)
            End If
            On Error Resume Next
            ItemList.MoveFirst()
            ItemList.Find("Style='" & Style.Text & "'", , ADODB.SearchDirectionEnum.adSearchForward)

            BadPic = False
            On Error GoTo BadPicture
            ResetPicture()
            'Code Edited by Robert - 5/11/2017 Many images are in a newer format that VB6 could not handle
            'Also added a library Freeimage.dll and mFreeImage.bas
            'BFH20170515 - Moved all FreeImage code to handler function:  ItemPictureByRN  Please refer to handler for comments regarding change.
            picItem.Image = ItemPictureByRN(RN)
            'picItem.ToolTipText = ItemPXByRN(RN)
            ToolTip1.SetToolTip(picItem, ItemPXByRN(RN))
            If Not BadPic Then
                MaintainPictureRatio(picItem)
                picItem.Visible = True
            End If
        End If
        DisposeDA(InvData)

        cmdEDI.Visible = Mfg.Text = "ASHLEY" ' BFH20101114
        Exit Sub

HandleErr:
        Resume Next

BadPicture:
        lblImageFileName.Visible = True
        txtImageFileName.Visible = True
        lblImageFileName.Text = "File Name For Item"
        ResetPicture()
        txtImageFileName.Text = ItemPXByRN(RN, False)
        BadPic = True
        Resume Next
    End Sub

    Public Sub DoInvRestore(Optional ByVal Style As String = "")
        MainMenu.Hide()
        Show()
        Text = "Restore Deleted Items"
        'RestoreInv.HelpContextID = 36000
        If Style <> "" Then RestoreInv.txtStyle.Text = Style
        'RestoreInv.Show vbModal, Me ' Can't show modal (or parented at all).  Microsoft bug causes VB to hang when a child form closes its parent.
        RestoreInv.ShowDialog(Me)
        'Unload Me
        Me.Close()
        MainMenu.Show()
    End Sub

    Public Sub ClearData(ByVal RetainMfg As Boolean)
        ' Called from GetRec before filling the screen with new data.
        ' Also called by InvANext, which comes after a new item is saved.
        ' Clear old lookup values on the screen!
        ' Otherwise if any fields in the database are null, old values will bleed through.
        Dim ControlsToClear As Object, El As Control

        '    ControlsToClear = Array(Style, Desc, Comments, SKU, txtImageFileName,
        'Write1, Write2, Write3, Write4, Pwrite1, Pwrite2, Pwrite3, Pwrite4, PoSold)

        ControlsToClear = New Control() {Style, Desc, Comments, SKU, txtImageFileName,
    Write1, Write2, Write3, Write4, Pwrite1, Pwrite2, Pwrite3, Pwrite4, PoSold}
        For Each El In ControlsToClear
            El.Text = ""
        Next

        ControlsToClear = New Control() {Ba1, Ba2, Ba3, Ba4, Ba5, Ba6, Ba7, Ba8, Ba9, Ba10, Ba11, Ba12, Ba13, Ba14, Ba15, Ba16}
        'For Each El In Ba
        For Each El In ControlsToClear
            El.Text = ""
        Next

        ControlsToClear = New Control() {OO1, OO2, OO3, OO4, OO5, OO6, OO7, OO8, OO9, OO10, OO11, OO12, OO13, OO14, OO15, OO16}
        'For Each El In OO
        For Each El In ControlsToClear
            El.Text = ""
        Next

        'ControlsToClear = Array(Ms, Cost, Landed, OnSale, List, Ns, Spiff, Cubes)
        ControlsToClear = New Control() {Ms, Cost, Landed, OnSale, List, Ns, Spiff, Cubes}
        For Each El In ControlsToClear
            El.Text = "0"
        Next

        'ControlsToClear = Array(Allocation)
        ControlsToClear = New Control() {Allocation}
        For Each El In ControlsToClear
            El.Text = "0"
        Next

        If Not RetainMfg Then
            'ControlsToClear = Array(Mfg, VendorNo, txtInvoiceNo)
            ControlsToClear = New Control() {Mfg, VendorNo, txtInvoiceNo}
            For Each El In ControlsToClear
                El.Text = ""
            Next

            'ControlsToClear = Array(txtFreight, txtGM, txtMU)
            ControlsToClear = New Control() {txtFreight, txtGM, txtMU}
            For Each El In ControlsToClear
                El.Text = "0"
            Next
            GM = 0 : Freight = 0 : Mu = 0
            cboMfgName.SelectedIndex = -1

            DeptNo.Text = ""
            cboDeptName.SelectedIndex = 0
        End If

        RDate.Text = DateFormat(Today)
        RDate2.Value = Today

        cboDeptName.Select()
    End Sub

    Private Sub LoadMfgName()
        ' This will do until we make a Vendor class (and maybe a vendor table).
        Dim OldVen As Integer
        OldVen = Val(VendorNo.Text)
        LoadMfgNamesIntoComboBox(cboMfgName)
        SelectManufacturer(OldVen)
    End Sub

    Public Sub SelectManufacturer(ByVal MfgNo As Integer, Optional ByVal MfgName As String = "")
        Dim I As Integer, BestFit As Integer
        On Error Resume Next
        BestFit = -1
        For I = 0 To cboMfgName.Items.Count
            'If cboMfgName.itemData(I) = MfgNo Then
            If CType(cboMfgName.SelectedItem, ItemDataClass).ItemData = MfgNo Then
                If cboMfgName.Text = MfgName Or MfgName = "" Then
                    cboMfgName.SelectedIndex = I
                    Exit Sub
                ElseIf BestFit = 0 Then
                    BestFit = I
                End If
            End If
        Next
        If BestFit >= 0 Then
            cboMfgName.SelectedIndex = BestFit
            Exit Sub
        End If
        cboMfgName.Text = "Select Manufacturer"
    End Sub

    Public Sub LoadDeptNames() ' Loads cboDeptName combo box.
        LoadDeptNamesIntoComboBox(cboDeptName, StoresSld)
    End Sub

    Public Sub SelectDeptName(ByVal selDept As String)
        Dim I As Integer
        On Error Resume Next
        For I = 0 To cboDeptName.Items.Count
            'If cboDeptName.itemData(I) = selDept Then
            If CType(cboDeptName.SelectedItem, ItemDataClass).ItemData = selDept Then
                'cboDeptName.ListIndex = I
                cboDeptName.SelectedIndex = I
                Exit Sub
            End If
        Next
    End Sub

    '    Public Property Get Balance(ByVal StoreNum as integer) As Single
    '  Balance = CSng(GetDouble(TxtBalance(StoreNum)))
    'End Property
    '    Public Property Let Balance(ByVal StoreNum as integer, ByVal nValue As Single)
    '  TxtBalance(StoreNum).Text = ZeroToEmptyString(nValue)
    'End Property

    'Public Property Balance() As Single
    '    Get
    '        Balance = CSng(GetDouble(TxtBalance(StoreNum)))
    '    End Get

    '    'Set(Optional ByVal StoreNum as integer = 1, Optional ByVal value As Single = 1)
    '    '    TxtBalance(StoreNum).Text = ZeroToEmptyString(nValue)
    '    'End Set
    '    Set(a As Single, b As Single)

    '    End Set
    'End Property

    Private Sub ResetPicture()
        picItem.Visible = False
        picItem.Image = LoadPictureStd("")
        picItem.Width = 5535
        fraStLocScroll.Width = IIf(LicensedNoOfStores > 8, 11412, 5772)
        picItem.Height = IIf(LicensedNoOfStores > 8, 3255, 4695)
    End Sub

End Class
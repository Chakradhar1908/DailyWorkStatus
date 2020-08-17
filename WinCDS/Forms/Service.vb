Imports Microsoft.VisualBasic.Compatibility.VB6
Public Class Service
    Public AccountFound As String
    Public MailIndex as integer
    Private ServiceOrderNumber as integer
    Private WithEvents mDBAccess As CDbAccessGeneral
    Private WithEvents mDBService As CDbAccessGeneral
    Private LoadingCheckBoxes As Boolean, SearchingSOID As Boolean
    Private Mail2 As MailNew2

    Public Sub LoadCustomer(ByVal NewMailIndex as integer, Optional ByVal CheckServiceCalls As Boolean = True)
        ' Load the customer info.
        ' How do we know when to bring up AddOnAcc form?

        ' What happens in this form:
        '  1: User chooses Service Call from main menu.
        '    Service form brings up mail info.
        '    If customer has prior calls, bring up AddOnAcc and act on its output.
        '  2: User chooses Next from Service form.
        '    Unload Service form.
        '    Service form brings up mail info.
        '    If customer has prior calls, bring up AddOnAcc and act on its output.
        '  3: User uses arrow buttons to navigate.
        '    Load target service call and customer info, no prompt.

        ' Functions:
        '   Load Customer Info
        '   Load Service Call Info
        '   Prompt for new/old Service Call.

        ' Sometimes we know Service Call first, sometimes Customer.

        If NewMailIndex > 0 Then
            ' Look up the customer's info.
            Dim mR As MailNew
            Mail_GetAtIndex(CStr(NewMailIndex), mR)
            If mR.Index > 0 Then
                LoadMailRecord(mR)
                FindItems()
            Else
                ' Bad mail record!
                MsgBox("Invalid mail index in Service module.", vbCritical, "Error")
                Exit Sub
            End If
        Else
            ' Bad mail info, what to do?
            MsgBox("No customer record available.", vbCritical, "Error")
            Exit Sub
        End If

        Dim NewCallNo as integer
        If CheckServiceCalls = True Then
            If MailCheck.ServiceCallNo > 0 Then
                ' Load this call..
                NewCallNo = MailCheck.ServiceCallNo
            Else
                ' Find a call to work with..
                AccountFound = ""
                CheckForService(CLng(MailIndex))   ' This should look but not load!
                If AccountFound = "Y" Then
                    If Val(NewMailIndex) > 0 Then
                        ' Show the form to select old or new service call.
                        AddOnAcc.Typee = ArAddOn_Nil
                        'AddOnAcc.Show vbModal, Me
                        AddOnAcc.ShowDialog(Me)
                        If AddOnAcc.Typee = ArAddOn_Add Then      ' Add to old service call.
                            NewCallNo = Val(AddOnAcc.ServiceNo)
                        Else
                            lblServiceOrderNo.Text = ""
                            ServiceOrderNumber = 0
                            FindItems()
                        End If
                        'Unload AddOnAcc
                        AddOnAcc.Close()
                        AddOnAcc = Nothing
                    End If
                End If
            End If
            If NewCallNo > 0 Then
                LoadServiceCall(NewCallNo)
            Else
                ' New Service Call.
                ' Clear old service call info..?
                ClearServiceOrder()
            End If
        End If
    End Sub
    Public Sub LoadServiceCall(ByVal SOID as integer, Optional ByVal Direction As String = "")
        ' Load the service call..
        ' If the Service Call's MailIndex doesn't match our current MailIndex,
        ' we need to also load the customer - no prompt.

        Me.AccountFound = "N"
        If Direction = "" Then
            mDBAccess_Init(SOID)
        Else
            mDBAccess_Init(, Direction)        ' this creates the sql string
        End If
        mDBAccess.GetRecord()
        On Error Resume Next
        mDBAccess.dbClose()
        mDBAccess = Nothing

        LoadPartsOrders()
    End Sub
    Public Sub LoadPartsOrders()
        Dim RS As ADODB.Recordset
        Dim SQL As String, Tot as integer, Closed as integer, N as integer

        lblPartsOrd.Visible = False
        On Error Resume Next
        If ServiceOrderNumber <= 0 Then Exit Sub

        SQL = "SELECT Count(ServicePartsOrderNo) As NOrders FROM ServicePartsOrder WHERE ServiceOrderNo=" & ServiceOrderNumber
        RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(), True, "Failed to access table [ServicePartsOrder] for Serivce.LoadPartsOrder()")
        Tot = RS("NOrders").Value
        RS = Nothing
        SQL = "SELECT Count(ServicePartsOrderNo) As NClosedOrders FROM ServicePartsOrder WHERE ServiceOrderNo=" & ServiceOrderNumber & " AND Trim(Status)='Closed'"
        RS = GetRecordsetBySQL(SQL, False, GetDatabaseAtLocation(), True, "Failed to access table [ServicePartsOrder] for Serivce.LoadPartsOrder()")
        Closed = RS("NClosedOrders").Value
        RS = Nothing

        If Tot = 0 Then
            lblPartsOrd.Visible = False
        Else
            lblPartsOrd.Visible = True
            lblPartsOrd.BackColor = Color.Yellow
            If Tot <= Closed Then
                lblPartsOrd.Text = "" & Tot & " Part" & IIf(Tot = 1, "", "s") & " Ordered"
                lblPartsOrd.BackColor = Color.Green
            ElseIf Closed = 0 Then
                lblPartsOrd.Text = "" & Tot & " Part" & IIf(Tot = 1, "", "s") & " On Order"
            Else
                N = Tot - Closed
                lblPartsOrd.Text = "" & (N) & " Part" & IIf(N = 1, "", "s") & " On Order" & vbCrLf &
                            "" & Closed & " Resolved"
            End If
        End If
    End Sub

    Private Sub mDBAccess_Init(Optional SOID as integer = 0, Optional f_strDirection As String = "", Optional MailIndex as integer = 0)
        ' Called by CheckForService, cmdSave_Click, GetServiceCall, MoveRecord.
        mDBService_Init()  ' Is this the best place for it?
        mDBAccess = New CDbAccessGeneral
        With mDBAccess
            .dbOpen(GetDatabaseAtLocation())
            If SOID <> 0 Then
                .SQL = "SELECT * From Service WHERE ServiceOrderNo  =" & SOID
            Else
                If MailIndex <> 0 Then
                    .SQL = "SELECT * From Service WHERE MailIndex  = " & MailIndex & ""
                Else
                    Select Case f_strDirection
                        Case "First"
                            .SQL = "SELECT TOP 1 * FROM Service ORDER BY ServiceOrderNo"
                        Case "Last"
                            .SQL = "SELECT TOP 1 * FROM Service ORDER BY ServiceOrderNo DESC"
                        Case "Previous"
                            .SQL = "SELECT TOP 1 * FROM Service WHERE ServiceOrderNo<" & ServiceOrderNumber & " ORDER BY ServiceOrderNo DESC"
                        Case "Next"
                            .SQL = "SELECT TOP 1 * FROM Service WHERE ServiceOrderNo>" & ServiceOrderNumber & " ORDER BY ServiceOrderNo"
                    End Select
                End If
            End If
        End With
        Exit Sub

HandleErr:
        MsgBox("ERROR mdbAccess_Init: " & Err.Description & ", " & Err.Source)
        Resume Next
    End Sub
    Private Sub mDBService_Init()
        mDBService = New CDbAccessGeneral
        Dim a As String
        a = mDBService.dbOpen(GetDatabaseAtLocation())
    End Sub

    Private Sub ClearServiceOrder()
        ServiceOrderNumber = 0
        lblServiceOrderNo.Text = ""

        lblSaleNo.Text = ""
        lblSaleNo.Visible = False
        lblSaleNoCaption.Visible = False

        dteServiceDate.Value = Today
        dteServiceDate.Value = Nothing
        lblClaimDate.Text = Today

        SelectStatus("")

        LoadCheckBoxes(1)

        txtItems.Text = ""
        Notes_Text.Text = ""
        Notes_New.Text = ""
    End Sub
    Private Sub SelectStatus(ByVal Stat As String)
        Select Case UCase(Trim(Stat))
            Case "", "OPEN"  ' allow "" for clearing
                cboStatus.SelectedIndex = 0
            Case "CLOSED"
                cboStatus.SelectedIndex = 1
            Case Else
                If cboStatus.DropDownStyle = 2 Then ' drop down list
                    'cboStatus.AddItem Stat
                    cboStatus.Items.Add(Stat)
                    'cboStatus.SelectedIndex = cboStatus.NewIndex
                    cboStatus.SelectedIndex = cboStatus.Items.Count - 1
                Else
                    cboStatus.Text = Stat
                End If
        End Select
    End Sub
    Private Sub LoadCheckBoxes(ByVal Val as integer, Optional ByVal ClearOnly As Boolean = False)
        LoadingCheckBoxes = True
        If Not ClearOnly Or Val <> 1 Then chkStoreService.Checked = IIf(Val = 1, 1, 0)
        If Not ClearOnly Or Val <> 2 Then chkOutsideService.Checked = IIf(Val = 2, 1, 0)
        If Not ClearOnly Or Val <> 3 Then chkPickupExchange.Checked = IIf(Val = 3, 1, 0)
        If Not ClearOnly Or Val <> 4 Then chkOther.Checked = IIf(Val = 4, 1, 0)
        LoadingCheckBoxes = False
    End Sub

    Private Sub CheckForService(ByVal MailIndex as integer)
        ' Called by LoadCustomer
        On Error GoTo HandleErr

        mDBAccess_Init(, , MailIndex)  'MailCheck.CustomerTele)
        mDBAccess.GetRecord()   ' this gets the record
        mDBAccess.dbClose()
        mDBAccess = Nothing

        'Set mDBService.db = mDBAccess.db
        mDBService_Init()
        mDBService_SqlSet(CStr(MailIndex))
        mDBService.GetRecord()   ' this gets the record

        mDBService.dbClose()
        mDBService = Nothing
        Exit Sub

HandleErr:
        MsgBox("Check for Service: " & Err.Description & ", " & Err.Source)
        Resume Next
    End Sub
    Private Sub mDBService_SqlSet(ByVal T As String)
        ' Only called from CheckForService
        If AddOnAcc.Typee = ArAddOn_New And Me.AccountFound = "Y" Then
            mDBService.SQL = "SELECT Service.* From Service WHERE ServiceOrderNo  = " & T
        Else
            mDBService.SQL = "SELECT Service.* From Service WHERE MailIndex  = " & T
        End If
    End Sub

    Public Sub FindItems()
        Dim Margin As CGrossMargin, Zz as integer, ItemsUpdated As Boolean, S As String
        Dim NN As Object, Selected As Boolean
        Dim ItemDescString As String, AckInv As String
        Dim RS As ADODB.Recordset
        Dim X As ADODB.Recordset, A as integer

        Margin = New CGrossMargin

        'Font.Name = "Arial"
        lstPurchases.Items.Clear()
        tvItemNotes.Visible = False
        tvItemNotes.Nodes.Clear()
        tvItemNotes.Nodes.Add("", "", "LABEL",
        ArrangeString("VENDOR", 17) & ArrangeString("STYLE", 17) & ArrangeString("SALE NO", 10) & ArrangeString("QUAN", 6) _
        & ArrangeString("DEL DATE", 12) & ArrangeString("DESCRIPTION", 32) & "ACK/INV NO")

        'tvItemNotes.Nodes("LABEL").Bold = True
        tvItemNotes.Nodes("LABEL").NodeFont = New Font(tvItemNotes.Font, FontStyle.Bold)

        A = Val(lblServiceOrderNo)
        If A = 0 And IsFormLoaded("MailCheck") Then
            A = Val(MailCheck.ServiceCallNo)
        End If
        If A <> 0 Then
            X = GetRecordsetBySQL("SELECT * FROM ServiceItemParts WHERE ServiceOrderNo=" & A & " AND MarginNo=0")

            Do While Not X.EOF
                ItemDescString =
          ArrangeString(UCase(IfNullThenNilString(X("Vendor"))), 17) & ArrangeString(UCase(IfNullThenNilString(X("Style"))), 17) & ArrangeString(UCase(IfNullThenNilString(X("SaleNo"))), 10) &
          ArrangeString(IfNullThenZeroDouble(X("Quantity")), 6) & ArrangeString(IfNullThenZeroDate(X("DelDate")), 12) &
          ArrangeString(UCase(IfNullThenNilString(X("Desc"))), 32) & ArrangeString("", 15)

                'NN = tvItemNotes.Nodes.Add(, , "EX-" & X("STYLE").Value & "-" & Random(1000), ItemDescString)
                NN = tvItemNotes.Nodes.Add("", "", "EX-" & X("STYLE").Value & "-" & Random(1000), ItemDescString)
                NN.ForeColor = Color.Red

                X.MoveNext()
            Loop
            X = Nothing
        End If

        Zz = 0
        Dim SQL As String
        SQL = ""
        SQL = SQL & "SELECT * FROM GrossMargin WHERE Trim(MailIndex)="""
        SQL = SQL & ProtectSQL(Trim(MailIndex)) & """ AND Trim(Style) NOT IN ('STAIN', 'DEL', 'LAB', "
        SQL = SQL & "'TAX1', 'TAX2', 'NOTES', 'SUB', 'PAYMENT', '--- Adj ---') "
        SQL = SQL & "AND Status NOT LIKE ""x%"" AND Status<>'VOID' AND Status NOT LIKE ""VD%"" "
        SQL = SQL & "ORDER BY SaleNo, MarginLine"

        Margin.DataAccess.Records_OpenSQL(SQL)
        Do While Margin.DataAccess.Records_Available

            'added detail 03/23/2003
            RS = GetRecordsetBySQL("SELECT * FROM Detail WHERE MarginRn=" & Margin.MarginLine & " AND Store=" & StoresSld, , GetDatabaseInventory)
            If Not RS.EOF Then
                AckInv = Trim(IfNullThenNilString(RS("Misc")))
            Else
                AckInv = ""
            End If
            DisposeDA(RS)
            ItemDescString = ""
            ItemDescString = ItemDescString & ArrangeString(Margin.Vendor, 17) & ArrangeString(Margin.Style, 17) & ArrangeString(Margin.SaleNo, 10)
            ItemDescString = ItemDescString & ArrangeString(Margin.Quantity, 6) & ArrangeString(DateFormat(Margin.DDelDat), 12)
            ItemDescString = ItemDescString & ArrangeString(Margin.Desc, 32) & ArrangeString(AckInv, 15)

            'lstPurchases.AddItem ItemDescString
            'lstPurchases.itemData(lstPurchases.NewIndex) = Margin.Detail

            '--> Note: replaced above two lines with the below one. created custom class ItemDataclass to implement
            '--> itemData property of vb6 in vb.net
            lstPurchases.Items.Add(New ItemDataClass(ItemDescString, Margin.Detail))

            'NN = tvItemNotes.Nodes.Add(, , "ML" & Margin.MarginLine, ItemDescString)
            NN = tvItemNotes.Nodes.Add("", "", "ML" & Margin.MarginLine, ItemDescString)
            ColorTaggedItem(Margin.MarginLine, IsItemTaggedForRepair(Margin.MarginLine))

            On Error Resume Next
            If Not Selected And IsItemTaggedForRepair(Margin.MarginLine) Then
                'tvItemNotes.
                'tvItemNotes.SelectedItem = NN
                tvItemNotes.SelectedNode = NN
                Selected = True
            End If
            On Error GoTo 0

            If ServiceOrderNumber > 0 Then
                If InStr(txtItems.Text, Microsoft.VisualBasic.Left(ItemDescString, 50)) > 0 Then
                    TagItemForRepair(Margin.MarginLine)
                    ItemsUpdated = True
                    ' Have to update txtItems in the database at this time too, but wait for the last item.
                End If
            End If

            Dim ServiceNote As clsServiceNotes
            ServiceNote = New clsServiceNotes
            With ServiceNote
                .DataAccess.Records_OpenSQL("SELECT * FROM ServiceNotes WHERE MarginNo=" & Margin.MarginLine & " AND ServiceCall=" & ServiceOrderNumber & " ORDER BY NoteDate, ServiceNoteID")
                If Not .DataAccess.Record_EOF Then
                    Do While .DataAccess.Records_Available
                        Dim Note As Object, I as integer
                        Note = SplitLongText(" --- " & .NoteTypeString & " entered at " & DateFormat(.NoteDate) & " ---" & vbCrLf & .Note, 75)
                        For I = LBound(Note) To UBound(Note)
                            tvItemNotes.Nodes.Add(IIf(I > LBound(Note), "SN" & .ServiceNoteID, "ML" & Margin.MarginLine), "4", "SN" & .ServiceNoteID & IIf(I > LBound(Note), "." & I, ""), Note(I))
                            tvItemNotes.Nodes(IIf(I > LBound(Note), "SN" & .ServiceNoteID, "ML" & Margin.MarginLine)).Expanded = True
                        Next
                    Loop
                    'tvItemNotes.Nodes("ML" & Margin.MarginLine).Expanded = True
                    tvItemNotes.Nodes("ML" & Margin.MarginLine).Expand()
                End If
            End With
            DisposeDA(ServiceNote)
        Loop
        'tvItemNotes.Style = tvwTreelinesPlusMinusPictureText
        DisposeDA(Margin)
        'lstPurchases.Clear
        lstPurchases.Items.Clear()


        If ItemsUpdated Then
            Dim cSR As clsServiceOrder
            cSR = New clsServiceOrder
            If cSR.Load(CStr(ServiceOrderNumber), "#ServiceOrderNo") Then
                cSR.Item = "" '  txtItems.Text  ' Just clear it.
                cSR.Save
            Else
                ' How can it not load, we're in it?
                MsgBox("Error upgrading service record structure.", vbCritical, "Error")
            End If
            DisposeDA(cSR, Margin)
        End If

        tvItemNotes.Visible = True

        LoadPartsOrders()
        'DisposeDA(cSR, Margin)

        Exit Sub

        LoadPartsOrders()
HandleErr:
        MsgBox("Check for Service: " & Err.Description & ", " & Err.Source)
        Resume Next
    End Sub
    Private Sub LoadMailRecord(ByRef MailRec As MailNew)
        Dim X as integer
        MailIndex = MailRec.Index
        If Not MailRec.Business Then
            lblFirstName.Text = Trim(MailRec.First)
            lblLastName.Text = Trim(MailRec.Last)
            'lblLastName.Move 2640, lblLastName.Top, 2175
            lblLastName.Location = New Point(2640, lblLastName.Top)
            lblLastName.Size = New Size(2175, lblLastName.Height)
        Else
            lblFirstName.Text = ""
            lblLastName.Text = Trim(MailRec.Last)
            'lblLastName.Move lblFirstName.Left, lblLastName.Top, lblLastName.Left - lblFirstName.Left + lblLastName.Width
            lblLastName.Location = New Point(lblFirstName.Left, lblLastName.Top)
            lblLastName.Size = New Size(lblLastName.Left - lblFirstName.Left + lblLastName.Width, lblLastName.Height)
        End If
        lblAddress.Text = Trim(MailRec.Address)
        lblAddress2.Text = Trim(MailRec.AddAddress)
        lblCity.Text = Trim(MailRec.City)
        lblZip.Text = MailRec.Zip
        lblTele.Text = DressAni(CleanAni(MailRec.Tele))
        lblTele2.Text = DressAni(MailRec.Tele2)
        lblSpecial.Text = MailRec.Special

        modMail.Mail2_GetAtIndex(MailRec.Index, Mail2, StoresSld)
        lblTele3 = DressAni(Mail2.Tele3)
        UpdateTelephoneLabels(MailRec.PhoneLabel1, MailRec.PhoneLabel2, Mail2.PhoneLabel3)
    End Sub
    Private Sub ColorTaggedItem(ByVal Item as integer, ByVal Tagged As Boolean)
        If Item <= 0 Then Exit Sub
        tvItemNotes.Nodes.Item("ML" & Item).ForeColor = IIf(Tagged, Color.Red, Color.Black)
    End Sub
    Private Function IsItemTaggedForRepair(ByVal ML as integer)
        Dim CSI As clsServiceItemParts
        CSI = New clsServiceItemParts
        With CSI
            .DataAccess.Records_OpenSQL("SELECT * FROM ServiceItemParts WHERE ServiceOrderNo=" & ServiceOrderNumber & " AND MarginNo=" & ML)
            IsItemTaggedForRepair = .DataAccess.Records_Available
        End With
        DisposeDA(CSI)
    End Function
    Private Sub TagItemForRepair(ByVal Item As Integer)
        Dim CSI As clsServiceItemParts
        Dim X As CGrossMargin

        CSI = New clsServiceItemParts
        With CSI
            .DataAccess.Records_OpenSQL("SELECT * FROM ServiceItemParts WHERE ServiceOrderNo=" & ServiceOrderNumber & " AND MarginNo=" & Item)
            ' Item was already tagged, delete it.
            If .DataAccess.Records_Available Then
                ExecuteRecordsetBySQL("DELETE * FROM ServiceItemParts WHERE ServiceOrderNo=" & ServiceOrderNumber & " AND MarginNo=" & Item)
                ColorTaggedItem(Item, False)
            Else
                If ServiceOrderNumber = 0 Then
                    'cmdSave.Value = True ' Save the whole SO.. This guarantees a SO#.
                    cmdSave.PerformClick()
                End If
                .ServiceOrderNumber = ServiceOrderNumber
                .MarginNo = Item
                X = New CGrossMargin
                If X.Load(Item, "#MarginLine") Then   ' Also fill in style, etc.
                    .Style = X.Style
                    .Desc = X.Desc
                    lblSaleNo.Text = X.SaleNo
                    lblSaleNo.Visible = True
                    lblSaleNoCaption.Visible = True
                End If
                DisposeDA(X)

                .Save
                ColorTaggedItem(Item, True)
            End If
        End With

        'cmdSave.Value = True          ' SALE NO MIGHT HAVE CHANGED
        cmdSave.PerformClick()
        cmdOrderParts.Enabled = True
        DisposeDA(CSI)
    End Sub

    Private Sub UpdateTelephoneLabels(ByVal Lbl1 As String, ByVal Lbl2 As String, ByVal Lbl3 As String)
        If Trim(Lbl1) = "" Then Lbl1 = "Tele: "
        If Trim(Lbl2) = "" Then Lbl2 = "Tele2: "
        If Trim(Lbl3) = "" Then Lbl3 = "Tele3: "
        If Microsoft.VisualBasic.Right(Trim(Lbl1), 1) <> ":" Then Lbl1 = Lbl1 & ": "
        If Microsoft.VisualBasic.Right(Trim(Lbl2), 1) <> ":" Then Lbl2 = Lbl2 & ": "
        If Microsoft.VisualBasic.Right(Trim(Lbl3), 1) <> ":" Then Lbl3 = Lbl3 & ": "
        lblCapTele.Text = Lbl1
        lblCapTele2.Text = Lbl2
        lblCapTele3.Text = Lbl3
        lblTele.Left = lblCapTele.Left + lblCapTele.Width + 60
        lblCapTele2.Left = lblTele.Left + lblTele.Width + 100
        lblTele2.Left = lblCapTele2.Left + lblCapTele2.Width + 60
        lblCapTele3.Left = lblTele2.Left + lblTele2.Width + 100
        lblTele3.Left = lblCapTele3.Left + lblCapTele3.Width + 60
    End Sub

    Public Sub QuickShowServiceCall(ByVal sC As String, Optional ByVal StoreNo As Integer = 0, Optional ByVal ReturnToOriginalStore As Boolean = True)
        Dim OldMMOrder As String, OldStoreNo As Integer
        If StoreNo = 0 Then StoreNo = StoresSld
        OldStoreNo = StoresSld
        If Microsoft.VisualBasic.Left(sC, 2) = "SO" Then sC = Mid(sC, 3)

        If sC <> "" Then
            If StoreNo <> StoresSld Then
                ' Sale was in a different store.  We have to switch to view it.
                StoresSld = StoreNo
                '      main_StoreChange StoreNo
                '      frmSetup .LoadStore
                If Not ReturnToOriginalStore Then
                    MessageBox.Show("This service call was made in store " & StoreNo & "." & vbCrLf &
               "Your current login has been changed to store " & StoreNo & "." & vbCrLf &
               "You may want to change it back before continuing.", "Current Store Changed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                Else
                    MessageBox.Show("This service call was made in store " & StoreNo & ", not in your current store (store " & OldStoreNo & ")" & vbCrLf &
               "Please note that you must log into the correct store to view this call normally.", "Service Call Store Different Than Login Store", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            End If

            ' This displays the Inventory Detail's matching customer record in a disabled BillOSale.  No edits allowed.
            OldMMOrder = Order
            Order = "S"

            LoadServiceCall(sC)
            cmdMenu.Text = "B&ack"
            Show()

            Order = OldMMOrder
        End If

        If ReturnToOriginalStore And OldStoreNo <> StoresSld Then
            StoresSld = OldStoreNo
            '    main_StoreChange OldStoreNo
            '    frmSetup .LoadStore
        End If
    End Sub

End Class
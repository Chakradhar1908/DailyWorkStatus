Imports Microsoft.VisualBasic.Compatibility.VB6
Imports Microsoft.VisualBasic.Interaction
Public Class Service
    Public AccountFound As String
    Public MailIndex As Integer
    Private ServiceOrderNumber As Integer
    Private WithEvents mDBAccess As CDbAccessGeneral
    Private WithEvents mDBService As CDbAccessGeneral
    Private LoadingCheckBoxes As Boolean, SearchingSOID As Boolean
    Private Mail2 As MailNew2
    Public ServiceStatus As String
    Private StartDate As String
    Private CurrentNoteMarginNo As Integer
    Private ServicePartsLoaded As Boolean
    Private Const AllowOrderParts As Boolean = True
    Private ServiceFormLoad As Boolean
    Public ServiceFormSetRecord As Boolean
    Private FromCmdNext As Boolean

    Public Sub LoadCustomer(ByVal NewMailIndex As Integer, Optional ByVal CheckServiceCalls As Boolean = True)
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
                Service_Load(Me, New EventArgs)
                ServiceFormLoad = True
                LoadMailRecord(mR)
                FindItems()
            Else
                ' Bad mail record!
                MessageBox.Show("Invalid mail index in Service module.", "Error")
                Exit Sub
            End If
        Else
            ' Bad mail info, what to do?
            MessageBox.Show("No customer record available.", "Error")
            Exit Sub
        End If

        Dim NewCallNo As Integer
        If CheckServiceCalls = True Then
            If MailCheck.ServiceCallNo > 0 Then
                ' Load this call..
                NewCallNo = MailCheck.ServiceCallNo
            Else
                ' Find a call to work with..
                AccountFound = ""
                'CheckForService(CLng(MailIndex))   ' This should look but not load!
                CheckForService(MailIndex)   ' This should look but not load!
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

    Public Sub LoadServiceCall(ByVal SOID As Integer, Optional ByVal Direction As String = "")
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
        Dim SQL As String, Tot As Integer, Closed As Integer, N As Integer

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

    Private Sub mDBAccess_Init(Optional SOID As Integer = 0, Optional f_strDirection As String = "", Optional MailIndex As Integer = 0)
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
        MessageBox.Show("ERROR mdbAccess_Init: " & Err.Description & ", " & Err.Source)
        Resume Next
    End Sub

    Private Sub mDBService_Init()
        mDBService = New CDbAccessGeneral
        Dim a As String
        a = mDBService.dbOpen(GetDatabaseAtLocation())
    End Sub

    Private Sub mDBService_GetRecordNotFound() Handles mDBService.GetRecordNotFound
        '  MsgBox ("No Prior Service Call")
    End Sub

    Private Sub ClearServiceOrder()
        ServiceOrderNumber = 0
        lblServiceOrderNo.Text = ""

        lblSaleNo.Text = ""
        lblSaleNo.Visible = False
        lblSaleNoCaption.Visible = False

        dteServiceDate.Value = Today
        'dteServiceDate.Value = Null
        'dteServiceDate.Value = Date.FromOADate(0)
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

    Private Sub LoadCheckBoxes(ByVal Val As Integer, Optional ByVal ClearOnly As Boolean = False)
        LoadingCheckBoxes = True
        If Not ClearOnly Or Val <> 1 Then chkStoreService.Checked = IIf(Val = 1, 1, 0)
        If Not ClearOnly Or Val <> 2 Then chkOutsideService.Checked = IIf(Val = 2, 1, 0)
        If Not ClearOnly Or Val <> 3 Then chkPickupExchange.Checked = IIf(Val = 3, 1, 0)
        If Not ClearOnly Or Val <> 4 Then chkOther.Checked = IIf(Val = 4, 1, 0)
        LoadingCheckBoxes = False
    End Sub

    Private Sub CheckForService(ByVal MailIndex As Integer)
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
        MessageBox.Show("Check for Service: " & Err.Description & ", " & Err.Source)
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
        Dim Margin As CGrossMargin, Zz As Integer, ItemsUpdated As Boolean, S As String
        Dim NN As Object, Selected As Boolean
        Dim ItemDescString As String, AckInv As String
        Dim RS As ADODB.Recordset
        Dim X As ADODB.Recordset, A As Integer

        Margin = New CGrossMargin

        'Font.Name = "Arial"
        lstPurchases.Items.Clear()
        tvItemNotes.Visible = False
        tvItemNotes.Nodes.Clear()
        'tvItemNotes.Nodes.Add("", "", "LABEL",
        'ArrangeString("VENDOR", 17) & ArrangeString("STYLE", 17) & ArrangeString("SALE NO", 10) & ArrangeString("QUAN", 6) _
        '& ArrangeString("DEL DATE", 12) & ArrangeString("DESCRIPTION", 32) & "ACK/INV NO")

        tvItemNotes.Nodes.Add("LABEL", ArrangeString("VENDOR", 17) & ArrangeString("STYLE", 17) & ArrangeString("SALE NO", 10) & ArrangeString("QUAN", 6) & ArrangeString("DEL DATE", 12) & ArrangeString("DESCRIPTION", 32) & "ACK/INV NO")
        'tvItemNotes.Nodes("LABEL").Bold = True
        tvItemNotes.Nodes("LABEL").NodeFont = New Font(tvItemNotes.Font, FontStyle.Bold)

        A = Val(lblServiceOrderNo.Text)
        If A = 0 And IsFormLoaded("MailCheck") Then
            A = Val(MailCheck.ServiceCallNo)
        End If
        If A <> 0 Then
            X = GetRecordsetBySQL("SELECT * FROM ServiceItemParts WHERE ServiceOrderNo=" & A & " AND MarginNo=0")

            Do While Not X.EOF
                ItemDescString =
          ArrangeString(UCase(IfNullThenNilString(X("Vendor").Value)), 17) & ArrangeString(UCase(IfNullThenNilString(X("Style").Value)), 17) & ArrangeString(UCase(IfNullThenNilString(X("SaleNo").Value)), 10) &
          ArrangeString(IfNullThenZeroDouble(X("Quantity").Value), 6) & ArrangeString(IfNullThenZeroDate(X("DelDate").Value), 12) &
          ArrangeString(UCase(IfNullThenNilString(X("Desc").Value)), 32) & ArrangeString("", 15)

                'NN = tvItemNotes.Nodes.Add(, , "EX-" & X("STYLE").Value & "-" & Random(1000), ItemDescString)
                NN = tvItemNotes.Nodes.Add("EX-" & X("STYLE").Value & "-" & Random(1000), ItemDescString)
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
            Margin.cDataAccess_GetRecordSet(Margin.DataAccess.RS)
            'added detail 03/23/2003
            RS = GetRecordsetBySQL("SELECT * FROM Detail WHERE MarginRn=" & Margin.MarginLine & " AND Store=" & StoresSld, , GetDatabaseInventory)
            If Not RS.EOF Then
                AckInv = Trim(IfNullThenNilString(RS("Misc").Value))
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

            '--> Note: replaced above two lines with the below one. created custom class ItemDataclass to implement itemData property of vb6 in vb.net
            lstPurchases.Items.Add(New ItemDataClass(ItemDescString, Margin.Detail))

            'NN = tvItemNotes.Nodes.Add(, , "ML" & Margin.MarginLine, ItemDescString)
            'NN = tvItemNotes.Nodes.Add("", "", "ML" & Margin.MarginLine, ItemDescString)
            NN = tvItemNotes.Nodes.Add("ML" & Margin.MarginLine, ItemDescString)
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
                        ServiceNote.cDataAccess_GetRecordSet(ServiceNote.DataAccess.RS)
                        Dim Note As Object, I As Integer
                        Note = SplitLongText(" --- " & .NoteTypeString & " entered at " & DateFormat(.NoteDate) & " ---" & vbCrLf & .Note, 75)
                        For I = LBound(Note) To UBound(Note)
                            'tvItemNotes.Nodes.Add(IIf(I > LBound(Note), "SN" & .ServiceNoteID, "ML" & Margin.MarginLine), "4", "SN" & .ServiceNoteID & IIf(I > LBound(Note), "." & I, ""), Note(I))

                            Dim Relative As String, Relationship As String
                            If I > LBound(Note) Then
                                Relative = "SN" & .ServiceNoteID
                                Relationship = 4 'tvwChild
                            Else
                                Relative = "ML" & Margin.MarginLine
                                Relationship = 4 'tvwChild
                            End If
                            'tvItemNotes.Nodes(0).Nodes.Add("SN" & .ServiceNoteID & IIf(I > LBound(Note), "." & I, ""), Note(I))
                            'tvItemNotes.Nodes(IIf(I > LBound(Note), "SN" & .ServiceNoteID, "ML" & Margin.MarginLine)).Expanded = True
                            tvItemNotes.SelectedNode.Nodes.Add("SN" & .ServiceNoteID & IIf(I > LBound(Note), "." & I, ""), Note(I))
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
                cSR.Save()
            Else
                ' How can it not load, we're in it?
                MessageBox.Show("Error upgrading service record structure.", "Error")
            End If
            DisposeDA(cSR, Margin)
        End If

        tvItemNotes.Visible = True

        LoadPartsOrders()
        'DisposeDA(cSR, Margin)

        Exit Sub

        LoadPartsOrders()
HandleErr:
        MessageBox.Show("Check for Service: " & Err.Description & ", " & Err.Source)
        Resume Next
    End Sub

    Private Sub LoadMailRecord(ByRef MailRec As MailNew)
        Dim X As Integer
        MailIndex = MailRec.Index
        If Not MailRec.Business Then
            lblFirstName.Text = Trim(MailRec.First)
            lblLastName.Text = Trim(MailRec.Last)
            'lblLastName.Move 2640, lblLastName.Top, 2175
            lblLastName.Location = New Point(180, lblLastName.Top)
            lblLastName.Size = New Size(152, lblLastName.Height)
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
        lblTele3.Text = DressAni(Mail2.Tele3)
        UpdateTelephoneLabels(MailRec.PhoneLabel1, MailRec.PhoneLabel2, Mail2.PhoneLabel3)
    End Sub

    Private Sub ColorTaggedItem(ByVal Item As Integer, ByVal Tagged As Boolean)
        If Item <= 0 Then Exit Sub
        tvItemNotes.Nodes.Item("ML" & Item).ForeColor = IIf(Tagged, Color.Red, Color.Black)
    End Sub

    Private Function IsItemTaggedForRepair(ByVal ML As Integer)
        Dim CSI As clsServiceItemParts
        CSI = New clsServiceItemParts

        CSI.DataAccess.Records_OpenSQL("SELECT * FROM ServiceItemParts WHERE ServiceOrderNo=" & ServiceOrderNumber & " AND MarginNo=" & ML)
        IsItemTaggedForRepair = CSI.DataAccess.Records_Available
        If IsItemTaggedForRepair = True Then
            CSI.cDataAccess_GetRecordSet(CSI.DataAccess.RS)
        End If
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

                .Save()
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
        lblTele.Left = lblCapTele.Left + lblCapTele.Width + 2
        lblCapTele2.Left = lblTele.Left + lblTele.Width + 40
        lblTele2.Left = lblCapTele2.Left + lblCapTele2.Width + 2
        lblCapTele3.Left = lblTele2.Left + lblTele2.Width + 100
        lblTele3.Left = lblCapTele3.Left + lblCapTele3.Width + 2
    End Sub

    Private Sub cmdAddItem_Click(sender As Object, e As EventArgs) Handles cmdAddItem.Click
        Dim SaleNo As String, Style As String, Desc As String, Quan As Double, Vendor As String, DelDate As Date
        Dim T() As Object
        Dim S As String

        If MessageBox.Show("This will add an item to this service call that is not in the inventory Database.", "Add Item?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Cancel Then Exit Sub

        'If lblServiceOrderNo = "" Then cmdSave.Value = True
        'If lblServiceOrderNo.Text = "" Then cmdSave.Value = True
        If lblServiceOrderNo.Text = "" Then cmdSave_Click(cmdSave, New EventArgs)

        'ServiceManualItem.Show 1
        ServiceManualItem.ShowDialog()

        'If ServiceManualItem.Cancelled Then Unload ServiceManualItem: Exit Sub
        If ServiceManualItem.Cancelled Then ServiceManualItem.Close() : Exit Sub

        SaleNo = UCase(ServiceManualItem.txtSaleNo.Text)
        Style = UCase(ServiceManualItem.txtStyle.Text)
        Desc = UCase(ServiceManualItem.txtDesc.Text)
        Quan = Val(ServiceManualItem.txtQuantity.Text)
        Vendor = UCase(ServiceManualItem.cboVendor.Text)
        DelDate = ServiceManualItem.dtpDelDate.Value

        'Unload ServiceManualItem
        ServiceManualItem.Close()

        S = ""
        S = S & "INSERT INTO [ServiceItemParts] "
        S = S & "([ServiceOrderNo], [MarginNo], [Style], [Desc], [Vendor], [SaleNo], [DelDate], [Quantity]) "
        S = S & "VALUES (" & lblServiceOrderNo.Text & ", 0, '" & ProtectSQL(Style) & "', '" & ProtectSQL(Desc) & "', '" & ProtectSQL(Vendor) & "', '" & ProtectSQL(SaleNo) & "', '" & ProtectSQL(DelDate) & "', " & Quan & ")"
        ExecuteRecordsetBySQL(S, , GetDatabaseAtLocation)
        MessageBox.Show("Item Added to service call #" & lblServiceOrderNo.Text, "Operation Completed.", MessageBoxButtons.OK, MessageBoxIcon.Information)
        S = lblServiceOrderNo.Text
        FindItems()
    End Sub
    Delegate Function d()
    Private Sub cmdMenu_Click(sender As Object, e As EventArgs) Handles cmdMenu.Click
        ClearServiceOrder()
        If cmdMenu.Text = "&Menu" Then
            modProgramState.Order = ""
            MainMenu.Show()
        End If
        'Unload Me
        'Me.Close() This line is throwing error of Notes_Frame frme is running on different thread. To clear the error, replaced the line with the below me.Invoke code.
        Me.Invoke(Sub()
                      Me.Close()
                  End Sub)
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

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        ClearServiceOrder()
        'Unload Me
        FromCmdNext = True
        Me.Close()
        FromCmdNext = False
        'MailCheck.optTelephone.Checked = True
        MailCheckSaleNoChecked = False
        'MailCheck.Show vbModal
        MailCheck.ShowDialog()
    End Sub

    Private Sub cmdMoveSearch_Click(sender As Object, e As EventArgs) Handles cmdMoveSearch.Click
        Dim X As Integer
        X = Val(InputBox("Search for ServiceOrder:", "New Service Order Number"))
        If X <= 0 Then Exit Sub
        SearchingSOID = True
        LoadServiceCall(X)
        SearchingSOID = False
    End Sub

    Private Sub EnableNavigation(ByVal OnOff As Boolean)
        cmdMenu.Enabled = OnOff
        cmdNext.Enabled = OnOff
        cmdMoveFirst.Enabled = OnOff
        cmdMovePrevious.Enabled = OnOff
        cmdMoveNext.Enabled = OnOff
        cmdMoveLast.Enabled = OnOff
        cmdMoveSearch.Enabled = OnOff
    End Sub

    Public Sub PartsOrderFormClosed()
        ServicePartsLoaded = False
        FindItems()
        EnableNavigation(True)
        ' Update lit-up parts order display.
    End Sub

    Private Sub cmdOrderParts_Click(sender As Object, e As EventArgs) Handles cmdOrderParts.Click
        Dim X As Integer
        X = SelectedMarginNode()

        If X = 0 Or Not MLItemIsTagged(X) Then
            MessageBox.Show("Please select a tagged item to order parts for.", "Wait!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        'order parts

        ' If there's more than one parts order, show AddOnAcc.

        ServicePartsLoaded = True ' Remember that we loaded ServiceParts.
        EnableNavigation(False)    ' Disable navigation while ServiceParts is showing.

        'ServiceParts.HelpContextID = 0 ' forces form_load
        ServiceParts.SetOwner(Me)  ' Make ServiceParts behave as a child form.

        ' ServiceParts will get this directly from the service call table.
        ServiceParts.lblFirstName = lblFirstName
        ServiceParts.lblLastName = lblLastName
        ServiceParts.lblAddress = lblAddress
        ServiceParts.lblAddress2 = lblAddress2
        ServiceParts.lblCity = lblCity
        ServiceParts.lblZip = lblZip
        ServiceParts.lblTele1Caption = lblCapTele
        ServiceParts.lblTele2Caption = lblCapTele2
        ServiceParts.lblTele3Caption = lblCapTele3
        ServiceParts.lblTele = lblTele
        ServiceParts.lblTele2 = lblTele2
        ServiceParts.lblTele3 = lblTele3
        If X > 0 Then
            ServiceParts.LoadInfoFromMarginLine(X)
        Else
            'ServiceParts.txtStyleNo = Trim(Mid(tvItemNotes.SelectedItem, 18, 17))
            ServiceParts.txtStyleNo.Text = Trim(Mid(tvItemNotes.SelectedNode.Text, 18, 17))
            'ServiceParts.txtDescription = Trim(Mid(tvItemNotes.SelectedItem, 63, 32))
            ServiceParts.txtDescription.Text = Trim(Mid(tvItemNotes.SelectedNode.Text, 63, 32))
        End If

        ' Tell ServiceParts it's working from this Service Call.
        If ServiceParts.LoadServiceCall(ServiceOrderNumber) Then
            ServiceParts.Show()
            ServiceParts.LoadRelativePartsOrder(-1, True, True)
            Hide()
        Else
            ' This should clean up ServicePartsLoaded and related flags..
            'Unload ServiceParts
            ServiceParts.Close()
            MessageBox.Show("Error ordering parts: Can't load current service call.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Function SelectedMarginNode() As Integer
        Dim CurRow As String
        'If tvItemNotes.SelectedItem Is Nothing Then Exit Function
        If tvItemNotes.SelectedNode Is Nothing Then Exit Function
        'CurRow = tvItemNotes.SelectedItem.Key
        CurRow = tvItemNotes.SelectedNode.Name


        If Microsoft.VisualBasic.Left(CurRow, 2) = "ML" Then
            SelectedMarginNode = Val(Mid(CurRow, 3))
        ElseIf Microsoft.VisualBasic.Left(CurRow, 2) = "EX" Then
            SelectedMarginNode = -1
        ElseIf Microsoft.VisualBasic.Left(CurRow, 2) = "SN" Then
            'CurRow = tvItemNotes.Nodes(CurRow).Parent.Key
            CurRow = tvItemNotes.Nodes(CurRow).Parent.SelectedImageKey
            'If Microsoft.VisualBasic.Left(CurRow, 2) = "SN" Then CurRow = tvItemNotes.Nodes(CurRow).Parent.Key
            If Microsoft.VisualBasic.Left(CurRow, 2) = "SN" Then CurRow = tvItemNotes.Nodes(CurRow).Parent.SelectedImageKey
            If Microsoft.VisualBasic.Left(CurRow, 2) <> "ML" Then
                Exit Function
            Else
                SelectedMarginNode = Val(Mid(CurRow, 3))
            End If
        Else
            Exit Function
        End If
    End Function

    Private Function MLItemIsTagged(ByVal MLItem As Integer) As Boolean
        If MLItem = 0 Then Exit Function
        If MLItem = -1 Then MLItemIsTagged = True : Exit Function
        MLItemIsTagged = tvItemNotes.Nodes.Item("ML" & MLItem).ForeColor = Color.Red
    End Function

    Private Sub cmdRepairTag_Click(sender As Object, e As EventArgs) Handles cmdRepairTag.Click
        Dim P As String
        If lblServiceOrderNo.Text = "" Then
            MessageBox.Show("You can only print a repair tag for active service calls.", "No Service Order", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        P = Printer.DeviceName
        If Not SetDymoPrinter() Then
            MessageBox.Show("Dymo Printer Required!", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Printer.FontName = "Arial"
        Printer.FontSize = 18
        Printer.FontBold = True
        Printer.FontUnderline = True
        'PrintAligned("REPAIR TAG", VBRUN.AlignConstants.vbAlignLeft, 0) '  vbCenter, Printer.ScaleWidth / 2
        PrintAligned2("REPAIR TAG", VBRUN.AlignConstants.vbAlignLeft, 0) '  vbCenter, Printer.ScaleWidth / 2
        Printer.FontUnderline = False
        Printer.FontSize = 14
        'PrintAligned("SrvOrd#: " & lblServiceOrderNo.Text & " [" & dteServiceDate.Value & "]")
        PrintAligned2("SrvOrd#: " & lblServiceOrderNo.Text & " [" & dteServiceDate.Value & "]")
        'PrintAligned("Claim Date: " & lblClaimDate.Text)
        PrintAligned2("Claim Date: " & lblClaimDate.Text)
        'PrintAligned("SaleNo: " & lblSaleNo.Text)
        PrintAligned2("SaleNo: " & lblSaleNo.Text)
        'PrintAligned("Name: " & lblLastName.Text)
        PrintAligned2("Name: " & lblLastName.Text)
        'PrintAligned "Phone: " & lblTele
        'PrintAligned(lblCapTele.Text & lblTele.Text)
        PrintAligned2(lblCapTele.Text & lblTele.Text)
        'PrintAligned(lblCapTele2.Text & lblTele2.Text)
        PrintAligned2(lblCapTele2.Text & lblTele2.Text)
        'PrintAligned(lblTele3.Text) 'lblCapTele3 & lblTele3
        PrintAligned2(lblTele3.Text) 'lblCapTele3 & lblTele3
        'PrintAligned("Type: " & Switch(chkStoreService.Checked = True, "Store", chkOutsideService.Checked = True, "Outside", chkPickupExchange.Checked = True, "P-Up/Exg", chkOther.Checked = True, "Other", True, "Other"))
        PrintAligned2("Type: " & Switch(chkStoreService.Checked = True, "Store", chkOutsideService.Checked = True, "Outside", chkPickupExchange.Checked = True, "P-Up/Exg", chkOther.Checked = True, "Other", True, "Other"))
        '  PrintAligned "Serv Date: " & dteServiceDate

        Printer.EndDoc()
        SetPrinter(P)
    End Sub

    Private Sub cmdSaveItemNote_Click(sender As Object, e As EventArgs) Handles cmdSaveItemNote.Click
        Dim NewID As Integer
        ' Validate, save note, and hide item notes frame.
        If CurrentNoteMarginNo = 0 Then
            MessageBox.Show("Error: Can't determine which item to save note for.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If txtItemNotes.Text = "" Then
            If MessageBox.Show("You can't save a blank note.  Would you like to enter one now?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                Exit Sub
            Else
                'cmdCancelItemNote.Value = True
                cmdCancelItemNote_Click(cmdCancelItemNote, New EventArgs)
                Exit Sub
            End If
        End If

        NewID = CreateServiceNote(CurrentNoteMarginNo, ServiceOrderNumber, txtItemNotes.Text)

        ' Refresh the items/notes treeview, making sure the new note is visible.
        FindItems()
        'tvItemNotes.Nodes("SN" & NewID).EnsureVisible()
        tvItemNotes.SelectedNode.Nodes("SN" & NewID).EnsureVisible()

        Notes_Frame.Visible = True
        ItemNotesFrame.Visible = False
        CurrentNoteMarginNo = 0
        lblItemNotesCaption.Text = ""
    End Sub

    Private Sub cmdCancelItemNote_Click(sender As Object, e As EventArgs) Handles cmdCancelItemNote.Click
        ' Hide item notes frame.
        Notes_Frame.Visible = True
        ItemNotesFrame.Visible = False
        CurrentNoteMarginNo = 0
        lblItemNotesCaption.Text = ""
        txtItemNotes.Text = ""
    End Sub

    Private Sub cmdAddItemNote_Click(sender As Object, e As EventArgs) Handles cmdAddItemNote.Click
        Dim MLRow As Integer
        On Error GoTo SayNo
        MLRow = SelectedMarginNode()
        If MLRow = 0 Then MessageBox.Show("Please select an item from the list first.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation) : Exit Sub
        CurrentNoteMarginNo = MLRow
        ItemNotesFrame.Visible = True
        Notes_Frame.Visible = False
        lblItemNotesCaption.Text = tvItemNotes.Nodes("ML" & MLRow).Text
        txtItemNotes.Text = ""
        Exit Sub

SayNo:
        MessageBox.Show("You cannot add a note to this item.", "Not Available", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ItemNotesFrame.Visible = False
        Notes_Frame.Visible = True
    End Sub

    Private Sub cmdTagForRepair_Click(sender As Object, e As EventArgs) Handles cmdTagForRepair.Click
        Dim ISelected As Integer
        ' Get the selected ML.
        ISelected = SelectedMarginNode()
        If ISelected = 0 Then MessageBox.Show("Please select an item from the list first.", "WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation) : Exit Sub
        TagItemForRepair(ISelected)
    End Sub

    Public Sub InitStatusList()
        cboStatus.Items.Clear()
        cboStatus.Items.Insert(0, "Open")
        cboStatus.Items.Insert(1, "Closed")
        cboStatus.SelectedIndex = 0
    End Sub

    Private Sub ShowTimeWindowBox(ByVal Show As Boolean, Optional ByVal Enabled As Boolean = False)
        'dtpDelWindow0.Value = "7:00 am"
        dtpDelWindow0.Value = DateTime.ParseExact("7:00 AM", "h:mm tt", System.Globalization.CultureInfo.InvariantCulture)
        'dtpDelWindow0.Value = ""
        'dtpDelWindow1.Value = "5:00 pm"
        dtpDelWindow1.Value = DateTime.ParseExact("5:00 PM", "h:mm tt", System.Globalization.CultureInfo.InvariantCulture)
        'dtpDelWindow1.Value = ""

        fraTimeWindow.Visible = Show And (StoreSettings.bUseTimeWindows)
        dtpDelWindow0.Enabled = Enabled
        dtpDelWindow1.Enabled = Enabled
    End Sub

    'This closeup event is replacement for change event of vb6.0 datetimepicker.
    Private Sub dtpDelWindowCloseUp(sender As Object, e As EventArgs) Handles dtpDelWindow0.CloseUp, dtpDelWindow1.CloseUp
        Dim D1 As Date, D2 As Date

        If IsDate(dtpDelWindow0.Value) And IsDate(dtpDelWindow1.Value) Then
            D1 = TimeValue(dtpDelWindow0.Value)
            If DateAfter(D1, "11:00p", False, "n") Then
                dtpDelWindow0.Value = "10:00 pm"
                D1 = TimeValue(dtpDelWindow0.Value)
            End If

            D2 = TimeValue(dtpDelWindow1.Value)
            If Not DateAfter(D2, D1, False, "n") Then
                dtpDelWindow1.Value = TimeValue(DateAdd("n", 30, D1))
            End If
        End If
    End Sub

    Private Sub dteServiceDate_CloseUp(sender As Object, e As EventArgs) Handles dteServiceDate.CloseUp
        'ShowTimeWindowBox(IsDate(dteServiceDate.Value), True)
        'If dteServiceDate.Value > Date.FromOADate(0) Then
        '    dteServiceDate.Checked = True
        '    ShowTimeWindowBox(True, True)
        'Else
        '    dteServiceDate.Checked = False
        '    ShowTimeWindowBox(False, True)
        'End If
    End Sub

    Private Sub Service_Load(sender As Object, e As EventArgs) Handles Me.Load
        If ServiceFormLoad = True Then Exit Sub
        'SetButtonImage cmdMoveFirst, "previous"
        SetButtonImage(cmdMoveFirst, 7)
        'SetButtonImage(cmdMovePrevious, "previous1")
        SetButtonImage(cmdMovePrevious, 4)
        'SetButtonImage cmdMoveNext, "next1"
        SetButtonImage(cmdMoveNext, 5)
        'SetButtonImage cmdMoveLast, "next"
        SetButtonImage(cmdMoveLast, 6)

        'SetButtonImage cmdSave
        SetButtonImage(cmdSave, 2)
        'SetButtonImage cmdPrint
        SetButtonImage(cmdPrint, 19)
        'SetButtonImage cmdNext
        SetButtonImage(cmdNext, 6)
        'SetButtonImage cmdMenu
        SetButtonImage(cmdMenu, 9)

        'SetButtonImage cmdSaveItemNote, "ok"
        SetButtonImage(cmdSaveItemNote, 2)
        'SetButtonImage cmdCancelItemNote, "cancel"
        SetButtonImage(cmdCancelItemNote, 3)

        'Testing
        'Left = (Screen.Width - Width) / 2
        Left = (Me.ClientSize.Width - Width) / 2
        lblClaimDate.Text = DateFormat(Today)
        StartDate = DateFormat(Today)
        'dteServiceDate.Value = Date.FromOADate(0)
        'dteServiceDate.Value = Today
        Notes_Frame.Visible = True
        InitStatusList()

        cmdOrderParts.Visible = False 'demo
        cmdOrderParts.Visible = AllowOrderParts
        lblPartsOrd.Visible = False

        On Error Resume Next
        'imgLogo.Picture = LoadPictureStd(StoreLogoFile())
        'imgLogo.Image = LoadPictureStd(StoreLogoFile())
        imgLogo.Image = Image.FromFile(StoreLogoFile())
    End Sub

    'form unload of vb6.0
    Private Sub Service_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Query Unload event of vb6.0
        'If UnloadMode = vbFormControlMenu Then cmdMenu.Value = True
        If FromCmdNext = False Then
            If e.CloseReason = CloseReason.UserClosing Then cmdMenu.PerformClick()
        End If
        'Form unload event of vb6.0
        On Error Resume Next
        mDBAccess.dbClose()
        mDBService.dbClose()
        mDBAccess = Nothing
        mDBService = Nothing
    End Sub

    Private Sub GetServiceNo()
        Dim SerNo As Integer
        SerNo = GetFileAutonumber(SerNoFile, 1001)
        lblServiceOrderNo.Text = SerNo
        ServiceOrderNumber = SerNo
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim QuickCheck As Integer
        QuickCheck = GetCheckBoxValue()

        Printer.FontName = "Arial"
        Printer.FontSize = 13
        Printer.DrawWidth = 2
        Printer.FontBold = True
        Printer.CurrentY = 100
        Printer.FontItalic = True
        PrintCentered("-Service Request-")
        Printer.FontItalic = False

        Printer.CurrentY = 100
        Printer.CurrentX = 8000
        Printer.FontSize = 11
        Printer.Print("Tech.___________________________")


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   Logo (center)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Printer.CurrentY = 600
        If imgLogo.Image Is Nothing Then
            Printer.FontSize = 18
            PrintCentered(StoreSettings.Name)
            PrintCentered(StoreSettings.Address)
            PrintCentered(StoreSettings.City)
            PrintCentered(StoreSettings.Phone)
        Else  'logo
            Printer.CurrentX = 4000
            'Printer.PaintPicture(imgLogo.Image, Printer.Width / 2 - imgLogo.Width / 2, 390, imgLogo.Width, imgLogo.Height)
            'Printer.PaintPicture(Image.FromFile(StoreLogoFile(0)), 4000, 200, 5000, 5000, 1200, 1000, 35000, 35000)
            Printer.PaintPicture(Image.FromFile(StoreLogoFile()), 4000, 400, 5000, 5000, 1200, 1000, 35000, 35000)
        End If

        Printer.FontBold = True
        Printer.CurrentX = 400
        Printer.CurrentY = 350
        Printer.FontSize = 10
        Printer.Print(" SERVICE ON:")
        Printer.DrawWidth = 8
        'Printer.Line(500, 600)-Step(2000, 1200), QBColor(0), B
        Printer.Line(500, 600, 2500, 1800, QBColor(0), True)
        Printer.DrawWidth = 1

        Printer.FontSize = 18
        Printer.CurrentX = 610
        Printer.CurrentY = 650

        If dteServiceDate.Value.ToString <> "" Then
            Printer.Print(Microsoft.VisualBasic.Left(dteServiceDate.Value, 10))
            Printer.FontSize = 14
            Printer.CurrentX = 1000
            Dim y As Integer
            y = Printer.CurrentY
            Printer.Print(Format(dteServiceDate.Value, "DDDD"))
            Printer.CurrentX = 610
            'PrintInBox(Printer, DescribeTimeWindow(dtpDelWindow0.Value, dtpDelWindow1.Value), 600, Printer.CurrentY, 1800, 300) <CT>'PrintInBox is notworking. Replaced it with direct Printer.Print(DescribeTimeWindow(dtpDelWindow0.Value, dtpDelWindow1.Value))</CT>
            Printer.FontSize = 8
            Printer.CurrentY = Printer.CurrentY + 40
            Printer.Print(DescribeTimeWindow(dtpDelWindow0.Value, dtpDelWindow1.Value))
        End If

        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.CurrentX = 10000
        Printer.CurrentY = 400
        Printer.Print("Service Order:")
        Printer.CurrentX = 10200
        Printer.FontSize = 14
        Printer.Print(ServiceOrderNumber)
        Printer.FontBold = False
        Printer.FontSize = 10

        Printer.CurrentX = 10000
        Printer.CurrentY = 1000
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.Print("Date Of Claim:")
        Printer.CurrentX = 10075
        Printer.Print(lblClaimDate.Text)

        Printer.CurrentX = 200
        Printer.CurrentY = 2600
        Printer.FontSize = 6
        If lblLastName.Text <> "" And lblFirstName.Text = "" Then
            Printer.Print("Business Name")
        Else
            Printer.Print("First Name", TAB(78), "Last Name")
        End If
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 200
        Printer.Print("Address")
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 200
        Printer.Print("City / State", TAB(100), "Zip")
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 200
        'Printer.Print "Telephone1"; Tab(58); "Telephone2"
        Printer.Print(IIf(lblCapTele.Text = "Tele: ", "Telephone1", lblCapTele.Text), TAB(120), IIf(lblCapTele2.Text = "Tele2: ", "Telephone2", lblCapTele2.Text)) '; Tab(58); IIf(lblCapTele3 = "Tele3: ", "Telephone3", lblCapTele3)
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 200
        Printer.Print()
        Printer.CurrentX = 200
        Printer.CurrentY = 5000
        Printer.FontSize = 10
        Printer.Print("Special Instructions   ")

        '.DrawWidth = 10 'box for spec inst.
        'Printer.Line (150, 5300)-Step(11220, 350), QBColor(0), B

        Dim SpInstLines() As String, Sp As String, I As Integer
        Sp = lblSpecial.Text
        Sp = WrapLongTextByPrintWidth(Printer, Sp, Printer.ScaleWidth, vbCrLf)
        SpInstLines = Split(Sp, vbCrLf)
        For I = LBound(SpInstLines) To UBound(SpInstLines)
            If I - LBound(SpInstLines) > 2 Then Exit For
            Printer.CurrentX = 400 ': Printer.CurrentY = Printer.CurrentY - 300
            Printer.Print(SpInstLines(I)) ' lblSpecial
        Next

        ' Ship to
        Printer.CurrentX = 6200 : Printer.CurrentY = 2400
        Printer.FontSize = 14
        Printer.Print("                SHIP TO ADDRESS:")

        Printer.FontSize = 6

        Printer.CurrentX = 6200 : Printer.CurrentY = 3139
        Printer.Print("Address")
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 6200
        'Printer.Print("City / State", SPC(58), "Zip")
        Printer.Print("City / State", SPC(88), "Zip")
        Printer.Print()
        Printer.Print()
        Printer.Print()
        Printer.CurrentX = 6200
        'Printer.Print "Telephone3 "
        Printer.Print(Mail2.PhoneLabel3)

        Printer.CurrentX = 200 : Printer.CurrentY = 5350
        'special inst
        Printer.FontSize = 10
        Printer.Print(MailCheck.SpecialIns)  ' This won't be available..

        Printer.CurrentX = 200 : Printer.CurrentY = 5350
        ' special inst
        Printer.FontSize = 10

        Printer_Location(200, 2700, 14) 'name
        If lblFirstName.Text = "" And lblLastName.Text <> "" Then
            Printer.Print(lblLastName.Text)
        Else
            Printer.Print(lblFirstName.Text, TAB(29), lblLastName.Text)
        End If

        Printer_Location(200, 3250, 12, lblAddress.Text) 'address
        Printer_Location(200, 3500, 12, lblAddress2.Text)

        Printer_Location(200, 4050, 12)
        Printer.Print(Trim(lblCity.Text), TAB(40), Trim(lblZip.Text))

        Printer_Location(200, 4600, 12)
        Printer.Print(lblTele.Text, TAB(25), lblTele2.Text)

        Printer_Location(6200, 2700, 12)
        Printer.Print(Mail2.ShipToFirst & "     " & Mail2.ShipToLast)

        Printer_Location(6200, 3250, 12)
        Printer.Print(Trim(Mail2.Address2))

        Printer_Location(6200, 4050, 12)
        Printer.Print(Trim(Mail2.City2), TAB(40), Trim(Mail2.Zip2))

        Printer_Location(6200, 4600, 12, DressAni(CleanAni(Mail2.Tele3)))

        'Check boxes
        Printer.FontSize = 14
        Printer.FontBold = True
        Printer.DrawWidth = 12

        If QuickCheck = 1 Then SetColor()  'get from data base
        'Printer.Line(500, 5900)-Step(500, 500), QBColor(0), B
        'Printer.Line(500, 600, 2500, 1800, QBColor(0), True)

        'Printer.Line(500, 5900, 500, 500, QBColor(0), True)
        'Printer.Line(500, 5900, 900, 8000, QBColor(0), True)
        Printer.Line(500, 5900, 1000, 6400, QBColor(0), True)
        If QuickCheck = 1 Then EndColor()

        If QuickCheck = 2 Then SetColor()  'get from data base
        'Printer.Line(3200, 5900, 500, 500, QBColor(0), True)
        Printer.Line(3200, 5900, 3700, 6400, QBColor(0), True)
        If QuickCheck = 2 Then EndColor()

        If QuickCheck = 3 Then SetColor()  'get from data base
        'Printer.Line(6000, 5900, 500, 500, QBColor(0), True)
        Printer.Line(6000, 5900, 6500, 6400, QBColor(0), True)
        If QuickCheck = 3 Then EndColor()

        If QuickCheck = 4 Then SetColor()  'get from data base
        'Printer.Line(9200, 5900, 500, 500, QBColor(0), True)
        Printer.Line(9200, 5900, 9700, 6400, QBColor(0), True)
        If QuickCheck = 4 Then EndColor()

        Printer.CurrentY = 6000
        Printer.CurrentX = 1100
        Printer.FontSize = 12

        Printer.Print("Store Service", TAB(32), "Outside Service", TAB(61), "Pick Up & Exchange", TAB(92), "Other")

        Printer.CurrentY = 7000
        Printer.CurrentX = 150
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.Print("Items Reported: ")

        Printer.CurrentX = 0 '1000
        Printer.FontSize = 9
        Printer.Print("  Vendor", TAB(30), "Style", TAB(65), "Sale No", TAB(80), "Quan", TAB(90), "Del Date", TAB(110), "Description", TAB(170), "Inv/Ack No.")
        Printer.FontBold = False
        Printer.FontName = "Courier New"
        Printer.FontSize = 9
        Printer.FontBold = True

        'Printer.Print txtItems 'items  and notes..
        Dim ind As Integer, PrintNotes As Boolean
        'For ind = 1 To tvItemNotes.Nodes.Count
        For ind = 0 To tvItemNotes.Nodes.Count - 1
            If IsIn(Microsoft.VisualBasic.Left(tvItemNotes.Nodes(ind).Name, 2), "ML", "EX") Then
                If tvItemNotes.Nodes(ind).ForeColor = Color.Red Then
                    Printer.Print(tvItemNotes.Nodes(ind).Text)
                    'Printer.Print(TAB(5), tvItemNotes.SelectedNode.Nodes(ind).Text)
                    'Printer.Print(TAB(5), tvItemNotes.SelectedNode.Text)
                    Dim Cn As Integer
                    For Cn = 0 To tvItemNotes.Nodes(ind).Nodes.Count - 1 '<CT> printing Notes </CT>
                        Printer.Print(TAB(5), tvItemNotes.Nodes(ind).Nodes(Cn).Text)
                    Next
                    ' Print this line..
                    PrintNotes = True
                Else
                    PrintNotes = False
                End If
                'ElseIf PrintNotes And (Microsoft.VisualBasic.Left(tvItemNotes.Nodes(ind).Name, 2)) = "SN" Then
            ElseIf PrintNotes And (Microsoft.VisualBasic.Left(tvItemNotes.SelectedNode.Nodes(ind).Name, 2)) = "SN" Then '<CT> This condition will not work. So replaced it with the above for loop to display service notes </CT>
                ' Print this line.
                'Printer.Print(TAB(5), tvItemNotes.Nodes(ind).Text)
                Printer.Print(TAB(5), tvItemNotes.SelectedNode.Nodes(ind).Text)
            End If
        Next

        Printer.FontBold = False
        Printer.FontName = "Arial"
        Printer.Print()
        Printer.Print()
        Printer.FontSize = 14
        Printer.CurrentX = 150
        Printer.FontBold = True
        Printer.Print("Customer Complaint:")
        Printer.FontBold = False
        Printer.FontSize = 12
        Printer.CurrentX = 0
        Printer.Print(WrapLongText(Notes_Text.Text, 92))

        Printer.Print()
        Printer.Print()
        Printer.FontSize = 14
        Printer.CurrentX = 150
        Printer.FontBold = True
        Printer.Print("Store Response:")
        Printer.FontBold = False
        Printer.FontSize = 12
        Printer.CurrentX = 0
        Printer.Print(WrapLongText(Notes_New.Text, 92))

        Printer.Print()
        Printer.Print()
        Printer.FontSize = 14
        Printer.CurrentX = 150
        Printer.FontBold = True
        Printer.Print("Technician's Report:")
        Printer.FontBold = False
        Printer.FontSize = 12
        Printer.CurrentX = 0

        Printer.FontSize = 10
        Printer.CurrentY = 14400 '14500
        Printer.CurrentX = 0

        Printer.FontBold = True
        Printer.Print("______________________________ ")

        Printer.CurrentX = 3200
        Printer.CurrentY = 14400 '14500
        Printer.FontBold = False
        Printer.Print(TAB(5), "Technician:______________________ Hours:__________ Charges:__________")

        Printer.FontBold = True
        Printer.Print("     Customer Satisfied:")
        Printer.FontBold = False
        Printer.EndDoc()
    End Sub

    Private Function GetCheckBoxValue() As Integer
        If chkOther.Checked = True Then GetCheckBoxValue = 4
        If chkPickupExchange.Checked = True Then GetCheckBoxValue = 3
        If chkOutsideService.Checked = True Then GetCheckBoxValue = 2
        If chkStoreService.Checked = True Then GetCheckBoxValue = 1
    End Function

    Private Sub SetColor()
        Printer.FillColor = QBColor(0)
        Printer.FillStyle = QBColor(0)
    End Sub

    Private Sub EndColor()
        Printer.FillColor = QBColor(15)
    End Sub

    Private Sub lstPurchases_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles lstPurchases.ItemCheck
        Dim InvoiceNo As String
        Dim InvDetail As CInventoryDetail
        InvDetail = New CInventoryDetail

        ' check for detail (S/O) and if so get invoice no.
        'If lstPurchases.ListCount < lstPurchases.ListIndex Or lstPurchases.ListIndex < 0 Then Exit Sub
        If lstPurchases.Items.Count < lstPurchases.SelectedIndex Or lstPurchases.SelectedIndex < 0 Then Exit Sub

        ' Later: load this info into the grid at search time.
        'If InvDetail.Load(CStr(lstPurchases.itemData(lstPurchases.ListIndex)), "#DetailID") Then
        If InvDetail.Load(CType(lstPurchases.Items(lstPurchases.SelectedIndex), ItemDataClass).ItemData, "#DetailID") Then
            InvoiceNo = InvDetail.Misc
            If InvoiceNo <> "" Then
                If Asc(InvoiceNo) = 0 Then InvoiceNo = ""
            End If
        Else
            ' Temporarily removed for demo, 20030716
            '    MsgBox "Invalid Detail Item in Service.lstPurchases_ItemCheck."
            InvoiceNo = ""
        End If
        DisposeDA(InvDetail)

        'If lstPurchases.Selected(Item) Then
        If lstPurchases.GetSelected(e.Index) Then
            'txtItems = txtItems & lstPurchases.List(lstPurchases.ListIndex) & "  " & InvoiceNo & vbCrLf
            txtItems.Text = txtItems.Text & lstPurchases.Items(lstPurchases.SelectedIndex) & "  " & InvoiceNo & vbCrLf
        Else
            'txtItems = Replace(txtItems.Text, lstPurchases.List(lstPurchases.ListIndex) & "  " & InvoiceNo & vbCrLf, "")
            txtItems.Text = Replace(txtItems.Text, lstPurchases.Items(lstPurchases.SelectedIndex) & "  " & InvoiceNo & vbCrLf, "")
        End If
    End Sub

    Private Sub Printer_Location(X As Single, Y As Single, FontSize As Integer, Optional Prt As String = "")
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.FontSize = FontSize
        If Len(Prt) <> 0 Then Printer.Print(Prt)
    End Sub

    Private Sub chkStoreService_CheckedChanged(sender As Object, e As EventArgs) Handles chkStoreService.CheckedChanged
        If LoadingCheckBoxes Then Exit Sub
        If chkStoreService.Checked = False Then chkStoreService.Checked = True : Exit Sub
        LoadCheckBoxes(1, True)
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        'save
        'MailIndex = MailCheck.Index  ' This should already be stored in Service!
        If ServiceOrderNumber < 1 Then 'new service
            mDBAccess_Init(, , MailIndex)
            'Service.ServiceStatus = "Open"
            ServiceFormSetRecord = True
            mDBAccess.SetRecord(True) ' this sets NEW record
            ServiceFormSetRecord = False
        Else
            mDBAccess_Init(ServiceOrderNumber)
            ServiceFormSetRecord = True
            mDBAccess.SetRecord()             ' this sets same record
            ServiceFormSetRecord = True
        End If
        mDBAccess.dbClose()
        mDBAccess = Nothing
    End Sub

    Private Sub chkOutsideService_CheckedChanged(sender As Object, e As EventArgs) Handles chkOutsideService.CheckedChanged
        If LoadingCheckBoxes Then Exit Sub
        If chkOutsideService.Checked = False Then chkOutsideService.Checked = True : Exit Sub
        LoadCheckBoxes(2, True)
    End Sub

    Private Sub chkPickupExchange_CheckedChanged(sender As Object, e As EventArgs) Handles chkPickupExchange.CheckedChanged
        If LoadingCheckBoxes Then Exit Sub
        If chkPickupExchange.Checked = False Then chkPickupExchange.Checked = True : Exit Sub
        LoadCheckBoxes(3, True)
    End Sub

    Private Sub chkOther_CheckedChanged(sender As Object, e As EventArgs) Handles chkOther.CheckedChanged
        If LoadingCheckBoxes Then Exit Sub
        If chkOther.Checked = False Then chkOther.Checked = True : Exit Sub
        LoadCheckBoxes(4, True)
    End Sub

    Private Sub mDBService_GetRecordEvent(RS As ADODB.Recordset) Handles mDBService.GetRecordEvent   ' called if record is found
        Dim Tid As String, NoteType As String

        Notes_Text.Text = ""
        AccountFound = "Y"
        AddOnAcc.lstAccounts.Items.Clear()

        Do While RS.EOF = False
            'Application.DoEvents()
            lblServiceOrderNo.Text = Trim(RS("ServiceOrderNo").Value)
            ServiceOrderNumber = Trim(RS("ServiceOrderNo").Value)
            cmdOrderParts.Enabled = True
            lblLastName.Text = Trim(RS("LastName").Value)
            UpdateTelephoneLabels("", "", "")
            lblTele.Text = DressAni(CleanAni(RS("Telephone").Value))
            MailIndex = RS("MailIndex").Value

            lblSaleNo.Text = Trim(IfNullThenNilString(RS("SaleNo").Value))
            lblSaleNo.Visible = lblSaleNo.Text <> ""
            lblSaleNoCaption.Visible = lblSaleNo.Text <> ""

            If IsNothing(RS("ServiceOnDate").Value) Then
                dteServiceDate.Value = Today
                'dteServiceDate.Value = Date.FromOADate(0)
            Else
                If RS("ServiceOnDate").Value = "NONE" Then
                    dteServiceDate.Value = Today
                    'dteServiceDate.Value = Date.FromOADate(0)
                Else
                    dteServiceDate.Value = CDate(Microsoft.VisualBasic.Left(RS("ServiceOnDate").Value, 10))
                End If
            End If
            lblClaimDate.Text = Trim(RS("DateOfClaim").Value)
            ServiceStatus = Trim(RS("Status").Value)

            SelectStatus(RS("Status").Value)
            LoadCheckBoxes(RS("QuickCheck").Value) 'this is for checkboxes

            NoteType = Trim(RS("Type").Value)
            txtItems.Text = Trim(RS("Item").Value)  ' Match this against the treeview?

            Notes_Text.Text = Trim(RS("Complaint").Value)
            Notes_New.Text = Trim(RS("StoreAction").Value)
            ' rs("Mfg")
            ' rs("InvoiceNo")

            'AddOnAcc.lstAccounts.AddItem " " & ServiceOrderNumber & "            " & lblLastName & "  " & lblTele
            'AddOnAcc.lstAccounts.Items.Add(" " & ServiceOrderNumber & "            " & lblLastName.Text & "  " & lblTele.Text)
            AddOnAcc.lstAccounts.Items.Add(" " & ServiceOrderNumber & "                 " & lblLastName.Text & "       " & lblTele.Text)
            RS.MoveNext()
        Loop

        LoadPartsOrders()
    End Sub

    Public Sub mDBAccess_SetRecordEvent(RS As ADODB.Recordset) Handles mDBAccess.SetRecordEvent    ' called to write the record
        If ServiceOrderNumber <= 0 Then
            GetServiceNo()
        End If
        RS("ServiceOrderNo").Value = Trim(ServiceOrderNumber)  ' Can't do this if it's an autonumber.
        RS("MailIndex").Value = MailIndex
        RS("LastName").Value = Trim(lblLastName.Text)
        RS("Telephone").Value = CleanAni(lblTele.Text)

        RS("SaleNo").Value = Trim(lblSaleNo.Text)
        RS("ServiceOnDate").Value = Trim(dteServiceDate.Value)
        RS("DateOfClaim").Value = Trim(lblClaimDate.Text)

        If cboStatus.SelectedIndex = -1 Then cboStatus.SelectedIndex = 0
        RS("Status").Value = Trim(cboStatus.Text)
        RS("QuickCheck").Value = GetCheckBoxValue()
        RS("Item").Value = Trim(txtItems.Text)
        RS("Complaint").Value = Trim(Notes_Text.Text)
        RS("Type").Value = "Store"                                      ' WHAT IS THIS FOR???!!
        RS("StoreAction").Value = Trim(Notes_New.Text)
        RS("Mfg").Value = Trim(Mid(txtItems.Text, 17, 16)) '###MANUFLENGTH16
        If RS("Mfg").Value = "" Then RS("Mfg").Value = " "
        RS("InvoiceNo").Value = " "
        RS("Detail").Value = ""
        RS("StopStart").Value = IIf(IsDate(dtpDelWindow0.Value), Format(dtpDelWindow0.Value, "h:mm ampm"), "")
        RS("StopEnd").Value = IIf(IsDate(dtpDelWindow1.Value), Format(dtpDelWindow1.Value, "h:mm ampm"), "")
    End Sub

    Private Sub mDBAccess_GetRecordEvent(RS As ADODB.Recordset) Handles mDBAccess.GetRecordEvent
        'This should be called by the mDBAccess component when it finds a record.
        On Error GoTo HandleErr

        lblServiceOrderNo.Text = RS("ServiceOrderNo").Value
        ServiceOrderNumber = RS("ServiceOrderNo").Value
        lblLastName.Text = RS("LastName").Value
        UpdateTelephoneLabels("", "", "")
        lblTele.Text = DressAni(CleanAni(RS("Telephone").Value))

        '  MailIndex = rs("MailIndex")

        lblSaleNo.Text = Trim(RS("SaleNo").Value)
        lblSaleNo.Visible = lblSaleNo.Text <> ""
        lblSaleNoCaption.Visible = lblSaleNo.Text <> ""

        If IsNothing(RS("ServiceOnDate").Value) Then
            dteServiceDate.Value = Today
            'dteServiceDate.Value = Date.FromOADate(0)
            ShowTimeWindowBox(False)
        Else
            'dteServiceDate.Value = Date.Parse(RS("ServiceOnDate").Value, Globalization.CultureInfo.InvariantCulture)
            dteServiceDate.Value = Date.Parse(DateFormat(RS("ServiceOnDate").Value), Globalization.CultureInfo.InvariantCulture)
            ShowTimeWindowBox(True, True)
        End If
        lblClaimDate.Text = Trim(RS("DateOfClaim").Value)

        cboStatus.Text = Trim(RS("Status").Value)  ' This may cause an error if the element's not in the list?

        LoadCheckBoxes(RS("QuickCheck").Value)
        cmdOrderParts.Enabled = True

        txtItems.Text = Trim(RS("Item").Value)
        Notes_Text.Text = Trim(RS("Complaint").Value)
        Notes_New.Text = Trim(RS("StoreAction").Value)
        '  FindItems

        LoadCustomer(RS("MailIndex").Value, False)  ' Load extended mail information

        'dtpDelWindow0.Value = RS("StopStart").Value
        'dtpDelWindow1.Value = RS("StopEnd").Value

        dtpDelWindow0.Value = DateTime.ParseExact(Trim(RS("StopStart").Value), "h:mm tt", System.Globalization.CultureInfo.InvariantCulture)
        dtpDelWindow1.Value = DateTime.ParseExact(Trim(RS("StopEnd").Value), "h:mm tt", System.Globalization.CultureInfo.InvariantCulture)
        Exit Sub

HandleErr:
        Err.Clear()
        Resume Next
    End Sub

    Private Sub mDBAccess_GetRecordNotFound() Handles mDBAccess.GetRecordNotFound
        If SearchingSOID Then
            MessageBox.Show("Unable to find that Service Order Number.", "Not Found!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        If ServiceOrderNumber > 0 Then Exit Sub  ' Don't clear if it's an error..
        ClearServiceOrder()
    End Sub

    Private Sub cmdMoveFirst_Click(sender As Object, e As EventArgs) Handles cmdMoveFirst.Click
        LoadServiceCall(0, "First")
    End Sub

    Private Sub cmdMoveLast_Click(sender As Object, e As EventArgs) Handles cmdMoveLast.Click
        LoadServiceCall(0, "Last")
    End Sub

    Private Sub cmdMoveNext_Click(sender As Object, e As EventArgs) Handles cmdMoveNext.Click
        LoadServiceCall(0, "Next")
    End Sub

    Private Sub cmdMovePrevious_Click(sender As Object, e As EventArgs) Handles cmdMovePrevious.Click
        LoadServiceCall(0, "Previous")
    End Sub

    Private Sub MoveRecord(ByVal strDirection As String)
        ' Implements the move first, last, next, previous
        ' by using strDirection
        ' and modifying the SQL created in mDBAccess_Init
        ' This affects the record(s) returned.
        LoadServiceCall(0, strDirection)
    End Sub

    Private Sub Notes_New_TextChanged(sender As Object, e As EventArgs) Handles Notes_New.TextChanged
        Dim Stamp As Boolean
        Select Case Notes_New.Tag
            Case "" : Notes_New.Tag = "EDIT1" : If Len(Notes_New.Text) = 1 Then Stamp = True
            Case "EDIT1" : Stamp = True
        End Select

        On Error Resume Next
        If Stamp Then
            Notes_New.Tag = "EDIT2"
            Notes_New.Text = Microsoft.VisualBasic.Left(Notes_New.Text, Len(Notes_New.Text) - 1) & vbCrLf & Now & ":" & vbCrLf & Microsoft.VisualBasic.Right(Notes_New.Text, 1)
            If Microsoft.VisualBasic.Left(Notes_New.Text, 2) = vbCrLf Then Notes_New.Text = Mid(Notes_New.Text, 3)
            Notes_New.SelectionStart = Len(Notes_New.Text)
        End If
    End Sub

    'Private Sub tvItemNotes_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvItemNotes.AfterSelect
    '    e.Node.Name
    'End Sub

    Private Sub chkServiceOnDate_CheckedChanged(sender As Object, e As EventArgs) Handles chkServiceOnDate.CheckedChanged
        'ShowTimeWindowBox(IsDate(dteServiceDate.Value), True)
        If chkServiceOnDate.Checked = True Then
            dteServiceDate.Enabled = True
            dtpDelWindow0.Enabled = True
            dtpDelWindow1.Enabled = True
            ShowTimeWindowBox(True, True)
        Else
            dteServiceDate.Enabled = False
            dtpDelWindow0.Enabled = False
            dtpDelWindow1.Enabled = False
            ShowTimeWindowBox(False, True)
        End If
    End Sub
End Class




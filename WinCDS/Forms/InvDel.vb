Public Class InvDel
    Public X as integer
    Dim Margin As New CGrossMargin  '+NEW 2003-01-31AA:
    Public TaxRec1 As Decimal, TaxRec2 As Decimal, MiscDisc As Decimal, Tax2Zone as integer
    Public TransDate As String, BSRowNum as integer
    Private ShowDept As Boolean, ShowVend As Boolean
    Private PollingSaleDate As Boolean

    Private Const FRM_W1 = 2520
    Private Const FRM_W2 = 5610

    Private DoDeliverAll As Boolean, ContinueDelivery As Boolean

    Public Sub ShowModal(ByRef ParentForm As Form, Optional ByVal Mdl As Boolean = True)
        If GetNextItemOrUnload() Then
            If Mdl Then
                'Show vbModal, ParentForm
                Me.ShowDialog(ParentForm)
            Else
                Show()
            End If
        End If
    End Sub

    Private Function GetNextItemOrUnload() As Boolean
        'If Not GetNextItem() Then UnloadForm : Exit Function
        If Not GetNextItem() Then
            'UnloadForm
            Me.Close()
            Exit Function
        End If
        GetNextItemOrUnload = True
        BillOSale.HiLiteRow(X)
    End Function
    Private Function GetNextItem() As Boolean ' True if found
        On Error GoTo AnError
        With Margin.DataAccess()
            ' DataAccess is already open to the Sale.
            ' All we have to do is get the next record and return true,
            ' or return false if there are no more records.
            If Not .Records_Available Then
                GetNextItem = False
            Else
                BSRowNum = BSRowNum + 1
                'If BSRowNum Mod 19 = 0 Then BillOSale.UGridIO1.GetDBGrid.FirstRow = BSRowNum \ 19
                MoveBox()
                GetNextItem = True
                X = X + 1
                With Margin
                    ' Sales tax needs to be recorded even if it was previously delivered.
                    If Trim(.Style) = "TAX1" Then TaxRec1 = TaxRec1 + .SellPrice
                    If Trim(.Style) = "TAX2" Then TaxRec2 = TaxRec2 + .SellPrice : Tax2Zone = Trim(.Quantity)
                    If Trim(.Style) = "PAYMENT" And Trim(.Quantity) = "10" Then MiscDisc = MiscDisc + .SellPrice

                    ' Don't give the option to deliver delivered items.
                    If IsNothing(.Status) Then GetNextItem = GetNextItem() : Exit Function
                    If IsDelivered(.Status) Then GetNextItem = GetNextItem() : Exit Function
                    If Microsoft.VisualBasic.Left(.Status, 1) = "x" Then GetNextItem = GetNextItem() : Exit Function
                    Style.Text = .Style

                    If Trim(.Status) = "SS" Or Trim(.Status) = "SSLAW" Or Trim(.Status) = "FND" Then
                        If .Vendor <> "" And .VendorNo <> "" Then
                            ' Select that vendor in the box..
                            ' Or better yet, don't show the boxes.
                            cboVendor.Visible = False
                            lblVendor.Visible = False
                            ShowVend = False
                        Else
                            If .Vendor = "" Then
                                'cboVendor.ListIndex = 0
                                cboVendor.SelectedIndex = 0
                            Else
                                cboVendor.Text = .Vendor
                            End If
                            cboVendor.Visible = True
                            lblVendor.Visible = True
                            ShowVend = True
                        End If

                        If .DeptNo <> "" Then
                            cboDept.Visible = False
                            lblDept.Visible = False
                            ShowDept = False
                        Else
                            lblDept.Visible = True
                            'cboDept.ListIndex = 0
                            cboDept.SelectedIndex = 0
                            cboDept.Visible = True  ' This doesn't work if the form isn't loaded already!
                            ShowDept = True
                        End If
                    Else
                        cboDept.Visible = False
                        cboVendor.Visible = False
                        lblDept.Visible = False
                        lblVendor.Visible = False
                        ShowDept = False
                        ShowVend = False
                    End If

                    If IsIn(Trim(.Status), "SS", "SO", "SOREC", "SSREC", "FND") Or IsIn(Trim(.Style), "DEL", "LAB", "NOTES", "STAIN") Then
                        CorrectPrice
                        .Cost = GetPrice(Cost.Text)
                        .ItemFreight = GetPrice(Freight.Text)
                        ' .Code = Code.Text
                        ' Code = Left(cboDept.Text, 1)
                        ' Mfg = Trim(Right(Combo1, 16))
                    End If
                End With
            End If
        End With
        CheckDeliverEnabled(True)
        Exit Function

AnError:
        MsgBox("ERROR in DeliverItems: " & Err.Description)
        Resume Next
    End Function
    Private Sub MoveBox()
        If X > 5 Then Top = 1570
    End Sub
    Private Sub CorrectPrice()
        On Error GoTo AnError

        ' S/O S/S part of form
        Cost.TabIndex = 2
        Freight.TabIndex = 3
        cmdDeliver.TabIndex = 5

        Cost.Visible = True
        'Cost.SelStart = 0
        Cost.SelectionStart = 0
        Freight.Visible = True
        lblCost.Visible = True
        lblFreight.Visible = True
        Cost.Text = CurrencyFormat(Margin.Cost)
        Freight.Text = CurrencyFormat(Margin.ItemFreight)
        Exit Sub

AnError:
        MsgBox("ERROR in DeliverItems: " & Err.Description)
        Resume Next
    End Sub
    Private Sub CheckDeliverEnabled(ByVal ResizeForm As Boolean)
        cmdDeliver.Enabled = True
        If Margin.Status <> "FND" Then  ' don't require dept/vendor for FND items (bfh20050815)
            'If cboVendor.Visible And cboVendor.ListIndex < 1 Then cmdDeliver.Enabled = False
            If cboVendor.Visible And cboVendor.SelectedIndex < 1 Then cmdDeliver.Enabled = False
            'If cboDept.Visible And cboDept.ListIndex < 1 Then cmdDeliver.Enabled = False
            If cboDept.Visible And cboDept.SelectedIndex < 1 Then cmdDeliver.Enabled = False
        End If

        If ResizeForm Then
            If Not cmdDeliver.Enabled Or IsDLS(Margin.Style) Or lblVendor.Visible Then
                Width = FRM_W2
            Else
                Width = FRM_W1
            End If
        End If
    End Sub

End Class
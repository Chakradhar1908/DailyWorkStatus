Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class SelectPrinter
    Public SmallTags As Boolean, TagSize As String
    Dim printer As New Printer
    Public sOriginalPrint As String
    Private TicketPath As String
    Private Xx As Integer, YY As Integer, Zz As Integer        ' Label finding tools..
    Private AdjustX As Integer
    Private KitTag As Boolean         ' Is this a kit?

    ' These variables are printed on the tags.
    Private List As String
    Private OnSale As String
    Private Desc As String
    Private Code As String
    Private Mfg As String
    Private mStyle As String
    Private Stock As String
    Private Comments As String
    Private SKU As String
    Private Landed As String
    Private PictureFile As String
    Private PrintPicture As Integer
    Private HidePricing As Boolean

    'Public PriceFontName As String, PriceFontSize As Double, OldPriceFontName As String, OldPricefontSize As Double
    ' Local Property Store
    Private mAllowRecLbl As Boolean
    Private mPrintingAllowed As Boolean
    Private SelectPrinterLoadFromPrintTags As Boolean

    Public Sub PrintTags(ByVal nStyle As String, ByVal nDesc As String, ByVal nLanded As String,
    ByVal nList As String, ByVal nOnSale As String, ByVal nDeptNo As String, ByVal nCode As String,
    ByVal nVendor As String, ByVal nStock As String, ByVal nComments As String,
    Optional ByVal ParentForm As Form = Nothing, Optional ByVal KitMode As Boolean = False,
    Optional ByVal nPictureFile As String = "", Optional ByVal nPrintPicture As Integer = 0,
    Optional ByVal AutoPrintTags As Boolean = False, Optional ByVal DefaultTagSize As String = "",
    Optional ByVal DefaultTicketPath As String = "", Optional ByVal nHidePricing As Boolean = False)
        ' This function gets the data from outside and shows the form..

        Dim X As String
        X = printer.DeviceName

        SelectPrinter_Load(Me, New EventArgs)
        SelectPrinterLoadFromPrintTags = True

        If PrSel.GetSelectedPrinter Is Nothing Then
            PrintingAllowed = False
        Else
            PrintingAllowed = True
        End If

        'Load Me
        LoadTagInfo(nStyle, nDesc, nLanded, nList, nOnSale, nDeptNo, nCode, nVendor, nStock, nComments)
        HidePricing = nHidePricing


        AllowRecLabelPrinting = True
        KitTag = KitMode  ' boolean, means this is a kit.
        PictureFile = nPictureFile
        PrintPicture = nPrintPicture

        If AutoPrintTags Then
            AutoPrint(DefaultTagSize, DefaultTicketPath, ParentForm)
        ElseIf ParentForm Is Nothing Then
            'Show vbModal
            ShowDialog()
        Else
            'Show vbModal, ParentForm
            ShowDialog(ParentForm)
            If UCase(ParentForm.Name) = "EDITPO" Then
                ' Awful hack, but oh well..
                EditPO.SaveTagPrintingOptions(TagSize, TicketPath)
            End If
        End If
        If Not SmallTags Then
            SetPrinter(X)
        End If
        KitTag = False
    End Sub

    Public Function PrintSoldTags(ByVal Style As String, Optional ByVal LastName As String = "", Optional ByVal SaleNo As String = "", Optional ByVal Q As Integer = 1) As Integer
        'print dymo labels
        Dim Counter As Byte, OriginalPrint As String, InvData As New CInvRec
        Dim P As Object, SQL As String

        Dim Tx As Integer
        'If Q <= 0 Then Q = Quantity

        On Error Resume Next
        If Not InvData.Load(Style, "Style") Then
            DisposeDA(InvData)
            Exit Function
        End If

        OriginalPrint = printer.DeviceName

        For Counter = 1 To Q
            If Not SetDymoPrinter() Then  ' Yes, it's inside the loop
                MessageBox.Show("Dymo Printer Required!", "WinCDS")
                Exit Function
            End If

            printer.FontSize = 14
            printer.CurrentX = 0
            printer.CurrentY = 0

            printer.Orientation = vbPRORLandscape

            printer.FontSize = 32
            printer.FontBold = True
            printer.Print("SOLD") 'PO.AckInv
            Tx = printer.CurrentX * 1.1
            printer.FontBold = False

            printer.CurrentY = 0
            printer.FontSize = 14
            If LastName <> "" Then printer.CurrentX = Tx : printer.Print("Cust: " & LastName)
            If SaleNo <> "" Then printer.CurrentX = Tx : printer.Print("Sale: " & SaleNo)
            printer.CurrentX = Tx : printer.Print("Date: " & DateFormat(Now))

            printer.EndDoc()

            printer.Orientation = vbPRORPortrait
            If OriginalPrint <> "" Then
                If Not SetPrinter(OriginalPrint) Then
                    MessageBox.Show("Could not restore the original printer!", "Original Printer")
                End If
            End If
        Next

        DisposeDA(InvData)

        PrintSoldTags = Q
    End Function

    Public ReadOnly Property LabelNoFile() As String
        Get
            'BFH20090622
            'LabelNoFile = InventFolder & "LabelNo.Dat"
            LabelNoFile = InventFolder() & "LblNo-" & StoresSld & ".Dat"
        End Get
    End Property

    Public Property AllowRecLabelPrinting() As Boolean
        Get
            AllowRecLabelPrinting = mAllowRecLbl
        End Get
        Set(value As Boolean)
            mAllowRecLbl = Trim(value)
            cmdRecLbl.Visible = IIf(Style = "", False, value)
            lblExtraRecLbl.Visible = IIf(Style = "", False, value)
        End Set
    End Property

    Private Sub cmdCustom_Click(sender As Object, e As EventArgs) Handles cmdCustom.Click
        If cboCustomTagTemplate.SelectedIndex = 0 Then
            MessageBox.Show("Please select a custom tag layout first!", "No Template Selected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        PrintCustomTags(Style, Quantity, cboCustomTagTemplate.Text)
    End Sub

    Private Sub SelectPrinter_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '  AllowRecLabelPrinting = False
        '  PrintingAllowed = False
        If SelectPrinterLoadFromPrintTags = True Then Exit Sub
        Quantity = 1

        '  If IsFennimore() Then
        '    cmdLarge.Visible = False
        '    cmdSmall.Visible = True
        '  End If

        LoadJustifications()
        LoadCustomTags()

        sOriginalPrint = printer.DeviceName
        SetPrinter(TicketPrinter)
        PrSel.SetSelectedPrinter(TicketPrinter)

        '  PriceFont = "Arial"
        '  PriceFontSize = 15

        'SetCustomFrame Me
    End Sub

    Private Sub LoadJustifications()
        On Error Resume Next
        cboTagJustify.Items.Clear()
        cboTagJustify.Items.Add("Center")
        cboTagJustify.Items.Add("Left")
        cboTagJustify.Items.Add("Right")
        cboTagJustify.Text = "Center"
        cboTagJustify.Text = StoreSettings.TagJustify
    End Sub

    Private Sub LoadCustomTags()
        LoadCustomTagLayoutsToComboBox(cboCustomTagTemplate)
    End Sub

    Public Property Style() As String
        Get
            Style = mStyle
        End Get
        Set(value As String)
            mStyle = Trim(value)
            lblDisplayStyle.Text = value
            AllowRecLabelPrinting = AllowRecLabelPrinting ' this resets according to whether style is provided
        End Set
    End Property

    Private Sub SelectPrinter_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'QueryUnload event of vb6.0 code
        '-------------------------------
        'If UnloadMode = vbFormControlMenu Then cmdDone.Value = True  ' Done
        If e.CloseReason = CloseReason.UserClosing Then cmdDone_Click(cmdDone, New EventArgs)

        'Unload event of vb6.0 code
        '---------------------------
        'RemoveCustomFrame Me
        'TicketPrinter = printer.DeviceName
        TicketPrinter(, printer.DeviceName, "Let")
        If Not SmallTags Then
            If sOriginalPrint <> "" Then
                If Not SetPrinter(sOriginalPrint) Then
                    MessageBox.Show("Could not restore the original printer!", "Original Printer")
                End If
            End If
        End If

        TagSize = ""
    End Sub

    Private Sub PrSel_Click() Handles PrSel.Click
        PrintingAllowed = True
    End Sub

    Private Sub cmdDone_Click(sender As Object, e As EventArgs) Handles cmdDone.Click
        'Unload Me
        Me.Close()
    End Sub

    Private Sub cmdDYMO_Click(sender As Object, e As EventArgs) Handles cmdDYMO.Click
        Dim I As Integer
        On Error GoTo ErrHand
        TagSize = "DYMO"
        For I = 1 To Quantity
            If Not MakeDYMOTags() Then Exit Sub
        Next
        DoBeep()
        Exit Sub
ErrHand:
        MessageBox.Show("DYMO Printer Error")
    End Sub

    Private Function MakeDYMOTags() As Boolean
        Dim OriginalPrint As String
        OriginalPrint = printer.DeviceName

        If Not SetDymoPrinter() Then  ' Yes, it's inside the loop
            MessageBox.Show("Dymo Printer Required!", "WinCDS")
            Exit Function
        End If

        printer.FontSize = 14
        printer.CurrentX = 0
        printer.CurrentY = 0

        printer.Orientation = vbPRORLandscape

        If StoreSettings.bShowRegularPrice Then  'list price
            ' If Not isveranda() Then
            printer.FontSize = 16
            printer.CurrentX = 100

            printer.Print("Reg: " & FormatCurrency(List))
        End If

        printer.FontBold = True

        'If Not isveranda() Then
        printer.CurrentX = 50
        printer.Print("On Sale:  ")
        printer.FontSize = 24
        printer.FontBold = True
        printer.Print(FormatCurrency(OnSale))
        printer.FontBold = False
        'End If

        printer.FontSize = 12
        If IsStudioD() Then
            printer.Print(Microsoft.VisualBasic.Left(SKU, 47))
        Else
            printer.Print(Microsoft.VisualBasic.Left(Desc, 47))
        End If

        If Not StoreSettings.bShowManufacturer Then
            printer.Print("Code: " & Code)
        Else
            printer.Print("Mfg.:" & Mfg)
        End If

        If StoreSettings.bCostInCode And Trim(Landed) <> "" Then
            printer.Print(New String(" ", 25))
            PrintCostCode(Landed, TagSize)
        Else
            printer.Print()
        End If

        If StoreSettings.bStyleNoInCode Then
            printer.Print("Style: " & ConvertCostToCode(Style))
        Else
            printer.Print("Style: " & Style)
        End If
        If StoreSettings.bShowAvailableStock Then
            printer.Print("Stock: " & Stock)
        End If
        printer.Print("Notes: " & Comments)

        If StoreSettings.bPrintBarCode Then
            printer.FontName = FONT_C39_WIDE
            printer.CurrentX = 500
            printer.FontSize = 20
            printer.Print(PrepareBarcode(Style))
            printer.FontName = "Arial" 'cdFont.FontName ' "Arial"
        End If

        printer.EndDoc()

        printer.Orientation = vbPRORPortrait
        If OriginalPrint <> "" Then
            If Not SetPrinter(OriginalPrint) Then
                MessageBox.Show("Could not restore the original printer!", "Original Printer")
            End If
        End If
        MakeDYMOTags = True
    End Function

    Private Sub cmdRecLbl_Click(sender As Object, e As EventArgs) Handles cmdRecLbl.Click
        PrintRecLabels(Style)
    End Sub

    Public Function PrintRecLabels(ByVal Style As String, Optional ByVal Q As Integer = 0) As Integer
        'print dymo labels
        Dim Counter As Byte, OriginalPrint As String, InvData As New CInvRec
        Dim P As Object, SQL As String
        If Q <= 0 Then Q = Quantity

        On Error Resume Next
        If Not InvData.Load(Style, "Style") Then
            DisposeDA(InvData)
            Exit Function
        End If

        OriginalPrint = printer.DeviceName

        For Counter = 1 To Q
            If Not SetDymoPrinter() Then  ' Yes, it's inside the loop
                MessageBox.Show("Dymo Printer Required!", "WinCDS")
                Exit Function
            End If

            printer.FontSize = 14
            printer.CurrentX = 0
            printer.CurrentY = 0

            printer.Orientation = vbPRORLandscape

            printer.FontSize = 14
            printer.Print("Cust: ")
            printer.FontSize = 20
            printer.FontBold = True
            printer.Print() 'Trim(PO.name)

            printer.FontSize = 14
            printer.Print("Sale: ")
            printer.Print() 'PO.SaleNo;
            printer.FontBold = False

            printer.Print("Inv/Ack: ") 'PO.AckInv

            printer.FontBold = True
            printer.Print("Style: ")
            printer.Print(InvData.Style)
            printer.FontBold = False

            printer.Print("Mfg: ")
            printer.Print(InvData.Vendor)

            printer.Print(Microsoft.VisualBasic.Left(InvData.Desc, 47))

            printer.Print("Date: ")
            printer.Print(DateFormat(Now))

            If StoreSettings.bPrintBarCode Then
                printer.FontName = FONT_C39_WIDE
                printer.CurrentX = 500
                printer.FontSize = 20
                printer.Print(PrepareBarcode(InvData.Style))
                printer.FontName = "Arial" 'cdFont.FontName ' "Arial"
            End If

            printer.EndDoc()

            printer.Orientation = vbPRORPortrait
            If OriginalPrint <> "" Then
                If Not SetPrinter(OriginalPrint) Then
                    MessageBox.Show("Could not restore the original printer!", "Original Printer")
                End If
            End If
        Next

        DisposeDA(InvData)
    End Function

    Private Sub txtCopies_TextChanged(sender As Object, e As EventArgs) Handles txtCopies.TextChanged
        Quantity = Val(txtCopies.Text)
    End Sub

    Public Property PrintingAllowed() As Boolean
        Get
            PrintingAllowed = mPrintingAllowed
        End Get
        Set(value As Boolean)
            mPrintingAllowed = value
            cmdCustom.Enabled = value
            lblCustomCaption.Enabled = value
            cboCustomTagTemplate.Enabled = value
            cmdLarge.Enabled = value
            cmdMedium.Enabled = value
            cmdSmall.Enabled = IIf(KitTag, False, value)
            cmdDYMO.Enabled = IIf(Style = "", False, True) 'nValue)
            cmdRecLbl.Enabled = IIf(Style = "", False, True) 'nValue
        End Set
    End Property

    Private Sub updQuantity_Change(sender As Object, e As EventArgs) Handles updQuantity.Change
        txtCopies.Text = updQuantity.Value
    End Sub

    Private Sub cmdSmall_Click(sender As Object, e As EventArgs) Handles cmdSmall.Click
        Dim II As Integer
        If Not SmallTags Then
            If Not SelectPrinter() Then Exit Sub
        End If
        SmallTags = True

        For II = 1 To Quantity
            PrintSmallTicket()
        Next
        DoBeep()
        printer.EndDoc()
    End Sub

    Private Sub PrintSmallTicket()
        On Error GoTo ErrHand
        TagSize = "SMALL"

        '  If IsFennimore() Then
        '    MakeFennimoreSmall
        '  ElseIf IsLegacy() Or IsHilliard() Then
        '    Legacy
        '  Else
        MakeSmallTags()
        '  End If
        Exit Sub
ErrHand:
        MessageBox.Show("Printer Error (SelectPrinter::PrintSmallTicket)")
    End Sub

    Private Sub MakeSmallTags()
        Dim CY As Integer
        Dim S As String

        FindLabel()
        printer.FontName = "Arial"
        printer.FontSize = 9
        If Val(Stock) < 0 Then Stock = 0

        'printer.CurrentX = Xx
        printer.CurrentX = 100
        printer.CurrentY = 500
        If IsStudioD() Then
            printer.Print(Microsoft.VisualBasic.Left(SKU, 29))
        Else
            printer.Print(Microsoft.VisualBasic.Left(Desc, 29))
        End If

        'printer.CurrentX = Xx
        printer.CurrentX = 100
        CY = printer.CurrentY
        If Not StoreSettings.bShowManufacturer Then
            printer.Print("Code:" & Code)
        Else
            printer.Print("Mfg:" & Mfg)
        End If

        printer.CurrentX = 1250
        printer.CurrentY = CY
        If StoreSettings.bShowAvailableStock Then   'show stock
            printer.Print(SPC(8), "Stock:" & Stock)
        Else
            printer.Print("")
        End If

        ''printer.CurrentX = Xx
        printer.CurrentX = 100
        If Not IsMiltons() Then
            If StoreSettings.bStyleNoInCode Then
                'PrintCostCode(Trim(Style), TagSize)
                CY = printer.CurrentY
                printer.Print("Style: ")
                S = ConvertCostToCode(Trim(Style))
                printer.FontBold = True
                printer.FontSize = 11
                printer.CurrentX = printer.CurrentX + printer.TextWidth("Style")
                printer.CurrentY = CY
                printer.Print(S)
                printer.FontBold = False
            Else
                printer.Print("Style: ", Style)
            End If
        End If

        Dim pY As Integer
        pY = printer.CurrentY
        If StoreSettings.bShowRegularPrice And Not IsClassicInteriors Then
            'printer.CurrentX = Xx
            printer.CurrentX = 1700
            printer.CurrentY = CY
            printer.FontSize = 8
            'printer.Print(SPC(22), "List: $", List)
            printer.Print("List: $" & List)
        End If

        'If StoreSettings.bShowRegularPrice And Not StoreSettings.bStyleNoInCode Then
        printer.CurrentY = pY
        printer.Print("")

        printer.CurrentX = Xx
        printer.FontSize = 9

        '  Printer.Print ""
        If Not HidePricing Then
            printer.CurrentX = 100
            printer.CurrentY = printer.CurrentY - 200
            CY = printer.CurrentY
            printer.Print(" Sale:")
            printer.FontSize = 12
            printer.CurrentX = printer.CurrentX + printer.TextWidth(" Sale:") + 90
            printer.CurrentY = CY
            If IsVeranda() Then
                printer.Print(FormatCurrency(OnSale), "   ")
            Else
                printer.Print(FormatCurrency(OnSale), "   ")
            End If

            If StoreSettings.bCostInCode Then
                'adds cost in code
                'PrintCostCode(Landed, TagSize)
                S = ConvertCostToCode(Trim(Landed))
                printer.FontSize = 11
                printer.CurrentX = printer.CurrentX + printer.TextWidth(OnSale) + 100
                printer.CurrentY = CY
                printer.Print(S)
            Else
                printer.Print()
            End If
        End If

        If StoreSettings.bPrintBarCode Then
            MainMenu.rtbn.Visible = True
            MainMenu.rtbn.SetBarcodeMed(Trim(Style))
            'MainMenu.rtbn.FilePrint(Xx + 100, printer.CurrentY, 3600, , True, False) '+ 250  ' was 450
            printer.CurrentX = 200
            printer.CurrentY = printer.CurrentY - 200
            ''printer.FontName = "Code39HalfInch-Regular"
            printer.FontName = FONT_C39_WIDE
            printer.FontSize = 15
            printer.FontBold = True
            printer.Print(MainMenu.rtbn.mRichTextBox.Text)
            MainMenu.rtbn.Visible = False
            'printer.FontName = "Arial"
        End If
        'printer.FontSize = 9
    End Sub

    Private Sub FindLabel()
        On Error GoTo HandleErr
        Zz = Val(ReadFile(LabelNoFile))
        Zz = Zz + 1

        ' Valid Zz are 1-30.
        If Zz = 31 Then
            Zz = 1
            printer.EndDoc()
            printer.FontName = "Arial"
            printer.FontSize = 9
        End If

        WriteFile(LabelNoFile, Zz, True)

        Select Case Zz Mod 3
            Case 1 : Xx = 100
            Case 2 : Xx = 3900 + 200
            Case 0 : Xx = 7900 + 180
        End Select

        If (Zz - 1) \ 3 >= 0 And (Zz - 1) \ 3 <= 9 Then
            YY = 525 + ((Zz - 1) \ 3) * 1440
        Else
            YY = 525
            Zz = Zz Mod 3
        End If
        Exit Sub
HandleErr:
        Zz = 0    ' create 1st label
        Resume Next
    End Sub

    Private Sub cmdMedium_Click(sender As Object, e As EventArgs) Handles cmdMedium.Click
        Dim II As Integer
        If Not SelectPrinter() Then Exit Sub

        If KitTag Then
            MakeMedPackage()
            Exit Sub
        End If

        For II = 1 To Quantity
            PrintMediumTicket()
        Next
        DoBeep()
    End Sub

    Private Sub PrintMediumTicket()
        On Error GoTo ErrHand
        TagSize = "MED"

        If IsFennimore() Then
            MakeFennimore()
        ElseIf IsWaters() Then
            MakeBassettMed()
        ElseIf IsLucas() Then
            MakeLucasMed()
        ElseIf IsMiltons Then
            MakeMiltonMed()
        ElseIf IsFurnOne Then
            MakeFurnitureOneMed(GetPrice(OnSale), Desc, Mfg, Style, Code)
        Else
            MakeTicketMed()
        End If
        Exit Sub
ErrHand:
        MessageBox.Show("Printer Error (SelectPrinter::PrintMediumTicket)")
    End Sub

    Private Sub MakeTicketMed()
        Dim CY As Integer
        Dim S As String
        AdjustX = JustificationAdjustment()
        TagSize = "MED"

        printer.FontName = "Arial"
        printer.FontBold = True

        If Not HidePricing Then
            If StoreSettings.bShowRegularPrice Then
                printer.CurrentX = 4900 + AdjustX
                printer.CurrentY = 4400
                printer.FontSize = 8
                printer.Print("")
                printer.CurrentX = 4000 + AdjustX
                printer.CurrentY = 4200

                If IsUFO() Then
                    printer.FontSize = 20
                    'printer.Print(Format(Int(GetPrice(List)), "$###,##0"))
                    printer.Print(Format(Convert.ToDecimal(GetPrice(List)), "$###,##0"))
                    printer.FontSize = 10
                    'printer.Print(Microsoft.VisualBasic.Right(Format(List, "0.00"), 2))
                    printer.Print(Microsoft.VisualBasic.Right(Format(Convert.ToDecimal(List), "0.00"), 2))
                    printer.FontSize = 20
                    printer.Print() ' Newline at size=40, or it overprints.
                Else
                    printer.FontSize = 20
                    'printer.Print(Format(List, "$###,##0.00"))
                    printer.Print(Format(Convert.ToDecimal(List), "$###,##0.00"))
                End If
            End If

            printer.FontSize = 10
            printer.CurrentX = 4000 + AdjustX
            printer.CurrentY = 4600   '4800

            If IsUFO() Then
                printer.FontSize = 40
                'printer.Print(Format(Int(GetPrice(OnSale)), "$###,##0"))
                printer.Print(Format(Convert.ToDecimal(GetPrice(OnSale)), "$###,##0"))
                printer.FontSize = 20
                'printer.Print(Microsoft.VisualBasic.Right(Format(OnSale, "0.00"), 2))
                printer.Print(Microsoft.VisualBasic.Right(Format(Convert.ToDecimal(OnSale), "0.00"), 2))
                printer.FontSize = 40
                printer.Print() ' Newline at size=40, or it overprints.
            Else
                printer.FontSize = 40
                printer.Print(Format(Convert.ToDecimal(OnSale), "$###,##0.00"))
            End If
        End If

        printer.FontSize = 10
        printer.CurrentX = 4000 + AdjustX

        If IsStudioD() Then
            printer.Print(" ", Microsoft.VisualBasic.Left(SKU, 38))
        Else
            If Len(Desc) > 46 Then printer.FontSize = 8
            printer.CurrentX = 4000 + AdjustX
            'printer.Print(" ", Microsoft.VisualBasic.Left(Desc, 46))
            printer.Print(Microsoft.VisualBasic.Left(Desc, 46))
            If Len(Desc) > 46 Then
                printer.CurrentX = 4000 + AdjustX
                printer.Print(" ", Mid(Desc, 47, 46))
            End If
            If Len(Desc) > 92 Then
                printer.CurrentX = 4000 + AdjustX
                printer.Print(" ", Mid(Desc, 93, 46))
            End If
            printer.FontSize = 10
        End If

        If Not StoreSettings.bShowManufacturer Then
            printer.CurrentX = 4000 + AdjustX
            printer.FontBold = True
            printer.Print("Code: " & Code)
            'printer.FontBold = True
            'printer.CurrentY = CY
            'printer.Print(Code)
            'Printer.FontBold = True
        Else
            printer.CurrentX = 4000 + AdjustX
            printer.FontBold = True
            printer.Print("Mfg: " & Mfg)
            'printer.FontBold = True
            'printer.Print(Mfg)
            'Printer.FontBold = True
        End If

        'Printer.CurrentX = 5500 + AdjustX
        'Printer.CurrentY = 6250

        printer.CurrentX = 4000 + AdjustX
        printer.CurrentY = printer.CurrentY + 40
        'Printer.CurrentY = 6250
        If StoreSettings.bStyleNoInCode Then
            S = ConvertCostToCode(Style)
            printer.FontBold = True
            CY = printer.CurrentY
            printer.Print("Style: ")
            printer.FontSize = 12
            printer.CurrentX = 4000 + AdjustX + printer.TextWidth("Style")
            printer.CurrentY = CY
            printer.Print(S)
            'printer.FontBold = True
            'PrintCostCode(Trim(Style), TagSize)  'style No. coded
            ' Printer.FontBold = True
        Else
            printer.FontBold = True
            printer.Print("Style: " & Style)
            ' Printer.FontBold = False
            'printer.FontBold = True
            'printer.Print(Style)
        End If

        If StoreSettings.bShowAvailableStock Then  'show stock
            printer.FontSize = 10
            printer.CurrentX = 4000 + AdjustX
            printer.CurrentY = printer.CurrentY + 40
            'printer.CurrentY = 6250   '6600
            printer.FontBold = True
            printer.Print("Stock: " & Stock)
            'printer.FontBold = True
            'printer.Print(Stock)
            ' Printer.FontBold = True
        End If

        printer.FontSize = 10
        printer.FontBold = True

        'jk next 2 lines old way.  Line wrap not correct
        printer.CurrentX = 4000 + AdjustX
        printer.CurrentY = printer.CurrentY + 40
        printer.Print(Comments)
        ' PrintInBox Printer, WrapLongText(Comments, 46), 4000 + AdjustX, Printer.CurrentY, 8000, 600

        If StoreSettings.bCostInCode Then
            'adds cost in code
            S = ConvertCostToCode(Landed)
            printer.CurrentX = 4000 + AdjustX
            printer.CurrentY = printer.CurrentY + 40
            printer.FontSize = 18
            printer.Print("  " & S)
            'printer.FontSize = 18
            'PrintCostCode(Landed, TagSize)
        End If

        If StoreSettings.bPrintBarCode Then 'bar code
            MainMenu.rtbn.SetBarcodeLarge(Trim(Style))
            'MainMenu.rtbn.FilePrint(4000 + AdjustX, 7100, 8000)
            printer.CurrentX = 4000 + AdjustX
            printer.CurrentY = 7100
            ''printer.FontName = "Code39HalfInch-Regular"
            printer.FontName = FONT_C39_HALFINCH
            printer.FontSize = 20
            printer.FontBold = False
            printer.Print(MainMenu.rtbn.mRichTextBox.Text)
            'printer.FontName = "Arial"
        End If

        'MousePointer = 0
        Me.Cursor = Cursors.Default
        printer.EndDoc()
    End Sub

    Private Sub MakeMiltonMed()
        AdjustX = JustificationAdjustment()

        TagSize = "MED"

        printer.FontName = "Arial"
        printer.FontBold = True
        printer.FontSize = 10
        printer.CurrentY = 4000
        printer.CurrentX = 5200 + AdjustX
        printer.Print(" Retail")

        If StoreSettings.bShowRegularPrice Then
            printer.CurrentX = 4900 + AdjustX
            printer.CurrentY = 4400
            printer.FontSize = 8
            printer.Print("")
            printer.CurrentY = 4200
            printer.FontSize = 20
            printer.Print(Format(List, "$###,##0.00"))
        End If

        printer.FontSize = 10
        printer.CurrentX = 5000 + AdjustX
        printer.Print(" Value Price")

        printer.FontSize = 10
        printer.CurrentX = 4300 + AdjustX
        printer.CurrentY = 4900   '4800

        printer.FontSize = 40
        printer.Print(Format(OnSale, "$###,##0.00"))

        printer.FontSize = 10
        printer.CurrentX = 4000 + AdjustX
        printer.Print(" ", Microsoft.VisualBasic.Left(Desc, 38))

        printer.CurrentX = 4000 + AdjustX
        printer.FontSize = 10
        printer.Print(Microsoft.VisualBasic.Left(Comments, 50))

        If Not StoreSettings.bShowManufacturer Then
            printer.CurrentX = 4000 + AdjustX
            printer.Print(" Code: ", Code)
        Else
            printer.CurrentX = 4000 + AdjustX
            printer.Print(" Mfg: ", Mfg)
        End If

        If StoreSettings.bPrintBarCode Then 'bar code
            MainMenu.rtbn.SetBarcodeLarge(Trim(Style))
            MainMenu.rtbn.FilePrint(4000 + AdjustX, 7100, 8000)
            printer.FontName = "Arial"
        End If

        printer.FontBold = True
        printer.FontBold = False
        'MousePointer = 0
        Me.Cursor = Cursors.Default
        printer.EndDoc()
    End Sub

    Private Sub MakeLucasMed()
        TagSize = "MED"

        printer.FontName = "Arial"
        printer.FontBold = True
        printer.FontSize = 10
        printer.CurrentX = 4900 + AdjustX
        printer.CurrentY = 6000   '4800

        If StoreSettings.bShowRegularPrice Then
            printer.CurrentX = 4700 + AdjustX
            printer.CurrentY = 5500
            printer.FontSize = 8
            printer.Print("")
            printer.CurrentY = 6000
            printer.FontSize = 30
            printer.Print(CurrencyFormat(List))
        End If

        printer.CurrentX = 4200 + AdjustX
        printer.CurrentY = 8000
        printer.FontSize = 60
        printer.Print(Format(OnSale, "$###,##0"))
        printer.FontSize = 20
        printer.Print(Microsoft.VisualBasic.Right(Format(OnSale, "0.00"), 2))
        printer.FontSize = 40
        printer.Print() ' Newline at size=40, or it overprints.

        If StoreSettings.bPrintBarCode Then  'bar code
            printer.CurrentY = 9500
            printer.CurrentX = 4500 + AdjustX '4000
            MainMenu.rtbn.SetBarcodeLarge(Trim(Style))
            MainMenu.rtbn.FilePrint()
            printer.FontName = "Arial"
        End If

        printer.FontBold = True
        printer.FontBold = False
        'MousePointer = 0
        Me.Cursor = Cursors.Default
        printer.EndDoc()

        'Back of card second side
        printer.FontSize = 16
        printer.CurrentY = 500
        printer.CurrentX = 3500 + AdjustX

        printer.Print(" ", Microsoft.VisualBasic.Left(Desc, 24))
        printer.CurrentY = printer.CurrentY + 240

        If Not StoreSettings.bShowManufacturer Then
            printer.CurrentX = 4000 + AdjustX
            printer.Print("       ", Code)
        Else
            printer.CurrentX = 4000 + AdjustX
            printer.Print("      ", Mfg)
        End If

        printer.CurrentY = printer.CurrentY + 275
        printer.CurrentX = 4000 + AdjustX
        If StoreSettings.bStyleNoInCode Then
            'style No. coded
            printer.Print("       ")
            PrintCostCode(Trim(Style), TagSize)
        Else
            printer.Print("       ", Style)
        End If

        printer.CurrentY = 5600
        printer.CurrentX = 4000 + AdjustX
        printer.Print(Trim(Microsoft.VisualBasic.Left(Comments, 50)))
    End Sub

    Private Sub MakeBassettMed()
        TagSize = "MED"

        printer.FontName = "Arial"
        printer.FontBold = True
        printer.FontSize = 10
        printer.CurrentX = 4200
        printer.CurrentY = 500
        printer.FontSize = 12
        printer.Print(" ", Microsoft.VisualBasic.Left(Desc, 38))

        If Not StoreSettings.bShowManufacturer Then
            printer.CurrentX = 4200
            printer.Print(" Code: ", Code)
        Else
            printer.CurrentX = 4200
            printer.Print(" Mfg: ", Mfg)
        End If

        printer.CurrentX = 4200
        If StoreSettings.bStyleNoInCode Then
            'style No. coded
            printer.Print("Style: ")
            PrintCostCode(Trim(Style), TagSize)
        Else
            printer.Print("Style: ", Microsoft.VisualBasic.Left(Style, 27))
        End If

        If StoreSettings.bShowAvailableStock Then  'show stock
            printer.FontSize = 10
            printer.CurrentX = 4200
            printer.Print("Stock: ", Stock)
        End If

        printer.CurrentX = 4200
        printer.CurrentY = 2800  '6900
        printer.FontSize = 10
        printer.Print("Notes: ", Microsoft.VisualBasic.Left(Comments, 50))

        If StoreSettings.bCostInCode Then
            'adds cost in code
            printer.CurrentX = 4200
            printer.Print("  ")
            printer.FontSize = 18
            PrintCostCode(Landed, TagSize)
        End If

        printer.FontSize = 10
        printer.CurrentX = 3800
        printer.CurrentY = 5000
        printer.FontSize = 80
        printer.Print(CInt(OnSale))

        printer.CurrentY = 7000
        If StoreSettings.bPrintBarCode Then
            printer.CurrentX = 4800  '4400
            MainMenu.rtbn.SetBarcodeLarge(Trim(Style))
            MainMenu.rtbn.FilePrint()
            printer.FontName = "Arial"
        End If

        printer.FontBold = True
        printer.FontBold = False
        'MousePointer = 0
        Me.Cursor = Cursors.Default
        printer.EndDoc()
    End Sub

    Private Sub MakeFennimore()
        '  TagSize = "MED"
        '
        '  Printer.FontName = "Arial"
        '  Printer.FontBold = True
        '  Printer.CurrentX = 4000
        '  Printer.CurrentY = 3200 'min w/barcode here
        '  Printer.FontSize = 10
        '  Printer.FontBold = True
        '  Printer.Print Left(Desc, 35)
        '  Printer.FontBold = False
        '
        '  Printer.CurrentX = 4000
        '  Printer.Print Trim(Style)
        '
        '  Printer.CurrentX = 4000
        '  Printer.Print Mfg
        '
        '  Printer.CurrentX = 4000
        '  Dim Length as integer, Zz as integer
        '  Length = Val(Landed)
        '  Length = Val(Length)
        '  For Zz = Length To 1 Step -1
        '    Printer.Print Mid(CInt(Landed), Zz, 1);
        '  Next
        '
        '  Printer.CurrentX = 5200
        '  Printer.FontSize = 10
        '  Printer.Print "Notes: "; Left(Comments, 50)
        '
        '  Printer.CurrentX = 4200
        '  Printer.CurrentY = 4300
        '  Printer.FontBold = True
        '  Printer.FontSize = 18
        '  Printer.Print "Price ";
        '
        '  Printer.CurrentX = 4800
        '  Printer.CurrentY = 5000
        '  Printer.FontSize = 30
        '  Printer.Print "$";
        '  Printer.FontSize = 50
        '  Printer.Print Trim(CInt(OnSale))   ' This will print whole dollars, cents removed.  Is this correct?
        '  Printer.FontBold = False
        '
        '  With mainmenu.rtbn
        '    Printer.CurrentX = 4800
        '    Printer.CurrentY = 2000
        '    .SetBarcodeLarge Trim(Style)
        '    .FilePrint
        '  End With
        '
        '  Printer.FontBold = False
        '  MousePointer = 0
        '  Printer.EndDoc
    End Sub

    Private Sub MakeFennimoreSmall()
        '  FindLabel
        '
        '  Printer.FontName = "Arial"
        '  Printer.FontSize = 9
        '
        '  Printer.CurrentX = Xx
        '  Printer.CurrentY = YY
        '  Printer.Print Left(Desc, 29)
        '
        '  Printer.CurrentX = Xx
        '  Printer.Print Style; "   ";
        '
        '  Printer.FontSize = 12
        '  Printer.Print OnSale
        '
        '  'cost backwards
        '  Printer.FontSize = 9
        '  Printer.CurrentX = Xx
        '  For Zz = Len(CStr(CInt(Landed))) To 1 Step -1
        '    Printer.Print Mid(CStr(CInt(Landed)), Zz, 1);
        '  Next
        '
        '  Printer.CurrentX = Printer.CurrentX + 400
        '  Printer.FontSize = 9
        '  Printer.Print "  "; Mfg
        '  Printer.CurrentX = Xx
        '
        '  If StoreSettings.bPrintBarCode Then
        '
        '    With mainmenu.rtbn
        '      Printer.CurrentY = Printer.CurrentY + 100  ' Was 250
        '      Printer.CurrentX = Xx + 400
        '      .SetBarcodeLarge Trim(Style)
        '      .FilePrint
        '    End With
        '    Printer.FontName = "Arial"
        '  End If
        '  Printer.FontSize = 9
    End Sub

    Private Sub MakeMedPackage()
        'med kits
        printer.FontName = "Arial"
        printer.FontBold = True

        ' It looks like a medium ticket is about 6000 units wide.
        If IsFurnOne Then
            MakeFurnitureOneMed(GetPrice(OnSale), Desc, "", Style, "")
            Exit Sub
        End If

        Dim AdjustX As Integer
        AdjustX = JustificationAdjustment()

        If IsLucas() Then
            MakeLucasPackage()
            Exit Sub
        End If

        If Not HidePricing Then
            If List <> "" Then
                printer.CurrentX = 4900 + AdjustX
                printer.CurrentY = 4200
                printer.FontSize = 20
                If IsUFO() Then
                    printer.FontSize = 20
                    printer.Print(Format(List, "$###,##0"))
                    printer.FontSize = 10
                    printer.Print(Microsoft.VisualBasic.Right(Format(List, "0.00"), 2))
                    printer.FontSize = 20
                    printer.Print() ' Newline at size=40, or it overprints.
                Else
                    printer.FontSize = 20
                    printer.Print(Format(List, "$###,##0.00"))
                End If
            Else
                printer.FontSize = 20
                printer.CurrentY = 3800 + printer.TextHeight("$")
            End If

            printer.CurrentX = 4000 + AdjustX

            If IsUFO() Then
                printer.FontSize = 30
                printer.Print(Format(Int(OnSale), "$###,##0"))
                printer.FontSize = 15
                printer.Print(Microsoft.VisualBasic.Right(Format(OnSale, "0.00"), 2))
                printer.FontSize = 30
                printer.Print() ' Newline at size=40, or it overprints.
            Else
                printer.FontSize = 30
                printer.Print(Format(OnSale, "$###,##0.00"))
            End If
        End If

        printer.FontSize = 14
        printer.CurrentX = 3000 + AdjustX
        printer.Print(Trim(Microsoft.VisualBasic.Left(Desc, 35)))

        printer.FontBold = False
        printer.FontSize = 14
        printer.CurrentX = 3000 + AdjustX
        printer.Print(Style)

        'Notes
        Dim El As Object
        For Each El In Split(Comments, vbCrLf)
            printer.CurrentX = 3200 + AdjustX
            printer.Print(El)
        Next

        If StoreSettings.bCostInCode Then
            printer.CurrentX = 3000 + AdjustX
            PrintCostCode(Landed, TagSize)
        End If

        If StoreSettings.bPrintBarCode Then
            MainMenu.rtbn.SetBarcodeLarge(Trim(Style))
            MainMenu.rtbn.FilePrint(4000 + AdjustX, , 8000)
        End If
        printer.EndDoc()
    End Sub

    Private Sub MakeLucasPackage()
        ' MEDIUM TAG
        TagSize = "MED"
        printer.FontName = "Arial"
        printer.FontBold = True

        If PackagePrice.optListPrice.Checked = True Then
            printer.CurrentX = 6000 + AdjustX
            printer.CurrentY = 3500
            printer.FontSize = 30
            printer.Print("$" & CurrencyFormat(List))
        End If

        printer.CurrentY = 3050
        printer.FontSize = 22
        'Notes
        If InStr(Comments, vbCrLf) <> 0 Then 'Test for "vbCrLf"
            Dim M As Integer, N As Integer, O As Integer, P As Integer
            'M is increment variable for For loops
            'N is the quantity of "vbCrLf"s in PackagePrice.txtNotes
            'O is the current location of "vbCrLf"
            'P is the previous location of "vbCrLf"
            N = 0
            For M = 1 To Len(Comments)
                If Mid(Comments, M, 2) = vbCrLf Then
                    N = N + 1   ' Count the vbCrLfs in the textbox.
                End If
            Next
            O = -1
            For M = 1 To N + 1
                P = O + 2
                O = IIf(M = N + 1, Len(Comments) + 1, InStr(P, Comments, vbCrLf))
                printer.CurrentX = 1000 + AdjustX
                printer.Print(Mid(Comments, P, O - P))

            Next
        Else 'There is no "vbCrLf"
            printer.Print(Comments)
        End If

        printer.CurrentX = 3000 + AdjustX

        printer.FontSize = 50
        printer.CurrentX = 5000 + AdjustX
        printer.CurrentY = 5200
        printer.Print(AlignString(Format(OnSale, "$###,###.00"), 7, VBRUN.AlignConstants.vbAlignRight, False))

        printer.CurrentY = 6700

        If StoreSettings.bPrintBarCode Then
            printer.CurrentX = 6500  '4400
            MainMenu.rtbn.SetBarcodeLarge(Trim(Style))
            MainMenu.rtbn.FilePrint()
            printer.FontName = "Arial"
        End If

        printer.EndDoc()
    End Sub

    Private Function JustificationAdjustment() As Integer
        Select Case cboTagJustify.Text
            Case "Left"
                JustificationAdjustment = -2500
            Case "Right"
                JustificationAdjustment = 3000
            Case "Center"
                JustificationAdjustment = 0
        End Select
    End Function

    Private Sub MakeFurnitureOneMed(ByVal Price As Decimal, ByVal Desc As String, ByVal Vendor As String, ByVal Style As String, ByVal Code As String)
        ' Special tags for FurnitureOne.
        Dim TicketWidth As Integer
        TicketWidth = 6500

        AdjustX = CalculateBorder(TicketWidth)
        TagSize = "MED"

        Dim Dollars As String, Cents As String
        Dim MainDollarTop As Integer, MainDollarHeight As Integer

        ' temporary borders
        '  Printer.Line (AdjustX, 0)-(AdjustX + TicketWidth, 0)
        '  Printer.Line (AdjustX + TicketWidth, 0)-(AdjustX + TicketWidth, 10000)
        '  Printer.Line (AdjustX, 10000)-(AdjustX + TicketWidth, 10000)
        '  Printer.Line (AdjustX, 0)-(AdjustX, 10000)

        ' Dollars print in (1300,3700)-step(3900,3500)
        ' What's actually used?
        '  1300+(3900-MainDollarHeight)/2
        PrintInBox(printer, DollarValue(Price), AdjustX + 1300, 3700, 3900, 3500)
        MainDollarHeight = printer.TextHeight(DollarValue(Price))
        MainDollarTop = 3700 + (3500 - MainDollarHeight) / 2

        PrintInBox(printer, "$", AdjustX, MainDollarTop, 1300, MainDollarHeight / 2, , VBRUN.AlignConstants.vbAlignRight)
        printer.FontUnderline = True
        PrintInBox(printer, CentValue(Price), AdjustX + 5200, MainDollarTop, 1300, MainDollarHeight / 2, , VBRUN.AlignConstants.vbAlignLeft)
        printer.FontUnderline = False
        PrintInBox(printer, Trim(Microsoft.VisualBasic.Left(Desc, 35)), AdjustX, 7200, TicketWidth, 1300)
        PrintInBox(printer, Trim(Vendor & " " & Style), AdjustX, 8500, 4000, 500, , VBRUN.AlignConstants.vbAlignLeft)
        PrintInBox(printer, Code, AdjustX + 4000, 8500, 2500, 500, , VBRUN.AlignConstants.vbAlignRight)
        printer.EndDoc()
    End Sub

    Private Function CalculateBorder(ByVal PageWidth As Integer) As Integer
        Select Case cboTagJustify.Text
            Case "Left"
                CalculateBorder = 0
            Case "Right"
                CalculateBorder = printer.ScaleWidth - PageWidth
            Case "Center"
                CalculateBorder = (printer.ScaleWidth - PageWidth) / 2
        End Select
    End Function

    Private Sub cmdLarge_Click(sender As Object, e As EventArgs) Handles cmdLarge.Click
        Dim II As Integer
        If Not SelectPrinter() Then Exit Sub

        If KitTag Then
            MakeTicketLarge(List, OnSale, Desc, "", "", Style, "", Comments, Landed, PictureFile, PrintPicture = 1, 1)
            Exit Sub
        End If

        For II = 1 To Quantity
            PrintLargeTicket()
        Next
        DoBeep()
    End Sub

    Private Sub PrintLargeTicket()
        On Error GoTo ErrHand
        TagSize = "LARGE"
        If IsWarehouseFurniture() Then
            MakeWFTicketLarge(List, OnSale, Style, , Desc, , , 1)
        ElseIf IsFurnOne Then
            MakeFurnitureOneLarge(GetPrice(OnSale), Desc, Mfg, Style, Code)
        Else
            MakeTicketLarge(List, OnSale, Desc, Code, Mfg, Style, Stock, Comments, Landed, PictureFile, PrintPicture = 1, 0)
        End If
        Exit Sub
ErrHand:
        MessageBox.Show("Printer Error (SelectPrinter::PrintMediumLarge)", "WinCDS")
    End Sub

    Private Sub MakeFurnitureOneLarge(ByVal Price As Decimal, ByVal Desc As String, ByVal Vendor As String, ByVal Style As String, ByVal Code As String)
        ' Special tags for FurnitureOne.
        Dim TicketWidth As Integer
        TicketWidth = printer.ScaleWidth

        AdjustX = 0
        TagSize = "LARGE"

        PrintInBox(printer, Trim(Microsoft.VisualBasic.Left(Desc, 35)), AdjustX, 7000, TicketWidth, 1500)
        PrintInBox(printer, "$" & CurrencyFormat(Price), AdjustX, 8500, TicketWidth, 3500)
        PrintInBox(printer, Trim(Vendor & " " & Style & " " & Code), AdjustX, 12250, TicketWidth, 500)
        printer.EndDoc()
    End Sub

    Private Sub MakeWFTicketLarge(ByVal ListPrice As Decimal, ByVal SalePrice As Decimal, ByVal Style As String, Optional ByVal Ex1 As String = "", Optional ByVal Ex2 As String = "", Optional ByVal Ex3 As String = "", Optional ByVal Ex4 As String = "", Optional ByVal Opt As Integer = 0)
        Dim W As Integer, border As VBRUN.BorderStyleConstants

        border = 0 'vbBSSolid
        SelectBarcodeFont(1, 30)
        W = printer.TextWidth(Style)
        PrintInBox(printer, PrepareBarcode(Style), (printer.ScaleWidth - W) / 2, 9300, W, 1000, 30, , , border)

        printer.FontName = "Arial"
        '  PrintInBox Printer, Style, 3500, 2780, 3000, 800, , , , Border
        PrintInBox(printer, Style, 3800, 2780, 2800, 800, , , , border)
        PrintInBox(printer, "1", 6600, 2780, 1500, 800, , , , border)
        PrintInBox(printer, Format(ListPrice, "###,##0.00"), 6200, 7350, 3000, 660, , , , border)
        PrintInBox(printer, Format(OnSale, "###,##0.00"), 6200, 8050, 3300, 1000, , , , border)

        If Len(Trim(Ex1)) > 0 Then PrintInBox(printer, Ex1, 5100, 3830, 4500, 710, , VBRUN.AlignConstants.vbAlignLeft, , border)
        If Len(Trim(Ex2)) > 0 Then
            If Opt = 1 Then
                PrintInBox(printer, Ex2, 2100, 4540, 7500, 710, , VBRUN.AlignConstants.vbAlignLeft, , border)
            Else
                PrintInBox(printer, Ex2, 5100, 4540, 4500, 710, , VBRUN.AlignConstants.vbAlignLeft, , border)
            End If
        End If
        If Len(Trim(Ex3)) > 0 Then PrintInBox(printer, Ex3, 5100, 5250, 4500, 710, , VBRUN.AlignConstants.vbAlignLeft, , border)
        If Len(Trim(Ex4)) > 0 Then PrintInBox(printer, Ex4, 5100, 5960, 4500, 710, , VBRUN.AlignConstants.vbAlignLeft, , border)

        printer.EndDoc()
    End Sub

    Private Sub MakeTicketLarge(
    List As String, OnSale As String, Desc As String, Code As String,
    Mfg As String, Style As String, Stock As String, Comments As String, Landed As String,
    PictureFile As String, PrintPicture As Boolean, BarcodeLocation As Integer)
        Dim AdjustX As Integer, TicketLeft As Integer, TicketRight As Integer, TicketCenter As Integer, TicketWidth As Integer
        Dim Box1Top As Integer, Box1Bottom As Integer
        Dim Box2Top As Integer, Box2Bottom As Integer
        Dim Box3Top As Integer, Box3Bottom As Integer
        Dim Box4Top As Integer, Box4Bottom As Integer
        Dim CY As Integer

        AdjustX = JustificationAdjustment()

        ' Set up the printing areas.
        If AdjustX > 0 Then
            ' Right-aligned
            TicketLeft = 4700
            TicketRight = 11200
        ElseIf AdjustX < 0 Then
            ' Left-aligned
            TicketLeft = 200
            TicketRight = 6700
        Else
            ' Centered
            TicketLeft = 2500
            TicketRight = 9000
        End If
        TicketWidth = TicketRight - TicketLeft
        TicketCenter = TicketLeft + TicketWidth / 2

        Box1Top = 200
        Box1Bottom = 4400
        Box2Top = 5400
        Box2Bottom = 6000
        Box3Top = 6500
        Box3Bottom = 8400
        Box4Top = 9000
        Box4Bottom = 11300

        If IsViking() Then        ' adj for smaller tags
            Box3Top = 6000
            Box3Bottom = 8200
            Box4Top = 8500
            Box4Bottom = 11300 '?
        End If


        ' And print the ticket..
        TagSize = "LARGE"
        printer.FontName = "Arial"
        On Error GoTo HandleErr

        If PrintPicture Then
            ' Print the item's picture or the store logo.
            printer.CurrentY = Box1Top + 300
            printer.FontSize = 20
            printer.FontBold = True
            PrintToPosition(printer, Trim(StoreSettings.Name), TicketCenter - printer.TextWidth(Trim(StoreSettings.Name)) / 2, VBRUN.AlignConstants.vbAlignLeft, True)

            printer.FontBold = False
            On Error Resume Next
            'pic.Picture = LoadPictureStd("")
            pic.Image = LoadPictureStd("")
            'pic.Picture = LoadPictureStd(PXFolder() & PictureFile)
            pic.Image = LoadPictureStd(PXFolder() & PictureFile)
            'pic.Picture = LoadPictureStd(PictureFile)
            pic.Image = LoadPictureStd(PictureFile)
            '      If .Picture = 0 Then .Picture = StoreLogoPicture
            'If pic.Picture <> 0 Then printer.PaintPicture pic.Picture, TicketLeft, printer.CurrentY, TicketWidth, Box1Bottom - printer.CurrentY
            If pic.Image IsNot Nothing Then printer.PaintPicture(pic.Image, TicketLeft, printer.CurrentY, TicketWidth, Box1Bottom - printer.CurrentY)
            Err.Clear()

            If Not HidePricing Then
                ' Print regular price in box 2, with a label.
                printer.FontBold = True
                printer.FontSize = 18
                printer.CurrentY = Box2Top + (Box2Bottom - Box2Top - printer.TextHeight("Reg:")) / 2
                printer.CurrentX = TicketLeft + 800
                printer.Print("Reg:")

                ' This price isn't formatted like the others!  It should be.
                ' It's also not checked against frmSetup here.
                If StoreSettings.bShowRegularPrice Then  'list price
                    PrintPrice(GetPrice(List), IIf(IsUFO(), 1, 0),
                    printer.CurrentX + 200, Box2Top, TicketWidth - (printer.CurrentX + 200 - TicketLeft), Box2Bottom - Box2Top, 28)
                End If
            End If
        Else
            ' Print regular price in box 2, with no label.
            If Not HidePricing Then
                printer.FontBold = True
                If StoreSettings.bShowRegularPrice Then  'list price
                    If List <> "" Then
                        PrintPrice(GetPrice(List), IIf(IsUFO(), 1, 0), TicketLeft, Box2Top, TicketWidth, Box2Bottom - Box2Top, 28)
                    End If
                End If
            End If
        End If

        ' This is a guess at the box size: 5200x2000
        ' Print at current x, y.  Max font size is 75.

        ' Print sale price in box 3.
        Dim Box3ImprovTop As Integer
        If StoreSettings.bPrintBarCode Then
            ' Leave room for the bar code
            printer.FontName = FONT_C39_HALFINCH
            printer.FontSize = 20
            Box3ImprovTop = Box3Top + printer.TextHeight("*") + 50
        Else
            Box3ImprovTop = Box3Top
        End If

        'printer.Font = "Arial"
        printer.FontName = "Arial"
        printer.FontBold = True
        PrintPrice(GetPrice(OnSale), IIf(IsUFO(), 1, 0), TicketLeft, Box3ImprovTop, TicketWidth, Box3Bottom - Box3ImprovTop, 75)

        ' Print comments, etc in box 4.
        printer.FontSize = 14
        printer.CurrentY = Box4Top
        Dim Col2 As Integer, Col1End As Integer
        Col2 = TicketLeft + 1000 'Printer.TextWidth("Stock: ") + TicketLeft + 200
        Col1End = Col2 - 100

        If IsStudioD() Then
            PrintToPosition(printer, Microsoft.VisualBasic.Left(SKU, 38), TicketLeft + 200, VBRUN.AlignConstants.vbAlignLeft, True)
        Else
            If Len(Desc) > 46 Then printer.FontSize = 12
            PrintToPosition(printer, Microsoft.VisualBasic.Left(Desc, 46), TicketLeft + 200, VBRUN.AlignConstants.vbAlignLeft, True)
            If Len(Desc) > 46 Then
                PrintToPosition(printer, Mid(Desc, 47, 46), TicketLeft + 200, VBRUN.AlignConstants.vbAlignLeft, True)
            End If
            If Len(Desc) > 92 Then
                PrintToPosition(printer, Mid(Desc, 93, 46), TicketLeft + 200, VBRUN.AlignConstants.vbAlignLeft, True)
            End If
            printer.FontSize = 14
        End If

        If Not StoreSettings.bShowManufacturer Then
            If Code <> "" Then
                CY = printer.CurrentY
                'PrintToPosition(printer, "Code:", Col1End, VBRUN.AlignConstants.vbAlignRight, False)
                PrintToPosition2(printer, "Code:", Col1End, VBRUN.AlignConstants.vbAlignRight, False, CY)
                'PrintToPosition(printer, Code, Col2, VBRUN.AlignConstants.vbAlignLeft, True)
                PrintToPosition2(printer, Code, Col2, VBRUN.AlignConstants.vbAlignLeft, True, CY)
            End If
        ElseIf Mfg <> "" Then
            CY = printer.CurrentY
            'PrintToPosition(printer, "Mfg:", Col1End, VBRUN.AlignConstants.vbAlignRight, False)
            'PrintToPosition(printer, Mfg, Col2, VBRUN.AlignConstants.vbAlignLeft, True)
            PrintToPosition2(printer, "Mfg:", Col1End, VBRUN.AlignConstants.vbAlignRight, False, CY)
            PrintToPosition2(printer, Mfg, Col2, VBRUN.AlignConstants.vbAlignLeft, True, CY)
        End If

        If Style <> "" Then
            If StoreSettings.bStyleNoInCode Then
                'style No. coded
                CY = printer.CurrentY
                'PrintToPosition(printer, "Style:", Col1End, VBRUN.AlignConstants.vbAlignRight, False)
                PrintToPosition2(printer, "Style:", Col1End, VBRUN.AlignConstants.vbAlignRight, False, CY)
                'PrintCostCode(Style, TagSize, Col2)
                'PrintCostCode(Style, TagSize, Col2, CY)
                Dim S As String
                S = ConvertCostToCode(Style)
                printer.CurrentX = Col2
                printer.CurrentY = CY
                printer.FontSize = 16
                printer.Print(S)
                printer.FontSize = 14
            Else
                CY = printer.CurrentY
                'PrintToPosition(printer, "Style:", Col1End, VBRUN.AlignConstants.vbAlignRight, False)
                'PrintToPosition(printer, Style, Col2, VBRUN.AlignConstants.vbAlignLeft, True)
                PrintToPosition2(printer, "Style:", Col1End, VBRUN.AlignConstants.vbAlignRight, False, CY)
                PrintToPosition2(printer, Style, Col2, VBRUN.AlignConstants.vbAlignLeft, True, CY)
            End If
        End If

        If StoreSettings.bShowAvailableStock And Stock <> "" Then 'show stock
            CY = printer.CurrentY
            'PrintToPosition(printer, "Stock:", Col1End, VBRUN.AlignConstants.vbAlignRight, False)
            PrintToPosition2(printer, "Stock:", Col1End, VBRUN.AlignConstants.vbAlignRight, False, CY)
            'PrintToPosition(printer, Stock, Col2, VBRUN.AlignConstants.vbAlignLeft, True)
            PrintToPosition2(printer, Stock, Col2, VBRUN.AlignConstants.vbAlignLeft, True, CY)
        End If

        If Comments <> "" Then
            PrintToPosition(printer, "Notes:", Col1End, VBRUN.AlignConstants.vbAlignRight, False)
            If IsUFO() Then
                printer.FontSize = 18
            End If
            PrintInBox(printer, WrapLongText(Comments, 60), Col2, printer.CurrentY, 5500, Box4Bottom - 100 - printer.CurrentY, printer.FontSize, VBRUN.AlignConstants.vbAlignLeft, VBRUN.AlignConstants.vbAlignTop)
            printer.FontSize = 14
        End If

        If StoreSettings.bCostInCode And Trim(Landed) <> "" Then
            'adds cost in code
            PrintCostCode(Landed, TagSize, TicketLeft + 200)
        End If

        If StoreSettings.bPrintBarCode Then
            MainMenu.rtbn.SetBarcodeLarge(Trim(Style))
            'MainMenu.rtbn.FilePrint(TicketLeft + 900, Box3Top + 25) ' This sometimes overwrites the price a bit..
            printer.CurrentX = TicketLeft + 900
            printer.CurrentY = Box3Top + 25
            'printer.FontName = "Code39HalfInch-Regular"
            printer.FontName = FONT_C39_HALFINCH
            printer.FontSize = 20
            printer.FontBold = False
            printer.Print(MainMenu.rtbn.mRichTextBox.Text)
        End If

        printer.FontBold = False
        'MousePointer = 0
        Me.Cursor = Cursors.Default
        printer.EndDoc()
        Exit Sub
HandleErr:

        MessageBox.Show("Error #: " & Err.Number & ": " & Err.Description, "WinCDS")
        Resume Next
    End Sub

    Private Sub PrintPrice(ByVal Amount As Decimal, ByVal Style As Integer, ByVal ALeft As Integer, ByVal ATop As Integer, ByVal AWidth As Integer, ByVal AHeight As Integer, ByVal MaxFont As Single)
        ' This is only used to print fancy prices in display box areas, always centered.
        ' We should always be able to supply top, left, maxwidth, maxheight, maxfontsize.
        Dim Dollars As String

        If Style = 1 Then
            ' Print cents up to half the fontsize of dollars.
            ' The whole thing still has to fit in the box..
            Dim Cents As String
            Dim DollarFontSize As Single, DollarWidth As Single
            Dim CentsFontSize As Single

            Dollars = Format(Int(Amount), "$##,##0")
            Cents = Microsoft.VisualBasic.Right(Format(Amount, "0.00"), 2)
            DollarFontSize = BestPrinterFontFit(Dollars, CLng(MaxFont), AWidth * 0.9, AHeight)  ' Save at least 10% of the area for cents.
            printer.FontSize = DollarFontSize
            DollarWidth = printer.TextWidth(Dollars)
            CentsFontSize = BestPrinterFontFit(Cents, DollarFontSize / 2, AWidth - DollarWidth, printer.TextHeight(Dollars) / 2)
            printer.FontSize = CentsFontSize
            DollarWidth = DollarWidth + printer.TextWidth(Cents)

            printer.FontSize = DollarFontSize
            ' I want to use AWidth here, but it's not right.  It has to be right.
            printer.CurrentX = ALeft + (AWidth - DollarWidth) / 2 ' Center price, offset to allow cents to follow.
            printer.CurrentY = ATop + (AHeight - printer.TextHeight(Dollars)) / 2
            printer.Print(Dollars)

            ' No printer position adjustment is required.
            printer.FontSize = CentsFontSize
            printer.Print(Cents)
        Else
            Dollars = Format(Amount, "$##,##0.00")
            printer.FontSize = BestPrinterFontFit(Dollars, CLng(MaxFont), AWidth, AHeight)
            'Printer.CurrentX = Printer.ScaleLeft + AdjustX + (Printer.ScaleWidth - Printer.TextWidth(Dollars)) / 2      ' Center it...
            ' I want to use AWidth here, but it's not right.
            printer.CurrentX = ALeft + (AWidth - printer.TextWidth(Dollars)) / 2    ' Center it horizontally..
            printer.CurrentY = ATop + (AHeight - printer.TextHeight(Dollars)) / 2   ' And vertically.
            printer.Print(Dollars)
        End If
    End Sub

    Public Property Quantity() As Integer
        Get
            Quantity = updQuantity.Value
        End Get
        Set(value As Integer)
            Try
                If value > updQuantity.Max Then value = updQuantity.Max
                If value < updQuantity.Min Then value = updQuantity.Min
                updQuantity.Value = value
            Catch ex As System.Windows.Forms.AxHost.InvalidActiveXStateException
                updQuantity.CreateControl()
                If value > updQuantity.Max Then value = updQuantity.Max
                If value < updQuantity.Min Then value = updQuantity.Min
                updQuantity.Value = value
            End Try
        End Set
    End Property

    Private Function SelectPrinter() As Boolean
        If Not PrSel.GetSelectedPrinter Is Nothing Then
            TicketPath = PrSel.GetSelectedPrinter.DeviceName
        End If

        If Not SetPrinter(TicketPath) Then
            MessageBox.Show("Failed to set printer.", "Set Printer")
            Exit Function
        End If

        If Not KitTag Then PrintingAllowed = True 'Package Tickets
        SelectPrinter = True
    End Function

    Private Sub txtCopies_Enter(sender As Object, e As EventArgs) Handles txtCopies.Enter
        SelectContents(txtCopies)
    End Sub

    Private Sub Legacy()
        Dim TSCTTP As Boolean, AdjustX As Integer  ' special jewelery label printer
        printer.FontName = "Arial"
        printer.FontSize = 5
        Dim OriginalPrint As String
        OriginalPrint = printer.DeviceName

        TSCTTP = True
        If TSCTTP Then
            AdjustX = 950
        Else
            AdjustX = 0
            If Not SetDymoPrinter() Then
                MessageBox.Show("Dymo Printer Required!", "WinCDS")
                Exit Sub
            End If
        End If

        printer.CurrentX = 100 + AdjustX
        printer.CurrentY = 0

        'Side 1
        printer.Print(Trim(Microsoft.VisualBasic.Left(InvenA.Desc.Text, 29)))
        printer.FontSize = 6

        MainMenu.rtbn.Visible = True
        MainMenu.rtbn.SetBarcodeSmallMediumRegular(Trim(Style))
        '.FilePrint Xx + 100, Printer.CurrentY, 0, , True, False
        MainMenu.rtbn.FilePrint(100 + AdjustX, printer.CurrentY + 50) ' jk 12/3/2007  'Printer.CurrentY, 3600, , True, False
        MainMenu.rtbn.Visible = False
        printer.CurrentX = 130 + AdjustX
        printer.CurrentY = 380
        printer.Print(Trim(Style))

        'Side 2
        printer.CurrentX = 2200 + AdjustX
        printer.CurrentY = 0
        printer.Print(" Apparaised: ")

        printer.CurrentX = 2400 + AdjustX
        printer.Print(Trim(InvenA.List.Text))
        printer.FontSize = 6
        printer.CurrentX = 2200 + AdjustX
        printer.FontBold = True
        printer.Print("Legacy: ")
        printer.CurrentX = 2400 + AdjustX
        printer.Print(Trim(InvenA.OnSale.Text))
        printer.FontBold = False
        printer.EndDoc()  'jk added 12-3-2007

        If OriginalPrint <> "" Then
            If Not SetPrinter(OriginalPrint) Then
                MessageBox.Show("Could not restore the original printer!", "Original Printer")
            End If
        End If
    End Sub

    Private Sub AutoPrint(ByVal DefaultTagSize As String, ByVal DefaultTicketPath As String, ByVal ParentForm As Form)
        TagSize = DefaultTagSize
        TicketPath = DefaultTicketPath
        If Trim(TicketPath) = "" Then
            'Show vbModal, ParentForm
            ShowDialog(ParentForm)
            Exit Sub
        End If
        Select Case TagSize
            'Case "MED", "MEDIUM" : cmdMedium.Value = True
            Case "MED", "MEDIUM" : cmdMedium_Click(cmdMedium, New EventArgs)
            'Case "SMALL" : cmdSmall.Value = True
            Case "SMALL" : cmdSmall_Click(cmdSmall, New EventArgs)
            'Case "LARGE" : cmdLarge.Value = True
            Case "LARGE" : cmdLarge_Click(cmdLarge, New EventArgs)
            'Case "DYMO" : cmdDYMO.Value = True
            Case "DYMO" : cmdDYMO_Click(cmdDYMO, New EventArgs)
            Case Else
                'Show vbModal, ParentForm ' Invalid options, so show the form.
                ShowDialog(ParentForm)
        End Select
    End Sub

    ' BFH20071118 - added this..
    ' for store num, -1 = current store, 0 = all stores
    Public Function PrintTagsByStyle(ByVal Style As String, Optional ByVal Count As Integer = 1, Optional ByVal Size As String = "LARGE", Optional ByVal StoreNum As Integer = -1) As Integer
        Dim M As CInvRec, ST As String, Path As String
        Dim I As Integer

        If Not IsIn(Size, "MED", "LARGE", "SMALL", "DYMO") Then Size = "LARGE"
        Path = TicketPrinter()
        PrSel.SetSelectedPrinter(Path)

        M = New CInvRec
        If M.Load(Style, "Style") Then
            ST = IIf(StoreNum = -1, M.QueryStock(StoresSld), IIf(StoreNum = 0, M.QueryTotalStock, M.QueryStock(StoreNum)))
            For I = 1 To Count
                PrintTags(M.Style, M.Desc, M.Landed, M.List, M.OnSale, M.DeptNo, M.GetItemCode(), M.Vendor, ST, M.Comments,
      , Microsoft.VisualBasic.Left(M.Style, 4) = KIT_PFX, , , True, Size, Path)
            Next
        End If
        DisposeDA(M)
    End Function

    Private Sub LoadTagInfo(ByVal nStyle As String, ByVal nDesc As String, ByVal nLanded As String,
  ByVal nList As String, ByVal nOnSale As String, ByVal nDeptNo As String, ByVal nCode As String,
  ByVal nVendor As String, ByVal nStock As String, ByVal nComments As String)

        Style = nStyle
        Landed = nLanded

        List = nList
        OnSale = nOnSale
        Desc = nDesc
        Code = nCode
        Mfg = nVendor
        Stock = nStock
        Comments = nComments
        SKU = nComments
    End Sub
End Class
﻿Public Class ServiceReports
    ' For printing service calls..
    Private WithEvents mDBAccess As CDbAccessGeneral
    Private WithEvents mDBAccessParts As CDbAccessGeneral
    Private WithEvents mDBAccessBilling As CDbAccessGeneral
    Dim CY As Integer = 900
    Private ReadOnly Property Mode() As String
        Get
            Mode = Order
        End Get
    End Property

    Public ReadOnly Property ReportTitle() As String
        Get
            Select Case Mode
                Case "SCR" : ReportTitle = "Open Service Call Report"
                Case "SPR" : ReportTitle = "Open Part Orders Report"
                Case "SBR" : ReportTitle = "Service Parts Billing Report"
                Case "SBU" : ReportTitle = "Unpaid Service Orders"
                Case Else : ReportTitle = "Service Report"
            End Select
        End Get
    End Property

    Public ReadOnly Property ReportHelpContext() As Integer
        Get
            Select Case Mode
                Case "SCR" : ReportHelpContext = 49690
                Case "SPR" : ReportHelpContext = 49690
                Case "SBR" : ReportHelpContext = 49690
                Case "SBU" : ReportHelpContext = 49690
                Case Else : ReportHelpContext = 0
            End Select
        End Get
    End Property

    Private Sub ServiceReports_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdPrint, 19)
        SetButtonImage(cmdPrintPreview, 20)
        SetButtonImage(cmdCancel, 3)
        Text = ReportTitle
        'HelpContextID = ReportHelpContext
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        'Unload Me
        Me.Close()
        MainMenu.Show()
    End Sub

    Private Sub RunReport()
        Select Case Mode
            Case "SCR" : repServiceCall()
            Case "SPR" : repServiceParts()
            Case "SBR" : repServiceBilling()
            Case "SBU" : repServiceUnpaid()
        End Select
    End Sub

    Private Sub repServiceCall()
        '  UndeliveredHeading
        mDBAccess = New CDbAccessGeneral
        mDBAccess.dbOpen(GetDatabaseAtLocation())
        mDBAccess.SQL = "SELECT * From Service WHERE Status='Open' ORDER BY Service.ServiceOrderNo"
        mDBAccess.GetRecord()
        On Error Resume Next
        mDBAccess.dbClose()
        mDBAccess = Nothing
    End Sub

    Private Sub repServiceParts()
        mDBAccessParts = New CDbAccessGeneral
        mDBAccessParts.dbOpen(GetDatabaseAtLocation())
        mDBAccessParts.SQL = "SELECT * From ServicePartsOrder WHERE Status='Open' ORDER BY ServicePartsOrder.ServicePartsOrderNo"
        mDBAccessParts.GetRecord()
        On Error Resume Next
        mDBAccessParts.dbClose()
        mDBAccessParts = Nothing
    End Sub

    Private Sub repServiceBilling()
        mDBAccessBilling = New CDbAccessGeneral
        mDBAccessBilling.dbOpen(GetDatabaseAtLocation())
        mDBAccessBilling.SQL = "SELECT * From ServicePartsOrder WHERE Status='Open' AND ChargeBackAmount > 0 AND Paid = 0 ORDER BY Vendor, DateOfClaim"
        mDBAccessBilling.GetRecord()
        On Error Resume Next
        mDBAccessBilling.dbClose()
        mDBAccessBilling = Nothing
    End Sub

    Private Sub repServiceUnpaid()
        mDBAccessBilling = New CDbAccessGeneral
        mDBAccessBilling.dbOpen(GetDatabaseAtLocation())
        mDBAccessBilling.SQL = "SELECT * From ServicePartsOrder WHERE ChargeBackAmount > 0 AND Paid = 0 ORDER BY Vendor, DateOfClaim"
        mDBAccessBilling.GetRecord()
        On Error Resume Next
        mDBAccessBilling.dbClose()
        mDBAccessBilling = Nothing
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Working(True)
        OutputToPrinter = True

        If Not IsDevelopment() Then
            OutputObject = Printer
        Else
            OutputObject = New cPrinter
            OutputObject.SetPrintToPDF("Test Report")
        End If
        RunReport()
        OutputObject.EndDoc
        Working(False)
        'cmdCancel.Value = True  ' we probably want to close this form as it's not that interesting!
        cmdCancel_Click(cmdCancel, New EventArgs)
    End Sub

    Private Sub Working(ByVal begin As Boolean)
        'MousePointer = IIf(Not begin, vbDefault, vbHourglass)
        Me.Cursor = IIf(Not begin, Cursors.Default, Cursors.WaitCursor)
        cmdPrint.Enabled = Not begin
        cmdPrintPreview.Enabled = Not begin
        cmdCancel.Enabled = Not begin
    End Sub

    Private Sub cmdPrintPreview_Click(sender As Object, e As EventArgs) Handles cmdPrintPreview.Click
        Working(True)

        'Load frmPrintPreviewMain
        OutputToPrinter = False
        OutputObject = frmPrintPreviewDocument.picPicture

        frmPrintPreviewDocument.CallingForm = Me
        frmPrintPreviewDocument.ReportName = Text
        RunReport()
        Hide()
        frmPrintPreviewDocument.DataEnd()

        Working(False)
    End Sub

    Public Sub Heading()
        Dim Y As Integer
        OutputObject.FontName = "Arial"

        OutputObject.FontSize = 18
        PrintCentered(ReportTitle, 100, True)

        OutputObject.FontSize = 8
        'PrintAligned("Time: " & Format(Now, "h:mm:ss am/pm"), , 10, 100)
        PrintAligned("Time: " & Format(Now, "h:mm:ss tt"), , 10, 100)

        If OutputToPrinter Then PageNumber = OutputObject.Page
        PrintAligned("Page: " & PageNumber, , 10100, 100)

        PrintCentered(StoreSettings.Name & "    " & StoreSettings.Address & "    " & StoreSettings.City, 500)

        OutputObject.FontSize = 9
        OutputObject.FontBold = True
        Y = OutputObject.CurrentY
        Select Case Mode
            Case "SCR"
                OutputObject.CurrentX = 0
                OutputObject.CurrentY = 700
                'PrintToTab(, "ServiceNo", 0)
                'PrintToTab(, "ServiceNo", 0,,, 700, 0, True)
                PrintToPosition2(, "ServiceNo", ,,, 700, 0)
                'PrintToTab(, "DateOfClaim", 20)
                'PrintToTab(, "DateOfClaim", 20,,, 700, 20, True)
                PrintToPosition2(, "DateOfClaim", ,,, 700, 1500)
                'PrintToTab(, "Last", 40)
                'PrintToTab(, "Last", 40,,, 700, 40, True)
                PrintToPosition2(, "Last", ,,, 700, 3000)
                'PrintToTab(, "Telephone", 60)
                'PrintToTab(, "Telephone", 60,,, 700, 60, True)
                PrintToPosition2(, "Telephone", ,,, 700, 4800)
                OutputObject.FontBold = False
            Case "SPR"
                PrintAligned("PartsOrderNo", , 10, Y, True)
                PrintAligned("Status", , 1300, Y, True)
                PrintAligned("ServiceNo", , 2000, Y, True)
                PrintAligned("Vendor", , 3000, Y, True)
                PrintAligned("DateOfClaim", , 5900, Y, True)
                PrintAligned("Repair Cost", , 7200, Y, True)
                PrintAligned("Paid", , 8400, Y, True)
            Case "SBR"
                PrintAligned("Vendor", , 10, Y, True)
                PrintAligned("Date", , 3200, Y, True)
                PrintAligned("Repair Cost", , 4500, Y, True)
                PrintAligned("Type", , 6000, Y, True)
                PrintAligned("PartsOrderNo", , 7500, Y, True)
                PrintAligned("Status", , 8800, Y, True)
                PrintAligned("Service No", , 9500, Y, True)
        End Select
        OutputObject.FontBold = False
    End Sub

    Private Sub DoNewPage(Optional ByVal ExtraLines As Integer = 2)
        Dim ExtraHeight As Integer
        Dim Newpage As Boolean

        ExtraHeight = OutputObject.TextHeight("X") * ExtraLines
        If OutputObject.ScaleHeight < OutputObject.CurrentY + ExtraHeight Then
            If OutputToPrinter Then
                Printer.NewPage()
                Newpage = True
                CY = 700
            Else
                frmPrintPreviewDocument.NewPage()
            End If
        End If
        If OutputObject.CurrentY = 0 And Newpage = False Then
            Heading()
            'If Not OutputToPrinter Then frmPrintPreviewDocument.NewPage
        End If
    End Sub

    Private Sub mDBAccess_GetRecordEvent(RS As ADODB.Recordset) Handles mDBAccess.GetRecordEvent   ' called if record is found
        Dim ServiceNo As String
        Dim Last As String
        Dim Tele As String
        Dim DateOfClaim As String
        Dim Item As String, ILine As Object, FN As String
        Dim GM As CGrossMargin

        Do Until RS.EOF
            ServiceNo = RS("ServiceOrderNo").Value
            Last = RS("LastName").Value
            Tele = RS("Telephone").Value
            DateOfClaim = RS("DateOfClaim").Value
            Item = Trim(RS("Item").Value)
            ' Item needs to be augmented..
            ' We need GrossMargin records (rendered).. based on what's in ServiceItemParts for this ServiceOrderNo.
            GM = New CGrossMargin
            GM.DataAccess.DataBase = GetDatabaseAtLocation()
            GM.DataAccess.Records_OpenSQL("SELECT GrossMargin.* FROM GrossMargin INNER JOIN ServiceItemParts on GrossMargin.MarginLine=ServiceItemParts.MarginNo WHERE ServiceItemParts.ServiceOrderNo=" & ServiceNo & " ORDER BY MarginLine")
            Do While GM.DataAccess.Records_Available
                GM.cDataAccess_GetRecordSet(GM.DataAccess.RS)
                Item = Item & vbCrLf & GM.Vendor & " " & GM.Style & " " & GM.SaleNo & " " & AlignString(CStr(GM.Quantity), 6, VBRUN.AlignConstants.vbAlignLeft) & " " & DateFormat(GM.DDelDat) & "  " & Trim(GM.Desc)
            Loop
            DisposeDA(GM)

            DoNewPage()

            'PrintToTab(, Trim(ServiceNo), 0)
            'PrintToTab(, Trim(ServiceNo), 0,,, CY, True)
            PrintToPosition2(, Trim(ServiceNo), ,,, CY, 0)
            'PrintToTab(, DateFormat(DateOfClaim), 20)
            'PrintToTab(, DateFormat(DateOfClaim), 20,,, CY, True)
            PrintToPosition2(, DateFormat(DateOfClaim), ,,, CY, 1500)
            'PrintToTab(, Microsoft.VisualBasic.Left(Last, 20), 40)
            'PrintToTab(, Microsoft.VisualBasic.Left(Last, 20), 40,,, CY, True)
            PrintToPosition2(, Microsoft.VisualBasic.Left(Last, 20), ,,, CY, 3000)
            'PrintToTab(, DressAni(CleanAni(Tele, 0)), 60, , True)
            'PrintToTab(, DressAni(CleanAni(Tele, 0)), 60, ,, CY, True)
            PrintToPosition2(, DressAni(CleanAni(Tele, 0)), ,,, CY, 4800)

            Do While Microsoft.VisualBasic.Left(Item, 2) = vbCrLf : Item = Trim(Mid(Item, 3)) : Loop                ' Trim off extra blank lines.
            Do While Microsoft.VisualBasic.Right(Item, 2) = vbCrLf : Item = Trim(Microsoft.VisualBasic.Left(Item, Len(Item) - 2)) : Loop  ' Trim off extra blank lines.

            FN = OutputObject.FontName
            OutputObject.FontName = "Lucida Console"
            For Each ILine In Split(Item, vbCrLf)
                'OutputObject.Print("      Item:", ILine)
                OutputObject.Print("      Item:" & " " & ILine)
            Next
            OutputObject.FontName = FN
            If Item = "" Then OutputObject.Print("  No Item Specified")
            OutputObject.Print
            RS.MoveNext()
            CY = CY + 600
        Loop
    End Sub

    Private Sub mDBAccessParts_GetRecordEvent(RS As ADODB.Recordset) Handles mDBAccessParts.GetRecordEvent   ' called if record is found
        Dim PartsOrderNo As String, ServiceNo As String
        Dim Status As String
        Dim DateOfClaim As String, Vendor As String
        '  Dim Style As String, Desc As String
        Dim RepairCost As String, Paid As String
        Dim CBType As String

        Do Until RS.EOF
            PartsOrderNo = IfNullThenNilString(RS("ServicePartsOrderNo").Value)
            Status = IfNullThenNilString(RS("Status").Value)
            ServiceNo = IfNullThenNilString(RS("ServiceOrderNo").Value)
            DateOfClaim = IfNullThenNilString(RS("DateOfClaim").Value)
            Vendor = IfNullThenNilString(RS("Vendor").Value)

            CBType = ServiceParts.ChargeBackTypeDesc(RS("ChargeBackType").Value)
            RepairCost = FormatCurrency(IfNullThenZero(RS("ChargeBackAmount").Value))
            Paid = YesNo(IfNullThenZero(RS("Paid").Value) <> 0)

            DoNewPage()

            Dim Y As Integer
            Y = OutputObject.CurrentY

            PrintAligned(PartsOrderNo, , 10, Y)
            PrintAligned(Status, , 1300, Y)
            PrintAligned(IIf(Val(ServiceNo) > 0, ServiceNo, "[none]"), , 2000, Y)
            PrintAligned(Microsoft.VisualBasic.Left(Vendor, 30), , 3000, Y)
            PrintAligned(DateOfClaim, , 5900, Y)
            PrintAligned(RepairCost, , 7200, Y)
            PrintAligned(Paid, , 8400, Y)
            RS.MoveNext()
        Loop
    End Sub

    Private Sub mDBAccessBilling_GetRecordEvent(RS As ADODB.Recordset) Handles mDBAccessBilling.GetRecordEvent    ' called if record is found
        Dim PartsOrderNo As String, ServiceNo As String
        Dim Status As String
        Dim DateOfClaim As String, Vendor As String
        '  Dim Style As String, Desc As String
        Dim RepairCost As String, Paid As String
        Dim CBType As String

        Dim TotCost As Decimal

        TotCost = 0
        Do Until RS.EOF
            PartsOrderNo = IfNullThenNilString(RS("ServicePartsOrderNo").Value)
            Status = IfNullThenNilString(RS("Status").Value)
            ServiceNo = IfNullThenNilString(RS("ServiceOrderNo").Value)
            DateOfClaim = IfNullThenNilString(RS("DateOfClaim").Value)
            Vendor = IfNullThenNilString(RS("Vendor").Value)

            CBType = ServiceParts.ChargeBackTypeDesc(RS("ChargeBackType").Value)
            RepairCost = FormatCurrency(IfNullThenZero(RS("ChargeBackAmount").Value))
            Paid = YesNo(IfNullThenZero(RS("Paid").Value) <> 0)

            DoNewPage()

            Dim Y As Integer
            Y = OutputObject.CurrentY
            PrintAligned(Microsoft.VisualBasic.Left(Vendor, 30), , 10, Y)
            PrintAligned(DateOfClaim, , 3200, Y)
            PrintAligned(RepairCost, , 4500, Y)
            PrintAligned(CBType, , 6000, Y)
            PrintAligned(PartsOrderNo, , 7500, Y)
            PrintAligned(Status, , 8800, Y)
            PrintAligned(ServiceNo, , 9500, Y)
            TotCost = TotCost + RS("ChargeBackAmount").Value
            RS.MoveNext()
        Loop

        DoNewPage(3)

        PrintAligned("----------", , 4500, , True)
        PrintAligned(FormatCurrency(TotCost), , 4500, , True)
        PrintAligned("----------", , 4500, , True)
    End Sub

End Class
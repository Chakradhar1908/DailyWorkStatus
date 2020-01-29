Module modOrigInv
    Public Sub BarcodeInventoryCheck()
        Dim InvData As New CInvRec

        GetAllBarcodes(MainMenu)
        If BarcodeFormType = 0 Then 'Barcodes found
            'Load frmPrintPreviewMain
            OutputToPrinter = False
            OutputObject = frmPrintPreviewDocument.picPicture
            frmPrintPreviewDocument.CallingForm = MainMenu
            frmPrintPreviewDocument.ReportName = "Barcode Inventory Check"
            BarcodeInventoryCheckHeading()
            Dim X As Integer, I As Integer
            For X = LBound(BarcodeList) To UBound(BarcodeList)
                If InvData.Load(BarcodeList(X)) Then
                    If OutputObject.Height < OutputObject.CurrentY + 1500 Then 'New Page
                        If OutputToPrinter Then
                            Printer.NewPage()
                        Else
                            frmPrintPreviewDocument.NewPage()
                        End If
                        BarcodeInventoryCheckHeading
                    End If
                    OutputObject.CurrentX = 500 'Align Left
                    '###STORECOUNT8
                    OutputObject.Print(InvData.Style, , InvData.QueryStock(1), InvData.QueryStock(2), InvData.QueryStock(3), InvData.QueryStock(4), InvData.QueryStock(5), InvData.QueryStock(6), InvData.QueryStock(7), InvData.QueryStock(8))
                    OutputObject.CurrentX = 500 'Align Left
                    '###STORECOUNT8
                    OutputObject.Print("Total Available: " & InvData.Available, , InvData.QueryOnOrder(1), InvData.QueryOnOrder(2), InvData.QueryOnOrder(3), InvData.QueryOnOrder(4), InvData.QueryOnOrder(5), InvData.QueryOnOrder(6), InvData.QueryOnOrder(7), InvData.QueryOnOrder(8))
                    OutputObject.CurrentX = 500 'Align Left
                    OutputObject.Print(InvData.Vendor, , InvData.Desc, InvData.OnSale, InvData.List)
                    'OutputObject.Line(500, OutputObject.CurrentY + 50)-(Printer.ScaleWidth - 150, OutputObject.CurrentY + 50)
                    OutputObject.Line(500, OutputObject.CurrentY + 50, Printer.ScaleWidth - 150, OutputObject.CurrentY + 50)
                    OutputObject.CurrentY = OutputObject.CurrentY + 100
                    Exit For
                End If
            Next
            If OutputToPrinter Then
                Printer.EndDoc()
            Else
                MainMenu.Hide()
                frmPrintPreviewDocument.DataEnd()
            End If
        End If
        DisposeDA(InvData)
    End Sub

    Private Sub BarcodeInventoryCheckHeading()
        OutputObject.CurrentX = 0
        OutputObject.CurrentY = 100
        OutputObject.DrawWidth = 2
        OutputObject.FontBold = True
        OutputObject.FontName = "Arial"
        OutputObject.FontSize = 18

        PrintAligned("Physical Inventory Discrepancy Report", VBRUN.AlignmentConstants.vbCenter)

        OutputObject.FontBold = False
        OutputObject.FontSize = 8
        OutputObject.CurrentY = 100 'Align Top

        If OutputToPrinter Then
            PrintAligned("Date: " & DateFormat(Now) & vbCrLf & "Time: " & Format(Now, "h:mm:ss am/pm") & vbCrLf & "Page: " & OutputObject.Page, VBRUN.AlignmentConstants.vbLeftJustify, 9900)
        Else
            PrintAligned("Date: " & DateFormat(Now) & vbCrLf & "Time: " & Format(Now, "h:mm:ss am/pm") & vbCrLf & "Page: " & PageNumber, VBRUN.AlignmentConstants.vbLeftJustify, 9900)
        End If

        OutputObject.CurrentX = 500 'Align Left
        OutputObject.CurrentY = 750 'Align Top

        OutputObject.FontBold = True
        '###STORECOUNT8
        OutputObject.Print("Style", , , "Loc1Bal", "Loc2Bal", "Loc3Bal", "Loc4Bal", "Loc5Bal", "Loc6Bal", "Loc7Bal", "Loc8Bal")
        OutputObject.CurrentX = 500 'Align Left
        '###STORECOUNT8
        OutputObject.Print("Available", , "OnOrder1", "OnOrder2", "OnOrder3", "OnOrder4", "OnOrder5", "OnOrder6", "OnOrder7", "OnOrder8")
        OutputObject.CurrentX = 500 'Align Left
        OutputObject.Print("Manufacturer", , "Description", , , , "On Sale", "List")
        'MainMenu.FontBold = False
        MainMenu.Font = New Font(MainMenu.Font.Name, MainMenu.Font.Size, FontStyle.Regular)
        'OutputObject.Line(500, OutputObject.CurrentY + 50)-(Printer.ScaleWidth - 150, OutputObject.CurrentY + 50)
        OutputObject.Line(500, OutputObject.CurrentY + 50, Printer.ScaleWidth - 150, OutputObject.CurrentY + 50)
        OutputObject.CurrentY = OutputObject.CurrentY + 100
    End Sub
End Module

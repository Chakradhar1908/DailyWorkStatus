Imports System.Drawing.Printing
Module modPrintRTF
    Private Structure FormatRange
        'Dim hDC as integer       ' Actual DC to draw on
        Dim hDC As IntPtr
        'Dim hdcTarget as integer ' Target DC for determining text formatting
        Dim hdcTarget As IntPtr ' Target DC for determining text formatting
        Dim rc As RECT        ' Region of the DC to draw to (in twips)
        Dim rcPage As RECT    ' Region of the entire DC (page size) (in twips)
        Dim chrg As CharRange ' Range of text to draw (see above declaration)
    End Structure

    Private Structure CharRange
        Dim cpMin As Integer     ' First character of range (0 for start of doc)
        Dim cpMax As Integer     ' Last character of range (-1 for end of doc)
    End Structure
    'Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC as integer, ByVal nIndex as integer) as integer
    Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As IntPtr, ByVal nIndex As Integer) As Integer
    Private Const WM_USER As Integer = &H400
    Private Const EM_FORMATRANGE As Integer = WM_USER + 57
    Private Const EM_SETTARGETDEVICE As Integer = WM_USER + 72
    Private Const PHYSICALOFFSETX As Integer = 112
    Private Const PHYSICALOFFSETY As Integer = 113

    '<CT>
    Public DeliveryticketMessageFileText As String
    '</CT>
    Public Sub PrintRTF(ByRef RTF As RichTextBox,
  Optional LeftMarginWidth As Integer = -1, Optional TopMarginHeight As Integer = -1,
  Optional PrintWidth As Integer = -1, Optional PrintHeight As Integer = -1,
  Optional NoAdjustment As Boolean = False, Optional AllowMultiplePages As Boolean = True)

        Dim LeftLimit As Integer, TopLimit As Integer
        Dim FR As FormatRange
        Dim rcDrawTo As RECT
        Dim rcPage As RECT
        Dim TextLength As Integer
        'Dim NextCharPosition As Integer
        Dim NextCharPosition As IntPtr

        ' Start a print job to get a valid Printer.hDC
        'If Printer.hdc = 0 Then
        Printer.Print(Space(0))
        Printer.ScaleMode = VBRUN.ScaleModeConstants.vbTwips
        'Printer.ScaleMode = VBRUN.ScaleModeConstants.vbPixels

        ' Get the offset to the printable area on the page in twips
        'LeftLimit = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
        'TopLimit = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
        Dim p As New PrinterSettings
        Dim g As Graphics = p.CreateMeasurementGraphics()
        'g.GetHdc()


        On Error Resume Next
        LeftLimit = Printer.ScaleX(GetDeviceCaps(g.GetHdc, PHYSICALOFFSETX), VBRUN.ScaleModeConstants.vbPixels, VBRUN.ScaleModeConstants.vbTwips)
        TopLimit = Printer.ScaleY(GetDeviceCaps(g.GetHdc, PHYSICALOFFSETY), VBRUN.ScaleModeConstants.vbPixels, VBRUN.ScaleModeConstants.vbTwips)
        ' Make sure the print area starts in a printable area.
        If LeftMarginWidth < 0 Then LeftMarginWidth = Printer.CurrentX
        If TopMarginHeight < 0 Then TopMarginHeight = Printer.CurrentY

        ' Set the right boundaries based on the width of the printable area.
        If PrintWidth < 0 Then PrintWidth = Printer.ScaleWidth - LeftMarginWidth
        If PrintHeight < 0 Then PrintHeight = Printer.ScaleHeight - TopMarginHeight

        '' Removed because this interferes with aligned tags, which don't use the whole print area.
        ''If SelectPrinter.TagSize = "SMALL" Or SelectPrinter.TagSize = "MED" Or SelectPrinter.TagSize = "LARGE" Then
        ''      If RightMargin > 11500 Then RightMargin = 11500  ' Was 12100
        ''  Else:
        ''       If RightMargin > 8350 Then RightMargin = 8350 'I put in to stop over printing
        '' End If

        ' Set printable area rect
        rcPage.Left = 0
        rcPage.Top = 0
        rcPage.Right = Printer.ScaleWidth
        rcPage.Bottom = Printer.ScaleHeight

        ' Set rect in which to print (relative to printable area)
        rcDrawTo.Left = LeftMarginWidth
        rcDrawTo.Top = TopMarginHeight
        rcDrawTo.Right = LeftMarginWidth + PrintWidth
        rcDrawTo.Bottom = TopMarginHeight + PrintHeight

        ' Set up the print instructions
        'FR.hDC = Printer.hDC        ' Use the same DC for measuring and rendering
        FR.hDC = g.GetHdc         ' Use the same DC for measuring and rendering
        'FR.hdcTarget = Printer.hDC  ' Point at printer hDC
        FR.hdcTarget = g.GetHdc   ' Point at printer hDC
        FR.rc = rcDrawTo            ' Indicate the area on page to draw to
        FR.rcPage = rcPage          ' Indicate entire size of page
        FR.chrg.cpMin = 0           ' Indicate start of text through
        FR.chrg.cpMax = -1          ' end of the text

        ' Get length of text in RTF
        TextLength = Len(RTF.Text)
        'Dim x As String = RTF.Text
        'Printer.CurrentX = LeftMarginWidth
        'Printer.CurrentY = TopMarginHeight
        'Printer.Print(x)

        ' Loop printing each page until done
        Dim NextCharPosition2 as integer
        Do

            ' Print the page by sending EM_FORMATRANGE message
            NextCharPosition = SendMessage(RTF.Handle, EM_FORMATRANGE, True, FR.chrg.ToString)   ' This prints the entire current page.
            'Public Function SendMessage(ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As IntPtr
            NextCharPosition2 = CLng(NextCharPosition)
            If NextCharPosition2 >= TextLength Or Not AllowMultiplePages Then Exit Do                       ' If done then exit

            FR.chrg.cpMin = NextCharPosition ' Starting position for next page
            Printer.NewPage()                  ' Move on to next page
            Printer.Print(Space(0)) ' Re-initialize hDC
            FR.hDC = g.GetHdc
            FR.hdcTarget = g.GetHdc
        Loop      ' Commit the print job

        g.ReleaseHdc()
        'r =
        SendMessage(RTF.Handle, EM_FORMATRANGE, False, CLng(0))
        '<CT>
        DeliveryticketMessageFileText = RTF.Text
        '</CT>
    End Sub
End Module

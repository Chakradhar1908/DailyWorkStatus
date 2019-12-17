Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module modPrintRTF
    Private Structure FormatRange
        Dim hDC As Long       ' Actual DC to draw on
        Dim hdcTarget As Long ' Target DC for determining text formatting
        Dim rc As RECT        ' Region of the DC to draw to (in twips)
        Dim rcPage As RECT    ' Region of the entire DC (page size) (in twips)
        Dim chrg As CharRange ' Range of text to draw (see above declaration)
    End Structure
    Private Structure CharRange
        Dim cpMin As Long     ' First character of range (0 for start of doc)
        Dim cpMax As Long     ' Last character of range (-1 for end of doc)
    End Structure
    Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

    Public Sub PrintRTF(ByRef RTF As RichTextBox,
  Optional LeftMarginWidth As Long = -1, Optional TopMarginHeight As Long = -1,
  Optional PrintWidth As Long = -1, Optional PrintHeight As Long = -1,
  Optional NoAdjustment As Boolean = False, Optional AllowMultiplePages As Boolean = True)

        Dim LeftLimit As Long, TopLimit As Long
        Dim FR As FormatRange
        Dim rcDrawTo As RECT
        Dim rcPage As RECT
        Dim TextLength As Long
        Dim NextCharPosition As Long

        ' Start a print job to get a valid Printer.hDC
        'If Printer.hdc = 0 Then
        Printer.Print(Space(0))
        Printer.ScaleMode = VBRUN.ScaleModeConstants.vbTwips

        ' Get the offset to the printable area on the page in twips
        LeftLimit = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
        TopLimit = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)

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
        FR.hDC = Printer.hDC        ' Use the same DC for measuring and rendering
        FR.hdcTarget = Printer.hDC  ' Point at printer hDC
        FR.rc = rcDrawTo            ' Indicate the area on page to draw to
        FR.rcPage = rcPage          ' Indicate entire size of page
        FR.chrg.cpMin = 0           ' Indicate start of text through
        FR.chrg.cpMax = -1          ' end of the text

        ' Get length of text in RTF
        TextLength = Len(RTF.Text)

        ' Loop printing each page until done
        Do
            ' Print the page by sending EM_FORMATRANGE message
            NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, FR)   ' This prints the entire current page.
            If NextCharPosition >= TextLength Or Not AllowMultiplePages Then Exit Do                       ' If done then exit

            FR.chrg.cpMin = NextCharPosition ' Starting position for next page
            Printer.NewPage()                  ' Move on to next page
            Printer.Print Space(0) ' Re-initialize hDC
            FR.hDC = Printer.hDC
            FR.hdcTarget = Printer.hDC
        Loop      ' Commit the print job

        'r =
        SendMessage RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0)
End Sub

End Module

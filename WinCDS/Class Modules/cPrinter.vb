Imports WinCDS.clsPDFPrinter
Imports stdole
Public Class cPrinter
    Private Const PI As Double = 3.14159265358979
    Private Const PI_2 As Double = 6.28318530717958
    Public mBuildPDF As Boolean
    Private mPreviewImage As Object
    Private mDocTitle As String
    Private mDocFile As String
    Private mDocKeywords As String
    Private pW As Integer, pH As Integer
    Private PDFPrinter As clsPDFPrinter

    Public Sub New()
        OutputToPrinter = True
        Orientation = VBRUN.PrinterObjectConstants.vbPRORPortrait
    End Sub

    Public Property Orientation() As Integer
        Get
            On Error Resume Next
            Orientation = oPrinter.Orientation
        End Get
        Set(value As Integer)
            On Error Resume Next
            oPrinter.Orientation = value
        End Set
    End Property

    Public ReadOnly Property oPrinter() As Object
        Get
            If BuildDLS Then
                oPrinter = DYMOObject
            ElseIf Preview Then
                oPrinter = PreviewImage
            Else
                oPrinter = Printer
            End If
        End Get
    End Property

    Public ReadOnly Property DYMOObject(Optional ByVal Reset As Boolean = False) As Object 'As Dymo.LabelEngine
        Get
            On Error Resume Next
            Static dymoAddInObj  'As Dymo.LabelEngine
            If Reset Then dymoAddInObj = Nothing
            If dymoAddInObj Is Nothing Then
                dymoAddInObj = CreateObject("DYMO.LabelEngine")
                dymoAddInObj.NewLabel(DateTimeStamp)
                'DYMO_Label_Framework.ope
                'Dim X As DYMO_Label_Framework.Label
            End If
            DYMOObject = dymoAddInObj
        End Get
    End Property

    Public ReadOnly Property DYMOAddIn(Optional ByVal Reset As Boolean = False) As Object 'As Dymo.DYMOAddIn
        Get
            On Error Resume Next
            Static dymoAddInObj 'As Dymo.DYMOAddIn
            If Reset Then DYMOAddIn = Nothing
            If IsNothing(DYMOAddIn) Then
                dymoAddInObj = CreateObject("DYMO.DYMOAddIn")
            End If
            DYMOAddIn = dymoAddInObj
        End Get
    End Property

    Public ReadOnly Property Preview() As Boolean
        Get
            Preview = Not (mPreviewImage Is Nothing)
        End Get
    End Property

    Public Property PrintQuality() As Integer
        Get
            PrintQuality = Printer.PrintQuality
        End Get
        Set(value As Integer)
            Printer.PrintQuality = value
        End Set
    End Property

    Public Property RightToLeft() As Boolean
        Get
            RightToLeft = Printer.RightToLeft
        End Get
        Set(value As Boolean)
            Printer.RightToLeft = value
        End Set
    End Property

    Public Property PreviewImage() As Object
        Get
            PreviewImage = mPreviewImage
        End Get
        Set(value As Object)
            mPreviewImage = value
        End Set
    End Property

    Public Property ScaleHeight() As Single
        Get
            ScaleHeight = oPrinter.ScaleHeight
        End Get
        Set(value As Single)
            oPrinter.ScaleHeight = value
        End Set
    End Property

    Public Property ScaleTop() As Single
        Get
            ScaleTop = oPrinter.ScaleTop
        End Get
        Set(value As Single)
            oPrinter.ScaleTop = value
        End Set
    End Property

    Public Property ScaleLeft() As Single
        Get
            ScaleLeft = oPrinter.ScaleLeft
        End Get
        Set(value As Single)
            oPrinter.ScaleLeft = value
        End Set
    End Property

    Public Property ScaleWidth() As Single
        Get
            ScaleWidth = oPrinter.ScaleWidth
        End Get
        Set(value As Single)
            oPrinter.ScaleWidth = value
        End Set
    End Property

    Public ReadOnly Property BuildDLS() As Boolean
        Get
            BuildDLS = False And IsDymo And HasDLS
        End Get
    End Property

    Public Property ScaleMode() As Integer
        Get
            ScaleMode = oPrinter.ScaleMode
        End Get
        Set(value As Integer)
            oPrinter.ScaleMode = value
        End Set
    End Property

    Public ReadOnly Property IsDymo() As Boolean
        Get
            IsDymo = IsInStr(DeviceName, "DYMO")
        End Get
    End Property

    Public ReadOnly Property DeviceName() As String
        Get
            On Error Resume Next
            DeviceName = Printer.DeviceName
        End Get
    End Property

    Public Property DrawMode() As Integer
        Get
            DrawMode = oPrinter.DrawMode
        End Get
        Set(value As Integer)
            oPrinter.DrawMode = value
        End Set
    End Property

    Public Property DrawStyle() As Integer
        Get
            DrawStyle = oPrinter.DrawStyle
        End Get
        Set(value As Integer)
            oPrinter.DrawStyle = value
        End Set
    End Property

    Public Property DrawWidth() As Integer
        Get
            DrawWidth = oPrinter.DrawWidth
        End Get
        Set(value As Integer)
            oPrinter.DrawWidth = value
        End Set
    End Property

    Public ReadOnly Property HasDLS() As Boolean
        Get
            Static vValue As TriState
            On Error Resume Next
            If vValue = vbFalse Then vValue = IIf(IsNotNothing(CreateObject("DYMO.LabelEngine")), vbTrue, vbUseDefault)
            HasDLS = (vValue = vbTrue)
        End Get
    End Property

    Public Property Height() As Integer
        Get
            Height = oPrinter.Height
        End Get
        Set(value As Integer)
            oPrinter.Height = value
        End Set
    End Property

    Public ReadOnly Property hDC() As Integer  ' Read-Only
        Get
            hDC = oPrinter.hDC
        End Get
    End Property

    Public Sub SetPrintToPDF(Optional ByVal vDocTitle As String = "", Optional ByVal vKeywords As String = "")
        If vDocTitle <> "" Then DocTitle = vDocTitle
        DocKeywords = vKeywords
        OutputToPrinter = True

        PDFInit()

        OutputObject = Me
    End Sub

    Public Property Duplex() As Integer
        Get
            Duplex = oPrinter.Duplex
        End Get
        Set(value As Integer)
            oPrinter.Duplex = value
        End Set
    End Property

    Public Property DocTitle() As String
        Get
            DocTitle = mDocTitle
        End Get
        Set(value As String)
            mDocTitle = value
        End Set
    End Property

    Public Property DocKeywords() As String
        Get
            DocKeywords = mDocKeywords
        End Get
        Set(value As String)
            mDocKeywords = value
        End Set
    End Property

    Public ReadOnly Property PDFSupportFolderExists() As Boolean
        Get
            PDFSupportFolderExists = FolderExists(PDFSupportFolder)
        End Get
    End Property

    Public ReadOnly Property Page() As Integer  ' Read-Only
        Get
            Page = oPrinter.Page
        End Get
    End Property

    Public Property PaperBin() As Integer
        Get
            PaperBin = oPrinter.PaperBin
        End Get
        Set(value As Integer)
            oPrinter.PaperBin = value
        End Set
    End Property

    Public Property PaperSize() As Integer
        Get
            PaperSize = oPrinter.PaperSize
        End Get
        Set(value As Integer)
            oPrinter.PaperSize = value
        End Set
    End Property

    Public ReadOnly Property PDFSupportFolder(Optional ByVal WithTrailingBS As Boolean = True) As String
        Get
            PDFSupportFolder = CleanPath(PDFFontsFolder, , False)
            If Not WithTrailingBS Then PDFSupportFolder = Left(PDFSupportFolder, Len(PDFSupportFolder) - 1)
        End Get
    End Property

    Public ReadOnly Property Port() As String
        Get
            'Port = Printer.Port
        End Get
    End Property

    Public Property TrackDefault() As Boolean
        Get
            TrackDefault = oPrinter.TrackDefault
        End Get
        Set(value As Boolean)
            oPrinter.TrackDefault = value
        End Set
    End Property

    Public ReadOnly Property TwipsPerPixelX() As Single
        Get
            TwipsPerPixelX = oPrinter.TwipsPerPixelX
        End Get
    End Property

    Public ReadOnly Property TwipsPerPixelY() As Single
        Get
            TwipsPerPixelY = oPrinter.TwipsPerPixelY
        End Get
    End Property

    Private ReadOnly Property ToDesktop() As Boolean
        Get
            ToDesktop = False
        End Get
    End Property

    Public ReadOnly Property Keywords() As String
        Get
            Keywords = DocKeywords & IIf(Len(DocKeywords) = 0, "", ",") & ProgramName & ",report,reports,archive,archived report," & CompanyName
        End Get
    End Property

    Public Property Width() As Integer
        Get
            Width = oPrinter.Width
        End Get
        Set(value As Integer)
            oPrinter.Width = value
        End Set
    End Property

    Public Property Zoom() As Integer
        Get
            Zoom = oPrinter.Zoom
        End Get
        Set(value As Integer)
            oPrinter.Zoom = value
        End Set
    End Property

    Public ReadOnly Property DocFile() As String
        Get
            DocFile = mDocFile
        End Get
    End Property

    Public Sub Circle_(ByVal X1 As Integer, ByVal X2 As Integer, ByVal Radius As String, Optional ByVal Color As Integer = 1, Optional ByVal cStart As Single = 0, Optional ByVal cEnd As Single = PI_2, Optional ByVal Aspect As Single = 1.0#)
        ' probably never used.
        If BuildDLS Then
        Else
            'oPrinter.Circle(X1, X2), Radius, Color, cStart, cEnd, Aspect
            If Color = 1 Then
                Color = Drawing.Color.Black.ToArgb
            End If
            oPrinter.Circle(X1, X2, Radius, Color, cStart, cEnd, Aspect)
        End If
    End Sub

    Public Sub EndDoc()
        On Error GoTo Failure
        TrackUsage(DocTitle)

        If BuildDLS Then
            DYMOObject.EndPrintJob
        ElseIf Preview Then
            frmPrintPreviewDocument.DataEnd()
        Else
            oPrinter.EndDoc
        End If

        If BuildPDF Then
            PDFFinish()
            DocTitle = ""
            DocKeywords = ""
            mBuildPDF = False
        End If
        Exit Sub

Failure:
        MessageBox.Show("Error in EndDoc: " & Err.Description)
    End Sub

    Private Sub PDFFinish()
        PDFPrinter.PDFEndDoc()
    End Sub

    Public Sub DataEnd()
        EndDoc()
    End Sub

    Public Sub Line_(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
        If BuildPDF Then
            PDFDraw_Setup()
            PDFPrinter.PDFDrawLine(PDFScaleX(X1), PDFScaleY(Y1), PDFScaleX(X2), PDFScaleX(Y2))
        End If
        oPrinter.Line(X1, Y1, X2, Y2)
    End Sub

    Private Function PDFDraw_Setup()
        PDFPrinter.PDFSetLineStyle = PDFStyleLgn.pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = DrawWidth
    End Function

    Private Sub PDFInit()
        If Not PDFSupportFolderExists Then Exit Sub

        mBuildPDF = True

        If ToDesktop Then
            mDocFile = UIOutputFolder() & "Report-" & DateTimeStamp() & ".pdf"
        Else
            mDocFile = ReportsFolder(Replace(DocTitle, " ", "")) & "Report-" & DateTimeStamp() & ".pdf"
        End If

        PDFPrinter = New clsPDFPrinter

        PDFPrinter.PDFTitle = DocTitle
        PDFPrinter.PDFAuthor = StoreSettings.Name & " - " & ProgramName
        PDFPrinter.PDFSubject = DocTitle & " - Archived Report"
        PDFPrinter.PDFCreator = SoftwareVersion(True, True, True)
        PDFPrinter.PDFProducer = SoftwareVersion(True, False, True)
        PDFPrinter.PDFKeywords = Keywords
        PDFPrinter.PDFView = False ' do not open the PDF file automatically

        PDFPrinter.PDFFileName = DocFile

        PDFPrinter.PDFLoadAfm = PDFSupportFolder(False)
        PDFPrinter.PDFConfirm = False
        PDFPrinter.PDFView = True
        'PDFPrinter.PDFFiligran = "P D F P r i n t e r   D e m o"

        'PDFPrinter.PDFSetViewerPreferences = VIEW_FITWINDOW
        PDFPrinter.PDFFormatPage = PDFFormatPgStr.FORMAT_LETTER
        PDFPrinter.PDFOrientation = PDFOrientationStr.ORIENT_PORTRAIT
        PDFPrinter.PDFSetUnit = PDFUnitStr.UNIT_PT
        PDFPrinter.PDFSetZoomMode = PDFZoomMd.ZOOM_REAL
        PDFPrinter.PDFSetLayoutMode = PDFLayoutMd.LAYOUT_DEFAULT
        PDFPrinter.PDFUseOutlines = False
        PDFPrinter.PDFUseThumbs = True

        PDFPrinter.PDFBeginDoc()

        PDFPrinter.PDFSetBookmark("Signet 1", 0, 40)
        PDFPrinter.PDFSetBookmark("Sous-Signet 2", 1, 60)

        PDFPrinter.PDFSetLineStyle = PDFStyleLgn.pPDF_SOLID
        PDFPrinter.PDFSetLineWidth = 1
    End Sub

    Public Property FillColor() As Color
        Get
            FillColor = oPrinter.FillColor
        End Get
        Set(value As Color)
            oPrinter.FillColor = value
        End Set
    End Property

    Public Property FillStyle() As Integer
        Get
            FillStyle = oPrinter.FillStyle
        End Get
        Set(value As Integer)
            oPrinter.FillStyle = value
        End Set
    End Property

    Public Property fontname() As String
        Get
            On Error Resume Next
            fontname = oPrinter.FontName
        End Get
        Set(value As String)
            On Error Resume Next
            oPrinter.FontName = value
        End Set
    End Property

    Public ReadOnly Property FontCount() As Integer
        Get
            On Error Resume Next
            FontCount = oPrinter.FontCount
        End Get
    End Property

    Public Property CurrentX() As Single
        Get
            On Error Resume Next
            CurrentX = oPrinter.CurrentX
        End Get
        Set(value As Single)
            On Error Resume Next
            oPrinter.CurrentX = value
        End Set
    End Property

    Public Property ColorMode() As Integer
        Get
            On Error Resume Next
            ColorMode = oPrinter.ColorMode
        End Get
        Set(value As Integer)
            On Error Resume Next
            oPrinter.ColorMode = value
        End Set
    End Property

    Public Property Copies() As Integer
        Get
            On Error Resume Next
            ColorMode = oPrinter.ColorMode
        End Get
        Set(value As Integer)
            On Error Resume Next
            oPrinter.ColorMode = value
        End Set
    End Property

    Public Property CurrentY() As Single
        Get
            On Error Resume Next
            CurrentY = oPrinter.CurrentY
        End Get
        Set(value As Single)
            On Error Resume Next
            oPrinter.CurrentY = value
        End Set
    End Property

    Public Property FontSize() As Single
        Get
            On Error Resume Next
            FontSize = oPrinter.FontSize
        End Get
        Set(value As Single)
            On Error Resume Next
            oPrinter.FontSize = value
        End Set
    End Property

    Public Property FontStrikeThru() As Boolean
        Get
            On Error Resume Next
            FontStrikeThru = oPrinter.FontStrikeThru
        End Get
        Set(value As Boolean)
            On Error Resume Next
            oPrinter.FontStrikeThru = value
        End Set
    End Property

    Public Property FontTransparent() As Boolean
        Get
            On Error Resume Next
            FontTransparent = oPrinter.FontTransparent
        End Get
        Set(value As Boolean)
            On Error Resume Next
            oPrinter.FontTransparent = value
        End Set
    End Property

    Public ReadOnly Property Fonts(ByVal Index As Integer) As String
        Get
            Fonts = oPrinter.Fonts(Index)
        End Get
    End Property

    Public Sub PrintNNL(ByVal Str1 As String, Optional ByVal Str2 As String = "", Optional ByVal Str3 As String = "", Optional ByVal Str4 As String = "", Optional ByVal Str5 As String = "", Optional ByVal Str6 As String = "", Optional ByVal Str7 As String = "", Optional ByVal Str8 As String = "", Optional ByVal Str9 As String = "", Optional ByVal Str10 As String = "")
        Print_(Str1, Str2, Str3, Str4, Str5, Str6, Str7, Str8, Str9, Str10)
    End Sub

    Public Sub PrintNL(Optional ByVal Str1 As String = "", Optional ByVal Str2 As String = "", Optional ByVal Str3 As String = "", Optional ByVal Str4 As String = "", Optional ByVal Str5 As String = "", Optional ByVal Str6 As String = "", Optional ByVal Str7 As String = "", Optional ByVal Str8 As String = "", Optional ByVal Str9 As String = "", Optional ByVal Str10 As String = "")
        ' let the other print routine handle it to standardize all printing
        On Error GoTo NoPrint
        Print_(Str1, Str2, Str3, Str4, Str5, Str6, Str7, Str8, Str9, Str10)
        oPrinter.Print
        Exit Sub
NoPrint:
        CheckStandardErrors()
    End Sub

    Private Sub Print_(ByVal Str1 As String, Optional ByVal Str2 As String = "", Optional ByVal Str3 As String = "", Optional ByVal Str4 As String = "", Optional ByVal Str5 As String = "", Optional ByVal Str6 As String = "", Optional ByVal Str7 As String = "", Optional ByVal Str8 As String = "", Optional ByVal Str9 As String = "", Optional ByVal Str10 As String = "")
        If BuildDLS Then
            '    DYMOObject.PrintObject.AddObject
        ElseIf BuildPDF Then
            PDFOutText_Setup()
            PDFPrinter.PDFTextOut(Str1 & Str2 & Str3 & Str4 & Str5 & Str6 & Str7 & Str8 & Str9 & Str10, PDFScaleX(CurrentX), PDFScaleY(CurrentY))
            ' printer handles location always.  PDF always mirrors. Newline is therefore irrelevant.
        End If
        On Error GoTo NoPrint

        Dim CY As Integer
        CY = CurrentY
        'oPrinter.Print(Str1)
        oPrinter.Print(Str1 & " " & Str2 & " " & Str3 & " " & Str4 & Str5 & " " & Str6 & " " & Str7 & " " & Str8 & Str9 & " " & Str10)
        'CurrentY = CY
        'If Str2 <> "" Then oPrinter.Print(Str2)
        'If Str3 <> "" Then oPrinter.Print(Str3)
        'If Str4 <> "" Then oPrinter.Print(Str4)
        'If Str5 <> "" Then oPrinter.Print(Str5)
        'If Str6 <> "" Then oPrinter.Print(Str6)
        'If Str7 <> "" Then oPrinter.Print(Str7)
        'If Str8 <> "" Then oPrinter.Print(Str8)
        'If Str9 <> "" Then oPrinter.Print(Str9)
        'If Str10 <> "" Then oPrinter.Print(Str10)

        Exit Sub
NoPrint:
        CheckStandardErrors()
    End Sub

    Public ReadOnly Property BuildPDF() As Boolean
        Get
            BuildPDF = mBuildPDF
        End Get
    End Property

    Private Function PDFOutText_Setup()
        Dim FID As Integer, FFlags As Integer

        FFlags = 0
        If FontBold Then FFlags = FFlags + PDFFontStl.FONT_BOLD
        If FontItalic Then FFlags = FFlags + PDFFontStl.FONT_ITALIC
        If FontUnderline Then FFlags = FFlags + PDFFontStl.FONT_UNDERLINE
        '  If FontStrikethru Then FFlags = FFlags + FONT_STRIKETHRU

        '    FONT_ARIAL = 0
        '    FONT_COURIER = 1
        '    FONT_TIMES = 2
        '    FONT_SYMBOL = 3
        '    FONT_ZAPFDINGBATS = 4
        Select Case LCase(fontname)
            Case "times new roman" : FID = PDFFontNme.FONT_TIMES
            Case "courier new" : FID = PDFFontNme.FONT_COURIER
            Case "symbol" : FID = PDFFontNme.FONT_SYMBOL
            Case Else : FID = PDFFontNme.FONT_ARIAL
        End Select

        PDFPrinter.PDFSetFont(FID, FontSize, FFlags)

        'PDFPrinter.PDFSetTextColor = ForeColor
        'PDFPrinter.PDFSetTextColor(ForeColor)
    End Function

    Public Property ForeColor() As Color
        Get
            ForeColor = oPrinter.ForeColor
        End Get
        Set(value As Color)
            oPrinter.ForeColor = value
        End Set
    End Property

    Private Function PDFScaleX(ByVal X As Single) As Single
        GetPDFDimensions()
        PDFScaleX = X / Printer.ScaleWidth * pW
    End Function

    Private Function PDFScaleY(ByVal Y As Single) As Single
        GetPDFDimensions()
        PDFScaleY = Y / Printer.ScaleHeight * pH + 10
    End Function

    Private Function GetPDFDimensions() As Boolean
        GetPDFDimensions = True
        If pH <> 0 And pW <> 0 Then Exit Function
        pH = PDFPrinter.PDFGetPageHeight
        pW = PDFPrinter.PDFGetPageWidth
    End Function

    Public Property FontBold() As Boolean
        Get
            FontBold = oPrinter.FontBold
        End Get
        Set(value As Boolean)
            oPrinter.FontBold = value
        End Set
    End Property

    Public Property FontItalic() As Boolean
        Get
            On Error Resume Next
            FontItalic = oPrinter.FontItalic
        End Get
        Set(value As Boolean)
            On Error Resume Next
            oPrinter.FontItalic = value
        End Set
    End Property

    Public Property FontUnderline() As Boolean
        Get
            On Error Resume Next
            FontUnderline = oPrinter.FontUnderline
        End Get
        Set(value As Boolean)
            On Error Resume Next
            oPrinter.FontUnderline = value
        End Set
    End Property

    Public Property Font() As Font
        Get
            Font = oPrinter.Font
        End Get
        Set(value As Font)
            oPrinter.Font = value
        End Set
    End Property

    Public Sub LineStep(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal W As Integer, ByVal H As Integer)
        Line_(X1, Y1, X1 + W, Y1 + H)
    End Sub

    Public Sub Box(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double)
        oPrinter.Line(X1, Y1, X2, Y2,, True)

        If BuildPDF Then
            PDFDraw_Setup()
            PDFPrinter.PDFDrawRectangle(X1, X2, X2 - X1, Y2 - Y1)
        End If
    End Sub

    Public Sub BoxStep(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal W As Integer, ByVal H As Integer)
        Box(X1, Y1, X1 + W, Y1 + H)
    End Sub

    Public Sub KillDoc()
        If Not Preview Then
            oPrinter.KillDoc
        End If

        If BuildPDF Then CancelPrintToPDF()
    End Sub

    Public Sub CancelPrintToPDF()
        mBuildPDF = False
        mDocTitle = ""
        mDocFile = ""
        mDocKeywords = ""
        PDFPrinter = Nothing
    End Sub

    Public Sub NewPage()
        If Preview Then
            frmPrintPreviewDocument.NewPage()
        Else
            oPrinter.NewPage
            PageNumber = PageNumber + 1
            '    TotalPages = TotalPages + 1
        End If
        If BuildPDF Then
            PDFPrinter.PDFEndPage()
            PDFPrinter.PDFNewPage()
        End If
    End Sub

    Public Sub PaintPicture(ByVal vPic As IPictureDisp, Optional ByVal X1 As Single = 0, Optional ByVal Y1 As Single = 0, Optional ByVal Width1 As Integer = 0, Optional ByVal Height1 As Integer = 0, Optional ByVal X2 As Integer = 0, Optional ByVal Y2 As Integer = 0, Optional ByVal Width2 As Integer = 0, Optional ByVal Height2 As Integer = 0, Optional ByVal OpCode As Integer = 0)
        '  Printer.PaintPicture
        'If IsMissing(Width1) Then
        If Width1 = 0 Then
            oPrinter.PaintPicture(vPic, X1, Y1)
            'ElseIf IsMissing(Height1) Then
        ElseIf Height1 = 0 Then
            oPrinter.PaintPicture(vPic, X1, Y1, Width1)
            'ElseIf IsMissing(X2) Then
        ElseIf X2 = 0 Then
            oPrinter.PaintPicture(vPic, X1, Y1, Width1, Height1)
            'ElseIf IsMissing(Y2) Then
        ElseIf Y2 = 0 Then
            oPrinter.PaintPicture(vPic, X1, Y1, Width1, Height1, X2)
            'ElseIf IsMissing(Width2) Then
        ElseIf Width2 = 0 Then
            oPrinter.PaintPicture(vPic, X1, Y1, Width1, Height1, X2, Y2)
            'ElseIf IsMissing(Height2) Then
        ElseIf Height2 = 0 Then
            oPrinter.PaintPicture(vPic, X1, Y1, Width1, Height1, X2, Y2, Width2)
            'ElseIf IsMissing(OpCode) Then
        ElseIf OpCode = 0 Then
            oPrinter.PaintPicture(vPic, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2)
        Else
            oPrinter.PaintPicture(vPic, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, OpCode)
        End If
    End Sub

    Public Sub PSet_(ByVal Stepp As Integer, ByVal X As Single, ByVal Y As Single, ByVal Color As Integer)
        '### probably implemented wrong
        oPrinter.PSet(X, Stepp, Color)
    End Sub

    Public Sub Scale_(ByVal Flags As Integer, Optional ByVal X1 As Integer = 0, Optional ByVal Y1 As Integer = 0, Optional ByVal X2 As Integer = 0, Optional ByVal Y2 As Integer = 0)
        oPrinter.Scale(X1, Y1, X2, Y2)
    End Sub

    Public Function ScaleX(ByVal Width As Single, Optional ByVal FromScale As Integer = 0, Optional ByVal ToScale As Integer = 0) As Single
        ScaleX = oPrinter.ScaleX(Width, FromScale, ToScale)
    End Function

    Public Function ScaleY(ByVal Height As Single, Optional ByVal FromScale As Integer = 0, Optional ByVal ToScale As Integer = 0) As Single
        ScaleY = oPrinter.ScaleY(Height, FromScale, ToScale)
    End Function

    Public Function TextHeight(ByVal Str As String) As Single
        TextHeight = oPrinter.TextHeight(Str)
    End Function

    Public Function TextWidth(ByVal Str As String) As Single
        TextWidth = oPrinter.TextWidth(Str)
    End Function

    Public Sub SetPreview(Optional ByVal vReportTitle As String = "", Optional ByVal vKeywords As String = "", Optional ByRef vCallingForm As Form = Nothing)
        If vReportTitle <> "" Then DocTitle = vReportTitle
        DocKeywords = vKeywords
        OutputToPrinter = False

        'Load frmPrintPreviewMain
        frmPrintPreviewDocument.CallingForm = vCallingForm

        OutputObject = Me

        PreviewImage = frmPrintPreviewDocument.picPicture
        '  Set frmPrintPreviewDocument.CallingForm = CallingForm
        frmPrintPreviewDocument.ReportName = DocTitle
    End Sub

End Class

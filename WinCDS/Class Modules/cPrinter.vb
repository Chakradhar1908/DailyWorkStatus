Imports WinCDS.clsPDFPrinter
Public Class cPrinter
    Private Const PI As Double = 3.14159265358979
    Private Const PI_2 As Double = 6.28318530717958
    Private mBuildPDF As Boolean
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

    Public ReadOnly Property Preview() As Boolean
        Get
            Preview = Not (mPreviewImage Is Nothing)
        End Get
    End Property

    Public ReadOnly Property PreviewImage() As Object
        Get
            PreviewImage = mPreviewImage
        End Get
    End Property

    Public ReadOnly Property BuildDLS() As Boolean
        Get
            BuildDLS = False And IsDymo And HasDLS
        End Get
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

    Public ReadOnly Property HasDLS() As Boolean
        Get
            Static vValue As TriState
            On Error Resume Next
            If vValue = vbFalse Then vValue = IIf(IsNotNothing(CreateObject("DYMO.LabelEngine")), vbTrue, vbUseDefault)
            HasDLS = (vValue = vbTrue)
        End Get
    End Property

    Public Sub SetPrintToPDF(Optional ByVal vDocTitle As String = "", Optional ByVal vKeywords As String = "")
        If vDocTitle <> "" Then DocTitle = vDocTitle
        DocKeywords = vKeywords
        OutputToPrinter = True

        PDFInit()

        OutputObject = Me
    End Sub

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

    Public ReadOnly Property PDFSupportFolder(Optional ByVal WithTrailingBS As Boolean = True) As String
        Get
            PDFSupportFolder = CleanPath(PDFFontsFolder, , False)
            If Not WithTrailingBS Then PDFSupportFolder = Left(PDFSupportFolder, Len(PDFSupportFolder) - 1)
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

    Public ReadOnly Property DocFile() As String
        Get
            DocFile = mDocFile
        End Get
    End Property

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
        oPrinter.Print(Str1)
        If Str2 <> "" Then oPrinter.Print(Str2)
        If Str3 <> "" Then oPrinter.Print(Str3)
        If Str4 <> "" Then oPrinter.Print(Str4)
        If Str5 <> "" Then oPrinter.Print(Str5)
        If Str6 <> "" Then oPrinter.Print(Str6)
        If Str7 <> "" Then oPrinter.Print(Str7)
        If Str8 <> "" Then oPrinter.Print(Str8)
        If Str9 <> "" Then oPrinter.Print(Str9)
        If Str10 <> "" Then oPrinter.Print(Str10)

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
        PDFPrinter.PDFSetTextColor(ForeColor)
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
        GetPDFDimensions
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
End Class

Public Class clsPDFPrinter
    Private FTitle As String
    Private FAuthor As String
    Private FSubject As String
    Private FCreator As String
    Private FProducer As String
    Private FFileCompress As Boolean
    Private FOrientation As String
    Private FKeywords As String

    Private boPDFUnderline As Boolean
    Private boPDFItalic As Boolean
    Private boPDFBold As Boolean
    Private boPDFConfirm As Boolean
    Private boPDFView As Boolean
    Private PDFboThumbs As Boolean
    Private PDFboOutlines As Boolean
    Private PDFboImage As Boolean

    Private FFileName As String
    Private FPageNumber As Integer
    Private FPageLink As Integer

    Private Fso As Object
    Private Strm As Object
    Private sPDFName As String
    Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal szClass$, ByVal szTitle$) As Integer
    Private Declare Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Object) As Integer
    Private Const WM_CLOSE As Integer = &H10
    Private Arr_Font() As Object
    Private wsPathConfig As String
    Private wsPathAdobe As String
    Private PDFCanvasWidth() As Object
    Private PDFCanvasHeight() As Object
    Private PDFCanvasOrientation() As Object
    Private In_offset As Integer
    Private in_FontNum As Integer
    Private in_PagesNum As Integer
    Private in_Ech As Double
    Private in_Canvas As Integer
    Private iWidthStr As Double
    Private PDFFontName As String
    Enum PDFFormatPgStr
        FORMAT_A4 = 0
        FORMAT_A3 = 1
        FORMAT_A5 = 2
        FORMAT_LETTER = 3
        FORMAT_LEGAL = 4
    End Enum
    Enum PDFOrientationStr
        ORIENT_PAYSAGE = 0
        ORIENT_PORTRAIT = 1
    End Enum
    Enum PDFUnitStr
        UNIT_PT = 0
        UNIT_MM = 1
        UNIT_CM = 2
    End Enum
    Enum PDFZoomMd
        ZOOM_FULLPAGE = 0
        ZOOM_FULLWIDTH = 1
        ZOOM_REAL = 2
        ZOOM_DEFAULT = 3
    End Enum
    Enum PDFLayoutMd
        LAYOUT_SINGLE = 0
        LAYOUT_CONTINOUS = 1
        LAYOUT_TWO = 2
        LAYOUT_DEFAULT = 3
    End Enum
    Enum PDFStyleLgn
        pPDF_SOLID = 0
        pPDF_DASH = 1
        pPDF_DASHDOT = 2
        pPDF_DASHDOTDOT = 3
    End Enum
    Enum PDFFontStl
        FONT_NORMAL = 0
        FONT_ITALIC = 1
        FONT_BOLD = 2
        FONT_UNDERLINE = 3
    End Enum
    Enum PDFFontNme
        FONT_ARIAL = 0
        FONT_COURIER = 1
        FONT_TIMES = 2
        FONT_SYMBOL = 3
        FONT_ZAPFDINGBATS = 4
    End Enum
    Private Structure PDFRGB
        Dim in_R As Integer
        Dim in_G As Integer
        Dim in_B As Integer
    End Structure

    Private Const mjwPDF As String = "1.3"
    Private Const mjwPDFVersion As String = "mjwPDF 1.0"

    Private PDFZoomMode As Object
    Private PDFLayoutMode As Object
    Private PDFViewerPref As Object
    Private bPDFViewerPref As Boolean
    Private bPDFWatermark As Boolean
    Private sPDFWatermark As String
    Private ParentNum As Object, ContentNum As Object, ResourceNum As Object, FontNum As Object, CatalogNum As Object, FontNumber As Object, CurrentPDFSetPageObject As Object, NumberofImages As Object, iOutlineRoot As Integer
    Private CurrentObjectNum As Integer
    Private ObjectOffset As Integer
    'Private ObjectOffsetList As Object
    Private ObjectOffsetList() As Object
    'Private PageNumberList As Object
    Private PageNumberList() As Object
    'Private PageLinksList(1 To 1000, 1 To 1000) As Object  '@NO-LINT
    Private PageLinksList(0 To 999, 0 To 999) As Object  '@NO-LINT
    'Private LinksList As Object
    Private LinksList() As Object
    'Private PageCanvasWidth As Object
    Private PageCanvasWidth() As Object
    'Private PageCanvasHeight As Object
    Private PageCanvasHeight() As Object
    'Private FontNumberList As Object
    Private FontNumberList() As Object
    Private CRCounter As Integer

    Private ColorSpace As String
    Private ColorCount As Byte
    Private ImageStream As String
    Private TempStream As String
    Private pTempStream As String
    Private sTempStream As String
    Private cTempStream As String
    Private dTempStream As String
    'Private boPageLinksList As Object
    Private boPageLinksList() As Object
    'Private NbPageLinksList As Object
    Private NbPageLinksList() As Object
    Private StreamSize1 As Integer, StreamSize2 As Integer
    Private in_xCurrent As Double
    Private in_yCurrent As Double
    Private aOutlines() As oOutlines
    Private iOutlines As Integer
    Private aPage() As Object
    Private PDFFontSize As Integer
    Private PDFFontNum As Integer

    Private Structure oOutlines
        Dim sText As String
        Dim iLevel As Integer
        Dim yPos As Double
        Dim iPageNb As Integer
        Dim bPrev As Boolean
        Dim bNext As Boolean
        Dim bFirst As Boolean
        Dim bLast As Boolean
        Dim iFirst As Integer
        Dim iNext As Integer
        Dim iPrev As Integer
        Dim iLast As Integer
        Dim iParent As Integer
    End Structure
    Private PDFLnStyle As String
    Private PDFLnWidth As Double
    Private bScanAdobe As Boolean
    Private PDFMargin As Integer
    Private PDFcMargin As Integer ' Center Margin
    Private PDFlMargin As Integer ' Left Margin
    Private PDFtMargin As Integer ' Top Margin
    Private PDFLineColor As String
    Private PDFDrawColor As String
    Private PDFTextColor As String
    Private PDFrMargin As Integer ' Right Margin
    Private PDFbMargin As Integer ' Bottom Margin
    Private PDFAngle As Double
    Private bAngle As Double
    Private str_TmpFont As String

    Public Sub New()
        PDFInit()
    End Sub

    Public Sub PDFInit()

        bScanAdobe = False
        Fso = CreateObject("scripting.filesystemobject")

        'If wsPathConfig = "" Then wsPathConfig = App.Path
        If wsPathConfig = "" Then wsPathConfig = My.Application.Info.DirectoryPath
        Dim Position As Integer
        Position = InStr(wsPathConfig, "bin")
        wsPathConfig = Mid(wsPathConfig, 1, Position - 2) 'c:\wincds\wincds\bin
        'If wsPathConfig = "" Then wsPathConfig = Application.StartupPath
        'If wsPathConfig = "" Then wsPathConfig = Assembly.GetExecutingAssembly.Location
        'wsPathConfig = Path.GetDirectoryName(wsPathConfig)
        PDFLoadAfm = wsPathConfig

        'ObjectOffsetList = Array()
        'PageNumberList = Array()
        'PageCanvasWidth = Array()
        'PageCanvasHeight = Array()

        'boPageLinksList = Array()
        'NbPageLinksList = Array()
        'LinksList = Array()

        'FontNumberList = Array()

        In_offset = 1
        in_FontNum = 1
        in_PagesNum = 1
        in_Canvas = 1
        FPageLink = 0

        boPDFUnderline = False
        boPDFBold = False
        boPDFItalic = False

        ' Unité de mesure par défaut : cm
        in_Ech = 72 / 2.54

        ' Marges de la page (1 cm)
        PDFMargin = in_Ech / 28.35
        PDFSetMargins(PDFMargin, PDFMargin)

        ' Marge interieure des cellules (1 mm)
        PDFcMargin = in_Ech * (PDFMargin / 10)

        ' Largeur de ligne (0.2 mm)
        PDFLnWidth = 0.567

        in_xCurrent = PDFlMargin
        in_yCurrent = PDFtMargin

        TempStream = ""
        ImageStream = ""
        pTempStream = ""
        sTempStream = ""
        cTempStream = ""
        dTempStream = ""

        FontNum = 1

        ' Définition dzes couleurs par défaut
        PDFLineColor = "0 G"
        PDFDrawColor = "0 g"
        PDFTextColor = "0 g"

        ' Format d'orientation de page par défaut : A4
        'ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
        ReDim Preserve PDFCanvasWidth(0 To in_Canvas - 1)
        'ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
        ReDim Preserve PDFCanvasHeight(0 To in_Canvas - 1)
        'ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)
        ReDim Preserve PDFCanvasOrientation(0 To in_Canvas - 1)

        PDFCanvasWidth(in_Canvas - 1) = 595.28
        PDFCanvasHeight(in_Canvas - 1) = 841.89
        PDFCanvasOrientation(in_Canvas - 1) = "p"

        FProducer = ""
        FAuthor = ""
        FCreator = ""
        FKeywords = ""
        FSubject = ""
        Exit Sub
    End Sub

    Public Sub PDFSetMargins(In_left As Integer, In_top As Integer, Optional In_right As Integer = -1, Optional In_bottom As Integer = -1)
        'Attribute PDFSetMargins.VB_HelpID = 2044

        PDFlMargin = In_left
        PDFtMargin = In_top

        If In_right = -1 Then In_right = In_left
        If In_bottom = -1 Then In_bottom = In_top

        PDFrMargin = In_right
        PDFbMargin = In_bottom
    End Sub

    Public Property PDFTitle() As String
        Get

        End Get
        Set(value As String)
            'Attribute PDFTitle.VB_HelpID = 2027
            FTitle = value
        End Set
    End Property

    Public Property PDFAuthor() As String
        Get

        End Get
        Set(value As String)
            'Attribute PDFAuthor.VB_HelpID = 2025
            FAuthor = value
        End Set
    End Property

    Public Property PDFSubject() As String
        Get

        End Get
        Set(value As String)
            'Attribute PDFSubject.VB_HelpID = 2022
            FSubject = value
        End Set
    End Property

    Public Property PDFCreator() As String
        Get

        End Get
        Set(value As String)
            'Attribute PDFCreator.VB_HelpID = 2024
            FCreator = value
        End Set
    End Property

    Public Property PDFProducer() As String
        Get

        End Get
        Set(value As String)
            'Attribute PDFProducer.VB_HelpID = 2021
            FProducer = value
        End Set
    End Property

    Public Property PDFKeywords() As String
        Get

        End Get
        Set(value As String)
            'Attribute PDFKeywords.VB_HelpID = 2023
            FKeywords = value
        End Set
    End Property

    Public Property PDFView() As Boolean
        Get

        End Get
        Set(value As Boolean)
            boPDFView = value
        End Set
    End Property

    Public Property PDFFileName() As String
        Get

        End Get
        Set(value As String)
            'Attribute PDFFileName.VB_HelpID = 2028

            Dim Items() As String
            Dim sFilePath As String
            Dim sFileName As String
            Dim hwnd As Integer
            Dim Retval As Integer
            Dim In_i As Integer

            On Error GoTo Err_File

            FFileName = value

            Items = Split(value, "\")
            If UBound(Items) = -1 Then Exit Property

            sFileName = Items(UBound(Items))
            sFilePath = Left(value, Len(value) - Len(Items(UBound(Items))))

            sPDFName = Fso.BuildPath(sFilePath, sFileName)
            Strm = Fso.CreateTextFile(sPDFName, True)

            Exit Property

Err_File:
            If Convert.ToInt32(Err()) = 70 Then
                hwnd = FindWindow(vbNullString, "Adobe Reader - [" & sFileName & "]")
                Retval = PostMessage(hwnd, WM_CLOSE, 0&, 0&)
                Sleep(17)

                Strm = Fso.CreateTextFile(sPDFName, True)
                Resume Next
            End If
        End Set
    End Property

    Public Property PDFLoadAfm() As String
        Get

        End Get
        Set(value As String)
            Dim Fso As Object
            Dim oRep As Object
            Dim oFiles As Object
            Dim in_Font As Integer

            Fso = CreateObject("Scripting.FileSystemObject")
            oRep = Fso.GetFolder(value)

            in_Font = -1
            For Each oFiles In oRep.Files
                If InStr(1, LCase(oFiles.Path), ".afm") <> 0 Then
                    in_Font = in_Font + 1
                    ReDim Preserve Arr_Font(0 To in_Font)
                    Arr_Font(in_Font) = Mid(oFiles.Name, 1, Len(oFiles.Name) - 4)
                End If
            Next

            If in_Font <> -1 Then wsPathConfig = value
        End Set
    End Property

    Public Property PDFConfirm() As Boolean
        Get

        End Get
        Set(value As Boolean)
            'Attribute PDFConfirm.VB_HelpID = 2029

            boPDFConfirm = value
        End Set
    End Property

    Public Property PDFFormatPage() As Object
        Get

        End Get
        Set(value As Object)
            'Attribute PDFFormatPage.VB_HelpID = 2018

            'ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
            ReDim Preserve PDFCanvasWidth(0 To in_Canvas - 1)
            'ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
            ReDim Preserve PDFCanvasHeight(0 To in_Canvas - 1)
            'ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)
            ReDim Preserve PDFCanvasOrientation(0 To in_Canvas - 1)

            Select Case TypeName(value)
                'Case "Long"
                Case "PDFFormatPgStr"
                    Select Case value
                        Case PDFFormatPgStr.FORMAT_A4
                            PDFCanvasWidth(in_Canvas - 1) = 595.28
                            PDFCanvasHeight(in_Canvas - 1) = 841.89
                        Case PDFFormatPgStr.FORMAT_A3
                            PDFCanvasWidth(in_Canvas - 1) = 841.89
                            PDFCanvasHeight(in_Canvas - 1) = 1190.55
                        Case PDFFormatPgStr.FORMAT_A5
                            PDFCanvasWidth(in_Canvas - 1) = 420.94
                            PDFCanvasHeight(in_Canvas - 1) = 595.28
                        Case PDFFormatPgStr.FORMAT_LETTER
                            PDFCanvasWidth(in_Canvas - 1) = 612
                            PDFCanvasHeight(in_Canvas - 1) = 792
                        Case PDFFormatPgStr.FORMAT_LEGAL
                            PDFCanvasWidth(in_Canvas - 1) = 612
                            PDFCanvasHeight(in_Canvas - 1) = 1008
                        Case Else
                            MessageBox.Show("Format page set incorrectly : " & value & "." & vbNewLine & "Format page set to A4.", "Format Page - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            PDFCanvasWidth(in_Canvas - 1) = 595.28
                            PDFCanvasHeight(in_Canvas - 1) = 841.89
                    End Select
                Case "Double()"
                    'PDFCanvasWidth(in_Canvas) = str_FormatPage(0)
                    PDFCanvasWidth(in_Canvas - 1) = value
                    'PDFCanvasHeight(in_Canvas) = str_FormatPage(1)
                    PDFCanvasHeight(in_Canvas - 1) = value
                Case Else
                    MessageBox.Show("Format page set incorrectly : " & value & "." & vbNewLine & "Format page set to A4", "Format Page - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    PDFCanvasWidth(in_Canvas - 1) = 595.28
                    PDFCanvasHeight(in_Canvas - 1) = 841.89
            End Select
        End Set
    End Property

    Public Property PDFOrientation() As PDFOrientationStr
        Get

        End Get
        Set(value As PDFOrientationStr)
            'Attribute PDFOrientation.VB_HelpID = 2017

            Dim tmp_PDFCanvasWidth As Integer
            Dim tmp_PDFCanvasHeight As Integer

            'ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
            ReDim Preserve PDFCanvasWidth(0 To in_Canvas - 1)
            'ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
            ReDim Preserve PDFCanvasHeight(0 To in_Canvas - 1)
            'ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)
            ReDim Preserve PDFCanvasOrientation(0 To in_Canvas - 1)

            tmp_PDFCanvasWidth = PDFCanvasWidth(in_Canvas - 1)
            tmp_PDFCanvasHeight = PDFCanvasHeight(in_Canvas - 1)

            Select Case value
                Case PDFOrientationStr.ORIENT_PORTRAIT
                    PDFCanvasWidth(in_Canvas - 1) = tmp_PDFCanvasWidth
                    PDFCanvasHeight(in_Canvas - 1) = tmp_PDFCanvasHeight
                    PDFCanvasOrientation(in_Canvas - 1) = "p"
                Case PDFOrientationStr.ORIENT_PAYSAGE
                    PDFCanvasWidth(in_Canvas - 1) = tmp_PDFCanvasHeight
                    PDFCanvasHeight(in_Canvas - 1) = tmp_PDFCanvasWidth
                    PDFCanvasOrientation(in_Canvas - 1) = "l"
                Case Else
                    MessageBox.Show("Orientation set incorrectly: " & value & "." & vbNewLine & "Orientation set to portrait.", "Error in orientation - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    PDFCanvasWidth(in_Canvas - 1) = tmp_PDFCanvasWidth
                    PDFCanvasHeight(in_Canvas - 1) = tmp_PDFCanvasHeight
                    PDFCanvasOrientation(in_Canvas - 1) = "p"
            End Select

            'ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
            ReDim Preserve PDFCanvasWidth(0 To in_Canvas - 1)
            'ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
            ReDim Preserve PDFCanvasHeight(0 To in_Canvas - 1)
            'ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)
            ReDim Preserve PDFCanvasOrientation(0 To in_Canvas - 1)
        End Set
    End Property

    Public Property PDFSetUnit() As PDFUnitStr
        Get

        End Get
        Set(value As PDFUnitStr)
            'Attribute PDFSetUnit.VB_HelpID = 2015
            Select Case value
                Case PDFUnitStr.UNIT_PT
                    in_Ech = 1
                Case PDFUnitStr.UNIT_MM
                    in_Ech = 72 / 25.4
                Case PDFUnitStr.UNIT_CM
                    in_Ech = 72 / 2.54
                Case Else
                    MessageBox.Show("Incorrect Unit of Measure : " & value & "." & vbNewLine & "Using centimeter ", "Error in measurement unit - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    in_Ech = 72 / 2.54
            End Select
        End Set
    End Property

    Public Property PDFSetZoomMode() As PDFZoomMd
        Get

        End Get
        Set(value As PDFZoomMd)
            'Attribute PDFSetZoomMode.VB_HelpID = 2009
            If value = PDFZoomMd.ZOOM_FULLPAGE Or value = PDFZoomMd.ZOOM_FULLWIDTH Or value = PDFZoomMd.ZOOM_REAL Or value = PDFZoomMd.ZOOM_DEFAULT Or
                (IsNumeric(value) And (value <> PDFZoomMd.ZOOM_FULLPAGE Or
                                            value <> PDFZoomMd.ZOOM_FULLWIDTH Or
                                            value <> PDFZoomMd.ZOOM_REAL Or
                                            value <> PDFZoomMd.ZOOM_DEFAULT)) Then
                If IsNumeric(value) Then
                    PDFZoomMode = Int(value)
                Else
                    PDFZoomMode = value
                End If
            Else
                MessageBox.Show("Incorrect Zoom Mode : " & value & "." & vbNewLine & "Focus will be set to full-page zoom", "Zoom Mode - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                PDFZoomMode = PDFZoomMd.ZOOM_FULLPAGE
            End If
        End Set
    End Property

    Public Property PDFSetLayoutMode() As PDFLayoutMd
        Get

        End Get
        Set(value As PDFLayoutMd)
            'Attribute PDFSetLayoutMode.VB_HelpID = 2013

            If value = PDFLayoutMd.LAYOUT_SINGLE Or value = PDFLayoutMd.LAYOUT_CONTINOUS Or value = PDFLayoutMd.LAYOUT_TWO Or value = PDFLayoutMd.LAYOUT_DEFAULT Then
                PDFLayoutMode = value
            Else
                MessageBox.Show("Layout incorrect : " & value & "." & vbNewLine & "Layout will be set to simple single page.", "Layout Mode - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                PDFLayoutMode = PDFLayoutMd.LAYOUT_SINGLE
            End If
        End Set
    End Property

    Public Property PDFUseOutlines() As Boolean
        Get

        End Get
        Set(value As Boolean)
            'Attribute PDFUseOutlines.VB_HelpID = 2012

            PDFboOutlines = value
        End Set
    End Property

    Public Property PDFUseThumbs() As Boolean
        Get

        End Get
        Set(value As Boolean)
            'Attribute PDFUseThumbs.VB_HelpID = 2011
            PDFboThumbs = value
        End Set
    End Property

    Public Sub PDFBeginDoc()

        FPageNumber = 1

        In_offset = 1

        NumberofImages = 0
        CurrentObjectNum = 0
        ObjectOffset = 0
        CurrentPDFSetPageObject = 0
        CRCounter = 0
        FontNumber = 0

        'ReDim ObjectOffsetList(1 To 1)
        ReDim ObjectOffsetList(0)
        'ReDim PageNumberList(1 to 1)
        ReDim PageNumberList(0)
        'ReDim PageCanvasHeight(1 to 1)
        ReDim PageCanvasHeight(0)
        'ReDim PageCanvasWidth(1 To 1)
        ReDim PageCanvasWidth(0)

        'ReDim boPageLinksList(1 To 1)
        ReDim boPageLinksList(0)
        'ReDim NbPageLinksList(1 To 1)
        ReDim NbPageLinksList(0)
        'ReDim LinksList(1 To 1)
        ReDim LinksList(0)
        'ReDim FontNumberList(1 To 1)
        ReDim FontNumberList(0)

        TempStream = ""
        ImageStream = ""

        PDFSetHeader()
        PDFSetDocInfo()
        PDFStartStream()
    End Sub

    Private Sub PDFSetHeader()
        'Attribute PDFSetHeader.VB_HelpID = 2080

        CurrentObjectNum = 0

        Strm.WriteLine("%PDF-" & mjwPDF)
        PDFAddToOffset(Len("%PDF-" & mjwPDF))
    End Sub

    Private Sub PDFAddToOffset(ByRef Offset As Integer)
        'Attribute PDFAddToOffset.VB_HelpID = 2096

        'ReDim Preserve ObjectOffsetList(1 To In_offset)
        ReDim Preserve ObjectOffsetList(0 To In_offset - 1)

        ObjectOffset = ObjectOffset + Offset
        ObjectOffsetList(In_offset - 1) = ObjectOffset

        In_offset = In_offset + 1

        CRCounter = 0
    End Sub

    Private Sub PDFSetDocInfo()
        'Attribute PDFSetDocInfo.VB_HelpID = 2081

        CurrentObjectNum = CurrentObjectNum + 1
        TempStream = ""

        PDFOutStream(sTempStream, "%DEBUT_OBJ/%")
        PDFOutStream(TempStream, CurrentObjectNum & " 0 obj")
        PDFOutStream(TempStream, "<<")
        PDFOutStream(TempStream, "/Producer (" + FProducer + ")")
        PDFOutStream(TempStream, "/Author (" + FAuthor + ")")
        PDFOutStream(TempStream, "/CreationDate (D:" + Format(Now, "YYYYMMDDHHmmSS") + ")")
        PDFOutStream(TempStream, "/Creator (" + FCreator + ")")
        PDFOutStream(TempStream, "/Keywords (" + FKeywords + ")")
        PDFOutStream(TempStream, "/Subject (" + FSubject + ")")
        PDFOutStream(TempStream, "/Title (" + FTitle + ")")
        PDFOutStream(TempStream, "/ModDate ()")
        PDFOutStream(TempStream, ">>")
        PDFOutStream(TempStream, "endobj")
        PDFOutStream(sTempStream, "%FIN_OBJ/%")

        PDFAddToOffset(Len(TempStream))
        Strm.WriteLine(TempStream)
    End Sub

    Private Sub PDFOutStream(ByRef Ms As String, ByRef S As String)
        'Attribute PDFOutStream.VB_HelpID = 2095
        CRCounter = CRCounter + 2
        Ms = Ms & S & vbCrLf
    End Sub

    Private Sub PDFStartStream()
        'Attribute PDFStartStream.VB_HelpID = 2088
        ContentNum = CurrentObjectNum
        CurrentObjectNum = CurrentObjectNum + 1

        TempStream = ""

        PDFOutStream(sTempStream, "%DEBUT_OBJ/%")
        PDFOutStream(TempStream, CurrentObjectNum & " 0 obj")
        PDFOutStream(TempStream, "<< /Length " & (CurrentObjectNum + 1) & " 0 R")
        PDFOutStream(TempStream, " >>")

        StreamSize1 = Len(TempStream)

        PDFOutStream(TempStream, "stream")
        sTempStream = ""
        dTempStream = ""
    End Sub

    Public Sub PDFSetBookmark(str_Text As String, Optional iLevel As Integer = 0, Optional Y As Double = -1)
        If Y = -1 Then Y = in_yCurrent

        ReDim Preserve aOutlines(0 To iOutlines)

        aOutlines(iOutlines).sText = str_Text
        aOutlines(iOutlines).iLevel = iLevel
        aOutlines(iOutlines).yPos = Y
        aOutlines(iOutlines).iPageNb = PDFPageNumber

        iOutlines = iOutlines + 1
    End Sub

    Public ReadOnly Property PDFPageNumber() As Integer
        Get
            'Attribute PDFPageNumber.VB_HelpID = 2019
            PDFPageNumber = FPageNumber
        End Get
    End Property

    Public Property PDFSetLineStyle() As PDFStyleLgn
        Get

        End Get
        Set(value As PDFStyleLgn)
            'Attribute PDFSetLineStyle.VB_HelpID = 2047
            PDFLnStyle = PDFLineStyle(value)
        End Set
    End Property

    Private Function PDFLineStyle(ByRef pLineStyle As PDFStyleLgn) As String
        'Attribute PDFLineStyle.VB_HelpID = 2050

        Dim pTmpLineStyle As PDFStyleLgn

        PDFLineStyle = ""
        pTmpLineStyle = pLineStyle

        Select Case pTmpLineStyle
            Case PDFStyleLgn.pPDF_SOLID
                PDFLineStyle = "[] 0 d"
            Case PDFStyleLgn.pPDF_DASH
                PDFLineStyle = "[" & Int(16 * in_Ech) & " " & Int(8 * in_Ech) & " ] 0 d"
            Case PDFStyleLgn.pPDF_DASHDOT
                PDFLineStyle = "[" & Int(8 * in_Ech) & " " & Int(7 * in_Ech) & " " &
                               Int(2 * in_Ech) & " " & Int(7 * in_Ech) & " ] 0 d"
            Case PDFStyleLgn.pPDF_DASHDOTDOT
                PDFLineStyle = "[" & Int(8 * in_Ech) & " " & Int(4 * in_Ech) & " " &
                               Int(2 * in_Ech) & " " & Int(4 * in_Ech) & " " &
                               Int(2 * in_Ech) & " " & Int(4 * in_Ech) & " ] 0 d"
            Case Else
                MessageBox.Show("Line style set incorrectly : " & pLineStyle & "." & vbNewLine & "Line style set to solid.", "Line Style - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                PDFLineStyle = "[] 0 d"
        End Select
    End Function

    Public Property PDFSetLineWidth() As Double
        Get

        End Get
        Set(value As Double)
            'Attribute PDFSetLineWidth.VB_HelpID = 2048
            PDFLnWidth = value
        End Set
    End Property

    Public Sub PDFTextOut(str_Text As String, Optional ByRef X As Double = -1, Optional ByRef Y As Double = -1)
        'Attribute PDFTextOut.VB_HelpID = 2072

        Dim J As Integer
        Dim in_PositionFont As Integer
        Dim str_Tmp As String
        Dim str_TmpText As String

        str_TmpText = Replace(str_Text, "\", "\\")
        str_TmpText = Replace(str_TmpText, "\\", "\\\\")
        str_TmpText = Replace(str_TmpText, "(", "\(")
        str_TmpText = Replace(str_TmpText, ")", "\)")

        str_Tmp = ""

        If X = -1 Then X = in_xCurrent
        If Y = -1 Then Y = in_yCurrent

        If PDFFontName = "" Then
            in_PositionFont = 1
        Else
            For J = 0 To UBound(Arr_Font)
                If Arr_Font(J) = PDFFontName Then
                    in_PositionFont = J + 1
                    Exit For
                End If
            Next
        End If

        If PDFFontSize = 0 Then PDFFontSize = 10
        If PDFTextColor <> "" Then PDFOutStream(sTempStream, "q " & PDFTextColor & " ")
        If boPDFUnderline Then str_Tmp = PDFUnderline(False, str_Text, CDbl(X * in_Ech), CDbl(Y * in_Ech))

        PDFOutStream(sTempStream, "%DEBUT_TEXT/%")
        PDFOutStream(sTempStream, "BT")

        If PDFAngle = 0 Then
            PDFOutStream(sTempStream, PDFFormatDouble((X + PDFlMargin) * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " Td")
        Else
            PDFStreamRotate(PDFAngle, X, Y)
            PDFAngle = 0
        End If

        PDFOutStream(sTempStream, "/F" & in_PositionFont & " " & PDFFormatDouble(PDFFontSize) & " Tf")
        PDFOutStream(sTempStream, "(" & str_TmpText & ") Tj")

        If PDFTextColor <> "" Then
            PDFOutStream(sTempStream, "ET")

            If boPDFUnderline = True Then
                PDFOutStream(sTempStream, str_Tmp)
            End If

            PDFOutStream(sTempStream, "Q")
        Else
            PDFOutStream(sTempStream, "ET")

            If boPDFUnderline = True Then
                PDFOutStream(sTempStream, str_Tmp)
            End If
        End If

        PDFOutStream(sTempStream, "%FIN_TEXT/%")

        boPDFUnderline = False

        in_xCurrent = X + PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize)
        in_yCurrent = Y + PDFFontSize
    End Sub

    Public Function PDFGetStringWidth(ByVal str_Txt As String, Optional ByVal str_FName As String = "", Optional ByVal in_FSize As Integer = 0) As Double
        'Attribute PDFGetStringWidth.VB_HelpID = 2097

        Dim str_TmpINI As String
        Dim in_Tmp As Integer
        Dim In_i As Integer
        Dim In_j As Integer
        Dim ArrFNT() As Integer
        Dim in_Asc As Integer
        Dim Fso As Object
        Dim F As Object
        Dim aTempFNT As Object
        Dim bWX As Boolean
        Dim iAscMin As Integer
        Dim iAscMax As Integer
        Dim aAsc As Object
        Dim aWX As Object
        Dim sReadLine As String

        If str_FName = "" Then
            str_FName = PDFFontName
        End If

        'ReDim ArrFNT(1 To 255)
        ReDim ArrFNT(0 To 254)
        iAscMin = 0
        iAscMax = 0

        bWX = False

        Fso = CreateObject("Scripting.FileSystemObject")
        F = Fso.OpenTextFile(wsPathConfig & "\" & str_FName & ".afm", 1, 0)
        Do While F.AtEndOfStream <> True
            sReadLine = F.ReadLine

            If InStr(1, sReadLine, "StartCharMetrics") <> 0 Then
                bWX = True
                sReadLine = F.ReadLine
            End If

            If InStr(1, sReadLine, "-1 ;") <> 0 Or
   InStr(1, sReadLine, "EndCharMetrics") <> 0 Then
                iAscMax = aAsc(1)
                Exit Do
            End If

            If bWX = True Then
                aTempFNT = Split(sReadLine, ";")
                aAsc = Split(Trim(aTempFNT(0)), " ")
                If iAscMin = 0 Then iAscMin = aAsc(1)

                aWX = Split(Trim(aTempFNT(1)), " ")

                ArrFNT(aAsc(1)) = Int(aWX(1))
            End If
        Loop
        F.Close

        For In_i = 1 To 255
            If In_i < iAscMin Then ArrFNT(In_i) = 0
            If In_i > iAscMax Then ArrFNT(In_i) = 0
        Next

        in_Tmp = 0
        For In_i = 1 To Len(str_Txt)
            in_Asc = Asc(Mid(str_Txt, In_i, 1))
            in_Tmp = in_Tmp + Int(ArrFNT(in_Asc)) ' + FontBBoxAbout
        Next

        PDFGetStringWidth = (in_Tmp * in_FSize) / 1000
    End Function

    Private Function PDFFormatDouble(In_dbl As Object, Optional ByRef nZero As Integer = 2) As String
        'Attribute PDFFormatDouble.VB_HelpID = 2100

        Dim sZero As String

        'sZero = String(nZero, "0")
        sZero = New String("0", nZero)
        PDFFormatDouble = Replace(Format(In_dbl, "###0." & sZero), ",", ".")
    End Function

    Private Sub PDFStreamRotate(ByRef pAngle As Double, ByRef X As Double, ByRef Y As Double)
        Dim dSin As Double
        Dim dCos As Double
        Dim CenterX As Double
        Dim CenterY As Double

        If pAngle <> 0 Then
            pAngle = pAngle * 3.1416 / 180
            dCos = Math.Cos(pAngle)
            dSin = Math.Sin(pAngle)
            CenterX = X * in_Ech
            CenterY = PDFCanvasHeight(in_Canvas) - Y * in_Ech

            PDFOutStream(sTempStream, PDFFormatDouble(dCos, 5) & " " &
                                  PDFFormatDouble(-1 * dSin, 5) & " " &
                                  PDFFormatDouble(dSin, 5) & " " &
                                  PDFFormatDouble(dCos, 5) & " " &
                                  PDFFormatDouble(CenterX) & " " &
                                  PDFFormatDouble(CenterY) & " Tm")
        End If

        bAngle = True
    End Sub

    Private Function PDFUnderline(ByRef boCell As Boolean, str_Text As String, ByRef X As Double, ByRef Y As Double) As String
        'Attribute PDFUnderline.VB_HelpID = 2091

        Dim in_wUp As Integer
        Dim in_wUt As Integer
        Dim in_wTxt As String

        Dim in_Px As Integer
        Dim in_Pw As String
        Dim in_Py As String

        Dim str_TmpUnderl As String

        Dim str_xLeft As String
        Dim str_yTop As String
        Dim str_wText As String
        Dim str_hLine As String
        Dim iNbSpace As Integer

        str_TmpUnderl = ""

        in_wUp = PDFGetStringWidth("up", PDFFontName, PDFFontSize)
        in_wUt = 2

        iNbSpace = PDFGetNumberOfCar(str_Text, " ")
        in_wTxt = PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize) +
  iNbSpace * PDFGetStringWidth(" ", PDFFontName, PDFFontSize) +
  iWidthStr * iNbSpace -
  IIf(iWidthStr <> 0, (iNbSpace + 1) * PDFcMargin, 0)

        in_Px = X + PDFlMargin * in_Ech
        in_Pw = (PDFCanvasHeight(in_Canvas) - (Y - in_wUp / 1000 * PDFFontSize) - 2)
        in_Py = -in_wUt / 1000 * in_wTxt
        str_hLine = PDFFormatDouble(in_Py)

        If boCell = False Then
            str_wText = PDFFormatDouble(in_wTxt)
            str_xLeft = PDFFormatDouble(in_Px)
            str_yTop = PDFFormatDouble(in_Pw)

            str_TmpUnderl = str_xLeft & " " & str_yTop & " " & str_wText & " " & str_hLine & " re f"
        Else
            str_wText = PDFFormatDouble(in_wTxt - PDFcMargin)
            str_xLeft = PDFFormatDouble(X)
            str_yTop = PDFFormatDouble(Y - 3)

            str_TmpUnderl = str_xLeft & " " & str_yTop & " " & str_wText & " " & str_hLine & " re f"
        End If

        PDFUnderline = str_TmpUnderl
    End Function

    Private Function PDFGetNumberOfCar(ByRef sText As String, ByRef sCar As String) As Integer
        Dim iNbCar As Integer
        Dim In_i As Integer

        iNbCar = 0
        In_i = InStr(1, sText, sCar)
        If In_i <> 0 Then iNbCar = 1

        Do While In_i <> 0
            In_i = InStr(In_i + 1, sText, sCar)
            If In_i <> 0 Then iNbCar = iNbCar + 1
        Loop

        PDFGetNumberOfCar = iNbCar
    End Function

    Public Sub PDFSetFont(str_Fontname As PDFFontNme, in_FontSize As Integer, Optional str_Style As PDFFontStl = PDFFontStl.FONT_NORMAL)
        'Attribute PDFSetFont.VB_HelpID = 2051
        Dim str_TmpFontName As String
        Dim str_TmpFontNm As String

        If str_Fontname <> PDFFontNme.FONT_ARIAL And
       str_Fontname <> PDFFontNme.FONT_COURIER And
       str_Fontname <> PDFFontNme.FONT_SYMBOL And
       str_Fontname <> PDFFontNme.FONT_TIMES And
       str_Fontname <> PDFFontNme.FONT_ZAPFDINGBATS Then
            MessageBox.Show("Font name set incorrectly : " & str_Style & "." & vbNewLine & "Font set to Times New Roman.", "Font name - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            str_TmpFontName = "TimesRoman"
            boPDFItalic = False
            boPDFBold = False

            PDFFontName = str_TmpFontName
            PDFFontNum = FontNum
            PDFFontSize = in_FontSize

            FontNum = FontNum + 1

            Exit Sub
        End If

        Select Case str_Fontname
            Case PDFFontNme.FONT_ARIAL
                str_TmpFontNm = "Arial"
            Case PDFFontNme.FONT_COURIER
                str_TmpFontNm = "Courier"
            Case PDFFontNme.FONT_TIMES
                str_TmpFontNm = "Times"
            Case PDFFontNme.FONT_SYMBOL
                str_TmpFontNm = "Symbol"
            Case PDFFontNme.FONT_ZAPFDINGBATS
                str_TmpFontNm = "ZapfDingbats"
        End Select

        If str_TmpFontNm = "Arial" Then
            str_TmpFontName = "Helvetica"
        Else
            str_TmpFontName = str_TmpFontNm
        End If

        boPDFItalic = False
        boPDFBold = False

        str_TmpFont = str_TmpFontName

        If InStr(1, str_Style.ToString, PDFFontStl.FONT_ITALIC.ToString) <> 0 Then boPDFItalic = True
        If InStr(1, str_Style.ToString, PDFFontStl.FONT_BOLD.ToString) <> 0 Then boPDFBold = True
        If InStr(1, str_Style.ToString, PDFFontStl.FONT_UNDERLINE.ToString) <> 0 Then boPDFUnderline = True

        If boPDFItalic = True And boPDFBold = False Then
            Select Case str_TmpFontName
                Case "Times"
                    str_TmpFontName = "TimesItalic"
                Case Else
                    str_TmpFontName = str_TmpFontName & "-Oblique"
            End Select
        End If

        If boPDFItalic = True And boPDFBold = True Then
            Select Case str_TmpFontName
                Case "Times"
                    str_TmpFontName = str_TmpFontName & "-BoldItalic"
                Case Else
                    str_TmpFontName = str_TmpFontName & "-BoldOblique"
            End Select
        End If

        If boPDFItalic = False And boPDFBold = True Then
            str_TmpFontName = str_TmpFontName & "-Bold"
        End If

        If boPDFItalic = False And boPDFBold = False Then
            Select Case str_TmpFontName
                Case "Times"
                    str_TmpFontName = str_TmpFontName & "-Roman"
                Case Else
                    str_TmpFontName = str_TmpFontName
            End Select
        End If

        PDFFontName = str_TmpFontName
        PDFFontNum = FontNum
        PDFFontSize = in_FontSize

        FontNum = FontNum + 1
    End Sub

    Public Function PDFSetTextColor(ByRef gColor As Object) As Object
        'Attribute PDFSetTextColor.VB_HelpID = 2063

        Dim TxtCl As PDFRGB
        Dim sColor As String

        Select Case TypeName(gColor)
            Case "Variant()"
                TxtCl.in_R = gColor(0)
                TxtCl.in_G = gColor(1)
                TxtCl.in_B = gColor(2)
            Case "String"
                If Left(gColor, 1) <> "#" Then
                    MessageBox.Show("Invalid HTMl color set" & gColor & "." & vbNewLine & "Set color to  black.", "Text Color " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    'TxtCl = PDFGetRGB(vbBlack)
                    TxtCl = PDFGetRGB(Color.Black.ToArgb)
                Else
                    TxtCl = PDFHtml2RgbColor(CStr(gColor))
                End If
            Case Else
                TxtCl = PDFGetRGB(Int(gColor))
        End Select

        PDFTextColor = PDFStreamColor(TxtCl, "TEXT")
    End Function

    Private Function PDFGetRGB(ByRef lColor As Integer) As PDFRGB
        'Attribute PDFGetRGB.VB_HelpID = 2099
        PDFGetRGB.in_B = CByte(Int(lColor / 65536))
        PDFGetRGB.in_G = CByte(Int((lColor - CLng(PDFGetRGB.in_B) * 65536) / 256))
        PDFGetRGB.in_R = CByte(lColor - CLng(PDFGetRGB.in_B) * 65536 - CLng(PDFGetRGB.in_G) * 256)
    End Function

    Private Function PDFHtml2RgbColor(ByRef sColor As String) As PDFRGB
        Dim sTmpColor As String

        sTmpColor = Right("000000" & sColor, 6)
        PDFHtml2RgbColor.in_R = CByte("&h" & Mid(sTmpColor, 1, 2))
        PDFHtml2RgbColor.in_G = CByte("&h" & Mid(sTmpColor, 3, 2))
        PDFHtml2RgbColor.in_B = CByte("&h" & Mid(sTmpColor, 5, 2))
    End Function

    Private Function PDFStreamColor(ByRef PDFRgbColor As PDFRGB, str_Type As String) As String
        'Attribute PDFStreamColor.VB_HelpID = 2069
        Dim Int_r As Integer
        Dim Int_g As Integer
        Dim Int_b As Integer
        Dim str_TxtColor As String

        Int_r = PDFRgbColor.in_R
        Int_g = PDFRgbColor.in_G
        Int_b = PDFRgbColor.in_B

        Select Case str_Type
            Case "TEXT", "BORDER"
                str_TxtColor = Replace(Format(Int_r / 255, "0.000"), ",", ".") & " " &
                           Replace(Format(Int_g / 255, "0.000"), ",", ".") & " " &
                           Replace(Format(Int_b / 255, "0.000"), ",", ".") & " rg"
            Case "LINE"
                str_TxtColor = Replace(Format(Int_r / 255, "0.000"), ",", ".") & " " &
                           Replace(Format(Int_g / 255, "0.000"), ",", ".") & " " &
                           Replace(Format(Int_b / 255, "0.000"), ",", ".") & " RG"
        End Select

        PDFStreamColor = str_TxtColor
    End Function

    Public ReadOnly Property PDFGetPageHeight() As Double
        Get
            'Attribute PDFGetPageHeight.VB_HelpID = 2031
            PDFGetPageHeight = PDFCanvasHeight(in_Canvas)
        End Get
    End Property

    Public ReadOnly Property PDFGetPageWidth() As Double
        Get
            'Attribute PDFGetPageWidth.VB_HelpID = 2033
            PDFGetPageWidth = PDFCanvasWidth(in_Canvas)
        End Get
    End Property

End Class

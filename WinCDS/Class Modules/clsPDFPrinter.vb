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
    Private ObjectOffsetList As Object
    Private PageNumberList As Object
    'Private PageLinksList(1 To 1000, 1 To 1000) As Object  '@NO-LINT
    Private PageLinksList(0 To 999, 0 To 999) As Object  '@NO-LINT
    Private LinksList As Object
    Private PageCanvasWidth As Object
    Private PageCanvasHeight As Object
    Private FontNumberList As Object
    Private CRCounter As Integer

    Private ColorSpace As String
    Private ColorCount As Byte
    Private ImageStream As String
    Private TempStream As String
    Private pTempStream As String
    Private sTempStream As String
    Private cTempStream As String
    Private dTempStream As String
    Private boPageLinksList As Object
    Private NbPageLinksList As Object
    Private StreamSize1 As Integer, StreamSize2 As Integer
    Private in_xCurrent As Double
    Private in_yCurrent As Double
    Private aOutlines() As oOutlines
    Private iOutlines As Integer
    Private aPage() As Object

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
                Case "Long"
                    Select Case value
                        Case PDFFormatPgStr.FORMAT_A4
                            PDFCanvasWidth(in_Canvas) = 595.28
                            PDFCanvasHeight(in_Canvas) = 841.89
                        Case PDFFormatPgStr.FORMAT_A3
                            PDFCanvasWidth(in_Canvas) = 841.89
                            PDFCanvasHeight(in_Canvas) = 1190.55
                        Case PDFFormatPgStr.FORMAT_A5
                            PDFCanvasWidth(in_Canvas) = 420.94
                            PDFCanvasHeight(in_Canvas) = 595.28
                        Case PDFFormatPgStr.FORMAT_LETTER
                            PDFCanvasWidth(in_Canvas) = 612
                            PDFCanvasHeight(in_Canvas) = 792
                        Case PDFFormatPgStr.FORMAT_LEGAL
                            PDFCanvasWidth(in_Canvas) = 612
                            PDFCanvasHeight(in_Canvas) = 1008
                        Case Else
                            MessageBox.Show("Format page set incorrectly : " & value & "." & vbNewLine & "Format page set to A4.", "Format Page - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            PDFCanvasWidth(in_Canvas) = 595.28
                            PDFCanvasHeight(in_Canvas) = 841.89
                    End Select
                Case "Double()"
                    'PDFCanvasWidth(in_Canvas) = str_FormatPage(0)
                    PDFCanvasWidth(in_Canvas) = value
                    'PDFCanvasHeight(in_Canvas) = str_FormatPage(1)
                    PDFCanvasHeight(in_Canvas) = value
                Case Else
                    MessageBox.Show("Format page set incorrectly : " & value & "." & vbNewLine & "Format page set to A4", "Format Page - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    PDFCanvasWidth(in_Canvas) = 595.28
                    PDFCanvasHeight(in_Canvas) = 841.89
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

            tmp_PDFCanvasWidth = PDFCanvasWidth(in_Canvas)
            tmp_PDFCanvasHeight = PDFCanvasHeight(in_Canvas)

            Select Case value
                Case PDFOrientationStr.ORIENT_PORTRAIT
                    PDFCanvasWidth(in_Canvas) = tmp_PDFCanvasWidth
                    PDFCanvasHeight(in_Canvas) = tmp_PDFCanvasHeight
                    PDFCanvasOrientation(in_Canvas) = "p"
                Case PDFOrientationStr.ORIENT_PAYSAGE
                    PDFCanvasWidth(in_Canvas) = tmp_PDFCanvasHeight
                    PDFCanvasHeight(in_Canvas) = tmp_PDFCanvasWidth
                    PDFCanvasOrientation(in_Canvas) = "l"
                Case Else
                    MessageBox.Show("Orientation set incorrectly: " & value & "." & vbNewLine & "Orientation set to portrait.", "Error in orientation - " & mjwPDFVersion, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    PDFCanvasWidth(in_Canvas) = tmp_PDFCanvasWidth
                    PDFCanvasHeight(in_Canvas) = tmp_PDFCanvasHeight
                    PDFCanvasOrientation(in_Canvas) = "p"
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
        ObjectOffsetList(In_offset) = ObjectOffset

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

End Class

Imports stdole
Public Class frmDesignTag
    Private mIsExternal As Boolean
    Private mLoadedForDisplay As Boolean
    Private Const Form_MinWidth As Long = 12000
    Private Const Form_MinHeight As Long = 8175

    Private Const TwipsPerInch As Long = 1440
    Private Const CurrentTagLayoutVersion As String = "1.2"
    Private Const DefaultFields As Long = 12

    Private Const SnapIncr As Long = 120
    Private Const FontIncr As Double = 0.5

    Private Const TAGDIM_MAX_X As Single = 11 '8.5    ' 11 b/c of landscape
    Private Const TAGDIM_MAX_Y As Single = 11

    Private Const TAGDIM_L_X As Single = 8.5
    Private Const TAGDIM_L_Y As Single = 11
    Private Const TAGDIM_M_X As Single = 3
    Private Const TAGDIM_M_Y As Single = 4
    Private Const TAGDIM_S_X As Single = 1
    Private Const TAGDIM_S_Y As Single = 2

    Private Const TAGDIM_C_X As Single = 3    ' these are the initial defaults for custom
    Private Const TAGDIM_C_Y As Single = 3    ' which of course can be adjusted then

    Private OrigPrinter As String
    Private PrintingStyle As String
    Private RN As Long

    Private lstItemsAdjust As Boolean
    Private mAllowPrintMany As Boolean

    Private OrigX As Single, OrigY As Single
    Private inX As Single, inY As Single
    Private Moving As Boolean

    Private CurrentField As Long
    Private Fields() As TagItemLayout

    Private Enum LayoutAlign
        lyaPosition = 0
        lyaLeft = 1
        lyaCenter = 2
        lyaRight = 3
        lyaPositionR = 4
    End Enum

    Private Structure TagItemLayout
        Dim Name As String
        Dim Caption As String
        Dim Visible As Boolean
        Dim ToolTipText As String

        Dim FontName As String
        Dim FontSize As Single
        Dim FontColor As String
        Dim CharSpecify As String

        Dim Left As Long
        Dim Top As Long
        Dim Alignment As LayoutAlign

        Dim ExtraFieldType As Long
        Dim PicWidth As Long
        Dim PicHeight As Long
        Dim PicLock As Boolean
    End Structure

    Public Function PrintCustomTags(ByVal Style As String, Optional ByVal Quantity As Long = 1, Optional ByVal TemplateName As String = "", Optional ByVal External As Boolean = True) As Boolean
        Dim I As Long
        TagLog("PrintCustomTags Style=" & Style & ", Qty=" & Quantity & ", Template=" & TemplateName)
        If Quantity < 0 Then Exit Function
        If External And TemplateName = "" Then Exit Function
        mIsExternal = External
        If TemplateName <> "" Then LoadTagLayout(TemplateName)
        PrintingStyle = Style
        RefreshItemCaptions(Style)

        For I = 1 To Quantity
            PrintCurrentTag(cmbPageAlign.SelectedIndex, False, IIf(Quantity = 1, -1, I - 1)) ' for 1 tag, print it w/ alignment options.  Multiple tags go where they specified.
        Next
        Printer.EndDoc()  ' an extra one, because it never hurts
        PrintingStyle = ""
        PrintCustomTags = True
    End Function

    Private ReadOnly Property MaxField() As Long
        Get
            On Error Resume Next
            MaxField = UBound(Fields)
        End Get
    End Property

    Private Function IsCharBold(ByVal Specify As String) As Boolean
        IsCharBold = InStr(Specify, "B") > 0
    End Function

    Private Function IsCharItalic(ByVal Specify As String) As Boolean
        IsCharItalic = InStr(Specify, "I") > 0
    End Function

    Private Function IsCharUnderline(ByVal Specify As String) As Boolean
        IsCharUnderline = InStr(Specify, "U") > 0
    End Function

    Private Sub PrintCurrentTag(Optional ByVal PageAlign As Long = 0, Optional ByVal WithBox As Boolean = True, Optional ByVal PageLoc As Long = -1)
        Dim I As Long
        Dim XOff As Long, YOff As Long, EoP As Boolean
        Dim X As Long, Y As Long
        Dim W As Long, H As Long
        Dim C As String

        PreparePrinter()

        ' ----- logging
        TagLog("PrintCurrentTag CurrentPrinter=" & Printer.DeviceName & ", Orientation=" & Printer.Orientation & ", ScaleMode=" & Printer.ScaleMode)
        TagLog("PrintCurrentTag PrinterDimensions=" & Printer.Width & "x" & Printer.Height & ", Scale=" & Printer.ScaleWidth & "x" & Printer.ScaleHeight)
        'Printer.ScaleMode = vbInches ' show in inches
        Printer.ScaleMode = VBRUN.ScaleModeConstants.vbInches
        TagLog("PrintCurrentTag PrinterDimensions=" & Printer.Width & """x" & Printer.Height & """, Scale=" & Printer.ScaleWidth & """x" & Printer.ScaleHeight & """")
        'Printer.ScaleMode = vbTwips
        Printer.ScaleMode = VBRUN.ScaleModeConstants.vbTwips
        ' -----

        GetPageAlignOffsets(PageAlign, PageLoc, XOff, YOff, EoP)
        On Error Resume Next
        If WithBox Then Printer.Line(XOff, YOff)-(XOff + TagWidth, YOff + TagHeight), , B
  
  For I = MaxField To 1 Step -1
            If Fields(I).Visible Then
                Printer.FontName = Fields(I).FontName
                Printer.FontSize = Fields(I).FontSize
                Printer.FontBold = IsCharBold(Fields(I).CharSpecify)
                Printer.FontItalic = IsCharItalic(Fields(I).CharSpecify)
                Printer.FontUnderline = IsCharUnderline(Fields(I).CharSpecify)
                Printer.ForeColor = IIf(Fields(I).FontColor = "", Color.Black, Fields(I).FontColor)
                Printer.FontTransparent = False
                Y = Fields(I).Top
                X = Fields(I).Left
                C = ParseTagCode(Fields(I).Caption)
                If IsIn(Fields(I).Name, "Description", "Comments") Then C = WrapLongText(C, 46)
                If Fields(I).Name = "Bar Code" Then C = PrepareBarcode(C)
                W = Printer.TextWidth(C)
                H = Printer.TextHeight(C)
                Select Case Fields(I).Alignment
                    Case LayoutAlign.lyaPositionR
                    Case LayoutAlign.lyaLeft : X = 0
                    Case LayoutAlign.lyaRight : X = TagWidth - W
                    Case LayoutAlign.lyaCenter : X = (TagWidth - W) / 2
                    Case LayoutAlign.lyaPositionR : X = X - W
                End Select
                If Fields(I).ExtraFieldType = 0 Then
                    Printer.CurrentX = XOff + X
                    Printer.CurrentY = YOff + Y
                    TagLog("Printing...  Loc=" & Printer.CurrentX & "x" & Printer.CurrentY, 7)
                    If IsIn(Fields(I).Name, "Description", "Comments") Then
                        Dim LLL As Object
                        For Each LLL In Split(C, vbCrLf)
                            Select Case Fields(I).Alignment
                                Case lyaPosition
                                Case lyaLeft : X = 0
                                Case lyaRight : X = TagWidth - W
                                Case lyaCenter : X = (TagWidth - W) / 2
                                Case lyaPositionR : X = X - W
                            End Select
                            Printer.CurrentX = XOff + X
                            Printer.Print(LLL)
                        Next
                    Else
                        Printer.Print(C)
                    End If
                Else
                    Dim P As IPictureDisp, pW As Long, pH As Long, aPW As Long, aPH As Long
                    P = LoadItemImage(Fields(I).Caption)
                    pW = P.Width
                    pH = P.Height
                    If Fields(I).PicWidth > 0 Then pW = Fields(I).PicWidth
                    If Fields(I).PicHeight > 0 Then pH = Fields(I).PicHeight
                    aPW = pW
                    aPH = pH

                    If Fields(I).PicLock Then
                        'imgPrintHelper.Picture = P
                        imgPrintHelper.Image = P
                        MaintainPictureRatio(imgPrintHelper, pW, pH, False)
                        X = X + (aPW - pW) / 2
                        Y = Y + (aPH - pH) / 2
                    End If

                    'If P <> 0 Then Printer.PaintPicture P, XOff + X, YOff + Y, pW, pH
                    If P IsNot Nothing Then Printer.PaintPicture(P, XOff + X, YOff + Y, pW, pH)

                    P = Nothing
                End If
            End If
        Next
        If EoP Then Printer.EndDoc()
    End Sub

    Private Function LoadItemImage(ByVal Caption As String) As IPictureDisp
        Caption = ImageFileName(Caption)
        If Caption = "" Then
    Set LoadItemImage = il.ListImages("invalid").Picture
  ElseIf Caption = "0" Then
    Set LoadItemImage = il.ListImages("blank").Picture
  Else
            If Dir(Caption) = "" Then
      Set LoadItemImage = il.ListImages("invalid").Picture
    Else
      Set LoadItemImage = LoadPictureStd(Caption)
    End If
        End If
    End Function

    Private Function ParseTagCode(ByVal Str As String) As String
        Dim N As Long, Style As String, Sp As String, L As Long, K As Long
        Dim F As String, I As Long, Tot As Currency
        If Left(Str, 1) <> "#" Then ParseTagCode = ParseExtraFieldToken(Str) : Exit Function
        Str = Mid(Str, 2)
        N = InStr(Str, ":")

        If N <= 0 Then ParseTagCode = "#" & Str : Exit Function

        F = LCase(Replace(Mid(Str, 1, N - 1), " ", ""))
        '  If F = "onsale" Then Stop
        Sp = Mid(Str, N + 1)
        ParseLineKey Sp, L, K
  If K = 0 Then
            Select Case F
                Case "list", "listprice", "landed", "onsale", "onsaleprice", "sale", "saleprice"
                    For I = 1 To 10
                        Style = GetMultipleStyle(L, I)
                        If Style <> "" Then Tot = Tot + GetPrice(GetItemField(Style, F))
                    Next
                    ParseTagCode = lCurrencyFormat(Tot)
                Case Else
                    Style = GetMultipleStyle(L, 1)
                    ParseTagCode = GetItemField(Style, F)
            End Select
        Else
            Style = GetMultipleStyle(L, K)
            ParseTagCode = GetItemField(Style, F)
        End If
    End Function

    Private Sub GetPageAlignOffsets(ByVal PageAlign As Long, ByVal PageLoc As Long, ByRef XOffset As Long, ByRef YOffset As Long, ByRef EndOfPage As Boolean)
        Dim XCenter As Long, YCenter As Long
        Dim XMax As Long, YMax As Long
        XCenter = (Printer.Width - TagWidth) / 2
        YCenter = (Printer.Height - TagHeight) / 2
        XMax = (Printer.Width - TagWidth)
        YMax = (Printer.Height - TagHeight)

        TagLog "GetPageAlignOffsets PageAlign=" & PageAlign & ", Pageloc=" & PageLoc & ", XOffset=" & XOffset & ", YOffset=" & YOffset & ", EndOfPage=" & EndOfPage
  TagLog "GetPageAlignOffsets XCenter=" & XCenter & ", YCenter=" & YCenter & ", XMax=" & XMax & ", YMax=" & YMax

  If PageLoc < 0 Then
            EndOfPage = True
            Select Case PageAlign
                Case 0 : XOffset = 0 : YOffset = 0
                Case 1 : XOffset = XCenter : YOffset = 0
                Case 2 : XOffset = XMax : YOffset = 0
                Case 3 : XOffset = 0 : YOffset = YCenter
                Case 4 : XOffset = XCenter : YOffset = YCenter
                Case 5 : XOffset = XMax : YOffset = YCenter
                Case 6 : XOffset = 0 : YOffset = YMax
                Case 7 : XOffset = XCenter : YOffset = YMax
                Case 8 : XOffset = XMax : YOffset = YMax
            End Select
        Else
            Dim Mx As Long, mY As Long, Tot As Long
            Mx = Int(CDbl(Printer.ScaleWidth) / CDbl(TagWidth + 120))
            mY = Int(CDbl(Printer.ScaleHeight) / CDbl(TagHeight + 120))
            Tot = Mx * mY
            PageLoc = PageLoc Mod Tot
            XOffset = Int(PageLoc Mod Mx) * (TagWidth + 120)
            YOffset = Int(PageLoc / Mx) * (TagHeight + 120)
            EndOfPage = (PageLoc = Tot - 1)
        End If

        TagLog "GetPageAlignOffsets PageAlign=" & PageAlign & ", Pageloc=" & PageLoc & ", XOffset=" & XOffset & ", YOffset=" & YOffset & ", EndOfPage=" & EndOfPage
End Sub

    Private Function TagLog(ByVal Msg As String, Optional ByVal Importance As Long = 3) As Boolean
        ActiveLog "frmDesignTag::" & Msg, Importance
  TagLog = True
    End Function

    Private Sub LoadTagLayout(ByVal TagName As String, Optional ByVal AsTemplate As Boolean = False)
        Dim FN As String, Opts As String
        If TagName = "-Select From List-" Or TagName = "(Default)" Then Exit Sub

        FN = TagLayoutFileName(TagName, AsTemplate)
        Opts = ReadFile(FN, 1, 1)

        If InStr(Opts, "Version=" & CurrentTagLayoutVersion) = 0 Then ConvertTagLayout FN

  DoLoadTagLayout FN, AsTemplate
End Sub

    Private Sub RefreshItemCaptions(Optional ByVal Style As String = "")
        Dim I As Long, C As CInvRec, K As cInvKit, X As String, Kit As Boolean
  Set C = New CInvRec
  If Style <> "" Then
            If Left(Style, 4) = KIT_PFX Then
      Set K = New cInvKit
      If Not K.Load(Style, "KitStyleNo") Then
                    Style = ""
                    MsgBox "Kit not found.", vbInformation, "Tag Designer"
      Else
                    RN = 0
                    Kit = True
                    C.Load K.Item1, "Style"
      End If
            Else
                If Not C.Load(Style, "Style") Then
                    Style = ""
                    MsgBox "Style not found.", vbInformation, "Tag Designer"
      Else
                    RN = C.RN
                    Kit = False
                End If
            End If
        End If
        If Style = "" Then RN = 0

        For I = LBound(Fields) To UBound(Fields)
            If I > DefaultFields Then
                If Fields(I).ExtraFieldType = 1 And Fields(I).Caption = "#" Then
                    Fields(I).Caption = RN
                Else
                    X = Fields(I).Caption
                End If
            Else
                If Style = "" Then
                    X = Fields(I).Name
                    If I = 1 Then X = "0000000000"    ' put something in barcode
                Else
                    If Kit Then
                        Select Case I       ' for kits
                            Case 1 : X = Replace(K.KitStyleNo, " ", "_") 'PrepareBarcode(C.Style)
                            Case 2 : X = IIf(StoreSettings.bStyleNoInCode, ConvertCostToCode(K.KitStyleNo), K.KitStyleNo)
                            Case 3 : X = K.Heading
                            Case 4 : X = C.GetItemCode   ' from first item
                            Case 5 : X = CurrencyFormat(K.List, chkHideCents.Value = 1, chkDollarSign.Value = 1)
                            Case 6 : X = CurrencyFormat(K.OnSale, chkHideCents.Value = 1, chkDollarSign.Value = 1)
                            Case 7 : X = ""
                            Case 8 : X = IIf(StoreSettings.bCostInCode, ConvertCostToCode(CurrencyFormat(K.Landed, chkHideCents.Value = 1, False)), CurrencyFormat(K.Landed, chkHideCents.Value = 1, False))
                            Case 9 : X = C.Vendor        ' from first item
                            Case 10 : X = C.VendorNo      ' from first item
                            Case 11 : X = CurrencyFormat(K.PackPrice, chkHideCents.Value = 1, chkDollarSign.Value = 1)
                            Case 12 : X = WrapLongText(K.MemoArea, 46)
                            Case Else : X = ""
                        End Select
                    Else
                        Select Case I       ' for items
                            Case 1 : X = Replace(C.Style, " ", "_") 'PrepareBarcode(C.Style)
                            Case 2 : X = IIf(StoreSettings.bStyleNoInCode, ConvertCostToCode(C.Style), C.Style)
                            Case 3 : X = WrapLongText(C.Desc, 46)
                            Case 4 : X = C.GetItemCode
                            Case 5 : X = CurrencyFormat(C.List, chkHideCents.Value = 1, chkDollarSign.Value = 1)
                            Case 6 : X = CurrencyFormat(C.OnSale, chkHideCents.Value = 1, chkDollarSign.Value = 1)
                            Case 7 : X = C.QueryStock(StoresSld)
                            Case 8 : X = IIf(StoreSettings.bCostInCode, ConvertCostToCode(CurrencyFormat(C.Landed, chkHideCents.Value = 1, False)), CurrencyFormat(C.Landed, chkHideCents.Value = 1, False))
                            Case 9 : X = C.Vendor
                            Case 10 : X = C.VendorNo
                            Case 11 : X = CurrencyFormat(C.OnSale, chkHideCents.Value = 1, chkDollarSign.Value = 1)
                            Case 12 : X = WrapLongText(C.Comments, 46)
                            Case Else : X = ""
                        End Select
                    End If
                End If
                Fields(I).Caption = X
            End If
        Next

        DisposeDA C, K

  RefreshFields
    End Sub

    Public Sub PreparePrinter()
        On Error Resume Next
        If IsDoubleR Then
        Else
            If Printer.DeviceName Like "*DYMO*" Then
                '    Printer.PaperSize = vbPRPSUser
                TagLog "PreparePrinter DYMO Sz: " & Printer.Width & """x" & Printer.Height & """"
  '    Printer.PaperSize = DYMO_PaperSize_30256
                Printer.ScaleMode = vbInches
                TagLog "PreparePrinter DYMO Pre-Sz: " & Printer.Width & """x" & Printer.Height & """"
      Printer.Width = txtCustomX
                Printer.Height = txtCustomY
                TagLog "PreparePrinter DYMO Set-Sz: " & Printer.Width & """x" & Printer.Height & """"
  '    Printer.Orientation = vbPRORLandscape
                Printer.ScaleMode = vbTwips
            End If
        End If
    End Sub

End Class

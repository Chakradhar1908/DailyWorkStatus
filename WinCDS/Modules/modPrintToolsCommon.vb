Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports VBPrnDlgLib
Imports VBRUN
Module modPrintToolsCommon
    Public OutputObject As Object
    Public OutputToPrinter As Boolean
    Public Const DYMO_PaperSize_30252 As Integer = 121                           ' 6x1.5 labels
    Public Const DYMO_PaperSize_30323 As Integer = 126                           ' 6x3 labels
    Public Const DYMO_PaperSize_30256 As Integer = 129                           ' 2x4 shipping
    Public Const DYMO_PaperSize_30270 As Integer = 186                           ' continuous tape
    Public Const DYMO_PaperSize_ContinuousWide As Integer = DYMO_PaperSize_30270 ' We didn't know it's SKU for a while..
    Public PageNumber As Integer
    Public ColCount As Integer, CommonReportPRG As ProgressBar
    Public CommonReportColSpacing As Integer, CommonReportIndent As Integer
    'Public Cols(1 To 50, 1 To 4) '@NO-LINT-NTYP
    Public Cols(0 To 49, 0 To 3) '@NO-LINT-NTYP
    Public Const DYMO_PaperBin_Left As Integer = 15                              '
    Public Const DYMO_PaperBin_Right As Integer = 16                             ' ?? Left is 15, we just did +1 for the twin
    Public Const DYMO_PaperBin_DEFAULT As Integer = DYMO_PaperBin_Left           ' Default to LEFT bin
    Public Const DYMO_PaperSize_ContW As Integer = DYMO_PaperSize_30270          ' A shorter alias...

    '  Public Sub PrintCentered(ByVal Text As String, Optional ByVal yPos as integer = -1, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
    '      Dim Ob As Boolean, oI As Boolean
    '      Dim OO As Object

    '      On Error Resume Next
    '      If OutputObject Is Nothing Then OutputObject = Printer
    '      OO = OutputObject
    '      Ob = OO.FontBold
    '      oI = OO.FontItalic
    '      If Bold Then OO.FontBold = True
    '      If Italic Then OO.FontItalic = True
    '      If yPos > 0 Then OO.CurrentY = yPos


    '      OO.CurrentX = (OutputObject.ScaleWidth - OO.TextWidth(Text)) / 2 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
    '      OO.CurrentX = (Printer.ScaleWidth - OO.TextWidth(Text)) / 2 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
    '      If Not IscPrinter(OutputObject) Then
    '          OutputObject.Print Text
    'Else
    '          OutputObject.PrintNL Text
    'End If

    '      If Bold Then OO.FontBold = Ob
    '      If Italic Then OO.FontItalic = oI
    '  End Sub
    Public Function NumLineBreaks(ByVal vStr As String) As Integer
        Dim tmpLines As Object, tmpStart As Object
        NumLineBreaks = ((Len(vStr) - 1) \ 46 + 1)

        tmpStart = 1
        Do
            tmpStart = NextPrintBreak(vStr, tmpStart)
            tmpLines = tmpLines + 1
        Loop Until tmpStart >= Len(vStr)

        If tmpLines > NumLineBreaks Then NumLineBreaks = tmpLines
    End Function

    Public Function SetDymoPrinter(Optional ByRef PaperType As Integer = 0) As Boolean
        ' bfh20051202 - we need to have error handling b/c of papersize setting
        ' w/o it, if they set a non-dymo up as their dymo, it will error b/c
        ' they would be invalid property values (i.e., paper size doesn't not exist)
        On Error Resume Next
        Dim Dymo As String
        Dymo = FindDymoPrinter()
        If Dymo = "" Then Exit Function
        SetPrinter(Dymo)
        SetDymoPrinter = True

        Select Case PaperType
            Case Is < 0  ' nothing, ignore...
            Case 1 : Printer.PaperSize = DYMO_PaperSize_ContinuousWide
            Case 2, 30252 : Printer.PaperSize = DYMO_PaperSize_30252
            Case 3, 30323 : Printer.PaperSize = DYMO_PaperSize_30323
            Case Else
                If Printer.PaperSize <> DYMO_PaperSize_30323 Then
                    If Printer.ScaleWidth < 2918 Then
                        Printer.PaperSize = DYMO_PaperSize_30323           ' this should set the paper type to 30323
                    End If
                End If
        End Select
    End Function

    Public Function SetPrinter(ByVal PrinterName As String) As Boolean
        Dim P As Object
        For Each P In Printers
            If P.DeviceName = PrinterName Then      '****finds selected printer
                If Printer.DeviceName <> P.DeviceName Then Printer = P
                SetPrinter = True
                ActiveLog("SetPrinter::EXACT MATCH PrinterName=" & PrinterName & ", DN=" & Printer.DeviceName, 7)
                Exit Function
            End If
        Next
        For Each P In Printers
            If InStr(UCase(P.DeviceName), UCase(PrinterName)) <> 0 Then '****finds near match printer
                If Printer.DeviceName <> P.DeviceName Then Printer = P
                SetPrinter = True
                ActiveLog("SetPrinter::NEAR MATCH PrinterName=" & PrinterName & ", DN=" & Printer.DeviceName, 7)
                Exit Function
            End If
        Next
    End Function

    Public Function NextPrintBreak(ByVal vStr As String, ByVal Start As Integer) As Integer
        Dim tmpBreak As Integer, Absolute As Boolean
        NextPrintBreak = Start + 46
        If InStr(1, Mid(vStr, Start, 46), "MasterCard") Then NextPrintBreak = Start + 54 : Absolute = True

        tmpBreak = InStr(Start + 1, vStr, "I Authorize")

        If tmpBreak > 1 And tmpBreak < NextPrintBreak Then NextPrintBreak = tmpBreak : Absolute = True

        tmpBreak = InStr(Start + 1, vStr, "X _________")
        If tmpBreak > 1 And tmpBreak < NextPrintBreak Then NextPrintBreak = tmpBreak : Absolute = True

        tmpBreak = InStr(Start + 1, vStr, "Approval=")
        If tmpBreak > 1 And tmpBreak < NextPrintBreak Then NextPrintBreak = tmpBreak : Absolute = True

        If Not Absolute Then
            Dim A As Integer, B As Integer
            A = InStr(NextPrintBreak, vStr, " ")
            If A - NextPrintBreak > 10 Then A = 0
            B = InStrRev(vStr, " ", NextPrintBreak)
            If NextPrintBreak - B > 10 Then B = 0
            If (A <> 0 And A <> NextPrintBreak) Or B <> 0 Then
                If B <> 0 And A <> 0 Then
                    If A - NextPrintBreak > NextPrintBreak - B Then A = 0 Else B = 0
                    If B = 0 Then
                        NextPrintBreak = A
                    Else
                        NextPrintBreak = B
                    End If
                End If
            End If
        End If
    End Function

    Public Function FindDymoPrinter(Optional ByVal IgnorePrevious As Boolean = False, Optional ByVal SelectBox As Integer = 1, Optional ByVal StoreNum As Integer = 0, Optional ByVal Current As String = "") As String
        Dim LabelPrinter As String, IsDymo As Boolean, DymoOnly As Boolean
        Dim P As Object
        Dim DN As String, SelList() As Object
        Dim Sel As Integer
        Const None As String = "No label printer."


        If StoreNum <= 0 Then StoreNum = StoresSld

        On Error GoTo CantFind
        FindDymoPrinter = ""
        If Printers.Count = 0 Then Exit Function  ' No printers defined!

        DymoOnly = SelectBox And &H10
        If DymoOnly Then SelectBox = SelectBox Xor &H10  'read and clear the flag

        If Current = "" Then
            LabelPrinter = GetConfigTableValue("Label Printer " & StoreNum, GetCDSSetting("Label Printer " & StoreNum, ""))
            If LabelPrinter = "" Then LabelPrinter = GetConfigTableValue("Label Printer", GetCDSSetting("Label Printer", ""))
        End If
        If LabelPrinter = "" Then LabelPrinter = Current

        If Not IgnorePrevious And LabelPrinter = None Then Exit Function ' They don't want a printer.
        If SelectBox = 0 And LabelPrinter <> "" Then
            FindDymoPrinter = LabelPrinter
            Exit Function
        End If

        If SelectBox <> 0 Then
            ReDim SelList(0)
            SelList(0) = None
        End If

        For Each P In Printers
            DN = P.DeviceName
            IsDymo = IsDymoPrinter(P) Or DN = LabelPrinter
            If SelectBox <> 0 Then
                If IsDymo Or Not DymoOnly Then
                    ReDim Preserve SelList(UBound(SelList) + 1)
                    If DN = Current Then
                        SelList(UBound(SelList)) = "x" & DN
                    Else
                        SelList(UBound(SelList)) = DN
                    End If
                End If
            End If
            If IgnorePrevious Or LabelPrinter = "" Then
                If FindDymoPrinter = "" And IsDymo Then FindDymoPrinter = DN
            Else
                If DN = LabelPrinter Then FindDymoPrinter = DN
            End If
        Next

        If SelectBox = 2 Or (SelectBox = 1 And FindDymoPrinter = "") Then
            Sel = SelectOptionArray("Choose your DYMO Printer", frmSelectOption.ESelOpts.SelOpt_List, SelList, "&Select")
            'Unload frmSelectOption
            frmSelectOption.Close()

            If Sel <= 0 Then FindDymoPrinter = LabelPrinter : Exit Function
            FindDymoPrinter = SelList(Sel - 1)
        End If

        ' we save both w/ and w/o the store designation..
        ' w/o serves as a system default, kinda..
        If SelectBox <> 0 Then ' this should only let it save if they selected...
            On Error GoTo CantSave
            SetConfigTableValue("Label Printer", FindDymoPrinter)
            SetConfigTableValue("Label Printer " & StoreNum, FindDymoPrinter)
        End If

        If FindDymoPrinter = None Then FindDymoPrinter = "" ' Don't return this special value.
        Exit Function
CantFind:
        MessageBox.Show("Couldn't find a DYMO printer.", "Couldn't Find")
        Exit Function
CantSave:
        MessageBox.Show("We could not save your settings." & vbCrLf & "ERR (" & Err.Number & "): " & Err.Description, "Couldn't save")
        Resume Next
    End Function

    Public Function IsDymoPrinter(Optional ByVal Device As Object = Nothing) As Boolean
        Dim T As String
        If IsNothing(Device) Then Device = Printer
        'If IsObject(Device) Then
        If Not Device Is Nothing Then
            If TypeName(Device) = "Printer" Then
                On Error Resume Next
                T = ""
                T = Device.DeviceName
                Device = T
            ElseIf TypeName(Device) = "String" Then
                T = Device
            Else
                Exit Function
            End If
        End If

        If TypeName(Device) = "String" Then
            Device = UCase(Device)
            IsDymoPrinter = (Device Like "*DYMO*" Or Device Like "DYMO*" Or Device Like "*DYMO" Or Device Like "DYMO")
        End If
    End Function

    Public Sub PrintInBox(ByVal PrintOb As Object, ByVal PrintText As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, Optional ByVal FontSize As Integer = -1, Optional ByVal HAlign As AlignConstants = AlignConstants.vbAlignNone, Optional ByVal VAlign As AlignConstants = AlignConstants.vbAlignNone, Optional ByVal BorderStyle As BorderStyleConstants = 0)
        If PrintText <> "" Then
            If FontSize = -1 Then FontSize = 300
            PrintOb.FontSize = BestFontFit(PrintOb, PrintText, FontSize, Width, Height)

            Select Case VAlign
                'Case vbAlignTop
                Case AlignConstants.vbAlignTop
                    PrintOb.CurrentY = Top
                'Case vbAlignBottom
                Case AlignConstants.vbAlignBottom
                    PrintOb.CurrentY = Top + Height - PrintOb.TextHeight(PrintText)
                Case Else ' center
                    PrintOb.CurrentY = Top + (Height - PrintOb.TextHeight(PrintText)) / 2
            End Select

            Dim El As Object
            For Each El In Split(PrintText, vbCrLf)
                Select Case HAlign
                    'Case vbAlignRight, 1
                    Case AlignConstants.vbAlignRight, 1
                        PrintOb.CurrentX = Left + Width - PrintOb.TextWidth(El)
                    'Case vbAlignLeft, 0
                    Case AlignConstants.vbAlignLeft, 0
                        PrintOb.CurrentX = Left
                    Case Else 'center
                        PrintOb.CurrentX = Left + (Width - PrintOb.TextWidth(El)) / 2
                End Select
                PrintOb.Print(El)
            Next
            '    PrintOb.Print PrintText
        End If

        If BorderStyle <> 0 Then
            PrintOb.Line(Left, Top - Left + Width, Top, BorderStyle)
            PrintOb.Line(Left, Top - Left, Top + Height, BorderStyle)
            PrintOb.Line(Left + Width, Top - Left + Width, Top + Height, BorderStyle)
            PrintOb.Line(Left, Top + Height - Left + Width, Top + Height, BorderStyle)
        End If
    End Sub

    Public Sub PrintToPosition(Optional ByVal OutOb As Object = Nothing, Optional ByVal OutText As String = "", Optional ByVal Position As Integer = -1, Optional ByVal Alignment As AlignConstants = AlignConstants.vbAlignLeft, Optional ByVal NewLine As Boolean = False)
        If OutOb Is Nothing Then OutOb = OutputObject
        If IsNothing(OutOb) Then Exit Sub
        If Position = -1 Then Position = OutOb.CurrentX

        Dim TruePos As Integer
        TruePos = Position  ' Already set to exact position.
        If TruePos = 0 And (Alignment = AlignConstants.vbAlignTop Or Alignment = AlignConstants.vbAlignNone) Then TruePos = OutOb.ScaleWidth / 2
        If TruePos <> 0 Then
            Select Case Alignment
                'Case vbAlignRight, vbRightJustify
                Case AlignConstants.vbAlignRight
                    OutOb.CurrentX = TruePos - OutOb.TextWidth(OutText)
                'Case vbAlignTop, vbCenter, 5 ' Center.. hmm.
                Case AlignConstants.vbAlignTop
                    OutOb.CurrentX = TruePos - (OutOb.TextWidth(OutText) / 2)
                Case Else ' Left
                    OutOb.CurrentX = TruePos
            End Select
        End If

        If Not IscPrinter(OutOb) Then
            If NewLine Then
                OutOb.Print(OutText)
            Else
                OutOb.Print(OutText)
            End If
        Else
            If NewLine Then
                OutOb.PrintNL(OutText)
            Else
                OutOb.PrintNNL(OutText)
            End If
        End If
    End Sub

    '<CT> Created this sub to include a parameter CY and CX to accept CurrentY and CurrentX values which are not there in PrintToPostion sub.
    Public Sub PrintToPosition2(Optional ByVal OutOb As Object = Nothing, Optional ByVal OutText As String = "", Optional ByVal Position As Integer = -1, Optional ByVal Alignment As AlignConstants = AlignConstants.vbAlignLeft, Optional ByVal NewLine As Boolean = False, Optional ByVal CY As Integer = 0, Optional ByVal CX As Integer = 0)
        If OutOb Is Nothing Then OutOb = OutputObject
        If IsNothing(OutOb) Then Exit Sub
        If Position = -1 Then Position = OutOb.CurrentX
        If CX > 0 Then
            Position = CX
        End If

        Dim TruePos As Integer
        TruePos = Position  ' Already set to exact position.
        If TruePos = 0 And (Alignment = AlignConstants.vbAlignTop Or Alignment = AlignConstants.vbAlignNone) Then TruePos = OutOb.ScaleWidth / 2
        If TruePos <> 0 Then
            Select Case Alignment
                'Case vbAlignRight, vbRightJustify
                Case AlignConstants.vbAlignRight
                    OutOb.CurrentX = TruePos - OutOb.TextWidth(OutText)
                'Case vbAlignTop, vbCenter, 5 ' Center.. hmm.
                Case AlignConstants.vbAlignTop
                    OutOb.CurrentX = TruePos - (OutOb.TextWidth(OutText) / 2)
                Case Else ' Left
                    OutOb.CurrentX = TruePos
            End Select
        End If

        OutOb.CurrentY = CY
        If Not IscPrinter(OutOb) Then
            If NewLine Then
                OutOb.Print(OutText)
            Else
                OutOb.Print(OutText)
            End If
        Else
            If NewLine Then
                OutOb.PrintNL(OutText)
            Else
                OutOb.PrintNNL(OutText)
            End If
        End If
    End Sub
    '</CT>

    Public Function IscPrinter(ByVal Ob As Object) As Boolean
        IscPrinter = TypeName(Ob) = "cPrinter"
    End Function

    Public Function BestFontFit(ByVal OutOb As Object, ByRef OutText As String, ByRef MaxFontSize As Integer, ByRef MaxWidth As Integer, ByRef MaxHeight As Integer) As Double
        On Error GoTo NoFit
        OutOb.FontSize = MaxFontSize
        Do While OutOb.TextHeight(OutText) > MaxHeight Or OutOb.TextWidth(OutText) > MaxWidth
            '    Debug.Print MaxFontSize, OutOb.TextHeight(OutText), OutOb.TextWidth(OutText)
            MaxFontSize = MaxFontSize - 1  ' Adjust this variable, in case printer.fontsize fails to set.  This can happen for unsupported font sizes.
            OutOb.FontSize = MaxFontSize
        Loop
        BestFontFit = OutOb.FontSize
        Exit Function
NoFit:
        If Err.Number = 6 Then Resume Next
        ' With any luck, the printer will still be set to a reasonable font.
        BestFontFit = OutOb.FontSize
    End Function

    Public Sub PrintCentered(ByVal Text As String, Optional ByVal yPos As Integer = -1, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
        Dim Ob As Boolean, oI As Boolean
        Dim OO As Object

        On Error Resume Next
        If OutputObject Is Nothing Then OutputObject = Printer
        OO = OutputObject
        Ob = OO.FontBold
        oI = OO.FontItalic
        If Bold Then OO.FontBold = True
        If Italic Then OO.FontItalic = True
        If yPos > 0 Then OO.CurrentY = yPos


        OO.CurrentX = (OutputObject.ScaleWidth - OO.TextWidth(Text)) / 2 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
        OO.CurrentX = (Printer.ScaleWidth - OO.TextWidth(Text)) / 2 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
        If Not IscPrinter(OutputObject) Then
            OutputObject.Print(Text)
        Else
            OutputObject.PrintNL(Text)
        End If

        If Bold Then OO.FontBold = Ob
        If Italic Then OO.FontItalic = oI
    End Sub

    Public Sub PrintLine(Optional ByRef X1 As Integer = -1, Optional ByRef Y1 As Integer = -1, Optional ByRef X2 As Integer = -1, Optional ByRef Y2 As Integer = -1)
        If X1 = -1 Then X1 = 0
        If Y1 = -1 Then Y1 = Printer.CurrentY
        If X2 = -1 Then X2 = Printer.ScaleWidth
        If Y2 = -1 Then Y2 = Printer.CurrentY
        'Printer.Line(X1, Y1)- (X2, Y2)
        Printer.Line(X2, Y2)
    End Sub

    ' Generic method to send text to printer or print preview.
    Public Sub PrintTo(Optional ByVal OutOb As Object = Nothing, Optional ByVal OutText As Object = Nothing, Optional ByVal Position As Integer = -1, Optional ByVal Alignment As AlignConstants = AlignConstants.vbAlignLeft, Optional ByVal NewLine As Boolean = False)
        If OutOb Is Nothing Then OutOb = OutputObject
        If Position = -1 Then Position = OutOb.CurrentX Else Position = Position * 80
        PrintToPosition(OutOb, OutText, Position, Alignment, NewLine)
    End Sub

    '    Public Property Get LegalContractPrinter(Optional ByVal StoreNo as integer = 0) As String
    '    Dim F As String, F1 As String
    '  F1 = "Legal Contract Printer"
    '  If StoreNo = 0 Then StoreNo = StoresSld
    '  F = IIf(StoreNo <= 1, F1, F1 & " " & StoreNo)
    '  LegalContractPrinter = GetConfigTableValue(F, "")
    'End Property
    '    Public Property Let LegalContractPrinter(Optional ByVal StoreNo as integer = 0, ByVal nValue As String)
    '    Dim F As String, F1 As String
    '  F1 = "Legal Contract Printer"
    '  If StoreNo = 0 Then StoreNo = StoresSld
    '  F = IIf(StoreNo <= 1, F1, F1 & " " & StoreNo)
    '  SetConfigTableValue F, nValue
    '    End Property

    Public Property LegalContractPrinter(Optional ByVal StoreNo As Integer = 0) As String
        Get
            Dim F As String, F1 As String
            F1 = "Legal Contract Printer"
            If StoreNo = 0 Then StoreNo = StoresSld
            F = IIf(StoreNo <= 1, F1, F1 & " " & StoreNo)
            LegalContractPrinter = GetConfigTableValue(F, "")
        End Get
        Set(value As String)
            Dim F As String, F1 As String
            F1 = "Legal Contract Printer"
            If StoreNo = 0 Then StoreNo = StoresSld
            F = IIf(StoreNo <= 1, F1, F1 & " " & StoreNo)
            SetConfigTableValue(F, value)
        End Set
    End Property

    Public Function PrinterSetupDialog(ByVal Loc As Form, ByRef DeviceName As String, ByRef Port As String, Optional ByRef doSet As Boolean = True, Optional ByVal Flags As Integer = VBPrinterConstants.cdlPDHidePrintToFile Or VBPrinterConstants.cdlPDNoPageNums Or VBPrinterConstants.cdlPDNoSelection Or VBPrinterConstants.cdlPDUseDevModeCopies Or VBPrinterConstants.cdlPDPrintSetup, Optional ByVal Min As Integer = 0, Optional ByVal Max As Integer = 0, Optional ByRef FromPage As Integer = 0, Optional ByRef ToPage As Integer = 0, Optional ByRef Copies As Integer = 0, Optional ByRef Orientation As Integer = 0) As Boolean
        Dim P As New PrinterDlg, Pri As Printer

        On Error Resume Next
        ' BFH20050107
        ' it did this using com dlg, but we're not using that anymore..
        ' cant find a similar way to do this, so i'm leaving it out.. its fairly trivial
        '   ComDlg.DialogTitle = MsgTitle$   'Set the title of the dialog box
        If Min <> 0 Then P.Min = Min
        If Max <> 0 Then P.Max = Max
        P.Flags = Flags

        P.PrinterName = Printer.DeviceName
        'P.DriverName = Printer.DriverName
        'P.Port = Printer.Port
        P.CancelError = True


        On Error GoTo PrinterDialogCancelled
        'P.ShowPrinter Loc.hwnd
        P.ShowPrinter(Loc.Handle)

        DeviceName = P.PrinterName
        Port = P.Port
        If P.Flags And VBPrinterConstants.cdlPDPageNums Then
            FromPage = P.FromPage
            ToPage = P.ToPage
        ElseIf P.Flags And VBPrinterConstants.cdlPDSelection Then
            ' no change
        Else
            FromPage = P.Min
            ToPage = P.Max
        End If
        Copies = P.Copies
        Orientation = P.Orientation

        P = Nothing

        On Error Resume Next
        If doSet Then
            SetPrinter(DeviceName)
            Printer.Orientation = Orientation
        End If
        PrinterSetupDialog = True
        Exit Function

PrinterDialogCancelled:
        Err.Clear()
    End Function

    Public Sub PrintOut(Optional ByVal X As Single = -1, Optional ByVal Y As Single = -1, Optional ByVal Text As String = "",
   Optional ByVal XCenter As Boolean = False _
  , Optional ByVal FontName As String = "", Optional ByVal FontBold As Boolean = False, Optional ByVal FontSize As String = "" _
  , Optional ByVal DrawWidth As Single = -1, Optional ByVal NewPage As Boolean = False, Optional ByVal BlankLines As Integer = -1 _
  , Optional ByVal Orientation As Integer = -1, Optional ByVal OutObj As Object = Nothing)

        Dim I As Integer
        If Not OutputToPrinter And OutObj Is Nothing Then OutObj = OutputObject
        If OutObj Is Nothing Then OutObj = Printer
        If NewPage Then
            If OutputToPrinter Then OutObj.NewPage Else frmPrintPreviewDocument.NewPage()
        End If
        If FontName <> "" Then OutObj.FontName = FontName
        If FontSize <> "" Then OutObj.FontSize = FontSize
        OutObj.FontBold = FontBold
        If Orientation <> -1 Then OutObj.Orientation = Orientation
        If X <> -1 Then OutObj.CurrentX = X
        If Y <> -1 Then OutObj.CurrentY = Y
        If DrawWidth <> -1 Then OutObj.DrawWidth = DrawWidth
        If XCenter Then OutObj.CurrentX = IIf(X <= 0, Printer.Width / 2, X) - OutObj.TextWidth(Trim(Text)) / 2
        If Text <> "" Then
            If Not IscPrinter(OutObj) Then
                OutObj.Print(Text)
            Else
                OutObj.PrintNL(Text)
            End If
        End If

        If (BlankLines <> -1) Then
            For I = 1 To BlankLines : OutObj.Print("") : Next
        End If
    End Sub

    Public Sub PrintPageOverflowIndicator()
        Dim Tx As Integer, Ty As Integer, Tds As Integer, Tdw As Integer

        On Error Resume Next
        If Not OutputToPrinter Then
            Tx = OutputObject.CurrentX
            Ty = OutputObject.CurrentY
            Tds = OutputObject.DrawStyle
            Tdw = OutputObject.DrawWidth
            OutputObject.DrawStyle = vbDot
            OutputObject.DrawWidth = 1
            'OutputObject.Line(Printer.ScaleWidth, 0)-(Printer.ScaleWidth, Printer.ScaleHeight)
            OutputObject.Line(Printer.ScaleWidth, 0, Printer.ScaleWidth, Printer.ScaleHeight)
            'OutputObject.Line(0, Printer.ScaleHeight)-(Printer.ScaleWidth, Printer.ScaleHeight)
            OutputObject.Line(0, Printer.ScaleHeight, Printer.ScaleWidth, Printer.ScaleHeight)
            OutputObject.DrawStyle = Tds
            OutputObject.DrawWidth = Tdw

            OutputObject.CurrentX = Tx
            OutputObject.CurrentY = Ty
        End If
    End Sub

    Public Sub Printer_Location(ByRef X As Single, ByRef Y As Single, ByRef FontSize As Single, Optional ByRef Prt As String = "")
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.FontSize = FontSize
        If Len(Prt) <> 0 Then Printer.Print(Prt)
    End Sub

    Public Sub PrintAligned(ByVal Text As String, Optional ByVal Align As Byte = AlignmentConstants.vbLeftJustify, Optional ByVal Location As Integer = 0, Optional ByVal yPos As Integer = -1, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
        Dim List() As String, X As Integer
        Dim Ob As Boolean, oI As Boolean
        Dim OO As Object

        OutputObject = Printer
        OO = OutputObject

        Ob = OO.FontBold
        oI = OO.FontItalic
        If Bold Then OO.FontBold = True
        If Italic Then OO.FontItalic = True
        If yPos > 0 Then OO.CurrentY = yPos

        If InStr(Text, vbCrLf) Then 'Multi-line string
            'ReDim List(UBound(List) + 1)
            'List() = Split(Text, vbCrLf)
            List = Split(Text, vbCrLf)
            For X = LBound(List) To UBound(List)
                Select Case Align
                    Case AlignmentConstants.vbLeftJustify, AlignConstants.vbAlignLeft '0
                        OO.CurrentX = Location
                    Case AlignmentConstants.vbRightJustify, AlignConstants.vbAlignRight '1
                        OO.CurrentX = Printer.ScaleWidth - OO.TextWidth(List(X)) + Location
                    Case AlignmentConstants.vbCenter '2
                        OO.CurrentX = (Printer.Width - OO.TextWidth(List(X))) / 2 + Location 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
                    Case Else
                        Debug.Print("UNKNOWN ALIGNMENT CONSTANT IN modPrintToolsCommon.PrintAligned: " & CStr(Align))
                End Select
                OutputObject.Print(List(X))
            Next
        Else 'Single-line string
            Select Case Align
                Case AlignmentConstants.vbLeftJustify, AlignConstants.vbAlignLeft '0
                    OO.CurrentX = Location
                Case AlignmentConstants.vbRightJustify, AlignConstants.vbAlignRight '1
                    If Location <= 0 Then
                        OO.CurrentX = Printer.ScaleWidth - OO.TextWidth(Text)
                    Else
                        OO.CurrentX = (Location) - OO.TextWidth(Text)
                    End If
                Case AlignmentConstants.vbCenter '2
                    OO.CurrentX = (Printer.Width - OO.TextWidth(Text)) / 2 + Location 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
                Case Else
                    Debug.Print("UNKNOWN ALIGNMENT CONSTANT IN modPrintToolsCommon.PrintAligned: " & CStr(Align))
            End Select
            If Not IscPrinter(OutputObject) Then
                OutputObject.Print(Text)
            Else
                OutputObject.PrintNL(Text)
            End If
        End If

        If Bold Then OO.FontBold = Ob         ' this is to preserve the original settings
        If Italic Then OO.FontItalic = oI
    End Sub

    '<CT>Created this PrintAligned2 as an alternative to PrintAligned sub. Cause OO as object, OO.FontBold, OO.TextWidth etc with as object will not work in vb.net </CT>
    Public Sub PrintAligned2(ByVal Text As String, Optional ByVal Align As Byte = AlignmentConstants.vbLeftJustify, Optional ByVal Location As Integer = 0, Optional ByVal yPos As Integer = -1, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
        Dim List() As String, X As Integer
        Dim Ob As Boolean, oI As Boolean
        Dim OO As Object
        OO = OutputObject

        'Ob = OO.FontBold
        Ob = Printer.FontBold
        'oI = OO.FontItalic
        oI = Printer.FontItalic

        'If Bold Then OO.FontBold = True
        If Bold Then Printer.FontBold = True
        'If Italic Then OO.FontItalic = True
        If Italic Then Printer.FontItalic = True
        'If yPos > 0 Then OO.CurrentY = yPos
        If yPos > 0 Then Printer.CurrentY = yPos

        If InStr(Text, vbCrLf) Then 'Multi-line string
            'ReDim List(UBound(List) + 1)
            'List() = Split(Text, vbCrLf)
            List = Split(Text, vbCrLf)
            For X = LBound(List) To UBound(List)
                Select Case Align
                    Case AlignmentConstants.vbLeftJustify, AlignConstants.vbAlignLeft '0
                        'OO.CurrentX = Location
                        Printer.CurrentX = Location
                    Case AlignmentConstants.vbRightJustify, AlignConstants.vbAlignRight '1
                        'OO.CurrentX = Printer.ScaleWidth - OO.TextWidth(List(X)) + Location
                        Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(List(X)) + Location
                    Case AlignmentConstants.vbCenter '2
                        'OO.CurrentX = (Printer.Width - OO.TextWidth(List(X))) / 2 + Location 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
                        Printer.CurrentX = (Printer.Width - Printer.TextWidth(List(X))) / 2 + Location 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
                    Case Else
                        Debug.Print("UNKNOWN ALIGNMENT CONSTANT IN modPrintToolsCommon.PrintAligned: " & CStr(Align))
                End Select
                'OutputObject.Print(List(X))
                Printer.Print(List(X))
            Next
        Else 'Single-line string
            Select Case Align
                Case AlignmentConstants.vbLeftJustify, AlignConstants.vbAlignLeft '0
                    'OO.CurrentX = Location
                    Printer.CurrentX = Location
                Case AlignmentConstants.vbRightJustify, AlignConstants.vbAlignRight '1
                    If Location <= 0 Then
                        'OO.CurrentX = Printer.ScaleWidth - OO.TextWidth(Text)
                        Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(Text)
                    Else
                        'OO.CurrentX = (Location) - OO.TextWidth(Text)
                        Printer.CurrentX = (Location) - Printer.TextWidth(Text)
                    End If
                Case AlignmentConstants.vbCenter '2
                    'OO.CurrentX = (Printer.Width - OO.TextWidth(Text)) / 2 + Location 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(Text)) / 2 + Location 'Use Printer.Width to keep the heading centered even with larger screen resolutions.
                Case Else
                    Debug.Print("UNKNOWN ALIGNMENT CONSTANT IN modPrintToolsCommon.PrintAligned: " & CStr(Align))
            End Select
            If Not IscPrinter(OutputObject) Then
                'OutputObject.Print(Text)
                Printer.Print(Text)
            Else
                'OutputObject.PrintNL(Text)
                Printer.Print(Text)
            End If
        End If

        'If Bold Then OO.FontBold = Ob         ' this is to preserve the original settings
        If Bold Then Printer.FontBold = Ob         ' this is to preserve the original settings
        'If Italic Then OO.FontItalic = oI
        If Italic Then Printer.FontItalic = oI
    End Sub

    Public Sub CommonReportAddColumn(Optional ByVal ColumnHeader As String = "", Optional ByVal Width As Integer = 0, Optional ByVal Reset As Boolean = False, Optional ByRef OptionString As String = "")
        On Error Resume Next
        If Reset Then
            ColCount = 1
            CommonReportColSpacing = 150
            CommonReportIndent = 100
        Else
            ColCount = ColCount + 1
        End If

        'Cols(ColCount, 1) = ColumnHeader
        Cols(ColCount - 1, 1 - 1) = ColumnHeader
        If ColCount = 1 Then
            'Cols(ColCount, 2) = CommonReportIndent
            Cols(ColCount - 1, 2 - 1) = CommonReportIndent
        Else
            'Cols(ColCount, 2) = Cols(ColCount - 1, 3) + CommonReportColSpacing
            Cols(ColCount - 1, 2 - 1) = Cols(ColCount - 2, 3 - 1) + CommonReportColSpacing
        End If
        'Cols(ColCount, 3) = Cols(ColCount, 2) + Width ' - CommonReportColSpacing
        Cols(ColCount - 1, 3 - 1) = Cols(ColCount - 1, 2 - 1) + Width ' - CommonReportColSpacing
        'Cols(ColCount, 4) = OptionString
        Cols(ColCount - 1, 4 - 1) = OptionString
    End Sub

    Public Function CommonReportHeader(ByRef ReportName As String, Optional ByRef PageNum As Integer = 1, Optional ByRef PageCount As Integer = 0, Optional ByVal OptionString As String = "") As Integer
        Dim ColumnHeadY As Integer, I As Integer, ColFormat As String, ColCap As String
        OptionString = UCase(OptionString)

        OutputObject.FontSize = 18
        PrintCentered(ReportName, 100, True)
        OutputObject.FontSize = 14
        PrintCentered(StoreSettings.Name, , True)
        PrintCentered(StoreSettings.Address, , True)
        PrintCentered(StoreSettings.City, , True)
        OutputObject.FontSize = 10
        PrintCentered("")
        ColumnHeadY = OutputObject.CurrentY

        PrintAligned("Page: " & PageNum & IIf(PageCount > 0, " of " & PageCount, ""), , 9500, 100)
        If Not OptionString Like "*NODATE*" Then PrintAligned("Date: " & Today, , 100, 100)

        OutputObject.CurrentY = ColumnHeadY

        If Not OptionString Like "*NOHEAD*" Then
            ' print column headers
            For I = 1 To ColCount
                ColCap = Cols(I, 1)
                ColFormat = "." & UCase(Cols(I, 4)) & "." ' extra stuff in case Like's * operator is picky
                If ColFormat Like "*CURRENCY*" Then
                    PrintAligned(ColCap, AlignmentConstants.vbRightJustify, ReportCol(I, True), ColumnHeadY, True)
                ElseIf ColFormat Like "*RIGHT*" Then
                    PrintAligned(ColCap, AlignmentConstants.vbRightJustify, ReportCol(I, True), ColumnHeadY, True)
                Else
                    PrintAligned(ColCap, AlignmentConstants.vbLeftJustify, ReportCol(I), ColumnHeadY, True)
                End If
            Next
            ColumnHeadY = OutputObject.CurrentY
            OutputObject.DrawWidth = 2
            'OutputObject.Line(ReportCol(1), ColumnHeadY)-(ReportCol(ColCount, True), ColumnHeadY)
            OutputObject.line(ReportCol(1), ColumnHeadY, ReportCol(ColCount, True), ColumnHeadY)
            OutputObject.DrawWidth = 1
        End If
        CommonReportHeader = OutputObject.CurrentY

        OutputObject.CurrentY = CommonReportHeader
    End Function

    Public Sub CommonReportPrintColumn(ByVal ColNum As Integer, ByVal ColVal As String, Optional ByRef Y As Integer = -1)
        Dim ColFormat As String, X As Integer, W As Integer, YY As Integer, DoRight As Boolean
        YY = IIf(Y < 0, OutputObject.CurrentY, Y)

        ColFormat = "." & UCase(Cols(ColNum, 4)) & "."
        If ColFormat Like "*CURRENCY*" Then
            DoRight = True
            ColVal = FormatCurrency(GetPrice(ColVal))
        ElseIf ColFormat Like "*RIGHT*" Then
            DoRight = True
        Else
            DoRight = (Left(ColVal, 1) = "$") ' Or IsNumeric(Val)
        End If
        X = ReportCol(ColNum, DoRight)
        W = ReportCol(ColNum, True) - ReportCol(ColNum)

        If ColFormat Like "*TRUNC*" Then
            Do While OutputObject.TextWidth(ColVal) > W
                ColVal = Left(ColVal, Len(ColVal) - 1)
            Loop
        End If
        PrintAligned(ColVal, IIf(DoRight, AlignConstants.vbAlignRight, AlignConstants.vbAlignLeft), X, YY)
    End Sub

    Private Function ReportCol(ByVal Which As Object, Optional ByVal Right As Boolean = False) As Integer
        Dim I As Integer
        If TypeName(Which) = "String" Then
            For I = 1 To ColCount
                If LCase(CStr(Which)) = LCase(CStr(Cols(I, 1))) Then
                    Which = I
                    Exit For
                End If
            Next
            If TypeName(Which) = "String" Then
                Debug.Print("modPrintToolsCommon.ReportCol - Invalid Field Name: " & CStr(Which))
                Exit Function
            End If
        End If
        If Which <= 0 Or Which > ColCount Then
            Debug.Print("modPrintToolsCommon.ReportCol - Invalid Field Name: " & CStr(Which))
            Exit Function
        End If

        ReportCol = Cols(Which, IIf(Right, 3, 2))
    End Function

    Public Function DescribeOutputObject() As String
        On Error Resume Next
        If OutputObject Is Nothing Then
            DescribeOutputObject = "[NOTHING]"
        Else
            If DescribeOutputObject = "" Then DescribeOutputObject = OutputObject.DeviceName
            If DescribeOutputObject = "" Then DescribeOutputObject = OutputObject.Name
            If OutputObject = Printer Then DescribeOutputObject = DescribeOutputObject & "[PRINTER]"
        End If
    End Function

    Public Function DescribeScaleMode(ByVal sC As Integer) As String
        Select Case sC
            Case vbCentimeters : DescribeScaleMode = "vbCentimeters - 7"
            Case vbCharacters : DescribeScaleMode = "vbCharacters - 4"
            Case VBRUN.ScaleModeConstants.vbContainerPosition : DescribeScaleMode = "vbCharacterPosition - 9"
            Case VBRUN.ScaleModeConstants.vbContainerSize : DescribeScaleMode = "vbContainerSize - 10"
            Case vbHimetric : DescribeScaleMode = "vbHimetric - 8"
            Case vbInches : DescribeScaleMode = "vbInches - 5"
            Case vbMillimeters : DescribeScaleMode = "vbMillimeters - 6"
            Case vbPixels : DescribeScaleMode = "vbPixels - 3"
            Case vbPoints : DescribeScaleMode = "vbPoints - 2"
            Case vbTwips : DescribeScaleMode = "vbTwips - 1"
            Case vbUser : DescribeScaleMode = "vbUser - 0"
            Case Else : DescribeScaleMode = "[UNKNOWN]"
        End Select
    End Function

    ' Generic method to print at a specified Tab(..) location..  corrects alignment for BOLD
    '<CT> Added three optional parameters (CY, CX and PP2) to PrintToTab sub</CT>
    Public Sub PrintToTab(Optional ByVal OutOb As Object = Nothing, Optional ByVal OutText As String = "", Optional ByVal TabLoc As Integer = 0, Optional ByVal Alignment As AlignConstants = VBRUN.AlignConstants.vbAlignLeft, Optional ByVal NewLine As Boolean = False, Optional ByVal CY As Integer = 0, Optional CX As Integer = 0, Optional ByVal PP2 As Boolean = False)
        Dim X As Boolean
        If OutOb Is Nothing Then OutOb = OutputObject
        X = OutOb.FontBold
        OutOb.FontBold = False
        If Not IscPrinter(OutOb) Then
            OutOb.Print(TAB(TabLoc))
        Else
            OutOb.CurrentX = OutOb.TextWidth("X") * TabLoc
        End If
        OutOb.FontBold = X
    End Sub

    Public Function PrinterStat(Optional ByVal vDeviceName As String = "") As Boolean
        Dim Op As String
        Dim S As String

        Op = Printer.DeviceName
        If vDeviceName <> "" Then SetPrinter(vDeviceName)

        S = ""
        S = S & "== Printer Info ==" & vbCrLf
        S = S & "dev: " & Printer.DeviceName & IIf(IsDymoPrinter(), " [DYMO]", "") & vbCrLf
        'S = S & "drv: " & Printer.DriverName & vbCrLf
        S = S & vbCrLf
        S = S & "dim: " & Printer.Width & "x" & Printer.Height & vbCrLf
        S = S & "scl: " & Printer.ScaleWidth & "x" & Printer.ScaleHeight & vbCrLf
        S = S & "pos: " & Printer.CurrentX & "x" & Printer.CurrentY & vbCrLf
        S = S & vbCrLf
        S = S & "ori: " & Printer.Orientation & " [" & DescribePrinterOrientation(Printer.Orientation) & "]" & vbCrLf
        S = S & "bin: " & Printer.PaperBin & " [" & DescribePrinterPaperBin(Printer.PaperBin) & "]" & vbCrLf
        S = S & "pap: " & Printer.PaperSize & " [" & DescribePrinterPaperSize(Printer.PaperSize) & "]" & vbCrLf
        S = S & "pqa: " & Printer.PrintQuality & " [" & DescribePrinterPrintQuality(Printer.PrintQuality) & "]" & vbCrLf
        'S = S & "zoo: " & Printer.Zoom & vbCrLf
        S = S & vbCrLf
        S = S & "dpx: " & Printer.Duplex & " [" & DescribePrinterDuplex(Printer.Duplex) & "]" & vbCrLf
        S = S & "clr: " & Printer.ColorMode & vbCrLf
        S = S & vbCrLf
        S = S & "fnt: " & Printer.FontName & " " & Printer.FontSize & IIf(Printer.FontBold, " [BOLD]", "") & IIf(Printer.FontItalic, " [ITAL]", "") & IIf(Printer.FontStrikethru, " [STRI]", "") & IIf(Printer.FontTransparent, " [TRANS]", "") & IIf(Printer.FontUnderline, " [UNDL]", "") & vbCrLf
        S = S & "fcl: " & Printer.ForeColor & " [" & DescribeColor(Printer.ForeColor) & "]" & vbCrLf
        S = S & "  Drawing" & vbCrLf
        'S = S & "mde: " & Printer.DrawMode & vbCrLf
        S = S & "sty: " & Printer.DrawStyle & vbCrLf
        S = S & "wid: " & Printer.DrawWidth & vbCrLf
        S = S & "fst: " & Printer.FillStyle & vbCrLf
        S = S & "fcl: " & Printer.FillColor & " [" & DescribeColor(Printer.FillColor) & "]" & vbCrLf
        '  S = S & vbCrLf

        MessageBox.Show(S, "PrinterStat", MessageBoxButtons.OK, MessageBoxIcon.Information)

        If vDeviceName <> "" Then SetPrinter(Op)

        PrinterStat = True
    End Function

    Public Function DescribePrinterOrientation(ByVal Orien As Integer) As String
        DescribePrinterOrientation = IIf(Orien = vbPRORLandscape, "Landscape", "Portrait")
    End Function

    Public Function DescribePrinterPaperBin(ByVal pBin As Integer) As String
        Select Case pBin
            Case vbPRBNAuto : DescribePrinterPaperBin = "vbPRBNAuto"
            Case vbPRBNCassette : DescribePrinterPaperBin = "vbPRBNCassette"
            Case vbPRBNEnvelope : DescribePrinterPaperBin = "vbPRBNEnvelope"
            Case vbPRBNEnvManual : DescribePrinterPaperBin = "vbPRBNEnvManual"
            Case vbPRBNLargeCapacity : DescribePrinterPaperBin = "vbPRBNLargeCapacity"
            Case vbPRBNLargeFmt : DescribePrinterPaperBin = "vbPRBNLargeFmt"
            Case vbPRBNLower : DescribePrinterPaperBin = "vbPRBNLower"
            Case vbPRBNManual : DescribePrinterPaperBin = "vbPRBNManual"
            Case vbPRBNMiddle : DescribePrinterPaperBin = "vbPRBNMiddle"
            Case vbPRBNSmallFmt : DescribePrinterPaperBin = "vbPRBNSmallFmt"
            Case vbPRBNTractor : DescribePrinterPaperBin = "vbPRBNTractor"
            Case vbPRBNUpper : DescribePrinterPaperBin = "vbPRBNUpper"

            Case DYMO_PaperBin_Left : DescribePrinterPaperBin = "DYMO - Left Roll"
            Case DYMO_PaperBin_Right : DescribePrinterPaperBin = "DYMO - Right Roll ?"

            Case 258 : DescribePrinterPaperBin = "Unknown - 258"

            Case Else : DescribePrinterPaperBin = "UNKNOWN"
        End Select
    End Function

    Public Function DescribePrinterPaperSize(ByVal pSize As Integer) As String
        Select Case pSize
            Case vbPRPSLetter : DescribePrinterPaperSize = "8.5"" x 11"""
            Case vbPRPSA3 : DescribePrinterPaperSize = "A3"
            Case vbPRPSA4 : DescribePrinterPaperSize = "A4"
            Case vbPRPSA5 : DescribePrinterPaperSize = "A5"
            Case vbPRPSLegal : DescribePrinterPaperSize = "Legal"
            Case DYMO_PaperSize_30252 : DescribePrinterPaperSize = "DYMO - 6"" x 1.5"" Addr Labels (NARROW)"
            Case DYMO_PaperSize_30323 : DescribePrinterPaperSize = "DYMO - 6"" x 3"" Addr Labels (Wide)"
            Case DYMO_PaperSize_30256 : DescribePrinterPaperSize = "DYMO - 2"" x 4"" Shipping"
            Case DYMO_PaperSize_ContW : DescribePrinterPaperSize = "Continuous, Wide"
            Case DYMO_PaperSize_30323 : DescribePrinterPaperSize = "DYMO - 6"" x 3"" labels"

            Case Else : DescribePrinterPaperSize = "UNKNOWN"
        End Select
    End Function

    'vbPRPQDraft, vbPRPQLow, PRPQMedium, PRPQHigh
    Public Function DescribePrinterPrintQuality(ByVal pPrintQuality As Integer) As String
        Select Case pPrintQuality
            Case vbPRPQDraft : DescribePrinterPrintQuality = "vbPRPQDraft"
            Case vbPRPQLow : DescribePrinterPrintQuality = "vbPRPQLow"
            Case vbPRPQMedium : DescribePrinterPrintQuality = "vbPRPQMedium"
            Case vbPRPQHigh : DescribePrinterPrintQuality = "vbPRPQHigh"

            Case Else : DescribePrinterPrintQuality = "UNKNOWN"
        End Select
    End Function

    Public Function DescribePrinterDuplex(ByVal pDuplex As Integer) As String
        Select Case pDuplex
            Case vbPRDPSimplex : DescribePrinterDuplex = "vbPRDPSimplex"
            Case vbPRDPVertical : DescribePrinterDuplex = "vbPRDPVertical"
            Case vbPRDPHorizontal : DescribePrinterDuplex = "vbPRDPHorizontal"

            Case Else : DescribePrinterDuplex = "UNKNOWN"
        End Select
    End Function

    Public Function TicketPrinter(Optional ByVal StoreNo As Integer = 0, Optional ByVal nValue As String = "", Optional ByVal GetOrLet As String = "Get")
        If GetOrLet = "Get" Then
            'Get
            If StoreNo = 0 Then StoreNo = StoresSld
            TicketPrinter = GetConfigTableValue("Ticket Printer", GetCDSSetting("Cash Register Printer"))
        Else
            'Let
            If StoreNo = 0 Then StoreNo = StoresSld
            SetConfigTableValue("Ticket Printer", nValue)        ' right now, we save to Config table (primary)
            SaveCDSSetting("Ticket Printer", nValue)             ' and registry (in case we want to use it in the future)
        End If
    End Function

    Public Sub PrintCostCode(ByRef Plaintext As String, ByRef TagSize As String, Optional ByRef PrintX As Integer = -1, Optional ByRef PrintY As Integer = -1)
        ' Encode and print PlainText.
        ' Encoded letters need to be in a different font than plain ones.

        Dim OldFontSize As Double
        Dim I As Integer
        Dim CodeLetter As String
        Dim S As String
        CodeLetter = ""

        OldFontSize = Printer.FontSize
        'If PrintX <> -1 Then Printer.CurrentX = PrintX
        'If PrintY <> -1 Then Printer.CurrentY = PrintY

        SetCostCodeFontSize(TagSize, True) ' CodeLetter <> Mid(Plaintext, I, 1)
        S = ConvertCostToCode(Plaintext)
        If PrintX <> -1 Then Printer.CurrentX = PrintX
        If PrintY <> -1 Then Printer.CurrentY = PrintY
        Printer.Print(S)

        '  For I = 1 To Len(Plaintext)
        '    CodeLetter = GetCostCode(Mid(Plaintext, I, 1))
        '    SetCostCodeFontSize TagSize, CodeLetter <> Mid(Plaintext, I, 1)
        '    Printer.Print CodeLetter;
        ' ' Debug.Print CodeLetter;
        '  Next
        Printer.FontSize = OldFontSize
        'Printer.Print()   ' Next line.
    End Sub

    Private Sub SetCostCodeFontSize(ByRef TagSize As String, ByRef Changed As Boolean)
        Dim SizeDiff As Integer
        If Changed Then SizeDiff = 2 Else SizeDiff = 0
        Select Case TagSize
            Case "DYMO", "SMALL" : Printer.FontSize = 9 + SizeDiff
            Case "MED" : Printer.FontSize = 10 + SizeDiff
            Case "LARGE" : Printer.FontSize = 14 + SizeDiff
        End Select
    End Sub

    Public Function BestPrinterFontFit(ByRef OutText As String, ByRef MaxFontSize As Integer, ByRef MaxWidth As Integer, ByRef MaxHeight As Integer) As Double
        BestPrinterFontFit = BestFontFit(Printer, OutText, MaxFontSize, MaxWidth, MaxHeight)
    End Function

    'bfh20051216 - added config table handling as primary interface for this value
    ' registry settings seemed to be misbehaving for United under a terminal-services environment
    ' config table should behave the same but not have any permissions/security issues
    ' so long as the program has been running smoothly already
    'Public ReadOnly Property CashRegisterPrinter(Optional ByVal StoreNo As Integer = 0) As String
    '    Get
    '        Dim F As String, F1 As String
    '        F1 = "Cash Register Printer"
    '        If StoreNo = 0 Then StoreNo = StoresSld
    '        F = IIf(StoreNo <= 1, F1, F1 & " " & StoreNo)
    '        CashRegisterPrinter = GetConfigTableValue(F, GetCDSSetting(F1))
    '    End Get
    'End Property
    Public Function CashRegisterPrinter(Optional ByVal StoreNo As Integer = 0, Optional ByVal nValue As String = "")
        If nValue = "" Then
            'Get
            Dim F As String, F1 As String
            F1 = "Cash Register Printer"
            If StoreNo = 0 Then StoreNo = StoresSld
            F = IIf(StoreNo <= 1, F1, F1 & " " & StoreNo)
            CashRegisterPrinter = GetConfigTableValue(F, GetCDSSetting(F1))
        Else
            'Let
            Dim F As String, F1 As String
            F1 = "Cash Register Printer"
            If StoreNo = 0 Then StoreNo = StoresSld
            F = IIf(StoreNo <= 1, F1, F1 & " " & StoreNo)
            SetConfigTableValue(F, nValue)
            SaveCDSSetting(F1, nValue)
        End If
    End Function

    Public Sub PrintAutoMailingLetterHeader(ByVal Name1 As String, ByVal addr1 As String, ByVal City1 As String, ByVal Tele1 As String, ByVal Name2 As String, ByVal addr2 As String, ByVal City2 As String, ByVal Tele2 As String, Optional ByVal ShowDate As Boolean = True)
        Dim oFN As String, oFS As Integer, oDW As Integer


        OutputObject = Printer

        oFN = OutputObject.FontName
        oFS = OutputObject.FontSize
        oDW = OutputObject.DrawWidth
        OutputObject.FontName = "Arial"
        OutputObject.FontSize = 12
        OutputObject.DrawWidth = 2

        If ShowDate Then PrintAligned(Today, , 8000, 1020, True)

        PrintAligned(Name1, , , 1020, True)
        PrintAligned(addr1,,,, True)
        PrintAligned(City1,,,, True)
        PrintAligned(Tele1,,,, True)

        PrintAligned(Name2, , , 3170, True)
        PrintAligned(addr2,,,, True)
        PrintAligned(City2,,,, True)
        PrintAligned(Tele2,,,, True)

        OutputObject.FontName = oFN
        OutputObject.FontSize = oFS
        OutputObject.DrawWidth = oDW
        OutputObject.CurrentY = 4514
    End Sub

End Module
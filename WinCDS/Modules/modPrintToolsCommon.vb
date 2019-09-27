Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports VBRUN

Module modPrintToolsCommon
    Public OutputObject As Object
    Public OutputToPrinter As Boolean
    Public Const DYMO_PaperSize_30252 as integer = 121                           ' 6x1.5 labels
    Public Const DYMO_PaperSize_30323 as integer = 126                           ' 6x3 labels
    Public Const DYMO_PaperSize_30256 as integer = 129                           ' 2x4 shipping
    Public Const DYMO_PaperSize_30270 as integer = 186                           ' continuous tape

    Public Const DYMO_PaperSize_ContinuousWide as integer = DYMO_PaperSize_30270 ' We didn't know it's SKU for a while..
    'Note: Printer code will be move to reporting software.

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
    Public Function NumLineBreaks(ByVal vStr As String) as integer
        Dim tmpLines As Object, tmpStart As Object
        NumLineBreaks = ((Len(vStr) - 1) \ 46 + 1)

        tmpStart = 1
        Do
            tmpStart = NextPrintBreak(vStr, tmpStart)
            tmpLines = tmpLines + 1
        Loop Until tmpStart >= Len(vStr)

        If tmpLines > NumLineBreaks Then NumLineBreaks = tmpLines
    End Function
    Public Function SetDymoPrinter(Optional ByRef PaperType as integer = 0) As Boolean
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
        Dim tmpBreak as integer, Absolute As Boolean
        NextPrintBreak = Start + 46
        If InStr(1, Mid(vStr, Start, 46), "MasterCard") Then NextPrintBreak = Start + 54 : Absolute = True

        tmpBreak = InStr(Start + 1, vStr, "I Authorize")

        If tmpBreak > 1 And tmpBreak < NextPrintBreak Then NextPrintBreak = tmpBreak : Absolute = True

        tmpBreak = InStr(Start + 1, vStr, "X _________")
        If tmpBreak > 1 And tmpBreak < NextPrintBreak Then NextPrintBreak = tmpBreak : Absolute = True

        tmpBreak = InStr(Start + 1, vStr, "Approval=")
        If tmpBreak > 1 And tmpBreak < NextPrintBreak Then NextPrintBreak = tmpBreak : Absolute = True

        If Not Absolute Then
            Dim A as integer, B as integer
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
    Public Function FindDymoPrinter(Optional ByVal IgnorePrevious As Boolean = False, Optional ByVal SelectBox as integer = 1, Optional ByVal StoreNum as integer = 0, Optional ByVal Current As String = "") As String
        Dim LabelPrinter As String, IsDymo As Boolean, DymoOnly As Boolean
        Dim P As Object
        Dim DN As String, SelList() As Object
        Dim Sel as integer
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
        MsgBox("Couldn't find a DYMO printer.", vbInformation, "Couldn't Find")
        Exit Function
CantSave:
        MsgBox("We could not save your settings." & vbCrLf & "ERR (" & Err.Number & "): " & Err.Description, vbCritical, "Couldn't save")
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
    Public Sub PrintInBox(ByVal PrintOb As Object, ByVal PrintText As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, Optional ByVal FontSize as integer = -1, Optional ByVal HAlign As AlignConstants = AlignConstants.vbAlignNone, Optional ByVal VAlign As AlignConstants = AlignConstants.vbAlignNone, Optional ByVal BorderStyle As BorderStyleConstants = 0)
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
    Public Sub PrintToPosition(Optional ByVal OutOb As Object = Nothing, Optional ByVal OutText As String = "", Optional ByVal Position as integer = -1, Optional ByVal Alignment As AlignConstants = AlignConstants.vbAlignLeft, Optional ByVal NewLine As Boolean = False)
        If OutOb Is Nothing Then OutOb = OutputObject
        If IsNothing(OutOb) Then Exit Sub
        If Position = -1 Then Position = OutOb.CurrentX

        Dim TruePos as integer
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
    Public Function IscPrinter(ByVal Ob As Object) As Boolean
        IscPrinter = TypeName(Ob) = "cPrinter"
    End Function

    Public Function BestFontFit(ByVal OutOb As Object, ByRef OutText As String, ByRef MaxFontSize as integer, ByRef MaxWidth as integer, ByRef MaxHeight as integer) As Double
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

End Module

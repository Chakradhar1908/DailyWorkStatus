Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class RichTextBoxNew
    Public Event key_press(ByRef KeyAscii As Integer)
    Public Event Change()
    Public Event Click_B()
    'Public Event DblClick()
    Public Event Key_Down(KeyCode As Integer, Shift As Integer)
    Public Event Key_Up(KeyCode As Integer, Shift As Integer)
    Public Event SelChange()
    Public Event Mouse_Down(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event Mouse_Move(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event Mouse_Up(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Public File As String
    'Private Rtb As RichTextBox = mRichTextBox
    Dim Printer As New Printer

    'Private Sub mRichTextBox_Change()
    '    RaiseEvent Change()
    'End Sub
    Private Sub mRichTextBox_Click()
        RaiseEvent Click_B()
    End Sub
    'Private Sub mRichTextBox_DblClick(): RaiseEvent DblClick: End Sub
    Private Sub mRichTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
        RaiseEvent Key_Down(KeyCode, Shift)
    End Sub
    Private Sub mRichTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
        RaiseEvent Key_Up(KeyCode, Shift)
    End Sub
    Private Sub mRichTextBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent Mouse_Down(Button, Shift, X, Y)
    End Sub
    Private Sub mRichTextBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent Mouse_Move(Button, Shift, X, Y)
    End Sub
    Private Sub mRichTextBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent Mouse_Up(Button, Shift, X, Y)
    End Sub
    Private Sub mRichTextBox_SelChange()
        RaiseEvent SelChange()
    End Sub

    Private Sub RichTextBoxNew_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        'mRichTextBox.Locked = False
        With mRichTextBox
            .Left = 0
            .Top = 0
            .Width = Width
            .Height = Height
        End With

    End Sub

    Public ReadOnly Property RichTextBox() As RichTextBox
        Get
            'Return Rtb
            RichTextBox = mRichTextBox
        End Get
    End Property

    Public Sub FileEdit()
        'for customer terms
        On Error GoTo CantOpen
        RunWordpad(File, , True)
        FileRead()
        Exit Sub
CantOpen:
        MessageBox.Show("Can't open Wordpad to edit the file.  We apologize for the inconvenience.", "Error opening MS WordPad", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub SetBarcodeLarge(ByVal T As String) ', Optional ByVal FontName As String = FONT_C39_SLIM)
        Dim fontsize As Font = mRichTextBox.SelectionFont

        SelectBarcodeFont(BarCodeFonts.bcfHalfInch)
        'Printer.FontName = FONT_C39_HALFINCH
        'mRichTextBox.Font = Printer.FontName
        mRichTextBox.Font = New Font(Printer.FontName, mRichTextBox.Font.Size)
        mRichTextBox.SelectionStart = 0
        mRichTextBox.SelectionLength = Len(PrepareBarcode(Trim(T)))
        mRichTextBox.SelectionFont = New Font(fontsize.FontFamily, 20)
        mRichTextBox.Text = PrepareBarcode(Trim(T))
    End Sub

    Public Sub SetBarcodeMed(ByVal T As String) ', Optional ByVal FontName As String = FONT_C39_WIDE)
        Dim fontsize As Font = mRichTextBox.SelectionFont

        'mRichTextBox.Font = FontName
        SelectBarcodeFont(BarCodeFonts.bcfWide)
        'mRichTextBox.Font = Printer.FontName
        mRichTextBox.Font = New Font(Printer.FontName, mRichTextBox.Font.Size)
        mRichTextBox.SelectionStart = 0
        mRichTextBox.SelectionLength = Len(PrepareBarcode(Trim(T)))
        mRichTextBox.SelectionFont = New Font(fontsize.FontFamily, 14)
        mRichTextBox.Text = PrepareBarcode(Trim(T))
    End Sub

    Public Sub SetBarcodeSmall(ByVal T As String) ', Optional ByVal FontName As String = FONT_C39_WIDE)
        Dim fontsize As Font = mRichTextBox.SelectionFont

        'mRichTextBox.Font = FontName
        SelectBarcodeFont(BarCodeFonts.bcfWide)
        'mRichTextBox.Font = Printer.FontName
        mRichTextBox.Font = New Font(Printer.FontName, mRichTextBox.Font.Size)
        mRichTextBox.SelectionStart = 0
        mRichTextBox.SelectionLength = Len(PrepareBarcode(Trim(T)))
        mRichTextBox.SelectionFont = New Font(fontsize.FontFamily, 14)
        mRichTextBox.Text = PrepareBarcode(Trim(T))
    End Sub

    Public Sub SetBarcodeSmallMediumRegular(ByVal T As String) ', Optional ByVal FontName As String = FONT_C39_SMALL_MEDIUM)
        Dim fontsize As Font = mRichTextBox.SelectionFont

        'Public Sub SetBarcodeSlimRegular(t As String, Optional FontName As String = FONT_C39_SLIM)
        'Public Sub SetBarcodeSlimRegular(t As String, Optional FontName As String = FONT_C39_SLIM)
        '  mRichTextBox.Font = FontName
        SelectBarcodeFont(BarCodeFonts.bcfSmallMedium)
        'mRichTextBox.Font = Printer.FontName
        mRichTextBox.Font = New Font(Printer.FontName, mRichTextBox.Font.Size)
        mRichTextBox.SelectionStart = 0
        mRichTextBox.SelectionLength = Len(PrepareBarcode(Trim(T)))
        mRichTextBox.SelectionFont = New Font(fontsize.FontFamily, 14)
        mRichTextBox.Text = PrepareBarcode(Trim(T))
    End Sub

    Public Sub FileRead(Optional ByVal CreateFileOnError As Boolean = False, Optional ByVal FileToOpen As String = "")
        On Error GoTo HandleErr

        If FileToOpen <> "" Then File = FileToOpen
        'mRichTextBox.LoadFile(File, RichTextBox.Rtf)
        mRichTextBox.LoadFile(File)
        Exit Sub

HandleErr:
        If Err.Number = 75 Then
            If CreateFileOnError And CreateBlankFile(File) Then
                FileRead()  ' Try again, without the error bypass.
            Else
                HideSplash()
                If File.Substring(0, 2) = LocalDrive Then
                    MsgBox("Could not load local file: " & File & ".", vbCritical, "Error")
                Else
                    MsgBox("Failed to load critical file: " & File & vbCrLf2 & "You appear to not be connected to the network!" & vbCrLf2 & "Use Windows Explorer to map Drive I: to the server, and try again!", vbCritical, "NetworkError")
                End If
                End
            End If
        End If
    End Sub

    Private Function CreateBlankFile(ByVal fName As String) As Boolean
        On Error GoTo NoGood
        '        Dim FNum as integer
        '        FNum = FreeFile()
        '        Open fName For Output As #FNum
        '        Print #FNum, "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}{\f1\fnil\fcharset0 Courier New;}}{\*\generator Msftedit 5.41.15.1503;}\viewkind4\uc1\pard\b\f0\fs20 \b0\f1\par}"
        '        Close #FNum
        '  CreateBlankFile = True
        '        Exit Function
        'NoGood:

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(fName, False)
        file.WriteLine("{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}{\f1\fnil\fcharset0 Courier New;}}{\*\generator Msftedit 5.41.15.1503;}\viewkind4\uc1\pard\b\f0\fs20 \b0\f1\par}")
        file.Close()
        CreateBlankFile = True
        Exit Function
NoGood:
    End Function

    Public Sub FilePrint(Optional LeftMarginWidth as integer = -1, Optional TopMarginHeight as integer = -1, Optional PrintWidth as integer = -1, Optional PrintHeight as integer = -1, Optional DontEndDoc As Boolean = False, Optional AllowMultiplePages As Boolean = True)
        PrintRTF(mRichTextBox, LeftMarginWidth, TopMarginHeight, PrintWidth, PrintHeight, , AllowMultiplePages)

        ' This really should move out of the common print area.
        ' Someday we might have two RTF boxes to print, or want to print something after the RTF.
        If Not DontEndDoc And SelectPrinter.TagSize <> "SMALL" Then Printer.EndDoc()         ' Allow the RTF to free up memory
    End Sub

    Private Sub mRichTextBox_DblClick()
        'for customer terms, etc
        FileEdit()
    End Sub

    Private Sub mRichTextBox_KeyPress(KeyAscii As Integer)
        RaiseEvent key_press(KeyAscii)
        KeyAscii = 0
    End Sub

    Public Sub SetInsuranceContract(ByVal T As String, Optional ByVal FontName As String = "Arial")
        Dim fontsize As Font = mRichTextBox.SelectionFont
        mRichTextBox.SelectionFont = New Font(fontsize.FontFamily, FontName)
        mRichTextBox.SelectionStart = 0
    End Sub

    Public Sub DoPrintFile(ByVal FileName As String, Optional ByVal LeftMarginWidth as integer = -1, Optional ByVal TopMarginHeight as integer = -1, Optional ByVal PrintWidth as integer = -1, Optional ByVal PrintHeight as integer = -1, Optional ByVal DontEndDoc As Boolean = False, Optional ByVal AllowMultiplePages As Boolean = True)

        Dim RTF As String, OF1 As String, OE As Boolean

        With RichTextBox
            'RTF = .TextRTF
            RTF = .Rtf
            OE = .Enabled
            OF1 = File


            File = FileName
            FileRead(False)
            .Enabled = False
            FilePrint(LeftMarginWidth, TopMarginHeight, PrintWidth, PrintHeight, DontEndDoc, AllowMultiplePages)

            '.TextRTF = RTF
            .Rtf = RTF
            .Enabled = OE
            File = OF1
        End With
    End Sub

    Private Sub mRichTextBox_TextChanged(sender As Object, e As EventArgs) Handles mRichTextBox.TextChanged
        RaiseEvent Change()
    End Sub

    Public Function asHtml() As String
        asHtml = RtfToHtml(mRichTextBox)
    End Function

End Class

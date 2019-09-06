Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module modBarcode
    Public Printer As New Printer
    Public BarcodeFormType As Byte
    Public Barcode As String

    Public Enum BarCodeFonts
        bcfWide = 1
        bcfOneInch = 2
        bcfSlim = 3
        bcfHalfInch = 4
        bcfQuarterInch = 5
        bcfSmallHigh = 6
        bcfSmallMedium = 7
        bcfSmallLow = 8

        bcf128Regular = 9

        bcfFirst = 1
        bcfLast = 9
        bcfDefault = 2

        bcfNone = 0
    End Enum

    Public Enum BarCodeTypes
        bctNA = 0
        bctCode39
        bctCode128
    End Enum

    Public Enum CommPorts
        COM1 = 0
        COM2 = 1
        COM3 = 2
        COM4 = 3
        COM5 = 4
        COM6 = 5
        COM7 = 6
        COM8 = 7
        COM9 = 8
        COM10 = 9
        COM11 = 10
        COM12 = 11
        COM13 = 12
        COM14 = 13
        COM15 = 14
        COM16 = 15
    End Enum

    Declare Function csp2Restore Lib "csp2.dll" () As Integer
    Declare Function csp2ReadData Lib "csp2.dll" () As Integer
    Declare Function csp2GetASCIIMode Lib "csp2.dll" () As Integer
    Public Const PARAM_ON As Integer = 1
    Declare Function csp2GetPacket Lib "csp2.dll" (stPacketData As Byte, ByVal lgBarcodeNumber As Integer, ByVal MaxLength As Integer) As Integer
    Declare Function csp2Init Lib "csp2.dll" (ByVal nComPort As Integer) As Integer
    Public Const STATUS_OK As Integer = 0
    Declare Function csp2WakeUp Lib "csp2.dll" () As Integer
    Public Const COMMUNICATIONS_ERROR As Integer = -1
    Public Const BAD_PARAM As Integer = -2
    Public Const SETUP_ERROR As Integer = -3
    Public Const INVALID_COMMAND_NUMBER As Integer = -4
    Public Const COMMAND_LRC_ERROR As Integer = -7
    Public Const RECEIVED_CHARACTER_ERROR As Integer = -8
    Public Const GENERAL_ERROR As Integer = -9
    Public Const FILE_NOT_FOUND As Integer = 2
    Public Const ACCESS_DENIED As Integer = 5
    Declare Function csp2ClearData Lib "csp2.dll" () As Integer
    Public BarcodeListQty As Integer
    Public BarcodeList() As String

    Public Function SelectBarcodeFont(Optional ByVal Alt As BarCodeFonts = BarCodeFonts.bcfDefault, Optional ByVal FSize As Integer = 39, Optional ByRef OldFontName As String = "", Optional ByRef OldFontSize As Integer = 1) As Boolean
        ' for some reason, setting the fontname to barcodes was failing randomly..
        ' so we'll try it a few times to make sure
        Dim T As String, X As Integer

        SelectBarcodeFont = False
        OldFontName = Printer.FontName
        OldFontSize = Printer.FontSize

        If UseCode128() Then Alt = BarCodeFonts.bcf128Regular
        T = BarCodeFontName(Alt)
        Printer.FontSize = FSize

        Do While Printer.FontName <> T
            Printer.FontName = T
            X = X + 1
            If X > 30 Then Exit Do
        Loop

        If Printer.FontName <> T Then
            MsgBox("Barcode " & T & " not installed.", vbExclamation, "No Barcodes")
            Exit Function
        End If
        SelectBarcodeFont = True
    End Function

    Public Function PrepareBarcode(ByVal Barcode As String) As String
        Select Case BarCodeFontNameType()
            Case BarCodeTypes.bctCode128
                PrepareBarcode = Code128(Barcode)
            Case Else ' both None and Code39
                PrepareBarcode = Barcode
                PrepareBarcode = Replace(PrepareBarcode, " ", "_")
                If Len(PrepareBarcode) > 0 Then PrepareBarcode = "*" & PrepareBarcode & "*"
        End Select
    End Function

    Public Function UseCode128() As Boolean
        '  UseCode128 = IsCDSComputer("LAPTOP")
        UseCode128 = False
    End Function

    Public Function BarCodeFontName(Optional ByVal Alt As BarCodeFonts = BarCodeFonts.bcfDefault) As String
        Select Case Alt
            Case BarCodeFonts.bcfWide : BarCodeFontName = FONT_C39_WIDE
            Case BarCodeFonts.bcfOneInch : BarCodeFontName = FONT_C39_ONEINCH
            Case BarCodeFonts.bcfSlim : BarCodeFontName = FONT_C39_SLIM
            Case BarCodeFonts.bcfHalfInch : BarCodeFontName = FONT_C39_HALFINCH
            Case BarCodeFonts.bcfQuarterInch : BarCodeFontName = FONT_C39_QUARTERINCH
            Case BarCodeFonts.bcfSmallHigh : BarCodeFontName = FONT_C39_SMALL_HIGH
            Case BarCodeFonts.bcfSmallMedium : BarCodeFontName = FONT_C39_SMALL_MEDIUM
            Case BarCodeFonts.bcfSmallLow : BarCodeFontName = FONT_C39_SMALL_LOW

            Case BarCodeFonts.bcf128Regular : BarCodeFontName = FONT_C128_REGULAR

            Case Else : BarCodeFontName = FONT_C39_ONEINCH
        End Select
    End Function

    Public Function BarCodeFontNameType(Optional ByVal Alt As String = "") As BarCodeTypes
        BarCodeFontNameType = BarCodeFontType(BarCodeFontEnum(Alt))
    End Function

    Public Function BarCodeFontType(Optional ByVal Alt As BarCodeFonts = BarCodeFonts.bcfDefault) As BarCodeTypes
        Select Case Alt
            Case BarCodeFonts.bcfNone : BarCodeFontType = BarCodeTypes.bctNA
            Case BarCodeFonts.bcf128Regular : BarCodeFontType = BarCodeTypes.bctCode128
            Case Else : BarCodeFontType = BarCodeTypes.bctCode39 '### shouldn't be a case else
        End Select
    End Function

    Public Function BarCodeFontEnum(Optional ByVal FNT As String = "") As BarCodeFonts
        FNT = BarcodeGetCurrentFont(FNT)
        Select Case FNT
            Case FONT_C39_WIDE : BarCodeFontEnum = BarCodeFonts.bcfWide
            Case FONT_C39_ONEINCH : BarCodeFontEnum = BarCodeFonts.bcfOneInch
            Case FONT_C39_SLIM : BarCodeFontEnum = BarCodeFonts.bcfSlim
            Case FONT_C39_HALFINCH : BarCodeFontEnum = BarCodeFonts.bcfHalfInch
            Case FONT_C39_QUARTERINCH : BarCodeFontEnum = BarCodeFonts.bcfQuarterInch
            Case FONT_C39_SMALL_HIGH : BarCodeFontEnum = BarCodeFonts.bcfSmallHigh
            Case FONT_C39_SMALL_MEDIUM : BarCodeFontEnum = BarCodeFonts.bcfSmallMedium
            Case FONT_C39_SMALL_LOW : BarCodeFontEnum = BarCodeFonts.bcfSmallLow
            Case FONT_C128_REGULAR : BarCodeFontEnum = BarCodeFonts.bcf128Regular
            Case Else : BarCodeFontEnum = BarCodeFonts.bcfNone
        End Select
    End Function

    Private Function BarcodeGetCurrentFont(Optional ByVal Specified As String = "") As String
        On Error Resume Next
        If Specified <> "" Then BarcodeGetCurrentFont = Specified : Exit Function
        If Not (OutputObject Is Nothing) Then BarcodeGetCurrentFont = OutputObject.FontName : Exit Function
        BarcodeGetCurrentFont = Printer.FontName
    End Function

    Public Function Code128(ByRef ChainE As String) As String
        'V 2.0.0
        'Parametres : une chaine
        'Parameters : a string
        'Retour : * une chaine qui, affichee avec la police CODE128.TTF, donne le code barre
        '         * une chaine vide si parametre fourni incorrect
        'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
        '         * an empty string if the supplied parameter is no good
        Dim I As Integer, Checksum As Integer, Mini As Integer, dummy As Integer, tableB As Boolean
        Code128 = ""
        If Len(ChainE) > 0 Then
            'Verifier si caracteres valides
            'Check for valid characters
            For I = 1 To Len(ChainE)
                Select Case Asc(Mid(ChainE, I, 1))
                    Case 32 To 126, 203
                    Case Else
                        I = 0
                        Exit For
                End Select
            Next
            'Calculer la chaine de code en optimisant l'usage des tables B et C
            'Calculation of the code string with optimized use of tables B and C
            Code128 = ""
            tableB = True
            If I > 0 Then
                I = 1 'i% devient l'index sur la chaine / i% become the string index
                Do While I <= Len(ChainE)
                    If tableB Then
                        'Voir si interessant de passer en table C / See if interesting to switch to table C
                        'Oui pour 4 chiffres au debut ou a la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
                        Mini = IIf(I = 1 Or I + 3 = Len(ChainE), 4, 6)

                        Mini = Mini - 1
                        If I + Mini <= Len(ChainE) Then
                            Do While Mini >= 0
                                If Asc(Mid(ChainE, I + Mini, 1)) < 48 Or Asc(Mid(ChainE, I + Mini, 1)) > 57 Then Exit Do
                                Mini = Mini - 1
                            Loop
                        End If

                        If Mini < 0 Then 'Choix table C / Choice of table C
                            If I = 1 Then 'Debuter sur table C / Starting with table C
                                Code128 = Chr(210)
                            Else 'Commuter sur table C / Switch to table C
                                Code128 = Code128 & Chr(204)
                            End If
                            tableB = False
                        Else
                            If I = 1 Then Code128 = Chr(209) 'Debuter sur table B / Starting with table B
                        End If
                    End If
                    If Not tableB Then
                        'On est sur la table C, essayer de traiter 2 chiffres / We are on table C, try to process 2 digits
                        Mini = 2

                        Mini = Mini - 1
                        If I + Mini <= Len(ChainE) Then
                            Do While Mini >= 0
                                If Asc(Mid(ChainE, I + Mini, 1)) < 48 Or Asc(Mid(ChainE, I + Mini, 1)) > 57 Then Exit Do
                                Mini = Mini - 1
                            Loop
                        End If

                        If Mini < 0 Then 'OK pour 2 chiffres, les traiter / OK for 2 digits, process it
                            dummy = Val(Mid(ChainE, I, 2))
                            dummy = IIf(dummy < 95, dummy + 32, dummy + 105)
                            Code128 = Code128 & Chr(dummy)
                            I = I + 2
                        Else 'On n'a pas 2 chiffres, repasser en table B / We haven't 2 digits, switch to table B
                            Code128 = Code128 & Chr(205)
                            tableB = True
                        End If
                    End If
                    If tableB Then
                        'Traiter 1 caractere en table B / Process 1 digit with table B
                        Code128 = Code128 & Mid(ChainE, I, 1)
                        I = I + 1
                    End If
                Loop
                'Calcul de la cle de controle / Calculation of the checksum
                For I = 1 To Len(Code128)
                    dummy = Asc(Mid(Code128, I, 1))
                    dummy = IIf(dummy < 127, dummy - 32, dummy - 105)
                    If I = 1 Then Checksum = dummy
                    Checksum = (Checksum + (I - 1) * dummy) Mod 103
                Next
                'Calcul du code ASCII de la cle / Calculation of the checksum ASCII code
                Checksum = IIf(Checksum < 95, Checksum + 32, Checksum + 105)
                'Ajout de la cle et du STOP / Add the checksum and the STOP
                Code128 = Code128 & Chr(Checksum) & Chr(211)
            End If
        End If
        Exit Function
    End Function

    Public Function GetNextBarcode(ByRef CallingForm As Object) As String
        BarcodeFormType = 0
        'frmBarcode.Show vbModal, CallingForm
        frmBarcode.ShowDialog(CallingForm)
        GetNextBarcode = Barcode
        Barcode = ""
    End Function

    Public Function ValidBarcode(ByVal Barcode As String) As Boolean
        ' Valid characters are: A-Z, 0-9, Space, $ % + - . /
        Dim I As Integer
        If Barcode = "" Then Exit Function
        If Len(Barcode) > Setup_2Data_StyleMaxLen Then Exit Function
        For I = 1 To Len(Barcode)
            Select Case Mid(Barcode, I, 1)
                Case "A" To "Z"
                Case "0" To "9"
                Case " ", "$", "%", "+", "-", ".", "/"
                Case Else
                    ValidBarcode = False
                    Exit Function
            End Select
        Next
        ValidBarcode = True
    End Function

    Public Function InterpretBarcode(ByVal Barcode As String) As String
        ' Unmap barcodes.
        InterpretBarcode = Barcode
        InterpretBarcode = Replace(InterpretBarcode, "_", " ")
        Exit Function
    End Function

End Module

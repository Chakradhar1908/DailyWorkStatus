Imports VBRUN
Module modDataValidation
    Public Function GetPrice(ByVal Value As String, Optional ByRef PriceError As Boolean = False) As Decimal
        ' This function is used when a string may contain a numerical value.
        ' Calling this will allow for handing of "" or "$1.24" type cases,
        ' without raising a type error.
        PriceError = False
        If IsNothing(Value) Then Value = 0

        On Error GoTo AnError
        If Len(Value) = 0 Then
            GetPrice = 0
        Else
            GetPrice = Value
        End If
        GetPrice = Math.Round(GetPrice, 2)

        Exit Function
AnError:
        GetPrice = 0
        PriceError = True
    End Function
    ':FUNCTION cleanani(ani, req)
    ':removes all non number charecters
    ' Provided by Krollmark Technologies 20030708
    Public Function CleanAni(ByVal Ani As String, Optional ByVal Req As Integer = 0)
        Dim I As Integer
        CleanAni = ""
        If Ani <> "" Then
            For I = 1 To Len(Ani)
                If IsNumeric(Mid(Ani, I, 1)) Then CleanAni = CleanAni & Mid(Ani, I, 1)
            Next
            If Len(CleanAni) < Req Then CleanAni = ""
        End If
    End Function
    Public Function FormatAniTextBox(ByRef tBox As TextBox) As Integer
        Dim OrigAni As String, OrigPos As Integer, OrigLen As Integer, OrigSel As Integer
        Dim tempAni As String, tempPos As Integer, tempLen As Integer, tempSel As Integer

        FormatAniTextBox = 0
        OrigAni = tBox.Text
        OrigLen = Len(CleanAni(OrigAni))
        OrigPos = tBox.SelectionStart
        OrigSel = tBox.SelectionLength

        ' For textbox purposes, we'll assume the user intends to input a 7/10 digit number.
        tempAni = CleanAni(OrigAni)
        If Len(tempAni) > 7 Then
            tempAni = "(" & Left(tempAni, 3) & ") " & Mid(tempAni, 4, 3) & "-" & Mid(tempAni, 7)
        ElseIf Len(tempAni) = 7 Then
            tempAni = Left(tempAni, 3) & "-" & Mid(tempAni, 4)
        ElseIf Len(tempAni) > 3 Then
            tempAni = "(" & Left(tempAni, 3) & ") " & Mid(tempAni, 4)
        End If
        tempLen = Len(tempAni)

        tBox.Text = tempAni
        ' Selected position starts at: OrigPos + (# characters added before OrigPos)
        ' How to figure characters added before OrigPos?  Handle type, paste, delete.

        ' 2481x     -> (248) 1x  => 5-8
        ' (248) 1x  -> 248       => 8-5
        ' (248) 12x -> (248) 12x => 9-9
        ' 2x481     -> (2x48) 1  => 2-3
        ' (2x48) 1  -> (2x48) 1  => 3-3
        ' 2x48) 1   -> (2x48) 1  => 2-3 ' Tried to delete formatting text..

        ' 2248-6732 -> (224) 867-32

        ' If len doesn't change, neither does sel.
        ' If len expands..
        '   What expanded before the cursor?
        '  If OrigLen = tempLen Then
        '    tempPos = OrigPos
        '    tempSel = OrigSel
        '  Else
        ' Let's revise this.. it's too hard to crack.
        ' Count the number of numbers before the current cursor.
        ' Put the new cursor after that many numbers in the formatted text.
        Dim I As Integer, NumCount As Integer
        For I = 1 To OrigPos
            If IsNumeric(Mid(OrigAni, I, 1)) Then
                NumCount = NumCount + 1
            End If
        Next
        For I = 1 To tempLen
            If IsNumeric(Mid(tempAni, I, 1)) Then
                NumCount = NumCount - 1
                If NumCount = 0 Then
                    tempPos = I
                End If
            End If
        Next

        '    tempPos = OrigPos + (tempLen - OrigLen)  ' Only add lengths from before selection!
        '    ' Can't assume formatting was in place..
        '    ' OrigLen<4 and OrigPos<4 and TempLen>5 -> tempPos=OrigPos+1?
        '    If OrigLen = 4 And OrigPos < 4 Then
        '      tempPos = OrigPos + 1
        '    End If
        '    If OrigLen >= 6 And tempLen <= 3 And OrigPos < 6 Then
        '      tempPos = OrigPos - 1
        '    End If
        '    If OrigLen = 10 And tempLen = 11 And OrigPos < 10 Then
        '      tempPos = OrigPos
        '    End If
        '    If OrigLen = 10 And tempLen = 9 And OrigPos < 10 Then
        '      tempPos = OrigPos
        '    End If
        tempSel = OrigSel
        '  End If
        If tempPos < 0 Then tempPos = 0
        tBox.SelectionStart = tempPos
        tBox.SelectionLength = tempSel
    End Function
    ':FUNCTION dressani(ani)
    ':formats phone number
    ' Provided by Krollmark Technologies 20030708
    Public Function DressAni(ByVal Ani As String)
        Dim tempAni As String
        If Not IsNumeric(Ani) Then DressAni = Ani : Exit Function
        DressAni = ""
        tempAni = Ani
        If Left(Ani, 1) = 1 Then tempAni = Mid(Ani, 2)
        If Len(tempAni) = 7 Then
            DressAni = Left(tempAni, 3) & "-" & Mid(tempAni, 4, 4)
        ElseIf Len(tempAni) >= 10 Then
            DressAni = "(" & Left(tempAni, 3) & ") " & Mid(tempAni, 4, 3) & "-" & Mid(tempAni, 7, 4)
            If Len(tempAni) > 10 Then DressAni = DressAni & " " & Mid(tempAni, 11)
            If Right(DressAni, 1) = "-" Then DressAni = Left(DressAni, Len(DressAni) - 1)
        Else
            DressAni = tempAni
        End If
    End Function
    Public Function CurrencyFormat(ByVal curMoney As Object, Optional ByVal Strip00 As Boolean = False, Optional ByVal DollarSign As Boolean = False, Optional ByVal NoCommas As Boolean = False) As String
        '::::CurrencyFormat
        ':::SUMMARY
        ':Used to format the currency Amount.
        ':::DESCRIPTION
        ':Format the currency as a string and we can also insert Dollar sign.
        ':::PARAMETERS
        ':-curMoney
        ':-Strip00
        ':-DollarSign
        ':::RETURN
        ':String-Returns the CurrencyFormat string.
        If IsNothing(curMoney) Then curMoney = 0
        If Not IsNumeric(curMoney) Then curMoney = 0
        CurrencyFormat = Format(curMoney, CurrencyFormatString)
        If Strip00 = True And Right(CurrencyFormat, 3) = ".00" Then CurrencyFormat = Left(CurrencyFormat, Len(CurrencyFormat) - 3)
        If DollarSign = True Then CurrencyFormat = "$" & CurrencyFormat
        If NoCommas Then CurrencyFormat = Replace(CurrencyFormat, ",", "")
    End Function
    Public Function CurrencyFormatString() As String

        '::::CurrencyFormatString
        ':::SUMMARY
        'Used to format the CurrencyFormatString.
        ':::DESCRIPTION
        ': Used to format the CurrencyFormatString in given custom numeric format string.
        ':::PARAMETERS
        ':::RETURN
        ':String-Returns the CurrencyFormatString as a String.
        ':::SEE ALSO
        ':-CurreencyFormat
        CurrencyFormatString = "###,##0.00"
    End Function

    Public Function IsNothingOrZero(ByVal nVal) As Boolean
        ':::IsNothingOrZero
        ':::SUMMARY
        ':This function just checks whether the value is Nothing or Zero.
        ':::DESCRIPTION
        ':This function is used to ensure whether the required data like Street Address  etc is present in Logo.
        ':::PARAMETERS
        ':-nVal
        ':::RETURN
        ':Boolean-Indicates whether the value is true or false.
        IsNothingOrZero = True
        If IsNothing(nVal) Then Exit Function
        If nVal = 0 Then Exit Function
        IsNothingOrZero = False
    End Function
    Public Function FormatGM(ByVal GM As Double, Optional ByVal DecimalPoints As Integer = 2) As String
        '::::FormatGM
        ':::SUMMARY
        ':FormatGM function is used to generate the format for GrossMargin.
        ':::DESCRIPTION
        ':Result from function GMFormatis assigned to FormatGM which is used to generate the format of GrossMargin.
        ':::PARAMETERS
        ':-GM-Denotes the current GrossMargin .
        ':-DecimalPoints-Based on number of decimal points, function GMFormat gets the result which is assigned to function FormatGM.
        ':::RETURN
        ':String-Returns the GrossMargin as a string.
        FormatGM = GMFormat(GM, DecimalPoints)
    End Function

    Public Function GMFormat(ByVal GM As Double, Optional ByVal DecimalPoints As Integer = 2) As String
        '::::GMFormat
        ':::SUMMARY
        ':Result from function Format is assigned to function GMFormat.
        ':::DESCRIPTION
        ':Here format of GrossMargin is designed by taking number of decimal points as a criteria.
        ':::PARAMETERS
        ':-GM-Denotes the current Gross Margin.
        ':-DecimalPoints-Denotes the decimal points after the GrossMargin value.
        ':::RETURN
        ':-String-Returns the GrossMargin as a string.

        Dim S As String, I As Integer
        If DecimalPoints < 0 Then DecimalPoints = 0
        S = "0"
        If DecimalPoints > 0 Then
            S = S & "."
            For I = 1 To DecimalPoints : S = S & "0" : Next
        End If
        GMFormat = Format(GM, S)
    End Function

    Public Function CurrencyFormat(ByVal curMoney As Decimal, Optional ByVal Strip00 As Boolean = False, Optional ByVal DollarSign As Boolean = False, Optional ByVal NoCommas As Boolean = False) As String
        '::::CurrencyFormat
        ':::SUMMARY
        ':Used to format the currency Amount.
        ':::DESCRIPTION
        ':Format the currency as a string and we can also insert Dollar sign.
        ':::PARAMETERS
        ':-curMoney
        ':-Strip00
        ':-DollarSign
        ':::RETURN
        ':String-Returns the CurrencyFormat string.
        If IsNothing(curMoney) Then curMoney = 0
        If Not IsNumeric(curMoney) Then curMoney = 0
        CurrencyFormat = Format(curMoney, CurrencyFormatString)
        If Strip00 = True And Right(CurrencyFormat, 3) = ".00" Then CurrencyFormat = Left(CurrencyFormat, Len(CurrencyFormat) - 3)
        If DollarSign = True Then CurrencyFormat = "$" & CurrencyFormat
        If NoCommas Then CurrencyFormat = Replace(CurrencyFormat, ",", "")
    End Function

    Public Function IsNotNothing(ByRef objReference) As Boolean
        IsNotNothing = Not IsNothing(objReference)
    End Function

    Public Function ZeroToEmptyString(ByVal Value As Object) As Object
        '::::ZeroToEmptyString
        ':::SUMMARY
        ':This function is used to display value from zero to empty string, i.e any value.
        ':::DESCRIPTION
        ':This function is used in Inventory Maintenance form ,used to handle any type errors.
        ':::PARAMETERS
        ':-Value-Indicates the input value given by user.
        ':::RETURN
        ':Variant
        On Error Resume Next
        If Val(Value) = 0 Then
            ZeroToEmptyString = ""
        Else
            ZeroToEmptyString = Value
        End If
    End Function

    Public Function QuantityFormat(ByVal Q As Double, Optional ByVal Decimals As Integer = 2, Optional ByVal BlankEmpty As Boolean = False) As String
        '::::QuantityFormat
        ':::SUMMARY
        ':Used to display the formatted Quantity when the order is loading.
        ':::DESCRIPTION
        ':Gets the result from FormatQuantity.This function is  called, after filling information related to quantity,i.e when order is loading while making new sale.
        ':::PARAMETERS
        ':-Q-Denotes the current Quantity.
        ':-Decimals-Decimal points after the Quantity value.
        ':-BlankEmpty-Boolean function which denotes Quantity Blank is not Empty.
        ':::RETURN
        ':String-Returns the QuantityFormat string.
        QuantityFormat = FormatQuantity(Q, Decimals, BlankEmpty)
    End Function

    Public Function FormatQuantity(ByVal Q As Double, Optional ByVal Decimals As Integer = 2, Optional ByVal BlankEmpty As Boolean = True) As String
        '::::FormQuantity
        ':::SUMMARY
        ':Used to format the Quantity.
        ':::DESCRIPTION
        ':Here format of Quantity is designed by using function GMFormat.
        ':::PARAMETERS
        ':-Q-Denotes the current Quantity.
        ':-Decimals-Decimal points after the Quantity value.
        ':-BlankEmpty-Boolean function which denotes Quantity Blank is Empty.
        ':::RETURN
        ':String-Returns the FormatQuantity string.
        ':::SEE ALSO
        ':-QuantityFormat


        If BlankEmpty And Q = 0 Then Exit Function
        FormatQuantity = GMFormat(Q, Decimals)
    End Function

    Public Function GetDouble(ByVal Value As String) As Double
        'If IsNull(Value) Then Value = 0
        If IsNothing(Value) Then Value = 0
        On Error GoTo AnError
        GetDouble = CDbl(Value)
        Exit Function
AnError:
    End Function

    Public Function PriceFormatFunc(ByVal Number As Object, Optional ByVal Style As String = "$###,##0.00", Optional ByVal Length As Integer = 12, Optional ByVal BlankForZero As Boolean = False) As String
        '::::PriceFormatFunc
        ':::Summmary
        ':PriceFormatFunc is used display the price in certain format($###,##0.00).
        ':::DESCRIPTION
        ':This Function is mainly used to denote the format for price and gets the result from AlignString function.
        ':AlignString Function aligns a string based on relevant criteria.Useful for forcing a fixed-width or left or right alignment.
        ':::PARAMETERS
        ':-Number-Denotes the Price.
        ':-Style-Denotes the format style of price.
        ':-Length-Denotes the length of price to be display.
        ':-BlankForZero-It is Boolean Function.Used to display blank space when the value is Zero.
        ':::RETURN
        ':String-Returns price in given format as a string.

        Dim NC As Decimal, Text As String
        NC = GetPrice(Number)
        If BlankForZero And NC = 0 Then PriceFormatFunc = "" : Exit Function
        Text = Format(NC, Style)
        PriceFormatFunc = AlignString(Text, Length, AlignConstants.vbAlignRight, False) ' Space(12 - Len(Text)) & Text

    End Function

    Public Function DressEmail(ByVal S As String) As String
        DressEmail = S
    End Function

    Public Function CleanEmail(ByVal S As String) As String
        '::::CleanEmail
        ':::SUMMARY
        ':This function is used to clear all the email addresses which are no longer used.
        ':::DESCRIPTION
        ':Clears all the email addresses which are no longer used and returns the result as a string.
        ':::PARAMETERS
        ':-S-This parameter is directly assigned to CleanEmail.
        ':::RETURN
        ':String-Returns the result as a string.
        CleanEmail = S
    End Function

    Public Function CleanAddress(ByVal Add As String, Optional ByVal UpperCase As Boolean = True, Optional ByVal RemovePunctuation As Boolean = True) As String
        '::::CleanAddress
        ':::SUMMARY
        ':This function  is used to Remove any extra spaces in Address.
        ':::DESCRIPTION
        ': This Function is used to remove punctuation marks, extra space in Address.
        ':::PARAMETERS
        ':-Add-Denotes the Current Address.
        ':-Uppercase-Checks whether the Address is in Capitals letters or not.
        ':-RemovePunctuation-Chceks whether the Punctuation marks are present in Address or not.
        ':::RETURN
        ':String-Returns the Result as a string.

        If UpperCase Then Add = UCase(Add)
        If RemovePunctuation Then
            Add = Replace(Add, ".", "")
            Add = Replace(Add, ";", "")
            Add = Replace(Add, ",", "")
            Add = Replace(Add, "/", "")
        End If
        Add = Trim(Add)
        Do While InStr(Add, "  ") > 0 : Add = Replace(Add, "  ", " ") : Loop
        CleanAddress = Add
    End Function
End Module

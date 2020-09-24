Imports Microsoft.VisualBasic.Compatibility.VB6
Module modService
    Public Function PartOrderToHTML(ByVal PartOrderNo As String, Optional ByVal StoreNo As Integer = 0, Optional ByRef GetEmailAddr As String = "", Optional ByRef GetEmailName As String = "", Optional ByVal NoPageSetup As Boolean = False, Optional ByRef Attach As String = "") As String

        '::::PartOrderToHTML
        ':::SUMMARY
        ': HTML Code for Parts Order form.
        ':::DESCRIPTION
        ': This function contains HTML code for designing of Page Setup,Store Information,Page Title (w/ date),Vendor Information,Claim Information,Page CleanUp in PartsOrder form.Useful to handle errors and to load required information in Service parts order form.
        ':::PARAMETERS
        ': - PartOrderNo - Indicates Parts Order number.
        ': - StoreNo - Indicates store number.
        ': - GetEmailAddr - Parameter to get email address.
        ': - GetEmailName - Parameter to get  email name.
        ': - NoPageSetup - Indicates whether there is PageSetup or not.
        ': - Attach -
        ':::RETURN
        ': String - Returns Parts Order form as a string.
        ':::SEE ALSO
        ': ChargeBackLetterHTML
        Dim S As String
        Dim Part As clsServicePartsOrder, ServiceOrder As clsServiceOrder

        Dim SI As StoreInfo
        Dim VendorName As String
        Dim vName As String, vAddress As String = "", vAddress2 As String = "", vAddress3 As String = ""
        Dim VZip As String = "", VPhone As String = "", VFax As String = "", vEMail As String = ""


        Dim PicID As Integer

        On Error Resume Next

        If StoreNo = 0 Then StoreNo = StoresSld

        Part = New clsServicePartsOrder
        With Part
            If Not .Load(PartOrderNo, "#ServicePartsOrderNo") Then
                DisposeDA(Part)
                Exit Function
            End If
            ServiceOrder = New clsServiceOrder
            If Not ServiceOrder.Load(Part.ServiceOrderNo, "#ServiceOrderNo") Then ServiceOrder = Nothing

            SI = StoreSettings(StoreNo)
            VendorName = .Vendor
            '    GetVendorName (VendorName), (VendorName), vAddress, vAddress2, vAddress3, VZip, VPhone, VFax, (VendorName), vEMail
            If UseQB() Then
                QBGetVendorName(VendorName, VendorName, vAddress, vAddress2, vAddress3, VZip, VPhone, VFax, , vEMail)
            Else
                GetVendorName(VendorName, VendorName, vAddress, vAddress2, vAddress3, VZip, VPhone, VFax, , vEMail)
            End If

            GetEmailAddr = vEMail
            GetEmailName = VendorName

            S = ""

            ' templates
            S = S & "" & vbCrLf
            '    S = S & "  <></>" & vbCrLf

            If Not NoPageSetup Then
                '------- Page Setup
                S = S & "<html>" & vbCrLf
                S = S & " <head>" & vbCrLf
                S = S & "  <title>Part Order #" & PartOrderNo & " - " & SI.Name & "</title>" & vbCrLf
                S = S & "  <meta name=""GENERATOR"" content=""" & SoftwareVersion(True, False, True) & """>" & vbCrLf
                S = S & " </head>" & vbCrLf
                S = S & " <body bgcolor=#FFFFFF text=#000000 vlink=#FF0000 link=#FFFF00 hspace=0 vspace=0>" & vbCrLf
            End If

            S = S & "  <table border=0 cellspacing=0 cellpadding=0 width='100%'>" & vbCrLf ' alignment table


            S = S & "   <tr><td colspan=3 align=center>" & vbCrLf
            '------- Page Title (w/ date)
            S = S & "<table border=0 width='100%'>" & vbCrLf
            S = S & "  <tr><td align=center><b><font size=+3>PARTS ORDER #" & PartOrderNo & "</font></b></td></tr>" & vbCrLf
            S = S & "  <tr><td align=right>" & Format(Today, "mm/dd/yyyy") & "</td></tr>" & vbCrLf
            S = S & "</table>" & vbCrLf

            S = S & "   </tr></td>" & vbCrLf    ' alignment table
            S = S & "   <tr><td colspan=3 align=left>" & vbCrLf

            '------- Store Information
            S = S & "      <table border=1 cellspacing=0 cellpadding=0  width=400><tr>" & vbCrLf
            S = S & "      <td>" & vbCrLf
            S = S & "        <table border=0>" & vbCrLf
            S = S & "          <tr><td><b>From:</b></td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & SI.Name & "</font></td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & SI.Address & "</td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & SI.City & "</td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & DressAni(SI.Phone) & "</td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & StoreSettings.Email & "</td></tr>" & vbCrLf
            S = S & "        </table>" & vbCrLf
            S = S & "      </td>" & vbCrLf
            S = S & "      </tr></table><br>" & vbCrLf

            S = S & "   </tr></td>" & vbCrLf
            S = S & "   <tr><td colspan=3 align=left>" & vbCrLf

            '------- Vendor Information
            S = S & "      <table border=1 cellspacing=0 cellpadding=0 width=400><tr>" & vbCrLf
            S = S & "      <td>" & vbCrLf
            S = S & "        <table border=0>" & vbCrLf
            S = S & "          <tr><td><b>To:</b></td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & VendorName & "</font></td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & vAddress & "</td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & vAddress2 & IIf(Len(vAddress3) > 0, "", " " & VZip) & "</td></tr>" & vbCrLf
            If Len(vAddress3) > 0 Then
                S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & vAddress3 & " " & VZip & "</td></tr>" & vbCrLf
            End If
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & DressAni(VPhone) & IIf(Len(VFax) > 0, "  FAX: " & DressAni(VFax), "") & " " & VZip & "</td></tr>" & vbCrLf
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & vEMail & "</td></tr>" & vbCrLf
            S = S & "        </table>" & vbCrLf
            S = S & "      </td>" & vbCrLf
            S = S & "      </tr></table><br>" & vbCrLf

            S = S & "   </tr></td>" & vbCrLf
            S = S & "   <tr><td colspan=3 align=center>" & vbCrLf

            '------- Claim Information
            S = S & "<table>" & vbCrLf
            S = S & " <tr><td>Parts Order Number:</td><td>" & .ServicePartsOrderNo & "</td></tr>" & vbCrLf
            If Not ServiceOrder Is Nothing Then
                S = S & " <tr><td>Service Order Number:</td><td>" & .ServiceOrderNo & "</td></tr>" & vbCrLf
            End If
            S = S & " <tr><td>Date of Claim:</td><td>" & Format(.DateOfClaim, "mmmm dd,yyyy") & "</td></tr>" & vbCrLf
            S = S & " <tr><td>Status:</td><td>" & .Status & "</td></tr>" & vbCrLf
            If .InvoiceNo <> "" Then
                S = S & " <tr bgcolor=lightblue><td>Vendor Invoice No:</td><td>" & .InvoiceNo & "</td></tr>" & vbCrLf
            End If
            If .InvoiceDate <> "" Then
                S = S & " <tr bgcolor=lightblue><td>Vendor Invoice Date:</td><td>" & .InvoiceDate & "</td></tr>" & vbCrLf
            End If
            If Not ServiceOrder Is Nothing Then
                If ServiceOrder.SaleNo <> "" Then
                    S = S & " <tr><td>Sale Number:</td><td>" & ServiceOrder.SaleNo & "</td></tr>" & vbCrLf
                End If
            End If
            S = S & " <tr bgcolor=lightyellow><td>Repair Cost:</td><td>" & FormatCurrency(.ChargeBackAmount) & "</td></tr>" & vbCrLf
            S = S & " <tr bgcolor=lightyellow><td>Paid:</td><td>" & YesNo(.Paid) & "</td></tr>" & vbCrLf
            S = S & " <tr bgcolor=lightyellow><td>Reimbursement:</td><td>" & DescribeChargeBackOption(.ChargeBackType) & "</td></tr>" & vbCrLf
            S = S & " <tr><td colspan=2>&nbsp;</td></tr>" & vbCrLf
            S = S & " <tr><td><b><i>Style No</i></b></td><td><b><i>Description</i></b></td></tr>" & vbCrLf
            S = S & " <tr><td>" & .Style & "</td><td>" & .Desc & "</td></tr>" & vbCrLf
            S = S & " <tr><td colspan=2>&nbsp;</td></tr>" & vbCrLf
            S = S & " <tr><td colspan=2><b><i>Notes</i></b></td></tr>" & vbCrLf
            S = S & " <tr><td colspan=2>" & .Notes & "</td></tr>" & vbCrLf
            S = S & "</table>" & vbCrLf



            S = S & "  </td></tr></table>" & vbCrLf  ' end of alignment table

            If Not NoPageSetup Then
                '------- Page CleanUp
                S = S & " </body>" & vbCrLf
                S = S & "</html>" & vbCrLf
            End If

            PicID = FindDatabasePictureID(StoreNo, cdspicType_PartsOrder, PartOrderNo)
            If PicID <> 0 Then Attach = GetDatabasePictureToTempFile(PicID)

            PartOrderToHTML = S
        End With
        DisposeDA(Part, ServiceOrder)
        Exit Function

        ' below was for testing
        '  Dim F As String
        '  F = localdesktopfolder & "PartOrder.htm"
        '  WriteFile F, S, True
        '  ShellOut_URL F
    End Function

    Public Function ChargeBackLetterHTML(ByVal PON As Integer, ByVal LetterType As Integer, ByVal StoreNum As Integer, ByVal Amount As Decimal, ByVal InvoiceNo As String, Optional ByRef vEMail As String = "", Optional ByRef vName As String = "", Optional ByRef Attach As String = "") As String
        '::::ChargeBackLetterHTML
        ':::SUMMARY
        ': HTML  code for ChargeBack Letter.
        ':::DESCRIPTION
        ': This function is contains HTML code for Page Setup,Store Information,Page Title (w/ date),Vendor Information,Claim Information,Page CleanUp in ChargeBack Letter.Useful to handle errors and get vendor name from Service parts order form to ChargeBack Letter.
        ':::PARAMETERS
        ': - PON - Indicates Part Order Number.
        ': - LetterType -
        ': - StoreNum - Indicates Store Number.
        ': - Amount - Indicates the ChargeBack amount.
        ': - InvoiceNo - Indicates the number given by Manufacturer.
        ': - vEMail - Indicates the vendor email.
        ': - Attach -
        ':::RETURN
        ': String - Returns ChargeBack letter as a string.
        Dim S As String, Oper As String
        Dim Part As clsServicePartsOrder

        Dim VendorName As String
        Dim vAddress As String, vAddress2 As String, vAddress3 As String
        Dim VZip As String, VPhone As String, VFax As String

        On Error Resume Next
        Part = New clsServicePartsOrder
        If Not Part.Load(PON, "#ServicePartsOrderNo") Then
            DisposeDA(Part)
            Exit Function
        End If

        VendorName = Part.Vendor
        vName = VendorName
        '  GetVendorName (VendorName), (VendorName), vAddress, vAddress2, vAddress3, VZip, VPhone, VFax, (VendorName), vEMail
        If UseQB() Then
            QBGetVendorName(VendorName, VendorName, vAddress, vAddress2, vAddress3, VZip, VPhone, VFax, , vEMail)
        Else
            GetVendorName(VendorName, VendorName, vAddress, vAddress2, vAddress3, VZip, VPhone, VFax, , vEMail)
        End If
        Oper = ChargeBackLetterOperationDesc(LetterType, Amount, InvoiceNo)


        S = ""

        ' templates
        S = S & "" & vbCrLf
        '    S = S & "  <></>" & vbCrLf

        '------- Page Setup
        S = S & "<html>" & vbCrLf
        S = S & " <head>" & vbCrLf
        S = S & "  <title>CREDIT DEPARTMENT</title>" & vbCrLf
        S = S & "  <meta name=""GENERATOR"" content=""" & SoftwareVersion(True, False, True) & """>" & vbCrLf
        S = S & " </head>" & vbCrLf
        S = S & " <body bgcolor=#FFFFFF text=#000000 vlink=#FF0000 link=#FFFF00 hspace=0 vspace=0>" & vbCrLf

        S = S & "  <table border=0 cellspacing=0 cellpadding=0 width='100%'>" & vbCrLf ' alignment table


        S = S & "   <tr><td colspan=3 align=center>" & vbCrLf
        '------- Page Title (w/ date)
        S = S & "<table border=0 width='100%'>" & vbCrLf
        S = S & "  <tr><td align=center><b><font size=+2>CREDIT DEPARTMENT</font></b></td></tr>" & vbCrLf
        S = S & "  <tr><td align=right>" & Format(Today, "mm/dd/yyyy") & "</td></tr>" & vbCrLf
        S = S & "</table>" & vbCrLf

        S = S & "   </tr></td>" & vbCrLf    ' alignment table
        S = S & "   <tr><td colspan=3 align=left>" & vbCrLf

        '------- Store Information
        S = S & "      <table border=1 cellspacing=0 cellpadding=0  width=400><tr>" & vbCrLf
        S = S & "      <td>" & vbCrLf
        S = S & "        <table border=0>" & vbCrLf
        S = S & "          <tr><td><b>From:</b></td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & StoreSettings.Name & "</font></td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & StoreSettings.Address & "</td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & StoreSettings.City & "</td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & DressAni(StoreSettings.Phone) & "</td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & StoreSettings.Email & "</td></tr>" & vbCrLf
        S = S & "        </table>" & vbCrLf
        S = S & "      </td>" & vbCrLf
        S = S & "      </tr></table><br>" & vbCrLf

        S = S & "   </tr></td>" & vbCrLf
        S = S & "   <tr><td colspan=3 align=left>" & vbCrLf

        '------- Vendor Information
        S = S & "      <table border=1 cellspacing=0 cellpadding=0 width=400><tr>" & vbCrLf
        S = S & "      <td>" & vbCrLf
        S = S & "        <table border=0>" & vbCrLf
        S = S & "          <tr><td><b>To:</b></td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & VendorName & "</font></td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & vAddress & "</td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & vAddress2 & IIf(Len(vAddress3) > 0, "", " " & VZip) & "</td></tr>" & vbCrLf
        If Len(vAddress3) > 0 Then
            S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & vAddress3 & " " & VZip & "</td></tr>" & vbCrLf
        End If
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & DressAni(VPhone) & IIf(Len(VFax) > 0, "  FAX: " & DressAni(VFax), "") & " " & VZip & "</td></tr>" & vbCrLf
        S = S & "          <tr><td>&nbsp;&nbsp;&nbsp;" & vEMail & "</td></tr>" & vbCrLf
        S = S & "        </table>" & vbCrLf
        S = S & "      </td>" & vbCrLf
        S = S & "      </tr></table><br>" & vbCrLf

        S = S & "   </tr></td>" & vbCrLf
        S = S & "   <tr><td colspan=3 align=left>" & vbCrLf

        '------- Claim Information
        S = S & "<b>RE: Service Parts Order #" & PON & "</b><br/>" & vbCrLf

        S = S & EmailChargeBackBodyHTML
        S = S & "<br/>" & vbCrLf

        S = S & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;As per the attached service order, we are " & Oper & ".<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "Thank You,<br/>" & vbCrLf
        S = S & "<br/>" & vbCrLf
        S = S & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & StoreSettings.Name & "<br/>" & vbCrLf

        S = S & "  </td></tr></table>" & vbCrLf  ' end of alignment table

        S = S & "<br/><br/><hr><br/><br/>" & vbCrLf

        S = S & PartOrderToHTML(PON, , , , , Attach)

        '------- Page CleanUp
        S = S & " </body>" & vbCrLf
        S = S & "</html>" & vbCrLf

        DisposeDA(Part)
        ChargeBackLetterHTML = S
        Exit Function

        ' below was for testing
        '  Dim F As String
        '  F = localdesktopfolder & "PartOrder.htm"
        '  WriteFile F, S, True
        '  ShellOut_URL F
    End Function

    Public Function DescribeChargeBackOption(ByRef nVal As Integer) As String
        '::::DescribeChargeBackOption
        ':::SUMMARY
        ': Describe each payment option for chargeback.
        ':::DESCRIPTION
        ': This function is mostly used to describe fileds like Charge Back,Deduct from Invoice,Request Credit in parts order form.
        ':::PARAMETERS
        ':::RETURN
        ': String - Returns ChargeBack option as a string.
        Select Case nVal
            Case 0 : DescribeChargeBackOption = "Charge Back"
            Case 1 : DescribeChargeBackOption = "Deduct from Invoice"
            Case 2 : DescribeChargeBackOption = "Credit"
            Case Else : DescribeChargeBackOption = "???"
        End Select
    End Function

    Public Function ChargeBackLetterOperationDesc(ByVal LetterType As Integer, ByVal Amount As Decimal, ByVal InvoiceNo As String) As String
        '::::ChargeBackLetterOperationDesc
        ':::SUMMARY
        ': Provides Description about ChargeBack letter operations.
        ':::DESCRIPTION
        ': This function is provides description about Charge Back letter operation.Depending up on selection of letter type, amount is deducted in different ways.
        ':::PARAMETERS
        ': - LetterType - Indicates the tpe of selection of ChargeBack amount.
        ': - Amount - Indicates the ChargeBack amount.
        ': - InvoiceNo - Indicates the number given by manufacturer
        ':::RETURN
        ': String - Returns the ChargeBack letter Operation description as a string.

        Dim Oper As String
        Select Case LetterType
            Case 0
                Oper = "charging back the repair cost of " & FormatCurrency(Amount)
                Oper = Oper & vbCrLf & "Please send credit memo for this amount"
            Case 1
                Oper = "deducting " & FormatCurrency(Amount)
                If Len(InvoiceNo) Then
                    Oper = Oper & " from Invoice No. " & InvoiceNo
                Else
                    Oper = Oper & " from the invoice."
                End If
            Case 2
                Oper = "requesting a credit of " & FormatCurrency(Amount)
            Case Else
                Oper = "requesting " & FormatCurrency(Amount)
        End Select
        ChargeBackLetterOperationDesc = Oper
    End Function

    Public Function GetMailIndexByServiceCallNo(ByRef ServiceCallNumber As Long) As Long
        '::::GetMailIndexByServiceCallNo
        ':::SUMMARY
        ': Gets Mail Index with Service call number.
        ':::DESCRIPTION
        ': By calling this function, we gets Mail Index based on service call number.
        ': Service Call Number is mostly used in part order form.
        ':::PARAMETERS
        ': - ServiceCallNumber - Indicates Service number.
        ':::RETURN
        ': - Long - Returns Mail Index as a long.
        Dim cServ As clsServiceOrder
        cServ = New clsServiceOrder
        If cServ.Load(CStr(ServiceCallNumber), "#ServiceOrderNo") Then
            GetMailIndexByServiceCallNo = cServ.MailIndex
        Else
            GetMailIndexByServiceCallNo = 0
        End If
        DisposeDA(cServ)
    End Function
End Module

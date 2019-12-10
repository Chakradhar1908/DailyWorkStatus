Imports Microsoft.VisualBasic.Interaction
Module modSetup
    Public Const LocalDriveLetter As String = "C"
    Public Const LocalDrive As String = LocalDriveLetter & ":"
#If True Then
    Public Const FONT_C39_ONEINCH As String = "Code39OneInch-Regular"
    Public Const FONT_C39_HALFINCH As String = "Code39HalfInch-Regular"
    Public Const FONT_C39_QUARTERINCH As String = "Code39QuarterInch-Regular"
    Public Const FONT_C39_SLIM As String = "Code39Slim-Regular"
    Public Const FONT_C39_WIDE As String = "Code39Wide-Regular"
    Public Const FONT_C39_SMALL_HIGH As String = "Code39SmallHigh-Regular"
    Public Const FONT_C39_SMALL_MEDIUM As String = "Code39SmallMedium-Regular"
    Public Const FONT_C39_SMALL_LOW As String = "Code39SmallLow-Regular"
    Public Const FONT_C128_REGULAR As String = "Code 128"
#Else
Public Const FONT_C39_ONEINCH As String = "xCode39OneInch-Regular"
Public Const FONT_C39_HALFINCH As String = "xCode39HalfInch-Regular"
Public Const FONT_C39_QUARTERINCH As String = "xCode39QuarterInch-Regular"
Public Const FONT_C39_SLIM As String = "xCode39Slim-Regular"
Public Const FONT_C39_WIDE As String = "xCode39Wide-Regular"
Public Const FONT_C39_SMALL_HIGH As String = "xCode39SmallHigh-Regular"
Public Const FONT_C39_SMALL_MEDIUM As String = "xCode39SmallMedium-Regular"
Public Const FONT_C39_SMALL_LOW As String = "xCode39SmallLow-Regular"
Public Const FONT_C128_REGULAR As String = "xCode 128 Regular"
#End If
    Public Const Setup_MaxStores_DB As Integer = 32              ' Store count in the Database
    Public Const ProgramName As String = "WinCDS"
    Public Const Setup_2Data_ManufMaxLen As Integer = 16         '###MANUFLENGTH16
    Public Const Setup_2Data_DescMaxLen As Integer = 138
    Public Const Setup_MaxStores As Integer = 32
    Public Const RegistryAppName As String = CompanyName
    Public Const RegistrySection As String = ProgramName
    Public Const LocalRoot As String = LocalDrive & "\"
    Public Const RemoteRoot As String = RemoteDrive & "\"
    Public Const RemoteDriveLetter As String = "I"
    Public Const LocalRootFolder As String = LocalRoot
    Public Const CompanyName As String = "Custom Design Software"
    Public Const RemoteDrive As String = RemoteDriveLetter & ":"
    Public Const cdsFirstDayOfWeek As Integer = vbSunday
    Public Const MaxLines As Integer = 260 ' Used to determine number of lines on BillOSale?
    Public Const CCPROC_TC As String = "Transaction Central"
    Public Const ProgramMessageTitle As String = ProgramName
    Public Const AdminContactProgram As String = ProgramName
    Public Const AdminContactName As String = CompanyName ' "CDS" ' "Jerry Katz"
    Public Const AdminContactCompany As String = CompanyName
    Public Const AdminContactEmail As String = "kris@wincdspro.com" '"jerryk@customdesignsoftware.net"
    Public Const AdminContactWeb As String = CompanyURL
    Public Const AdminContactPhone As String = "(800) 884-0806"
    Public Const AdminContactPhone2 As String = "(248) 669-9512"
    Public Const AdminContactFax As String = "(248) 694-2104" '"(248) 669-9514"
    Public Const StoreFinanceAsDelivered As Boolean = False             'BFH20161026 True was what the software always was...   We are changing that
    Public Const CompanyURL As String = "http://" & CompanyURL_BARE & "/"
    Public Const CompanyURL_BARE As String = "www.wincdspro.com"
    Public Const ProgramErrorTitle As String = ProgramName & " Error"
    Public Const CopyrightYear As String = "2019"
    Public Const CopyrightFirstYear As String = "1991"
    Public Const CopyrightYears As String = CopyrightFirstYear & "-" & CopyrightYear
    Public Const CCPROC_NONE As String = "None"
    Public Const CCPROC_NA As String = ""
    Public Const CCPROC_XC As String = "X-Charge"
    Public Const CCPROC_XL As String = "eXpressLink"
    Public Const CCPROC_CM As String = "Credomatic"
    Public Const CCPROC_CI As String = "ChargeItPro"
    Public Const Setup_MaxKitItems As Integer = 10
    Public Const Setup_2Data_CommMaxLen As Integer = 138
    Public Const FONTTTF_C39_WIDE As String = "C39WIDE.TTF"
    Public Const FONTTTF_C39_ONEINCH As String = "C39ONE.TTF"
    Public Const FONTTTF_C39_SLIM As String = "C39SLIM.TTF"
    Public Const FONTTTF_C39_HALFINCH As String = "C39HALF.TTF"
    Public Const FONTTTF_C39_QUARTERINCH As String = "C39QRTR.TTF"
    Public Const FONTTTF_C39_SMALL_HIGH As String = "C39SHIGH.TTF"
    Public Const FONTTTF_C39_SMALL_MEDIUM As String = "C39SMED.TTF"
    Public Const FONTTTF_C39_SMALL_LOW As String = "C39SLOW.TTF"
    Public Const FONTTTF_C128_REGULAR As String = "CODE128.TTF"
    Public URLTryAlt As Long
    Public Const CompanyURL2 As String = "http://" & CompanyURL_BARE2 & "/"
    Public Const CompanyURL_BARE2 As String = "www.wincds.net"

    Public ReadOnly Property Setup_2Data_StyleMaxLen() As Integer
        Get
            '::::Setup_2Data_StyleMaxLen
            ':::SUMMARY
            ':Returns the maximum style length.
            ':::DESCRIPTION
            ': Returns the maximum length of the style field.
            ':
            ':::NOTES
            ': Formerly was a constant, the original value of which was 16.  Was changed to a property
            ': to support development of expanding this, although practical considerations are the hard-
            ': coded reports which will not support an arbitrary increase in style length.
            ':
            ': Along with this change, the database field was extended to 128 characters for experimentation.
            ': While the software works with a 128 length style, it is not considered stable...
            ':::RETURNS
            ': String - A path to a file.
            ':::SEE ALSO
            ': - FileBanking, FileGenledger, FilePayroll, FileAccountPayable
            '
            Setup_2Data_StyleMaxLen = 16                                '###STYLELENGTH16
            'Setup_2Data_StyleMaxLen = 5                                '###STYLELENGTH16
            If IsCDSComputer("laptop") Then Setup_2Data_StyleMaxLen = 128
            If IsCDSComputer("laptop") Then Setup_2Data_StyleMaxLen = 32
        End Get
    End Property
    Public Sub UpdatePermStatus()
        '::::UpdatePermStatus
        ':::SUMMARY
        ': Hook to the frmPermissionMontior form.
        If Not IsFormLoaded("frmPermissionMonitor") Then Exit Sub
        frmPermissionMonitor.Update()
    End Sub
    Public Function GetStoreTax1() As Double
        '::::GetStoreTax1
        ':::SUMMARY
        ':Returns the TAX1 rate
        ':::DESCRIPTION
        ':Returns the primary tax rate.  The rate can be entered on the store setup page.
        GetStoreTax1 = StoreSettings.SalesTax
    End Function
    Public Function AdminContactString(
    Optional ByVal Format As Integer = 0,
    Optional ByVal Version As Boolean = True,
    Optional ByVal Company As Boolean = True,
    Optional ByVal Name As Boolean = False,
    Optional ByVal WebSite As Boolean = False,
    Optional ByVal Email As Boolean = False,
    Optional ByVal Phone As Boolean = True,
    Optional ByVal Phone2 As Boolean = False,
    Optional ByVal Fax As Boolean = False,
    Optional ByVal Copyright As Boolean = False) As String
        '::::AdminContactString
        ':::SUMMARY
        ': The standard way to display contact information to the user.
        ':::DESCRIPTION
        ': Whenever admin contact is displayed to the user, this is the function to call.  It is referenced
        ': by its various parameters to display as much or as little information as is appropriate given
        ': the calling circumstances.
        ':
        ': As such, when calling this function, find similar usages of the functino to generate the desired
        ': output.
        ':
        ':::PARAMETERS
        ': - Format
        ': - Version
        ': - Company
        ': - Name
        ': - WebSite
        ': - Email
        ': - Phone
        ': - Phone2
        ': - Fax
        ': - Copyright
        ':
        ':::RETURN
        ': String - Admin Contact String
        Dim S As String
        S = ""

        Select Case Format
            Case 3
                S = S & SoftwareCopyright(True)
            Case 2
                S = S & AdminContactName & " at " & AdminContactPhone2 & " or " & AdminContactEmail
            Case 1
                If Company And Name Then
                    S = S & AdminContactName & " at " & AdminContactCompany & " at "
                ElseIf Company Or Name Then
                    S = S & IIf(Company, AdminContactCompany, AdminContactName) & " at "
                Else
                    S = S & AdminContactCompany & " at "
                End If
                If Phone And Email Then
                    S = S & AdminContactPhone & " or " & AdminContactEmail
                ElseIf Phone Or Email Then
                    S = S & IIf(Phone, AdminContactPhone, AdminContactEmail)
                Else
                    S = S & AdminContactPhone2
                End If

                If WebSite Then S = S & AdminContactWeb & vbCrLf
                If Fax Then S = S & "Phone: " & AdminContactPhone2 & vbCrLf
            Case Else
                If Version Then S = S & "Version: " & SoftwareVersion(False, True, True) & vbCrLf
                If Company Then S = S & "Company: " & AdminContactCompany & vbCrLf
                If Name Then S = S & "Name:    " & AdminContactName & vbCrLf
                If WebSite Then S = S & "Web:     " & AdminContactWeb & vbCrLf
                If Email Then S = S & "eMail:   " & AdminContactEmail & vbCrLf
                If Phone Then S = S & "Phone:   " & AdminContactPhone & vbCrLf
                If Phone2 Then S = S & "Phone 2: " & AdminContactPhone2 & vbCrLf
                If Fax Then S = S & "Fax:     " & AdminContactFax & vbCrLf
                If Copyright Then
                    S = S & vbCrLf
                    S = S & SoftwareCopyright()
                End If
        End Select
        AdminContactString = S
    End Function
    Public Function SoftwareVersion(Optional ByVal ForDisplay As Boolean = True, Optional ByVal WithRevision As Boolean = True, Optional ByVal WithProgramName As Boolean = False, Optional ByVal wHash As Boolean = False, Optional ByVal ShowStoreName As Boolean = False, Optional ByVal ShowOSVersion As Boolean = False) As String
        '::::SoftwareVersion
        ':::SUMMARY
        ': Returns the version of the software running.
        ':::DESCRIPTION
        ': Returns in various forms the current software version.
        ':
        ': This is the main go-to source for generating version numbers and so can be parameterized in several
        ': ways to provide the version required in the situation.
        ':
        ':::PAREMETERS
        ': - ForDisplay - Boolean.  Optional.  True if this is for display (suppresses last numbers).
        ': - WithRevision - Boolean.  Optional.  True if the revision (build number) should be displayed.
        ': - WithProgramName - Boolean.  Optional.  True if the program name should be included.  False for version only.
        ': - wHash - Boolean.  Optional.  True to include the version hash.
        ': - ShowStoreName - Boolean.  Optional.  True to include Store Name.
        ': - ShowOSVersion - Boolean.  Optional.  True to display Operating System version in result.
        ':
        ':::RETURN
        ': String - The version as a string.
        ':::SEE ALSO
        ': - SoftwareVersionForLog

        'LogStartup "SoftwareVersion: a"
        If ForDisplay Then
            'LogStartup "SoftwareVersion: b"
            SoftwareVersion = WinCDSMajorVersion() & ".0" ' & WinCDSMinorVersion
            'LogStartup "SoftwareVersion: c"
            If WithRevision Then SoftwareVersion = SoftwareVersion & " (Revision: " & WinCDSRevisionNumber() & ")"
            'LogStartup "SoftwareVersion: d"
        Else
            'LogStartup "SoftwareVersion: e"
            SoftwareVersion = WinCDSMajorVersion() & "." & WinCDSMinorVersion()
            'LogStartup "SoftwareVersion: f"
            If WithRevision Then SoftwareVersion = SoftwareVersion & "." & WinCDSRevisionNumber() & "." & WinCDSBuildNumber()
            'LogStartup "SoftwareVersion: g"
        End If
        'LogStartup "SoftwareVersion: h"
        If WithProgramName Then SoftwareVersion = AdminContactProgram & " v" & SoftwareVersion
        'LogStartup "SoftwareVersion: i"
        If wHash Then SoftwareVersion = SoftwareVersion & " (HASH=" & SoftwareVersionHash() & ")"
        'LogStartup "SoftwareVersion: j"
        If ShowStoreName Then SoftwareVersion = SoftwareVersion '& " @ " & StoreSettings(1).Name
        'LogStartup "SoftwareVersion: k"
        If ShowOSVersion Then SoftwareVersion = SoftwareVersion & " {" & GetWinVerNumber() & "}"
        'LogStartup "SoftwareVersion: l"
        'LogStartup "SoftwareVersion: m"
    End Function
    Public Function SoftwareCopyright(Optional ByVal Shortx As Boolean = False) As String
        '::::SoftwareCopyright
        ':::SUMMARY
        ': The official copyright string
        ':::DESCRIPTION
        ': Returns the software copyright string.  An long and short form are possible.
        SoftwareCopyright = ""
        SoftwareCopyright = SoftwareCopyright & "(c) " & CopyrightYears & ", " & AdminContactCompany & "."
        If Not Shortx Then SoftwareCopyright = SoftwareCopyright & vbCrLf & "All rights reserved." & vbCrLf
    End Function
    Public Function SoftwareVersionHash() As String
        '::::SoftwareVersionHash
        ':::SUMMARY
        ': Returns a version hash.  Used as an attempt to bypass user improvisation of version numbers.
        ':::DESCRIPTION
        ':Returns a base 64 encoding of the current version.
        SoftwareVersionHash = EncodeBase64String(BuildDate() & " " & BuildTime() & "-" & WinCDSRevisionNumber() & "." & WinCDSBuildNumber() & "-" & BuildComputer()) ' & "-" & GetWinVerNumber)
    End Function

    Public Function GetXCTransactionFolder(Optional ByVal StoreNum As Integer = 0, Optional ByVal MOTO As Boolean = False) As String
        '::::GetXCTransactionFolder
        ':::SUMMARY
        ': Returns the XCharge Transaction folder.
        Dim K As String
        K = "XCFolder" & IIf(StoreNum <= 0, "", Format(StoreNum, "00")) & IIf(MOTO, "-MOTO", "")
        GetXCTransactionFolder = GetCDSSetting(K, , "XCharge")
        If Right(GetXCTransactionFolder, 1) = "\" Then GetXCTransactionFolder = Left(GetXCTransactionFolder, Len(GetXCTransactionFolder) - 1)
    End Function

    Public ReadOnly Property WebUpdateURL() As String
        Get
            WebUpdateURL = Switch(URLTryAlt = 2, CompanyURL2, True, CompanyURL) & "webupdate/"
        End Get
    End Property

End Module

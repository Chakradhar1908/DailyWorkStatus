Module modWinCDSLicenses
    Public Const LICENSE_DEMO As String = "DEMO"
    Public Const LICENSE_DEMO_STORE_ALLOWANCE As Integer = 2
    Public Const LICENSE_STORES_1 As String = "23986"
    Public Const LICENSE_STORES_2 As String = "46523"
    Public Const LICENSE_STORES_3 As String = "59893"
    Public Const LICENSE_STORES_4 As String = "43782"
    Public Const LICENSE_STORES_5 As String = "12643"
    Public Const LICENSE_STORES_6 As String = "45861"
    Public Const LICENSE_STORES_7 As String = "43982"
    Public Const LICENSE_STORES_8 As String = "13259"

    Public Const LICENSE_STORES_9 As String = "28932"
    Public Const LICENSE_STORES_10 As String = "90892"
    Public Const LICENSE_STORES_11 As String = "10987"
    Public Const LICENSE_STORES_12 As String = "33897"
    Public Const LICENSE_STORES_13 As String = "62891"
    Public Const LICENSE_STORES_14 As String = "80196"
    Public Const LICENSE_STORES_15 As String = "44321"
    Public Const LICENSE_STORES_16 As String = "48941"

    Public Const LICENSE_STORES_17 As String = "38927"
    Public Const LICENSE_STORES_18 As String = "09823"
    Public Const LICENSE_STORES_19 As String = "1-384"
    Public Const LICENSE_STORES_20 As String = "55304"
    Public Const LICENSE_STORES_21 As String = "91284"
    Public Const LICENSE_STORES_22 As String = "87432"
    Public Const LICENSE_STORES_23 As String = "22-14"
    Public Const LICENSE_STORES_24 As String = "78643"

    Public Const LICENSE_STORES_25 As String = "92834"
    Public Const LICENSE_STORES_26 As String = "28309"
    Public Const LICENSE_STORES_27 As String = "44623"
    Public Const LICENSE_STORES_28 As String = "40928"
    Public Const LICENSE_STORES_29 As String = "48493"
    Public Const LICENSE_STORES_30 As String = "94532"
    Public Const LICENSE_STORES_31 As String = "94801"
    Public Const LICENSE_STORES_32 As String = "01024"
    Public Const LICENSE_INSTALLMENT As String = "I589423"
    Private mActiveLocations As Integer
    Public Const LICENSE_DISPATCHTRACK As String = "DDT34892"
    '    Public Property Get LicensedNoOfStores() as integer :     
    '    LicensedNoOfStores = ConvertWinCDSLicenseCode(GetWinCDSLicense) 
    'End Function
    'Public ReadOnly Property LicensedNoOfStores() As Integer
    '    Get
    '        LicensedNoOfStores = ConvertWinCDSLicenseCode(GetWinCDSLicense)
    '    End Get
    'End Property

    Public Function LicensedNoOfStores() As Integer
        LicensedNoOfStores = ConvertWinCDSLicenseCode(GetWinCDSLicense)
    End Function

    Public Function ConvertWinCDSLicenseCode(ByVal Code As String, Optional ByRef Valid As Boolean = False) As Integer
        Dim I As Integer

        ConvertWinCDSLicenseCode = 0
        ' Takes a license code, returns the number of licensed stores.
        Valid = False
        For I = 1 To Setup_MaxStores_DB
            If Code = WinCDSLicenseCode(I) Then ConvertWinCDSLicenseCode = I : Valid = True : Exit Function
        Next

        If Code = LICENSE_DEMO Then
            ConvertWinCDSLicenseCode = LICENSE_DEMO_STORE_ALLOWANCE
            Valid = True
            Exit Function
        End If
    End Function

    Public Function GetWinCDSLicense() As String
        GetWinCDSLicense = GetCDSSetting("License")
    End Function

    Public Function WinCDSLicenseCode(Optional ByVal StoreCount As Integer = 0) As String
        WinCDSLicenseCode = 0
        If StoreCount <= 0 Or StoreCount > Setup_MaxStores_DB Then Exit Function
        '  If StoreCount > Setup_MaxStores Then Exit Function
        '###STORECOUNT32
        WinCDSLicenseCode = Choose(StoreCount,
      LICENSE_STORES_1, LICENSE_STORES_2, LICENSE_STORES_3, LICENSE_STORES_4,
      LICENSE_STORES_5, LICENSE_STORES_6, LICENSE_STORES_7, LICENSE_STORES_8,
      LICENSE_STORES_9, LICENSE_STORES_10, LICENSE_STORES_11, LICENSE_STORES_12,
      LICENSE_STORES_13, LICENSE_STORES_14, LICENSE_STORES_15, LICENSE_STORES_16,
      LICENSE_STORES_17, LICENSE_STORES_18, LICENSE_STORES_19, LICENSE_STORES_20,
      LICENSE_STORES_21, LICENSE_STORES_22, LICENSE_STORES_23, LICENSE_STORES_24,
      LICENSE_STORES_25, LICENSE_STORES_26, LICENSE_STORES_27, LICENSE_STORES_28,
      LICENSE_STORES_29, LICENSE_STORES_30, LICENSE_STORES_31, LICENSE_STORES_32)
    End Function

    Public ReadOnly Property ActiveNoOfLocations() As Integer
        Get
            ActiveNoOfLocations = NoOfActiveLocations
        End Get
    End Property

    Public ReadOnly Property NoOfActiveLocations() As Integer
        Get
            Dim I As Integer
            If mActiveLocations > 0 Then NoOfActiveLocations = mActiveLocations : Exit Property

            For I = 1 To Setup_MaxStores
                If Dir(StoreFile(I)) = "" Then Exit For
                mActiveLocations = I
            Next
            If LicensedNoOfStores() > mActiveLocations Then mActiveLocations = LicensedNoOfStores()
            NoOfActiveLocations = mActiveLocations
        End Get
    End Property

    Public Property License() As String
        Get
            License = GetWinCDSLicense()
        End Get
        Set(value As String)
            On Error Resume Next
            If value = "" Then value = LICENSE_DEMO ' LICENSE_STORES_1
            SetWinCDSLicense(value)
        End Set
    End Property

    Public Function SetWinCDSLicense(ByVal vData As String) As String
        SetWinCDSLicense = SaveCDSSetting("License", vData)
        mActiveLocations = 0
    End Function

    Public ReadOnly Property Installment() As Boolean
        Get
            Installment = InstallmentLicenseValid(InstallmentLicense)
        End Get
    End Property

    Public Function InstallmentLicenseValid(ByVal S As String) As Boolean
        InstallmentLicenseValid = IsIn(S, LICENSE_INSTALLMENT, "TEST")
    End Function

    Public Function IsDemo() As Boolean
        IsDemo = (License = LICENSE_DEMO) Or (Not LicenseValid(License)) Or (UCase(StoreSettings.Name) = "DEMO")
        If IP_CONTROL Then IsDemo = IsDemo Or IPAddressIsBanned
    End Function

    Public Function DemoExpirationDate(Optional ByVal Reset As Boolean = False) As Date
        Dim R As String, T As String
        Const fDemoExpirationDate = "DemoExpirationDate"
        If Not IsDemo() Then
            DemoExpirationDate = YearAdd(Today, 1)
            Exit Function
        End If

        R = GetConfigTableValue(fDemoExpirationDate)
        T = DateAdd("d", 30, Today)
        If IsDate(R) Then
            If DateAfter(R, T) Then R = ""
        End If
        If Not IsDate(R) Or Reset Then
            R = T
            SetConfigTableValue(fDemoExpirationDate, R)
        End If

        DemoExpirationDate = DateValue(R)
    End Function

    Public Function LicenseValid(Optional ByVal vData As String = "") As Boolean
        LicenseValid = WinCDSLicenseValid(IIf(vData <> "", vData, License))
    End Function

    Public Function WinCDSLicenseValid(ByVal S As String) As Boolean
        ConvertWinCDSLicenseCode(S, WinCDSLicenseValid)
    End Function

    Public Function IsDemoExpired() As Boolean
        If Not IsDemo() Then Exit Function
        IsDemoExpired = DateAfter(Today, DemoExpirationDate)
    End Function

    Public Function NotifyDemoExpired(Optional ByVal CommandLine As String = "") As Boolean
        Const tCaption As String = "DEMO Trial Period Expired"
        If IsDemoExpired Then
            If IsIDE() Then
                Dim R As VBA.VbMsgBoxResult, Discard, S As String
                S = ""
                S = S & "You are running in the VB6 IDE." & vbCrLf2
                S = S & "This data is set as a DEMO INSTALL and is EXPIRED." & vbCrLf
                S = S & "Expiration Date: " & DemoExpirationDate() & vbCrLf
                S = S & "Please select from the following options: " & vbCrLf
                S = S & "    Abort - Encounter the user fail message." & vbCrLf
                S = S & "    Retry - Open " & ProgramName & " regardless of this error." & vbCrLf
                S = S & "   Ignore - Reset the demo expiry date."
                R = MessageBox.Show(S, tCaption, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation)
                Select Case R
                    Case vbAbort
                    Case vbRetry : Exit Function
                    Case vbIgnore : Discard = DemoExpirationDate(True) : Exit Function
                End Select
            End If

            NotifyDemoExpired = True
            If CommandLine = "" Then
                If frmDemoNotify.RequireLicenseOrQuit Then
                    NotifyDemoExpired = False
                    Exit Function
                End If
            End If
        End If
    End Function
End Module

Imports QBFC10Lib
Module modQuickBooks_CommBase
    Public mQBActiveStore As Integer, mQBAS_JustSet As Boolean
    Public QBConnOpen As Boolean, QBSessOpen As Boolean
    Private mQBFCVer As Integer
    Private Const QBFCVerDefault As Integer = 6
    Public Function QB_VendorQuery_Vendor(ByVal Vendor As String, Optional ByRef RET As Integer = 0, Optional ByRef RetMsg As String = "") As Object
        Select Case QBFCVersion
            Case 5 : QB_VendorQuery_Vendor = QB5_VendorQuery_Vendor(Vendor, RET, RetMsg)
            Case 6 : QB_VendorQuery_Vendor = QB6_VendorQuery_Vendor(Vendor, RET, RetMsg)
            Case 7 : QB_VendorQuery_Vendor = QB7_VendorQuery_Vendor(Vendor, RET, RetMsg)
            Case 8 : QB_VendorQuery_Vendor = qb8_VendorQuery_Vendor(Vendor, RET, RetMsg)
            Case 10 : QB_VendorQuery_Vendor = QB10_VendorQuery_Vendor(Vendor, RET, RetMsg)
            Case 11 : QB_VendorQuery_Vendor = QB11_VendorQuery_Vendor(Vendor, RET, RetMsg)
            Case 12 : QB_VendorQuery_Vendor = QB12_VendorQuery_Vendor(Vendor, RET, RetMsg)
            Case 13 : QB_VendorQuery_Vendor = QB13_VendorQuery_Vendor(Vendor, RET, RetMsg)
        End Select
    End Function

    Public Function IfNNGetValue(ByVal Field As Object) As Object
        If Not Field Then Field = Field : Exit Function
        If Field Is Nothing Then Exit Function
        IfNNGetValue = Field.GetValue
    End Function

    Public ReadOnly Property QB_File(Optional ByVal StoreNo As Integer = 0) As String
        Get
            If StoreNo = 0 Then StoreNo = QBActiveStore
            QB_File = GetQBSetupValue("file", StoreNo)
        End Get
    End Property

    Public Function QBObjectsExist(Optional ByRef FailReason As String = "") As Boolean
        QBObjectsExist = QBVersionExist(QBFCVersion, FailReason)
    End Function

    Public Function QB_ClassExists(ByVal ClassID As String) As Boolean
        Dim RET As Integer, RetMsg As String
        On Error GoTo Failed
        Select Case QBFCVersion
            Case 5
                Dim X5 As QBFC5Lib.IClassRet
                X5 = QB5_ClassQuery_Class(ClassID, RET, RetMsg)
                If RET = 0 Then QB_ClassExists = True
            Case 6
                Dim X6 As QBFC6Lib.IClassRet
                X6 = QB5_ClassQuery_Class(ClassID, RET, RetMsg)
                If RET = 0 Then QB_ClassExists = True
            Case 7
                Dim X7 As QBFC7Lib.IClassRet
                X7 = QB7_ClassQuery_Class(ClassID, RET, RetMsg)
                If RET = 0 Then QB_ClassExists = True
            Case 8
                Dim X8 As QBFC8Lib.IClassRet
                X8 = QB8_ClassQuery_Class(ClassID, RET, RetMsg)
                If RET = 0 Then QB_ClassExists = True
            Case 10
                Dim X10 As QBFC10Lib.IClassRet
                X10 = QB10_ClassQuery_Class(ClassID, RET, RetMsg)
                If RET = 0 Then QB_ClassExists = True
            Case 11
                Dim X11 As QBFC11Lib.IClassRet
                X11 = QB11_ClassQuery_Class(ClassID, RET, RetMsg)
                If RET = 0 Then QB_ClassExists = True
            Case 12
                Dim X12 As QBFC12Lib.IClassRet
                X12 = QB12_ClassQuery_Class(ClassID, RET, RetMsg)
                If RET = 0 Then QB_ClassExists = True
            Case 13
                Dim X13 As QBFC13Lib.IClassRet
                X13 = QB13_ClassQuery_Class(ClassID, RET, RetMsg)
                If RET = 0 Then QB_ClassExists = True
        End Select
Failed:
        If RET <> 0 Then ActiveLog("QB::QB_ClassExists: Error [" & RET & "] " & RetMsg, 4)
    End Function

    Public Function QB_CustomerExistsByName(ByVal CustomerName As String) As Boolean
        Dim RET As Integer, RetMsg As String
        On Error GoTo Failed
        Select Case QBFCVersion
            Case 5
                Dim X5 As QBFC5Lib.ICustomerRet
                X5 = QB5_CustomerQuery_Name(CustomerName, RET)
                If RET = 0 Then QB_CustomerExistsByName = True : Exit Function
            Case 6
                Dim X6 As QBFC6Lib.ICustomerRet
                X6 = QB6_CustomerQuery_Name(CustomerName, RET)
                If RET = 0 Then QB_CustomerExistsByName = True : Exit Function
            Case 7
                Dim X7 As QBFC7Lib.ICustomerRet
                X7 = QB7_CustomerQuery_Name(CustomerName, RET)
                If RET = 0 Then QB_CustomerExistsByName = True : Exit Function
            Case 8
                Dim X8 As QBFC8Lib.ICustomerRet
                X8 = qb8_CustomerQuery_Name(CustomerName, RET)
                If RET = 0 Then QB_CustomerExistsByName = True : Exit Function
            Case 10
                Dim X10 As QBFC10Lib.ICustomerRet
                X10 = QB10_CustomerQuery_Name(CustomerName, RET)
                If RET = 0 Then QB_CustomerExistsByName = True : Exit Function
            Case 11
                Dim X11 As QBFC11Lib.ICustomerRet
                X11 = QB11_CustomerQuery_Name(CustomerName, RET)
                If RET = 0 Then QB_CustomerExistsByName = True : Exit Function
            Case 12
                Dim X12 As QBFC12Lib.ICustomerRet
                X12 = QB12_CustomerQuery_Name(CustomerName, RET)
                If RET = 0 Then QB_CustomerExistsByName = True : Exit Function
            Case 13
                Dim X13 As QBFC13Lib.ICustomerRet
                X13 = QB13_CustomerQuery_Name(CustomerName, RET)
                If RET = 0 Then QB_CustomerExistsByName = True : Exit Function
        End Select
Failed:
        If RET <> 0 Then ActiveLog("QB::QB_CustomerExistsByName(" & CustomerName & "): Error [" & RET & "] " & RetMsg, 4)
    End Function

    Public Function QB_AccountQuery_All() As Object
        Select Case QBFCVersion
            Case 5 : QB_AccountQuery_All = QB5_AccountQuery_All()
            Case 6 : QB_AccountQuery_All = QB6_AccountQuery_All()
            Case 7 : QB_AccountQuery_All = QB7_AccountQuery_All()
            Case 8 : QB_AccountQuery_All = QB8_AccountQuery_All()
            Case 10 : QB_AccountQuery_All = QB10_AccountQuery_All()
            Case 11 : QB_AccountQuery_All = QB11_AccountQuery_All()
            Case 12 : QB_AccountQuery_All = QB12_AccountQuery_All()
            Case 13 : QB_AccountQuery_All = QB13_AccountQuery_All()
        End Select
    End Function

    Public Function QB_VendorQuery_All(Optional ByRef RET As Integer = 0) As Object
        Select Case QBFCVersion
            Case 5 : QB_VendorQuery_All = QB5_VendorQuery_All(RET)
            Case 6 : QB_VendorQuery_All = QB6_VendorQuery_All(RET)
            Case 7 : QB_VendorQuery_All = QB7_VendorQuery_All(RET)
            Case 8 : QB_VendorQuery_All = qb8_VendorQuery_All(RET)
            Case 10 : QB_VendorQuery_All = QB10_VendorQuery_All(RET)
            Case 11 : QB_VendorQuery_All = QB11_VendorQuery_All(RET)
            Case 12 : QB_VendorQuery_All = QB12_VendorQuery_All(RET)
            Case 13 : QB_VendorQuery_All = QB13_VendorQuery_All(RET)
        End Select
    End Function

    Public Property QBFCVersion() As Integer
        Get
            If mQBFCVer <> 0 Then QBFCVersion = mQBFCVer
            QBFCVersion = GetQBSetupValue("qbfcver")
            If Not QBFCVersionSupported(QBFCVersion) Then QBFCVersion = QBFCVerDefault
            mQBFCVer = QBFCVersion
        End Get
        Set(value As Integer)
            QBShutdown()
            If value = 0 Then Exit Property
            If value < 0 Then value = QBFCVersionSuggestion
            If Not QBFCVersionSupported(value) Then value = QBFCVerDefault
            SetQBSetupValue("qbfcver", value)
            mQBFCVer = value
        End Set
    End Property

    Private ReadOnly Property QBFCVersionSuggestion() As Integer
        Get
            If QB13ObjectsExist() Then QBFCVersionSuggestion = 13 : Exit Property
            If QB12ObjectsExist() Then QBFCVersionSuggestion = 12 : Exit Property
            If QB11ObjectsExist() Then QBFCVersionSuggestion = 11 : Exit Property
            If QB10ObjectsExist() Then QBFCVersionSuggestion = 10 : Exit Property
            If QB8ObjectsExist() Then QBFCVersionSuggestion = 8 : Exit Property
            If QB7ObjectsExist() Then QBFCVersionSuggestion = 7 : Exit Property
            If QB6ObjectsExist() Then QBFCVersionSuggestion = 6 : Exit Property
            '  If QB5ObjectsExist Then QBFCVersionSuggestion = 5: Exit Property
            QBFCVersionSuggestion = QBFCVerDefault
        End Get
    End Property

    Public ReadOnly Property QBFCVersionSupported(ByVal VerNo As Integer) As Boolean
        Get
            QBFCVersionSupported = IsIn(VerNo, 6, 7, 8, 10, 11, 12, 13)
        End Get
    End Property

    Public Property QBActiveStore() As Integer
        Get
            QBActiveStore = mQBActiveStore
            If QBActiveStore = 0 Then QBActiveStore = StoresSld
        End Get
        Set(value As Integer)
            If value <> mQBActiveStore Then
                QBShutdown()
            End If

            mQBActiveStore = value
            mQBAS_JustSet = True
        End Set
    End Property

    Public ReadOnly Property QBVersionExist(ByVal VerNo As Integer, Optional FailReason As String = "") As Boolean
        Get
            On Error Resume Next
            Select Case VerNo
                Case 5 : QBVersionExist = QB5ObjectsExist(FailReason)
                Case 6 : QBVersionExist = QB6ObjectsExist(FailReason)
                Case 7 : QBVersionExist = QB7ObjectsExist(FailReason)
                Case 8 : QBVersionExist = QB8ObjectsExist(FailReason)
                Case 10 : QBVersionExist = QB10ObjectsExist(FailReason)
                Case 11 : QBVersionExist = QB11ObjectsExist(FailReason)
                Case 12 : QBVersionExist = QB12ObjectsExist(FailReason)
                Case 13 : QBVersionExist = QB13ObjectsExist(FailReason)
            End Select
        End Get
    End Property

    Public ReadOnly Property QB_Country() As String
        Get
            QB_Country = "US"
        End Get
    End Property

    Public ReadOnly Property QB_XML_MajorVer() As String
        Get
            QB_XML_MajorVer = GetQBSetupValue("xmlmajor")
        End Get
    End Property

    Public ReadOnly Property QB_XML_MinorVer() As String
        Get
            QB_XML_MinorVer = GetQBSetupValue("xmlminor")
        End Get
    End Property

    Public Function QB_SendRequests(Optional ByRef ErrString As String = "", Optional ByRef ErrNo As Integer = 0, Optional ByVal OnErr As ENRqOnError = ENRqOnError.roeContinue) As Boolean
        Select Case QBFCVersion
            Case 5 : QB_SendRequests = QB5_SendRequests(ErrString, ErrNo, OnErr)
            Case 6 : QB_SendRequests = QB6_SendRequests(ErrString, ErrNo, OnErr)
            Case 7 : QB_SendRequests = QB7_SendRequests(ErrString, ErrNo, OnErr)
            Case 8 : QB_SendRequests = QB8_SendRequests(ErrString, ErrNo, OnErr)
            Case 10 : QB_SendRequests = QB10_SendRequests(ErrString, ErrNo, OnErr)
            Case 11 : QB_SendRequests = QB11_SendRequests(ErrString, ErrNo, OnErr)
            Case 12 : QB_SendRequests = QB12_SendRequests(ErrString, ErrNo, OnErr)
            Case 13 : QB_SendRequests = QB13_SendRequests(ErrString, ErrNo, OnErr)
        End Select
    End Function

    Public Sub QBShutdown()
        Select Case QBFCVersion
            Case 5 : QB5Shutdown
            Case 6 : QB6Shutdown
            Case 7 : QB7Shutdown
            Case 8 : QB8Shutdown
            Case 10 : QB10Shutdown
            Case 11 : QB11Shutdown
            Case 12 : QB12Shutdown
            Case 13 : QB13Shutdown
        End Select
    End Sub

    Public ReadOnly Property QB_AppID() As String
        Get
            QB_AppID = ProgramName
        End Get
    End Property

    Public ReadOnly Property QB_AppNm() As String
        Get
            QB_AppNm = ProgramShort
        End Get
    End Property

End Module

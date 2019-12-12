Module modQuickBooks_CommQBFC11
    Public Function QB11_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As QBFC11Lib.IVendorRet
        Dim VQ As QBFC11Lib.IVendorQuery

        On Error Resume Next
  Set VQ = MSREQ11.AppendVendorQueryRq
  With VQ
            .ORVendorListQuery.FullNameList.Add Vendor
'    .ORVendorListQuery.VendorListFilter.ORNameFilter.NameFilter.Name.SetValue Vendor
        End With
  Set VQ = Nothing
  
  Set QB11_VendorQuery_Vendor = QB11_SendRequestsGetRet(RET, RetMsg)
  Exit Function
    End Function

    Public Function QB11_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC11Lib.IClassQuery
        QBSM11_Reset

        On Error Resume Next
  Set Q = MSREQ11.AppendClassQueryRq
  With Q
            .ORListQuery.FullNameList.Add ClassName
  End With
  Set Q = Nothing
  
  Set QB11_ClassQuery_Class = Nothing
  Set QB11_ClassQuery_Class = QB11_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB11_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String) As QBFC11Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC11Lib.ICustomerQuery
        QBSM11_Reset
  Set Q = MSREQ11.AppendCustomerQueryRq
  With Q
            Q.ORCustomerListQuery.FullNameList.Add Name
'    Q.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.Name.SetValue Name
        End With
  Set Q = Nothing

  Set QB11_CustomerQuery_Name = QB11_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB11_AccountQuery_All(Optional ByRef RET As Long, Optional ByRef RetMsg As String)
        On Error Resume Next
        QBSM11_Reset
        MSREQ11.AppendAccountQueryRq ' nothing to set
        QB11_AccountQuery_All = QB11_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB11_VendorQuery_All(Optional ByRef RET As Long)
        Dim ES As String
        On Error Resume Next
        QBSM11_Reset
        MSREQ11.AppendVendorQueryRq ' nothing to set
        QB11_VendorQuery_All = QB11_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB11ObjectsExist(Optional ByRef Msg As String) As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC11)"
  Set X = New QBFC11Lib.QBSessionManager
'  Set X = CreateObject("QBFC11.QBSessionManager")
  Msg = "Could not create MsgSetRequest (QBFC11)"
  Set Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
  Msg = ""
        QB11ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
  Set Y = Nothing
  Set X = Nothing
End Function

    Public Function QB11_SendRequests(Optional ByRef ErrString As String, Optional ByRef ErrNo As Long, Optional ByRef OnErr As ENRqOnError = roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB11Startup(ErrNo, ErrString) Then Exit Function
        MSREQ11.Attributes.OnError = OnErr
  Set MSRSP11 = QBSM11.DoRequests(MSREQ11)
'  Debug.Print mmsrsp11.ToXMLString
  
  QB11_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Public Function QB11Shutdown()
        ActiveLog "QB::QBShutdown", 1
  If QBSessOpen Then QBSM11.EndSession : QBSessOpen = False
        If QBConnOpen Then QBSM11.CloseConnection : QBConnOpen = False
  Set mQBSM11 = Nothing
End Function

End Module

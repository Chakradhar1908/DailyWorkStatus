Module modQuickBooks_CommQBFC6
    Public Function QB6_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As QBFC6Lib.IVendorRet
        Dim VQ As QBFC6Lib.IVendorQuery

        On Error Resume Next
  Set VQ = MSREQ6.AppendVendorQueryRq
  With VQ
            .ORVendorListQuery.FullNameList.Add Vendor
'    .ORVendorListQuery.VendorListFilter.ORNameFilter.NameFilter.Name.SetValue Vendor
        End With
  Set VQ = Nothing
  
  Set QB6_VendorQuery_Vendor = QB6_SendRequestsGetRet(RET, RetMsg)
  Exit Function
    End Function

    Public Function QB6_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String) As QBFC6Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC6Lib.ICustomerQuery
        QBSM6_Reset
  Set Q = MSREQ6.AppendCustomerQueryRq
  With Q
            Q.ORCustomerListQuery.FullNameList.Add Name
'    Q.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.Name.SetValue Name
        End With
  Set Q = Nothing

  Set QB6_CustomerQuery_Name = QB6_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB6_AccountQuery_All(Optional ByRef RET As Long, Optional ByRef RetMsg As String)
        On Error Resume Next
        QBSM6_Reset
        MSREQ6.AppendAccountQueryRq ' nothing to set
        QB6_AccountQuery_All = QB6_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB6_VendorQuery_All(Optional ByRef RET As Long)
        Dim ES As String
        On Error Resume Next
        QBSM6_Reset
        MSREQ6.AppendVendorQueryRq ' nothing to set
        QB6_VendorQuery_All = QB6_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB6ObjectsExist(Optional ByRef Msg As String) As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC6)"
  Set X = New QBFC6Lib.QBSessionManager
'  Set X = CreateObject("QBFC6.QBSessionManager")
  Msg = "Could not create MsgSetRequest (QBFC6)"
  Set Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
  Msg = ""
        QB6ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
  Set Y = Nothing
  Set X = Nothing
End Function

    Public Function QB6_SendRequests(Optional ByRef ErrString As String, Optional ByRef ErrNo As Long, Optional ByRef OnErr As ENRqOnError = roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB6Startup(ErrNo, ErrString) Then Exit Function
        MSREQ6.Attributes.OnError = OnErr
  Set MSRSP6 = QBSM6.DoRequests(MSREQ6)
'  Debug.Print mmsrsp6.ToXMLString
  
  QB6_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Public Function QB6Shutdown()
        ActiveLog "QB::QBShutdown", 1
  If QBSessOpen Then QBSM6.EndSession : QBSessOpen = False
        If QBConnOpen Then QBSM6.CloseConnection : QBConnOpen = False
  Set mQBSM6 = Nothing
End Function

End Module

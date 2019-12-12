Module modQuickBooks_CommQBFC7
    Public Function QB7_VendorQuery_Vendor(ByRef Vendor As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "") As QBFC7Lib.IVendorRet
        Dim VQ As QBFC7Lib.IVendorQuery

        On Error Resume Next
  Set VQ = MSREQ7.AppendVendorQueryRq
  With VQ
            .ORVendorListQuery.FullNameList.Add Vendor
'    .ORVendorListQuery.VendorListFilter.ORNameFilter.NameFilter.Name.SetValue Vendor
        End With
  Set VQ = Nothing
  
  Set QB7_VendorQuery_Vendor = QB7_SendRequestsGetRet(RET, RetMsg)
  Exit Function
    End Function

    Public Function QB7_ClassQuery_Class(ByRef ClassName As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String = "")
        Dim Q As QBFC7Lib.IClassQuery
        QBSM7_Reset

        On Error Resume Next
  Set Q = MSREQ7.AppendClassQueryRq
  With Q
            .ORListQuery.FullNameList.Add ClassName
  End With
  Set Q = Nothing
  
  Set QB7_ClassQuery_Class = Nothing
  Set QB7_ClassQuery_Class = QB7_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB7_CustomerQuery_Name(ByRef Name As String, Optional ByRef RET As Long = 0, Optional ByRef RetMsg As String) As QBFC7Lib.ICustomerRet
        On Error Resume Next
        Dim Q As QBFC7Lib.ICustomerQuery
        QBSM7_Reset
  Set Q = MSREQ7.AppendCustomerQueryRq
  With Q
            Q.ORCustomerListQuery.FullNameList.Add Name
'    Q.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.Name.SetValue Name
        End With
  Set Q = Nothing

  Set QB7_CustomerQuery_Name = QB7_SendRequestsGetRet(RET, RetMsg)
End Function

    Public Function QB7_AccountQuery_All(Optional ByRef RET As Long, Optional ByRef RetMsg As String)
        On Error Resume Next
        QBSM7_Reset
        MSREQ7.AppendAccountQueryRq ' nothing to set
        QB7_AccountQuery_All = QB7_SendRequestsGetRetList(RET, RetMsg)
    End Function

    Public Function QB7_VendorQuery_All(Optional ByRef RET As Long)
        Dim ES As String
        On Error Resume Next
        QBSM7_Reset
        MSREQ7.AppendVendorQueryRq ' nothing to set
        QB7_VendorQuery_All = QB7_SendRequestsGetRetList(RET, ES)
    End Function

    Public Function QB7ObjectsExist(Optional ByRef Msg As String) As Boolean
        Dim X, Y
        On Error GoTo FailPoint
        Msg = "Could not create session manager (QBFC7)"
  Set X = New QBFC7Lib.QBSessionManager
'  Set X = CreateObject("QBFC7.QBSessionManager")
  Msg = "Could not create MsgSetRequest (QBFC7)"
  Set Y = X.CreateMsgSetRequest(QB_Country, QB_XML_MajorVer, QB_XML_MinorVer)
  Msg = ""
        QB7ObjectsExist = True
FailPoint:
        If Msg <> "" Then Msg = Msg & ": " & Err.Description
  Set Y = Nothing
  Set X = Nothing
End Function

    Public Function QB7_SendRequests(Optional ByRef ErrString As String, Optional ByRef ErrNo As Long, Optional ByRef OnErr As ENRqOnError = roeContinue) As Boolean
        On Error GoTo NoComm
        If Not QB7Startup(ErrNo, ErrString) Then Exit Function
        MSREQ7.Attributes.OnError = OnErr
  Set MSRSP7 = QBSM7.DoRequests(MSREQ7)
'  Debug.Print mmsrsp7.ToXMLString
  
  QB7_SendRequests = True
        Exit Function ' no more auto-cleanup... persistant session

NoComm:
        ErrString = Err.Description
        ErrNo = Err.Number
        Err.Clear()
        QBShutdown()
    End Function

    Public Function QB7Shutdown()
        ActiveLog "QB::QBShutdown", 1
  If QBSessOpen Then QBSM7.EndSession : QBSessOpen = False
        If QBConnOpen Then QBSM7.CloseConnection : QBConnOpen = False
  Set mQBSM7 = Nothing
End Function

End Module

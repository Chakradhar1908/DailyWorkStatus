Imports System.Runtime.InteropServices
Module modShareFolder
    Private Const STYPE_DISKTREE As Integer = 0
    Private Const ACCESS_READ As Integer = &H1
    Private Const ACCESS_WRITE As Integer = &H2
    Private Const ACCESS_CREATE As Integer = &H4
    Private Const ACCESS_EXEC As Integer = &H8
    Private Const ACCESS_DELETE As Integer = &H10
    Private Const ACCESS_ATRIB As Integer = &H20
    Private Const ACCESS_PERM As Integer = &H40
    Private Const ACCESS_ALL As Integer = ACCESS_READ Or ACCESS_WRITE Or ACCESS_CREATE Or ACCESS_EXEC Or ACCESS_DELETE Or ACCESS_ATRIB Or ACCESS_PERM
    Private Declare Function NetShareAdd Lib "netapi32" (ByVal ServerName As IntPtr, ByVal level As Integer, buf As Object, ParmErr As Integer) As Integer
    Private Const NERR_Success As Integer = 0&
    Private Const NERR_Base As Integer = 2100
    Private Const NERR_DuplicateShare As Integer = NERR_Base + 18
    'Private Declare Function NetShareCheck Lib "netapi32" (ByVal ServerName as integer, ByRef DeviceName as integer, ByRef pType as integer) as integer --vb6.0
    Private Declare Function NetShareCheck Lib "netapi32" (ByVal ServerName As IntPtr, ByRef DeviceName As IntPtr, ByRef pType As Integer) As Integer  'vb.net
    Private Const NERR_DeviceNotShared As Integer = NERR_Base + 211

    Private Structure SHARE_INFO_2
        'Dim shi2_netname as integer
        Dim shi2_netname As IntPtr

        Dim shi2_type As Integer
        'Dim shi2_remark as integer
        Dim shi2_remark As IntPtr
        Dim shi2_permissions As Integer
        Dim shi2_max_uses As Integer
        Dim shi2_current_uses As Integer
        'Dim shi2_path as integer
        Dim shi2_path As IntPtr
        'Dim shi2_passwd as integer
        Dim shi2_passwd As IntPtr
    End Structure

    Public Function CreateCDSDataShare() As Boolean
        On Error Resume Next
        If ShareCheck(UCase(GetLocalComputerName), "C:\CDSData") Then Exit Function
        CreateCDSDataShare = ShareAdd(UCase(GetLocalComputerName), "C:\CDSData", "CDSData")
    End Function

    Public Function ShareCheck(ByVal sServerName As String, ByVal sDevice As String) As Boolean
        Dim Res As Integer
        Dim pType As Integer
        Dim dwServer As IntPtr, dwDevice As IntPtr

        sDevice = UCase(sDevice)


        'Res = NetShareCheck(StrPtr(sServerName), StrPtr(sDevice), pType)
        dwServer = Marshal.UnsafeAddrOfPinnedArrayElement(sServerName.ToArray, 0)
        dwDevice = Marshal.UnsafeAddrOfPinnedArrayElement(sDevice.ToArray, 0)

        Res = NetShareCheck(dwServer, dwDevice, pType)

        Select Case Res
            Case NERR_Success : ShareCheck = True
            Case NERR_DeviceNotShared : ShareCheck = False
            Case Else
                Debug.Print("ShareCheckRemote: ERROR: ** " & WinSysError(Res))
                ShareCheck = False
        End Select
    End Function

    Public Function ShareAdd(ByVal sServer As String, ByVal sSharePath As String, ByVal sShareName As String, Optional ByVal sShareRemark As String = "", Optional ByVal sSharePw As String = "") As Boolean
        'Dim dwServer as integer
        Dim dwServer As IntPtr

        'Dim dwNetname as integer
        Dim dwNetname As IntPtr

        'Dim dwPath as integer
        Dim dwPath As IntPtr

        'Dim dwRemark as integer
        Dim dwRemark As IntPtr

        'Dim dwPw as integer
        Dim dwPw As IntPtr

        Dim ParmErr As Integer
        Dim SI2 As SHARE_INFO_2
        Dim Res As Integer

        'obtain pointers to the server, share and path
        'dwServer = StrPtr(sServer)

        dwServer = Marshal.UnsafeAddrOfPinnedArrayElement(sServer.ToArray, 0)

        'dwNetname = StrPtr(sShareName)
        dwNetname = Marshal.UnsafeAddrOfPinnedArrayElement(sShareName.ToArray, 0)

        'dwPath = StrPtr(sSharePath)
        dwPath = Marshal.UnsafeAddrOfPinnedArrayElement(sSharePath.ToArray, 0)

        'if the remark or password specified obtain pointer to those as well
        'If Len(sShareRemark) > 0 Then dwRemark = StrPtr(sShareRemark)
        If Len(sShareRemark) > 0 Then dwRemark = Marshal.UnsafeAddrOfPinnedArrayElement(sShareRemark.ToArray, 0)

        'If Len(sSharePw) > 0 Then dwPw = StrPtr(sSharePw)
        If Len(sSharePw) > 0 Then dwPw = Marshal.UnsafeAddrOfPinnedArrayElement(sSharePw.ToArray, 0)

        'prepare the SHARE_INFO_2 structure
        With SI2
            .shi2_netname = dwNetname
            .shi2_path = dwPath
            .shi2_remark = dwRemark
            .shi2_type = STYPE_DISKTREE
            .shi2_permissions = ACCESS_ALL
            .shi2_max_uses = -1
            .shi2_passwd = dwPw
        End With

        'add the share
        Res = NetShareAdd(dwServer, 2, SI2, ParmErr)

        '  MsgBox WinSysError(Res)
        ShareAdd = (Res = NERR_Success Or Res = NERR_DuplicateShare)
    End Function

End Module

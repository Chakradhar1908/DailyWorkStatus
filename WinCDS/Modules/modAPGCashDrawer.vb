Imports APG554QFAX21Lib
Module modAPGCashDrawer
    Private Declare Function APGOpenDrawer Lib "554QFDLL11.dll" Alias "OpenDrawer" () As Integer
    Private Declare Function APGDrawerStatus Lib "554QFDLL11.dll" Alias "DrawerStatus" () As Integer

    Public Function OpenAPGCashDrawer() As Boolean
        '::::OpenAPGCashDrawer
        ':::SUMMARY
        ': Opens the APG Cash Drawer
        ':::DESCRIPTION
        ': Calls the DLL to open the drawer
        ':::PARAMETERS
        ':::RETURN
        ': Boolean - Returns True.
        Dim C As APG554QFAX21
        On Error Resume Next
        If Not HasAPGDLL() Then
            C = APGControl()
            C.OpenDrawer()
            APGControlDestroy(C)
        Else
            APGOpenDrawer
        End If
        OpenAPGCashDrawer = True
    End Function

    Public Function HasAPGDLL() As Boolean
        '::::HasAPGDLL
        ':::SUMMARY
        ': System has the APG DLL
        ':::DESCRIPTION
        ': Checks existence of APG Cash Drawer DLL.  Makes remaining functions safe.
        ':::PARAMETERS
        ':::RETURN
        ': Boolean

        HasAPGDLL = FileExists(AppFolder() & "554QFDLL11.dll")
    End Function

    Private Function APGControl() As APG554QFAX21
        On Error Resume Next
        Dim c As New APG554QFAX21


        'APGControl = MainMenu.Controls.Add("APG554QFAX21Ctrl.1", "apg_" + Timestamp(, True) & "_" & GetTickCount())
        MainMenu.Controls.Add(c)
        APGControl = c
    End Function

    Private Function APGControlDestroy(ByVal C As APG554QFAX21)
        On Error Resume Next
        MainMenu.Controls.Remove(C)
    End Function
End Module

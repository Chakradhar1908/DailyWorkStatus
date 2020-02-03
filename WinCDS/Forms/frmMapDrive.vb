Public Class frmMapDrive
    Public Function InitialSetup() As Boolean
        Dim S As String, R As VbMsgBoxResult
        Dim C As String

        ' Be careful.  Don't use anything!  Post-installer setup

        S = "SOFTWARE\" & CompanyName & "\" & ProgramName
        S = QueryValue(HKLM, S, "IsServer")
        Select Case S
            Case "server"
                CreateCDSDataShare
            Case "station"
                If Not IDriveIsMapped Then
                    txtServerName = "INVENTORY"
                    txtShareName = "CDSData"
                    txtUserName = ""
                    txtPassword = ""

Again:
                    Canceled = False
                    Show vbModal

        If Not Canceled Then
                        C = "net.exe"
                        C = C & " use I:"
                        C = C & " \\" & txtServerName & "\" & txtShareName
                        C = C & IIf(txtUserName = "", "", " /USER:" & txtUserName)
                        C = C & IIf(txtPassword = "", "", " " & txtPassword)
                        C = C & " /PERSISTENT:YES"

                        MousePointer = vbHourglass
                        ShellOut.ShellAndWait C, enSW_MINIMIZE
'          Shell C ' Here we use the VB shell method, not our own.  "Dont use anything"
                        Sleep 5000
          MousePointer = vbDefault

                        If Not IDriveIsMapped Then
                            If MsgBox("Could not map I:\ drive." & vbCrLf2 & "Would you like to try again?", vbOKCancel + vbExclamation, "Map Drive Failed") = vbOK Then
                                GoTo Again
                            End If
                            MsgBox "You will have to map the I:\ drive manually in order to use the software in workstation mode.", vbInformation, "Map Drive Failed"
            End
                        End If
                    Else
                        MsgBox "You have selected to install WinCDS as a workstation, but have not set up an I:\ mapping." & vbCrLf2 & "You will have to set up your drive mapping manually in order for WinCDS to start properly." & vbCrLf2 & "You must map your I:\ to your server manually before running WinCDS.", vbExclamation, ProgramName & " Drive Mapping Cancelled"
          End
                    End If
                End If
        End Select

        Unload Me
  InitialSetup = True
    End Function

End Class
Public Class frmMapDrive
    Private Canceled As Boolean
    Public Function IDriveIsMapped() As Boolean
        IDriveIsMapped = Dir("I:\Invent\CDSInvent.mdb") <> ""
    End Function

    Public Function InitialSetup() As Boolean
        Dim S As String, R As VBA.VbMsgBoxResult
        Dim C As String

        ' Be careful.  Don't use anything!  Post-installer setup

        S = "SOFTWARE\" & CompanyName & "\" & ProgramName
        S = QueryValue(HKLM, S, "IsServer")
        Select Case S
            Case "server"
                CreateCDSDataShare()
            Case "station"
                If Not IDriveIsMapped() Then
                    txtServerName.Text = "INVENTORY"
                    txtShareName.Text = "CDSData"
                    txtUserName.Text = ""
                    txtPassword.Text = ""

Again:
                    Canceled = False
                    ShowDialog()

                    If Not Canceled Then
                        C = "net.exe"
                        C = C & " use I:"
                        C = C & " \\" & txtServerName.Text & "\" & txtShareName.Text
                        C = C & IIf(txtUserName.Text = "", "", " /USER:" & txtUserName.Text)
                        C = C & IIf(txtPassword.Text = "", "", " " & txtPassword.Text)
                        C = C & " /PERSISTENT:YES"

                        'MousePointer = vbHourglass
                        Me.Cursor = Cursors.WaitCursor
                        ShellOut.ShellAndWait(C, EnSW.enSW_MINIMIZE)
                        '          Shell C ' Here we use the VB shell method, not our own.  "Dont use anything"
                        Sleep(5000)
                        'MousePointer = vbDefault
                        Me.Cursor = Cursors.Default

                        If Not IDriveIsMapped() Then
                            If MsgBox("Could not map I:\ drive." & vbCrLf2 & "Would you like to try again?", vbOKCancel + vbExclamation, "Map Drive Failed") = vbOK Then
                                GoTo Again
                            End If
                            MessageBox.Show("You will have to map the I:\ drive manually in order to use the software in workstation mode.", "Map Drive Failed")
                            End
                        End If
                    Else
                        MessageBox.Show("You have selected to install WinCDS as a workstation, but have not set up an I:\ mapping." & vbCrLf2 & "You will have to set up your drive mapping manually in order for WinCDS to start properly." & vbCrLf2 & "You must map your I:\ to your server manually before running WinCDS.", ProgramName & " Drive Mapping Cancelled")
                        End
                    End If
                End If
        End Select

        'Unload Me
        Me.Close()
        InitialSetup = True
    End Function

End Class
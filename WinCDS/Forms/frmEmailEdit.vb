Public Class frmEmailEdit
    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Select Case Inven
            Case "AK-E"
                If Trim(RTBEditTemplate.Text) = "" Then
                    If MessageBox.Show("Do you want to reset to the default message?", "Edit Template - WinCDS", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        DeleteFileIfExists(EmailTemplateFactOrdNotAckFile)
                    Else
                        Exit Sub
                    End If
                Else
                    RTBEditTemplate.SaveFile(EmailTemplateFactOrdNotAckFile)
                    MessageBox.Show("Template saved.", "Edit Template - WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    'Unload Me
                    Me.Close()
                End If
            Case "OverdueOrders-E"
                If Trim(RTBEditTemplate.Text) = "" Then
                    If MessageBox.Show("Do you want to reset to the default message?", "Edit Template - WinCDS", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        DeleteFileIfExists(EmailTemplateOverdueOrdersFile)
                    Else
                        Exit Sub
                    End If
                Else
                    RTBEditTemplate.SaveFile(EmailTemplateOverdueOrdersFile)
                    MessageBox.Show("Template saved.", "Edit Template - WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    'Unload Me
                    Me.Close()
                End If
        End Select

        Select Case Order
            Case "SParts"
                If Trim(RTBEditTemplate.Text) = "" Then
                    If MessageBox.Show("Do you want to reset to the default message?", "Edit Template - WinCDS", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        DeleteFileIfExists(EmailTemplateChargeBackFile)
                    Else
                        Exit Sub
                    End If
                Else
                    RTBEditTemplate.SaveFile(EmailTemplateChargeBackFile)
                    MessageBox.Show("Template saved.", "Edit Template - WinCDS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    'Unload Me
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub frmEmailEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdCancel, 3)
        LoadTemplate
    End Sub

    Public Sub LoadTemplate()
        Select Case Inven
            Case "AK-E"
                If FileExists(EmailTemplateFactOrdNotAckFile) Then
                    RTBEditTemplate.LoadFile(EmailTemplateFactOrdNotAckFile)
                Else
                    RTBEditTemplate.Text = DefaultEmailFactOrdNotAck(True)
                End If
            Case "OverdueOrders-E"
                If FileExists(EmailTemplateOverdueOrdersFile) Then
                    RTBEditTemplate.LoadFile(EmailTemplateOverdueOrdersFile)
                Else
                    RTBEditTemplate.Text = DefaultEmailOverdueOrders(True)
                End If
        End Select

        Select Case Order
            Case "SParts"
                If FileExists(EmailTemplateChargeBackFile) Then
                    RTBEditTemplate.LoadFile(EmailTemplateChargeBackFile)
                Else
                    RTBEditTemplate.Text = DefaultEmailChargeBack(True)
                End If
        End Select
    End Sub

End Class
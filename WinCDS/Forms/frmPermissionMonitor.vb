Public Class frmPermissionMonitor
    Public Sub ShowLog(Optional ByVal ClearIt As Boolean = False)
        If ClearIt Then ActiveLogClear()
        ActiveLogLoadClasses(cmbLogType)
        txtLog = ActiveLogLines(cmbLogType.Text, Val(txtLogLvl), chkLogTS.Checked = 1)
    End Sub

End Class
Imports Microsoft.VisualBasic.Interaction
Public Class frmPermissionMonitor
    Private SettingsLoaded As Boolean
    Public Sub ShowLog(Optional ByVal ClearIt As Boolean = False)
        If ClearIt Then ActiveLogClear()
        ActiveLogLoadClasses(cmbLogType)
        txtLog = ActiveLogLines(cmbLogType.Text, Val(txtLogLvl), chkLogTS.Checked = 1)
    End Sub

    Public Property OnTop() As Boolean
        Get
            OnTop = Not lblTop.Visible
        End Get
        Set(value As Boolean)
            SetAlwaysOnTop(Me, value)
            lblBot.Visible = Not value
            lblTop.Visible = value
        End Set
    End Property

    Public Property Display() As Integer
        Get
            Display = Switch(optPerm, 1, optStatus, 2, optLog, 3, optExtra, 4, optFormList, 5, optMemory, 6, True, 1)
        End Get
        Set(value As Integer)
            tmrExtra.Enabled = False
            tmrFormList.Enabled = False
            tmrMemory.Enabled = False
            Select Case value
                Case 2
                    'If Not optStatus Then optStatus = True
                    If optStatus.Checked = False Then optStatus.Checked = True
                    'fraStatus.ZOrder 0
                    'fraStatus.ZOrder 0
                    fraStatus.BringToFront()
                Case 3
                    'If Not optLog Then optLog = True
                    If optLog.Checked = False Then optLog.Checked = True
                    'fraLog.ZOrder 0
                    fraLog.BringToFront()
                    ShowLog()
                Case 4
                    'If Not optExtra Then optExtra = True
                    If optExtra.Checked = False Then optExtra.Checked = True
                    LoadExtras()
                    'fraExtras.ZOrder 0
                    fraExtras.BringToFront()
                    tmrExtra.Interval = 1000
                    tmrExtra.Enabled = True
                Case 5
                    'If Not optFormList Then optFormList = True
                    If optFormList.Checked = False Then optFormList.Checked = True
                    LoadForms()
                    'fraFormList.ZOrder 0
                    fraFormList.BringToFront()
                    tmrFormList.Interval = 1000
                    tmrFormList.Enabled = True
                Case 6
                    'If Not optMemory Then optMemory = True
                    If optMemory.Checked = False Then optMemory.Checked = True
                    LoadMemoryInfo()
                    'fraMemory.ZOrder 0
                    fraMemory.BringToFront()
                    tmrMemory.Interval = 1000
                    tmrMemory.Enabled = True
                Case Else
                    'If Not optPerm Then optPerm = True
                    If optPerm.Checked = False Then optPerm.Checked = True
                    'fraPerm.ZOrder 0
                    fraPerm.BringToFront()
                    tmrPerm.Interval = 400
                    tmrPerm.Enabled = True
            End Select
        End Set
    End Property

    Private Sub LoadMemoryInfo()
        Dim S As String, M As String, N As String
        On Error Resume Next

        S = "" : N = vbCrLf
        S = S & M
        S = S & M & "Total Mem.: " & Mid(FormatCurrency(BenchmarkMemInfo, 0), 2)
        'S = S & N & "Form Count: " & Forms.Count
        S = S & N & "Form Count: " & My.Application.OpenForms.Count
        S = S & N
        S = S & N & BenchmarkMemoryProfile()

        txtMemory.Text = S
    End Sub

    Private Sub LoadForms()
        Dim I As Integer, S As String, OIX As Integer, ONM As String

        On Error Resume Next
        OIX = lstForms.SelectedIndex
        'ONM = lstForms.List(OIX)
        ONM = lstForms.SelectedItem.ToString

        fraFormList.Text = "Form List (" & My.Application.OpenForms.Count & "):"
        'LockWindowUpdate hwnd
        LockWindowUpdate(Handle)
        lstForms.Items.Clear()

        For I = 0 To My.Application.OpenForms.Count - 1
            'S = Format(I, "#0") & ": " & Forms(I).Name
            S = Format(I, "#0") & ": " & My.Application.OpenForms(I).Name
            If Not (fActiveForm() Is Nothing) Then
                If fActiveForm() Is My.Application.OpenForms(I) Then
                    '      If fActiveForm.Name = Forms(I).Name Then
                    S = S & "**"
                    'If Not (Forms(I).ActiveControl Is Nothing) Then
                    If Not (My.Application.OpenForms(I).ActiveControl Is Nothing) Then
                        S = S & "  -->" & My.Application.OpenForms(I).ActiveControl.Name
                    End If
                End If
            End If
            'lstForms.AddItem S
            lstForms.Items.Add(S)
        Next
        'LockWindowUpdate 0
        LockWindowUpdate(IntPtr.Zero)

        'If lstForms.List(OIX) = ONM Then lstForms.ListIndex = OIX
        If lstForms.SelectedItem.ToString = ONM Then lstForms.SelectedIndex = OIX
    End Sub

    Private Sub LoadExtras()
        Dim F As Form, T As String
        F = fActiveForm()
        If F Is Nothing Then Exit Sub
        On Error Resume Next
        T = "F: " & F.Name & vbCrLf
        'T = T & F.DeveloperEx
        txtExtras.Text = T
    End Sub

    Public Sub LoadSettings()
        Dim S As String, X As Object
        If SettingsLoaded Then Exit Sub
        SettingsLoaded = True

        S = GetCDSSetting("Permission Monitor")
        If S = "" Then Exit Sub
        X = Val(CSVField(S, 1))
        If X > 0 Then Width = X
        X = Val(CSVField(S, 2))
        If X > 0 Then Height = X
        X = Val(CSVField(S, 3))
        If X > 0 Then Left = X
        X = Val(CSVField(S, 4))
        If X > 0 Then Top = X

        OnTop = Val(CSVField(S, 5)) <> 0
        Display = Val(CSVField(S, 6))
        cmbLogType.Text = CSVField(S, 7)

        X = Val(CSVField(S, 8, "200"))
        X = FitRange(50, X, 9999)
        txtMaxLogLines = X
    End Sub

End Class
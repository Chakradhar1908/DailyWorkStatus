Module modActiveLog
    Private mActiveLogLines As clsHashTable, mActiveLogClasses As clsHashTable
    Public Sub ActiveLog(ByVal Msg As String, Optional ByVal Priority as integer = 9) ', Optional ByVal ToFile As String = "")
        '::::ActiveLog
        ':::SUMMARY
        ': Enter an ActiveLog entry
        ':::DESCRIPTION
        ': This function is used to add Active Log lines.
        ':::PARAMETERS
        ': - Msg
        ': - Priority
        Dim N as integer, Clss As String
        'MsgBox "ActiveLog"
        N = InStr(Msg, "::")
        If N <= 0 Then
            Clss = "General"
        Else
            Clss = Left(Msg, N - 1)
            Msg = Mid(Msg, N + 2)
        End If
        ActiveLogAdd(Msg, Clss, Priority)
        If IsFormLoaded("frmPermissionMonitor") Then frmPermissionMonitor.ShowLog()
    End Sub

    Private Function ActiveLogAdd(ByVal Msg As String, Optional ByVal ClassX As String = "General", Optional ByVal Priority As Integer = 9) As Boolean
        'Dim E
        Dim E As Object

        'MsgBox "ActiveLogAdd"
        If mActiveLogLines Is Nothing Then mActiveLogLines = New clsHashTable
        ActiveLogLimit()
        ActiveLogClassNumber(ClassX)
        'E = Array(Replace(Msg, vbCrLf, ""), ClassX, FitRange(1, Priority, 9), Now)
        E = {Replace(Msg, vbCrLf, ""), ClassX, FitRange(1, Priority, 9), Now}
        mActiveLogLines.Add("", E)

        ' of course, if we do this here, we will fail in other places.
        ' Logfile is not a safe function, because it relies on LogFolder to
        ' know where to write.  This cannot be deduced except by IsServer(),
        ' which calls MessageBox if it can't be determined, which in turn
        ' calls ActiveLog...  Something in that change would need to be
        ' changed to make this work, and that have to be for another day..

        '  If FileExists(InventFolder() & "activelog.txt") Then
        '    LOGFILE "ActiveLog", Priority & " " & Class & " - " & Msg
        '  End If
    End Function
    Private Function ActiveLogLimit()
        Dim N as integer, I as integer, X as integer, S As String
        On Error Resume Next
        S = GetCDSSetting("Permission Monitor")
        X = Val(CSVField(S, 8, "200"))
        X = FitRange(50, X, 9999)

        If mActiveLogLines Is Nothing Then Exit Function
        If mActiveLogLines.Count < 200 Then Exit Function
        For I = 1 To 10
            N = MinArray(mActiveLogLines.Keys)
            mActiveLogLines.Remove(N)
        Next
    End Function
    Private Function ActiveLogClassNumber(ByVal ClassName As String, Optional ByVal CreateIt As Boolean = True) as integer

        Dim I as integer
        If mActiveLogClasses Is Nothing Then
            mActiveLogClasses = New clsHashTable
            mActiveLogClasses.Add(0, "General")
        End If
        For I = 0 To mActiveLogClasses.Count - 1
            If LCase(ClassName) = LCase(mActiveLogClasses.Item(I)) Then ActiveLogClassNumber = I : Exit Function
        Next
        If CreateIt Then
            ActiveLogClassNumber = mActiveLogClasses.Count
            mActiveLogClasses.Add(ActiveLogClassNumber, ClassName)
        Else
            ActiveLogClassNumber = -1
        End If
    End Function
    Public Sub ActiveLogClear()
        '::::ActiveLogClear
        ':::SUMMARY
        ': Clear Active Log
        ':::DESCRIPTION
        ': This function is used to clear all Active Log Lines and Active Log Classes.
        mActiveLogLines = Nothing
        mActiveLogClasses = Nothing
    End Sub
    Public Sub ActiveLogLoadClasses(ByRef Cmb As ComboBox)
        '::::ActiveLogLoadClasses
        ':::SUMMARY
        ': Load Active Log Classes to Combo Box
        ':::DESCRIPTION
        ': This funcion is used to load Active Log Classes and used to display using ComboBox.
        ':::PARAMETERS
        ': - Cmb
        Dim X As String, I as integer
        X = Cmb.Text
        If mActiveLogClasses Is Nothing Then mActiveLogClasses = New clsHashTable
        'If Cmb.ListCount = mActiveLogClasses.Count + 1 Then Exit Sub
        If Cmb.Items.Count = mActiveLogClasses.Count + 1 Then Exit Sub
        'Cmb.Clear
        Cmb.Items.Clear()
        'Cmb.AddItem "All"
        'Cmb.itemData(Cmb.NewIndex) = -1

        '-->  Commented above two lines (cmb.additem and cmb.itemdata) and replaced with the below line. 
        '-->  Because, itemdata property does not exist in vb.net. For this, one custom class ItemDataClass created.
        Cmb.Items.Add(New ItemDataClass("All", -1))

        For I = 0 To mActiveLogClasses.Count - 1
            'Cmb.AddItem mActiveLogClasses.Item(I)
            'Cmb.itemData(Cmb.NewIndex) = I

            '-->  Commented above two lines (cmb.additem and cmb.itemdata) and replaced with the below line. 
            '-->  Because, itemdata property does not exist in vb.net. For this, one custom class ItemDataClass created.
            Cmb.Items.Add(New ItemDataClass(mActiveLogClasses.Item(I), I))
        Next
        On Error Resume Next
        Cmb.Text = X
    End Sub
    Public Function ActiveLogLines(Optional ByVal Classx As String = "All", Optional ByVal MaxPriority as integer = 9, Optional ByVal Timestamp As Boolean = False)
        '::::ActiveLogLines
        ':::SUMMARY
        ': Load Output Lines
        ':::DESCRIPTION
        ': This function is used to fetch the log file lines afer formatting them based on class, priority.
        ':::PARAMETERS
        ': - Class
        ': - MaxPriority
        ': - Timestamp
        ':::RETURN
        ': - Variant of String()
        Dim N as integer, L As String, E
        Dim M As String, C As String, P as integer, T As Date
        On Error Resume Next
        If mActiveLogLines Is Nothing Then mActiveLogLines = New clsHashTable
        N = MinArray(mActiveLogLines.Keys)
        Do While mActiveLogLines.Exists(N)
            E = mActiveLogLines.Item(N)
            M = E(0)
            C = E(1)
            P = Val(E(2))
            T = E(3)
            If P <= MaxPriority And (Classx = "All" Or Classx = C) Then
                L = L & IIf(Timestamp, Format(T, "hh:mm:ss") & ": ", "")
                L = L & M
                L = L & vbCrLf
            End If
            N = N + 1
        Loop
        ActiveLogLines = L
    End Function

End Module

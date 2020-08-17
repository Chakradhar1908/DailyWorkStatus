Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Module Functions
    Public Const vbCrLf2 As String = vbCrLf & vbCrLf
    Public Function DescribeColor(ByVal RGB as integer) As String
        Dim R as integer, G as integer, B as integer
        Select Case RGB
            Case vbBlack : DescribeColor = "BLACK"
            Case vbBlue : DescribeColor = "BLUE"
            Case vbCyan : DescribeColor = "CYAN"
            Case vbGreen : DescribeColor = "GREEN"
            Case vbMagenta : DescribeColor = "MAGENTA"
            Case vbRed : DescribeColor = "RED"
            Case vbWhite : DescribeColor = "WHITE"
            Case vbYellow : DescribeColor = "YELLOW"

            Case Else
                DescribeColor = "OTHER"
                R = RGB And 255
                G = (RGB And 65280) / 256
                B = (RGB And 16711680) / 65536
                DescribeColor = DescribeColor & " (R:" & R & ",G:" & G & ",:" & B & ")"
        End Select
    End Function
    Public Function DisposeDA(ByRef X As Object, Optional ByRef X2 As Object = Nothing, Optional ByRef X3 As Object = Nothing, Optional ByRef X4 As Object = Nothing, Optional ByRef X5 As Object = Nothing, Optional ByRef X6 As Object = Nothing, Optional ByRef X7 As Object = Nothing, Optional ByRef X8 As Object = Nothing, Optional ByRef X9 As Object = Nothing, Optional ByRef X0 As Object = Nothing)
        On Error Resume Next
        X.Close
        X.Dispose
        X.DataAccess.Dispose
        X.dbClose
        X.DataSource.Close
        X = Nothing

        If Not X2 Is Nothing Then DisposeDA(X2)
        If Not X3 Is Nothing Then DisposeDA(X3)
        If Not X4 Is Nothing Then DisposeDA(X4)
        If Not X5 Is Nothing Then DisposeDA(X5)
        If Not X6 Is Nothing Then DisposeDA(X6)
        If Not X7 Is Nothing Then DisposeDA(X7)
        If Not X8 Is Nothing Then DisposeDA(X8)
        If Not X9 Is Nothing Then DisposeDA(X9)
        If Not X0 Is Nothing Then DisposeDA(X0)

        Err.Clear()
    End Function
    Public Function SpeechActive() As Boolean
        SpeechActive = IsFormLoaded("frmSpeech")
    End Function
    Public Function GetFileAutonumber(ByVal fName As String, ByVal Defaultx as integer)
        On Error GoTo BadFile
        Dim FileVal As String, FNum as integer

        If fName = "" Then
            MsgBox("No Autonumber filename.", vbCritical, "Error")
            Exit Function
        End If

        If Mid(fName, 2, 1) <> ":" Then
            fName = InventFolder() & fName
        End If

        FNum = FreeFile()
        'Open fName For Input As #FNum
        FileOpen(FNum, fName, OpenMode.Input)
        'Line Input #FNum, FileVal)
        FileVal = LineInput(FNum)
        'Close #FNum
        FileClose(FNum)
        FNum = 0

        If IsNumeric(FileVal) Then
            GetFileAutonumber = Val(FileVal) + 1
        Else
            '    MsgBox "Invalid value " & FileVal & " received from " & FName & ", resetting to " & Defaultx & ".", vbCritical
            GetFileAutonumber = Defaultx
        End If

        FNum = FreeFile()
        'Open fName For Output As #FNum
        FileOpen(FNum, fName, OpenMode.Output)
        Print(FNum, CStr(GetFileAutonumber))
        'Close #FNum
        FileClose(FNum)
        Application.DoEvents()

        Exit Function
BadFile:
        Select Case Err.Number
            Case 52, 53, 62 ' File not found, bad file name or number, input past end of file
                ' These errors mean the file didn't exist, but can be created.
                Resume Next
            Case 70, 75 'Permission denied., Path/File access error
                '      If MsgBox(Err.Description & vbCrLf & "Can't access " & FName & ".  Try again?", vbCritical + vbYesNo, "BOS Num File Error") = vbYes Then Resume Else End
                Application.DoEvents()
                Resume
            Case 76 ' Path not found
                'If MsgBox("Can't access " & fName & ", try again?", vbCritical + vbYesNo, "File Error") = vbYes Then
                If MessageBox.Show("Can't access " & fName & ", try again?", "File Error", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                    Resume
                Else
                    End
                End If
            Case Else  ' An error we didn't foresee.
                'MsgBox("An unforseen error [" & Err.Number & "] occurred accessing " & fName & "." & vbCrLf & "Your sale could not be completed." & vbCrLf & Err.Description, vbCritical, "BOS Number Error")
                MessageBox.Show("An unforseen error [" & Err.Number & "] occurred accessing " & fName & "." & vbCrLf & "Your sale could not be completed." & vbCrLf & Err.Description, "BOS Number Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End
        End Select
    End Function

    Public Sub ChooseSet(ByVal Idx, ByVal Value, ParamArray List())
        On Error Resume Next
        List(Idx - 1) = Value
        'List(Idx) = Value
    End Sub

    Public Function OpenCashDrawer(Optional ByVal PortNum As Integer = -1, Optional ByVal ForceUSB7 As TriState = vbUseDefault) As Boolean
        Dim USB7 As Boolean

        USB7 = QueryStringQueryL(StoreSettings.CashDrawerConfig, "USB7") <> 0
        Select Case ForceUSB7
            Case vbTrue : USB7 = True
            Case vbFalse : USB7 = False
        End Select

        'Cash Drawer
        On Error GoTo HandleErr
        If PortNum = -1 Then PortNum = CashDrawerCOMPort

        If PortNum = 253 Then         ' 253 is the number we set aside to indicate USB
            OpenAPGCashDrawer()
            '    Load frmOpenCashDrawer      ' form load fires the activex control
            '    Unload frmOpenCashDrawer    ' simply loading it should open the drawer
            OpenCashDrawer = True
        ElseIf PortNum <> 0 Then  'Make sure Cash Drawer is Enabled
            Dim MSComm1 As MSCommLib.MSComm
            MSComm1 = MainMenu.MSComm1
            '    Set MSComm1 = CreateObject("MSCommlib.MSComm")  ' It'd be great if we could test this. :)
            ' Remove the control from MainMenu if we ever figure out how to load one here.
            MSComm1.CommPort = PortNum 'Choose COM port
            MSComm1.Settings = "9600,N,8,1" 'Set default settings
            MSComm1.PortOpen = True 'Open the port
            If USB7 Then
                MSComm1.Output = Chr(7) ' "7"
            Else
                MSComm1.Output = "AAAAAAAAAA" 'Write to the drawer
            End If
            MSComm1.PortOpen = False 'Close the port
            OpenCashDrawer = True
        End If
        Exit Function
HandleErr:
        ' This shouldn't be a fatal error, but we have to the user know we tried.
        OpenCashDrawer = False
        MessageBox.Show("Error " & Err.Number & " opening cash drawer: " & Err.Description)
        Err.Clear()
    End Function

    Public Function InitLineBorderForm(ByVal LineControlArray As Object, ByVal vFrm As Form, Optional ByVal BorderWidth as integer = 2) As Boolean
        InitLineBorderForm = InitLineBorder(LineControlArray, 0, 0, vFrm.Width, vFrm.Height, BorderWidth)
    End Function

    Public Function InitLineBorder(ByVal LineControlArray As Object, ByVal L as integer, ByVal T as integer, ByVal W as integer, ByVal H as integer, Optional ByVal BorderWidth As Integer = 2) As Boolean
        Dim I as integer, lW as integer, BC

        On Error Resume Next
        If False Then
            lW = 15
            'BC = Array(vbRed, vbBlue, vbGreen, vbCyan)
            BC = New String() {vbRed, vbBlue, vbGreen, vbCyan}
        Else
            'lW = LineControlArray(0).BorderWidth * Screen.TwipsPerPixelX  'lin(0).BorderWidth
            lW = LineControlArray(0).BorderWidth

            'BC = Array(&H80000014, &H80000015, &H80000016, &H80000010)
            BC = New String() {&H80000014, &H80000015, &H80000016, &H80000010}
        End If

        For I = 0 To (BorderWidth * 4 - 1)
            If LineControlArray.UBound < I Then
                'Load(LineControlArray(I))  Load method is not supported in vb.net
                LineControlArray(I).Visible = True
            End If
            LineControlArray(I).Bordercolor = BC((I \ 2) Mod 4)
        Next

        For I = 0 To (BorderWidth - 1)
            On Error Resume Next
            'Load LineControlArray(0 + (I * 4))   -> Load method is not supported in vb.net
            'Load LineControlArray(1 + (I * 4))
            'Load LineControlArray(2 + (I * 4))
            'Load LineControlArray(3 + (I * 4))
            On Error Resume Next
            MoveControl(LineControlArray(0 + (I * 4)), L + lW * I, T, L + lW * I, T + H)                      ' left
            MoveControl(LineControlArray(1 + (I * 4)), L, T + lW * I, L + W, T + lW * I)                      ' top
            MoveControl(LineControlArray(2 + (I * 4)), L + W - lW * (I + 1), T, L + W - lW * (I + 1), T + H)  ' right
            MoveControl(LineControlArray(3 + (I * 4)), L, T + H - (I + 1) * lW, L + W, T + H - (I + 1) * lW) ' bottom
        Next

        InitLineBorder = True
    End Function

    Public Function VersionControlDialog()
        frmVersionControl.ShowDialog()
    End Function

    Public Function MakeLong(ByVal WordHi As Object, ByVal WordLo As Integer) as integer
        ' it to overflow limits of multiplication which shifts
        ' it left.
        MakeLong = (WordHi * &H10000) + (WordLo And &HFFFF&)
    End Function

End Module


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
        FileClose(fName)
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
                If MsgBox("Can't access " & fName & ", try again?", vbCritical + vbYesNo, "File Error") = vbYes Then
                    Resume
                Else
                    End
                End If
            Case Else  ' An error we didn't foresee.
                MsgBox("An unforseen error [" & Err.Number & "] occurred accessing " & fName & "." & vbCrLf & "Your sale could not be completed." & vbCrLf & Err.Description, vbCritical, "BOS Number Error")
                End
        End Select
    End Function

End Module

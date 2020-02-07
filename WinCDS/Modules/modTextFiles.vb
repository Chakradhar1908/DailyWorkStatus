Module modTextFiles
    Public Function WriteFile(ByVal Filename As String, ByVal Str As String, Optional ByVal OverWrite As Boolean = False, Optional ByVal PreventNL As Boolean = False) As Boolean
        '::::WriteFile
        ':::SUMMARY
        ':Write the given string to a file.
        ':::DESCRIPTION
        ':Writes a given text string to a file.
        ':
        ':Text may or may not contain new lines (multi-line write supported).
        ':
        ':A New-line is appended by default if not specified in thes tring.
        ':::PARAMETERS
        ':- File - The name of the file to read.
        ':- str - The text to write to the file.  Can be an empty string (blank line).
        ':- [OverWrite] - Default is to append.  Set to TRUE to delete file before write (overwrite contents).
        ':- [PreventNL] - By default, the end of the string is checked for a new line.  Use this to write to a file without a new-line.
        ':::RETURN
        ':  Boolean - Returns True.
        ':::SEE ALSO
        ':  ReadEntireFile, WriteFile, CountLines

        'Dim FNo As Integer
        'On Error Resume Next
        'FNo = FreeFile()

        '      If OverWrite Then
        '          Kill file
        '  Open file For Output As #FNo
        'Else
        '          Open file For Append As #FNo
        'End If
        '      If PreventNL Or Right(Str, 2) = vbCrLf Then
        '          Print #FNo, Str;
        'Else
        '          Print #FNo, Str
        'End If
        '      Close #FNo
        'WriteFile = True

        Dim file As System.IO.StreamWriter

        On Error Resume Next
        If OverWrite Then
            My.Computer.FileSystem.DeleteFile(Filename) 'kill filename
            file = My.Computer.FileSystem.OpenTextFileWriter(Filename, False)
            'FileOpen(FNo, Filename, OpenMode.Output)
        Else
            'FileOpen(FNo, Filename, OpenMode.Append)
            file = My.Computer.FileSystem.OpenTextFileWriter(Filename, True)
        End If

        If PreventNL Or Right(Str, 2) = vbCrLf Then
            file.Write(Str)
            'FileSystem.PrintLine(FNo, Str)
        Else
            file.WriteLine(Str)
            'FileSystem.Print(FNo, Str)
        End If
        file.Close()
        WriteFile = True
    End Function

    Public Function ReadEntireFile(ByVal FileName As String) As String
        '::::ReadEntireFile
        ':::SUMMARY
        ':Read an entire file.
        ':::DESCRIPTION
        ':Reads  the full contents of a file and returns the value as a string (without modification).
        ':::PARAMETERS
        ':- FileName - The name of the file to read.
        ':::RETURN
        ':  String - The string contents of the file.
        ':::SEE ALSO
        ':  ReadFile, WriteFile, ReadEntireFileAndDelete

        On Error Resume Next
        'With CreateObject("Scripting.FileSystemObject")
        'ReadEntireFile = .OpenTextFile(FileName, 1).ReadAll
        'End With
        Dim f As System.IO.StreamReader
        f = My.Computer.FileSystem.OpenTextFileReader(FileName)
        ReadEntireFile = f.ReadToEnd
        f.Close()
        Dim fl = My.Computer.FileSystem.GetFileInfo(FileName)

        'If FileLen(FileName) / 10 <> Len(ReadEntireFile) / 10 Then
        'MessageBox.Show("ReadEntireFile was short: " & FileLen(FileName) & " vs " & Len(ReadEntireFile))
        'End If
        If fl.Length / 10 <> Len(ReadEntireFile) / 10 Then
            MessageBox.Show("ReadEntireFile was short: " & fl.Length & " vs " & Len(ReadEntireFile))
        End If

        '
        '  Dim intFile as integer
        '  intFile = FreeFile
        'On Error Resume Next
        '  Open FileName For Input As #intFile
        '  ReadEntireFile = Input$(LOF(intFile), #intFile)  '  LOF returns Length of File
        '  Close #intFile
    End Function

    Public Function ReadFile(ByVal FileName As String, Optional ByVal Startline As Integer = 1, Optional ByVal NumLines As Integer = 0) ', Optional ByRef WasEOF As Boolean = False)
        '::::ReadFile
        ':::SUMMARY
        ':Random Access Read a given file based on line number.
        ':::DESCRIPTION
        ':Reads the specified lines from a given file.
        ':
        ':If the file does not exist, no error is thrown, and an empty string is returned.
        ':::PARAMETERS
        ':- FileName - The name of the file to read.
        ':- StartLine - The line number to begin reading (the first line is 1).  If you try to read beyond the end of the file, an empty string is returned.
        ':- NumLines - If passed, attempts to read the specified number of lines.  Reading beyond the end of the file simply returns as many lines as possible.  Zero means read rest of file.  Default is zero.
        ':- WasEOF - If EOF checking is required, this ByRef parameter can be passed and checked later.  True if the file's EOF was reached.  False otherwise.
        ':::RETURN
        ':  String - The string contents of the file.
        ':::SEE ALSO
        ':  ReadEntireFile, WriteFile, CountLines, TailFile, HeadFile
        'Dim FNum as integer, Line As String, LineNum as integer, Count as integer
        Static CacheFileName As String
        Static CacheFileDate As String
        Static CacheFileLoad() As String

        ReadFile = Nothing
        If Not FileExists(FileName) Then
            '    WasEOF = True
            Exit Function
        End If

        If FileName = CacheFileName Then
            If FileDateTime(FileName) <> CacheFileDate Then CacheFileName = ""
        End If

        If FileName <> CacheFileName Then
            CacheFileName = FileName
            CacheFileDate = FileDateTime(FileName)
            CacheFileLoad = Split(Replace(ReadEntireFile(FileName), vbLf, ""), vbCr)
        End If

        If Startline = 1 And NumLines = 0 Then
            ReadFile = Join(CacheFileLoad, vbCrLf)
        Else
            ReadFile = Join(SubArr(CacheFileLoad, Startline - 1, NumLines), vbCrLf)
            '    ReadFile = LineByNumber(CacheFileLoad, Startline, NumLines)
        End If

        Exit Function

        '  If Startline < 1 Then Startline = 1
        '  LineNum = 0
        '  FNum = FreeFile
        '  Open FileName For Input As #FNum
        '  Do While Not EOF(FNum)
        '    LineNum = LineNum + 1
        '    Line Input #FNum, Line
        '    If LineNum >= Startline Then
        '      ReadFile = ReadFile & IIf(Len(ReadFile) > 0, vbCrLf, "") & Line
        '      Count = Count + 1
        '    End If
        '    If NumLines > 0 And Count >= NumLines Then GoTo Done
        ''    DoEvents
        '  Loop
        ''  WasEOF = True
        'Done:
        '  Close #FNum
    End Function
    Public Function ReadEntireFileAndDelete(ByVal FileName As String) As String
        '::::ReadEntireFileAndDelete
        ':::SUMMARY
        ':Read an entire file and safely delete it..
        ':::DESCRIPTION
        ':Reads the full contents of the file and then safely deletes it.
        ':
        ':If the file does not exist, no error is thrown, and an empty string is returned.
        ':::PARAMETERS
        ':- FileName - The name of the file to read.
        ':::RETURN
        ':  String - The string contents of the file.
        ':::SEE ALSO
        ':  ReadEntireFile

        On Error Resume Next
        ReadEntireFileAndDelete = ReadEntireFile(FileName)
        Kill(FileName)
    End Function

    Public Function CountFileLines(ByVal SourceFile As String, Optional ByVal IgnoreBlank As Boolean = False, Optional ByVal IgnorePrefix As String = "") As Integer
        '::::CountFileLines
        ':::SUMMARY
        ':Returns the number of lines in a given file.
        ':::DESCRIPTION
        ':Retruns the number of lines in a file, based on the number of vbCr characters.
        ':
        ':- vbLf is completely ignored.
        ':- Blank lines can be optionally ignored
        ':- A prefix (such as # or ') can also be omitted from the count.
        ':
        ':If the file does not exist, no error is thrown, and an empty string is returned.
        ':::PARAMETERS
        ':- Source - The name of the file to read.
        ':- IgnoreBlank - Ignore blank lines in count.  Set to False to count all lines.  Default == TRUE
        ':- IgnorePrefix - Specify a string prefix to ignore in the count.  Popular options are the VB comment character (') and the utility file comment character (#).
        ':::RETURN
        ':  Long - The number of lines.
        ':::SEE ALSO
        ':  WriteFile, ReadFile, VBFileCountLines, CountLines
        CountFileLines = CountLines(ReadEntireFile(SourceFile), IgnoreBlank, IgnorePrefix)
    End Function

    Public Function CountLines(ByVal Source As String, Optional ByVal IgnoreBlank As Boolean = True, Optional ByVal IgnorePrefix As String = "'") As Integer
        '::::CountLines
        ':::SUMMARY
        ':Returns the number of lines in a given string (not a file).
        ':::DESCRIPTION
        ':Retruns the number of lines in a string, based on the number of vbCr characters.
        ':
        ':- vbLf is completely ignored.
        ':- Blank lines can be optionally ignored
        ':- A prefix (such as # or ') can also be omitted from the count.
        ':
        ':If the file does not exist, no error is thrown, and an empty string is returned.
        ':::PARAMETERS
        ':- Source - The string to count lines in.
        ':- IgnoreBlank - Ignore blank lines in count.  Set to False to count all lines.  Default == TRUE
        ':- IgnorePrefix - Specify a string prefix to ignore in the count.  Popular options are the VB comment character (') and the utility file comment character (#).
        ':::RETURN
        ':  Long - The number of lines.
        ':::SEE ALSO
        ':  WriteFile, ReadFile, VBFileCountLines, CountFileLines, LineByNumber
        Dim L As Object
        Source = Replace(Source, vbLf, "")
        For Each L In Split(Source, vbCr)
            If Trim(L) = "" And IgnoreBlank Then
                ' Don't count...
            ElseIf IgnorePrefix <> "" And Left(LTrim(L), Len(IgnorePrefix)) = IgnorePrefix Then
                ' Don't count...
            Else
                CountLines = CountLines + 1
            End If
        Next
    End Function

End Module

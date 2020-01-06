Module modPath
    Public Const DIRSEP As String = "\"
    Public Function FileExists(ByVal F As String) As Boolean
        '::::FileExists
        ':::SUMMARY
        ':Returns T/F whether the file exists.
        ':::DESCRIPTION
        ':  Use to test whether a path exists.
        ':
        ':  This is the preferred method to test whether a file exists.
        ':  Equivalent to:  Dir(sFilePath) <> ""
        ':  The standard VB6 method for upgrade is not as convenient for eventual software migration.
        ':
        ':  Error safe.
        ':::PARAMETERS
        ': - sFilePath - The file path to test
        ':::RETURN
        ':  Boolean - Returns Success.
        ':::SEE ALSO
        ':  DirExists
        On Error Resume Next
        '  FileExists = LCase(Dir(F)) = GetFileName(F) And GetFileName(F) <> ""
        FileExists = (GetAttr(F) And vbDirectory) = 0
        Err.Clear()
    End Function

    Public Function AppFolder(Optional ByVal NoTrailingBackslash As Boolean = False) As String
        '::::AppFolder
        ':::SUMMARY
        ':Functional Replacement for the App.Folder setting in VB6.
        ':::DESCRIPTION
        ': Replacement function for the App.Folder.  Allows programatic control and handling within this
        ': typically Read-Only VB6 property.  Should be used as the only go-to function.
        ':
        ': In general, we don't use this much, simply because we because the Program Files folder
        ': is not generally writable.  We store things in the CDSData or AppData folders.
        ':::PARAMETERS
        ': - Boolean [bNoTrailingBackslash = False] - True for no trailing Backslash (App.Folder Compliance).  False for WinCDS standard.
        ':::RETURN
        ':  Returns the App folder:
        ':  - C:\Program Files\
        ':  - C:\Program Files (x86)\
        ':::SEE ALSO
        ':  CleanDir, FolderExists
        Dim A As String = ""

        ' Try to force the correct folder first..
#If False Then
  A = "C:\WinCDS\WinCDS\"
  If IsIDE Then AppFolder = A: Exit Function

  A = "C:\Program Files\WinCDS\"
  If FolderExists(A) Then AppFolder = A: Exit Function
  A = "C:\Program Files (x86)\WinCDS\"
  If FolderExists(A) Then AppFolder = A: Exit Function
#End If

        AppFolder = CleanDir(AppDomain.CurrentDomain.BaseDirectory, NoTrailingBackslash)
    End Function

    Public Function CleanDir(ByVal S As String, Optional ByVal NoTrailingBackslash As Boolean = False) As String
        '::::CleanDir
        ':::SUMMARY
        ':Performs basic sanitization of path string (no validation, no extension).
        ':::DESCRIPTION
        ':  Used to perform basic path cleaning.
        ':
        ':  - Remove '.'
        ':  - Remove '..' (todo - if requried)
        ':  - Remove '//'
        ':  - Appends or removes trailing DIRSEP (backslash)
        ':
        ':::PARAMETERS
        ': - sPath - The directory path to clean
        ': - [bNoTrailingBackslash = False] - True to remove trailing backslash.  False to force trailing Backslash.
        ':::RETURN
        ': Returns the cleaned path string.
        ':
        ':::SEE ALSO
        ':  FileExists
        ':
        '::Aliases
        ':  FolderExists
        If S = "" Then CleanDir = Nothing : Exit Function
        CleanDir = S

        Do While IsInStr(CleanDir, DIRSEP & DIRSEP)
            CleanDir = Replace(CleanDir, DIRSEP & DIRSEP, DIRSEP)
        Loop
        If Left(CleanDir, 2) = "." & DIRSEP Then CleanDir = Mid(CleanDir, 3)
        If Left(CleanDir, 2) = ".." Then CleanDir = Mid(CleanDir, 4)
        CleanDir = Replace(CleanDir, DIRSEP & "." & DIRSEP, DIRSEP)

        If NoTrailingBackslash Then
            If Right(CleanDir, 1) = DIRSEP Then CleanDir = Left(CleanDir, Len(CleanDir) - 1)
        Else
            If Right(CleanDir, 1) <> DIRSEP Then CleanDir = CleanDir & DIRSEP
        End If
    End Function
    Public Function DirExists(ByVal F As String) As Boolean
        '::::DirExists
        ':::SUMMARY
        ':Returns T/F whether the file exists.
        ':::DESCRIPTION
        ':  Use to test whether a path exists.
        ':
        ':  This is the preferred method to test whether a file exists.
        ':  Equivalent to:  DirExists = (GetAttr(F) And vbDirectory) = vbDirectory
        ':  Which is better than Dir(F, vbDirectory) <> ""
        ':  The standard VB6 method is complicated, not reliable, and not encapsulated.  Use DirExists exclusively
        ':  throughout WinCDS
        ':
        ':  Error safe.
        ':::PARAMETERS
        ': - sDirPath - The directory path to test
        ':::RETURN
        ':  Boolean - Returns Success.
        ':::SEE ALSO
        ':  FileExists
        ':
        '::Aliases
        ':  FolderExists
        On Error Resume Next
        DirExists = (GetAttr(F) And vbDirectory) = vbDirectory
        Err.Clear()
        '  If FileExists(F) Then Exit Function
        '  DirExists = Dir(F, vbDirectory) <> ""
    End Function
    Public Function DriveMapped(ByVal Drive As String) As Boolean
        '::::DriveMapped
        ':::SUMMARY
        ':Returns whether the drive is available (generally used on a mapped drive).
        ':::DESCRIPTION
        ':Returns whether the drive exists.
        ':::PAREMETERS
        ': - sDrive - The drive to check.
        ':::RETURNS
        ':Returns true if the drive exists.
        ':::SEE ALSO
        ': IsDriveRoot
        DriveMapped = DirExists(Left(Drive, 1) & ":")
    End Function
    Public Function GetFilePath(ByVal FN As String, Optional ByVal Sep As String = DIRSEP) As String
        '::::GetFilePath
        ':::SUMMARY
        ':Returns the directory-only portion of the given path.
        ':::DESCRIPTION
        ':Returns the path to the given file.
        ':If a directory is passed, it returns the (cleaned) directory parameter.
        ':Result will be cleaned and have a trailing backslash (always, as the root has a \ already).
        ':::PAREMETERS
        ': - sFileName - String - The file Path to parse.
        ': - [sSep] = DIRSEP - String - Allows for an alternate directory separator (such as / instead of \).
        ':::EXAMPLES
        ': - GetFileBase("C:\WinCDS\WinCDS\WinCDS.vbp")
        ':     - "C:\WinCDS\WinCDS\"
        ':::SEE ALSO
        ': CleanPath, FolderExists, FileExists, ParentDirectory
        ': GetFileName, GetFileExt
        Dim X
        On Error Resume Next
        If Sep = DIRSEP Then
            If IsDriveRoot(FN) Then GetFilePath = "" : Exit Function
            If DirExists(FN) Then
                GetFilePath = FN
                If Right(GetFilePath, 1) <> DIRSEP Then GetFilePath = GetFilePath & DIRSEP
                Exit Function
            End If
            If DirExists(FN & DIRSEP) Then GetFilePath = FN & DIRSEP : Exit Function
        End If
        X = Split(FN, Sep)
        'ReDim Preserve X(LBound(X) To UBound(X) - 1)
        ReDim Preserve X(0 To UBound(X) - 1)
        GetFilePath = Join(X, Sep)
        If Right(GetFilePath, 1) <> Sep Then GetFilePath = GetFilePath & Sep
    End Function

    Public Function IsDriveRoot(ByVal Path As String) As Boolean
        '::::IsDriveRoot
        ':::SUMMARY
        ':Returns whether the path specifies a volume root.
        ':::DESCRIPTION
        ': Predicate returns whether the path consists of "X:\", where X is A-Z.
        ':::PARAMETERS
        ': - sPath - The path to test
        ':::RETURN
        ':  Returns True/False if the path points to a volume root.
        ':::SEE ALSO
        ':  FolderExists, CleanDir
        IsDriveRoot = (Mid(Path, 2) = ":\")
    End Function
    Public Function GetFileName(ByVal FN As String, Optional ByVal NoLCase As Boolean = False, Optional ByVal Sep As String = DIRSEP) As String
        '::::GetFileName
        ':::SUMMARY
        ':Returns the file portion of a given path (no directory)
        ':::DESCRIPTION
        ':Returns the filename portion of the path, only, removing any absolute or relative directory information.
        ':::PAREMETERS
        ': - sFileName - String - The Path to parse.
        ': - [bNoLCase] = False - Boolean - Pass TRUE to not automatically lower-case the result.  Default is to lower-case all filenames.
        ': - [Sep] = DIRSEP - String - To use this funciton with a different directory separator (e.g., '/' instead of '\').
        ':::EXAMPLES
        ': - GetFileName("C:\WinCDS\WinCDS\WinCDS.vbp")
        ':     - "wincds.vbp"
        ':     - Note the result is automatically lower-cased unless over-ridden.
        ': - GetFileName("C:\WinCDS\WinCDS\WinCDS.vbp", True)
        ':     - "WinCDS.vbp"
        ':::NOTES
        ': If you just want the base name of the file, use GetFileBase.  If you just need an extension, GetFileExt
        ':::SEE ALSO
        ': GetFilePath, CleanPath, FolderExists, FileExists, ParentDirectory
        ': GetFileBase, GetFileExt
        Dim X
        If InStr(FN, Sep) = 0 Then GetFileName = FN : Exit Function
        X = Split(FN, Sep)
        GetFileName = X(UBound(X))
        If Not NoLCase Then GetFileName = LCase(GetFileName)
    End Function
    'Public Function RemoveFolder(ByVal F As String, Optional ByVal WithContents = True) As Boolean
    '::::RemoveFolder
    ':::SUMMARY
    ':Removes a folder completely, deleting all files if necessary.
    ':::DESCRIPTION
    ':Deletes a folder along with its contents.
    ':Functions as a shortcut for ClearFolder(f): RmDir f
    ':::PAREMETERS
    ': - sDirPath - Valid directory to delete files in.
    ': - [bWithContents] = TRUE - This will cause directory removal to fail if it is not empty.  Used as a precaution.
    ':::RETURNS
    ':Returns True on success, false otherwise.
    ':::SEE ALSO
    ': DeleteFileIfExists, ClearFolder
    'On Error Resume Next
    '   ClearFolder F, True
    'RmDir F
    'End Function
    Public Function CleanPath(ByVal F As String, Optional ByVal Path As String = "", Optional ByVal ForceTrailingSlash As Boolean = False) As String
        '::::CleanPath
        ':::SUMMARY
        ':Cleans a path.  Most Paths (ideally) should use this.
        ':::DESCRIPTION
        ':Removes relative paths, adds trailing Slash (or removes).
        ':
        ':::PARAMETERS
        ': - sPath - The Path to convert.
        ': - sCurr - The Current Path
        ': - [bForceTrailingSlash] = False - Whether the result should have a slash on the end or not
        ':::RETURN
        ':  Returns a string containing the clean path.
        ':::SEE ALSO
        ':  FolderExists, CleanDir, IsPathAbsolute, IsDriveRoot, MakePathAbsolute, RemoveTrailingSlash
        ':
        '::Aliases
        ':  IsAbsolutePath
        Dim N As String, X as integer, Y as integer
        'If IsPathAbsolute(F) Then
        '    N = F
        'ElseIf Path = "" Then
        '    N = AppFolder() & F
        'Else
        '    N = CleanDir(Path) & F
        'End If

        Do While InStr(N, "\\") <> 0
            N = Replace(N, "\\", "\")
        Loop

        Do While InStr(N, "\.\") <> 0
            N = Replace(N, "\.\", "\")
        Loop

        Do While InStr(N, "\..\") <> 0
            X = InStr(N, "\..\")
            Y = InStrRev(N, "\", X - 1)
            N = Replace(N, Mid(N, Y, X - Y + 3), "")
        Loop

        CleanPath = N
        If ForceTrailingSlash And FolderExists(CleanPath) Then
            If FolderExists(CleanPath) Then
                CleanPath = CleanDir(CleanPath)
            End If
        End If
    End Function
    Public Function FolderExists(ByVal F As String) As Boolean
        FolderExists = DirExists(F)
    End Function

    Public Function CanWriteToFolder(ByVal S As String) As Boolean
        '::::CanWriteToFolder
        ':::SUMMARY
        ':Verify write permissions to folder by doing a test-write to a temp file.
        ':::DESCRIPTION
        ':The directory must exist, and the software will attempt to write to a test file in the folder.
        ':Only if the folder exists and is writable will the function return TRUE.  Otherwise returns FALSE.
        ':Error safe.  Any error returns FALSE.
        ':::PAREMETERS
        ': - sDirPath - Valid directory to verify
        ':::EXAMPLES
        ': - CountFiles("C:\Windows\")
        ':::RETURNS
        ':Returns TRUE if the folder exists and can be written to.
        ':::SEE ALSO
        ': FolderExists, DateStampFile, WriteFile, FileExists
        ': DateStampFile, TempFile
        Dim F As String
        On Error GoTo Fail
        F = CleanDir(S) & DateStampFile("TESTWRITE-$.txt")
        WriteFile(F, "TEST WRITE", True)
        If Not FileExists(F) Then Exit Function
        DeleteFileIfExists(F)
        If FileExists(F) Then Exit Function
        CanWriteToFolder = True
Fail:
    End Function
    Public Function DeleteFileIfExists(ByVal sFIle As String, Optional ByVal bNoAttributeClearing As Boolean = False) As Boolean
        '::::DeleteFileIfExists
        ':::SUMMARY
        ':Deletes a file whehther or not it exists.
        ':::DESCRIPTION
        ':This is an ERROR-SAFE alternative to the VB6 Kill statement.  Because `Kill sFile` results in an error if the delete fails, or simply
        ':if the file does not exist, this functions as a safe way to delete a file without having to provide error-checking.
        ':
        ':It is the preferred WinCDS method for deleting a file.
        ':
        ':In addition, this function ALSO handles attribute clearing.
        ':Functions as a shortcut for ClearFolder(f): RmDir f
        ':::PAREMETERS
        ': - sFile - Valid directory to delete files in.
        ': - [bNoAttributeClearing] = False - Set to TRUE if you do NOT want to clear any Read-Only or other file attributes.  It may make the function fail silently, of course, but that would be the point.
        ':::RETURNS
        ':Returns True.  Checking is expected of calling function.
        ':::SEE ALSO
        ': DeleteFileIfExists, ClearFolder
        On Error Resume Next
        If Not FileExists(sFIle) Then Exit Function
        If Not bNoAttributeClearing Then SetAttr(sFIle, 0)
        If FileExists(sFIle) Then Kill(sFIle)
        '  DeleteFileIfExists = FileExists(sFile)
        DeleteFileIfExists = True
    End Function

    Public Function GetFileBase(ByVal vFile As String, Optional ByVal NoLCase As Boolean = False, Optional ByVal WithPath As Boolean = False) As String
        '::::GetFileBase
        ':::SUMMARY
        ':Returns the base-name only of the given file path (no path or extension).
        ':::DESCRIPTION
        ':Returns the base-name of the given file path (no extension).
        ':::PAREMETERS
        ': - sFileName - String - The file Path to parse.
        ': - [bNoLCase] = False - Boolean - Pass TRUE to not automatically lower-case the result.  Default is to lower-case all filenames.
        ': - [bWithPath] = False - Boolean - Optionally, sometimes it is useful to have the path included.  Pass TRUE to include the path with the base name (stripping extension, however)
        ':::EXAMPLES
        ': - GetFileBase("C:\WinCDS\WinCDS\WinCDS.VBP")
        ':     - "wincds"
        ':     - Note the result is automatically lower-cased unless over-ridden.
        ': - GetFileName("C:\WinCDS\WinCDS\WinCDS.VBP", True)
        ':     - "WinCDS"
        ': - GetFileName("C:\WinCDS\WinCDS\WinCDS.VBP", , True)
        ':     - "c:\wincds\wincds\wincds"
        ':::SEE ALSO
        ': GetFilePath, CleanPath, FolderExists, FileExists, ParentDirectory
        ': GetFileName, GetFileExt
        Dim FN As String, P As String, X
        P = GetFilePath(vFile)
        FN = GetFileName(vFile)
        If InStr(FN, ".") = 0 Then GetFileBase = FN : Exit Function
        X = Split(FN, ".")
        'ReDim Preserve X(LBound(X) To UBound(X) - 1)
        ReDim Preserve X(0 To UBound(X) - 1)
        GetFileBase = Join(X, ".")
        If Not NoLCase Then GetFileBase = LCase(GetFileBase)
        If WithPath Then GetFileBase = CleanDir(P) & GetFileBase
    End Function

    Public Function GetFileExt(ByVal FN As String, Optional ByVal NoLCase As Boolean = False) As String
        '::::GetFileExt
        ':::SUMMARY
        ':Returns the extension only of the given file path.
        ':::DESCRIPTION
        ':Returns the extension after the exention separator (".").
        ':::PAREMETERS
        ': - sFileName - String - The file Path to parse.
        ': - [bNoLCase] = False - Boolean - Pass TRUE to not automatically lower-case the result.  Default is to lower-case all filenames.
        ':::EXAMPLES
        ': - GetFileName("C:\WinCDS\WinCDS\WinCDS.VBP")
        ':     - "wincds.vbp"
        ':     - Note the result is automatically lower-cased unless over-ridden.
        ': - GetFileName("C:\WinCDS\WinCDS\WinCDS.VBP", True)
        ':     - "WinCDS.VBP"
        ':::SEE ALSO
        ': GetFilePath, CleanPath, FolderExists, FileExists, ParentDirectory
        ': GetFileName, GetFileExt
        ': FileIsImage
        Dim X
        FN = GetFileName(FN)
        If InStr(FN, ".") = 0 Then GetFileExt = "" : Exit Function
        X = Split(FN, ".")
        GetFileExt = LCase(X(UBound(X)))
        If Not NoLCase Then GetFileExt = LCase(GetFileExt)
    End Function

    Public Function ParentDirectory(ByVal FN As String) As String
        '::::ParentDirectory
        ':::SUMMARY
        ':Returns the parent folder of the one passed
        ':::DESCRIPTION
        ':Returns a path to the parent of the directory given.
        ':::REMARKS
        ':This must be smart/aware.
        ': - If the path is not absolute, nothing can be done.  Returns passed (could be improved).
        ': - If the path is root already, we cannot go higher (returns passed).
        ': - If the path is a file (as opposed to a directory), return the parent of the path the file is in.
        ': - Otherwise, return equivalent of [...]\Dir\..\
        ': - Maintains trailing backslash of parent directory.  Includes trailing backslash if pointed to a file.
        ':::EXAMPLES
        ': -ParentDirectory("C:\CDSData\Store1\")
        ':     - Returns "C:\CDSData\"
        ':     - Passed Parameter had trailing backslash, so it is preserved in result.
        ': -ParentDirectory("C:\WinCDS\Backups")
        ':     - Returns "C:\WinCDS"
        ':     - Passed Parameter did NOT have trailing backslash, so form is preserved in result.
        ': -ParentDirectory("C:\WinCDS\WinCDS\WinCDS.vbp"
        ':     - Returns "C:\WinCDS\"
        ':     - WinCDS.vbp is a file.  The directory is \WinCDS\WinCDS\, so function returns \WinCDS
        ':     - Passed parameter was a file, so result contains trailing backslash.
        ':
        ':::PARAMETERS
        ': - sPath - The Path to convert.
        ':::RETURN
        ':  Returns a string containing the clean path.
        ':::SEE ALSO
        ':  GetFilePath, CleanPath, FolderExists, FileExists
        ':::Aliases
        ': ParentFolder, ParentDir
        Dim X
        If Not IsPathAbsolute(FN) Then ParentDirectory = FN : Exit Function
        If IsDriveRoot(FN) Then ParentDirectory = FN : Exit Function
        If FileExists(FN) Then FN = GetFilePath(FN)
        If Right(FN, 1) = DIRSEP Then FN = Left(FN, Len(FN) - 1)
        X = Split(FN, DIRSEP)
        'ReDim Preserve X(LBound(X) To UBound(X) - 1)
        ReDim Preserve X(0 To UBound(X) - 1)
        ParentDirectory = Join(X, DIRSEP)
        If Right(ParentDirectory, 1) <> DIRSEP Then ParentDirectory = ParentDirectory & DIRSEP
    End Function

    Public Function IsPathAbsolute(ByVal Path As String) As Boolean
        '::::IsPathAbsolute
        ':::SUMMARY
        ':Determines whether the path specified is an absolute path.
        ':::DESCRIPTION
        ':Predicate returns true when the path is absolute.
        ':Absolute paths are those that refer to the root path and have no relative path components.
        ':::PARAMETERS
        ': - sPath - The path to test
        ':::RETURN
        ':  Returns True/False if the path points to an absolute path.
        ':::SEE ALSO
        ':  FolderExists, CleanDir, IsDriveRoot, MakePathAbsolute
        ':
        '::Aliases
        ':  IsAbsolutePath
        IsPathAbsolute = (Mid(Path, 2, 2) = ":\") And Not IsInStr(Path, "\.\") And Not IsInStr(Path, "\..\")
    End Function

    Public Function AllFiles(ByVal DirPath As String) As String()
        '::::AllFiles
        ':::SUMMARY
        ':Return an array containing name of all files in directory specified
        ':::DESCRIPTION
        ':Given a valid directory, will return a String() array containing all filenames (without path).
        ':::PAREMETERS
        ': - DirPath - Valid directory ending in DIRSEP ("\") or with a valid wildcard specifier.
        ':::EXAMPLES
        ': - DirPath("C:\Windows\")
        ': - DirPath("C:\Windows\*.exe")
        ':
        ':Dim sFiles() As String
        ':Dim lCtr As Long
        ':
        ':sFiles = AllFiles("C:\windows\")
        ':For lCtr = 0 To UBound(sFiles)
        ':  Debug.Print sFiles(lCtr)
        ':Next
        ':
        ':::RETURNS
        ':Returns a string array of the files in the folder.
        ':::SEE ALSO
        ': FolderExists, CleanPath
        ': AllFolders
        '***************************************************
        'PURPOSE: RETURN AN ARRAY CONTAINING NAME OF ALL FILES IN
        'DIRECTORY SPECIFIED BY DIR PATH
        '
        'PARAMETER:
        '
        '  'DIRPATH: A VALID DRIVE OR SUBDIRECTORY ON YOUR SYSTEM,
        '  'ENDING WITH FORWARD SLASH (\) CHARACTER, OR
        '  'A DRIVE OR SUBDIRECTORY FOLLOWED BY A WILD CARD
        '  'STRING (e.g., C:\WINDOWS\*.txt)
        '
        'RETURNS: A STRING ARRAY WITH THE NAMES OF ALL FILENAMES
        'IN THE DIRECTORY, INCLUDING HIDDEN, SYSTEM, AND READ-ONLY FILES
        'THE FUNCTION IS NON RECURSIVE, I.E., IT DOES NOT SEARCH
        'SUBDIRECTORIES UNDERNEATH DIRPATH

        'REQUIRES: VB6, BECAUSE IT RETURNS A STRING ARRAY

        'EXAMPLE
        'Dim sFiles() As String
        'Dim lCtr As Long

        'sFiles = AllFiles("C:\windows\")
        'For lCtr = 0 To UBound(sFiles)
        '    Debug.Print sFiles(lCtr)
        'Next
        '********************************************************

        Dim sFIle As String
        Dim lElement As Long
        Dim sAns() As String
        ReDim sAns(0)

        If InStr(DirPath, "*") = 0 Then
            If Right(DirPath, 1) <> DIRSEP Then DirPath = DirPath & DIRSEP
        End If


        sFIle = Dir(DirPath, vbNormal + vbHidden + vbReadOnly + vbSystem + vbArchive)
        If sFIle <> "" Then
            sAns(0) = sFIle
            Do
                sFIle = Dir()
                If sFIle = "" Then Exit Do
                lElement = IIf(sAns(0) = "", 0, UBound(sAns) + 1)
                ReDim Preserve sAns(lElement) As String
      sAns(lElement) = sFIle
            Loop
        End If
        AllFiles = sAns
    End Function

End Module

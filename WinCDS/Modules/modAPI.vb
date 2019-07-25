'Imports System.Runtime.InteropServices

Module modAPI
    Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Integer) As Integer
    ' ------------------------------
    ' GetCompName
    ' Gets the computer network name.
    ' ------------------------------
    Public Declare Function GetDC Lib "USER32" (ByVal hwnd As Integer) As Integer
    Public Declare Function GetClientRect Lib "USER32" (ByVal hwnd As Integer, lpRect As RECT) As Integer
    Public Declare Function GetWindowRect Lib "USER32" (ByVal hwnd As Integer, lpRect As RECT) As Integer
    Public Declare Function RedrawWindow Lib "USER32" (ByVal hwnd As Integer, lprcUpdate As RECT, ByVal hrgnUpdate As Integer, ByVal fuRedraw As Integer) As Integer
    Public Declare Function InvalidateRect Lib "USER32" (ByVal hwnd As Integer, lpRect As RECT, ByVal bErase As Integer) As Integer
    Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Object, Source As Object, ByVal Length As Integer)
    Private Declare Function GdiAlphaBlend Lib "GDI32" (ByVal hDC As Integer, ByVal linT As Integer, ByVal linT As Integer, ByVal linT As Integer, ByVal linT As Integer, ByVal hDC As Integer, ByVal linT As Integer, ByVal linT As Integer, ByVal linT As Integer, ByVal linT As Integer, ByVal BLENDFUNCT As Integer) As Integer
    Public Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer

    Public Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Integer, ByVal hwnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Public Declare Function GetDesktopWindow Lib "USER32" () As Integer
    Public Declare Function GetTickCount Lib "kernel32" () As Integer  ' use for timerless timers
    Private Declare Function GetFileVersionInfoSize Lib "Version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Integer) As Integer
    Private Declare Function GetFileVersionInfo Lib "Version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Integer, ByVal dwlen As Integer, lpData As Object) As Integer
    Private Declare Function VerQueryValue Lib "Version" Alias "VerQueryValueA" (pBlock As Object, ByVal lpSubBlock As String, lplpBuffer As Object, puLen As Integer) As Integer
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Object, ByVal Source As Integer, ByVal Length As Integer)
    Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Integer) As Integer
    Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Integer
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    'Public Declare Function LockWindowUpdate Lib "USER32" (ByVal hwnd As Integer) As Integer  -> This is for vb6.0.
    '<DllImport("user32.dll")>
    'Public Function LockWindowUpdate(ByVal hWndLock As IntPtr) As Boolean
    'End Function
    Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hWndLock As IntPtr) As Boolean

    Private Const RDW_INTERNALPAINT As Integer = &H2
    Private Const RDW_UPDATENOW As Integer = &H100
    Public Const AC_SRC_OVER As Integer = &H0 '(0)
    Public Const GWL_WNDPROC As Integer = (-4)
    Public Const WM_PAINT As Integer = &HF
    Public Const SW_SHOWDEFAULT As Integer = 10
    Public Const SW_SHOWNORMAL As Integer = 1
    Const VS_FFI_SIGNATURE As Integer = &HFEEF04BD
    Const VS_FFI_STRUCVERSION As Integer = &H10000
    Const VS_FFI_FILEFLAGSMASK As Integer = &H3F&
    Const VS_FF_DEBUG As Integer = &H1
    Const VS_FF_PRERELEASE As Integer = &H2
    Const VS_FF_PATCHED As Integer = &H4
    Const VS_FF_PRIVATEBUILD As Integer = &H8
    Const VS_FF_INFOINFERRED As Integer = &H10
    Const VS_FF_SPECIALBUILD As Integer = &H20
    Const VOS_UNKNOWN As Integer = &H0
    Const VOS_DOS As Integer = &H10000
    Const VOS_OS216 As Integer = &H20000
    Const VOS_OS232 As Integer = &H30000
    Const VOS_NT As Integer = &H40000
    Const VOS__BASE As Integer = &H0
    Const VOS__WINDOWS16 As Integer = &H1
    Const VOS__PM16 As Integer = &H2
    Const VOS__PM32 As Integer = &H3
    Const VOS__WINDOWS32 As Integer = &H4
    Const VOS_DOS_WINDOWS16 As Integer = &H10001
    Const VOS_DOS_WINDOWS32 As Integer = &H10004
    Const VOS_OS216_PM16 As Integer = &H20002
    Const VOS_OS232_PM32 As Integer = &H30003
    Const VOS_NT_WINDOWS32 As Integer = &H40004
    Const VFT_UNKNOWN As Integer = &H0
    Const VFT_APP As Integer = &H1
    Const VFT_DLL As Integer = &H2
    Const VFT_DRV As Integer = &H3
    Const VFT_FONT As Integer = &H4
    Const VFT_VXD As Integer = &H5
    Const VFT_STATIC_LIB As Integer = &H7
    Const VFT2_UNKNOWN As Integer = &H0
    Const VFT2_DRV_PRINTER As Integer = &H1
    Const VFT2_DRV_KEYBOARD As Integer = &H2
    Const VFT2_DRV_LANGUAGE As Integer = &H3
    Const VFT2_DRV_DISPLAY As Integer = &H4
    Const VFT2_DRV_MOUSE As Integer = &H5
    Const VFT2_DRV_NETWORK As Integer = &H6
    Const VFT2_DRV_SYSTEM As Integer = &H7
    Const VFT2_DRV_INSTALLABLE As Integer = &H8
    Const VFT2_DRV_SOUND As Integer = &H9
    Const VFT2_DRV_COMM As Integer = &HA

    Dim UGridIOHooks As Collection, UGridIO_Painting As Boolean
    Public Const WINVER_WINXP As String = "5.1.2600"
    Public Const VER_PLATFORM_WIN32s As Integer = 0
    Public Const VER_PLATFORM_WIN32_WINDOWS As Integer = 1
    Public Const VER_PLATFORM_WIN32_NT As Integer = 2

    Public Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure
    Public Structure BLENDFUNCTION
        Dim BlendOp As Byte
        Dim BlendFlags As Byte
        Dim SourceConstantAlpha As Byte
        Dim AlphaFormat As Byte
    End Structure
    Public Structure VersionInformationType
        Dim StructureVersion As String
        Dim FileVersion As String
        Dim FileVersion_Major As Integer
        Dim FileVersion_Minor As Integer
        Dim FileVersion_vbBuild As Integer
        Dim FileVersion_vbRevision As Integer
        Dim ProductVersion As String
        Dim FileFlags As String
        Dim TargetOperatingSystem As String
        Dim FileType As String
        Dim FileSubtype As String
    End Structure
    Private Structure VS_FIXEDFILEINFO
        Dim dwSignature As Integer
        Dim dwStrucVersionl As Integer      ' e.g. = &h0000 = 0
        Dim dwStrucVersionh As Integer      ' e.g. = &h0042 = .42
        Dim dwFileVersionMSl As Integer     ' e.g. = &h0003 = 3
        Dim dwFileVersionMSh As Integer     ' e.g. = &h0075 = .75
        Dim dwFileVersionLSl As Integer     ' e.g. = &h0000 = 0
        Dim dwFileVersionLSh As Integer     ' e.g. = &h0031 = .31
        Dim dwProductVersionMSl As Integer  ' e.g. = &h0003 = 3
        Dim dwProductVersionMSh As Integer  ' e.g. = &h0010 = .1
        Dim dwProductVersionLSl As Integer  ' e.g. = &h0000 = 0
        Dim dwProductVersionLSh As Integer  ' e.g. = &h0031 = .31
        Dim dwFileFlagsMask As Integer         ' = &h3F for version "0.42"
        Dim dwFileFlags As Integer             ' e.g. VFF_DEBUG Or VFF_PRERELEASE
        Dim dwFileOS As Integer                ' e.g. VOS_DOS_WINDOWS16
        Dim dwFileType As Integer              ' e.g. VFT_DRIVER
        Dim dwFileSubtype As Integer           ' e.g. VFT2_DRV_KEYBOARD
        Dim dwFileDateMS As Integer            ' e.g. 0
        Dim dwFileDateLS As Integer            ' e.g. 0
    End Structure
    Public Structure OSVERSIONINFO 'windows-defined type OSVERSIONINFO
        Dim OSVSize As Integer         'size, in bytes, of this data structure
        Dim dwVerMajor As Integer         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
        Dim dwVerMinor As Integer         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
        Dim dwBuildNumber As Integer         'NT: build number of the OS
        'Win9x: build number of the OS in low-order word.
        '       High-order word contains major & minor ver nos.
        Dim PlatformID As Integer         'Identifies the operating system platform.
        <VBFixedString(128)> Dim szCSDVersion As String 'NT: string, such as "Service Pack 3"
        'Win9x: 'arbitrary additional information'
    End Structure


    Public Function GetLocalComputerName() As String
        Dim sBuffer As String
        Dim lReturn As Integer
        sBuffer = Space(255)
        lReturn = GetComputerName(sBuffer, Len(sBuffer))
        GetLocalComputerName = Trim(Left(sBuffer, InStr(sBuffer, vbNullChar) - 1))
    End Function
    Public Function DrawRectangle(ByVal hwnd As Integer, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer, ByVal Color As Integer, Optional ByVal Transparency As Integer = 75, Optional ByVal Invalidate1st As Boolean = True) As Boolean
        Dim DC As Integer
        DC = GetDC(hwnd)

        If Invalidate1st Then
            InvalidateRectangle(hwnd, L, T, W, H)
            RedrawWindowRectangle(hwnd, L, T, W, H)
        End If

        DrawRectangle = DrawRectangleToDC(DC, L, T, W, H, Color, Transparency, Invalidate1st)
    End Function
    Public Function InvalidateRectangle(ByVal hwnd As Integer, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As Boolean
        On Error Resume Next
        Dim tr As RECT
        GetWindowRect(hwnd, tr)
        tr.Left = L
        tr.Top = T
        tr.Right = L + W
        tr.Bottom = T + H
        InvalidateRect(hwnd, tr, 1)
    End Function
    Public Function RedrawWindowRectangle(ByVal hwnd As Integer, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer) As Boolean
        On Error Resume Next
        Dim tr As RECT
        GetWindowRect(hwnd, tr)
        tr.Left = L
        tr.Top = T
        tr.Right = L + W
        tr.Bottom = T + H
        RedrawWindow(hwnd, tr, 0, RDW_INTERNALPAINT + RDW_UPDATENOW)
    End Function
    Public Function DrawRectangleToDC(ByVal DC As Integer, ByVal L As Integer, ByVal T As Integer, ByVal W As Integer, ByVal H As Integer, ByVal Color As Integer, Optional ByVal Transparency As Integer = 75, Optional ByVal Invalidate1st As Boolean = True) As Boolean
        Dim PTB As PictureBox
        Dim lBlend As Integer, Bf As BLENDFUNCTION
        Dim N As Integer

        Bf.BlendOp = AC_SRC_OVER
        Bf.BlendFlags = 0
        ' 255 = 100% opaque, 128 = 50% opaque, 64 = 75% transparent
        Bf.SourceConstantAlpha = 255 * Transparency / 100
        Bf.AlphaFormat = 0
        RtlMoveMemory(lBlend, Bf, 4)

        'PTB = MainMenu.picAlpha

        'PTB.Width = W * Screen.TwipsPerPixelX
        'PTB.Height = H * Screen.TwipsPerPixelY
        '  PTB.Line (0, 0)-(W, H), vbWhite, BF
        '  N = GdiAlphaBlend(dC, L, T, W, H, PTB.hdc, 0, 0, 1, 1, lBlend)

        'PTB.Line(0, 0)-(W, H), Color, BF

        Dim g As Graphics = PTB.CreateGraphics


        'N = GdiAlphaBlend(DC, L, T, W, H, PTB.hDC, 0, 0, 1, 1, lBlend)
        N = GdiAlphaBlend(DC, L, T, W, H, g.GetHdc, 0, 0, 1, 1, lBlend)
        '  Debug.Print "n=" & N
        DrawRectangleToDC = N <> 0
        '  EditPO.UGridIO1.ColorColumn 2, vbRed
    End Function
    Public Sub UGridIO_AddHook(ByRef hwnd As Integer, ByRef Obj As Object)
        Dim X As Object
        If UGridIOHooks Is Nothing Then UGridIOHooks = New Collection
        Debug.Print("uGridIO-" & IIf(Obj Is Nothing, "rem", "add") & " hook " & hwnd)

        On Error Resume Next
        If Obj Is Nothing Then
            UGridIOHooks.Remove(CStr(hwnd))
        Else
            X = UGridIOHooks(CStr(hwnd))
            If Not X Is Nothing Then UGridIOHooks.Remove(CStr(hwnd))
            UGridIOHooks.Add(Obj, CStr(hwnd))
        End If
    End Sub
    Public Function UGridIO_Paint(ByVal hwnd As Integer, ByVal uMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
        Dim X As Object
        '  Debug.Print "ugridio_paint: " & hwnd & ", Msg = " & uMsg

        If UGridIOHooks Is Nothing Then Exit Function
        On Error Resume Next
        X = UGridIOHooks(CStr(hwnd))
        If X Is Nothing Then Exit Function

        Select Case uMsg
            Case WM_PAINT
                Dim A As Integer, B As Integer, R As RECT
                UGridIO_Paint = CallWindowProc(X.PrevProc, hwnd, uMsg, wParam, lParam)
                Debug.Print("Painting UGridIO... wparam=" & wParam & ", lparam=" & lParam)
                If Not UGridIO_Painting Then
                    UGridIO_Painting = True

                    '        RedrawWindowL hwnd, 0, 0, RDW_INTERNALPAINT + RDW_UPDATENOW

                    X.ColorColumns(wParam, lParam)
                    UGridIO_Painting = False
                End If
                '    Case WM_ERASEBKGND
                '      UGridIO_Paint = CallWindowProc(X.PrevProc, hwnd, uMsg, Wparam, lParam)
            Case Else
                UGridIO_Paint = CallWindowProc(X.PrevProc, hwnd, uMsg, wParam, lParam)


                '    Case WM_ERASEBKGND
                '    Case Else:   Debug.Print "UgridIO unkonwn-message: " & uMsg
        End Select
    End Function
    Public Delegate Function UGridIO_PaintDelegate(ByVal h As Integer, ByVal u As Integer, ByVal w As Integer, ByVal l As Integer) As Integer
    Public Sub RunShellExecute(ByVal sTopic As String, ByVal sFIle As Object, Optional ByVal sParams As Object = "", Optional ByVal sDirectory As Object = "", Optional ByVal nShowCmd As Integer = SW_SHOWNORMAL)
        'execute the passed operation, passing the desktop as the window to receive any error messages
        '  If sDirectory = "" Then sDirectory = AppFolder
        LastProcessID = ShellExecute(GetDesktopWindow(), sTopic, sFIle, sParams, sDirectory, nShowCmd)
    End Sub
    Public Function EnsureFolderExists(ByVal sFileName As String, Optional ByVal bCreate As Boolean = False) As String
        If Len(sFileName) = 0 Then EnsureFolderExists = UpdateFolder() : 
        Exit Function

        EnsureFolderExists = CleanPath(sFileName, , True)

        On Error Resume Next
        If Not FolderExists(sFileName) Then MkDir(sFileName)
    End Function
    Public Function GetWinVerNumber() As String
        Dim OS As String, N As Integer
        OS = RunCmdToOutput("ver")
        N = InStr(OS, "[")
        OS = Trim(Replace(Replace(Replace(Mid(OS, N + 1), "]", ""), "Version", ""), vbCrLf, ""))
        GetWinVerNumber = OS

        'If Left(OS, 4) = "4.0." Then OS = "Win95"
        'If Left(OS, 4) = "4.1." Then OS = "Win98"
        'If Left(OS, 4) = "4.90" Then OS = "WinME"
        'If Left(OS, 4) = "5.0." Then OS = "Win2000"
        'If Left(OS, 4) = "5.1." Then OS = "WinXP"
        'If Left(OS, 4) = "5.2." Then OS = "WinXPx64 or Server 2003"
        'If Left(OS, 4) = "6.0." Then OS = "Vista or Server 2008"
        'If Left(OS, 4) = "6.1." Then OS = "Win7 or Server 2008R2"
        'If Left(OS, 4) = "6.2." Then OS = "Win8 or Server 2012"
        'If Left(OS, 4) = "6.3." Then OS = "Win8.1 or Server 2012R2"
        'If Left(OS, 5) = "10.0." Then OS = "Win10"
    End Function
    Public Function VersionInformation(ByVal FILE_Name As String) As VersionInformationType
        Dim Dummy_handle As Integer
        Dim Buffer() As Byte
        Dim Info_size As Integer
        Dim Info_address As Integer
        Dim Fixed_File_Info As VS_FIXEDFILEINFO
        Dim Fixed_File_Info_Size As Integer
        Dim Result As VersionInformationType

        ' Get the version information buffer size.
        Info_size = GetFileVersionInfoSize(FILE_Name, Dummy_handle)
        If Info_size = 0 Then
            '      MsgBox "No version information available"
            Exit Function
        End If

        ' Load the fixed file information into a buffer.
        ReDim Buffer(0 To Info_size - 1)
        If GetFileVersionInfo(FILE_Name, 0&, Info_size,
      Buffer(1)) = 0 Then
            MsgBox("Error getting version information")
            Exit Function
        End If
        If VerQueryValue(Buffer(1), "\", Info_address, Fixed_File_Info_Size) = 0 Then
            MsgBox("Error getting fixed file version information")
            Exit Function
        End If

        ' Copy the information from the buffer into a usable structure.
        MoveMemory(Fixed_File_Info, Info_address, Len(Fixed_File_Info))

        ' Get the version information.
        ' Structure version.
        Result.StructureVersion = Format(Fixed_File_Info.dwStrucVersionh) & "." & Format(Fixed_File_Info.dwStrucVersionl)

        ' File version number.
        Result.FileVersion = Format(Fixed_File_Info.dwFileVersionMSh) & "." & Format(Fixed_File_Info.dwFileVersionMSl) & "." & Format(Fixed_File_Info.dwFileVersionLSh) & "." & Format(Fixed_File_Info.dwFileVersionLSl)
        Result.FileVersion_Major = Fixed_File_Info.dwFileVersionMSh
        Result.FileVersion_Minor = Fixed_File_Info.dwFileVersionMSl
        Result.FileVersion_vbBuild = Fixed_File_Info.dwFileVersionLSh
        Result.FileVersion_vbRevision = Fixed_File_Info.dwFileVersionLSl

        ' Product version number.
        Result.ProductVersion = Format(Fixed_File_Info.dwProductVersionMSh) & "." & Format(Fixed_File_Info.dwProductVersionMSl) & "." & Format(Fixed_File_Info.dwProductVersionLSh) & "." & Format(Fixed_File_Info.dwProductVersionLSl)

        ' File attributes.
        Result.FileFlags = ""
        If Fixed_File_Info.dwFileFlags And VS_FF_DEBUG Then Result.FileFlags = Result.FileFlags & " Debug"
        If Fixed_File_Info.dwFileFlags And VS_FF_PRERELEASE Then Result.FileFlags = Result.FileFlags & " PreRel"
        If Fixed_File_Info.dwFileFlags And VS_FF_PATCHED Then Result.FileFlags = Result.FileFlags & " Patched"
        If Fixed_File_Info.dwFileFlags And VS_FF_PRIVATEBUILD Then Result.FileFlags = Result.FileFlags & " Private"
        If Fixed_File_Info.dwFileFlags And VS_FF_INFOINFERRED Then Result.FileFlags = Result.FileFlags & " Info"
        If Fixed_File_Info.dwFileFlags And VS_FF_SPECIALBUILD Then Result.FileFlags = Result.FileFlags & " Special"
        If Fixed_File_Info.dwFileFlags And VFT2_UNKNOWN Then Result.FileFlags = Result.FileFlags + " Unknown"
        If Len(Result.FileFlags) > 0 Then Result.FileFlags = Mid(Result.FileFlags, 2)

        ' Target operating system.
        Select Case Fixed_File_Info.dwFileOS
            Case VOS_DOS_WINDOWS16 : Result.TargetOperatingSystem = "DOS-Win16"
            Case VOS_DOS_WINDOWS32 : Result.TargetOperatingSystem = "DOS-Win32"
            Case VOS_OS216_PM16 : Result.TargetOperatingSystem = "OS/2-16 PM-16"
            Case VOS_OS232_PM32 : Result.TargetOperatingSystem = "OS/2-16 PM-32"
            Case VOS_NT_WINDOWS32 : Result.TargetOperatingSystem = "NT-Win32"
            Case 4 : Result.TargetOperatingSystem = "Win32"
            Case Else : Result.TargetOperatingSystem = "Unknown"
        End Select

        ' File type.
        Select Case Fixed_File_Info.dwFileType
            Case VFT_APP : Result.FileType = "App"
            Case VFT_DLL : Result.FileType = "DLL"
            Case VFT_DRV
                Result.FileType = "Driver"
                Select Case Fixed_File_Info.dwFileSubtype
                    Case VFT2_DRV_PRINTER : Result.FileSubtype = "Printer drv"
                    Case VFT2_DRV_KEYBOARD : Result.FileSubtype = "Keyboard drv"
                    Case VFT2_DRV_LANGUAGE : Result.FileSubtype = "Language drv"
                    Case VFT2_DRV_DISPLAY : Result.FileSubtype = "Display drv"
                    Case VFT2_DRV_MOUSE : Result.FileSubtype = "Mouse drv"
                    Case VFT2_DRV_NETWORK : Result.FileSubtype = "Network drv"
                    Case VFT2_DRV_SYSTEM : Result.FileSubtype = "System drv"
                    Case VFT2_DRV_INSTALLABLE : Result.FileSubtype = "Installable"
                    Case VFT2_DRV_SOUND : Result.FileSubtype = "Sound drv"
                    Case VFT2_DRV_COMM : Result.FileSubtype = "Comm drv"
                    Case VFT2_UNKNOWN : Result.FileSubtype = "Unknown"
                End Select
            Case VFT_FONT : Result.FileType = "Font"
            Case VFT_VXD : Result.FileType = "VxD"
            Case VFT_STATIC_LIB : Result.FileType = "Lib"
            Case Else : Result.FileType = "Unknown"
        End Select

        VersionInformation = Result
    End Function
    Public Sub RunShellExecuteAdmin(ByVal App As String, Optional ByVal nHwnd As Integer = 0, Optional ByVal WindowState As Integer = SW_SHOWNORMAL)
        If nHwnd = 0 Then nHwnd = GetDesktopWindow()
        LastProcessID = ShellExecute(nHwnd, "runas", App, vbNullString, vbNullString, WindowState)
        '  ShellExecute nHwnd, "runas", App, Command & " /admin", vbNullString, SW_SHOWNORMAL
    End Sub
    Public Function GetShortName(ByVal sLongFileName As String) As String
        Dim lRetVal As Integer, sShortPathName As String, iLen As Integer
        'Set up buffer area for API function call return

        sShortPathName = Space(255)
        iLen = Len(sShortPathName)

        'Call the function
        lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
        'Strip away unwanted characters.
        GetShortName = Left(sShortPathName, lRetVal)
    End Function

End Module

Imports System.Runtime.InteropServices
Imports stdole
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

    Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Integer
    '<DllImport("kernel32")>
    'Private Function GetVersionEx(ByRef osvi As OSVERSIONINFO) As Boolean
    'End Function

    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Integer, ByVal lpBuffer As String) As Integer
    'Public Declare Function LockWindowUpdate Lib "USER32" (ByVal hwnd As Integer) As Integer  -> This is for vb6.0.
    '<DllImport("user32.dll")>
    'Public Function LockWindowUpdate(ByVal hWndLock As IntPtr) As Boolean
    'End Function
    Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hWndLock As IntPtr) As Boolean
    Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Integer, ByVal nFolder As Integer, ByVal hToken As Integer, ByVal dwFlags As Integer, ByVal pszPath As String) As Integer
    Public Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As Integer) As Boolean
    End Function
    Public Declare Function GetCurrentProcessId Lib "kernel32" () As Integer
    Public Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Integer) As Integer
    'Public Declare Function GetActiveWindow Lib "USER32" () as integer  --> vb6.0 
    Public Declare Function GetActiveWindow Lib "USER32" () As IntPtr 'vb.net
    'Public Declare Function GetWindow Lib "USER32" (ByVal hwnd as integer, ByVal wCmd as integer) as integer vb6.0
    Public Declare Function GetWindow Lib "USER32" (ByVal hwnd As IntPtr, ByVal wCmd As Integer) As IntPtr  'vb.net
    Public Declare Function LoadImageAsString Lib "USER32" Alias "LoadImageA" (ByVal hinst As Integer, ByVal lpsz As String, ByVal uType As Integer, ByVal cxDesired As Integer, ByVal cyDesired As Integer, ByVal fuLoad As Integer) As Integer
    Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
    Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Integer, ByVal th32ProcessID As Integer) As Integer
    Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Integer, lppe As PROCESSENTRY32) As Integer
    Public Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal lHandle As Integer, ByRef lpStructure As PROCESS_MEMORY_COUNTERS, ByVal lSize As Integer) As Integer  '--> vb6.0
    'Public Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal lHandle As IntPtr, ByRef lpStructure As PROCESS_MEMORY_COUNTERS, ByVal lSize as integer) as integer ' --> vb.net
    Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Integer, lppe As PROCESSENTRY32) As Integer
    'Private Declare Function GetWindowWord Lib "USER32" (ByVal hwnd as integer, ByVal nIndex as integer) As Integer -vb6.0
    Private Declare Function GetWindowWord Lib "USER32" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer 'vb.net
    Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer
    Public Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer
    'Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd as integer, ByVal wMsg as integer, ByVal wParam as integer, lParam As Any) as integer - vb6.0
    'Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As IntPtr
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Function SendMessage(ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As IntPtr
    End Function

    'The below line is commented, because it is for vb6.0. Replaced with the next line for vb.net.
    'Public Declare Function SetWindowPos Lib "USER32" (ByVal hwnd as integer, ByVal hWndInsertAfter as integer, ByVal X as integer, ByVal Y as integer, ByVal cX as integer, ByVal cy as integer, ByVal wFlags as integer) as integer

    '<DllImport("user32.dll", SetLastError:=True)>
    'Private Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As SetWindowPosFlags) As Boolean
    'End Function

    'NOTE: THIS ENUM "SetWindowPosFlags" is used in the above commented SetWindowPos api.
    '<Flags>
    'Private Enum SetWindowPosFlags As UInteger
    '    ''' <summary>If the calling thread and the thread that owns the window are attached to different input queues,
    '    ''' the system posts the request to the thread that owns the window. This prevents the calling thread from
    '    ''' blocking its execution while other threads process the request.</summary>
    '    ''' <remarks>SWP_ASYNCWINDOWPOS</remarks>
    '    SynchronousWindowPosition = &H4000
    '    ''' <summary>Prevents generation of the WM_SYNCPAINT message.</summary>
    '    ''' <remarks>SWP_DEFERERASE</remarks>
    '    DeferErase = &H2000
    '    ''' <summary>Draws a frame (defined in the window's class description) around the window.</summary>
    '    ''' <remarks>SWP_DRAWFRAME</remarks>
    '    DrawFrame = &H20
    '    ''' <summary>Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to
    '    ''' the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE
    '    ''' is sent only when the window's size is being changed.</summary>
    '    ''' <remarks>SWP_FRAMECHANGED</remarks>
    '    FrameChanged = &H20
    '    ''' <summary>Hides the window.</summary>
    '    ''' <remarks>SWP_HIDEWINDOW</remarks>
    '    HideWindow = &H80
    '    ''' <summary>Does not activate the window. If this flag is not set, the window is activated and moved to the
    '    ''' top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter
    '    ''' parameter).</summary>
    '    ''' <remarks>SWP_NOACTIVATE</remarks>
    '    DoNotActivate = &H10
    '    ''' <summary>Discards the entire contents of the client area. If this flag is not specified, the valid
    '    ''' contents of the client area are saved and copied back into the client area after the window is sized or
    '    ''' repositioned.</summary>
    '    ''' <remarks>SWP_NOCOPYBITS</remarks>
    '    DoNotCopyBits = &H100
    '    ''' <summary>Retains the current position (ignores X and Y parameters).</summary>
    '    ''' <remarks>SWP_NOMOVE</remarks>
    '    IgnoreMove = &H2
    '    ''' <summary>Does not change the owner window's position in the Z order.</summary>
    '    ''' <remarks>SWP_NOOWNERZORDER</remarks>
    '    DoNotChangeOwnerZOrder = &H200
    '    ''' <summary>Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to
    '    ''' the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent
    '    ''' window uncovered as a result of the window being moved. When this flag is set, the application must
    '    ''' explicitly invalidate or redraw any parts of the window and parent window that need redrawing.</summary>
    '    ''' <remarks>SWP_NOREDRAW</remarks>
    '    DoNotRedraw = &H8
    '    ''' <summary>Same as the SWP_NOOWNERZORDER flag.</summary>
    '    ''' <remarks>SWP_NOREPOSITION</remarks>
    '    DoNotReposition = &H200
    '    ''' <summary>Prevents the window from receiving the WM_WINDOWPOSCHANGING message.</summary>
    '    ''' <remarks>SWP_NOSENDCHANGING</remarks>
    '    DoNotSendChangingEvent = &H400
    '    ''' <summary>Retains the current size (ignores the cx and cy parameters).</summary>
    '    ''' <remarks>SWP_NOSIZE</remarks>
    '    IgnoreResize = &H1
    '    ''' <summary>Retains the current Z order (ignores the hWndInsertAfter parameter).</summary>
    '    ''' <remarks>SWP_NOZORDER</remarks>
    '    IgnoreZOrder = &H4
    '    ''' <summary>Displays the window.</summary>
    '    ''' <remarks>SWP_SHOWWINDOW</remarks>
    '    ShowWindow = &H40
    'End Enum
    '    <DllImport(
    '"user32.dll",
    'CharSet:=CharSet.Auto,
    'CallingConvention:=CallingConvention.StdCall
    ')>

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
    Private Const GW_OWNER As Integer = 4
    Private Const SM_CXICON As Integer = 11
    Private Const SM_CYICON As Integer = 12
    Public Const TH32CS_SNAPPROCESS As Integer = &H2
    Public Const INVALID_HANDLE_VALUE As Integer = -1
    Public Const PROCESS_QUERY_LIMITED_INFORMATION As Integer = &H1000
    Public Const PROCESS_QUERY_INFORMATION As Integer = 1024
    Const GWW_HINSTANCE As Integer = (-6)

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

    Public Enum FolderEnum
        feCDBurnArea = 59               ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
        feCommonAppData = 35            ' \Docs & Settings\All Users\Application Data
        feCommonAdminTools = 47         ' \Docs & Settings\All Users\Start Menu\Programs\Administrative Tools
        feCommonDesktop = 25            ' \Docs & Settings\All Users\Desktop
        feCommonDocs = 46               ' \Docs & Settings\All Users\Documents
        feCommonPics = 54               ' \Docs & Settings\All Users\Documents\Pictures
        feCommonMusic = 53              ' \Docs & Settings\All Users\Documents\Music
        feCommonStartMenu = 22          ' \Docs & Settings\All Users\Start Menu
        feCommonStartMenuPrograms = 23  ' \Docs & Settings\All Users\Start Menu\Programs
        feCommonTemplates = 45          ' \Docs & Settings\All Users\Templates
        feCommonVideos = 55             ' \Docs & Settings\All Users\Documents\My Videos
        feLocalAppData = 28             ' \Docs & Settings\User\Local Settings\Application Data
        feLocalCDBurning = 59           ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
        feLocalHistory = 34             ' \Docs & Settings\User\Local Settings\History
        feLocalTempInternetFiles = 32   ' \Docs & Settings\User\Local Settings\Temporary Internet Files
        feProgramFiles = 38             ' \Program Files
        feProgramFilesCommon = 43       ' \Program Files\Common Files
        'feRecycleBin = 10               ' ???
        feUser = 40                     ' \Docs & Settings\User
        feUserAdminTools = 48           ' \Docs & Settings\User\Start Menu\Programs\Administrative Tools
        feUserAppData = 26              ' \Docs & Settings\User\Application Data
        feUserCache = 32                ' \Docs & Settings\User\Local Settings\Temporary Internet Files
        feUserCookies = 33              ' \Docs & Settings\User\Cookies
        feUserDesktop = 16              ' \Docs & Settings\User\Desktop
        feUserDocs = 5                  ' \Docs & Settings\User\My Documents
        feUserFavorites = 6             ' \Docs & Settings\User\Favorites
        feUserMusic = 13                ' \Docs & Settings\User\My Documents\My Music
        feUserNetHood = 19              ' \Docs & Settings\User\NetHood
        feUserPics = 39                 ' \Docs & Settings\User\My Documents\My Pictures
        feUserPrintHood = 27            ' \Docs & Settings\User\PrintHood
        feUserRecent = 8                ' \Docs & Settings\User\Recent
        feUserSendTo = 9                ' \Docs & Settings\User\SendTo
        feUserStartMenu = 11            ' \Docs & Settings\User\Start Menu
        feUserStartMenuPrograms = 2     ' \Docs & Settings\User\Start Menu\Programs
        feUserStartup = 7               ' \Docs & Settings\User\Start Menu\Programs\Startup
        feUserTemplates = 21            ' \Docs & Settings\User\Templates
        feUserVideos = 14               ' \Docs & Settings\User\My Documents\My Videos
        feWindows = 36                  ' \Windows
        feWindowsFonts = 20             ' \Windows\Fonts
        feWindowsResources = 56         ' \Windows\Resources
        feWindowsSystem = 37            ' \Windows\System32

        feWindowsSysWow64 = 111137            ' \Windows\SysWOW64
    End Enum

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

    Public Structure PROCESSENTRY32
        Dim dwSize As Integer
        Dim cntUsage As Integer
        Dim th32ProcessID As Integer
        Dim th32DefaultHeapID As Integer
        Dim th32ModuleID As Integer
        Dim cntThreads As Integer
        Dim th32ParentProcessID As Integer
        Dim pcPriClassBase As Integer
        Dim dwFlags As Integer
        <VBFixedString(260)> Dim szExeFile As String
    End Structure

    Public Structure PROCESS_MEMORY_COUNTERS
        Dim Cb As Integer
        Dim PageFaultCount As Integer
        Dim PeakWorkingSetSize As Integer
        Dim WorkingSetSize As Integer
        Dim QuotaPeakPagedPoolUsage As Integer
        Dim QuotaPagedPoolUsage As Integer
        Dim QuotaPeakNonPagedPoolUsage As Integer
        Dim QuotaNonPagedPoolUsage As Integer
        Dim PagefileUsage As Integer
        Dim PeakPagefileUsage As Integer
    End Structure

    Public Function GetLocalComputerName() As String
        Dim sBuffer As String
        Dim lReturn As Integer

        Try
            sBuffer = Space(255)
            lReturn = GetComputerName(sBuffer, Len(sBuffer))
            GetLocalComputerName = Trim(Left(sBuffer, InStr(sBuffer, vbNullChar) - 1))
        Catch ex As System.AccessViolationException
            GetLocalComputerName = ""
        Catch ex As System.Runtime.InteropServices.COMException
            GetLocalComputerName = ""
        Catch ex As Exception
            GetLocalComputerName = ""
        End Try
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
        'Dim PTB As PictureBox
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

        'Dim g As Graphics = PTB.CreateGraphics


        'N = GdiAlphaBlend(DC, L, T, W, H, PTB.hDC, 0, 0, 1, 1, lBlend)
        'N = GdiAlphaBlend(DC, L, T, W, H, g.GetHdc, 0, 0, 1, 1, lBlend)
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
        If Len(sFileName) = 0 Then EnsureFolderExists = UpdateFolder() : Exit Function

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
        Dim Result As VersionInformationType = Nothing

        VersionInformation = Nothing
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

    Public Function CaptureForm(ByRef frmSrc As Form) As Picture
        ' Call CaptureWindow to capture the entire form given its window
        ' handle and then return the resulting Picture object.
        'CaptureForm = CaptureWindow(frmSrc.hWnd, False, 0, 0,
        'CaptureForm = CaptureWindow(frmSrc.Handle, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
        Return Nothing
    End Function

    'Public Function CaptureWindow(ByVal hWndSrc as integer,
    '      ByVal Client As Boolean, ByVal LeftSrc as integer,
    '      ByVal TopSrc as integer, ByVal WidthSrc as integer,
    '      ByVal HeightSrc as integer) As Picture

    '    Dim hDCMemory as integer
    '    Dim hBmp as integer
    '    Dim hBmpPrev as integer
    '    Dim R as integer
    '    Dim hDCSrc as integer
    '    Dim hPal as integer
    '    Dim hPalPrev as integer
    '    Dim RasterCapsScrn as integer
    '    Dim HasPaletteScrn as integer
    '    Dim PaletteSizeScrn as integer
    '    Dim LogPal As LOGPALETTE

    '    ' Depending on the value of Client get the proper device context.
    '    If Client Then
    '        hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
    '    Else
    '        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire window.
    '    End If

    '    ' Create a memory device context for the copy process.
    '    hDCMemory = CreateCompatibleDC(hDCSrc)
    '    ' Create a bitmap and place it in the memory DC.
    '    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    '    hBmpPrev = SelectObject(hDCMemory, hBmp)

    '    ' Get screen properties.
    '    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities
    '    HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette support
    '    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of pallet

    '    ' If the screen has a palette make a copy and realize it.
    '    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    '        ' Create a copy of the system palette.
    '        LogPal.palVersion = &H300
    '        LogPal.palNumEntries = 256
    '        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
    '        hPal = CreatePalette(LogPal)
    '        ' Select the new palette into the memory DC and realize it.
    '        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
    '        R = RealizePalette(hDCMemory)
    '    End If

    '    ' Copy the on-screen image into the memory DC.
    '    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc,
    '    LeftSrc, TopSrc, vbSrcCopy)

    '    ' Remove the new copy of the  on-screen image.
    '    hBmp = SelectObject(hDCMemory, hBmpPrev)

    '    ' If the screen has a palette get back the palette that was
    '    ' selected in previously.
    '    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    '        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    '    End If

    '    ' Release the device context resources back to the system.
    '    R = DeleteDC(hDCMemory)
    '    R = ReleaseDC(hWndSrc, hDCSrc)

    '    ' Call CreateBitmapPicture to create a picture object from the
    '    ' bitmap and palette handles. Then return the resulting picture
    '    ' object.
    '    CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    'End Function

    Public Function GetTempDir() As String
        Dim Item As String, Size As Integer, Pos As Integer, Tmp As String
        Tmp = Space(256)
        Size = Len(Tmp)
        GetTempPath(Size, Tmp)
        Pos = InStr(Tmp, Chr(0))
        If Pos Then
            GetTempDir = Left(Tmp, Pos - 1)
        Else
            GetTempDir = Tmp
        End If
    End Function

    Public Function GetWinVerMajor() As Integer
        Dim OSV As OSVERSIONINFO
        Dim R As Integer
        Dim Pos As Integer
        Dim sVer As String
        Dim sBuild As String

        Try
            OSV.OSVSize = Len(OSV)
            If GetVersionEx(OSV) = 1 Then
                If OSV.PlatformID <> VER_PLATFORM_WIN32_NT Then Exit Function
                GetWinVerMajor = OSV.dwVerMajor
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function GetWindowsDir(Optional ByVal AddTrailingDirSep As Boolean = False) As String
        Dim Buffer As String, RET As Integer, X As Integer
        Buffer = Space(255)
        RET = GetWindowsDirectory(Buffer, 255)
        X = InStr(Buffer, Chr(0))
        If X > 0 Then Buffer = Left(Buffer, X - 1)
        GetWindowsDir = Buffer

        If AddTrailingDirSep Then GetWindowsDir = GetWindowsDir & DIRSEP
    End Function

    Public Function GetWindowsSystemDir(Optional ByVal No64 As Boolean = False) As String
        Dim Buffer As String, RET As Integer, X As Integer
        Buffer = Space(255)
        RET = GetSystemDirectory(Buffer, 255)
        X = InStr(Buffer, Chr(0))
        If X > 0 Then Buffer = Left(Buffer, X - 1)
        GetWindowsSystemDir = Buffer
        If No64 Then Exit Function
        Dim T As String
        T = ParentDirectory(GetWindowsSystemDir) & "SysWOW64"
        If DirExists(T) Then GetWindowsSystemDir = T
    End Function

    Public Function SpecialFolder(ByRef pFe As FolderEnum) As String
        Const MAX_PATH = 260
        Dim strPath As String
        Dim strBuffer As String

        strBuffer = Space(MAX_PATH)
        If SHGetFolderPath(0, pFe, 0, 0, strBuffer) = 0 Then strPath = Left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
        If Right(strPath, 1) = "\" Then strPath = Left(strPath, Len(strPath) - 1)
        SpecialFolder = strPath
    End Function

    Public Sub SetAlwaysOnTop(ByRef frm As Form, Optional ByRef OnTop As Boolean = True)
        If OnTop Then
            'SetWindowPos(frm.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2)
            SetWindowPos(frm.Handle, New IntPtr(-1), 0, 0, 0, 0, &H1 Or &H2)
        Else
            'SetWindowPos(frm.hwnd, -2, 0, 0, 0, 0, &H1 Or &H2)
            SetWindowPos(frm.Handle, New IntPtr(-2), 0, 0, 0, 0, &H1 Or &H2)
        End If
    End Sub

    Public Function SessionIsRemote() As Boolean 'BFH20081013
        ' this is used for detecting whether the program is running under a terminal services environment (remote desktop)
        ' if this causes problems, simply comment out the contents of this function
        On Error Resume Next
        Const SM_REMOTESESSION As Integer = &H1000
        SessionIsRemote = GetSystemMetrics(SM_REMOTESESSION) <> 0
    End Function

    Public Function fActiveForm() As Form
        'Dim X as integer, L As Variant
        Dim X As IntPtr, L As Form

        On Error Resume Next
        X = GetActiveWindow()
        'For Each L In Forms
        For Each L In My.Application.OpenForms
            'If L.hWnd = X Then
            If L.Handle = X Then
                fActiveForm = L
                Exit Function
            End If
        Next
    End Function

    Public Sub SetAppIcon()
        'SetIcon MainMenu.hWnd, "AAA", True
        SetIcon(MainMenu.Handle, "AAA", True)
    End Sub

    Private Sub SetIcon(ByVal hwnd As IntPtr, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
        Dim lhWndTop As IntPtr
        Dim lHwnd As IntPtr
        Dim cX As Integer
        Dim cY As Integer
        Dim hIconLarge As Integer
        Dim hIconSmall As Integer

        If (bSetAsAppIcon) Then ' Find VB's hidden parent window:
            lHwnd = hwnd
            lhWndTop = lHwnd
            Do While Not (lHwnd = 0)
                lHwnd = GetWindow(lHwnd, GW_OWNER)
                If Not (lHwnd = 0) Then
                    lhWndTop = lHwnd
                End If
            Loop
        End If

        cX = GetSystemMetrics(SM_CXICON)
        cY = GetSystemMetrics(SM_CYICON)

        '------------------------------------------COMMENTED BELOW CODE BECAUSE LOADIMAGEASSTRING(APP.HINSTANCE) PARAMETER IS NOT SUPPORTING IN VB.NET ---------------
        'hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cX, cY, LR_SHARED)

        '      If (bSetAsAppIcon) Then
        '          SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
        '     End If

        '      SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge

        'cX = GetSystemMetrics(SM_CXSMICON)
        '      cY = GetSystemMetrics(SM_CYSMICON)
        '      hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cX, cY, LR_SHARED)

        '      If (bSetAsAppIcon) Then
        '          SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
        'End If

        '      SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
    End Sub

    Public Function GetSystemUserName() As String
        Dim Buffer As String, RET As Integer, X As Integer
        Buffer = Space(255)
        RET = GetUserName(Buffer, 255)
        X = InStr(Buffer, Chr(0))
        If X > 0 Then Buffer = Left(Buffer, X - 1)
        GetSystemUserName = Buffer
    End Function

    Public Function GetDirectoryUserName() As String
        GetDirectoryUserName = SplitWord(UserFolder, -2, "\")
    End Function

    Public Function FileVersion(ByVal FileName As String, Optional ByRef A As Integer = 0, Optional ByRef B As Integer = 0, Optional ByRef C As Integer = 0, Optional ByRef D As Integer = 0) As String
        Dim T As Object
        On Error Resume Next
        FileVersion = VersionInformation(FileName).FileVersion
        T = Split(FileVersion, ".")
        A = Val(T(0))
        B = Val(T(1))
        C = Val(T(2))
        D = Val(T(3))
    End Function

    Public Function CurrentEXEDirectory(Optional ByVal DoUCase As Boolean = True, Optional ByVal DoShort As Boolean = True) As String
        On Error Resume Next
        Dim S As String, A As Integer
        S = CurrentEXEFileName()
        If S = "" Then Exit Function
        A = InStrRev(S, "\")
        If A = 0 Then Exit Function
        S = Left(S, A - 1)

        If DoShort Then S = GetShortName(S)
        S = CleanDir(S)
        If DoUCase Then S = UCase(S)

        CurrentEXEDirectory = S
    End Function

    Public Sub RestartComputer()
        ' Can't get API call to work, so we'll go with the internet suggested fix.
        '  ExitWindowsEx EWX_REBOOT, 0
        ShellOut.ShellOut("shutdown.exe -r -f -t 0")
    End Sub

    Public Sub EnableRichCHMContent()
        On Error Resume Next
        Const Section1 As String = "SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
        Const Section2 As String = "SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
        Const KeyName As String = "hh.exe"
        SaveRegistrySetting(HKEYS.regHKLM, Section1, KeyName, 9999, REG_TYPE.vtDWord)
        SaveRegistrySetting(HKEYS.regHKLM, Section2, KeyName, 9999, REG_TYPE.vtDWord)
    End Sub

    Public Function CurrentEXEFileName(Optional ByVal DoUCase As Boolean = True, Optional ByVal DoShort As Boolean = True) As String
        On Error Resume Next
        'Dim ModuleName As String * 128 '@NO-LINT-NTYP
        Dim ModuleName As String

        Dim FileName As String, hinst As Integer
        'create a buffer
        ModuleName = New String(Chr(0), 128)
        'get the hInstance application:
        'hinst = GetWindowWord(App.hInstance, GWW_HINSTANCE)
        hinst = GetWindowWord(Process.GetCurrentProcess.Handle, GWW_HINSTANCE)
        GetModuleFileName(hinst, ModuleName, Len(ModuleName))

        CurrentEXEFileName = ModuleName

        If DoShort Then CurrentEXEFileName = GetShortName(CurrentEXEFileName)
        If DoUCase Then CurrentEXEFileName = UCase(CurrentEXEFileName)
    End Function

    Public Function TicksElapsed(ByRef Ref As Integer, ByVal Limit As Integer) As Boolean
        If Ref = 0 Then Ref = GetTickCount
        TicksElapsed = GetTickCount < Ref + Limit
    End Function

    Public Function TicksSecondsRemaining(ByRef Ref As Integer, ByVal Limit As Integer) As Integer
        TicksSecondsRemaining = (Limit - (GetTickCount - Ref)) / 1000 + 1
    End Function

    Public Sub RunShellExecuteAdminArgs(ByVal App As String, ByVal Args As String, Optional ByVal cDir As String = "", Optional ByVal nHwnd As Integer = 0, Optional ByVal WindowState As Integer = SW_SHOWNORMAL)
        If nHwnd = 0 Then nHwnd = GetDesktopWindow()
        LastProcessID = ShellExecute(nHwnd, "runas", App, Args, cDir, WindowState)
        '  ShellExecute nHwnd, "runas", App, Command & " /admin", vbNullString, SW_SHOWNORMAL
    End Sub

    Public Function TicksNotElapsed(ByRef Ref As Integer, ByVal Limit As Integer) As Boolean
        TicksNotElapsed = Not TicksElapsed(Ref, Limit)
    End Function
End Module

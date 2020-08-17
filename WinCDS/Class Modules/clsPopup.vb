Imports System.Runtime.InteropServices
Public Class clsPopup
    Dim hMenu As Integer
    Private colHandles As Collection
    Public Event ItemClick(ByVal sItemKey As String)
    Private Declare Function CreatePopupMenu Lib "USER32" () As Integer
    Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Integer
    'Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Integer, ByVal uItem As Integer, ByVal fByPosition As Integer, lpmii As MENUITEMINFO) As Integer
    Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Integer, ByVal un As Integer, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Integer
    Private Declare Function TrackPopupMenu Lib "USER32" (ByVal hMenu As Integer, ByVal wFlags As enumTrackPopupMenu, ByVal X As Integer, ByVal Y As Integer, ByVal nReserved As Integer, ByVal hWnd As IntPtr, lprc As RECT) As Integer

    'NOTE: ABOVE TrackPopupMenu is fo vb6.0. BELOW ONE IS FOR VB.NET. If the above one will not work, replace it with the below one.
    '<DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    'Friend Function TrackPopupMenu(ByVal hMenu As IntPtr, ByVal wFlags As Integer,
    'ByVal x As Integer, ByVal y As Integer, ByVal nReserved As Integer,
    'ByVal hWnd As IntPtr, ByVal ignored As IntPtr) As Boolean
    'End Function
    Public Enum enumMenuItemStates
        MFS_GRAYED = &H3
        MFS_DISABLED = MFS_GRAYED
        MFS_CHECKED = enumMenuFlags.MF_CHECKED
        MFS_HILITE = enumMenuFlags.MF_HILITE
        MFS_ENABLED = enumMenuFlags.MF_ENABLED
        MFS_UNCHECKED = enumMenuFlags.MF_UNCHECKED
        MFS_UNHILITE = enumMenuFlags.MF_UNHILITE
        MFS_DEFAULT = enumMenuFlags.MF_DEFAULT
        MFS_MASK = &H108B                   ' #if(WINVER >= 0x0500)
        MFS_HOTTRACKDRAWN = &H10000000      ' #if(WINVER >= 0x0500)
        MFS_CACHEDBMP = &H20000000          ' #if(WINVER >= 0x0500)
        MFS_BOTTOMGAPDROP = &H40000000      ' #if(WINVER >= 0x0500)
        MFS_TOPGAPDROP = &H80000000         ' #if(WINVER >= 0x0500)
        MFS_GAPDROP = &HC0000000            ' #if(WINVER >= 0x0500)
    End Enum
    Private Enum enumMenuFlags
        MF_APPEND = &H100&
        MF_BITMAP = &H4&
        MF_BYCOMMAND = &H0&
        MF_BYPOSITION = &H400&
        MF_CALLBACKS = &H8000000
        MF_CHANGE = &H80&
        MF_CHECKED = &H8&
        MF_CONV = &H40000000
        MF_DEFAULT = &H1000    ' #if(WINVER >= 0x0400)
        MF_DELETE = &H200&
        MF_DISABLED = &H2&
        MF_ENABLED = &H0&
        MF_END = &H80
        MF_ERRORS = &H10000000
        MF_GRAYED = &H1&
        MF_HELP = &H4000&
        MF_HILITE = &H80&
        MF_HSZ_INFO = &H1000000
        MF_INSERT = &H0&
        MF_LINKS = &H20000000
        MF_MASK = &HFF000000
        MF_MENUBARBREAK = &H20&
        MF_MENUBREAK = &H40&
        MF_MOUSESELECT = &H8000&
        MF_OWNERDRAW = &H100&
        MF_POPUP = &H10&
        MF_POSTMSGS = &H4000000
        MF_REMOVE = &H1000&
        MF_RIGHTJUSTIFY = &H4000   ' #if(WINVER >= 0x0400)
        MF_SENDMSGS = &H2000000
        MF_SEPARATOR = &H800&
        MF_STRING = &H0&
        MF_SYSMENU = &H2000&
        MF_UNCHECKED = &H0&
        MF_UNHILITE = &H0&
        MF_USECHECKBITMAPS = &H200&
    End Enum
    Private Enum enumTrackPopupMenu
        TPM_CENTERALIGN = &H4
        TPM_LEFTALIGN = &H0
        TPM_RIGHTALIGN = &H8
        TPM_BOTTOMALIGN = &H20
        TPM_TOPALIGN = &H0
        TPM_VCENTERALIGN = &H10
        TPM_NONOTIFY = &H80
        TPM_RETURNCMD = &H100
        TPM_LEFTBUTTON = &H0
        TPM_RIGHTBUTTON = &H2
    End Enum
    Public Enum enumMenuItemTypes
        MFT_STRING = enumMenuFlags.MF_STRING
        MFT_BITMAP = enumMenuFlags.MF_BITMAP
        MFT_MENUBARBREAK = enumMenuFlags.MF_MENUBARBREAK
        MFT_MENUBREAK = enumMenuFlags.MF_MENUBREAK
        MFT_OWNERDRAW = enumMenuFlags.MF_OWNERDRAW
        MFT_RADIOCHECK = &H200
        MFT_SEPARATOR = enumMenuFlags.MF_SEPARATOR
        MFT_RIGHTORDER = &H2000
        MFT_RIGHTJUSTIFY = enumMenuFlags.MF_RIGHTJUSTIFY
    End Enum
    Private Enum enumMenuItemInfoMembers
        MIIM_STATE = &H1
        MIIM_ID = &H2
        MIIM_SUBMENU = &H4
        MIIM_CHECKMARKS = &H8
        MIIM_TYPE = &H10
        MIIM_DATA = &H20
        MIIM_STRING = &H40
        MIIM_BITMAP = &H80
        MIIM_FTYPE = &H100
    End Enum

    Private Structure POINTAPI
        Dim X As Integer
        Dim Y As Integer
    End Structure

    Private Structure MENUITEMINFO
        Dim cbSize As Integer
        Dim fMask As enumMenuItemInfoMembers
        Dim fType As enumMenuItemTypes
        Dim fState As enumMenuItemStates
        Dim wID As Integer
        Dim hSubMenu As Integer
        Dim hbmpChecked As Integer
        Dim hbmpUnchecked As Integer
        Dim dwItemData As Integer
        Dim dwTypeData As String
        Dim cch As Integer
    End Structure
    Private lCurItem As Integer

    Public Sub New()
        hMenu = CreatePopupMenu()
        colHandles = New Collection
    End Sub

    Public Sub AddItem(ByVal sKey As String, Optional ByVal sCaption As String = "", Optional ByVal eType As enumMenuItemTypes = enumMenuItemTypes.MFT_STRING, Optional ByVal eState As enumMenuItemStates = enumMenuItemStates.MFS_ENABLED, Optional ByVal lItemData As Integer = 0)
        Dim newMenuItem As MENUITEMINFO
        Dim lHandle As Integer

        If sCaption = "" Then sCaption = sKey

        lCurItem = lCurItem + 1

        newMenuItem.cbSize = Len(newMenuItem)          ' The size of this structure.
        ' Which elements of the structure to use.
        newMenuItem.fMask = enumMenuItemInfoMembers.MIIM_STATE Or enumMenuItemInfoMembers.MIIM_ID Or enumMenuItemInfoMembers.MIIM_TYPE Or enumMenuItemInfoMembers.MIIM_DATA
        newMenuItem.fType = eType                      ' The type of item: a string.
        newMenuItem.fState = eState                    ' This item is currently enabled and is the default item.
        newMenuItem.dwItemData = lItemData             ' Set The ItemData
        newMenuItem.wID = lCurItem                     ' Assign this item an item identifier.
        newMenuItem.dwTypeData = sCaption              ' Display the following text for the item.
        ' We would set submenu to the handle of an existing
        ' popup to bind them together
        '.hSubMenu = ??
        newMenuItem.cch = Len(newMenuItem.dwTypeData)

        'lHandle = InsertMenuItem(hMenu, lCurItem - 1, 1, newMenuItem)
        lHandle = InsertMenuItem(hMenu, lCurItem - 1, True, newMenuItem)

        If lHandle <> 0 Then colHandles.Add(sKey, "h" & lCurItem)
    End Sub

    Public Function PopupMenu(Optional lHwnd As IntPtr = Nothing) As Integer  ', ByVal X As Single, ByVal Y As Single)
        Dim pt As POINTAPI
        Dim Rec As RECT
        'If lHwnd = 0 Then lHwnd = MainMenu.hWnd
        If IsNothing(lHwnd) Then lHwnd = MainMenu.Handle

        GetCursorPos(pt)

        pt.X = Screen.PrimaryScreen.Bounds.Width / 2
        pt.Y = Screen.PrimaryScreen.Bounds.Height / 2
        PopupMenu = TrackPopupMenu(hMenu, enumTrackPopupMenu.TPM_NONOTIFY Or enumTrackPopupMenu.TPM_RETURNCMD Or enumTrackPopupMenu.TPM_LEFTALIGN, pt.X, pt.Y, 0, lHwnd, Rec)
        If PopupMenu <> 0 Then RaiseEvent ItemClick(colHandles(PopupMenu))
    End Function

End Class

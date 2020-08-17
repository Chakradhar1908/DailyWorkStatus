Imports SSubTimer6
Public Class cRegHotKey
    Private m_hWnd As IntPtr
    Private Const WM_DESTROY As Integer = &H2&
    Private Const WM_HOTKEY As Integer = &H312&
    Private m_iAtomCount As Integer
    Private m_tAtoms() As tHotKeyInfo
    Private Declare Function UnregisterHotKey Lib "USER32" (ByVal hwnd As IntPtr, ByVal ID As Integer) As Integer
    Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
    Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
    Private Declare Function RegisterHotKey Lib "USER32" (ByVal hwnd As IntPtr, ByVal ID As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Integer
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, lpSource As Object, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, Arguments As Integer) As Integer
    Private Const FORMAT_MESSAGE_FROM_SYSTEM As Integer = &H1000
    Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Integer = &H200
    Public Enum EHKModifiers
        MOD_ALT = &H1&
        MOD_CONTROL = &H2&
        MOD_SHIFT = &H4&
        MOD_WIN = &H8&
    End Enum
    Private Structure tHotKeyInfo
        Dim sName As String
        Dim sAtomName As String
        Dim lID As Integer
        Dim eKey As VBRUN.KeyCodeConstants
        Dim eModifiers As EHKModifiers
    End Structure

    Public Sub Attach(ByVal hwndA As IntPtr)
        Dim g As New GSubclass

        Clear()

        If (hwndA <> 0) Then
            m_hWnd = hwndA
            'GSubclass.AttachMessage(Me, m_hWnd, WM_HOTKEY)
            g.AttachMessage(Me, m_hWnd, WM_HOTKEY)
            'GSubclass.AttachMessage(Me, m_hWnd, WM_DESTROY)
            g.AttachMessage(Me, m_hWnd, WM_DESTROY)
        End If
    End Sub

    Public Sub Clear()
        Dim I As Integer
        Dim g As New GSubclass

        ' Remove all hot keys and atoms:
        For I = 1 To m_iAtomCount
            UnregisterKey(m_tAtoms(I).sName)
        Next
        ' Stop subclassing:
        If (m_hWnd <> 0) Then
            g.DetachMessage(Me, m_hWnd, WM_HOTKEY)
            g.DetachMessage(Me, m_hWnd, WM_DESTROY)
            m_hWnd = 0
        End If
    End Sub

    Public Sub UnregisterKey(ByVal sName As String)
        Dim lIndex As Integer
        Dim I As Integer
        lIndex = AtomIndex(sName)
        If (lIndex > 0) Then
            ' Unregister the key:
            UnregisterHotKey(m_hWnd, m_tAtoms(lIndex).lID)
            ' Unregister the atom:
            GlobalDeleteAtom(m_tAtoms(lIndex).lID)
            ' Remove from internal array:
            If (m_iAtomCount > 1) Then
                For I = lIndex To m_iAtomCount - 1
                    'LSet m_tAtoms(lIndex) = m_tAtoms(lIndex + 1)
                    m_tAtoms(lIndex) = m_tAtoms(lIndex + 1)
                Next
                m_iAtomCount = m_iAtomCount - 1
                'ReDim Preserve m_tAtoms(1 To m_iAtomCount) As tHotKeyInfo
                ReDim Preserve m_tAtoms(0 To m_iAtomCount - 1)
            Else
                m_iAtomCount = 0
                Erase m_tAtoms
            End If
        End If
    End Sub

    Private ReadOnly Property AtomIndex(ByVal sName As String) As Integer
        Get
            Dim I As Integer
            For I = 1 To m_iAtomCount
                If (m_tAtoms(I).sName = sName) Then
                    AtomIndex = I
                    Exit Property
                End If
            Next
            Err.Raise(vbObjectError + 1048 + 1, WinCDSEXEName() & ".cRegHotKey", "No hot key registered under the name '" & sName & "'")
        End Get
    End Property

    Public Sub RegisterKey(ByVal sName As String, ByVal eKey As VBRUN.KeyCodeConstants, ByVal eModifiers As EHKModifiers)
        Dim lID As Integer
        Dim lErr As Integer
        Dim lR As Integer
        Dim sError As String
        Dim sMsg As String
        Dim I As Integer
        Dim sAtomName As String

        ' Check for valid user name:
        If Len(sName) > 32 Then
            Err.Raise(vbObjectError + 1048 + 3, WinCDSEXEName() & ".cRegHotKey", "Key Name too long (max 32 characters).")
            Exit Sub
        Else
            For I = 1 To m_iAtomCount
                If (m_tAtoms(I).sName = sName) Then
                    Err.Raise(vbObjectError + 1048 + 4, WinCDSEXEName() & ".cRegHotKey", "The Key Name '" & sName & "' is already registered.")
                    Exit Sub
                End If
            Next
        End If

        ' Modify the user supplied name to get a more random system name:
        sAtomName = sName & "_" & WinCDSEXEName() & "_" & GetTickCount()
        If (Len(sAtomName) > 254) Then
            sAtomName = Left(sAtomName, 254)
        End If

        ' Create a new atom:
        lID = GlobalAddAtom(sAtomName)
        If (lID = 0) Then
            lErr = Err.LastDllError
            sError = WinError(lErr)
            sMsg = "Failed to add GlobalAtom"
            If (sError <> "") Then
                sMsg = sMsg & " [" & sError & "]"
            End If
            Err.Raise(vbObjectError + 1048 + 2, WinCDSEXEName() & ".cRegHotKey", sMsg)
        Else
            ' We have added the atom, now try to Register the
            ' key:
            lR = RegisterHotKey(m_hWnd, lID, eModifiers, eKey)
            If (lR = 0) Then
                lErr = Err.LastDllError
                ' Remove the atom:
                GlobalDeleteAtom(lID)
                ' Raise the error:
                WinError(lErr)
                sError = WinError(lErr)
                sMsg = "Failed to Register Hot Key"
                If (sError <> "") Then
                    sMsg = sMsg & " [" & sError & "]"
                End If
                Err.Raise(vbObjectError + 1048 + 3, WinCDSEXEName() & ".cRegHotKey", sMsg)
            Else
                ' Succeeded in adding the hot key:
                m_iAtomCount = m_iAtomCount + 1
                'ReDim Preserve m_tAtoms(1 To m_iAtomCount) As tHotKeyInfo
                ReDim Preserve m_tAtoms(0 To m_iAtomCount - 1)
                m_tAtoms(m_iAtomCount).sName = sName
                m_tAtoms(m_iAtomCount).sAtomName = sAtomName
                m_tAtoms(m_iAtomCount).lID = lID
                m_tAtoms(m_iAtomCount).eModifiers = eModifiers
                m_tAtoms(m_iAtomCount).eKey = eKey
            End If
        End If
    End Sub

    Private Function WinError(ByVal lLastDLLError As Integer) As String
        Dim sBuff As String
        Dim lCount As Integer

        ' Return the error message associated with LastDLLError:
        'sBuff = String(256, 0)
        sBuff = New String("0", 256)
        lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), 0)
        If lCount Then
            WinError = Left(sBuff, lCount)
        End If
    End Function
End Class

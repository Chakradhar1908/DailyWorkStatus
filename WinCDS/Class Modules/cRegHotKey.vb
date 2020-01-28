Imports SSubTimer6
Public Class cRegHotKey
    Private m_hWnd As Integer
    Private Const WM_DESTROY As Integer = &H2&
    Private Const WM_HOTKEY As Integer = &H312&
    Public Enum EHKModifiers
        MOD_ALT = &H1&
        MOD_CONTROL = &H2&
        MOD_SHIFT = &H4&
        MOD_WIN = &H8&
    End Enum

    Public Sub Attach(ByVal hwndA As Integer)
        Clear()

        If (hwndA <> 0) Then
            m_hWnd = hwndA
            GSubclass.AttachMessage(Me, m_hWnd, WM_HOTKEY)
            GSubclass.AttachMessage(Me, m_hWnd, WM_DESTROY)
        End If
    End Sub

    Public Sub Clear()
        Dim I As Long
        ' Remove all hot keys and atoms:
        For I = 1 To m_iAtomCount
            UnregisterKey m_tAtoms(I).sName
  Next
        ' Stop subclassing:
        If (m_hWnd <> 0) Then
            DetachMessage Me, m_hWnd, WM_HOTKEY
    DetachMessage Me, m_hWnd, WM_DESTROY
    m_hWnd = 0
        End If
    End Sub

    Public Sub RegisterKey(ByVal sName As String, ByVal eKey As KeyCodeConstants, ByVal eModifiers As EHKModifiers)
        Dim lID As Long
        Dim lErr As Long
        Dim lR As Long
        Dim sError As String
        Dim sMsg As String
        Dim I As Long
        Dim sAtomName As String

        ' Check for valid user name:
        If Len(sName) > 32 Then
            Err.Raise vbObjectError + 1048 + 3, WinCDSEXEName() & ".cRegHotKey", "Key Name too long (max 32 characters)."
    Exit Sub
        Else
            For I = 1 To m_iAtomCount
                If (m_tAtoms(I).sName = sName) Then
                    Err.Raise vbObjectError + 1048 + 4, WinCDSEXEName() & ".cRegHotKey", "The Key Name '" & sName & "' is already registered."
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
            Err.Raise vbObjectError + 1048 + 2, WinCDSEXEName() & ".cRegHotKey", sMsg
  Else
            ' We have added the atom, now try to Register the
            ' key:
            lR = RegisterHotKey(m_hWnd, lID, eModifiers, eKey)
            If (lR = 0) Then
                lErr = Err.LastDllError
                ' Remove the atom:
                GlobalDeleteAtom lID
         ' Raise the error:
                WinError lErr
      sError = WinError(lErr)
                sMsg = "Failed to Register Hot Key"
                If (sError <> "") Then
                    sMsg = sMsg & " [" & sError & "]"
                End If
                Err.Raise vbObjectError + 1048 + 3, WinCDSEXEName() & ".cRegHotKey", sMsg
    Else
                ' Succeeded in adding the hot key:
                m_iAtomCount = m_iAtomCount + 1
                ReDim Preserve m_tAtoms(1 To m_iAtomCount) As tHotKeyInfo
      m_tAtoms(m_iAtomCount).sName = sName
                m_tAtoms(m_iAtomCount).sAtomName = sAtomName
                m_tAtoms(m_iAtomCount).lID = lID
                m_tAtoms(m_iAtomCount).eModifiers = eModifiers
                m_tAtoms(m_iAtomCount).eKey = eKey
            End If

        End If
    End Sub


End Class

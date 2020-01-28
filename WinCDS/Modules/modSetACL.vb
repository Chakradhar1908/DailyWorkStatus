Module modSetACL
    Private Const ACLEXE As String = "SetACL.exe"
    Public Function ACLHasFullAccess(ByVal sDir As String) As Boolean
        '::::ACLHasFullAccess
        ':::SUMMARY
        ': Test Folder has FULL ACCESS set using SetACL.exe command
        ':::DESCRIPTION
        ': Executes the SetACL.exe command to determine if a given folder is set to FULL ACCESS for user EVERYONE
        ':::PARAMETERS
        ': - sDir - The folder to test access on
        ':::RETURN
        ': Boolean

        If Not ACLExists Then ACLHasFullAccess = True : Exit Function

        Dim O As String
        O = ACLList(sDir)
        ACLHasFullAccess = IsInStr(O, "Everyone,full,allow")
    End Function

    Public Function ACLList(ByVal sDir As String, Optional ByVal AsTable As Boolean = False) As String
        '::::ACLList
        ':::SUMMARY
        ': Access Control List
        ':::DESCRIPTION
        ': Return the ACL for the given Dir via external command execution.
        ':::PARAMETERS
        ': - sDir - Indicates the Directory where we get ACL files.
        ': - AsTable
        ':::RETURN
        ': String : Returns the ACL List string.
        On Error Resume Next
        Dim C As String, O As String, E As String

        If Not ACLExists Then ACLList = "" : Exit Function

        C = ACLLstCmd(sDir, AsTable)
        O = ShellOut.RunCmdToOutput(C, E)
        If E <> "" Then MessageBox.Show(E)
        '  Debug.Print O
        O = NLTrim(Replace(O, "SetACL finished successfully.", ""))
        ACLList = O
    End Function

    Private ReadOnly Property ACLLstCmd(ByVal File As String, Optional ByVal AsTable As Boolean = False, Optional ByVal ForDebug As Boolean = False) As String
        Get
            ACLLstCmd = ACLCmd & " -on " & File & " -ot file -actn list -lst ""i:y;f:" & IIf(AsTable, "tab", "csv") & """"
            If ForDebug Then ACLLstCmd = """" & Replace(ACLLstCmd, """", """""") & """"
        End Get
    End Property

    Private ReadOnly Property ACLCmd() As String
        Get
            If Not ACLExists Then Exit Property
            ACLCmd = ACLDir & ACLEXE
        End Get
    End Property

    Private ReadOnly Property ACLDir() As String
        Get
            ACLDir = GetWindowsDir() & "\"
        End Get
    End Property

    Public ReadOnly Property ACLExists() As Boolean
        Get
            ACLExists = FileExists(ACLDir & ACLEXE)
        End Get
    End Property

    Public Function ACL_FA(ByVal sDir As String) As String
        '::::ACL_FA
        ':::SUMMARY
        ': ACL Test macro for quick-display
        ':::DESCRIPTION
        ': This function is used to return FA if ACL has Full Access or !FA if not.
        ':::PARAMETERS
        ': - sDir - The folder to test
        ':::RETURN
        ': String : Returns ACL_FA as a string.

        On Error Resume Next
        If Not ACLExists() Then ACL_FA = "[-NoACL]" : Exit Function
        ACL_FA = IIf(ACLHasFullAccess(sDir), "[FA]", "[!FA]")
    End Function
End Module

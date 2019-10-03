Module modRegistry
    Private Const LocalConf As String = "WinCDS Local Configuration"
    Private bUseIni As TriState
    Public Const HKEY_LOCAL_MACHINE as integer = &H80000002
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey as integer, ByVal lpSubKey As String, ByVal ulOptions as integer, ByVal samDesired as integer, phkResult as integer) as integer
    Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey as integer, ByVal lpSubKey As String) as integer
    Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey as integer, ByVal lpValueName As String, ByVal lpReserved as integer, lpType as integer, ByVal lpData as integer, lpcbData as integer) as integer
    Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey as integer, ByVal lpValueName As String, ByVal lpReserved as integer, lpType as integer, lpData as integer, lpcbData as integer) as integer
    Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey as integer, ByVal lpValueName As String, ByVal Reserved as integer, ByVal dwType as integer, ByVal lpValue As String, ByVal cbData as integer) as integer
    Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey as integer, ByVal lpValueName As String, ByVal lpReserved as integer, lpType as integer, ByVal lpData As String, lpcbData as integer) as integer
    Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey as integer, ByVal lpSubKey As String, ByVal Reserved as integer, ByVal lpClass As String, ByVal dwOptions as integer, ByVal samDesired as integer, ByVal lpSecurityAttributes as integer, phkResult as integer, lpdwDisposition as integer) as integer
    Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey as integer, ByVal lpValueName As String, ByVal Reserved as integer, ByVal dwType as integer, lpValue as integer, ByVal cbData as integer) as Integer
    Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String) As Integer
    Public Const KEY_QUERY_VALUE As Integer = &H1
    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
    Public Const REG_SZ As Integer = REG_TYPE.vtString
    Public Const KEY_ALL_ACCESS As Integer = &H3F
    Public Const ERROR_NONE As Integer = 0
    Public Const REG_DWORD As Integer = REG_TYPE.vtDWord
    Public Const KEY_SET_VALUE As Integer = &H2
    Public Const REG_OPTION_NON_VOLATILE As Integer = 0

    Public Enum REG_TYPE
        vtNone = &H0                      ' No value type - REG_NONE
        vtString = &H1                    ' Nul terminated string - REG_SZ
        vtExpandString = &H2              ' Nul terminated string (with environment variable references) - REG_EXPAND_SZ
        vtBinary = &H3                    ' Free form binary - REG_BINARY
        vtDWord = &H4                     ' 32-bit number - REG_DWORD
        vtDWordBigEndian = &H5            ' 32-bit number. In big-endian format, the most significant byte of a word is the low-order byte - REG_DWORD_BIG_ENDIAN
        vtLink = &H6                      ' Symbolic Link (unicode) - REG_LINK
        vtMultiString = &H7               ' Multiple strings - REG_MULTI_SZ
        vtResourceList = &H8              ' Resource list in the resource map - REG_RESOURCE_LIST
        vtFullResourceDescriptor = &H9    ' Resource list in the hardware description
        vtResourceRequirementsList = &HA
    End Enum

    Public Function GetCDSSetting(ByVal vKEY As String, Optional ByVal Defaults As String = "", Optional ByVal SubSection As String = "", Optional ByVal ForceRegistry As Boolean = False) As String
        '::::GetCDSSetting
        ':::SUMMARY
        ': Get a WinCDS setting
        ':::DESCRIPTION
        ': Gets a WinCDS setting from the appropriate place in the registry or the CDS INI file (if it exists).
        ': - Located within HKLM (as opposed to HKCU) to make it available to to all users on the machine.
        ': - If the CDS.INI file exists, settings are pulled from there as opposed to registry (for windows security workaround)
        ':::PARAMETERS
        ': - vKEY
        ': - Default
        ': - SubSection
        ': - SubSection
        ': - ForceRegistry
        ':::RETURN
        ': String

        Dim X As String
        If UseINIFile() And Not ForceRegistry Then
            If SubSection = "" Then SubSection = LocalConf
            GetCDSSetting = ReadIniValue(CDSINI, SubSection, vKEY)
            If GetCDSSetting = "" And InStr(LCase(vKEY), "license") = 0 Then
                X = GetSystemSetting(RegistryAppName, RegistrySection & IIf(SubSection = LocalConf, "", "\" & SubSection), vKEY, Defaults)
                If X <> "" Then
                    GetCDSSetting = X
                    WriteIniValue(CDSINI, SubSection, vKEY, X)
                End If
            End If
            If GetCDSSetting = "" Then GetCDSSetting = Defaults
        Else
            GetCDSSetting = GetSystemSetting(RegistryAppName, RegistrySection & IIf(SubSection = "", "", "\" & SubSection), vKEY, Defaults)
        End If
    End Function

    Private Function UseINIFile() As Boolean
        Const Fld = "TestWrite"

        UseINIFile = False
        If bUseIni <> vbFalse Then UseINIFile = IIf(bUseIni = vbTrue, True, False) : Exit Function

        If FileExists(CDSINI) Then
            '    ActiveLog "Settings::UseINIFile:  YES - EXISTS", 8
            UseINIFile = True
            bUseIni = vbTrue
            Exit Function
        End If

        SaveSystemSetting(RegistryAppName, RegistrySection, Fld, "1")
        If GetSystemSetting(RegistryAppName, RegistrySection, Fld) = "1" Then
            DeleteSystemSetting(RegistryAppName, RegistrySection, Fld)
            bUseIni = vbUseDefault
            Exit Function
        End If

        '  ActiveLog "Settings::UseINIFile:  YES - WILL CREATE", 8
        UseINIFile = True
        bUseIni = vbTrue
    End Function
    Public Function CDSINI() As String
        '::::CDSINI
        ':::SUMMARY
        ': WinCDS.ini File
        ':::DESCRIPTION
        ': Returns filename of WinCDS.ini file.
        ':
        ': Used to store global software settings as an alternative to the registry (since Windows Vista restricted it).
        ':::RETURN
        ': String
        Const cINI As String = "WinCDS.ini"
        'Dim Old As String, Contents As String

        ' New Location is in our data folder (local, because these are per-computer settings)
        CDSINI = LocalCDSDataFolder() & cINI
        If FileExists(CDSINI) Then Exit Function  ' Base Case...

        ''BFH20170727 - We are discontinuing upgrading WinCDS.ini from AppFolder to CDSData
        '' It's been in long enough, and the new install procedure had CDSData set as the working directory for the
        '' App Shortcut on the desktop.  This meant the software was getting stuck in here trying to move the
        '' ini file onto itself.
        '
        '' This will upgrade it to the new location if needed
        '  Old = AppFolder & cINI
        'On Error Resume Next
        '  If LCase(GetShortName(Old)) <> LCase(GetShortName(CDSINI)) Then
        '    If Not FileExists(CDSINI) Then            ' If the correct file does not exist....
        '      If FileExists(Old) Then                   ' and it previously was in AppFolder, move it..
        '        Contents = ReadEntireFile(Old)
        '        WriteFile CDSINI, Contents, True        ' Set overwrite, even though we know it doesn't exist (base case)
        '        DeleteFileIfExists Old                  ' We will at least try to delete the old.  If it fails, no issue.
        '      End If
        '    End If
        '  End If

        ' BFH20160624 - Used to be in Program Files...  Can't do that.
        '  CDSINI = AppFolder & "WinCDS.ini"
    End Function
    Public Function GetSystemSetting(ByVal AppName As String, ByVal Section As String, ByVal vKEY As String, Optional ByVal Defaults As String = "") As String
        '::::GetSystemSetting
        ':::SUMMARY
        ': Get value from HKLM.
        ':::DESCRIPTION
        ': Returns a value from the HKLM entry in registry.  Returns default value if key is not set.
        ':::PARAMETERS
        ': - AppName
        ': - Section
        ': - vKEY
        ': - Default
        ':::RETURN
        ': String
        GetSystemSetting = QueryValue(HKEY_LOCAL_MACHINE, "Software\" & AppName & "\" & Section, vKEY)
        If GetSystemSetting = "" Then GetSystemSetting = Defaults
    End Function

    Public Sub DeleteSystemSetting(ByVal AppName As String, ByVal Section As String, ByVal vKEY As String)
        '::::DeleteSystemSetting
        ':::SUMMARY
        ': Remove key from HKLM.
        ':::DESCRIPTION
        ': This function is used to clear all the given system settings previously.
        ':::PARAMETERS
        ': - AppName
        ': - Section
        ': - vKEY
        DeleteRegValue(HKEY_LOCAL_MACHINE, "Software\" & AppName & "\" & Section, vKEY)
    End Sub
    Public Function QueryValue(ByVal hKey As Integer, ByRef sKeyName As String, ByRef sValueName As String) As Object
        '::::QueryValue
        ':::SUMMARY
        ': Query Registry Value
        ':::DESCRIPTION
        ': Queries a value from registry
        ':::PARAMETERS
        ': - hKey
        ': - sKeyName
        ': - sValueName
        ':::RETURN
        ': Variant
        Dim lRetVal As Integer         'result of the API functions
        Dim vValue As Object = Nothing        'setting of queried value

        'lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, KEY_QUERY_VALUE, hKey)
        lRetVal = QueryValueEx(hKey, sValueName, vValue)
        'RegCloseKey(hKey)
        QueryValue = vValue
    End Function
    Public Sub SaveSystemSetting(ByVal AppName As String, ByVal Section As String, ByVal vKEY As String, ByVal Setting As String)
        '::::SaveSystemSetting
        ':::SUMMARY
        ': Save to HKLM.
        ':::DESCRIPTION
        ': SAve to the HKLM entry in registry.  Handles creating/opening key, saving, and closing.
        ':::PARAMETERS
        ': - AppName
        ': - Section
        ': - vKEY
        ': - Setting
        SetKeyValue(HKEY_LOCAL_MACHINE, "Software\" & AppName & "\" & Section, vKEY, Setting, REG_SZ)
    End Sub
    Public Sub DeleteRegValue(ByVal hKey As Integer, ByVal sSection As String, ByVal sKeyName As String)
        '::::DeleteRegValue
        ':::SUMMARY
        ': Delete registry value
        ':::DESCRIPTION
        ': This function is used to Delete the Registry Value.
        ':::PARAMETERS
        ': - hKey
        ': - sSection
        ': - sKeyName
        Dim lRetVal As Integer         'result of the API functions

        lRetVal = RegOpenKeyEx(hKey, sSection, 0, KEY_ALL_ACCESS, hKey)
        lRetVal = RegDeleteValue(hKey, sKeyName)
        RegCloseKey(hKey)
    End Sub
    Public Function QueryValueEx(ByVal lhKey As Integer, ByVal szValueName As String, ByRef vValue As Object) As Integer
        '::::QueryValueEx
        ':::SUMMARY
        ': Return Registry Value
        ':::DESCRIPTION
        ': Query a registry value.
        ':::PARAMETERS
        ': - lhKey
        ': - szValueName
        ': - vValue
        ':::RETURN
        ': Long
        Dim cch As Integer
        Dim lRc As Integer
        Dim lType As Integer
        Dim lValue As Integer
        Dim sValue As String
        On Error GoTo QueryValueExError

        ' Determine the size and type of data to be read
        'lRc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
        If lRc <> ERROR_NONE Then Error 5

        Select Case lType
            Case REG_SZ  ' For strings
                'sValue = String(cch, 0)
                sValue = New String("0"c, cch)
                lRc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                If lRc = ERROR_NONE Then
                    vValue = Left(sValue, cch - 1)
                Else
                    vValue = String.Empty
                End If

            Case REG_DWORD  ' For DWORDS
                lRc = RegQueryValueExLong(lhKey, szValueName, 0&, lType,
                                lValue, cch)
                If lRc = ERROR_NONE Then vValue = lValue
            Case Else
                'all other data types not supported
                lRc = -1
        End Select

QueryValueExExit:
        QueryValueEx = lRc
        Exit Function

QueryValueExError:
        Resume QueryValueExExit
    End Function
    Public Sub SetKeyValue(ByVal hKey As Integer, ByVal sKeyName As String, ByVal sValueName As String, ByVal vValueSetting As Object, ByVal lValueType As Integer)
        '::::SetKeyValue
        ':::SUMMARY
        ': Used to Set Key Value.
        ':::DESCRIPTION
        ': This function is used to Set Key Value according to our specified conditions,.
        ':::PARAMETERS
        ': - sKeyName
        ': - sValueName
        ': - vValueSetting
        ': - lValueType
        ':::RETURN

        Dim lRetVal As Integer         'result of the SetValueEx function
        Dim hKeyResult As Integer         'handle of open key

        'open the specified key
        'lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, KEY_SET_VALUE, hKeyResult)
        If lRetVal <> 0 Then
            CreateNewKey(sKeyName, hKey)
            lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, KEY_SET_VALUE, hKeyResult)
        End If
        lRetVal = SetValueEx(hKeyResult, sValueName, lValueType, vValueSetting)
        'RegCloseKey(hKeyResult)
    End Sub
    Private Sub CreateNewKey(ByRef sNewKeyName As String, ByRef lPredefinedKey As Integer)
        Dim hNewKey As Integer         'handle to the new key
        Dim lRetVal As Integer         'result of the RegCreateKeyEx function

        lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
        RegCloseKey(hNewKey)
    End Sub
    Public Function SetValueEx(ByVal hKey As Integer, ByRef sValueName As String, ByRef lType As Integer, ByRef vValue As Object) As Integer
        '::::SetValueEx
        ':::SUMMARY
        ': Set Registry Value EX
        ':::DESCRIPTION
        ': Set Long value in registry
        ':::PARAMETERS
        ': - hKey
        ': - sValueName
        ': - lType
        ': - vValue
        ':::RETURN
        Dim lValue As Integer
        Dim sValue As String

        SetValueEx = 0
        Select Case lType
            Case REG_SZ
                sValue = vValue & Chr(0)
                'SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
            Case REG_DWORD
                lValue = vValue
                SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select
        '  WinSysError SetValueEx
    End Function
    Public Function SaveCDSSetting(ByVal vKEY As String, ByVal Value As String, Optional ByVal SubSection As String = "", Optional ByVal ForceRegistry As Boolean = False) As String
        '::::SaveCDSSetting
        ':::SUMMARY
        ': Save a WinCDS Setting
        ':::DESCRIPTION
        ': Saves a setting to either the registry or the CDS INI file (if exists)
        ':::PARAMETERS
        ': - vKEY
        ': - Value
        ': - SubSection
        ': - SubSection
        ': - ForceRegistry
        ':::SEE ALSO
        ': - GetCDSSetting
        ':::RETURN
        ': String - Returns the Result as a String.

        If UseINIFile() And Not ForceRegistry Then
            If SubSection = "" Then SubSection = LocalConf
            WriteIniValue(CDSINI, SubSection, vKEY, Value)
        Else
            SaveSystemSetting(RegistryAppName, RegistrySection & IIf(SubSection = "", "", "\" & SubSection), vKEY, Value)
        End If
        SaveCDSSetting = Value
    End Function

    Public Function DeleteSystemKey(ByVal AppName As String, ByVal Section As String, ByVal vKEY As String) As String
        '::::DeleteSystemKey
        ':::SUMMARY
        ': Used to Delete System Key.
        ':::DESCRIPTION
        ': This function is used to Delete System Key,after deleting Registered key.
        ':::PARAMETERS
        ': - AppName
        ': - Section
        ': - vKEY
        ':::RETURN
        ': String
        DeleteRegKey(HKEY_LOCAL_MACHINE, "Software\" & AppName & "\" & Section, vKEY)
        DeleteSystemKey = ""
    End Function

    Public Sub DeleteRegKey(ByVal hKey As Integer, ByVal sSection As String, ByVal sKeyName As String)
        '::::DeleteRegKey
        ':::SUMMARY
        ': Delete a registry key.
        ':::DESCRIPTION
        ': Delete a registry key from anywhere
        ':::PARAMETERS
        ': - hKey
        ': - sSection
        ': - sKeyName
        Dim lRetVal As Integer         'result of the API functions

        lRetVal = RegOpenKeyEx(hKey, sSection, 0, KEY_ALL_ACCESS, hKey)
        lRetVal = RegDeleteKey(hKey, sKeyName)
        RegCloseKey(hKey)
    End Sub

End Module

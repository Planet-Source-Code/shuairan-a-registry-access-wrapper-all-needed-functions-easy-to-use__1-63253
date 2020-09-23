Attribute VB_Name = "modRegistry"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE INFO:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' module name:       modRegistry (modRegistry.bas)
' created by:        shuairan
'
' public functions/procedures:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  __________________________________________________________________________
'  RegWrite (RegKey), (ValueData) [,(ValueDataType)]
'  ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'   writes data at the specified registry location.
'   if the specified value does not exist, it will be created.
'   if the specified registry key does not exist, it will be created.
'
'   (RegKey)        specifies the registry location where the value data
'                   should be written to.
'                   must start with one of this rootkeys (a short version
'                   can be used):
'
'                   HKEY_CLASSES_ROOT     (HKCR)
'                   HKEY_CURRENT_USER     (HKCU)
'                   HKEY_LOCAL_MACHINE    (HKLM)
'                   HKEY_USERS            (HKUS)
'                   HKEY_PERFORMANCE_DATA (HKPD)
'                   HKEY_CURRENT_CONFIG   (HKCC)
'                   HKEY_DYN_DATA         (HKDD)
'
'                   note: which root key actually exists, depends on your
'                         Windows Version!
'                   note: to access the registry of a remote computer, you
'                         have to put the computers name and a leading "\\"
'                         at the start of (RegKey):
'                         "\\remoteMachine\HKEY_LOCAL_MACHINE\Software\..."
'
'   (ValueData)     specifies the data that should be written.
'
'                   note: if you set (ValueDataType) to REG_DWORD or
'                         REG_BINARY, (ValueData) must be a numeric value
'                         (LONG for REG_DWORD, INTEGER for REG_BINARY).
'
'   (ValueDataType) optional, defines the data type of the specified value.
'                   can be one of the following public constants:
'                   REG_SZ, REG_EXPAND_SZ, REG_BINARY or REG_DWORD
'                   if omitted (ValueDataType) is set automatically to REG_SZ.
'
'   examples: RegWrite "HKLM\Software\AppValues\Path", "C:\App\"
'             > creates the sub key "AppValues\" and the Value "Path", if they
'               do not exist, and sets "Path" to "C:\App\" as data type REG_SZ.
'
'             RegWrite "HKLM\Software\AppValues\Size", 1024, REG_DWORD
'             > sets/creates "Size" to 1024& as REG_DWORD.
'
'             RegWrite "HKLM\Software\AppValues\", "this is my app"
'             > sets the "AppValues\"-key's (Default)-value to "this is my app".
'
'  __________________________________________________________________________
'  RegCreateKey (RegKey)
'  ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'   Creates one (or more) key(s) (and its sub keys) at the specified loction.
'
'   (RegKey)        specifies the registry location where the key should be
'                   created.
'                   must start with one of this rootkeys (a short version
'                   can be used):
'
'                   HKEY_CLASSES_ROOT     (HKCR)
'                   HKEY_CURRENT_USER     (HKCU)
'                   HKEY_LOCAL_MACHINE    (HKLM)
'                   HKEY_USERS            (HKUS)
'                   HKEY_PERFORMANCE_DATA (HKPD)
'                   HKEY_CURRENT_CONFIG   (HKCC)
'                   HKEY_DYN_DATA         (HKDD)
'
'                   note: which root key actually exists, depends on your
'                         Windows Version!
'                   note: to access the registry of a remote computer, you
'                         have to put the computers name and a leading "\\"
'                         at the start of (RegKey):
'                         "\\remoteMachine\HKEY_LOCAL_MACHINE\Software\..."
'
'  __________________________________________________________________________
'  RegRead ( (RegKey) )
'  ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'   Reads and returns the data of the value at the specified registry
'   location.
'
'   note: the returned value is of the type VARIANT. its sub type depends
'         on the reg data type of the data to be read:
'         REG_SZ and REG_EXPAND_SZ return a string.
'         REG_DWORD and REG_BINARY return a numeric value of type LONG.
'
'   (RegKey)        specifies the registry location of the value which data
'                   should be read.
'                   must start with one of this rootkeys (a short version
'                   can be used):
'
'                   HKEY_CLASSES_ROOT     (HKCR)
'                   HKEY_CURRENT_USER     (HKCU)
'                   HKEY_LOCAL_MACHINE    (HKLM)
'                   HKEY_USERS            (HKUS)
'                   HKEY_PERFORMANCE_DATA (HKPD)
'                   HKEY_CURRENT_CONFIG   (HKCC)
'                   HKEY_DYN_DATA         (HKDD)
'
'                   note: which root key actually exists, depends on your
'                         Windows Version!
'                   note: to access the registry of a remote computer, you
'                         have to put the computers name and a leading "\\"
'                         at the start of (RegKey):
'                         "\\remoteMachine\HKEY_LOCAL_MACHINE\Software\..."
'
'   examples: RegRead "HKLM\Software\AppValues\Path"
'             > reads the data of the value "Path".
'
'             RegRead "HKLM\Software\AppValues\"
'             > reads the "AppValues\"-key's (Default)-value.
'
'  __________________________________________________________________________
'  RegDelete (RegKey) [,(DeleteKeyDefaultValue)]
'  ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'   deletes the specified key and all of its subkeys and values or a single
'   value at the specified location.
'
'   (RegKey)        specifies the registry key/value that should be deleted.
'                   must start with one of this rootkeys (a short version
'                   can be used):
'
'                   HKEY_CLASSES_ROOT     (HKCR)
'                   HKEY_CURRENT_USER     (HKCU)
'                   HKEY_LOCAL_MACHINE    (HKLM)
'                   HKEY_USERS            (HKUS)
'                   HKEY_PERFORMANCE_DATA (HKPD)
'                   HKEY_CURRENT_CONFIG   (HKCC)
'                   HKEY_DYN_DATA         (HKDD)
'
'                   note: which root key actually exists, depends on your
'                         Windows Version!
'                   note: to access the registry of a remote computer, you
'                         have to put the computers name and a leading "\\"
'                         at the start of (RegKey):
'                         "\\remoteMachine\HKEY_LOCAL_MACHINE\Software\..."
'
'   (DeleteKeyDefaultValue) optional, defines if the key's (Default)-value or
'                           the key itself should be deleted.
'                   can be one of the following public constants:
'
'                   KEY_DELETE_KEY           (if the key should be deleted)
'                   KEY_DELETE_DEFAULT_VALUE (if the key's (Default)-value
'                                             should be deleted)
'
'                   if omitted (DeleteKeyDefaultValue) is automatically set
'                   to KEY_DELETE_KEY
'
'   examples: RegDelete "HKLM\Software\AppValues\"
'             > deletes the key "AppValues\".
'
'             RegDelete "HKLM\Software\AppValues\", KEY_DELETE_DEFAULT_VALUE
'             > deletes the "AppValues\"-key's (Default)-value.
'
'             RegDelete "HKLM\Software\AppValues\Path"
'             > deletes the value "Path"
'
'  __________________________________________________________________________
'  RegEnumKeyItems ( (RegKey) [,(Options)] )
'  ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'   Returns the names of all subkeys and/or all values of a registry key in
'   an array.
'
'   note: the returned value is of the type VARIANT. if the key has no subkeys
'         and/or values, the returned value will be 0 otherwise it will be a
'         array containing all subkeys and/or all values of the key.
'
'   note: if the key's (Default)-value will be enumerated, it will be returned
'         as a null-string ("") in the array ((Default)-values have no name).
'
'   (RegKey)        specifies the registry location of the key which
'                   subkeys and/or values should be enumerated.
'                   must start with one of this rootkeys (a short version
'                   can be used):
'
'                   HKEY_CLASSES_ROOT     (HKCR)
'                   HKEY_CURRENT_USER     (HKCU)
'                   HKEY_LOCAL_MACHINE    (HKLM)
'                   HKEY_USERS            (HKUS)
'                   HKEY_PERFORMANCE_DATA (HKPD)
'                   HKEY_CURRENT_CONFIG   (HKCC)
'                   HKEY_DYN_DATA         (HKDD)
'
'                   note: which root key actually exists, depends on your
'                         Windows Version!
'                   note: to access the registry of a remote computer, you
'                         have to put the computers name and a leading "\\"
'                         at the start of (RegKey):
'                         "\\remoteMachine\HKEY_LOCAL_MACHINE\Software\..."
'
'   (Options)       optional, defines what key itmes should be enumerated
'                   and if the array should be sorted.
'                   allowed options are the following public constants:
'
'                   KEY_ENUM_SUBKEYS (only enumerate the subkeys of (RegKey))
'                   KEY_ENUM_VALUES  (only enumerate the values of (RegKey))
'                   KEY_ENUM_DEFAULT (enumerate the key's (Default)-value, too
'                                     only allowed, if KEY_ENUM_VALUES is
'                                     given)
'                   KEY_ENUM_ALL     (enumerate the subkeys, values and the
'                                     key's (Default)-Value)
'                   KEY_ENUM_SORT    (sort the returned array)
'
'                   if omitted or if (Options) is only KEY_ENUM_SORT or only
'                   KEY_ENUM_DEFAULT, (Options) will be set automatically to
'                   (Options) + KEY_ENUM_SUBKEYS + KEY_ENUM_VALUES.
'
'   examples: RegEnumKeyItems "HKLM\Software\AppValues\"
'             > enumerates and returns all subkeys and values of "AppValues\"
'               without the "AppValues\"-key's (Default)-value.
'
'             RegEnumKeyItems "HKLM\Software\AppValues\", KEY_ENUM_VALUES + KEY_ENUM_DEFAULT
'             > enumerates all values and the "AppValues\"-key's
'               (Default)-value. if there is a (Default)-value, it will be
'               saved as a null-string ("") in the array.
'
'             RegEnumKeyItems "HKLM\Software\AppValues\", KEY_ENUM_DEFAULT + KEY_ENUM_SORT
'             > enumerates all subkeys, values of "AppValues\" and the
'               "AppValues\"-key's (Default)-value and sorts the returned
'               array.
'
'             RegEnumKeyItems "HKLM\Software\AppValues\", KEY_ENUM_ALL + KEY_ENUM_SORT
'             > same as above
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WIN32 API FUNCTIONS DECLERATION                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' functions concerning registry access
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long   ' changed [lpData As Byte] to [ByVal lpData As Any] to support REG_SZ values
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PRIVATE CONSTANTS                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' key access constants
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

' root key constants
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

' error constants
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_MORE_DATA = 234           ' dderror - More data is available.
Private Const ERROR_NO_MORE_ITEMS = 259

' used for private Win32 API function RegCreateKeyEx
Private Const REG_OPTION_NON_VOLATILE = 0   ' key is preserved when system is rebooted

' used for private function RegKeyGetPart
Private Const KEY_GETROOT = 1     ' the root of a registry key
Private Const KEY_GETSUB = 2      ' all subkeys between the root key and the value
Private Const KEY_GETVALUE = 3    ' the value a key points to
Private Const KEY_GETREMOTE = 4   ' the PC name in a key for remote accessing registry keys


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PUBLIC CONSTANTS                                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' reg value data type constants, used for public procedure RegWrite
Public Const REG_SZ = 1          ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2   ' Unicode nul terminated string
Public Const REG_BINARY = 3      ' Free form binary
Public Const REG_DWORD = 4       ' 32-bit number (LONGINTEGER)

' used for public function RegEnumItems
Public Const KEY_ENUM_SUBKEYS = 1
Public Const KEY_ENUM_VALUES = 2
Public Const KEY_ENUM_DEFAULT = 4   ' enumerate the key's default value, too
Public Const KEY_ENUM_ALL = KEY_ENUM_SUBKEYS Or KEY_ENUM_VALUES Or KEY_ENUM_DEFAULT
Public Const KEY_ENUM_SORT = 8      ' sort the returned array

' used for public procedure RegDelete
Public Const KEY_DELETE_DEFAULT_VALUE = True   ' delete the default value of a key and not the key itself
Public Const KEY_DELETE_KEY = False            ' delete a key


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' USERDEFINED DATA TYPES                                                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' used for private Win32 API function RegCreateKeyEx, has to be NULL
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

' used for private Win32 API functions RegEnumKeyEx and RegQueryInfoKey,
' has to be NULL
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PRIVATE FUNCTIONS AND PROCEDURES                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function RegKeyGetPart(ByVal RegKey As String, ByVal Part As Long) As Variant
' returns the specified part (Part&) of the given registry key RegKey$
'
' Part& can be one of this public onstants:
' - KEY_GETREMOTE
' - KEY_GETROOT
' - KEY_GETSUB
' - KEY_GETVALUE
'
' if Part& is KEY_GETROOT, RegKeyGetPart will return one of these private
' numeric root key constants:
' - HKEY_CLASSES_ROOT
' - HKEY_CURRENT_USER
' - HKEY_LOCAL_MACHINE
' - HKEY_USERS
' - HKEY_PERFORMANCE_DATA
' - HKEY_CURRENT_CONFIG
' - HKEY_DYN_DATA
'
' otherwise RegKeyGetPart returns the corresponding string
'
' "\\remotePC\HKEY_LOCAL_MACHINE\Software\AppValues\Value"
'    ,------- ,----------------- ,------------------,----
'    |        |                  '- KEY_GETSUB      '- KEY_GETVALUE
'    |        '- KEY_GETROOT (returns a root key constant)
'    '- KEY_GETREMOTE

   Dim rootKey As String
   Dim remoteName As String

   Select Case Part
   Case KEY_GETROOT
      remoteName = RegKeyGetPart(RegKey, KEY_GETREMOTE)

      If remoteName = "" Then
         rootKey = UCase(Mid(RegKey, 1, InStr(RegKey, "\") - 1))
      Else
         rootKey = UCase(Mid(RegKey, Len(remoteName) + 4, _
                             InStr(Len(remoteName) + 4, RegKey, "\") - (Len(remoteName) + 4)))
      End If

      Select Case rootKey
      Case "HKCR", "HKEY_CLASSES_ROOT"
         RegKeyGetPart = HKEY_CLASSES_ROOT
      Case "HKCU", "HKEY_CURRENT_USER"
         RegKeyGetPart = HKEY_CURRENT_USER
      Case "HKLM", "HKEY_LOCAL_MACHINE"
         RegKeyGetPart = HKEY_LOCAL_MACHINE
      Case "HKUS", "HKEY_USERS"
         RegKeyGetPart = HKEY_USERS
      Case "HKPD", "HKEY_PERFORMANCE_DATA"
         RegKeyGetPart = HKEY_PERFORMANCE_DATA
      Case "HKCC", "HKEY_CURRENT_CONFIG"
         RegKeyGetPart = HKEY_CURRENT_CONFIG
      Case "HKDD", "HKEY_DYN_DATA"
         RegKeyGetPart = HKEY_DYN_DATA
      Case Else
         Err.Raise Number:=vbObjectError + 1, _
                   Description:="Invalid root in registry key '" & RegKey & "'"
         Exit Function
      End Select
   Case KEY_GETSUB
      remoteName = RegKeyGetPart(RegKey, KEY_GETREMOTE)

      If remoteName = "" Then
         RegKeyGetPart = Mid(RegKey, InStr(RegKey, "\") + 1, _
                             InStrRev(RegKey, "\") - _
                             Len(Mid(RegKey, 1, InStr(RegKey, "\"))))
      Else
         rootKey = Mid(RegKey, Len(remoteName) + 4, _
                       InStr(Len(remoteName) + 4, RegKey, "\") - (Len(remoteName) + 4))

         RegKeyGetPart = Mid(RegKey, Len(remoteName) + Len(rootKey) + 5, _
                             InStrRev(RegKey, "\") - _
                             (Len(remoteName) + Len(rootKey) + 4))
      End If
   Case KEY_GETVALUE
      RegKeyGetPart = Mid(RegKey, InStrRev(RegKey, "\") + 1)
   Case KEY_GETREMOTE
      If Left(RegKey, 2) = "\\" Then
         RegKeyGetPart = Mid(RegKey, 3, InStr(3, RegKey, "\") - 3)
      Else
         RegKeyGetPart = ""
      End If
   Case Else
      Err.Raise Number:=5
      Exit Function
   End Select
End Function

Private Function RegOpen(ByVal RegKey As String, Optional ByVal samDesired As Long = KEY_ALL_ACCESS) As Long
' opens the specified registry key RegKey$ for samDesired& actions and returns
' a handle to the key
' if samDesired& is omitted, the key will be acessed with full rights (KEY_ALL_ACCESS)

   Dim retVal As Long
   Dim remoteName As String
   Dim rootKey As Long
   Dim subKey As String
   Dim hKey As Long
   Dim rhKey As Long

   remoteName = RegKeyGetPart(RegKey, KEY_GETREMOTE)
   rootKey = RegKeyGetPart(RegKey, KEY_GETROOT)
   subKey = RegKeyGetPart(RegKey, KEY_GETSUB)

   If remoteName = "" Then
      ' access registry key
      retVal = RegOpenKeyEx(rootKey, subKey, 0&, samDesired, hKey)

      If retVal <> ERROR_NONE Then
         Err.Raise Number:=vbObjectError + 2, _
                   Description:="Unable to open registry key '" & RegKey & "'"
         Exit Function
      End If
   Else
      ' establish connection to remote machine remoteName$
      retVal = RegConnectRegistry(remoteName, rootKey, rhKey)

      If retVal <> ERROR_NONE Then
         Err.Raise Number:=vbObjectError + 3, _
                   Description:="Unable to connect to remote computer '\\" & _
                                remoteName & "'"
         Exit Function
      End If

      ' access remote registry key
      retVal = RegOpenKeyEx(rhKey, subKey, 0&, samDesired, hKey)

      If retVal <> ERROR_NONE Then
         Err.Raise Number:=vbObjectError + 4, _
                   Description:="Unable to open remote registry key '" & _
                                RegKey & "'"
         Exit Function
      End If
   End If

   RegOpen = hKey
End Function

Private Function RegOpenCreate(ByVal RegKey As String, Optional ByVal samDesired As Long = KEY_ALL_ACCESS)
' opens the specified registry key RegKey$ for samDesired& actions and returns
' a handle to the key
' if the key does not not exist it will be created and the new key will be accessed for
' samDesired& actions
' if samDesired& is omitted, the opened/created key will be acessed with full rights (KEY_ALL_ACCESS)

   Dim retVal As Long
   Dim rootKey As Long
   Dim subKey As String
   Dim hKey As Long
   Dim hNewKey As Long
   Dim lpSecurityAttributes As SECURITY_ATTRIBUTES   ' is NULL
   Dim lpdwDisposition As Long   ' is NULL

   rootKey = RegKeyGetPart(RegKey, KEY_GETROOT)
   subKey = RegKeyGetPart(RegKey, KEY_GETSUB)

   retVal = RegCreateKeyEx(rootKey, subKey, 0&, vbNullString, _
                           REG_OPTION_NON_VOLATILE, samDesired, _
                           lpSecurityAttributes, hNewKey, lpdwDisposition)

   If retVal <> ERROR_NONE Then
      Err.Raise Number:=vbObjectError + 5, _
                Description:="Unable to create registry key '" & RegKey & "'"
      Exit Function
   End If

   RegOpenCreate = hNewKey
End Function

Private Sub RegClose(hKey As Long)
' closes a open registry key
   Dim retVal As Long
   retVal = RegCloseKey(hKey)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PUBLIC FUNCTIONS AND PROCEDURES                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function RegRead(ByVal RegKey As String) As Variant
' see documentation in module information

   Dim retVal As Long
   Dim hKey As Long
   Dim value As String
   Dim ValueDataType As Long
   Dim valueDataSize As Long
   Dim valueDataStr As String
   Dim valueDataLng As Long

   hKey = RegOpen(RegKey, KEY_QUERY_VALUE)
   value = RegKeyGetPart(RegKey, KEY_GETVALUE)

   ' determine the size (valueDataSize&) and type (valueDataType&) of the
   ' data to be read
   retVal = RegQueryValueEx(hKey, value, 0&, ValueDataType, ByVal 0&, _
                            valueDataSize)

   If retVal <> ERROR_NONE Then
      Err.Raise Number:=vbObjectError + 6, _
                Description:="Unable to read registry value '" & RegKey & "'"
      Exit Function
   End If

   If ValueDataType = REG_SZ Or ValueDataType = REG_EXPAND_SZ Then
      valueDataStr = String(valueDataSize, vbNullChar)

      retVal = RegQueryValueEx(hKey, value, 0&, 0&, ByVal valueDataStr, _
                               valueDataSize)
   ElseIf ValueDataType = REG_BINARY Or ValueDataType = REG_DWORD Then
      retVal = RegQueryValueEx(hKey, value, 0&, 0&, valueDataLng, _
                               valueDataSize)
   Else
      Err.Raise Number:=vbObjectError + 7, _
                Description:="Unsupported data type in registry value '" & _
                             RegKey & "'"
      Exit Function
   End If

   If retVal <> ERROR_NONE Then
      Err.Raise Number:=vbObjectError + 6, _
                Description:="Unable to read registry value '" & RegKey & "'"
      Exit Function
   End If

   RegClose hKey

   If ValueDataType = REG_SZ Or ValueDataType = REG_EXPAND_SZ Then
      If valueDataSize = 0 Then
         RegRead = ""
      Else
         RegRead = Left(valueDataStr, valueDataSize - 1)
      End If
   ElseIf ValueDataType = REG_BINARY Or ValueDataType = REG_DWORD Then
      RegRead = valueDataLng
   End If
End Function

Public Sub RegWrite(ByVal RegKey As String, ByVal ValueData As Variant, Optional ByVal ValueDataType As Long = REG_SZ)
' see documentation in module information

   Dim retVal As Long
   Dim hKey As Long
   Dim value As String
   Dim valueDataStr As String
   Dim valueDataLng As Long
   Dim valueDataInt As Integer

   hKey = RegOpenCreate(RegKey, KEY_WRITE)
   value = RegKeyGetPart(RegKey, KEY_GETVALUE)

   Select Case ValueDataType
   Case REG_SZ, REG_EXPAND_SZ
      valueDataStr = CStr(ValueData) & vbNullChar
      retVal = RegSetValueEx(hKey, value, 0&, ValueDataType, _
                             ByVal valueDataStr, Len(valueDataStr))
   Case REG_BINARY
      valueDataInt = CInt(ValueData)
      retVal = RegSetValueEx(hKey, value, 0&, ValueDataType, valueDataInt, 2)
   Case REG_DWORD
      valueDataLng = CLng(ValueData)
      retVal = RegSetValueEx(hKey, value, 0&, ValueDataType, valueDataLng, 4)
   Case Else
      Err.Raise Number:=5
      Exit Sub
   End Select

   If retVal <> ERROR_NONE Then
      Err.Raise Number:=vbObjectError + 8, _
                Description:="Unable to write registry value '" & RegKey & "'"
      Exit Sub
   End If

   RegClose hKey
End Sub

Public Sub RegCreateKey(ByVal RegKey As String)
' see documentation in module information

   Dim hKey As Long

   If Right(RegKey, 1) <> "\" Then RegKey = RegKey & "\"

   hKey = RegOpenCreate(RegKey, KEY_WRITE)

   RegCloseKey hKey
End Sub

Public Function RegEnumKeyItems(ByVal RegKey As String, Optional ByVal Options As Byte = (KEY_ENUM_SUBKEYS Or KEY_ENUM_VALUES)) As Variant
' see documentation in module information

' Options is a bitwise parameter:
'
' 0 0 0 0 1 1 1 1
' ------- ^ ^ ^ ^
'   \ /   | | | '- bit 0: if this bit is set subkeys will be enumerated
'    |    | | '--- bit 1: values will be enumerated
'    |    | '----- bit 2: the key's default value will be enumerated
'    |    |               (set only if bit 1 is set, otherwise error 5
'    |    |               will be raised)
'    |    '------- bit 3: the returned array wil be sorted
'    '------------ bit 4-7: if one of this bits is set, error 5 will be raised

   Dim retVal As Long
   Dim hKey As Long
   Dim items() As String
   Dim n As Long
   Dim m As Long
   Dim temp As String
   Dim lpftLastWriteTime As FILETIME   ' is NULL

   Dim numSubKeys As Long
   Dim subKeyName As String
   Dim maxSubKeyNameLen As Long
   Dim curSubKeyNameLen As Long

   Dim numValues As Long
   Dim value As String
   Dim maxValueLen As Long
   Dim curValueLen As Long

   ' if optional parameter Options = KEY_ENUM_SORT
   ' (bit 2 is set and bit 0 and 1 are unset)
   If (Options And (KEY_ENUM_SUBKEYS Or KEY_ENUM_VALUES)) = 0 Then
      Options = Options Or KEY_ENUM_SUBKEYS Or KEY_ENUM_VALUES
   End If

   If (Options And (KEY_ENUM_ALL Or KEY_ENUM_SORT)) <> Options Or _
      (((Options And KEY_ENUM_DEFAULT) = KEY_ENUM_DEFAULT) And _
       ((Options And KEY_ENUM_VALUES) <> KEY_ENUM_VALUES)) Then
      Err.Raise Number:=5
      Exit Function
   End If

   If Right(RegKey, 1) <> "\" Then RegKey = RegKey & "\"

   ' open the key RegKey$ to query its sub keys and values
   hKey = RegOpen(RegKey, KEY_READ)

   ' get information about the sub keys anf values of the key RegKey$
   retVal = RegQueryInfoKey(hKey, vbNullChar, 0&, 0&, numSubKeys, _
                            maxSubKeyNameLen, 0&, numValues, maxValueLen, _
                            0&, 0&, lpftLastWriteTime)

   If retVal <> ERROR_NONE Then
      Err.Raise Number:=vbObjectError + 9, _
                Description:="Unable to read information from registry key '" & _
                             RegKey & "'"
      Exit Function
   End If

   If (Options And KEY_ENUM_SUBKEYS) = KEY_ENUM_SUBKEYS And numSubKeys Then
      numSubKeys = numSubKeys - 1   ' count should start with 0
      ReDim items(numSubKeys)

      ' enumerate the sub keys of the key RegKey$
      For n = 0 To numSubKeys
         ' save maxSubKeyNameLen&, because RegEnumKeyEx would change it
         curSubKeyNameLen = maxSubKeyNameLen + 1
         subKeyName = String(curSubKeyNameLen, vbNullChar)

         retVal = RegEnumKeyEx(hKey, n, subKeyName, curSubKeyNameLen, _
                               0&, vbNullChar, 0&, lpftLastWriteTime)

         If retVal <> ERROR_NONE And retVal <> ERROR_MORE_DATA Then
         ' to ignore the constantly returned "More Data available" error
            Err.Raise Number:=vbObjectError + 9, _
                      Description:="Unable to read information from registry key '" & _
                                   RegKey & "'"
            Exit Function
         End If

         items(n) = Left(subKeyName, curSubKeyNameLen) & "\"
      Next
   End If

   If (Options And KEY_ENUM_VALUES) = KEY_ENUM_VALUES And numValues Then
      ' change numSubKeys& to the actual number of saved sub key names
      numSubKeys = n
      numValues = numValues - 1   ' count should start with 0
      ReDim Preserve items(numSubKeys + numValues)

      ' enumerate the sub keys of the key RegKey$
      For n = 0 To numValues
         ' save maxValueLen&, because RegEnumValueEx would change it
         curValueLen = maxValueLen + 1
         value = String(curValueLen, vbNullChar)

         retVal = RegEnumValue(hKey, n, value, curValueLen, _
                               0&, 0&, ByVal 0&, 0&)

         If retVal <> ERROR_NONE Then
            Err.Raise Number:=vbObjectError + 9, _
                      Description:="Unable to read information from registry key '" & _
                                   RegKey & "'"
            Exit Function
         End If

         If (Options And KEY_ENUM_DEFAULT) = KEY_ENUM_DEFAULT Then
            items(numSubKeys + n) = Left(value, curValueLen)
         Else
            If Left(value, curValueLen) = "" Then
            ' the key's default value that is to be ignored
               ReDim Preserve items(UBound(items) - 1)
            Else
               items(numSubKeys + m) = Left(value, curValueLen)
               m = m + 1
            End If
         End If
      Next
   End If

   RegClose hKey

   If (Options And KEY_ENUM_SORT) = KEY_ENUM_SORT And n Then
      For n = 0 To UBound(items)
         For m = n To UBound(items)
            If items(n) > items(m) And _
               Right(items(n), 1) = "\" And _
               Right(items(m), 1) = "\" Then
            ' sort sub keys
               temp = items(n)
               items(n) = items(m)
               items(m) = temp
            ElseIf items(n) > items(m) And _
               Right(items(n), 1) <> "\" And _
               Right(items(m), 1) <> "\" Then
            ' sort values
               temp = items(n)
               items(n) = items(m)
               items(m) = temp
            End If
         Next
      Next
   End If

   If n Then
      RegEnumKeyItems = items()
   Else
      RegEnumKeyItems = 0&
   End If
End Function

Public Sub RegDelete(ByVal RegKey As String, Optional ByVal DeleteKeyDefaultValue As Boolean = KEY_DELETE_KEY)
' see documentation in module information

   Dim retVal As Long
   Dim value As String
   Dim hKey As Long
   Dim subKeys As Variant
   Dim n As Long
   Dim mainKey As String
   Dim subKey As String

   value = RegKeyGetPart(RegKey, KEY_GETVALUE)

   If value = "" And Not DeleteKeyDefaultValue Then
   ' delete the specified key and its subs keys
      ' enumerate the sub keys of the key RegKey$
      subKeys = RegEnumKeyItems(RegKey, KEY_ENUM_SUBKEYS)

      If IsArray(subKeys) Then
         ' delete the enumerated sub keys of the key RegKey$
         For n = 0 To UBound(subKeys)
            RegDelete (RegKey & subKeys(n))
         Next
      End If

      subKey = Mid(RegKey, InStrRev(RegKey, "\", Len(RegKey) - 1) + 1, _
                   InStrRev(RegKey, "\"))
      mainKey = Left(RegKey, Len(RegKey) - Len(subKey))

      ' open the key that RegKey$ is sub key to
      hKey = RegOpen(mainKey, KEY_WRITE)

      ' delete the key RegKey$
      retVal = RegDeleteKey(hKey, subKey)

      If retVal <> ERROR_NONE Then
         Err.Raise Number:=vbObjectError + 10, _
                   Description:="Unable to delete registry key '" & RegKey & "'"
         Exit Sub
      End If

      RegClose hKey
   Else
   ' delete specified value
      hKey = RegOpen(RegKey, KEY_WRITE)

      retVal = RegDeleteValue(hKey, value)

      If retVal <> ERROR_NONE Then
         If DeleteKeyDefaultValue Then
            Err.Raise Number:=vbObjectError + 11, _
                      Description:="Unable to delete registry key default value '" & _
                                   RegKey & "(Default)'"
         Else
            Err.Raise Number:=vbObjectError + 12, _
                      Description:="Unable to delete registry value '" & _
                                   RegKey & "'"
         End If

         Exit Sub
      End If

      RegClose hKey
   End If
End Sub
Attribute VB_Name = "modReg"

Option Explicit

   Public Const REG_SZ As Long = 1
   Public Const REG_DWORD As Long = 4

   Public Const HKEY_CLASSES_ROOT = &H80000000
   Public Const HKEY_CURRENT_USER = &H80000001
   Public Const HKEY_LOCAL_MACHINE = &H80000002
   Public Const HKEY_USERS = &H80000003

   Public Const ERROR_NONE = 0
   Public Const ERROR_BADDB = 1
   Public Const ERROR_BADKEY = 2
   Public Const ERROR_CANTOPEN = 3
   Public Const ERROR_CANTREAD = 4
   Public Const ERROR_CANTWRITE = 5
   Public Const ERROR_OUTOFMEMORY = 6
   Public Const ERROR_ARENA_TRASHED = 7
   Public Const ERROR_ACCESS_DENIED = 8
   Public Const ERROR_INVALID_PARAMETERS = 87
   Public Const ERROR_NO_MORE_ITEMS = 259

   Public Const KEY_QUERY_VALUE = &H1
   Public Const KEY_SET_VALUE = &H2
   Public Const KEY_ALL_ACCESS = &H3F

   Public Const REG_OPTION_NON_VOLATILE = 0

   Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
   As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long
   Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
   Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As _
   Long, lpcbData As Long) As Long
   Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
   Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
   String, ByVal cbData As Long) As Long
   Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
   ByVal cbData As Long) As Long
                

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, _
   lType As Long, vValue As Variant) As Long
             Dim lValue As Long
             Dim sValue As String
10           Select Case lType
                 Case REG_SZ
20                   sValue = vValue & Chr$(0)
30                   SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
                                                    lType, sValue, Len(sValue))
40               Case REG_DWORD
50                   lValue = vValue
60                   SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
         lType, lValue, 4)
70               End Select
   End Function

   Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
   String, vValue As Variant) As Long
             Dim cch As Long
             Dim lrc As Long
             Dim lType As Long
             Dim lValue As Long
             Dim sValue As String

10           On Error GoTo QueryValueExError

             ' Determine the size and type of data to be read
20           lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
30           If lrc <> ERROR_NONE Then Error 5

40           Select Case lType
                 ' For strings
                 Case REG_SZ:
50                   sValue = String(cch, 0)

60       lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
         sValue, cch)
70                   If lrc = ERROR_NONE Then
80                       vValue = Left$(sValue, cch - 1)
90                   Else
100                      vValue = Empty
110                  End If
                 ' For DWORDS
     Case REG_DWORD:
120      lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
         lValue, cch)
130                  If lrc = ERROR_NONE Then vValue = lValue
140              Case Else
                     'all other data types not supported
150                  lrc = -1
160          End Select

QueryValueExExit:
170          QueryValueEx = lrc
180          Exit Function

QueryValueExError:
190          Resume QueryValueExExit
   End Function


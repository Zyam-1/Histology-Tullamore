Attribute VB_Name = "modNoConstant"
Option Explicit
  
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" _
    (ByVal lpSectionName As String, ByVal lpKeyName As String, _
     ByVal lpDefault As String, ByVal lpbuffurnedString As String, _
     ByVal nBuffSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSectionNames Lib "Kernel32.dll" Alias _
    "GetPrivateProfileSectionNamesA" _
    (ByVal lpszReturnBuffer As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

Public Sub ConnectToDatabase()

      Dim dbConnectRemoteBB As String
10    ReDim Cnxn(0 To 0) As Connection
20    ReDim CnxnBB(0 To 0) As Connection
30    ReDim CnxnRemoteBB(0 To 0) As Connection
      'ReDim HospName(0 To 0) As String
      Dim Con As String
      Dim ConBB As String

40    On Error GoTo ConnectToDatabase_Error

50    HospName(0) = GetcurrentConnectInfo(Con, ConBB)
60    If IsIDE And HospName(0) = "" Then
        'iMsg "INI Error"
70      frmMsgBox.Msg "INI Error"
80      End
90    ElseIf HospName(0) = "" Then
  
100      GetConnectInfo "Active", Con, HospName(0)
110       GetConnectInfo "BB", ConBB
120       GetConnectInfo "RemoteBB", dbConnectRemoteBB


130     If dbConnectRemoteBB <> "" Then
140       Set CnxnRemoteBB(0) = New Connection
150       CnxnRemoteBB(0).Open dbConnectRemoteBB
160     End If
170   End If

180   Set Cnxn(0) = New Connection
190   Cnxn(0).Open Con
200   ConnectionString = Con
210   If ConBB <> "" Then
220     Set CnxnBB(0) = New Connection
230     CnxnBB(0).Open ConBB
240   End If

250   CheckGroupedHospitalsInDb


260   Exit Sub

ConnectToDatabase_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   MsgBox "Error in modNoConstant, ConnectToDatabase line " & intEL & _
             " Con=" & Con & _
             " ConBB=" & ConBB & _
             " Error=" & strES & _
             " Hosp=" & HospName(0)

End Sub
Public Sub ConnectToLimNetacquireDb()

      Dim dbConnect As String
      Dim strConnection As String


10    On Error GoTo ConnectToLimNetacquireDb_Error

20    ReDim Preserve Cnxn(0 To 1) As Connection

      'Get Connection string from registry
30    strConnection = QueryValue("SOFTWARE\CustomSoftware\Netacquire", "ConnectionString")

40    strConnection = Obfuscate(strConnection)

50    If Len(strConnection) > 0 Then
60        dbConnect = strConnection
70    Else
          'iMsg "NetAcquire database connection string not specified in registry!"
80        End
90    End If

100   Set Cnxn(1) = New Connection
110   Cnxn(1).Open dbConnect

120   Exit Sub

ConnectToLimNetacquireDb_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "modNoConstant", "ConnectToLimNetacquireDb", intEL, strES

End Sub

Public Sub connectDb()

          Dim dbConnect As String
          Dim strConnection As String


10        ReDim Cnxn(0 To 0) As Connection

20        If App.PrevInstance Then End

          'Get Hospital(site) Name from registry
          HospName(0) = QueryValue("SOFTWARE\CustomSoftware", "Site")

          If Len(HospName(0)) = 0 Then
              'iMsg "Site not specified in registry!"
              End
          End If
          'Uncommented by Ibrahim

          'Get Connection string from registry
30        strConnection = QueryValue("SOFTWARE\CustomSoftware\HistologyCS", "ConnectionString")

40        strConnection = Obfuscate(strConnection)

50        If Len(strConnection) > 0 Then
60            dbConnect = strConnection
70        Else
              'iMsg "NetAcquire database connection string not specified in registry!"
'80            End
90        End If

100       Set Cnxn(0) = New Connection
110       Cnxn(0).Open dbConnect
    
          'ConnectionString = dbConnect
End Sub

Public Function QueryValue(sKeyName As String, sValueName As String) As String
          Dim lRetVal As Long         'result of the API functions
          Dim hKey As Long         'handle of opened key
          Dim vValue As Variant      'setting of queried value

10        QueryValue = ""

20        lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, _
                                 KEY_QUERY_VALUE, hKey)
30        lRetVal = QueryValueEx(hKey, sValueName, vValue)

          'MsgBox vValue
40        RegCloseKey (hKey)

50        QueryValue = vValue

End Function
Public Sub CheckGroupedHospitalsInDb()

      Dim sql As String

10    On Error GoTo CheckGroupedHospitalsInDb_Error

20    If IsTableInDatabase("GroupedHospitals") = False Then 'There is no table  in database
30      sql = "CREATE TABLE GroupedHospitals " & _
              "( [HospName] nvarchar(50), " & _
              "  [Connect] nvarchar(250), " & _
              "  [ConnectBB] nvarchar(250), " & _
              "  [UseInIDE] bit )"
40      Cnxn(0).Execute sql
50    End If

60    Exit Sub

CheckGroupedHospitalsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modNoConstant", "CheckGroupedHospitalsInDb", intEL, strES, sql

End Sub
Public Function IsTableInDatabase(ByVal TableName As String) As Boolean

      Dim tbExists As Recordset
      Dim sql As String
      Dim RetVal As Boolean

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist
      'if it has a record then the table does exist.

10    On Error GoTo IsTableInDatabase_Error

20    sql = "SELECT name FROM sysobjects WHERE " & _
            "xtype = 'U' " & _
            "AND name = '" & TableName & "'"
30    Set tbExists = Cnxn(0).Execute(sql)

40    RetVal = True

50    If tbExists.EOF Then 'There is no table <TableName> in database
60      RetVal = False
70    End If
80    IsTableInDatabase = RetVal

90    Exit Function

IsTableInDatabase_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modNoConstant", "IsTableInDatabase", intEL, strES, sql

  
End Function



'Private Sub GetHospitalsInGroup()
'
'      Dim sql As String
'      Dim tb As Recordset
'
'10    On Error GoTo GetHospitalsInGroup_Error
'
'20    intOtherHospitalsInGroup = 0
'
'30    sql = "SELECT * FROM GroupedHospitals " & _
'            "WHERE UseInIDE = " & IIf(IsIDE, 1, 0)
'
'40    Set tb = New Recordset
'50    RecOpenServer 0, tb, sql
'60    Do While Not tb.EOF
'
'70      intOtherHospitalsInGroup = intOtherHospitalsInGroup + 1
'80      ReDim Preserve Cnxn(0 To intOtherHospitalsInGroup)
'90      ReDim Preserve CnxnBB(0 To intOtherHospitalsInGroup)
'100     ReDim Preserve HospName(0 To intOtherHospitalsInGroup)
'
'110     HospName(intOtherHospitalsInGroup) = tb!HospName & ""
'
'120     If tb!Connect & "" <> "" Then
'130       Set Cnxn(intOtherHospitalsInGroup) = New Connection
'140       Cnxn(intOtherHospitalsInGroup).Open tb!Connect
'150     End If
'160     If tb!ConnectBB & "" <> "" Then
'170       Set CnxnBB(intOtherHospitalsInGroup) = New Connection
'180       CnxnBB(intOtherHospitalsInGroup).Open tb!ConnectBB
'190     End If
'
'200     tb.MoveNext
'210   Loop
'
'220   Exit Sub
'
'GetHospitalsInGroup_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'230   intEL = Erl
'240   strES = Err.Description
'250   LogError "modNoConstant", "GetHospitalsInGroup", intEL, strES, sql
'
'
'End Sub

Public Function GetConnectInfo(ByVal ConnectTo As String, _
                               ByRef ReturnConnectionString As String, _
                               Optional ByRef HospName As Variant) As Boolean

      'ConnectTo = "Active"
      '            "BB"
      '            "Active" & n - HospitalGroup
      '            "BB" & n - HospitalGroup

10    GetConnectInfo = False

20    If Not IsMissing(HospName) Then
30      HospName = GetSetting("NetAcquire", "HospName", ConnectTo, "")
40      If Left$(UCase$(HospName), 5) = "LOCAL" Then
50        HospName = Mid$(HospName, 6)
60      End If
70    End If

80    ReturnConnectionString = GetSetting("NetAcquire", "Cnxn", ConnectTo, "")

90    If Trim$(ReturnConnectionString) <> "" Then
  
100     ReturnConnectionString = Obfuscate(ReturnConnectionString)
  
110     GetConnectInfo = True
  
120   End If

End Function


Public Function GetcurrentConnectInfo(ByRef Con As String, ByRef ConBB As String) As String

      'Returns Hospital Name

      Dim HospitalNames() As String
      Dim n As Long
      Dim HospitalName As String
      Dim retHospitalName As String
      Dim ServerName As String
      Dim NetAcquireDB As String
      Dim TransfusionDB As String
      Dim uId As String
      Dim PWD As String
      Dim CurrentPath As String

10    On Error GoTo GetcurrentConnectInfo_Error
'20    If IsIDE Then
'30        If Dir("C:\ClientCode\NetAcquire.INI") <> "" Then
'40            CurrentPath = "C:\ClientCode\NetAcquire.INI"
'50        Else
'60            GetcurrentConnectInfo = "StJohns"
'70            Exit Function
'80        End If
'90    Else
20        If Dir(App.Path & "\NetAcquire.INI") <> "" Then
30            CurrentPath = App.Path & "\NetAcquire.INI"
40        Else
50            GetcurrentConnectInfo = "StJohns"
60            Exit Function
70        End If
'160   End If

80    HospitalNames = GetINISectionNames(CurrentPath, n)
90    HospitalName = HospitalNames(0)
100   If Left$(UCase$(HospitalName), 5) = "LOCAL" Then
110     retHospitalName = Mid$(HospitalName, 6)
120   Else
130     retHospitalName = HospitalName
140   End If


''170   HospitalNames = "Tullamore"
''180   HospitalName = "Tullamore"
''190   If Left$(UCase$(HospitalName), 5) = "LOCAL" Then
''200     retHospitalName = "Tullamore"
''210   Else
''220     retHospitalName = "Tullamore"
''230   End If

150   ServerName = ProfileGetItem(HospitalName, "N", "", CurrentPath)
160   NetAcquireDB = ProfileGetItem(HospitalName, "D", "", CurrentPath)
170   TransfusionDB = ProfileGetItem(HospitalName, "T", "", CurrentPath)

180   PWD = GetPass(uId)
'PWD = "DfySiywtgtw$1>)*"
190   PWD = "DfySiywtgtw$1>)="

200   Con = "DRIVER={SQL Server};" & _
            "Server=" & Obfuscate(ServerName) & ";" & _
            "Database=" & Obfuscate(NetAcquireDB) & ";" & _
            "uid=" & uId & ";" & _
            "pwd=" & PWD & ";"
'281     Con = "DRIVER={SQL Server};" & _
'            "Server=" & "192.168.20.21" & ";" & _
'            "Database=" & "Histology_TullTest" & ";" & _
'            "uid=" & "LabUser" & ";" & _
'            "pwd=" & "DfySiywtgtw$1>)=" & ";"
'        MsgBox (Con)
''''
'210   Con = "DRIVER={SQL Server};" & _
'      "Server=DESKTOP-3OMS1N5\SQLEXPRESS;" & _
'      "Database=Histology_TullTest;" & _
'      "Trusted_Connection=Yes;"

220   If TransfusionDB <> "" Then
230     ConBB = "DRIVER={SQL Server};" & _
                "Server=" & Obfuscate(ServerName) & ";" & _
                "Database=" & Obfuscate(TransfusionDB) & ";" & _
                "uid=" & uId & ";" & _
                "pwd=" & PWD & ";"
'        ConBB = "DRIVER={SQL Server};" & _
'                "Server=" & "192.168.20.21" & ";" & _
'                "Database=" & "Transfusion_TullLive" & ";" & _
'                "uid=" & "usman" & ";" & _
'                "pwd=" & "usman123" & ";"
240   End If



250   GetcurrentConnectInfo = retHospitalName

'320   GetcurrentConnectInfo = "Tullamore"

260   Exit Function

GetcurrentConnectInfo_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   MsgBox (strES)
300   MsgBox "GetCurrentConnectInfo Error Line " & intEL
310   LogError "modNoConstant", "GetcurrentConnectInfo", intEL, strES

End Function
Private Function ProfileGetItem(ByRef sSection As String, _
                                ByRef sKeyName As String, _
                                ByRef sDefValue As String, _
                                ByRef sIniFile As String) As String

          'retrieves a value from an ini file
          'corresponding to the section and
          'key name passed.

      Dim dwSize As Integer
      Dim nBuffSize As Integer
      Dim buff As String
      Dim RetVal As String

      'Call the API with the parameters passed.
      'nBuffSize is the length of the string
      'in buff, including the terminating null.
      'If a default value was passed, and the
      'section or key name are not in the file,
      'that value is returned. If no default
      'value was passed (""), then dwSize
      'will = 0 if not found.
      '
      'pad a string large enough to hold the data
10    buff = Space(2048)
20    nBuffSize = Len(buff)
30    dwSize = GetPrivateProfileString(sSection, sKeyName, sDefValue, buff, nBuffSize, sIniFile)

40    If dwSize > 0 Then
50      RetVal = Left$(buff, dwSize)
60    End If

70    ProfileGetItem = RetVal

End Function

Private Function GetINISectionNames(ByRef inFile As String, ByRef outCount As Long) As String()

      Dim StrBuf As String
      Dim BufLen As Long
      Dim RetVal() As String
      Dim Count As Long

10    BufLen = 16

20    Do
30      BufLen = BufLen * 2
40      StrBuf = Space$(BufLen)
50      Count = GetPrivateProfileSectionNames(StrBuf, BufLen, inFile)
60    Loop While Count = BufLen - 2

70    If (Count) Then
80      RetVal = Split(Left$(StrBuf, Count - 1), vbNullChar)
90      outCount = UBound(RetVal) + 1
100   End If

110   GetINISectionNames = RetVal

End Function
  

Public Function Obfuscate(ByVal strData As String) As String

      Dim lngI As Long
      Dim lngJ As Long
   
10    For lngI = 0 To Len(strData) \ 4
20      For lngJ = 1 To 4
30         Obfuscate = Obfuscate & Mid$(strData, (4 * lngI) + 5 - lngJ, 1)
40      Next
50    Next

End Function

Private Function GetPass(ByRef uId As String) As String

      Dim p As String
      Dim A As String
      Dim n As Integer

10    A = ""
20    For n = 97 To 122
30      A = A & Chr$(n)
40    Next
50    For n = 65 To 90
60      A = A & Chr$(n)
70    Next

80    A = A & "!£$%^&*()<>-_+={}[]:@~||;'#,./?"
90    For n = 48 To 57
100     A = A & Chr$(n)
110   Next

      '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
      '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
      '             1         2         3         4         5         6         7         8         9

      'p = ""
      'UID = "sa"

      'LabUser
120   uId = Mid$(A, 38, 1) & Mid$(A, 1, 1) & Mid$(A, 2, 1) & Mid$(A, 47, 1) & _
            Mid$(A, 19, 1) & Mid$(A, 5, 1) & Mid$(A, 18, 1)
    
      'DfySiywtgtw$1>)*
130   p = Mid$(A, 30, 1) & Mid$(A, 6, 1) & Mid$(A, 25, 1) & Mid$(A, 45, 1) & _
          Mid$(A, 9, 1) & Mid$(A, 25, 1) & Mid$(A, 23, 1) & Mid$(A, 20, 1) & _
          Mid$(A, 7, 1) & Mid$(A, 20, 1) & Mid$(A, 23, 1) & Mid$(A, 55, 1) & _
          Mid$(A, 85, 1) & Mid$(A, 63, 1) & Mid$(A, 61, 1) & Mid$(A, 59, 1)

140   GetPass = p

End Function


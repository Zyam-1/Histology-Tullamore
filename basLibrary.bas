Attribute VB_Name = "basLibrary"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public strEvent As String

Public Enum InputValidation
    NumericFullStopDash = 0
    Char = 1
    YorN = 2
    AlphaNumeric_NoApos = 3
    AlphaNumeric_AllowApos = 4
    Numeric_Only = 5
    AlphaOnly = 6
    NumericSlash = 7
    AlphaAndSpaceonly = 8
    CharNumericDashSlash = 9
    AlphaAndSpaceApos = 10
    DecimalNumericOnly = 11
    CharNumericDashSlashFullStop = 12
    ivSampleID = 13
    AlphaNumeric = 14
    AlphaNumericSpace = 15
    NumericDMY = 16
    
End Enum

Public Const FORWARD = -1
Public Const BACKWARD = 1
Public Const DONTCARE = 0


Public UserName As String
Public UserCode As String
Public UserInitials As String

Public blnVIstatus As Boolean
Public strDeptLetter4Histo As String

Public Custom As Boolean
'*********************Flex Grid format constants (Babar Shahzad 2008-09-25)**********
'***NOTE: Change in constant values will affect all FlexGrids in application
Public Const fgcBackColorFixed As Long = &H80000001     'Desktop
Public Const fgcForeColorFixed As Long = &H80000005     'White
Public Const fgcFontBoldFixed As Boolean = False
Public Const fgcBackColor As Long = &H80000018          'Tooltip Text
Public Const fgcForeColor As Long = &H80000008          'Black
Public Const fgcFontSize As Integer = 9
Public Const fgcFontName As String = "MS Sans Serif"
Public Const fgcExtraSpace As Integer = 150

'*********************End Flex Grid format constants*********************************

Public Sub WriteInitGridCode(g As MSFlexGrid)
      Dim i As Integer

      '10    Debug.Print "private sub InitializeGrid()"
      '
      '20    Debug.Print "dim I as integer"
      '30    Debug.Print "with " & g.Name
      '
      '40    Debug.Print vbTab & ".rows=2: .fixedrows=1"
      '50    Debug.Print vbTab & ".cols = " & g.Cols & ": .fixedcols= 0 "
      '60    Debug.Print vbTab & ".rows=1"
      '70    Debug.Print vbTab & ".font.size = fgcFontSize"
      '80    Debug.Print vbTab & ".font.name = fgcFontName"
      '90    Debug.Print vbTab & ".forecolor = fgcForeColor"
      '100   Debug.Print vbTab & ".backcolor = fgcbackColor"
      '110   Debug.Print vbTab & ".forecolorFixed = fgcForeColorfixed"
      '120   Debug.Print vbTab & ".BackColorFixed = fgcBackColorFixed"
      '130   Debug.Print vbTab & ".ScrollBars = flexScrollBarBoth"
      '140   Debug.Print vbTab & "'" & g.FormatString
10    For i = 0 To g.Cols - 1
    
20        If g.TextMatrix(0, i) = "" And g.ColWidth(i) = 0 Then
30            Debug.Print ".colwidth(" & i & ") = 0"
40        Else
50            Debug.Print vbTab & ".textmatrix(0," & i & ") = "; "NetAcquire"; " :" & _
                          ".colwidth(" & i & ") = " & g.ColWidth(i) & " : " & _
                          ".colalignment(" & i & ")= flexalignleftcenter"
60        End If
70    Next i
80    Debug.Print vbTab & "For i = 0 To .Cols - 1"
90    Debug.Print vbTab & vbTab & "If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then"
100   Debug.Print vbTab & vbTab & vbTab & ".ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace"
110   Debug.Print vbTab & vbTab & "End If"
120   Debug.Print vbTab & "Next i"
130   Debug.Print "end with"
140   Debug.Print "end sub"
End Sub


Public Function VI(KeyAscii As Integer, _
                   iv As InputValidation, _
                   Optional NextFieldOnEnter As Boolean) As Integer

      Dim sTemp As String

10    If blnVIstatus Then
20        sTemp = Chr$(KeyAscii)
30        If KeyAscii = 13 Then 'Enter Key
40          If NextFieldOnEnter = True Then
50            VI = 9 'Return Tab Keyascii if User Selected NextFieldOnEnter Option
60          Else
70            VI = 13
80          End If
90          Exit Function
100       ElseIf KeyAscii = 8 Then 'BackSpace
110         VI = 8
120         Exit Function
130       End If
          
          ' turn input to upper case
          
140       Select Case iv
            Case InputValidation.NumericFullStopDash:
150           Select Case sTemp
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", "-"
160               VI = Asc(sTemp)
170             Case Else
180               VI = 0
190           End Select
          
200         Case InputValidation.ivSampleID
210           Select Case sTemp
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
220               VI = Asc(sTemp)
230             Case "A" To "Z"
240               VI = Asc(sTemp)
250             Case "a" To "z"
260               VI = Asc(sTemp) - 32 'Convert to upper case
270             Case Else
280               VI = 0
290           End Select
          
300         Case InputValidation.AlphaNumeric
310           Select Case sTemp
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
320               VI = Asc(sTemp)
330             Case "A" To "Z"
340               VI = Asc(sTemp)
350             Case "a" To "z"
360               VI = Asc(sTemp)
370             Case Else
380               VI = 0
390           End Select
          
400         Case InputValidation.AlphaNumericSpace
410           Select Case sTemp
                Case " ", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "<", ">"
420               VI = Asc(sTemp)
430             Case "A" To "Z"
440               VI = Asc(sTemp)
450             Case "a" To "z"
460               VI = Asc(sTemp)
470             Case Else
480               VI = 0
490           End Select
          
500         Case InputValidation.Char
510           Select Case sTemp
                Case " ", "-"
520               VI = Asc(sTemp)
530             Case "A" To "Z"
540               VI = Asc(sTemp)
550             Case "a" To "z"
560               VI = Asc(sTemp)
570             Case Else
580               VI = 0
590           End Select
          
600         Case InputValidation.YorN
610           sTemp = UCase(Chr$(KeyAscii))
620           Select Case sTemp
                Case "Y", "N"
630               VI = Asc(sTemp)
640             Case Else
650               VI = 0
660           End Select
          
670         Case InputValidation.AlphaNumeric_NoApos
680           Select Case sTemp
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", _
                     " ", "/", ";", ":", "\", "-", ">", "<", "(", ")", "@", _
                     "%", "!", """", "+", "^", "~", "`", "Ç", "´", "Ã", "Á", _
                     "Â", "È", "É", "Ê", "Ì", "Í", "Î", "Ò", "Ó", "Ô", "Õ", _
                     "Ù", "Ú", "Û", "Ü", "à", "á", "â", "ã", "ç", "è", "é", _
                     "ê", "ì", "í", "î", "ò", "ó", "ô", "õ", "ö", "ù", "ú", _
                     "û", "ü", "Æ", "æ", ",", "?", "&", "="
690               VI = Asc(sTemp)
700             Case "A" To "Z"
710               VI = Asc(sTemp)
720             Case "a" To "z"
730               VI = Asc(sTemp)
740             Case Else
750               VI = 0
760           End Select
          
770         Case InputValidation.AlphaNumeric_AllowApos
780           Select Case sTemp
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", " ", "'"
790               VI = Asc(sTemp)
800             Case "A" To "Z"
810               VI = Asc(sTemp)
820             Case "a" To "z"
830               VI = Asc(sTemp)
840             Case Else
850               VI = 0
860           End Select
              
870         Case InputValidation.Numeric_Only
880           Select Case sTemp
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
890               VI = Asc(sTemp)
900             Case Else
910               VI = 0
920           End Select
          
930         Case InputValidation.AlphaOnly
940           Select Case sTemp
                Case "A" To "Z"
950               VI = Asc(sTemp)
960             Case "a" To "z"
970               VI = Asc(sTemp)
980             Case Else
990               VI = 0
1000          End Select
          
1010        Case InputValidation.NumericSlash
1020          Select Case sTemp
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/"
1030              VI = Asc(sTemp)
1040            Case Else
1050              VI = 0
1060          End Select
          
1070        Case InputValidation.AlphaAndSpaceonly
1080          Select Case sTemp
                Case " "
1090              VI = Asc(sTemp)
1100            Case "A" To "Z"
1110              VI = Asc(sTemp)
1120            Case "a" To "z"
1130              VI = Asc(sTemp)
1140            Case Else
1150              VI = 0
1160          End Select
              
1170          Case InputValidation.CharNumericDashSlash
1180            Select Case sTemp
                  Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/", "-"
1190                VI = Asc(sTemp)
1200              Case "A" To "Z"
1210                VI = Asc(sTemp)
1220              Case "a" To "z"
1230                VI = Asc(sTemp) - 32 'Convert to upper case
1240              Case Else
1250                VI = 0
1260            End Select
          
1270          Case InputValidation.AlphaAndSpaceApos
1280          Select Case sTemp
                Case " ", "'"
1290              VI = Asc(sTemp)
1300            Case "A" To "Z"
1310              VI = Asc(sTemp)
1320            Case "a" To "z"
1330              VI = Asc(sTemp)
1340            Case Else
1350              VI = 0
1360          End Select
              
1370          Case InputValidation.DecimalNumericOnly
1380          Select Case sTemp
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "."
1390              VI = Asc(sTemp)
1400            Case Else
1410              VI = 0
1420          End Select
              
1430          Case InputValidation.CharNumericDashSlashFullStop
1440            Select Case sTemp
                  Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "/", "-", "."
1450                VI = Asc(sTemp)
1460              Case "A" To "Z"
1470                VI = Asc(sTemp)
1480              Case "a" To "z"
1490                VI = Asc(sTemp)
1500              Case Else
1510                VI = 0
1520            End Select
              
1530          Case InputValidation.NumericDMY
1540            Select Case sTemp
                  Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "D", "M", "Y", "d", "m", "y"
1550                VI = Asc(sTemp)
1560              Case Else
1570                VI = 0
1580            End Select
              
1590      End Select
          
1600      If VI = 0 Then Beep
1610  Else
1620      VI = KeyAscii
1630  End If

End Function
Public Function vbGetComputerName() As String
  
      'Gets the name of the machine
      Const MAXSIZE As Integer = 256
      Dim sTmp As String * MAXSIZE
      Dim lLen As Long
 
10    On Error GoTo vbGetComputerName_Error

20    lLen = MAXSIZE - 1
30    If (GetComputerName(sTmp, lLen)) Then
40      vbGetComputerName = Left$(sTmp, lLen)
50    Else
60      vbGetComputerName = ""
70    End If

80    Exit Function

vbGetComputerName_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "basLibrary", "vbGetComputerName", intEL, strES


End Function
Public Function CheckPhoneLog(ByVal SID As String, ByVal Year As String) As Boolean

      'Returns True if an entry in phone log

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo CheckPhoneLog_Error

20    sql = "Select * FROM PhoneLog WHERE " & _
            "SampleID = '" & Val(SID) & "' " & _
            "AND Year = '" & Year & "'"
30    Set tb = Cnxn(0).Execute(sql)

40    CheckPhoneLog = Not tb.EOF

50    Exit Function

CheckPhoneLog_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "basLibrary", "CheckPhoneLog", intEL, strES, sql

End Function


Public Function AddTicks(ByVal s As String) As String

10    On Error GoTo AddTicks_Error

20    s = Trim$(s)

30    s = Replace(s, "'", "''")

40    AddTicks = s

50    Exit Function

AddTicks_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "basLibrary", "AddTicks", intEL, strES

End Function

Public Function RemoveTicks(ByVal s As String) As String

10    On Error GoTo RemoveTicks_Error

20    s = Trim$(s)

30    s = Replace(s, "'", " ")

40    RemoveTicks = s

50    Exit Function

RemoveTicks_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "basLibrary", "RemoveTicks", intEL, strES

End Function


Public Function BetweenDates(ByVal Index As Integer, _
                             ByRef upto As String) _
                             As String

      Dim From As String
      Dim m As Long

10    On Error GoTo BetweenDates_Error

20    Select Case Index
        Case 0: 'last week
30              From = DateAdd("ww", -1, Now)
40              upto = Now
50      Case 1: 'last month
60              From = DateAdd("m", -1, Now)
70              upto = Now
80      Case 2: 'last fullmonth
90              From = DateAdd("m", -1, Now)
100             From = "01/" & Mid$(From, 4)
110             upto = DateAdd("m", 1, From)
120             upto = DateAdd("d", -1, upto)
130     Case 3: 'last quarter
140             From = DateAdd("q", -1, Now)
150             upto = Now
160     Case 4: 'last full quarter
170             From = DateAdd("q", -1, Now)
180             m = Val(Mid$(From, 4, 2))
190             m = ((m - 1) \ 3) * 3 + 1
200             From = "01/" & Format$(m, "00") & Mid$(From, 6)
210             upto = DateAdd("q", 1, From)
220             upto = DateAdd("d", -1, upto)
230     Case 5: 'year to date
240             From = "01/01/" & Format(Now, "yyyy")
250             upto = Now
260     Case 6: 'today
270             From = Now
280             upto = From
290   End Select

300   BetweenDates = From

310   Exit Function

BetweenDates_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "basLibrary", "BetweenDates", intEL, strES


End Function
Public Function CalcpAge(ByVal DoB As String) As String
      Dim diff As Long
      Dim DobYr As Single


10    On Error GoTo CalcpAge_Error

20    DoB = Format$(DoB, "Short Date")
30    If IsDate(DoB) Then
40      diff = DateDiff("d", (DoB), (Now))
50      DobYr = diff / 365.25
60      If DobYr > 1 Then
70        CalcpAge = Int(DobYr)
80        ElseIf diff < 30.43 Then
90        CalcpAge = diff
100       Else
110       CalcpAge = Int(diff / 30.43)
120       End If
130   Else
140     CalcpAge = ""
150   End If


160   Exit Function

CalcpAge_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "basLibrary", "CalcpAge", intEL, strES


End Function

Public Function CalcOldAge(ByVal DoB As String, ByVal Old As String) As String

      Dim diff As Long
      Dim DobYr As Single

10    On Error GoTo CalcOldAge_Error

20    DoB = Format$(DoB, "Short Date")
30    If IsDate(DoB) Then
40      diff = DateDiff("d", (DoB), (Old))
50      DobYr = diff / 365.25
60      If DobYr > 1 Then
70        CalcOldAge = Format$(Int(DobYr), "###\" & " Years")
80        ElseIf diff < 30.43 Then
90        CalcOldAge = Format$(diff, "##\" & " Days")
100       Else
110       CalcOldAge = Format$(Int(diff / 30.43), "##\" & " Months")
120       End If
130   Else
140     CalcOldAge = ""
150   End If


160   Exit Function

CalcOldAge_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "basLibrary", "CalcOldAge", intEL, strES


End Function





Public Sub ClearFGrid(ByVal g As MSFlexGrid)


10    On Error GoTo ClearFGrid_Error

20    With g
30      .Rows = .FixedRows + 1
40      .AddItem ""
50      .RemoveItem .FixedRows
60      .Visible = False
70    End With

80    Exit Sub

ClearFGrid_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "basLibrary", "ClearFGrid", intEL, strES


End Sub


Public Function Convert62Date(ByVal s As String, _
                              ByVal Direction As Long) _
                              As String

      Dim d As String

10    On Error GoTo Convert62Date_Error

20    If Len(s) <> 6 Then
30      Convert62Date = s
40      Exit Function
50    End If

60    d = Left(s, 2) & "/" & Mid(s, 3, 2) & "/" & Right(s, 2)
70    If IsDate(d) Then
80      Select Case Direction
          Case BACKWARD:
90          If DateValue(d) > DateValue(Now) Then
100           d = DateAdd("yyyy", -100, d)
110         End If
120         Convert62Date = Format$(d, "Short Date")
130       Case FORWARD:
140         If DateValue(d) < Now Then
150           d = DateAdd("yyyy", 100, d)
160         End If
170         Convert62Date = Format$(d, "Short Date")
180       Case DONTCARE:
190         Convert62Date = Format$(d, "Short Date")
200     End Select
210   Else
220     Convert62Date = s
230   End If

240   Exit Function

Convert62Date_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "basLibrary", "Convert62Date", intEL, strES


End Function

Public Function dmyFromCount(ByVal Days As Long) As String

      Dim d As Long
      Dim m As Long
      Dim Y As Long
      Dim s As String


10    On Error GoTo dmyFromCount_Error

20    Y = Int(Days / 365)

30    Days = Days - (Y * 365)

40    m = Days \ 30

50    d = Days - (m * 30)

60    If Y > 0 Then
70      s = Format$(Y) & "Y "
80    End If

90    If m > 0 Then
100     s = s & Format$(m) & "M "
110   End If
  
120   dmyFromCount = s & Format$(d, "0") & "D"


130   Exit Function

dmyFromCount_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "basLibrary", "dmyFromCount", intEL, strES

  
End Function




Public Sub FixG(ByVal g As MSFlexGrid)

10    On Error GoTo FixG_Error

20    With g
30      .Visible = True
40      If .Rows > .FixedRows + 1 And .TextMatrix(.FixedRows, 0) = "" Then
50        .RemoveItem .FixedRows
60      End If
70    End With



80    Exit Sub

FixG_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "basLibrary", "FixG", intEL, strES


End Sub

Public Sub FlashNoPrevious(ByVal f As Form)

      Dim T As Single
      Dim n As Long


10    On Error GoTo FlashNoPrevious_Error

20    With f.lNoPrevious
30      For n = 1 To 5
40        .Visible = True
50        .Refresh
60        T = Timer
70        Do While Timer - T < 0.1: DoEvents: Loop
80        .Visible = False
90        .Refresh
100       T = Timer
110       Do While Timer - T < 0.1: DoEvents: Loop
120     Next
130   End With


140   Exit Sub

FlashNoPrevious_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "basLibrary", "FlashNoPrevious", intEL, strES


End Sub






Function initial2upper(ByVal s As String) As String
    
      Dim n As Long


10    On Error GoTo initial2upper_Error
     

20    s = Trim$(s & "")
30    If s = "" Then
40        initial2upper = ""
50        Exit Function
60    End If
  
70    If InStr(UCase$(s), "MAC") > 0 Or InStr(UCase$(s), "MC") > 0 Or InStr(s, "'") > 0 Then
80    s = LCase$(s)
90    s = UCase$(Left$(s, 1)) & Mid(s, 2)

100   For n = 1 To Len(s) - 1
110       If Mid(s, n, 1) = " " Or Mid(s, n, 1) = "'" Or Mid(s, n, 1) = "." Then
120           s = Left$(s, n) & UCase$(Mid(s, n + 1, 1)) & Mid(s, n + 2)
130       End If
140       If n > 1 Then
150           If Mid(s, n, 1) = "c" And Mid(s, n - 1, 1) = "M" Then
160               s = Left$(s, n) & UCase$(Mid(s, n + 1, 1)) & Mid(s, n + 2)
170           End If
180       End If
190   Next
200   Else
210     s = StrConv(s, vbProperCase)
220   End If
230   initial2upper = s
    


240   Exit Function

initial2upper_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "basLibrary", "initial2upper", intEL, strES


End Function

Public Function IsRoutine() As Boolean

      'Returns True if time now is between
      '09:30 and 16:30 Mon to Fri
      'else returns False


10    On Error GoTo IsRoutine_Error

20    IsRoutine = False

30    If Weekday(Now) <> vbSaturday And Weekday(Now) <> vbSunday Then
40      If TimeValue(Now) > TimeValue("09:29") And _
           TimeValue(Now) < TimeValue("16:31") Then
50        IsRoutine = True
60      End If
70    End If




80    Exit Function

IsRoutine_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "basLibrary", "IsRoutine", intEL, strES


End Function



Public Function ParseForeName(ByVal Name As String) As String


10    On Error GoTo ParseForeName_Error

20    Name = Trim$(UCase$(Name))
      Dim n As Long
      Dim temp As String


30    If InStr(Name, "B/O") Or _
         InStr(Name, "BABY") Then
40      ParseForeName = ""
50      Exit Function
60    End If

70    temp = Name

80    n = InStr(temp, " ")
90    If n = 0 Then
100     ParseForeName = ""
110     Exit Function
120   End If

130   temp = Mid$(temp, n + 1)

      Rem Code Change 16/01/2006
      'checks if a double barreled name
140   n = InStr(temp, " ")
150   temp = Mid$(temp, n + 1)
160   If Trim(temp) = "" Then
170     Exit Function
180   End If

190   If InStr(temp, " ") Or _
           temp Like "*[!A-Z]*" Or _
           Len(temp) = 1 Then
200       ParseForeName = ""
210   Else
220     ParseForeName = temp
230   End If



240   Exit Function

ParseForeName_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "basLibrary", "ParseForeName", intEL, strES


End Function



Public Sub RecClose(ByVal rs As Recordset)


10    On Error GoTo RecClose_Error

20    rs.Close
30    Set rs = Nothing


40    Exit Sub

RecClose_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "basLibrary", "RecClose", intEL, strES


End Sub

Public Sub RecOpenClient(ByVal n As Long, ByVal RecSet As Recordset, ByVal sql As String)

10    With RecSet
20      .CursorLocation = adUseClient
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = Cnxn(n)
60      .Source = sql
70      .Open
80    End With

End Sub





Public Sub RecOpenServer(ByVal n As Long, ByVal RecSet As Recordset, ByVal sql As String)

10    With RecSet
20      .CursorLocation = adUseServer
30      .CursorType = adOpenDynamic
40      .LockType = adLockOptimistic
50      .ActiveConnection = Cnxn(n)
60      .Source = sql
70      .Open
80    End With

End Sub





Public Function Split_Comm(ByVal Comm As String) As String
      Dim n As Long
      Dim s As String
      Dim Cnt  As Long

10    On Error GoTo Split_Comm_Error

20    For n = 1 To Len(Comm)
30      If Asc(Mid(Comm, n, 1)) = Asc(vbCr) Or Asc(Mid(Comm, n, 1)) = 10 Then
40        If Cnt = 0 Then
50          If n > 1 Then s = s & vbCrLf
60           Cnt = 1
70        End If
80      Else
90        s = s & Mid(Comm, n, 1)
100       Cnt = 0
110     End If
120   Next

130   Split_Comm = s


140   Exit Function

Split_Comm_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "basLibrary", "Split_Comm", intEL, strES


End Function

'Public Sub LogError(ByVal ModuleName As String, _
'                    ByVal ProcedureName As String, _
'                    ByVal ErrorLineNumber As Integer, _
'                    ByVal ErrorDescription As String, _
'                    Optional ByVal SQLStatement As String, _
'                    Optional ByVal EventDesc As String)
'
'
'      Dim tb As Recordset
'      Dim sql As String
'      Dim MyMachineName As String
'      Dim Vers As String
'
'10    Vers = App.Major & "-" & App.Minor & "-" & App.Revision
'
'20    MyMachineName = vbGetComputerName()
'
'30    sql = "INSERT INTO ErrorLog " & _
'            "( DateTimeOfRecord, AppName, AppVersion, ModuleName, ProcedureName, " & _
'            "  ErrorLineNumber, SQLStatement, ErrorDescription, " & _
'            "  UserName, MachineName, Eventdesc) VALUES " & _
'            "('" & Format$(Now, "yyyymmdd hh:mm:ss") & "', " & _
'            "'" & App.EXEName & "'," & _
'            "'" & Vers & "'," & _
'            "'" & ModuleName & "', " & _
'            "'" & ProcedureName & "', " & _
'            "'" & ErrorLineNumber & "', " & _
'            "'" & AddTicks(SQLStatement) & "', " & _
'            "'" & AddTicks(ErrorDescription) & "', " & _
'            "'" & AddTicks(UserName) & "', " & _
'            "'" & AddTicks(MyMachineName) & "', " & _
'            "'" & AddTicks(EventDesc) & "')"
'40    Set tb = New Recordset
'50    RecOpenClient 0, tb, sql
'
'End Sub

Public Sub LogError(ByVal ModuleName As String, _
                    ByVal ProcedureName As String, _
                    ByVal ErrorLineNumber As Integer, _
                    ByVal ErrorDescription As String, _
                    Optional ByVal SQLStatement As String, _
                    Optional ByVal EventDesc As String)

Dim sql As String
Dim MyMachineName As String
Dim Vers As String

10    On Error Resume Next

20    ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "[MSSQL]")
30    ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver]", "[SQL]")
40    ErrorDescription = AddTicks(ErrorDescription)

50    SQLStatement = AddTicks(SQLStatement)
60    EventDesc = AddTicks(EventDesc)

70    Vers = App.Major & "-" & App.Minor & "-" & App.Revision

80    MyMachineName = vbGetComputerName()

90    sql = "IF NOT EXISTS " & _
      "    (SELECT * FROM ErrorLog WHERE " & _
      "     ModuleName = '" & ModuleName & "' " & _
      "     AND ProcedureName = '" & ProcedureName & "' " & _
      "     AND ErrorLineNumber = '" & ErrorLineNumber & "') " & _
      "  INSERT INTO ErrorLog (" & _
      "    ModuleName, ProcedureName, ErrorLineNumber, SQLStatement, " & _
      "    ErrorDescription, UserName, MachineName, Eventdesc, AppName, AppVersion, EventCounter, Emailed) " & _
      "  VALUES  ('" & ModuleName & "', " & _
      "           '" & ProcedureName & "', " & _
      "           '" & ErrorLineNumber & "', " & _
      "           '" & SQLStatement & "', " & _
      "           '" & ErrorDescription & "', " & _
      "           '" & AddTicks(UserName) & "', " & _
      "           '" & MyMachineName & "', " & _
      "           '" & EventDesc & "', " & _
      "           '" & App.EXEName & "', " & _
      "           '" & Vers & "', " & _
      "           '1', '0') " & _
      "ELSE "
100   sql = sql & "  UPDATE ErrorLog " & _
      "  SET SQLStatement = '" & SQLStatement & "', " & _
      "  ErrorDescription = '" & ErrorDescription & "', " & _
      "  MachineName = '" & MyMachineName & "', " & _
      "  UserName = '" & AddTicks(UserName) & "', " & _
      "  AppName = '" & App.EXEName & "', " & _
      "  AppVersion = '" & Vers & "', " & _
      "  DateTime = getdate(), " & _
      "  EventCounter = COALESCE(EventCounter, 0) + 1 " & _
      "WHERE ModuleName = '" & ModuleName & "' " & _
      "AND ProcedureName = '" & ProcedureName & "' " & _
      "AND ErrorLineNumber = '" & ErrorLineNumber & "'"

110   Cnxn(0).Execute sql

End Sub


Public Function CheckNewEXE(ByVal NameOfExe As String) As String

      Dim FileName As String
      Dim Current As String
      Dim found As Boolean
      Dim Path As String

10    On Error GoTo CheckNewEXE_Error

20    found = False

30    Path = App.Path & "\"
40    Current = UCase$(NameOfExe) & ".EXE"
50    FileName = UCase$(Dir(Path & NameOfExe & "*.exe", vbNormal))

60    Do While FileName <> ""
70      If FileName > Current Then
80        Current = FileName
90        found = True
100     End If
110     FileName = UCase$(Dir)
120   Loop

130   If found And UCase$(App.EXEName) & ".EXE" <> Current Then
140     CheckNewEXE = Path & Current
150   Else
160     CheckNewEXE = ""
170   End If

180   Exit Function

CheckNewEXE_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "basLibrary", "CheckNewEXE", intEL, strES


End Function
Public Function MaskInput(KeyAscii As Integer, Text As String, InputMask As String) As Integer

      '---------------------------------------------------------------------------------------
      ' Procedure : MaskInput
      ' DateTime  : 06/06/2008 15:28
      ' Author    : Babar Shahzad
      ' Purpose   : Masks input and doesnt allow user to add more than masked number of chars.
      '               X,x = Alphabets
      '               # = Number
      '               all other chars should match mask
      '---------------------------------------------------------------------------------------

10    If KeyAscii = 8 Then
20        MaskInput = KeyAscii
30        Exit Function
40    End If

50    If Len(Text) = Len(InputMask) Then
60        MaskInput = 0
70        Exit Function
80    End If

90    If Mid(InputMask, Len(Text) + 1, 1) = "X" Then
100       If KeyAscii >= 65 And KeyAscii <= 90 Then
110           MaskInput = KeyAscii
120           Exit Function
130       ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
140           MaskInput = KeyAscii - 32
150           Exit Function
160       Else
170           MaskInput = 0
180           Exit Function
190       End If
200   ElseIf Mid(InputMask, Len(Text) + 1, 1) = "x" Then
210       If KeyAscii >= 65 And KeyAscii <= 90 Then
220           MaskInput = KeyAscii + 32
230           Exit Function
240       ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
250           MaskInput = KeyAscii
260           Exit Function
270       Else
280           MaskInput = 0
290           Exit Function
300       End If
310   ElseIf Mid(InputMask, Len(Text) + 1, 1) = "#" Then
320       If KeyAscii >= 48 And KeyAscii <= 57 Then
330           MaskInput = KeyAscii
340           Exit Function
350       Else
360           MaskInput = 0
370           Exit Function
380       End If
390   Else
          'FOR ALL OTHER CHARACTERS
400       If KeyAscii = Asc(Mid(InputMask, Len(Text) + 1, 1)) Then
410           MaskInput = KeyAscii
420           Exit Function
430       Else
440           MaskInput = 0
450           Exit Function
460       End If
470   End If




End Function



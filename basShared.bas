Attribute VB_Name = "basShared"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Public intOtherHospitalsInGroup As Long

'User information
Public UserMemberOf As String
Public UserPass As String
Public strVersion As String

'logoff information
Public LogOffDelayMin As Long
Public LogOffDelaySecs As Long

'connections
Public Cnxn() As Connection

Public blnIsTestMode As Boolean
Public strTestSystemColor As String
Public TimedOut As Boolean
Public Answer As Long

Public strReportPath As String

Public ConnectionString As String

Public Rada As Integer
'hospital information
Public HospName(100) As String
Public CaseNo As String

'Public ListId As Integer

Public DataChanged As Boolean
Public TreeChanged As Boolean

Public Enum DataModeValues
    DataModeNew = 1
    DataModeEdit = 2
End Enum

Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160
Private Const LB_GETITEMHEIGHT = &H1A1

Public Declare Function MoveWindow Lib "user32" _
                                   (ByVal hwnd As Long, _
                                    ByVal X As Long, ByVal Y As Long, _
                                    ByVal nWidth As Long, _
                                    ByVal nHeight As Long, _
                                    ByVal bRepaint As Long) As Long

Public Declare Function SendMessage Lib "user32" _
                                    Alias "SendMessageA" _
                                    (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long

Public DataMode As DataModeValues

Public Function GetListID(Code As String, ListType As String) As Integer

    Dim tb As Recordset
    Dim sql As String

10    On Error GoTo GetListID_Error

20    sql = "Select ListID From Lists Where Code = N'" & Code & "' " & _
          "AND ListType = '" & ListType & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If tb.EOF Then
60      GetListID = 0
70    Else
80      GetListID = tb!ListId
90    End If

100   Exit Function

GetListID_Error:

    Dim strES As String
    Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basShared", "GetListID", intEL, strES, sql

End Function
'Public Function ConvertNull(Data As Variant, Default As Variant) As Variant
'    On Error GoTo ERROR_ConvertNull
'    If IsNull(Data) = True Then
'        ConvertNull = Default
'    Else
'        ConvertNull = Data
'    End If
'    Exit Function
'ERROR_ConvertNull:
'       Dim strES As String
'       Dim intEL As Integer
'
'110   intEL = Erl
'120   strES = Err.Description
'      LogError "basShared", "ConvertNull", intEL, strES
'
'
'End Function

Public Function ListCodeFor(ByVal ListType As String, ByVal Text As String) As String

    Dim tb As New Recordset
    Dim sql As String

10    On Error GoTo ListCodeFor_Error

20    ListCodeFor = ""

30    sql = "SELECT Code FROM Lists WHERE " & _
          "ListType = '" & ListType & "' " & _
          "AND Description = N'" & AddTicks(Text) & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      ListCodeFor = Trim(tb!Code)
80    End If

90    Exit Function

ListCodeFor_Error:

    Dim strES As String
    Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "basShared", "ListCodeFor", intEL, strES, sql

End Function
Public Function ListDescriptionFor(ByVal ListType As String, ByVal Code As String) As String

    Dim tb As New Recordset
    Dim sql As String

10    On Error GoTo ListCodeFor_Error

20    ListDescriptionFor = ""

30    sql = "SELECT Description FROM Lists WHERE " & _
          "ListType = '" & ListType & "' " & _
          "AND Code = N'" & AddTicks(Code) & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      ListDescriptionFor = Trim(tb!Description)
80    End If

90    Exit Function

ListCodeFor_Error:

    Dim strES As String
    Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "basShared", "ListCodeFor", intEL, strES, sql

End Function

Public Function UpdateListRank(ListId As Integer)

    Dim tb As Recordset
    Dim sql As String

10    On Error GoTo UpdateListRank_Error

20    sql = "Select * From Lists Where ListId = " & ListId & ""
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then

60      tb!Rank = IIf(IsNull(tb!Rank), 0, tb!Rank) + 1
70      tb.Update
80    End If


90    Exit Function

UpdateListRank_Error:

    Dim strES As String
    Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "basShared", "UpdateListRank", intEL, strES, sql


End Function


Public Function GetUniqueID() As Double

    Dim tb As Recordset
    Dim sql As String

10    On Error GoTo GetUniqueID_Error

20    sql = "SELECT RecordId + 1 AS NewId FROM UniqueRecordId " & _
          "UPDATE UniqueRecordId " & _
          "SET RecordId = RecordId + 1 "

30    Set tb = New Recordset
40    Set tb = Cnxn(0).Execute(sql)

50    If tb.EOF Then
60      GetUniqueID = 0
70    Else
80      GetUniqueID = tb!NewId
90    End If

100   Exit Function

GetUniqueID_Error:

    Dim strES As String
    Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basShared", "GetUniqueID", intEL, strES, sql


End Function
Public Function SaveUniqueID(Id As String)

    Dim tb As Recordset
    Dim sql As String

10    On Error GoTo SaveUniqueID_Error

20    sql = "SELECT RecordId FROM UniqueRecordId"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    If tb.EOF Then tb.AddNew
60    tb!RecordId = Id

70    tb.Update



80    Exit Function

SaveUniqueID_Error:

    Dim strES As String
    Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "basShared", "SaveUniqueID", intEL, strES, sql


End Function

Public Sub EnableTestMode(ByVal TestForm As Form)
    Dim cc As Object

10    For Each cc In TestForm
20      If TypeOf cc Is CommandButton Or TypeOf cc Is TextBox Or TypeOf cc Is ComboBox Then
30          cc.BackColor = strTestSystemColor
40      End If
50    Next
60    TestForm.Caption = TestForm.Caption & "   ----- Test Application and Database in use ----- "


End Sub

Public Sub ExportFlexGrid(ByVal objGrid As MSFlexGrid, _
                          Optional ByVal CallingForm As Form = Nothing, _
                          Optional ByVal HeadingMatrix As String = "", _
                          Optional ByVal ImgGreenTick As Image = Nothing, _
                          Optional ByVal ImgRedCross As Image = Nothing, _
                          Optional ByVal IsTATGrid As Boolean = False)

    Dim objXL As Object
    Dim objWB As Object
    Dim objWS As Object
    Dim r As Long
    Dim c As Long

    'Assume the calling form has a MSFlexGrid (grdToExport),
    'CommandButton (cmdXL) and Label (lblExcelInfo) (Visible set to False)
    'In the calling form:
    'Private Sub cmdXL_Click()
    'ExportFlexGrid grdToExport, Me
    'End Sub

10    On Error GoTo ehEFG

    'With CallingForm.lblExcelInfo
    '  .Caption = "Exporting..."
    '  .Visible = True
    '  .Refresh
    'End With

20    Set objXL = CreateObject("Excel.Application")
30    Set objWB = objXL.Workbooks.Add
40    Set objWS = objWB.Worksheets(1)

    Dim intLineCount As Integer
    '****Change: Babar_Ahh Shahzad 2007-11-19
    'Heading for export to excel can be passed as string which would be
    'a string having TABS as column breaks and CR as row break.

50    intLineCount = 0
60    If HeadingMatrix <> "" Then
70      With objWS
            Dim strTokens() As String
80          strTokens = Split(HeadingMatrix, vbCr)
90          intLineCount = UBound(strTokens)

100         For r = LBound(strTokens) To UBound(strTokens) - 1
                'For C = 0 To objGrid.Cols - 1
                'The "'" is required to format the cells as text in Excel
                'otherwise entries like "4/2" are interpreted as a date
110             .Range(.Cells(r + 1, 1), .Cells(r + 1, objGrid.Cols)).MergeCells = True
120             .Range(.Cells(r + 1, 1), .Cells(r + 1, objGrid.Cols)).HorizontalAlignment = 3
130             .Range(.Cells(r + 1, 1), .Cells(r + 1, objGrid.Cols)).Font.Bold = True
140             objWS.Cells(r + 1, 1) = "'" & strTokens(r)

150         Next
160     End With

170   End If

180   With objWS
190     For r = 0 To objGrid.Rows - 1
200         For c = 0 To objGrid.Cols - 1
                'The "'" is required to format the cells as text in Excel
                'otherwise entries like "4/2" are interpreted as a date
210             If r = 0 Then
220                 If IsTATGrid Then
230                     .Range(.Cells(r + 1 + intLineCount, 1), .Cells(r + 1 + intLineCount, objGrid.Cols)).MergeCells = True
240                     .Range(.Cells(r + 1 + intLineCount, 1), .Cells(r + 1 + intLineCount, objGrid.Cols)).HorizontalAlignment = 3
250                 End If
                    'For j = 0 To objGrid.FixedRows
260                 .Range(.Cells(r + 1 + intLineCount, 1), .Cells(r + 1 + intLineCount, objGrid.Cols)).Font.Bold = True
270                 .Cells(r + 1 + intLineCount, c + 1) = "'" & objGrid.TextMatrix(r, c)
                    'Next
280             Else
290                 objGrid.Row = r: objGrid.Col = c
300                 If Not ImgGreenTick Is Nothing Then
310                     If objGrid.CellPicture = ImgGreenTick.Picture Then
320                         .Cells(r + 1 + intLineCount, c + 1).Font.Name = "Marlett"
330                         .Cells(r + 1 + intLineCount, c + 1) = "a"
340                     End If
350                 ElseIf Not ImgRedCross Is Nothing Then
360                     If objGrid.CellPicture = ImgRedCross.Picture Then
370                         .Cells(r + 1 + intLineCount, c + 1).Font.Name = "Marlett"
380                         .Cells(r + 1 + intLineCount, c + 1) = "r"
390                     End If
400                 Else
410                     .Cells(r + 1 + intLineCount, c + 1) = "'" & objGrid.TextMatrix(r, c)
420                 End If
430             End If

440         Next
450     Next

460     .Cells.Columns.AutoFit
470   End With

480   objXL.Visible = True

490   Set objWS = Nothing
500   Set objWB = Nothing
510   Set objXL = Nothing

    'CallingForm.lblExcelInfo.Visible = False

520   Exit Sub

ehEFG:
    Dim er As Long
    Dim es As String

530   er = Err.Number
540   es = Err.Description

    'iMsg es
550   frmMsgBox.Msg es

    'With CallingForm.lblExcelInfo
    '  .Caption = "Error " & Format(er)
    '  .Refresh
    '  t = Timer
    '  Do While Timer - t < 1: Loop
    '  .Visible = False
    'End With

560   Exit Sub

End Sub

Public Function PopulateGenericList(ByVal ListType As String, _
                                    ByRef Combo As ComboBox, _
                                    Optional ByVal KeyCol As String = "Description") _
                                    As Boolean

10    On Error GoTo PopulateGenericList_Error

    Dim sql As String
    Dim tb As Recordset

20    sql = "SELECT " & KeyCol & " KeyCol FROM Lists WHERE " & _
          "ListType = '%listtype' " & _
          "AND InUse = 1 " & _
          "ORDER BY ListOrder"
30    sql = Replace(sql, "%listtype", ListType)

40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    Combo.Clear
70    Combo.AddItem ""
80    If tb.EOF And tb.BOF Then
90      PopulateGenericList = False
100     Exit Function
110   Else

120     While Not tb.EOF
130         Combo.AddItem tb!KeyCol & ""
140         tb.MoveNext
150     Wend

160     tb.Close
170     Set tb = Nothing

180     PopulateGenericList = True
190   End If


200   Exit Function

PopulateGenericList_Error:

210   PopulateGenericList = False
    Dim strES As String
    Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "basShared", "PopulateGenericList", intEL, strES, sql

End Function


Public Function PopulateGenericList_2Cols(ByVal ListType As String, _
                                    ByRef Combo As ComboBox, _
                                    Optional ByVal KeyCol As String = "Code", _
                                    Optional ByVal KeyCol2 As String = "Description") _
                                    As Boolean
    Dim sql As String
    Dim tb As Recordset

20    sql = "SELECT " & KeyCol & " KeyCol, " & KeyCol2 & " KeyCol2  FROM Lists WHERE " & _
          "ListType = '%listtype' " & _
          "AND InUse = 1 " & _
          "ORDER BY ListOrder"
30    sql = Replace(sql, "%listtype", ListType)

40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    Combo.Clear
70    Combo.AddItem ""
80    If tb.EOF And tb.BOF Then
90      PopulateGenericList_2Cols = False
100     Exit Function
110   Else

120     While Not tb.EOF
130         Combo.AddItem tb!KeyCol & "" & " - " & tb!KeyCol2 & ""
140         tb.MoveNext
150     Wend

160     tb.Close
170     Set tb = Nothing

180     PopulateGenericList_2Cols = True
190   End If


End Function



Public Sub InitializeGridCodes(Grid As MSFlexGrid)
10    With Grid
20      .Clear
30      .Rows = 2
40      .Cols = 7
50      .FixedCols = 0
60      .FixedRows = 1
70      .Rows = 1
80      .ColWidth(0) = 1200
90      .ColWidth(1) = Grid.Width - 1200 - 250
100     .ColWidth(2) = 0
110     .ColWidth(3) = 0
120     .ColWidth(4) = 0
130     .ColWidth(5) = 0
140     .ColWidth(6) = 0
'ali
150     .TextMatrix(0, 0) = "Code"
'---------
160     .TextMatrix(0, 1) = "Description"
170     .TextMatrix(0, 2) = "Unique ID"
180     .TextMatrix(0, 3) = "Tissue Type ID"
190     .TextMatrix(0, 4) = "Tissue Type Letter"
200     .TextMatrix(0, 5) = "Tissue Path"
210     .TextMatrix(0, 6) = "Tissue Type List Id"
220     .GridLines = flexGridNone

230   End With
End Sub

Public Sub UpdateLoggedOnUsers()

    Dim tb As Recordset
    Dim sql As String
    Dim MyMachineName As String



10    On Error GoTo UpdateLoggedOnUsers_Error

20    MyMachineName = vbGetComputerName()

30    sql = "SELECT * FROM LoggedOnUsers WHERE " & _
          "MachineName = N'" & MyMachineName & "' " & _
          "AND AppName = N'Histology'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If tb.EOF Then
70      tb.AddNew
80    End If
90    tb!MachineName = MyMachineName
100   tb!AppName = "Histology"
110   tb!UserName = UserName
120   tb.Update


130   Exit Sub

UpdateLoggedOnUsers_Error:

    Dim strES As String
    Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "basShared", "UpdateLoggedOnUsers", intEL, strES, sql


End Sub

Public Sub NameLostFocus(ByRef strName As String, _
                         ByRef strSex As String)

    Dim ForeName As String
    Dim tb As New Recordset
    Dim sql As String

10    On Error GoTo NameLostFocus_Error

20    strName = Replace(strName, ",", "")

30    strName = initial2upper(strName)

40    ForeName = strName

50    If ForeName = "" Then
60      strSex = ""
70    End If

80    sql = "SELECT * from SexNames WHERE " & _
          "Name = N'" & AddTicks(ForeName) & "'"
90    Set tb = New Recordset
100   RecOpenServer 0, tb, sql
110   If tb.EOF Then
120     If strSex <> "" Then
130         tb.AddNew
140         tb!Name = ForeName
150         tb!Sex = UCase$(Left$(strSex, 1))
160         tb.Update
170     End If
180   Else
190     Select Case UCase(tb!Sex & "")
        Case "M": strSex = "Male"
200     Case "F": strSex = "Female"
210     End Select
220   End If

230   Exit Sub

NameLostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "basShared", "NameLostFocus", intEL, strES, sql

End Sub

Public Sub SexLostFocus(ByVal tSex As TextBox, ByVal tName As TextBox)

    Dim sql As String
    Dim tb As New Recordset
    Dim ForeName As String

10    On Error GoTo SexLostFocus_Error

20    If Trim$(tSex) = "" Then Exit Sub
30    If UCase$(Left$(tSex, 1)) <> "Female" And UCase$(Left$(tSex, 1)) <> "Male" Then Exit Sub

40    ForeName = tName
50    If ForeName = "" Then Exit Sub

60    sql = "SELECT * from SexNames WHERE " & _
          "Name = N'" & AddTicks(ForeName) & "'"
70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    If tb.EOF Then
100     tb.AddNew
110   End If

120   tb!Name = ForeName
130   tb!Sex = UCase$(Left$(tSex, 1))
140   tb.Update

150   Exit Sub

SexLostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "basShared", "SexLostFocus", intEL, strES, sql


End Sub

Public Function DisplayReport(ReportNumber As String, PageNo As String) As Boolean

    Dim strReportData As String
    Dim tb As Recordset
    Dim sql As String
    Dim BinaryStream As New Stream
    Dim baEnCrypt() As Byte
    Dim sEncrypt As String
    Dim OriginalSize As Long

10    On Error GoTo DisplayReport_Error

20    sql = "SELECT RepNo, ReportRTF, ReportParameters, UncompressedSize From Reports WHERE " & _
          "SampleId = N'" & CaseNo & "' " & _
          "AND RepNo = '%reportnumber' " & _
          "AND PageNumber = '" & PageNo & "' " & _
          "AND ReportRTF IS NOT NULL"
    '"AND PrintTime = '" & Format(PrintTime, "dd/MMM/yyyy HH:mm:ss") & "' " & _

30         sql = Replace(sql, "%reportnumber", ReportNumber)
40    Set tb = New Recordset

50    RecOpenClient 0, tb, sql
60    If tb.EOF Then
70      DisplayReport = False
80      Exit Function
90    End If


100   strReportData = tb!ReportRTF.GetChunk(10000000#) & ""

110   OriginalSize = tb!UncompressedSize & 0

120   baEnCrypt = strReportData

130   UnCompressBytes baEnCrypt, OriginalSize

140   Set BinaryStream = CreateObject("ADODB.Stream")
150   BinaryStream.Type = adTypeBinary
160   BinaryStream.Open
170   BinaryStream.Write baEnCrypt
180   BinaryStream.Position = 0
190   BinaryStream.Type = adTypeText
200   BinaryStream.Charset = "us-ascii"    'unicode
210   sEncrypt = BinaryStream.ReadText    'stream in text
220   BinaryStream.Close
230   Set BinaryStream = Nothing

240   If strReportData <> "" Then

        'Call UnCompressToFile(strReportData, strFileName)
        'frmReportViewer.rtb.LoadFile strFileName, rtfRTF
250     frmReportViewer.rtb.TextRTF = sEncrypt

260   End If

270   tb.Close
280   Set tb = Nothing

290   DisplayReport = True
300   Exit Function

DisplayReport_Error:
    Dim strES As String
    Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "basShared", "DisplayReport", intEL, strES, sql

340   DisplayReport = False

End Function

Public Sub SaveRTF(ByVal SampleID As String, _
                   ByVal YYYY As String, _
                   ByVal PageNumber As String, _
                   ByVal Initiator As String, _
                   Optional RepNo As String = "")

    Dim tb As Recordset
    Dim sql As String
    Dim strCompressed As String
    Dim BinaryStream As New Stream
    Dim baEnCrypt() As Byte

10    On Error GoTo SaveRTF_Error

20    frmRichText.rtb.SelStart = 0
30    frmRichText.rtb.SelLength = 100000




40    Set BinaryStream = CreateObject("ADODB.Stream")
50    BinaryStream.Type = adTypeText
60    BinaryStream.Charset = "us-ascii"
70    BinaryStream.Open
80    BinaryStream.WriteText frmRichText.rtb.TextRTF
90    BinaryStream.Position = 0
100   BinaryStream.Type = adTypeBinary
    Dim OriginalSize As Long
110   OriginalSize = BinaryStream.Size
120   ReDim baEnCrypt(OriginalSize - 1) As Byte
130   baEnCrypt = BinaryStream.Read    'stream out ascii byte array
140   BinaryStream.Close

150   CompressBytes baEnCrypt

160   strCompressed = baEnCrypt
    'strCompressed = Format(OriginalSize, "000000") & strCompressed


170   sql = "SELECT * FROM Reports WHERE 0 = 1"
180   Set tb = New Recordset
190   RecOpenServer 0, tb, sql
200   tb.AddNew
210   tb!SampleID = SampleID
220   tb!Year = YYYY
230   tb!Initiator = Initiator
240   tb!PrintTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
250   tb!RepNo = SampleID & RepNo
260   tb!ReportRTF.AppendChunk strCompressed
270   tb!UncompressedSize = OriginalSize
280   tb!Printer = Printer.DeviceName
290   tb!PageNumber = PageNumber
    '180   tb!ReportParameters = "HOSPITALNAME" & vbTab & GetOptionSetting("HOSPITALNAME", "") & vbTab & _
     '                            "HOSPITALADDRESS" & vbTab & GetOptionSetting("HOSPITALADDRESS", "") & vbTab & _
     '                            "HOSPITALDISTRICT" & vbTab & GetOptionSetting("HOSPITALDISTRIC", "") & vbTab & _
     '                            "HOSPITALREGION" & vbTab & GetOptionSetting("HOSPITALREGION", "") & vbTab & _
     '                            "DEPARTMENT" & vbTab & Department & vbTab & _
     '                            "INITIATOR" & vbTab & Initiator
300   tb.Update

310   tb.Close
320   Set tb = Nothing



330   Exit Sub

SaveRTF_Error:

    Dim strES As String
    Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "basShared", "SaveRTF", intEL, strES, sql


End Sub

Public Function QueryKnown(ByVal ClinOrGP As String, _
                           ByVal CodeOrText As String, _
                           Optional ByVal Source As String) _
                           As String
'Returns either "" = not known
'        or CodeOrText = known

    Dim SourceCode As String
    Dim County As String
    Dim Original As String
    Dim tb As New Recordset
    Dim sql As String

10    On Error GoTo QueryKnown_Error

20    QueryKnown = ""
30    Original = CodeOrText

40    CodeOrText = Trim$(UCase$(CodeOrText))
50    If CodeOrText = "" Then Exit Function

60    If UCase(ClinOrGP) = "GP" Then
70      If Not IsMissing(Source) Then
            'County = Trim$(Mid$(Source, InStr(Source, "-") + 1))
80          County = Source
90      End If

100     sql = "SELECT * from GPs WHERE " & _
              "County = N'" & County & "' And InUse = '1'"
110     Set tb = New Recordset
120     RecOpenServer 0, tb, sql
130     Do While Not tb.EOF
140         If Trim$(UCase$(tb!Code)) = CodeOrText Then
150             QueryKnown = tb!GPName
160             Exit Function
170         ElseIf Trim$(UCase$(tb!GPName)) = CodeOrText Then
180             QueryKnown = Original
190             Exit Function
200         End If
210         tb.MoveNext
220     Loop

230   ElseIf UCase(ClinOrGP) = "CLINICIAN" Or UCase(ClinOrGP) = "CORONER" _
           Or UCase(ClinOrGP) = "WARD" Then
240     If Not IsMissing(Source) Then
250         SourceCode = ListCodeFor("Source", Source)
260     Else
270         SourceCode = "T"
280     End If

290     sql = "SELECT * from SourceItemLists WHERE " & _
              "ListType = '" & ClinOrGP & "' " & _
              "AND (Source = N'" & SourceCode & "' OR Source = '') AND InUse = '1'"
300     Set tb = New Recordset
310     RecOpenServer 0, tb, sql
320     Do While Not tb.EOF
330         If Trim$(UCase$(tb!Code)) = CodeOrText Then
340             QueryKnown = tb!Description
350             Exit Function
360         ElseIf Trim$(UCase$(tb!Description)) = CodeOrText Then
370             QueryKnown = Original
380             Exit Function
390         End If
400         tb.MoveNext
410     Loop
420   End If



430   Exit Function

QueryKnown_Error:

    Dim strES As String
    Dim intEL As Integer

440   intEL = Erl
450   strES = Err.Description
460   LogError "basShared", "QueryKnown", intEL, strES, sql

End Function

Public Function MarkGridRow(flxGrid As MSFlexGrid, GridRow As Integer, Optional intBackColor As Long = 0, Optional boolStrikeThru As Boolean = False, Optional boolBold As Boolean = False, Optional boolItalic As Boolean = False) As Boolean
10    If GridRow > flxGrid.Rows Then Exit Function
20    If GridRow = 0 Then Exit Function
    Dim X As Integer
30    flxGrid.Row = GridRow
40    For X = 0 To flxGrid.Cols - 1
50      With flxGrid
60          .Col = X
70          .CellBackColor = intBackColor
80          .CellFontStrikeThrough = boolStrikeThru
90          .CellFontBold = boolBold
100         .CellFontItalic = boolItalic
110     End With

120   Next X
130   MarkGridRow = True
End Function

Public Function VerifyCaseIdFormat(ByVal strCaseId As String) As Boolean

10       On Error GoTo VerifyCaseIdFormat_Error

20    VerifyCaseIdFormat = False
      'C00002/23
'      If Left(strCaseId, 2) = "PA" Or Left(strCaseId, 2) = "TA" Or Left(strCaseId, 2) = "MA" Then 'Portlaoise/Tullamore/Mulingar Autopsies
'120       If Len(strCaseId) = 9 Then 'Lenght = 9
130           If IsNumeric(Right(strCaseId, Len(strCaseId) - 2)) Then 'rest is numeric
140               If Val(Right(strCaseId, 2)) <= Val(Format(Now, "YY")) Then
150                   VerifyCaseIdFormat = True
160               End If
170           End If
'180       End If
'190   Else
'200       VerifyCaseIdFormat = False
'210   End If


220      Exit Function

VerifyCaseIdFormat_Error:

Dim strES As String
Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "basShared", "VerifyCaseIdFormat", intEL, strES

End Function

Public Function CaseIdDemoEntered(ByVal strCaseId As String) As Boolean
          Dim sql As String
          Dim tb As Recordset

10       On Error GoTo CaseIdDemoEntered_Error

20        sql = "SELECT CaseID FROM Demographics WHERE CaseId = N'" & strCaseId & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If Not tb.EOF Then
60            CaseIdDemoEntered = True
70        Else
80            CaseIdDemoEntered = False
90        End If

100      Exit Function

CaseIdDemoEntered_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basShared", "CaseIdDemoEntered", intEL, strES, sql

End Function

Public Function FixComboWidth(Combo As ComboBox) As Boolean

      Dim i As Integer
      Dim ScrollWidth As Long

10    With Combo
20        For i = 0 To .ListCount
30            If .Parent.TextWidth(.List(i)) > ScrollWidth Then
40                ScrollWidth = .Parent.TextWidth(.List(i))
50            End If
60        Next i
70        FixComboWidth = SendMessage(.hwnd, CB_SETDROPPEDWIDTH, _
                                      ScrollWidth / 15 + 30, 0) > 0

80    End With

End Function

Public Function IsTestSystemExpired(NoOfEnteries As Integer) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo IsTestSystemExpired_Error

20    sql = "SELECT Count(CaseId) AS Cnt FROM Demographics"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb!Cnt > NoOfEnteries Then
60        IsTestSystemExpired = True
70    Else
80        IsTestSystemExpired = False
90    End If

100   Exit Function

IsTestSystemExpired_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basShared", "IsTestSystemExpired", intEL, strES

End Function

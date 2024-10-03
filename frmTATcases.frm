VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTATcases 
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1635
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   17
      Top             =   5460
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   30
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fetching results ..."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -15
         TabIndex        =   19
         Top             =   285
         Width           =   3840
      End
   End
   Begin VB.Frame fraDates 
      Caption         =   "Between Dates"
      Height          =   2175
      Left            =   90
      TabIndex        =   4
      Top             =   330
      Width           =   5415
      Begin VB.ComboBox cmbPathologist 
         Height          =   315
         Left            =   1770
         TabIndex        =   22
         Top             =   1830
         Width           =   3550
      End
      Begin VB.ComboBox cmbTissue 
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   1485
         Width           =   3550
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "&Calculate"
         Default         =   -1  'True
         Height          =   975
         Left            =   4080
         Picture         =   "frmTATcases.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Search"
         Top             =   720
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   360
         Left            =   2160
         TabIndex        =   6
         Top             =   735
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110100481
         CurrentDate     =   37951
      End
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   735
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110100481
         CurrentDate     =   37951
      End
      Begin VB.Label lblPathologist 
         Caption         =   "Pathologist"
         Height          =   210
         Left            =   1800
         TabIndex        =   23
         Top             =   1605
         Width           =   3510
      End
      Begin VB.Label lblTissueType 
         AutoSize        =   -1  'True
         Caption         =   "Tissue Type"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   975
      Left            =   6915
      Picture         =   "frmTATcases.frx":03E8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   465
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   8115
      Picture         =   "frmTATcases.frx":0803
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   465
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   90
      TabIndex        =   1
      Top             =   2580
      Width           =   5415
      Begin VB.TextBox txtDaysUnReported 
         Height          =   285
         Left            =   2565
         TabIndex        =   15
         Text            =   "10"
         Top             =   315
         Width           =   525
      End
      Begin VB.Label lblDays1 
         Caption         =   "days"
         Height          =   240
         Left            =   3195
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblUnReported1 
         Caption         =   "Cases unreported after"
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   345
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   975
      Left            =   5715
      Picture         =   "frmTATcases.frx":0B45
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   465
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   90
      TabIndex        =   10
      Top             =   105
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4650
      Left            =   90
      TabIndex        =   13
      Top             =   3540
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   8202
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   16711680
      ForeColorSel    =   65280
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "Tissue                                                                         |<Case id           |<Turnaround time"
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   150
      TabIndex        =   12
      Top             =   8265
      Width           =   855
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1110
      TabIndex        =   11
      Top             =   8265
      Width           =   525
   End
End
Attribute VB_Name = "frmTATcases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pCode As String
Private mclsToolTip As New clsToolTip

Public Property Let Code(ByVal Id As String)

10  pCode = Id

End Property

Private Sub cmdCalc_Click()



10  ClearFGrid g

20  Select Case pCode
    Case "T"
30      FillGridT
40  Case "P"
50      FillGridP
60  End Select

70  cmdPrint.Enabled = True
80  g.Visible = True
90  If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1

End Sub

Private Sub FillGridP()
    Dim sn As New Recordset
    Dim sql As String
    Dim strG As String

10  On Error GoTo FillGridP_Error

20  sql = "select l.description, C.CaseId,([dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate)/24) as TAT " & _
          "From cases c left join CaseListLink ct on c.caseid = ct.caseid " & _
          "inner join lists l on ct.listid = l.listid " & _
          "Where OrigValDate Is Not Null " & _
          "AND ct.Type = 'P' " & _
          "AND SampleReceived BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
          "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59' " & _
          "AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * " & txtDaysUnReported & ")"

30  If Len(cmbPathologist) > 0 Then
40      sql = sql & " AND C.OrigValBy = '" & AddTicks(cmbPathologist) & "' "
50  End If
60  Set sn = New Recordset
70  RecOpenClient 0, sn, sql

80  If Not sn.EOF Then
90      pbProgress.Max = sn.RecordCount + 1
100     fraProgress.Visible = True
110     Do While Not sn.EOF
120         pbProgress.Value = pbProgress.Value + 1
130         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
140         lblProgress.Refresh
150         strG = sn!Description & "" & vbTab & sn!CaseId & "" & vbTab & sn!TAT & ""
160         g.AddItem strG
170         sn.MoveNext
180     Loop
190     fraProgress.Visible = False
200     pbProgress.Value = 1

210 End If

220 Exit Sub

FillGridP_Error:

    Dim strES As String
    Dim intEL As Integer

230 intEL = Erl
240 strES = Err.Description
250 LogError "frmTATcases", "FillGridP", intEL, strES, sql
End Sub

Private Sub FillGridT()

    Dim sn As New Recordset
    Dim sql As String
    Dim strG As String

10  On Error GoTo FillGridT_Error

20  sql = "SELECT l.description, C.CaseId, ([dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate)/24) as TAT From cases C left join casetree CT on C.CaseId = CT.caseid " & _
          "inner join lists L on ct.tissuetypelistid = l.listid " & _
          "Where OrigValDate Is Not Null " & _
          "AND SampleReceived BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
          "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59' " & _
          "AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * " & txtDaysUnReported & ")"

30  If Len(cmbTissue) > 0 Then
40      sql = sql & " AND L.Description = '" & cmbTissue & "' "
50  End If
60  Set sn = New Recordset
70  RecOpenClient 0, sn, sql

80  If Not sn.EOF Then
90      pbProgress.Max = sn.RecordCount + 1
100     fraProgress.Visible = True
110     Do While Not sn.EOF
120         pbProgress.Value = pbProgress.Value + 1
130         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
140         lblProgress.Refresh
150         strG = sn!Description & "" & vbTab & sn!CaseId & "" & vbTab & sn!TAT & ""
160         g.AddItem strG
170         sn.MoveNext
180     Loop
190     fraProgress.Visible = False
200     pbProgress.Value = 1

210 End If

220 Exit Sub

FillGridT_Error:

    Dim strES As String
    Dim intEL As Integer

230 intEL = Erl
240 strES = Err.Description
250 LogError "frmTATcases", "FillGridT", intEL, strES, sql

End Sub

Private Sub cmdExit_Click()
10  Unload Me
End Sub

Private Sub cmdExport_Click()
10  On Error GoTo cmdExport_Click_Error

20  ExportFlexGrid g, Me, , , , True

30  Exit Sub

cmdExport_Click_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmTATcases", "cmdExport_Click", intEL, strES
End Sub

Private Sub cmdPrint_Click()
    Const GAP = 60

    Dim xmax As Single
    Dim ymax As Single
    Dim xmin As Single
    Dim ymin As Single
    Dim X As Single
    Dim c As Integer


10  On Error GoTo cmdPrint_Click_Error

20  xmin = 1000
30  ymin = 1660

    Dim lRowsPrinted As Long, lRowsPerPage As Long
    Dim lThisRow As Long, lNumRows As Long
    Dim lPrinterPageHeight As Long
    Dim lPrintPage As Long
    Dim lNoOfPages As Long

40  g.TopRow = 1
50  lNumRows = g.Rows - 1
60  lPrinterPageHeight = Printer.Height

70  lRowsPrinted = 1

80  xmax = xmin + GAP
90  For c = 0 To g.Cols - 1
100     If g.ColWidth(c) <> 0 Then
110         xmax = xmax + g.ColWidth(c) + 2 * GAP
120     End If
130 Next c

140 lPrintPage = 1

150 Do
160     Printer.Orientation = 1
170     lRowsPerPage = 45
180     lNoOfPages = Int(lNumRows / lRowsPerPage) + 1


190     Do

200         With g

210             PrintHeadingWorkLog "Page " & lPrintPage & " of " & lNoOfPages, "Statistics", 9

                ' Print each row.
220             Printer.CurrentY = ymin

230             Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)

240             Printer.CurrentY = Printer.CurrentY + GAP

250             X = xmin + GAP
260             For c = 0 To .Cols - 1
270                 Printer.CurrentX = X
280                 If .ColWidth(c) <> 0 Then
290                     PrintText BoundedText(Printer, .TextMatrix(0, c), .ColWidth(c)), "MS Sans Serif", , True
300                     X = X + .ColWidth(c) + 2 * GAP
310                 End If

320             Next c
330             Printer.CurrentY = Printer.CurrentY + GAP

                ' Move to the next line.
340             PrintText vbCrLf

350             For lThisRow = lRowsPrinted To lRowsPerPage * lPrintPage

360                 If (lThisRow - 1) < lNumRows Then
370                     If lThisRow > 0 Then Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)
380                     Printer.CurrentY = Printer.CurrentY + GAP

                        ' Print the entries on this row.
                        ' Print the entries on this row.
390                     X = xmin + GAP
400                     For c = 0 To .Cols - 1
410                         Printer.CurrentX = X
420                         If .ColWidth(c) <> 0 Then
430                             PrintText BoundedText(Printer, .TextMatrix(lThisRow, c), .ColWidth(c)), "MS Sans Serif", 8
440                             X = X + .ColWidth(c) + 2 * GAP
450                         End If

460                     Next c
470                     Printer.CurrentY = Printer.CurrentY + GAP

                        ' Move to the next line.
480                     PrintText vbCrLf

490                     lRowsPrinted = lRowsPrinted + 1
500                 Else
510                     Exit Do
520                 End If
530             Next
540         End With
550     Loop While (lRowsPrinted - 1) < lRowsPerPage * lPrintPage

560     ymax = Printer.CurrentY

        ' Draw a box around everything.
570     Printer.Line (xmin, ymin)-(xmax, ymax), , B

        ' Draw lines between the columns.
580     X = xmin
590     For c = 0 To g.Cols - 2
600         If g.ColWidth(c) <> 0 Then
610             X = X + g.ColWidth(c) + 2 * GAP
620             Printer.Line (X, ymin)-(X, ymax)
630         End If
640     Next c

650     Printer.EndDoc
660     lPrintPage = lPrintPage + 1

670 Loop While (lRowsPrinted - 1) < lNumRows

680 Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

690 intEL = Erl
700 strES = Err.Description
710 LogError "frmTATcases", "cmdPrint_Click", intEL, strES


End Sub

Private Sub Form_Load()
'frmTATcases_ChangeLanguage

10  calTo = Format(Now, "dd/MM/yyyy")
20  calFrom = Format(Now - 7, "dd/MM/yyyy")
30  lblLoggedIn = UserName
40  g.MergeCells = flexMergeRestrictAll
50  g.MergeRow(0) = True
60  g.MergeCol(0) = True
70  g.Col = 0
80  g.Row = 0
90  g.CellAlignment = flexAlignCenterCenter

100 If pCode = "P" Then
110     FillPathologist
120     cmbPathologist.Visible = True
130     lblPathologist.Visible = True
140     lblTissueType.Visible = False
150     cmbTissue.Visible = False
160     lblPathologist.Left = 240
170     cmbPathologist.Left = 240
180     lblPathologist.Top = 1230
190     cmbPathologist.Top = 1485
200 Else
210     cmbPathologist.Visible = False
220     cmdCalc.Top = 240
230     fraDates.Height = 2175
240     lblPathologist.Visible = False
250     cmdPrint.Top = 550
260     cmdExit.Top = 550
270     cmdExport.Top = 550
280     g.Top = 3420
290     FillTissueList
300     cmbTissue.Visible = True
310     lblTissueType.Visible = True
320 End If

330 txtDaysUnReported = GetOptionSetting("DefaultUnReportedDaysTAT", "7")

340 loadtooltip

350 If blnIsTestMode Then EnableTestMode Me

End Sub

Private Sub FillTissueList()
    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo LoadOrientaion_Error

20  cmbTissue.AddItem ""
30  sql = "SELECT * FROM Lists WHERE ListType = 'T' order by Description"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql

60  Do While Not tb.EOF
70      cmbTissue.AddItem tb!Description & ""
80      tb.MoveNext
90  Loop
100 cmbTissue.ListIndex = -1

110 Exit Sub

LoadOrientaion_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmTATcases", "FillTissueList", intEL, strES, sql

End Sub

Private Sub FillPathologist()

    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo FillPathologist_Error

20  sql = "SELECT * FROM Users WHERE AccessLevel = 'Consultant'"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  cmbPathologist.AddItem ""
60  Do While Not tb.EOF
70      cmbPathologist.AddItem tb!UserName & ""
80      cmbPathologist.ItemData(cmbPathologist.NewIndex) = tb!UserId & ""
90      tb.MoveNext
100 Loop
110 cmbPathologist.ListIndex = -1

120 Exit Sub

FillPathologist_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmTATcases", "FillPathologist", intEL, strES, sql


End Sub

Private Sub loadtooltip()

10  With mclsToolTip
        '
        '
20      Call .Create(Me)
        '

        '
30      .MaxTipWidth = 240
        '


40      Call .AddTool(g)

50  End With
End Sub

Private Sub txtDaysUnReported_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

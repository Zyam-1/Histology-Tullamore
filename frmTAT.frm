VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13650
   Icon            =   "frmTAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4620
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   18
      Top             =   4860
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   240
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
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   975
      Left            =   9960
      Picture         =   "frmTAT.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   5820
      TabIndex        =   14
      Top             =   240
      Width           =   1875
      Begin VB.OptionButton optCases 
         Caption         =   "% Of Cases"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   1140
         Width           =   1335
      End
      Begin VB.OptionButton optCases 
         Caption         =   "No. Of Cases"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   780
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   12360
      Picture         =   "frmTAT.frx":1014
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   975
      Left            =   11160
      Picture         =   "frmTAT.frx":1356
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame fraDates 
      Caption         =   "Between Dates"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.ComboBox cmbPathologist 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   3550
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "&Calculate"
         Default         =   -1  'True
         Height          =   975
         Left            =   4080
         Picture         =   "frmTAT.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Search"
         Top             =   720
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   360
         Left            =   2160
         TabIndex        =   3
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
         Format          =   147718145
         CurrentDate     =   37951
      End
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   360
         Left            =   240
         TabIndex        =   4
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
         Format          =   147718145
         CurrentDate     =   37951
      End
      Begin VB.Label lblPathologist 
         AutoSize        =   -1  'True
         Caption         =   "Pathologist"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   12
      Top             =   120
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
      Height          =   5445
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   9604
      _Version        =   393216
      Cols            =   17
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
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   11
      Top             =   8160
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   8160
      Width           =   855
   End
End
Attribute VB_Name = "frmTAT"
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
90  If g.Rows > 3 And g.TextMatrix(2, 0) = "" Then g.RemoveItem 2



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
60  LogError "frmTAT", "cmdExport_Click", intEL, strES

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
30  ymin = 1960

    Dim lRowsPrinted As Long, lRowsPerPage As Long
    Dim lThisRow As Long, lNumRows As Long
    Dim lPrinterPageHeight As Long
    Dim lPrintPage As Long
    Dim lNoOfPages As Long

40  g.TopRow = 2
50  lNumRows = g.Rows - 1
60  lPrinterPageHeight = Printer.Height

70  lRowsPrinted = 1



80  xmax = xmin + GAP
90  For c = 0 To g.Cols - 1
100     If g.ColWidth(c) <> 0 Then
110         If c = 0 Then
120             xmax = xmax + g.ColWidth(c) + 840
130         Else
140             xmax = xmax + g.ColWidth(c) + 40    '1 * GAP
150         End If
160     End If
170 Next c

180 lPrintPage = 1

190 Do
200     Printer.Orientation = 2
210     lRowsPerPage = 29
220     lNoOfPages = Int(lNumRows / lRowsPerPage) + 1


230     Do

240         With g

250             PrintHeadingWorkLog "Page " & lPrintPage & " of " & lNoOfPages, "Turnaround Times", 9

260             PrintText Space(9)
270             PrintText FormatString("Results from " & Format(calFrom, "dd/mm/yyyy") & " to " & Format(calTo, "dd/mm/yyyy") & " ", 40, , Alignleft), , 10

280             If cmbPathologist <> "" Then
290                 PrintText FormatString("(Pathologist : " & cmbPathologist & ")", 70, , Alignleft), , 10
300             End If

310             PrintText vbCrLf, , 10

                ' Print each row.
320             Printer.CurrentY = ymin

330             Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)

340             Printer.CurrentY = Printer.CurrentY + GAP

350             X = xmin + GAP

360             Printer.CurrentX = X
370             PrintText FormatString(BoundedText(Printer, .TextMatrix(0, 0), xmax), 300, , AlignCenter), "MS Sans Serif", , True

380             X = X + .ColWidth(0) + 840    '1 * GAP

390             Printer.CurrentY = Printer.CurrentY + GAP

                ' Move to the next line.
400             PrintText vbCrLf

410             For lThisRow = lRowsPrinted To lRowsPerPage * lPrintPage

420                 If (lThisRow - 1) < lNumRows Then
430                     If lThisRow > 0 Then Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)
440                     Printer.CurrentY = Printer.CurrentY + GAP

                        ' Print the entries on this row.
                        ' Print the entries on this row.
450                     X = xmin + GAP
460                     For c = 0 To .Cols - 1
470                         Printer.CurrentX = X
480                         If .ColWidth(c) <> 0 Then
490                             If lThisRow = 1 Then
500                                 If c = 0 Then
510                                     PrintText BoundedText(Printer, .TextMatrix(lThisRow, c), .ColWidth(c) + 800), "MS Sans Serif", 8, True
520                                     X = X + .ColWidth(c) + 840
530                                 Else
540                                     PrintText BoundedText(Printer, .TextMatrix(lThisRow, c), .ColWidth(c)), "MS Sans Serif", 8, True
550                                     X = X + .ColWidth(c) + 40
560                                 End If
570                             Else
580                                 If c = 0 Then
590                                     PrintText BoundedText(Printer, .TextMatrix(lThisRow, c), .ColWidth(c) + 800), "MS Sans Serif", 8
600                                     X = X + .ColWidth(c) + 840
610                                 Else
620                                     PrintText BoundedText(Printer, .TextMatrix(lThisRow, c), .ColWidth(c)), "MS Sans Serif", 8
630                                     X = X + .ColWidth(c) + 40
640                                 End If
650                             End If
660                         End If

670                     Next c
680                     Printer.CurrentY = Printer.CurrentY + GAP

                        ' Move to the next line.
690                     PrintText vbCrLf

700                     lRowsPrinted = lRowsPrinted + 1
710                 Else
720                     Exit Do
730                 End If
740             Next
750         End With
760     Loop While (lRowsPrinted - 1) < lRowsPerPage * lPrintPage

770     ymax = Printer.CurrentY

        ' Draw a box around everything.
780     Printer.Line (xmin, ymin)-(xmax, ymax), , B

        ' Draw lines between the columns.
790     X = xmin
800     For c = 0 To g.Cols - 2
810         If g.ColWidth(c) <> 0 Then
820             If c = 0 Then
830                 X = X + g.ColWidth(c) + 840
840             Else
850                 X = X + g.ColWidth(c) + 40    '1 * GAP
860             End If
870             Printer.Line (X, ymin + 320)-(X, ymax)
880         End If
890     Next c

900     Printer.EndDoc
910     lPrintPage = lPrintPage + 1

920 Loop While (lRowsPrinted - 1) < lNumRows




930 Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

940 intEL = Erl
950 strES = Err.Description
960 LogError "frmTAT", "cmdPrint_Click", intEL, strES


End Sub

Private Sub Form_Load()
    Dim i As Integer
'frmtat_ChangeLanguage

10  calTo = Format(Now, "dd/MM/yyyy")
20  calFrom = Format(Now - 7, "dd/MM/yyyy")
30  lblLoggedIn = UserName

40  InitializeGrid


50  For i = 0 To 31
60      g.TextMatrix(0, i) = g.TextMatrix(0, 6)
70  Next


80  g.MergeCells = flexMergeRestrictAll
90  g.MergeRow(0) = True
100 g.MergeCol(0) = True
110 g.Col = 0
120 g.Row = 0
130 g.CellAlignment = flexAlignCenterCenter

140 If pCode = "P" Then
150     FillPathologist
160 Else
170     cmbPathologist.Visible = False
180     cmdCalc.Top = 240
190     fraDates.Height = 1335
200     optCases(0).Top = 400
210     optCases(1).Top = 760
220     Frame1.Height = 1335
230     lblPathologist.Visible = False
240     cmdPrint.Top = 550
250     cmdExit.Top = 550
260     cmdExport.Top = 550
270     g.Top = 1800
280 End If

290 loadtooltip

300 If blnIsTestMode Then EnableTestMode Me
End Sub

Private Sub GridToolTip(Grid As MSFlexGrid, X As Single, Y As Single)

    Dim lngRow As Long
    Dim lngCol As Long



10  On Error GoTo GridToolTip_Error

20  lngCol = 0
30  lngRow = 0


40  lngCol = Grid.MouseCol


50  lngRow = Grid.MouseRow

60  If lngRow = Grid.Rows Or lngCol = Grid.Cols Then

        ' Off the grid just blank the tooltip
70      Grid.ToolTipText = vbNullString

80  Else


90      mclsToolTip.ToolText(Grid) = Replace(Grid.TextMatrix(lngRow, lngCol), "<<tab>>", " ")

100 End If




110 Exit Sub

GridToolTip_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmTAT", "GridToolTip", intEL, strES


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
150 LogError "frmTAT", "FillPathologist", intEL, strES, sql


End Sub


Private Sub InitializeGrid()
    Dim i As Integer
10  With g
20      .Rows = 3: .FixedRows = 2
30      .Cols = 32: .FixedCols = 0
40      .Rows = 2
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 6) = " Completed By Day": .ColWidth(6) = 1000: .ColAlignment(6) = flexAlignLeftCenter
70      .TextMatrix(1, 0) = "Description": .ColWidth(0) = 1550: .ColAlignment(0) = flexAlignLeftCenter
80      For i = 1 To .Cols - 2
90          .TextMatrix(1, i) = i: .ColWidth(i) = 360: .ColAlignment(i) = flexAlignLeftCenter
100     Next
110     .TextMatrix(1, 31) = "> 30": .ColWidth(i) = 410: .ColAlignment(i) = flexAlignLeftCenter

120 End With

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
Private Sub FillGridT()

    Dim sn As New Recordset
    Dim sql As String
    Dim s As String
    Dim tot As Long


10  On Error GoTo FillGridT_Error

20  sql = "select l.description, "
30  sql = sql & "sum(case "
40  sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 1) then 1 "
50  sql = sql & "else 0 "
60  sql = sql & "end) as 'Day1', "
70  sql = sql & "sum(case "
80  sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 1) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 2) then 1 "
90  sql = sql & "else 0 "
100 sql = sql & "end)as 'Day2', "
110 sql = sql & "sum(case "
120 sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 2) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 3) then 1 "
130 sql = sql & "else 0 "
140 sql = sql & "end)as 'Day3', "
150 sql = sql & "sum(case "
160 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 3) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 4) then 1 "
170 sql = sql & "   else 0 "
180 sql = sql & "end)as 'Day4', "
190 sql = sql & "sum(case "
200 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 4) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 5) then 1 "
210 sql = sql & "   else 0 "
220 sql = sql & "end)as 'Day5', "
230 sql = sql & "sum(case "
240 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 5) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 6) then 1 "
250 sql = sql & "   else 0 "
260 sql = sql & "end)as 'Day6', "
270 sql = sql & "sum(case "
280 sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 6) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 7) then 1 "
290 sql = sql & "else 0 "
300 sql = sql & "end)as 'Day7', "
310 sql = sql & "sum(case "
320 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 7) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 8) then 1 "
330 sql = sql & "   else 0 "
340 sql = sql & "end)as 'Day8', "
350 sql = sql & "sum(case "
360 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 8) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 9) then 1 "
370 sql = sql & "   else 0 "
380 sql = sql & "end)as 'Day9', "
390 sql = sql & "sum(case "
400 sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 9) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 10) then 1 "
410 sql = sql & "else 0 "
420 sql = sql & "end)as 'Day10', "
430 sql = sql & "sum(case "
440 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 10) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 11) then 1 "
450 sql = sql & "   else 0 "
460 sql = sql & "end)as 'Day11', "
470 sql = sql & "sum(case "
480 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 11) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 12) then 1 "
490 sql = sql & "   else 0 "
500 sql = sql & "end)as 'Day12', "
510 sql = sql & "sum(case "
520 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 12) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 13) then 1 "
530 sql = sql & "   else 0 "
540 sql = sql & "end)as 'Day13', "
550 sql = sql & "sum(case "
560 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 13) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 14) then 1 "
570 sql = sql & "   else 0 "
580 sql = sql & "end)as 'Day14', "
590 sql = sql & "sum(case "
600 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 14) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 15) then 1 "
610 sql = sql & "   else 0 "
620 sql = sql & "end)as 'Day15', "
630 sql = sql & "sum(case "
640 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 15) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 16) then 1 "
650 sql = sql & "   else 0 "
660 sql = sql & "end)as 'Day16', "
670 sql = sql & "sum(case "
680 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 16) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 17) then 1 "
690 sql = sql & "   else 0 "
700 sql = sql & "end)as 'Day17', "
710 sql = sql & "sum(case "
720 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 17) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 18) then 1 "
730 sql = sql & "   else 0 "
740 sql = sql & "end)as 'Day18', "
750 sql = sql & "sum(case "
760 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 18) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 19) then 1 "
770 sql = sql & "   else 0 "
780 sql = sql & "end)as 'Day19', "
790 sql = sql & "sum(case "
800 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 19) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 20) then 1 "
810 sql = sql & "   else 0 "
820 sql = sql & "end)as 'Day20', "
830 sql = sql & "sum(case "
840 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 20) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 21) then 1 "
850 sql = sql & "   else 0 "
860 sql = sql & "end)as 'Day21', "
870 sql = sql & "sum(case "
880 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 21) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 22) then 1 "
890 sql = sql & "   else 0 "
900 sql = sql & "end)as 'Day22', "
910 sql = sql & "sum(case "
920 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 22) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 23) then 1 "
930 sql = sql & "   else 0 "
940 sql = sql & "end)as 'Day23', "
950 sql = sql & "sum(case "
960 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 23) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 24) then 1 "
970 sql = sql & "   else 0 "
980 sql = sql & "end)as 'Day24', "
990 sql = sql & "sum(case "
1000 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 24) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 25) then 1 "
1010 sql = sql & "   else 0 "
1020 sql = sql & "end)as 'Day25', "
1030 sql = sql & "sum(case "
1040 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 25) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 26) then 1 "
1050 sql = sql & "   else 0 "
1060 sql = sql & "end)as 'Day26', "
1070 sql = sql & "sum(case "
1080 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 26) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 27) then 1 "
1090 sql = sql & "   else 0 "
1100 sql = sql & "end)as 'Day27', "
1110 sql = sql & "sum(case "
1120 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 27) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 28) then 1 "
1130 sql = sql & "   else 0 "
1140 sql = sql & "end)as 'Day28', "
1150 sql = sql & "sum(case "
1160 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 28) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 29) then 1 "
1170 sql = sql & "   else 0 "
1180 sql = sql & "end)as 'Day29', "
1190 sql = sql & "sum(case "
1200 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 29) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 30) then 1 "
1210 sql = sql & "   else 0 "
1220 sql = sql & "end)as 'Day30', "
1230 sql = sql & "sum(case "
1240 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 30) then 1 "
1250 sql = sql & "   else 0 "
1260 sql = sql & "end)as 'Day31' "
1270 sql = sql & "From cases c "
1280 sql = sql & "left join casetree ct on c.caseid = ct.caseid "
1290 sql = sql & "inner join lists l on ct.tissuetypelistid = l.listid "
1300 sql = sql & "Where OrigValDate Is Not Null "
1310 sql = sql & "AND SampleReceived BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' "
1320 sql = sql & "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59' "
1330 sql = sql & "group by l.description"
1340 Set sn = New Recordset
1350 RecOpenClient 0, sn, sql

1360 If Not sn.EOF Then
1370    pbProgress.Max = sn.RecordCount + 1
1380    fraProgress.Visible = True
1390    Do While Not sn.EOF
1400        pbProgress.Value = pbProgress.Value + 1
1410        lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
1420        lblProgress.Refresh

1430        If optCases(0) Then
1440            s = sn!Description & vbTab & sn!Day1 & vbTab & _
                    sn!Day2 & vbTab & sn!Day3 & vbTab & sn!Day4 & vbTab & sn!Day5 & vbTab & _
                    sn!Day6 & vbTab & sn!Day7 & vbTab & sn!Day8 & vbTab & sn!Day9 & vbTab & _
                    sn!Day10 & vbTab & sn!Day11 & vbTab & _
                    sn!Day12 & vbTab & sn!Day13 & vbTab & sn!Day14 & vbTab & sn!Day15 & vbTab & _
                    sn!Day16 & vbTab & sn!Day17 & vbTab & sn!Day18 & vbTab & sn!Day19 & vbTab & _
                    sn!Day20 & vbTab & sn!Day21 & vbTab & _
                    sn!Day22 & vbTab & sn!Day23 & vbTab & sn!Day24 & vbTab & sn!Day25 & vbTab & _
                    sn!Day26 & vbTab & sn!Day27 & vbTab & sn!Day28 & vbTab & sn!Day29 & vbTab & _
                    sn!Day30 & vbTab & sn!Day31
1450        Else
1460            tot = sn!Day1 + sn!Day2 + sn!Day3 + sn!Day4 + sn!Day5 + sn!Day6 + sn!Day7 + sn!Day8 + sn!Day9 + sn!Day10 + sn!Day11 + _
                      sn!Day12 + sn!Day13 + sn!Day14 + sn!Day15 + sn!Day16 + sn!Day17 + sn!Day18 + sn!Day19 + sn!Day20 + sn!Day21 + _
                      sn!Day22 + sn!Day23 + sn!Day24 + sn!Day25 + sn!Day26 + sn!Day27 + sn!Day28 + sn!Day29 + sn!Day30 + sn!Day31

1470            s = sn!Description & vbTab & Format$(sn!Day1 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day2 / tot, "##.00") * 100 & vbTab & Format$(sn!Day3 / tot, "##.00") * 100 & vbTab & Format$(sn!Day4 / tot, "##.00") * 100 & vbTab & Format$(sn!Day5 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day6 / tot, "##.00") * 100 & vbTab & Format$(sn!Day7 / tot, "##.00") * 100 & vbTab & Format$(sn!Day8 / tot, "##.00") * 100 & vbTab & Format$(sn!Day9 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day10 / tot, "##.00") * 100 & vbTab & Format$(sn!Day11 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day12 / tot, "##.00") * 100 & vbTab & Format$(sn!Day13 / tot, "##.00") * 100 & vbTab & Format$(sn!Day14 / tot, "##.00") * 100 & vbTab & Format$(sn!Day15 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day16 / tot, "##.00") * 100 & vbTab & Format$(sn!Day17 / tot, "##.00") * 100 & vbTab & Format$(sn!Day18 / tot, "##.00") * 100 & vbTab & Format$(sn!Day19 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day20 / tot, "##.00") * 100 & vbTab & Format$(sn!Day21 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day22 / tot, "##.00") * 100 & vbTab & Format$(sn!Day23 / tot, "##.00") * 100 & vbTab & Format$(sn!Day24 / tot, "##.00") * 100 & vbTab & Format$(sn!Day25 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day26 / tot, "##.00") * 100 & vbTab & Format$(sn!Day27 / tot, "##.00") * 100 & vbTab & Format$(sn!Day28 / tot, "##.00") * 100 & vbTab & Format$(sn!Day29 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day30 / tot, "##.00") * 100 & vbTab & Format$(sn!Day31 / tot, "##.00") * 100
1480        End If



1490        g.AddItem s
1500        sn.MoveNext
1510    Loop
1520    fraProgress.Visible = False
1530    pbProgress.Value = 1

1540 End If

1550 Exit Sub

FillGridT_Error:

    Dim strES As String
    Dim intEL As Integer

1560 intEL = Erl
1570 strES = Err.Description
1580 LogError "frmTAT", "FillGridT", intEL, strES, sql

End Sub

Private Sub FillGridP()

    Dim sn As New Recordset
    Dim sql As String
    Dim s As String
    Dim tot As Long

10  On Error GoTo FillGridP_Error

20  sql = "select l.description, "
30  sql = sql & "sum(case "
40  sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 1) then 1 "
50  sql = sql & "else 0 "
60  sql = sql & "end) as 'Day1', "
70  sql = sql & "sum(case "
80  sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 1) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 2) then 1 "
90  sql = sql & "else 0 "
100 sql = sql & "end)as 'Day2', "
110 sql = sql & "sum(case "
120 sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 2) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 3) then 1 "
130 sql = sql & "else 0 "
140 sql = sql & "end)as 'Day3', "
150 sql = sql & "sum(case "
160 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 3) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 4) then 1 "
170 sql = sql & "   else 0 "
180 sql = sql & "end)as 'Day4', "
190 sql = sql & "sum(case "
200 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 4) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 5) then 1 "
210 sql = sql & "   else 0 "
220 sql = sql & "end)as 'Day5', "
230 sql = sql & "sum(case "
240 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 5) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 6) then 1 "
250 sql = sql & "   else 0 "
260 sql = sql & "end)as 'Day6', "
270 sql = sql & "sum(case "
280 sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 6) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 7) then 1 "
290 sql = sql & "else 0 "
300 sql = sql & "end)as 'Day7', "
310 sql = sql & "sum(case "
320 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 7) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 8) then 1 "
330 sql = sql & "   else 0 "
340 sql = sql & "end)as 'Day8', "
350 sql = sql & "sum(case "
360 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 8) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 9) then 1 "
370 sql = sql & "   else 0 "
380 sql = sql & "end)as 'Day9', "
390 sql = sql & "sum(case "
400 sql = sql & "When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 9) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 10) then 1 "
410 sql = sql & "else 0 "
420 sql = sql & "end)as 'Day10', "
430 sql = sql & "sum(case "
440 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 10) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 11) then 1 "
450 sql = sql & "   else 0 "
460 sql = sql & "end)as 'Day11', "
470 sql = sql & "sum(case "
480 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 11) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 12) then 1 "
490 sql = sql & "   else 0 "
500 sql = sql & "end)as 'Day12', "
510 sql = sql & "sum(case "
520 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 12) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 13) then 1 "
530 sql = sql & "   else 0 "
540 sql = sql & "end)as 'Day13', "
550 sql = sql & "sum(case "
560 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 13) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 14) then 1 "
570 sql = sql & "   else 0 "
580 sql = sql & "end)as 'Day14', "
590 sql = sql & "sum(case "
600 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 14) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 15) then 1 "
610 sql = sql & "   else 0 "
620 sql = sql & "end)as 'Day15', "
630 sql = sql & "sum(case "
640 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 15) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 16) then 1 "
650 sql = sql & "   else 0 "
660 sql = sql & "end)as 'Day16', "
670 sql = sql & "sum(case "
680 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 16) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 17) then 1 "
690 sql = sql & "   else 0 "
700 sql = sql & "end)as 'Day17', "
710 sql = sql & "sum(case "
720 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 17) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 18) then 1 "
730 sql = sql & "   else 0 "
740 sql = sql & "end)as 'Day18', "
750 sql = sql & "sum(case "
760 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 18) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 19) then 1 "
770 sql = sql & "   else 0 "
780 sql = sql & "end)as 'Day19', "
790 sql = sql & "sum(case "
800 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 19) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 20) then 1 "
810 sql = sql & "   else 0 "
820 sql = sql & "end)as 'Day20', "
830 sql = sql & "sum(case "
840 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 20) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 21) then 1 "
850 sql = sql & "   else 0 "
860 sql = sql & "end)as 'Day21', "
870 sql = sql & "sum(case "
880 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 21) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 22) then 1 "
890 sql = sql & "   else 0 "
900 sql = sql & "end)as 'Day22', "
910 sql = sql & "sum(case "
920 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 22) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 23) then 1 "
930 sql = sql & "   else 0 "
940 sql = sql & "end)as 'Day23', "
950 sql = sql & "sum(case "
960 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 23) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 24) then 1 "
970 sql = sql & "   else 0 "
980 sql = sql & "end)as 'Day24', "
990 sql = sql & "sum(case "
1000 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 24) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 25) then 1 "
1010 sql = sql & "   else 0 "
1020 sql = sql & "end)as 'Day25', "
1030 sql = sql & "sum(case "
1040 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 25) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 26) then 1 "
1050 sql = sql & "   else 0 "
1060 sql = sql & "end)as 'Day26', "
1070 sql = sql & "sum(case "
1080 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 26) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 27) then 1 "
1090 sql = sql & "   else 0 "
1100 sql = sql & "end)as 'Day27', "
1110 sql = sql & "sum(case "
1120 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 27) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 28) then 1 "
1130 sql = sql & "   else 0 "
1140 sql = sql & "end)as 'Day28', "
1150 sql = sql & "sum(case "
1160 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 28) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 29) then 1 "
1170 sql = sql & "   else 0 "
1180 sql = sql & "end)as 'Day29', "
1190 sql = sql & "sum(case "
1200 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 29) AND [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) <= (24 * 30) then 1 "
1210 sql = sql & "   else 0 "
1220 sql = sql & "end)as 'Day30', "
1230 sql = sql & "sum(case "
1240 sql = sql & "   When [dbo].fnGetHoursExcludingWeekends(SampleReceived,OrigValDate) > (24 * 30) then 1 "
1250 sql = sql & "   else 0 "
1260 sql = sql & "end)as 'Day31' "
1270 sql = sql & "From cases c "
1280 sql = sql & "left join CaseListLink ct on c.caseid = ct.caseid "
1290 sql = sql & "inner join lists l on ct.listid = l.listid "
1300 sql = sql & "Where OrigValDate Is Not Null AND ct.Type = '" & pCode & "' "
1310 sql = sql & "AND SampleReceived BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' "
1320 sql = sql & "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59' "
1330 If cmbPathologist <> "" Then
1340    sql = sql & "AND OrigValBy = '" & AddTicks(cmbPathologist) & "' "
1350 End If
1360 sql = sql & "group by l.description"
1370 Set sn = New Recordset
1380 RecOpenClient 0, sn, sql


1390 If Not sn.EOF Then
1400    pbProgress.Max = sn.RecordCount + 1
1410    fraProgress.Visible = True
1420    Do While Not sn.EOF
1430        pbProgress.Value = pbProgress.Value + 1
1440        lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
1450        lblProgress.Refresh


1460        If optCases(0) Then
1470            s = sn!Description & vbTab & sn!Day1 & vbTab & _
                    sn!Day2 & vbTab & sn!Day3 & vbTab & sn!Day4 & vbTab & sn!Day5 & vbTab & _
                    sn!Day6 & vbTab & sn!Day7 & vbTab & sn!Day8 & vbTab & sn!Day9 & vbTab & _
                    sn!Day10 & vbTab & sn!Day11 & vbTab & _
                    sn!Day12 & vbTab & sn!Day13 & vbTab & sn!Day14 & vbTab & sn!Day15 & vbTab & _
                    sn!Day16 & vbTab & sn!Day17 & vbTab & sn!Day18 & vbTab & sn!Day19 & vbTab & _
                    sn!Day20 & vbTab & sn!Day21 & vbTab & _
                    sn!Day22 & vbTab & sn!Day23 & vbTab & sn!Day24 & vbTab & sn!Day25 & vbTab & _
                    sn!Day26 & vbTab & sn!Day27 & vbTab & sn!Day28 & vbTab & sn!Day29 & vbTab & _
                    sn!Day30 & vbTab & sn!Day31
1480        Else
1490            tot = sn!Day1 + sn!Day2 + sn!Day3 + sn!Day4 + sn!Day5 + sn!Day6 + sn!Day7 + sn!Day8 + sn!Day9 + sn!Day10 + sn!Day11 + _
                      sn!Day12 + sn!Day13 + sn!Day14 + sn!Day15 + sn!Day16 + sn!Day17 + sn!Day18 + sn!Day19 + sn!Day20 + sn!Day21 + _
                      sn!Day22 + sn!Day23 + sn!Day24 + sn!Day25 + sn!Day26 + sn!Day27 + sn!Day28 + sn!Day29 + sn!Day30 + sn!Day31

1500            s = sn!Description & vbTab & Format$(sn!Day1 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day2 / tot, "##.00") * 100 & vbTab & Format$(sn!Day3 / tot, "##.00") * 100 & vbTab & Format$(sn!Day4 / tot, "##.00") * 100 & vbTab & Format$(sn!Day5 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day6 / tot, "##.00") * 100 & vbTab & Format$(sn!Day7 / tot, "##.00") * 100 & vbTab & Format$(sn!Day8 / tot, "##.00") * 100 & vbTab & Format$(sn!Day9 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day10 / tot, "##.00") * 100 & vbTab & Format$(sn!Day11 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day12 / tot, "##.00") * 100 & vbTab & Format$(sn!Day13 / tot, "##.00") * 100 & vbTab & Format$(sn!Day14 / tot, "##.00") * 100 & vbTab & Format$(sn!Day15 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day16 / tot, "##.00") * 100 & vbTab & Format$(sn!Day17 / tot, "##.00") * 100 & vbTab & Format$(sn!Day18 / tot, "##.00") * 100 & vbTab & Format$(sn!Day19 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day20 / tot, "##.00") * 100 & vbTab & Format$(sn!Day21 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day22 / tot, "##.00") * 100 & vbTab & Format$(sn!Day23 / tot, "##.00") * 100 & vbTab & Format$(sn!Day24 / tot, "##.00") * 100 & vbTab & Format$(sn!Day25 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day26 / tot, "##.00") * 100 & vbTab & Format$(sn!Day27 / tot, "##.00") * 100 & vbTab & Format$(sn!Day28 / tot, "##.00") * 100 & vbTab & Format$(sn!Day29 / tot, "##.00") * 100 & vbTab & _
                    Format$(sn!Day30 / tot, "##.00") * 100 & vbTab & Format$(sn!Day31 / tot, "##.00") * 100

1510        End If

1520        g.AddItem s
1530        sn.MoveNext
1540    Loop
1550    fraProgress.Visible = False
1560    pbProgress.Value = 1

1570 End If



1580 Exit Sub

FillGridP_Error:

    Dim strES As String
    Dim intEL As Integer

1590 intEL = Erl
1600 strES = Err.Description
1610 LogError "frmTAT", "FillGridP", intEL, strES, sql


End Sub
Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then
20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
10  mclsToolTip.RemoveTool g
End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10  GridToolTip g, X, Y
End Sub

Private Sub optCases_Click(Index As Integer)
10  ClearFGrid g

20  Select Case pCode
    Case "T"
30      FillGridT
40  Case "P"
50      FillGridP
60  End Select

70  cmdPrint.Enabled = True
80  g.Visible = True
90  If g.Rows > 3 And g.TextMatrix(2, 0) = "" Then g.RemoveItem 2

End Sub





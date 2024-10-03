VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDiscrepancyReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discrepancy Report"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15930
   Icon            =   "frmDiscrepancyReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   15930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   975
      Left            =   13560
      Picture         =   "frmDiscrepancyReport.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   1035
   End
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   6120
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   14760
      Picture         =   "frmDiscrepancyReport.frx":1014
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   975
      Left            =   12360
      Picture         =   "frmDiscrepancyReport.frx":1356
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7335
      Begin VB.ComboBox cmbDiscrepType 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   5700
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "&Calculate"
         Default         =   -1  'True
         Height          =   975
         Left            =   6120
         Picture         =   "frmDiscrepancyReport.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Search"
         Top             =   720
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   360
         Left            =   3240
         TabIndex        =   4
         Top             =   735
         Width           =   2715
         _ExtentX        =   4789
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
         TabIndex        =   5
         Top             =   735
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Discrepancy Type"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7065
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   15540
      _ExtentX        =   27411
      _ExtentY        =   12462
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   13
      Top             =   9720
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   9720
      Width           =   855
   End
End
Attribute VB_Name = "frmDiscrepancyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()
    Dim rsRec As Recordset
    Dim sql As String
    Dim FromTime As String, ToTime As String
    Dim s As String
    Dim TempCaseId As String

10  On Error GoTo cmdCalc_Click_Error

20  g.Visible = False

30  ClearFGrid g

40  FromTime = Format(calFrom, "yyyymmdd") & " 00:00:00"
50  ToTime = Format(calTo, "yyyymmdd") & " 23:59:59"

60  sql = "SELECT * FROM Discrepancy DL " & _
          "INNER JOIN Demographics D ON DL.CaseId = D.CaseId " & _
          "INNER JOIN Cases C ON DL.CaseId = C.CaseId " & _
          "WHERE DL.DateTimeOfRecord BETWEEN '" & FromTime & "' AND '" & ToTime & "' "
70  If cmbDiscrepType <> "" Then
80      sql = sql & " AND DL.DiscrepancyType LIKE '" & AddTicks(cmbDiscrepType) & "' "
90  End If
100 sql = sql & " ORDER BY DL.DateTimeOfRecord desc"
110 Set rsRec = New Recordset
120 RecOpenClient 0, rsRec, sql
130 If Not rsRec.EOF Then
140     pbProgress.Max = rsRec.RecordCount + 1
150     g.Visible = False
160     fraProgress.Visible = True
170     Do While Not rsRec.EOF
180         pbProgress.Value = pbProgress.Value + 1
190         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
200         lblProgress.Refresh

210         If Mid(rsRec!CaseId & "", 2, 1) = "P" Or Mid(rsRec!CaseId & "", 2, 1) = "A" Then
220             TempCaseId = Left(rsRec!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(rsRec!CaseId, 2)
230         Else
240             TempCaseId = Left(rsRec!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(rsRec!CaseId, 2)
250         End If

260         s = TempCaseId & "" & vbTab & rsRec!PatientName & "" & vbTab

270         s = s & rsRec!DiscrepancyType & "" & vbTab & rsRec!PersonResponsible & "" & vbTab
280         s = s & rsRec!NatureOfDiscrepancy & "" & vbTab & rsRec!PersonDealingWith & "" & vbTab
290         s = s & rsRec!Resolution & "" & vbTab & rsRec!DateOfDiscrepancy
300         s = s & vbTab & rsRec!DateOfResolution
310         g.AddItem s
320         rsRec.MoveNext
330     Loop
340     fraProgress.Visible = False
350     pbProgress.Value = 1
360 End If

370 g.Visible = True

380 If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1


390 Exit Sub

cmdCalc_Click_Error:

    Dim strES As String
    Dim intEL As Integer

400 intEL = Erl
410 strES = Err.Description
420 LogError "frmDiscrepancyReport", "cmdCalc_Click", intEL, strES, sql


End Sub

Private Sub cmdExit_Click()
10  Unload Me
End Sub

Private Sub cmdExport_Click()
10  On Error GoTo cmdExport_Click_Error

20  ExportFlexGrid g, Me

30  Exit Sub

cmdExport_Click_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmDiscrepancyReport", "cmdExport_Click", intEL, strES

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
160     Printer.Orientation = 2
170     lRowsPerPage = 29
180     lNoOfPages = Int(lNumRows / lRowsPerPage) + 1


190     Do

200         With g

210             PrintHeadingWorkLog "Page " & lPrintPage & " of " & lNoOfPages, "Discrepancy Report", 9

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
710 LogError "frmDiscrepancyReport", "cmdPrint_Click", intEL, strES

End Sub

Private Sub Form_Load()
'frmDiscrepancyReport_ChangeLanguage
10  calTo = Format(Now, "dd/MM/yyyy")
20  calFrom = Format(Now - 7, "dd/MM/yyyy")
30  lblLoggedIn = UserName

40  FillDiscrepancyType

50  InitializeGrid
60  If blnIsTestMode Then EnableTestMode Me
End Sub

Private Sub FillDiscrepancyType()
    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo FillDiscrepancyType_Error

20  cmbDiscrepType.AddItem ""
30  sql = "SELECT * FROM Lists WHERE ListType = 'DiscrepType'"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql

60  Do While Not tb.EOF
70      cmbDiscrepType.AddItem tb!Description & ""
80      tb.MoveNext
90  Loop

100 Exit Sub

FillDiscrepancyType_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmDiscrepancyReport", "FillDiscrepancyType", intEL, strES, sql

End Sub

Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then
20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub
Private Sub InitializeGrid()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 9: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "Case Id": .ColWidth(0) = 1100: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Patient Name": .ColWidth(1) = 2000: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "Discrepancy": .ColWidth(2) = 2000: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Responsible Person": .ColWidth(3) = 1800: .ColAlignment(3) = flexAlignLeftCenter
160     .TextMatrix(0, 4) = "Corrective Action": .ColWidth(4) = 2000: .ColAlignment(4) = flexAlignLeftCenter
170     .TextMatrix(0, 5) = "Person Dealing With": .ColWidth(5) = 1800: .ColAlignment(5) = flexAlignLeftCenter
180     .TextMatrix(0, 6) = "Resolution": .ColWidth(6) = 1500: .ColAlignment(6) = flexAlignLeftCenter
190     .TextMatrix(0, 7) = "Date of Discrepancy": .ColWidth(7) = 1650: .ColAlignment(7) = flexAlignLeftCenter
200     .TextMatrix(0, 8) = "Resolution Date": .ColWidth(8) = 1400: .ColAlignment(8) = flexAlignLeftCenter

210 End With
End Sub

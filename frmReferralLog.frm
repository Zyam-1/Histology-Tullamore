VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReferralLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Referrals"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13545
   Icon            =   "frmReferralLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   5820
      TabIndex        =   14
      Top             =   300
      Width           =   1815
      Begin VB.OptionButton optSelect 
         Caption         =   "All"
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Complete"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Not Complete"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   975
      Left            =   9720
      Picture         =   "frmReferralLog.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   975
      Left            =   10920
      Picture         =   "frmReferralLog.frx":1014
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   12120
      Picture         =   "frmReferralLog.frx":142F
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   1035
   End
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   5040
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   6
      Top             =   4380
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.Frame fraDates 
      Caption         =   "Between Dates"
      Height          =   1395
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.CommandButton cmdCalc 
         Caption         =   "&Calculate"
         Default         =   -1  'True
         Height          =   975
         Left            =   4080
         Picture         =   "frmReferralLog.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Search"
         Top             =   240
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   360
         Left            =   2160
         TabIndex        =   2
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
         TabIndex        =   3
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
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5745
      Left            =   180
      TabIndex        =   9
      Top             =   1800
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   10134
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
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   180
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   7620
      Width           =   855
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1140
      TabIndex        =   18
      Top             =   7620
      Width           =   525
   End
End
Attribute VB_Name = "frmReferralLog"
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

20  On Error GoTo cmdCalc_Click_Error

30  g.Visible = False

40  ClearFGrid g

50  FromTime = Format(calFrom, "yyyymmdd") & " 00:00:00"
60  ToTime = Format(calTo, "yyyymmdd") & " 23:59:59"

70  sql = "SELECT L.Code AS RefToCode,CM.Code AS RefCode, CM.Description AS RefDescription, " & _
          "CM.CaseId,CM.Type,CM.ReferralReason,CM.Destination,CM.DateSent, " & _
          "D.PatientName, D.DateOfBirth " & _
          "FROM CaseMovements CM " & _
          "INNER JOIN Demographics D ON D.CaseId = CM.CaseId " & _
          "INNER JOIN Lists L ON CM.Destination = L.Description " & _
          "WHERE CM.DateSent BETWEEN '" & FromTime & "' AND '" & ToTime & "' "


80  If optSelect(1) Then
90      sql = sql & "AND Agreed = '1'"
100 ElseIf optSelect(2) Then
110     sql = sql & "AND Agreed <> '1'"
120 End If

130 sql = sql & " ORDER BY DateSent desc"
140 Set rsRec = New Recordset
150 RecOpenClient 0, rsRec, sql
160 If Not rsRec.EOF Then
170     pbProgress.Max = rsRec.RecordCount + 1
180     g.Visible = False
190     fraProgress.Visible = True
200     Do While Not rsRec.EOF
210         pbProgress.Value = pbProgress.Value + 1
220         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
230         lblProgress.Refresh

240         If Mid(rsRec!CaseId & "", 2, 1) = "P" Or Mid(rsRec!CaseId & "", 2, 1) = "A" Then
250             TempCaseId = Left(rsRec!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(rsRec!CaseId, 2)
260         Else
270             TempCaseId = Left(rsRec!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(rsRec!CaseId, 2)
280         End If

290         s = TempCaseId & "" & vbTab & rsRec!PatientName & "" & vbTab
300         If Not IsNull(rsRec!DateOfBirth) Then
310             s = s & Format$(rsRec!DateOfBirth, "Short Date") & vbTab
320         Else
330             s = s & "" & vbTab
340         End If

350         s = s & rsRec!RefCode & "" & vbTab

360         s = s & rsRec!RefDescription & "" & vbTab & rsRec!Type & "" & vbTab
370         s = s & rsRec!ReferralReason & "" & vbTab & rsRec!Destination & "" & vbTab
380         s = s & rsRec!RefToCode & "" & vbTab
390         If Not IsNull(rsRec!DateSent) Then
400             s = s & Format(rsRec!DateSent, "dd/mm/yyyy hh:mm") & vbTab
410         Else
420             s = s & vbTab
430         End If
440         g.AddItem s

450         rsRec.MoveNext
460     Loop
470     fraProgress.Visible = False
480     pbProgress.Value = 1
490 End If

500 g.Visible = True

510 If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1

520 Exit Sub

cmdCalc_Click_Error:

    Dim strES As String
    Dim intEL As Integer

530 intEL = Erl
540 strES = Err.Description
550 LogError "frmReferralLog", "cmdCalc_Click", intEL, strES, sql

End Sub

Private Sub cmdExit_Click()
10  Unload Me
End Sub

Private Sub cmdExport_Click()
10  ExportFlexGrid g, Me
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
710 LogError "frmReferralLog", "cmdPrint_Click", intEL, strES


End Sub
Private Sub InitializeGrid()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 10: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "Case Id": .ColWidth(0) = 1150: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Patient Name": .ColWidth(1) = 1500: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "DOB": .ColWidth(2) = 1000: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Code": .ColWidth(3) = 950: .ColAlignment(3) = flexAlignLeftCenter
160     .TextMatrix(0, 4) = "Description": .ColWidth(4) = 2000: .ColAlignment(4) = flexAlignLeftCenter
170     .TextMatrix(0, 5) = "Type": .ColWidth(5) = 2000: .ColAlignment(5) = flexAlignLeftCenter
180     .TextMatrix(0, 6) = "Reason For Referral": .ColWidth(6) = 2000: .ColAlignment(6) = flexAlignLeftCenter
190     .TextMatrix(0, 7) = "Referred To": .ColWidth(7) = 2000: .ColAlignment(7) = flexAlignLeftCenter
200     .TextMatrix(0, 8) = "Ref. Code": .ColWidth(8) = 1000: .ColAlignment(8) = flexAlignLeftCenter
210     .TextMatrix(0, 9) = "Date Sent": .ColWidth(9) = 1500: .ColAlignment(9) = flexAlignLeftCenter

220 End With
End Sub

Private Sub Form_Load()

ChangeFont Me, "Arial"
'frmReferralLog_ChangeLanguage
calTo = Format(Now, "dd/MM/yyyy")
calFrom = Format(Now - 7, "dd/MM/yyyy")
lblLoggedIn = UserName

InitializeGrid
If blnIsTestMode Then EnableTestMode Me

End Sub
Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub optSelect_Click(Index As Integer)
10  cmdCalc_Click
End Sub

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportViewer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cellular Pathology -  Report Viewer"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   12000
      Picture         =   "frmReportViewer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Re-Print this page"
      Height          =   855
      Left            =   13200
      Picture         =   "frmReportViewer.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame fraPage 
      Caption         =   "Viewing Page"
      Height          =   1335
      Left            =   9000
      TabIndex        =   6
      Top             =   360
      Width           =   1635
      Begin ComCtl2.UpDown udPage 
         Height          =   285
         Left            =   195
         TabIndex        =   7
         Top             =   600
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblCurrentPage"
         BuddyDispid     =   196613
         OrigLeft        =   25
         OrigTop         =   8
         OrigRight       =   25
         OrigBottom      =   9
         Max             =   99
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCurrentPage 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   210
         TabIndex        =   10
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "of"
         Height          =   195
         Left            =   720
         TabIndex        =   9
         Top             =   420
         Width           =   135
      End
      Begin VB.Label lblTotalPages 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   930
         TabIndex        =   8
         Top             =   390
         Width           =   465
      End
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8385
      Left            =   90
      TabIndex        =   5
      Top             =   1740
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   14790
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReportViewer.frx":075D
   End
   Begin MSFlexGridLib.MSFlexGrid grdRep 
      Height          =   1215
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ForeColorSel    =   65535
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      SelectionMode   =   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1080
      TabIndex        =   14
      Top             =   10200
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   10200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Report to View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   210
      Width           =   7320
   End
   Begin VB.Label lblNoRep 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   7545
      TabIndex        =   2
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number of Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   7545
      TabIndex        =   1
      Top             =   480
      Width           =   1275
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mSampleID As String
Private mDept As String
Private PreviousReport As String
Private mYear As String

Public Property Let Year(ByVal NewValue As String)

10  mYear = NewValue

End Property

Private Sub cmdExit_Click()

10  On Error GoTo cmdExit_Click_Error

20  If PreviousReport <> "" Then
30      If Dir(PreviousReport & ".rpt") <> "" Then
40          Kill (PreviousReport & ".rpt")
50      End If
60  End If

70  Unload Me

80  Exit Sub

cmdExit_Click_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmReportViewer", "cmdExit_Click", intEL, strES

End Sub

Private Sub cmdPrint_Click()

10  rtb.SelStart = 0
20  rtb.SelLength = 100000
30  rtb.SelPrint Printer.hDC

End Sub


Private Sub Form_Load()

   Dim sql As String

10  On Error GoTo Form_Load_Error

20  InitializeGrid
30  PreviousReport = ""
40  FillGrid

50  lblLoggedIn = UserName
60  If blnIsTestMode Then EnableTestMode Me

70  Exit Sub

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmReportViewer", "Form_Load", intEL, strES, sql

End Sub
Private Sub FillGrid()

    Dim tb As Recordset
    Dim tb1 As Recordset
    Dim sql As String
    Dim TotalReports As Integer
    Dim Y As Integer
    Dim Target As String
    Dim TempCaseId As String


10  If Mid(mSampleID & "", 2, 1) = "P" Or Mid(mSampleID & "", 2, 1) = "A" Then
20      TempCaseId = Left(mSampleID, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(mSampleID, 2)
30  Else
40      TempCaseId = Left(mSampleID, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(mSampleID, 2)
50  End If

60  With grdRep
70      sql = "SELECT PrintTime FROM Reports WHERE " & _
              "SampleID = '" & mSampleID & "' " & _
              "AND Year = '" & mYear & "' " & _
              "ORDER BY PrintTime DESC"

80      Set tb = New Recordset
90      RecOpenServer 0, tb, sql
100     If Not tb.EOF Then
110         Do While Not tb.EOF
120             .AddItem TempCaseId & vbTab & Format(tb!PrintTime, "dd/MM/yy HH:mm:ss")
130             tb.MoveNext
140         Loop
150         If .Rows > 2 Then
160             .RemoveItem 1
170         End If
180         If .Rows > 2 Then
190             Target = .TextMatrix(1, 1)
200             For Y = 2 To .Rows - 1
210                 If DateDiff("s", .TextMatrix(Y, 1), Target) < 10 And .TextMatrix(Y, 1) <> Target Then
220                     sql = "UPDATE Reports SET PrintTime = '" & Format(Target, "dd/MMM/yyyy HH:mm:ss") & "' " & _
                              "WHERE SampleID = '" & mSampleID & "' " & _
                              "AND PrintTime = '" & Format(.TextMatrix(Y, 1), "dd/MMM/yyyy HH:mm:ss") & "' "
230                     Set tb1 = New Recordset
240                     RecOpenServer 0, tb1, sql
250                 Else
260                     Target = .TextMatrix(Y, 1)
270                 End If
280             Next
290         End If
300     End If

310     .Rows = 2
320     .AddItem ""
330     .RemoveItem 1

340     sql = "SELECT DISTINCT(PrintTime), Initiator, Printer, RepNo FROM Reports WHERE " & _
              "SampleID = '" & mSampleID & "' " & _
              "ORDER BY PrintTime DESC"
350     Set tb = New Recordset
360     RecOpenServer 0, tb, sql
370     Do While Not tb.EOF
380         .AddItem TempCaseId & vbTab & Format(tb!PrintTime, "dd/MM/yy HH:mm:ss") & vbTab & _
                     tb!Initiator & vbTab & _
                     tb!Printer & vbTab & tb!RepNo & ""
390         tb.MoveNext
400     Loop
410     If .Rows > 2 Then
420         .RemoveItem 1
430     End If
440 End With

450 TotalReports = grdRep.Rows - 1
460 lblNoRep = TotalReports
470 If TotalReports > 0 Then
480     grdRep.row = 1
490     HighlightRow
500     fraPage.Visible = True
510     lblTotalPages = PagesPerReport(grdRep.TextMatrix(1, 4))
520     lblCurrentPage = "1"
530     udPage.Max = lblTotalPages
540     PreviousReport = grdRep.TextMatrix(grdRep.row, 4)
550     DisplayReport grdRep.TextMatrix(grdRep.row, 4), Val(lblCurrentPage)
560 Else
570     fraPage.Visible = False
580 End If

End Sub



Public Property Let SampleID(ByVal SID As String)

10  On Error GoTo SampleID_Error

20  mSampleID = SID

30  Exit Property

SampleID_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmReportViewer", "SampleID", intEL, strES


End Property

Public Property Let Dept(ByVal Dep As String)

10  On Error GoTo Dept_Error

20  mDept = Dep

30  Exit Property

Dept_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmReportViewer", "Dept", intEL, strES


End Property

Private Sub InitializeGrid()

10  With grdRep
20      .Rows = 2: .FixedRows = 1
30      .Cols = 5: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "CaseId": .ColWidth(0) = 1170: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "PrintTime": .ColWidth(1) = 1800: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "Printed By": .ColWidth(2) = 1500: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Printer": .ColWidth(3) = 2500: .ColAlignment(3) = flexAlignLeftCenter
160     .TextMatrix(0, 4) = "ReportNo": .ColWidth(4) = 0: .ColAlignment(4) = flexAlignLeftCenter
170 End With
End Sub


Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub


Private Sub grdRep_Click()
10  HighlightRow
20  lblTotalPages = PagesPerReport(grdRep.TextMatrix(grdRep.row, 4))
30  lblCurrentPage = "1"
40  udPage.Max = lblTotalPages

50  If PreviousReport <> "" Then
60      If Dir(PreviousReport & ".rpt") <> "" Then
70          Kill (PreviousReport & ".rpt")
80      End If
90  End If

100 PreviousReport = grdRep.TextMatrix(grdRep.row, 4)
110 DisplayReport grdRep.TextMatrix(grdRep.row, 4), Val(lblCurrentPage)
End Sub

Private Sub HighlightRow()

    Dim X As Integer
    Dim Y As Integer
    Dim ySave As Integer

10  With grdRep
20      ySave = .row

30      .col = 0
40      For Y = 1 To .Rows - 1
50          .row = Y
60          If .CellBackColor = vbYellow Then
70              For X = 0 To .Cols - 1
80                  .col = X
90                  .CellBackColor = 0
100             Next
110             Exit For
120         End If
130     Next

140     .row = ySave
150     For X = 0 To .Cols - 1
160         .col = X
170         .CellBackColor = vbYellow
180     Next

190 End With

End Sub

Private Sub UdPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10  DisplayReport grdRep.TextMatrix(grdRep.row, 4), Val(lblCurrentPage)

End Sub


Private Function PagesPerReport(ByVal RepNo As String) As Integer

    Dim sql As String
    Dim tb As Recordset



10  On Error GoTo PagesPerReport_Error

20  sql = "SELECT COUNT(*) Tot FROM Reports WHERE " & _
          "SampleID = '" & mSampleID & "' " & _
          "AND RepNo = '" & RepNo & "' "
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql
50  PagesPerReport = tb!tot



60  Exit Function

PagesPerReport_Error:

    Dim strES As String
    Dim intEL As Integer

70  intEL = Erl
80  strES = Err.Description
90  LogError "frmReportViewer", "PagesPerReport", intEL, strES, sql


End Function


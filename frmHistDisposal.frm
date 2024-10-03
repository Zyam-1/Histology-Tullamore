VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmHistDisposal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   11325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15360
   Icon            =   "frmHistDisposal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelectInclude 
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6375
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2940
      Width           =   315
   End
   Begin VB.CommandButton cmdSelectInclude 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2940
      Width           =   315
   End
   Begin VB.TextBox txtBetween 
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   6840
      TabIndex        =   29
      Top             =   360
      Width           =   2355
   End
   Begin VB.TextBox txtBetween 
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   3960
      TabIndex        =   27
      Top             =   360
      Width           =   2355
   End
   Begin VB.OptionButton optList 
      Caption         =   "List of all specimens between"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   420
      Value           =   -1  'True
      Width           =   3615
   End
   Begin VB.OptionButton optList 
      Caption         =   "List of all specimens between"
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Tag             =   "NOT Scheduled For Disposal"
      Top             =   1380
      Width           =   2535
   End
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   5760
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   13
      Top             =   6780
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.OptionButton optList 
      Caption         =   "List of all specimens scheduled for disposal on"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Tag             =   "Scheduled For Disposal"
      Top             =   660
      Width           =   3615
   End
   Begin VB.OptionButton optList 
      Caption         =   "List of all Kept specimens"
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Tag             =   "Kept Specimen"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.OptionButton optList 
      Caption         =   "List of all disposed specimens between"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Tag             =   "Disposed Specimen"
      Top             =   2460
      Width           =   3255
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   855
      Left            =   13320
      Picture         =   "frmHistDisposal.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdReloadList 
      Caption         =   "&Reload"
      Height          =   855
      Left            =   11400
      Picture         =   "frmHistDisposal.frx":1291
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   14280
      Picture         =   "frmHistDisposal.frx":1604
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   855
      Left            =   12360
      Picture         =   "frmHistDisposal.frx":1946
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Scheduled For Disposal"
      Top             =   1920
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7725
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   13626
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
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker calFrom 
      Height          =   345
      Index           =   0
      Left            =   3960
      TabIndex        =   9
      Top             =   2400
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " "
      Format          =   146735107
      CurrentDate     =   37753
   End
   Begin MSComCtl2.DTPicker calTo 
      Height          =   345
      Index           =   0
      Left            =   6870
      TabIndex        =   10
      Top             =   2400
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " "
      Format          =   146735107
      CurrentDate     =   37753
   End
   Begin MSComCtl2.DTPicker calSchedOn 
      Height          =   345
      Index           =   1
      Left            =   3960
      TabIndex        =   12
      Top             =   840
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " "
      Format          =   146735107
      CurrentDate     =   37753
   End
   Begin MSComCtl2.DTPicker calNotSchedOn 
      Height          =   345
      Left            =   11640
      TabIndex        =   17
      Top             =   1320
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " "
      Format          =   146735107
      CurrentDate     =   37753
   End
   Begin MSComCtl2.DTPicker calFrom 
      Height          =   345
      Index           =   1
      Left            =   3960
      TabIndex        =   18
      Top             =   1320
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " "
      Format          =   146735107
      CurrentDate     =   37753
   End
   Begin MSComCtl2.DTPicker calTo 
      Height          =   345
      Index           =   1
      Left            =   6870
      TabIndex        =   19
      Top             =   1320
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " "
      Format          =   146735107
      CurrentDate     =   37753
   End
   Begin MSComCtl2.DTPicker calSchedOn 
      Height          =   345
      Index           =   0
      Left            =   11640
      TabIndex        =   25
      Top             =   360
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " "
      Format          =   146735107
      CurrentDate     =   37753
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "and"
      Height          =   195
      Left            =   6480
      TabIndex        =   28
      Top             =   420
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "scheduled for disposal on"
      Height          =   195
      Left            =   9480
      TabIndex        =   26
      Top             =   420
      Width           =   1800
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   11040
      Width           =   855
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   22
      Top             =   11040
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "NOT scheduled for disposal on"
      Height          =   195
      Left            =   9360
      TabIndex        =   21
      Top             =   1380
      Width           =   2190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "and"
      Height          =   195
      Left            =   6480
      TabIndex        =   20
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   14760
      Picture         =   "frmHistDisposal.frx":1D61
      Top             =   840
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   14760
      Picture         =   "frmHistDisposal.frx":2037
      Top             =   360
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "and"
      Height          =   195
      Left            =   6480
      TabIndex        =   11
      Top             =   2460
      Width           =   270
   End
End
Attribute VB_Name = "frmHistDisposal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DataSaved As Boolean
Private pDisposalType As String

Public Property Let DisposalType(ByVal Value As String)

10  pDisposalType = Value

End Property

Private Sub cmdExit_Click()

10  If DataChanged = True Then
20      If frmMsgBox.Msg("Do you want to save changes to grid?", mbYesNo, "Save Changes", mbQuestion) = 1 Then
30          SaveGrid
40      End If
50  End If
60  Unload Me
End Sub

Private Sub cmdPrint_Click()
    Const GAP = 60

    Dim xmax As Single
    Dim ymax As Single
    Dim xmin As Single
    Dim ymin As Single
    Dim X As Single
    Dim c As Integer

    Dim i As Integer


10  On Error GoTo cmdPrint_Click_Error

20  xmin = 1000    '1440
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
210             For i = 0 To 4
220                 If optList(i) Then
230                     PrintHeadingWorkLog "Page " & lPrintPage & " of " & lNoOfPages, "Worklog : " & optList(i).Tag, 9
240                     Exit For
250                 End If
260             Next

                ' Print each row.
270             Printer.CurrentY = ymin

280             Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)

290             Printer.CurrentY = Printer.CurrentY + GAP

300             X = xmin + GAP
310             For c = 0 To .Cols - 1
320                 Printer.CurrentX = X
330                 If .ColWidth(c) <> 0 Then
340                     PrintText BoundedText(Printer, .TextMatrix(0, c), .ColWidth(c)), "MS Sans Serif", , True
350                     X = X + .ColWidth(c) + 2 * GAP
360                 End If

370             Next c
380             Printer.CurrentY = Printer.CurrentY + GAP

                ' Move to the next line.
390             PrintText vbCrLf

400             For lThisRow = lRowsPrinted To lRowsPerPage * lPrintPage

410                 If (lThisRow - 1) < lNumRows Then
420                     If lThisRow > 0 Then Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)
430                     Printer.CurrentY = Printer.CurrentY + GAP

                        ' Print the entries on this row.
440                     X = xmin + GAP
450                     For c = 0 To .Cols - 1
460                         Printer.CurrentX = X
470                         If .ColWidth(c) <> 0 Then
480                             PrintText BoundedText(Printer, .TextMatrix(lThisRow, c), .ColWidth(c)), "MS Sans Serif", 8
490                             If c = 3 Or c = 4 Then
500                                 g.col = c
510                                 g.row = lThisRow
520                                 If g.CellPicture = imgSquareTick.Picture Then
530                                     Call Printer.PaintPicture(imgSquareTick.Picture, (Printer.CurrentX + (.ColWidth(c) / 2)) - 105, Printer.CurrentY)
540                                 End If

550                             End If
560                             Printer.FontName = "MS Sans Serif"
570                             Printer.FontSize = 8
580                             X = X + .ColWidth(c) + 2 * GAP
590                         End If

600                     Next c
610                     Printer.CurrentY = Printer.CurrentY + GAP

                        ' Move to the next line.
620                     PrintText vbCrLf

630                     lRowsPrinted = lRowsPrinted + 1
640                 Else
650                     Exit Do
660                 End If
670             Next
680         End With
690     Loop While (lRowsPrinted - 1) < lRowsPerPage * lPrintPage

700     ymax = Printer.CurrentY

        ' Draw a box around everything.
710     Printer.Line (xmin, ymin)-(xmax, ymax), , B

        ' Draw lines between the columns.
720     X = xmin
730     For c = 0 To g.Cols - 2
740         If g.ColWidth(c) <> 0 Then
750             X = X + g.ColWidth(c) + 2 * GAP
760             Printer.Line (X, ymin)-(X, ymax)
770         End If
780     Next c

790     Printer.EndDoc
800     lPrintPage = lPrintPage + 1

810 Loop While (lRowsPrinted - 1) < lNumRows




820 Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

830 intEL = Erl
840 strES = Err.Description
850 LogError "frmHistDisposal", "cmdPrint_Click", intEL, strES


End Sub

Private Sub cmdReloadList_Click()
10  If DataChanged = True Then
20      If frmMsgBox.Msg("Do you want to save changes to grid?", mbYesNo, "Save Changes", mbQuestion) = 1 Then
30          SaveGrid
40      End If
50  End If
60  LoadList
End Sub

Private Sub LoadList()
10  ClearFGrid g
20  If optList(0) Then
30      LoadSchedDisposalList
40  ElseIf optList(1) Then
50      LoadSchedDisposalList
60  ElseIf optList(2) Then
70      LoadNotSchedDisposalList
80  ElseIf optList(3) Then
90      LoadKeptList
100 ElseIf optList(4) Then
110     LoadDisposedList

120 End If
130 cmdPrint.Visible = True
End Sub

Private Sub cmdSave_Click()
10  If Not optList(2) Then
20      If frmMsgBox.Msg("Are you sure you want to save grid?", mbYesNo, "Save Grid", mbQuestion) = 1 Then
30          SaveGrid
40      End If
50  End If

End Sub



Private Sub cmdSelectInclude_Click(Index As Integer)

10  On Error GoTo cmdSelectInclude_Click_Error

    Dim i As Integer

20  If g.TextMatrix(1, 0) = "" Then Exit Sub

30  With g
40      For i = 1 To g.Rows - 1
50          g.row = i
60          g.col = 4
70          Select Case Index
            Case 0
80              Set g.CellPicture = imgSquareTick
90              .col = 3
100             Set g.CellPicture = imgSquareCross

110         Case 1
120             Set g.CellPicture = imgSquareCross
130             .col = 3
140             Set g.CellPicture = imgSquareTick
                '            Case 2
                '                If g.CellPicture = imgSquareTick Then
                '                    Set g.CellPicture = imgSquareCross
                '                    .Col = 3
                '                    Set g.CellPicture = imgSquareTick
                '                ElseIf g.CellPicture = imgSquareCross Then
                '                    Set g.CellPicture = imgSquareTick
                '                    .Col = 3
                '                    Set g.CellPicture = imgSquareCross
                '                End If
150         End Select

160     Next i
170 End With

180 Exit Sub

cmdSelectInclude_Click_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmHistDisposal", "cmdSelectInclude_Click", intEL, strES

End Sub

Private Sub Form_Activate()
10  If pDisposalType = "H" Then
20      Me.Caption = "Histology Disposal"
30  ElseIf pDisposalType = "C" Then
40      Me.Caption = "Cytology Disposal"
50  ElseIf pDisposalType = "A" Then
60      Me.Caption = "Autopsy Disposal"
70  End If

End Sub

Private Sub Form_Load()

ChangeFont Me, "Arial"
'frmHistDisposal_ChangeLanguage
calSchedOn(0).Enabled = True
calSchedOn(0).CustomFormat = "dd/MM/yyyy"
calSchedOn(0) = Format(Now, "dd/MM/yyyy")
txtBetween(0).Enabled = True
txtBetween(1).Enabled = True
lblLoggedIn = UserName

InitializeGrid

LoadList
If blnIsTestMode Then EnableTestMode Me
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

120     .TextMatrix(0, 0) = "Case Id": .ColWidth(0) = 1300: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Specimen Type": .ColWidth(1) = 2800: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "AE": .ColWidth(2) = 500: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Keep": .ColWidth(3) = 1000: .ColAlignment(3) = flexAlignCenterCenter
160     .TextMatrix(0, 4) = "Dispose": .ColWidth(4) = 1000: .ColAlignment(4) = flexAlignCenterCenter
170     .TextMatrix(0, 5) = "Comment": .ColWidth(5) = 2500: .ColAlignment(5) = flexAlignLeftCenter
180     .TextMatrix(0, 6) = "Specimen Rec": .ColWidth(6) = 1700: .ColAlignment(6) = flexAlignLeftCenter
190     .TextMatrix(0, 7) = "Disposed By": .ColWidth(7) = 1700: .ColAlignment(7) = flexAlignLeftCenter
200     .TextMatrix(0, 8) = "Date Disposed": .ColWidth(8) = 1500: .ColAlignment(8) = flexAlignLeftCenter
210     .TextMatrix(0, 9) = "Changed": .ColWidth(9) = 0: .ColAlignment(9) = flexAlignLeftCenter
220 End With
End Sub
Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then
20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub
Private Sub LoadDisposedList()
    Dim sql As String
    Dim tb As Recordset
    Dim TempCaseId As String
    Dim s As String

10  On Error GoTo LoadDisposedList_Error

20  sql = "SELECT * FROM CaseTree CT " & _
          "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
          "WHERE CT.DisposalDate BETWEEN '" & Format(calFrom(0), "yyyymmdd") & " 00:00:00' " & _
          "AND '" & Format(calTo(0), "yyyymmdd") & " 23:59:59' " & _
          "AND CT.Disposal = 'D' " & _
          "AND TissueTypeListId IS NOT NULL "

30  Select Case pDisposalType
    Case "H", "C"
40      sql = sql & "AND SUBSTRING(CT.CaseId,1,1) = '" & pDisposalType & "'"
50  Case "A"
60      sql = sql & "AND SUBSTRING(CT.CaseId,2,1) = '" & pDisposalType & "'"
70  End Select

80  Set tb = New Recordset
90  RecOpenClient 0, tb, sql

100 If Not tb.EOF Then
110     pbProgress.Max = tb.RecordCount + 1
120     g.Visible = False
130     fraProgress.Visible = True
140     Do While Not tb.EOF
150         pbProgress.Value = pbProgress.Value + 1
160         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
170         lblProgress.Refresh

180         If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
190             TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
200         Else
210             TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
220         End If

230         s = TempCaseId
240         s = s & vbTab & tb!LocationName & vbTab
250         s = s & vbTab & vbTab & vbTab & tb!DisposalComment & vbTab & tb!SampleReceived
260         s = s & vbTab & tb!DisposedBy & vbTab
270         If IsDate(tb!DisposalDate) Then s = s & Format(tb!DisposalDate, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
280         g.AddItem s

290         g.row = g.Rows - 1
300         g.col = 3
310         g.CellPictureAlignment = flexAlignCenterCenter
320         Set g.CellPicture = imgSquareCross.Picture

330         g.col = 4
340         g.CellPictureAlignment = flexAlignCenterCenter
350         Set g.CellPicture = imgSquareTick.Picture

360         tb.MoveNext
370     Loop
380     fraProgress.Visible = False
390     pbProgress.Value = 1
400 End If

410 g.Visible = True
420 If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1


430 Exit Sub

LoadDisposedList_Error:

    Dim strES As String
    Dim intEL As Integer

440 intEL = Erl
450 strES = Err.Description
460 LogError "frmHistDisposal", "LoadDisposedList", intEL, strES, sql


End Sub

Private Sub LoadKeptList()
    Dim sql As String
    Dim tb As Recordset
    Dim TempCaseId As String
    Dim s As String

10  On Error GoTo LoadKeptList_Error

20  sql = "SELECT * FROM CaseTree CT " & _
          "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
          "WHERE CT.Disposal = 'K' " & _
          "AND TissueTypeListId IS NOT NULL "

30  Select Case pDisposalType
    Case "H", "C"
40      sql = sql & "AND SUBSTRING(CT.CaseId,1,1) = '" & pDisposalType & "'"
50  Case "A"
60      sql = sql & "AND SUBSTRING(CT.CaseId,2,1) = '" & pDisposalType & "'"
70  End Select

80  Set tb = New Recordset
90  RecOpenClient 0, tb, sql

100 If Not tb.EOF Then
110     pbProgress.Max = tb.RecordCount + 1
120     g.Visible = False
130     fraProgress.Visible = True
140     Do While Not tb.EOF
150         pbProgress.Value = pbProgress.Value + 1
160         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
170         lblProgress.Refresh

180         If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
190             TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
200         Else
210             TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
220         End If

230         s = TempCaseId
240         s = s & vbTab & tb!LocationName & vbTab
250         s = s & vbTab & vbTab & vbTab & tb!DisposalComment & vbTab & tb!SampleReceived
260         s = s & vbTab & tb!DisposedBy & vbTab
270         If IsDate(tb!DisposalDate) Then s = s & Format(tb!DisposalDate, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
280         g.AddItem s

290         g.row = g.Rows - 1
300         g.col = 3
310         g.CellPictureAlignment = flexAlignCenterCenter
320         Set g.CellPicture = imgSquareTick.Picture

330         g.col = 4
340         g.CellPictureAlignment = flexAlignCenterCenter
350         Set g.CellPicture = imgSquareCross.Picture

360         tb.MoveNext
370     Loop
380     fraProgress.Visible = False
390     pbProgress.Value = 1
400 End If

410 g.Visible = True
420 If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1

430 Exit Sub

LoadKeptList_Error:

    Dim strES As String
    Dim intEL As Integer

440 intEL = Erl
450 strES = Err.Description
460 LogError "frmHistDisposal", "LoadKeptList", intEL, strES, sql


End Sub

Private Sub LoadSchedDisposalList()
    Dim sql As String
    Dim tb As Recordset
    Dim TempCaseId As String
    Dim s As String

10  On Error GoTo LoadDisposalList_Error

20  sql = "SELECT * FROM CaseTree CT " & _
          "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
          "WHERE C.State = 'Authorised' " & _
          "AND (CT.Disposal IS NULL OR CT.Disposal = '')" & _
          "AND TissueTypeListId IS NOT NULL "

30  If optList(0) Then
40      sql = sql & "AND REPLACE(CONVERT(VARCHAR(10),DATEADD(dd, 14, ValReportDate),120),'-','') <= '" & Format(calSchedOn(0), "yyyymmdd") & "' "
50  Else
60      sql = sql & "AND REPLACE(CONVERT(VARCHAR(10),DATEADD(dd, 14, ValReportDate),120),'-','') <= '" & Format(calSchedOn(1), "yyyymmdd") & "' "
70  End If

80  Select Case pDisposalType
    Case "H", "C", "J"
90      sql = sql & "AND SUBSTRING(CT.CaseId,1,1) = '" & pDisposalType & "' "
100     If optList(0) Then
110         sql = sql & "AND CONVERT(INTEGER,SUBSTRING(CT.CaseId,2,5),120) BETWEEN " & Val(Mid(txtBetween(0), 2, 5)) & " AND " & Val(Mid(txtBetween(1), 2, 5)) & _
                  "AND REPLACE(CONVERT(VARCHAR(10),DATEADD(wk, 5, SampleReceived),120),'-','') <= '" & Format(calSchedOn(0), "yyyymmdd") & "' "
120     Else
130         sql = sql & "AND REPLACE(CONVERT(VARCHAR(10),DATEADD(wk, 5, SampleReceived),120),'-','') <= '" & Format(calSchedOn(1), "yyyymmdd") & "' "
140     End If

150 Case "A"
160     sql = sql & "AND SUBSTRING(CT.CaseId,2,1) = '" & pDisposalType & "'"

170     If optList(0) Then
180         sql = sql & "AND CONVERT(INTEGER,SUBSTRING(CT.CaseId,3,5),120) BETWEEN " & Val(Mid(txtBetween(0), 3, 5)) & " AND " & Val(Mid(txtBetween(1), 3, 5))
190     End If
200 End Select


210 Set tb = New Recordset
220 RecOpenClient 0, tb, sql

230 If Not tb.EOF Then
240     pbProgress.Max = tb.RecordCount + 1
250     g.Visible = False
260     fraProgress.Visible = True
270     Do While Not tb.EOF
280         pbProgress.Value = pbProgress.Value + 1
290         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
300         lblProgress.Refresh

310         If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
320             TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
330         Else
340             TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
350         End If

360         s = TempCaseId
370         s = s & vbTab & tb!LocationName & vbTab
380         s = s & vbTab & vbTab & vbTab & tb!DisposalComment & vbTab & tb!SampleReceived
390         s = s & vbTab & tb!DisposedBy & vbTab
400         If IsDate(tb!DisposalDate) Then s = s & Format(tb!DisposalDate, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
410         s = s & "1"
420         g.AddItem s

430         g.row = g.Rows - 1
440         g.col = 3
450         g.CellPictureAlignment = flexAlignCenterCenter
460         Set g.CellPicture = imgSquareCross.Picture

470         g.col = 4
480         g.CellPictureAlignment = flexAlignCenterCenter
490         Set g.CellPicture = imgSquareTick.Picture

500         tb.MoveNext
510     Loop
520     fraProgress.Visible = False
530     pbProgress.Value = 1
540 End If

550 g.Visible = True
560 If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1

570 Exit Sub

LoadDisposalList_Error:

    Dim strES As String
    Dim intEL As Integer

580 intEL = Erl
590 strES = Err.Description
600 LogError "frmHistDisposal", "LoadDisposalList", intEL, strES, sql


End Sub

Private Sub LoadNotSchedDisposalList()
    Dim sql As String
    Dim tb As Recordset
    Dim TempCaseId As String
    Dim s As String

10  On Error GoTo LoadNotSchedDisposalList_Error

20  sql = "SELECT * FROM CaseTree CT " & _
          "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
          "WHERE ((C.State = 'Authorised' " & _
          "AND REPLACE(CONVERT(VARCHAR(10),DATEADD(dd, 14, C.ValReportDate),120),'-','') > '" & Format(calNotSchedOn, "yyyymmdd") & "') " & _
          "OR C.State <> 'Authorised') " & _
          "AND (CT.Disposal IS NULL OR CT.Disposal = '') " & _
          "AND TissueTypeListId IS NOT NULL "

30  Select Case pDisposalType
    Case "H", "C"
40      sql = sql & "AND SampleReceived BETWEEN '" & Format(calFrom(1), "yyyymmdd") & " 00:00:00' " & _
              "AND '" & Format(calTo(1), "yyyymmdd") & " 23:59:59' " & _
              "AND SUBSTRING(CT.CaseId,1,1) = '" & pDisposalType & "'"
50  Case "A"
60      sql = sql & "AND SampleReceived BETWEEN '" & Format(calFrom(1), "yyyymmdd") & " 00:00:00' " & _
              "AND '" & Format(calTo(1), "yyyymmdd") & " 23:59:59' " & _
              "AND SUBSTRING(CT.CaseId,2,1) = '" & pDisposalType & "'"
70  End Select


80  Set tb = New Recordset
90  RecOpenClient 0, tb, sql

100 If Not tb.EOF Then
110     pbProgress.Max = tb.RecordCount + 1
120     g.Visible = False
130     fraProgress.Visible = True
140     Do While Not tb.EOF
150         pbProgress.Value = pbProgress.Value + 1
160         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
170         lblProgress.Refresh

180         If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
190             TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
200         Else
210             TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
220         End If

230         s = TempCaseId
240         s = s & vbTab & tb!LocationName & vbTab
250         s = s & vbTab & vbTab & vbTab & tb!DisposalComment & vbTab & tb!SampleReceived
260         s = s & vbTab & tb!DisposedBy & vbTab
270         If IsDate(tb!DisposalDate) Then s = s & Format(tb!DisposalDate, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
280         g.AddItem s

290         g.row = g.Rows - 1
300         g.col = 3
310         g.CellPictureAlignment = flexAlignCenterCenter
320         Set g.CellPicture = imgSquareCross.Picture

330         g.col = 4
340         g.CellPictureAlignment = flexAlignCenterCenter
350         Set g.CellPicture = imgSquareCross.Picture

360         tb.MoveNext
370     Loop
380     fraProgress.Visible = False
390     pbProgress.Value = 1
400 End If

410 g.Visible = True
420 If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1


430 Exit Sub

LoadNotSchedDisposalList_Error:

    Dim strES As String
    Dim intEL As Integer

440 intEL = Erl
450 strES = Err.Description
460 LogError "frmHistDisposal", "LoadNotSchedDisposalList", intEL, strES, sql


End Sub

Private Sub g_Click()

10  On Error GoTo g_Click_Error

20  Rada = g.MouseRow
30  If g.col = 3 Then
40      If g.CellPicture = imgSquareTick.Picture Then
50          Set g.CellPicture = imgSquareCross.Picture
60          g.col = 4
70          Set g.CellPicture = imgSquareTick.Picture
80      Else
90          Set g.CellPicture = imgSquareTick.Picture
100         g.col = 4
110         Set g.CellPicture = imgSquareCross.Picture
120     End If
130     g.col = 3

140     DataChanged = True
150 End If

160 If g.col = 4 Then
170     If g.CellPicture = imgSquareTick.Picture Then
180         Set g.CellPicture = imgSquareCross.Picture
190         g.col = 3
200         Set g.CellPicture = imgSquareTick.Picture
210     Else
220         Set g.CellPicture = imgSquareTick.Picture
230         g.col = 3
240         Set g.CellPicture = imgSquareCross.Picture
250     End If
260     g.col = 4

270     DataChanged = True

280 End If

290 If g.col = 5 Then

300     With frmComment
310         .CommentType = "DISPOSAL"
320         .GridRow = Rada
330         .GeneralComment = g.Text
340         .cmdExit.Visible = True
350         .Move Me.Left + g.Left, frmWorkSheet.Top + g.Top + (g.row * g.RowHeight(0))
360         .Show 1
370     End With
380 End If

390 If DataChanged = True Then
400     If optList(2) Or optList(3) Then
410         g.TextMatrix(Rada, 9) = "1"
420     End If
430     cmdPrint.Visible = False
440 End If

450 Exit Sub

g_Click_Error:

    Dim strES As String
    Dim intEL As Integer

460 intEL = Erl
470 strES = Err.Description
480 LogError "frmHistDisposal", "g_Click", intEL, strES


End Sub

Private Sub optList_Click(Index As Integer)

10  Select Case Index
    Case 0
20      If DataChanged = True Then
30          If frmMsgBox.Msg("Do you want to save changes to grid?", mbYesNo, "Save Changes", mbQuestion) = 1 Then
40              SaveGrid
50          End If
60      End If
70      g.Enabled = True
80      calFrom(1).Enabled = False
90      calFrom(1).CustomFormat = " "
100     calTo(1).Enabled = False
110     calTo(1).CustomFormat = " "
120     calFrom(0).Enabled = False
130     calFrom(0).CustomFormat = " "
140     calTo(0).Enabled = False
150     calTo(0).CustomFormat = " "
160     calSchedOn(0).Enabled = True
170     calSchedOn(0).CustomFormat = "dd/MM/yyyy"
180     calSchedOn(0) = Format(Now, "dd/MM/yyyy")
190     calSchedOn(1).Enabled = False
200     calSchedOn(1).CustomFormat = " "
210     txtBetween(0) = ""
220     txtBetween(0).Enabled = True
230     txtBetween(1) = ""
240     txtBetween(1).Enabled = True
250     calNotSchedOn.Enabled = False
260     calNotSchedOn.CustomFormat = " "
270     g.TextMatrix(0, 7) = "Disposed By"
280     g.TextMatrix(0, 4) = "Dispose"
290 Case 1
300     If DataChanged = True Then
310         If frmMsgBox.Msg("Do you want to save changes to grid?", mbYesNo, "Save Changes", mbQuestion) = 1 Then
320             SaveGrid
330         End If
340     End If
350     g.Enabled = True
360     calFrom(1).Enabled = False
370     calFrom(1).CustomFormat = " "
380     calTo(1).Enabled = False
390     calTo(1).CustomFormat = " "
400     calFrom(0).Enabled = False
410     calFrom(0).CustomFormat = " "
420     calTo(0).Enabled = False
430     calTo(0).CustomFormat = " "
440     calSchedOn(0).Enabled = False
450     calSchedOn(0).CustomFormat = " "
460     calSchedOn(1).Enabled = True
470     calSchedOn(1).CustomFormat = "dd/MM/yyyy"
480     calSchedOn(1) = Format(Now, "dd/MM/yyyy")
490     txtBetween(0) = ""
500     txtBetween(0).Enabled = False
510     txtBetween(1) = ""
520     txtBetween(1).Enabled = False
530     calNotSchedOn.Enabled = False
540     calNotSchedOn.CustomFormat = " "
550     g.TextMatrix(0, 7) = "Disposed By"
560     g.TextMatrix(0, 4) = "Dispose"
570 Case 2
580     If DataChanged = True Then
590         If frmMsgBox.Msg("Do you want to save changes to grid?", mbYesNo, "Save Changes", mbQuestion) = 1 Then
600             SaveGrid
610         End If
620     End If
630     g.Enabled = False
640     calFrom(0).Enabled = False
650     calFrom(0).CustomFormat = " "
660     calTo(0).Enabled = False
670     calTo(0).CustomFormat = " "
680     calFrom(1).Enabled = True
690     calTo(1).Enabled = True
700     calTo(1).CustomFormat = "dd/MM/yyyy"
710     calFrom(1).CustomFormat = "dd/MM/yyyy"
720     calTo(1) = Format(Now - 35, "dd/MM/yyyy")
730     calFrom(1) = Format(Now - 42, "dd/MM/yyyy")
740     txtBetween(0) = ""
750     txtBetween(0).Enabled = False
760     txtBetween(1) = ""
770     txtBetween(1).Enabled = False
780     calSchedOn(1).Enabled = False
790     calSchedOn(1).CustomFormat = " "
800     calSchedOn(0).Enabled = False
810     calSchedOn(0).CustomFormat = " "
820     calNotSchedOn.Enabled = True
830     calNotSchedOn.CustomFormat = "dd/MM/yyyy"
840     calNotSchedOn = Format(Now, "dd/MM/yyyy")


850 Case 3
860     If DataChanged = True Then
870         If frmMsgBox.Msg("Do you want to save changes to grid?", mbYesNo, "Save Changes", mbQuestion) = 1 Then
880             SaveGrid
890         End If
900     End If
910     g.Enabled = True
920     calFrom(1).Enabled = False
930     calFrom(1).CustomFormat = " "
940     calTo(1).Enabled = False
950     calTo(1).CustomFormat = " "
960     calFrom(0).Enabled = False
970     calFrom(0).CustomFormat = " "
980     calTo(0).Enabled = False
990     calTo(0).CustomFormat = " "
1000    txtBetween(0) = ""
1010    txtBetween(0).Enabled = False
1020    txtBetween(1) = ""
1030    txtBetween(1).Enabled = False
1040    calSchedOn(1).Enabled = False
1050    calSchedOn(1).CustomFormat = " "
1060    calSchedOn(0).Enabled = False
1070    calSchedOn(0).CustomFormat = " "
1080    calNotSchedOn.Enabled = False
1090    calNotSchedOn.CustomFormat = " "
1100    g.TextMatrix(0, 7) = "Kept By"
1110    g.TextMatrix(0, 4) = "Dispose"

1120 Case 4
1130    If DataChanged = True Then
1140        If frmMsgBox.Msg("Do you want to save changes to grid?", mbYesNo, "Save Changes", mbQuestion) = 1 Then
1150            SaveGrid
1160        End If
1170    End If
1180    g.Enabled = True
1190    calSchedOn(0).Enabled = False
1200    calSchedOn(0).CustomFormat = " "
1210    calSchedOn(1).Enabled = False
1220    calSchedOn(1).CustomFormat = " "
1230    calNotSchedOn.Enabled = False
1240    calNotSchedOn.CustomFormat = " "
1250    calFrom(0).Enabled = True
1260    calTo(0).Enabled = True
1270    calTo(0).CustomFormat = "dd/MM/yyyy"
1280    calFrom(0).CustomFormat = "dd/MM/yyyy"
1290    calTo(0) = Format(Now, "dd/MM/yyyy")
1300    calFrom(0) = Format(Now - 7, "dd/MM/yyyy")
1310    calFrom(1).Enabled = False
1320    calFrom(1).CustomFormat = " "
1330    calTo(1).Enabled = False
1340    calTo(1).CustomFormat = " "
1350    txtBetween(0) = ""
1360    txtBetween(0).Enabled = False
1370    txtBetween(1) = ""
1380    txtBetween(1).Enabled = False
1390    g.TextMatrix(0, 7) = "Disposed By"
1400    g.TextMatrix(0, 4) = "Disposed"
1410 End Select
1420 DataChanged = False
1430 LoadList
End Sub

Private Sub SaveGrid()
    Dim sql As String
    Dim tb As Recordset
    Dim r As Integer
    Dim SID As String
    Dim Dispose As String
    Dim DisposalDate As String

10  On Error GoTo SaveGrid_Error

20  r = 1
30  Do Until r = g.Rows
40      If g.TextMatrix(r, 9) = "1" Then

50          SID = Replace(g.TextMatrix(r, 0), " " & sysOptCaseIdSeperator(0) & " ", "")

60          sql = "SELECT * FROM CaseTree " & _
                  "WHERE CaseId = '" & SID & "' " & _
                  "AND LocationName = '" & AddTicks(g.TextMatrix(r, 1)) & "'"

70          Set tb = New Recordset
80          RecOpenClient 0, tb, sql

90          If Not tb.EOF Then
100             g.row = r
110             g.col = 4
120             Dispose = IIf(g.CellPicture = imgSquareTick.Picture, "D", "K")
130             If tb!Disposal & "" <> Dispose Then
140                 If Dispose = "D" Then
150                     CaseUpdateLogEvent SID, Disposal, " Disposed (Comments: " & g.TextMatrix(r, 5) & ")", g.TextMatrix(r, 1)
160                 Else
170                     CaseUpdateLogEvent SID, Disposal, " Kept (Comments: " & g.TextMatrix(r, 5) & ")", g.TextMatrix(r, 1)
180                 End If
190             End If
200             If Dispose = "D" Then
210                 DisposalDate = Format(Now, "dd/MM/yyyy")
220             Else
230                 DisposalDate = ""
240             End If

250             tb!Disposal = Dispose
260             tb!DisposalComment = g.TextMatrix(r, 5)
270             tb!DisposedBy = UserName
280             If DisposalDate <> "" Then
290                 tb!DisposalDate = DisposalDate
300             Else
310                 tb!DisposalDate = Null
320             End If
330             tb.Update
340         End If
350     End If
360     r = r + 1
370 Loop

380 DataSaved = True
390 DataChanged = False

400 LoadList


410 Exit Sub

SaveGrid_Error:

    Dim strES As String
    Dim intEL As Integer

420 intEL = Erl
430 strES = Err.Description
440 LogError "frmHistDisposal", "SaveGrid", intEL, strES, sql


End Sub


Private Sub txtBetween_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngSel As Long, lngLen As Long

10  With txtBetween(Index)
20      lngSel = .SelStart
30      lngLen = .SelLength

40      If lngSel = 0 Then
50          Select Case pDisposalType
            Case "H"
60              Select Case KeyAscii
                Case 104, 72
70                  KeyAscii = 72
80              Case 8, 127
90              Case Else
100                 KeyAscii = 0
110             End Select
120         Case "C"
130             Select Case KeyAscii
                Case 99, 67
140                 KeyAscii = 67
150             Case 8, 127
160             Case Else
170                 KeyAscii = 0
180             End Select
190         Case "A"
200             Select Case KeyAscii
                Case 112, 80
210                 KeyAscii = 80
220             Case 109, 77
230                 KeyAscii = 77
240             Case 116, 84
250                 KeyAscii = 84
260             Case 8, 127
270             Case Else
280                 KeyAscii = 0
290             End Select
300         End Select
310     ElseIf lngSel = 1 Then

320         If pDisposalType = "A" Then
330             Select Case KeyAscii
                Case 97, 65
340                 KeyAscii = 65
350                 lngMaxDigits = 12
360             Case 8, 127
370             Case Else
380                 KeyAscii = 0
390             End Select
400         Else
410             Select Case KeyAscii
                Case 48 To 57
420                 lngMaxDigits = 11
430             Case 8, 127
440             Case Else
450                 KeyAscii = 0
460             End Select
470         End If

480     ElseIf lngSel < lngMaxDigits Then
490         Select Case KeyAscii
            Case 32
500             If lngMaxDigits = 12 Then
510                 If lngSel = 2 Or lngSel = 3 Or lngSel = 4 _
                       Or lngSel = 5 Or lngSel = 6 Then
520                     .Text = Left(.Text, 2) & formatLeadingZero(Int(Val(Mid(.Text, 3, 5))), 5) & " /"
530                     lngSel = 10
540                 End If
550             Else

560                 If lngSel = 1 Or lngSel = 2 Or lngSel = 3 _
                       Or lngSel = 4 Or lngSel = 5 Then
570                     .Text = Left(.Text, 1) & formatLeadingZero(Int(Val(Mid(.Text, 2, 5))), 5) & " /"
580                     lngSel = 9
590                 End If
600             End If

610         Case 48 To 57
620             If lngMaxDigits = 12 Then
630                 If lngSel = 7 Or lngSel = 8 Or lngSel = 9 Then
640                     .Text = Left(.Text, 7) & " / "
650                     lngSel = 10
660                 End If
670             Else

680                 If lngSel = 6 Or lngSel = 7 Or lngSel = 8 Then
690                     .Text = Left(.Text, 6) & " / "
700                     lngSel = 9
710                 End If
720             End If
730         Case 8, 127
740         Case Else
750             KeyAscii = 0
760         End Select
770     ElseIf KeyAscii <> 8 And KeyAscii <> 127 Then
780         KeyAscii = 0
790     End If
800     .SelStart = lngSel
810     .SelLength = lngLen
820 End With
End Sub

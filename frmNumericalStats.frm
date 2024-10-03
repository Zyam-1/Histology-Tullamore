VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNumericalStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAquire - Cellular Pathology"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   Icon            =   "frmNumericalStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   975
      Left            =   7320
      Picture         =   "frmNumericalStats.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   975
      Left            =   8520
      Picture         =   "frmNumericalStats.frx":1014
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   1035
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   9720
      Picture         =   "frmNumericalStats.frx":142F
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.CommandButton cmdCalc 
         Caption         =   "&Calculate"
         Default         =   -1  'True
         Height          =   975
         Left            =   4080
         Picture         =   "frmNumericalStats.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Search"
         Top             =   720
         Width           =   1035
      End
      Begin VB.ComboBox cmbSource 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   3550
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
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   510
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7155
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   12621
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      FormatString    =   "Statistics                                          |                   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   11
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1080
      TabIndex        =   12
      Top             =   9840
      Width           =   525
   End
End
Attribute VB_Name = "frmNumericalStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()


10  ClearFGrid g

20  FillStats
30  g.Visible = True
40  cmdPrint.Enabled = True

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
60  LogError "frmNumericalStats", "cmdExport_Click", intEL, strES


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
710 LogError "frmNumericalStats", "cmdPrint_Click", intEL, strES


End Sub

Private Sub Form_Load()

10    ChangeFont Me, "Arial"
'20    frmNumericalStats_ChangeLanguage
30    calTo = Format(Now, "dd/MM/yyyy")
40    calFrom = Format(Now - 7, "dd/MM/yyyy")
50    lblLoggedIn = UserName

60    InitializeGrid

70    FillSource
80    If blnIsTestMode Then EnableTestMode Me
End Sub
Private Sub InitializeGrid()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 2: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "Statistics": .ColWidth(0) = 8500: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Count": .ColWidth(1) = 1050: .ColAlignment(1) = flexAlignLeftCenter

140 End With
End Sub

Private Sub FillSource()
    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo FillSource_Error

20  cmbSource.AddItem ""
30  sql = "SELECT * FROM Lists WHERE ListType = 'Source'"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql

60  Do While Not tb.EOF
70      cmbSource.AddItem tb!Description & ""
80      tb.MoveNext
90  Loop
100 cmbSource.ListIndex = -1

110 Exit Sub

FillSource_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmNumericalStats", "FillSource", intEL, strES, sql


End Sub

Private Sub FillStats()

    Dim rsRec As Recordset
    Dim sql As String
    Dim FromTime As String
    Dim ToTime As String
    Dim blnNotFinished As Boolean
    Dim n As Integer
    Dim strStatName As String

10  On Error GoTo FillStats_Error

20  strStatName = ""
30  blnNotFinished = True

    'Search between these dates
40  FromTime = Format(calFrom, "yyyymmdd") & " 00:00:00"
50  ToTime = Format(calTo, "yyyymmdd") & " 23:59:59"

60  Do While blnNotFinished
70      For n = 0 To 26

80          Select Case n

            Case 0:    ' Number of Histology Cases
90              sql = "SELECT COUNT(DISTINCT(C.CaseId)) AS Tot FROM Cases C " & _
                      "INNER JOIN Demographics D ON C.CaseId = D.CaseId " & _
                      "WHERE SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(C.CaseId,1,1) = 'H' "
100             If cmbSource <> "" Then
110                 sql = sql & "AND D.Source = '" & cmbSource & "' "
120             End If
130             strStatName = "Number of Histology Cases"
140         Case 1:    ' Number of Histology Specimens
150             sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "WHERE (CT.TissueTypeListId IS NOT NULL AND CT.TissueTypeListId <> '') " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(CT.CaseId,1,1) = 'H' "
160             If cmbSource <> "" Then
170                 sql = sql & "AND D.Source = '" & cmbSource & "' "
180             End If
190             strStatName = "Number of Histology Specimens"
200         Case 2:    ' Number of Histology Blocks
210             sql = "SELECT COUNT(B.CaseId) as Tot FROM BlockDetails B " & _
                      "INNER JOIN Cases C ON B.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON B.CaseId = D.CaseId " & _
                      "WHERE C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(B.CaseId,1,1) = 'H' "
220             If cmbSource <> "" Then
230                 sql = sql & "AND D.Source = '" & cmbSource & "' "
240             End If
250             strStatName = "Number of Histology Blocks"
260         Case 3:    ' Number of Histology/Autopsy Special Stain Slides
270             sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '4' OR CT.LocationLevel = '5') " & _
                      "AND L.ListType = 'SS' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND (SUBSTRING(CT.CaseId,1,1) = 'H' OR SUBSTRING(CT.CaseId,2,1) = 'A') "
280             If cmbSource <> "" Then
290                 sql = sql & "AND D.Source = '" & cmbSource & "' "
300             End If
310             strStatName = "Number of Histology/Autopsy Special Stain Slides"
320         Case 4:    ' Number of Histology/Autopsy Immunohistochemistry Stain Slides
330             sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '4' OR CT.LocationLevel = '5') " & _
                      "AND L.ListType = 'IS' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND (SUBSTRING(CT.CaseId,1,1) = 'H' OR SUBSTRING(CT.CaseId,2,1) = 'A') "
340             If cmbSource <> "" Then
350                 sql = sql & "AND D.Source = '" & cmbSource & "' "
360             End If
370             strStatName = "Number of Histology/Autopsy Immunohistochemistry Stain Slides"

380         Case 5:    ' Number of Histology/Autopsy H&E Stain Slides
390             sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '4' OR CT.LocationLevel = '5') " & _
                      "AND L.Code = 'H&E' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND (SUBSTRING(CT.CaseId,1,1) = 'H' OR SUBSTRING(CT.CaseId,2,1) = 'A') "
400             If cmbSource <> "" Then
410                 sql = sql & "AND D.Source = '" & cmbSource & "' "
420             End If
430             strStatName = "Number of Histology/Autopsy H&E Stain Slides"
440         Case 6:    ' Total Number of Slides (excluding Cytology)
450             sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "WHERE (CT.Type = 'S') " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND (SUBSTRING(CT.CaseId,1,1) = 'H' OR SUBSTRING(CT.CaseId,2,1) = 'A') "
460             If cmbSource <> "" Then
470                 sql = sql & "AND D.Source = '" & cmbSource & "' "
480             End If
490             strStatName = "Total Number of Histology Slides (excluding Cytology)"
500         Case 7:    ' Number of Cases with Special Stains Requested
510             sql = "SELECT COUNT(DISTINCT(CT.CaseId)) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '4' OR CT.LocationLevel = '5') " & _
                      "AND L.ListType = 'SS' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
520             If cmbSource <> "" Then
530                 sql = sql & "AND D.Source = '" & cmbSource & "' "
540             End If
550             strStatName = "Number of Cases with Special Stains Requested"
560         Case 8:    ' Number of Cases with Immunohistochemistry Stains Requested
570             sql = "SELECT COUNT(DISTINCT(CT.CaseId)) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '4' OR CT.LocationLevel = '5') " & _
                      "AND L.ListType = 'IS' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
580             If cmbSource <> "" Then
590                 sql = sql & "AND D.Source = '" & cmbSource & "' "
600             End If
610             strStatName = "Number of Cases with Immunohistochemistry Stains Requested"
620         Case 9:    ' Number of Cases for each Immunohistochemistry Stain Requested
630             sql = "SELECT CT.LocationName AS Filter, COUNT(DISTINCT(CT.CaseId)) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '4' OR CT.LocationLevel = '5') " & _
                      "AND L.ListType = 'IS' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "

640             If cmbSource <> "" Then
650                 sql = sql & "AND D.Source = '" & cmbSource & "' "
660             End If
670             sql = sql & "GROUP BY CT.LocationName "
680             strStatName = "Number of Cases for each Immunohistochemistry Stain Requested"
690         Case 10:    ' Number of Cases for each Special Stain Requested
700             sql = "SELECT CT.LocationName AS Filter, COUNT(DISTINCT(CT.CaseId)) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '4' OR CT.LocationLevel = '5') " & _
                      "AND L.ListType = 'SS' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
710             If cmbSource <> "" Then
720                 sql = sql & "AND D.Source = '" & cmbSource & "' "
730             End If
740             sql = sql & "GROUP BY CT.LocationName "
750             strStatName = "Number of Cases for each Special Stain Requested"
760         Case 11:    ' Number of Cytology Cases
770             sql = "SELECT COUNT(DISTINCT(C.CaseId)) AS Tot FROM Cases C " & _
                      "INNER JOIN Demographics D ON C.CaseId = D.CaseId " & _
                      "WHERE SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(C.CaseId,1,1) = 'C' "
780             If cmbSource <> "" Then
790                 sql = sql & "AND D.Source = '" & cmbSource & "' "
800             End If
810             strStatName = "Number of Cytology Cases"
820         Case 12:    ' Number of Cytology Specimens
830             sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "WHERE (CT.TissueTypeListId IS NOT NULL AND CT.TissueTypeListId <> '') " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(CT.CaseId,1,1) = 'C' "
840             If cmbSource <> "" Then
850                 sql = sql & "AND D.Source = '" & cmbSource & "' "
860             End If
870             strStatName = "Number of Cytology Specimens"
880         Case 13:    ' Number of Cytology Slides
890             sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "WHERE (CT.Type = 'S') " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(CT.CaseId,1,1) = 'C' "
900             If cmbSource <> "" Then
910                 sql = sql & "AND D.Source = '" & cmbSource & "' "
920             End If
930             strStatName = "Number of Cytology Slides"
940         Case 14:    ' Number of MGG Slides
950             sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '3' OR CT.LocationLevel = '4') " & _
                      "AND L.Code = 'MGG' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(CT.CaseId,1,1) = 'C' "
960             If cmbSource <> "" Then
970                 sql = sql & "AND D.Source = '" & cmbSource & "' "
980             End If
990             strStatName = "Number of MGG Slides"
1000        Case 15:    ' Number of PAP Slides
1010            sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '3' OR CT.LocationLevel = '4') " & _
                      "AND L.Code = 'PAP' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(CT.CaseId,1,1) = 'C' "
1020            If cmbSource <> "" Then
1030                sql = sql & "AND D.Source = '" & cmbSource & "' "
1040            End If
1050            strStatName = "Number of PAP Slides"
1060        Case 16:    'Number of Cytoclots
1070            sql = "SELECT COUNT(DISTINCT(C.CaseId)) AS Tot FROM Cases C " & _
                      "INNER JOIN Demographics D ON C.CaseId = D.CaseId " & _
                      "WHERE (C.LinkedCaseId IS NOT NULL AND C.LinkedCaseId <> '') " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(C.CaseId,1,1) = 'C' "
1080            If cmbSource <> "" Then
1090                sql = sql & "AND D.Source = '" & cmbSource & "' "
1100            End If
1110            strStatName = "Number of Cytoclots"
1120        Case 17:    'Number of Autopsies
1130            sql = "SELECT COUNT(DISTINCT(C.CaseId)) AS Tot FROM Cases C " & _
                      "INNER JOIN Demographics D ON C.CaseId = D.CaseId " & _
                      "WHERE C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(C.CaseId,2,1) = 'A' "
1140            If cmbSource <> "" Then
1150                sql = sql & "AND D.Source = '" & cmbSource & "' "
1160            End If
1170            strStatName = "Number of Autopsies"
1180        Case 18:    'Number of Autopsy Blocks
1190            sql = "SELECT COUNT(B.CaseId) as Tot FROM BlockDetails B " & _
                      "INNER JOIN Cases C ON B.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON B.CaseId = D.CaseId " & _
                      "WHERE C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(B.CaseId,2,1) = 'A' "
1200            If cmbSource <> "" Then
1210                sql = sql & "AND D.Source = '" & cmbSource & "' "
1220            End If
1230            strStatName = "Number of Autopsy Blocks"
1240        Case 19:    'Number of Autopsy Slides (H&E)
1250            sql = "SELECT COUNT(CT.CaseId) as Tot FROM CaseTree CT " & _
                      "INNER JOIN Cases C ON CT.CaseId = C.CaseId " & _
                      "INNER JOIN Demographics D ON CT.CaseId = D.CaseId " & _
                      "INNER JOIN Lists L ON CT.LocationName = L.Description " & _
                      "WHERE (CT.LocationLevel = '4' OR CT.LocationLevel = '5') " & _
                      "AND L.Code = 'H&E' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' " & _
                      "AND SUBSTRING(CT.CaseId,2,1) = 'A' "
1260            If cmbSource <> "" Then
1270                sql = sql & "AND D.Source = '" & cmbSource & "' "
1280            End If
1290            strStatName = "Number of Autopsy Slides (H&E)"
1300        Case 20:    ' Number of Histology/Cytology Cases by County
1310            sql = "SELECT County As Filter,COUNT(DISTINCT(D.CaseId)) as Tot FROM Demographics D " & _
                      "INNER JOIN Cases C ON D.CaseId = C.CaseId " & _
                      "WHERE SUBSTRING(D.CaseId,2,1) <> 'A' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
1320            If cmbSource <> "" Then
1330                sql = sql & "AND D.Source = '" & cmbSource & "' "
1340            End If
1350            sql = sql & "GROUP BY County "
1360            strStatName = "Number of Histology/Cytology Cases by County"
1370        Case 21:    'Number of Histology/Cytology Cases per Clinician
1380            sql = "SELECT Clinician As Filter,COUNT(DISTINCT(D.CaseId)) as Tot FROM Demographics D " & _
                      "INNER JOIN Cases C ON D.CaseId = C.CaseId " & _
                      "WHERE SUBSTRING(D.CaseId,2,1) <> 'A' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
1390            If cmbSource <> "" Then
1400                sql = sql & "AND D.Source = '" & cmbSource & "' "
1410            End If
1420            sql = sql & "GROUP BY Clinician "
1430            strStatName = "Number of Histology/Cytology Cases per Clinician"
1440        Case 22:    'Number of Histology/Cytology Cases per GP
1450            sql = "SELECT GP As Filter,COUNT(DISTINCT(D.CaseId)) as Tot FROM Demographics D " & _
                      "INNER JOIN Cases C ON D.CaseId = C.CaseId " & _
                      "WHERE SUBSTRING(D.CaseId,2,1) <> 'A' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
1460            If cmbSource <> "" Then
1470                sql = sql & "AND D.Source = '" & cmbSource & "' "
1480            End If
1490            sql = sql & "GROUP BY GP "
1500            strStatName = "Number of Histology/Cytology Cases per GP"
1510        Case 23:    'Number of Histology/Cytology Specimens per Clinician
1520            sql = "SELECT Clinician As Filter,COUNT(D.CaseId) as Tot FROM Demographics D " & _
                      "INNER JOIN Cases C ON D.CaseId = C.CaseId " & _
                      "LEFT JOIN CaseTree CT ON D.CaseId = CT.CaseId " & _
                      "WHERE (CT.TissueTypeListId IS NOT NULL AND CT.TissueTypeListId <> '') " & _
                      "AND SUBSTRING(D.CaseId,2,1) <> 'A' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
1530            If cmbSource <> "" Then
1540                sql = sql & "AND D.Source = '" & cmbSource & "' "
1550            End If
1560            sql = sql & "GROUP BY Clinician "
1570            strStatName = "Number of Histology/Cytology Specimens per Clinician"
1580        Case 24:    'Number of Histology/Cytology Specimens per GP
1590            sql = "SELECT GP As Filter,COUNT(D.CaseId) as Tot FROM Demographics D " & _
                      "INNER JOIN Cases C ON D.CaseId = C.CaseId " & _
                      "LEFT JOIN CaseTree CT ON D.CaseId = CT.CaseId " & _
                      "WHERE (CT.TissueTypeListId IS NOT NULL AND CT.TissueTypeListId <> '') " & _
                      "AND SUBSTRING(D.CaseId,2,1) <> 'A' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
1600            If cmbSource <> "" Then
1610                sql = sql & "AND D.Source = '" & cmbSource & "' "
1620            End If
1630            sql = sql & "GROUP BY GP "
1640            strStatName = "Number of Histology/Cytology Specimens per GP"
1650        Case 25:    'Number of Histology/Cytology Cases By Ward
1660            sql = "SELECT Ward As Filter,COUNT(DISTINCT(D.CaseId)) as Tot FROM Demographics D " & _
                      "INNER JOIN Cases C ON D.CaseId = C.CaseId " & _
                      "WHERE SUBSTRING(D.CaseId,2,1) <> 'A' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
1670            If cmbSource <> "" Then
1680                sql = sql & "AND Source = '" & cmbSource & "' "
1690            End If
1700            sql = sql & "GROUP BY Ward "
1710            strStatName = "Number of Histology/Cytology Cases By Ward"
1720        Case 26:    'Number of Histology/Cytology Specimens By Ward
1730            sql = "SELECT Ward As Filter,COUNT(D.CaseId) as Tot FROM Demographics D " & _
                      "INNER JOIN Cases C ON D.CaseId = C.CaseId " & _
                      "LEFT JOIN CaseTree CT ON D.CaseId = CT.CaseId " & _
                      "WHERE (CT.TissueTypeListId IS NOT NULL AND CT.TissueTypeListId <> '') " & _
                      "AND SUBSTRING(D.CaseId,2,1) <> 'A' " & _
                      "AND C.SampleReceived BETWEEN " & _
                    " '" & FromTime & "'  AND '" & ToTime & "' "
1740            If cmbSource <> "" Then
1750                sql = sql & "AND D.Source = '" & cmbSource & "' "
1760            End If
1770            sql = sql & "GROUP BY Ward "
1780            strStatName = "Number of Histology/Cytology Specimens By Ward"
1790        End Select



1800        If sql <> "" Then
1810            If n = 9 Or n = 10 Or n = 20 Or n = 21 Or n = 22 Or _
                 n = 23 Or n = 24 Or n = 25 Or n = 26 Then
1820                Set rsRec = New Recordset
1830                RecOpenServer 0, rsRec, sql
1840                Do While Not rsRec.EOF
1850                    If rsRec!Filter & "" <> "" Then
1860                        g.AddItem strStatName & " (" & rsRec!Filter & ")" & vbTab & rsRec!tot
1870                    End If
1880                    rsRec.MoveNext
1890                Loop
1900                sql = ""
1910            Else
1920                Set rsRec = New Recordset
1930                RecOpenServer 0, rsRec, sql
1940                If Not rsRec.EOF Then
1950                    g.AddItem strStatName & vbTab & rsRec!tot & ""
1960                    rsRec.MoveNext
1970                End If
1980                sql = ""
1990            End If
2000        End If
2010    Next

2020    g.Visible = True
2030    If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1

2040    blnNotFinished = False
2050 Loop

2060 Exit Sub

FillStats_Error:

    Dim strES As String
    Dim intEL As Integer

2070 intEL = Erl
2080 strES = Err.Description
2090 LogError "frmNumericalStats", "FillStats", intEL, strES, sql

End Sub
Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then
20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

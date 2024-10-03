VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmCaseEventLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15225
   Icon            =   "frmCaseEventLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   5640
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   735
      Left            =   10080
      Picture         =   "frmCaseEventLog.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   11400
      TabIndex        =   6
      Top             =   120
      Width           =   3495
      Begin VB.Label lblAll 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "All"
         Height          =   270
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblChanges 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Changes"
         Height          =   270
         Left            =   195
         TabIndex        =   8
         Top             =   225
         Width           =   915
      End
      Begin VB.Label lblEvents 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Events"
         Height          =   270
         Left            =   1320
         TabIndex        =   7
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   735
      Left            =   3240
      Picture         =   "frmCaseEventLog.frx":12E5
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   4440
      Picture         =   "frmCaseEventLog.frx":1663
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2730
      Begin VB.TextBox txtCaseId 
         Height          =   285
         Left            =   900
         MaxLength       =   12
         TabIndex        =   2
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Case ID"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   285
         Width           =   570
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdReport 
      Height          =   8100
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   14288
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   12648447
      ForeColor       =   -2147483625
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      AllowUserResizing=   1
      FormatString    =   $"frmCaseEventLog.frx":19A5
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   14655
      _ExtentX        =   25850
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
      TabIndex        =   16
      Top             =   9120
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   9120
      Width           =   1335
   End
End
Attribute VB_Name = "frmCaseEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SID As String
Private SortOrder As Boolean
Public ExternalCaseID As String

Private Sub cmdExit_Click()
10  Unload Me
End Sub
Public Sub RunReport()

10  ClearGrid
20  LoadDemographicChanges
30  LoadCaseDetailChanges
40  LoadChangesOtherCodes
50  LoadDiscrepancyChanges
60  LoadEventLog

70  grdReport.Visible = True
    'Sort by datetime
80  grdReport.Col = 0
90  grdReport.Sort = 9

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

20  xmin = 1440
30  ymin = 1560

    Dim lRowsPrinted As Long, lRowsPerPage As Long
    Dim lThisRow As Long, lNumRows As Long
    Dim lPrinterPageHeight As Long
    Dim lPrintPage As Long
    Dim lNoOfPages As Long

40  grdReport.TopRow = 1
50  lNumRows = grdReport.Rows - 1
60  lPrinterPageHeight = Printer.Height
70  lRowsPerPage = 29
80  lRowsPrinted = 1



90  xmax = xmin + GAP
100 For c = 0 To grdReport.Cols - 1
110     xmax = xmax + grdReport.ColWidth(c) + 2 * GAP
120 Next c

130 lPrintPage = 1
140 lNoOfPages = Int(lNumRows / lRowsPerPage) + 1
150 Do

160     Printer.Orientation = 2
170     Do

180         With grdReport

190             PrintHeadingAuditTrail txtCaseId, "Page " & lPrintPage & " of " & lNoOfPages

                ' Print each row.
200             Printer.CurrentY = ymin

210             Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)

220             Printer.CurrentY = Printer.CurrentY + GAP

230             X = xmin + GAP
240             For c = 0 To .Cols - 1
250                 Printer.CurrentX = X
260                 PrintText BoundedText(Printer, .TextMatrix(0, c), .ColWidth(c)), "MS Sans Serif", , True
270                 X = X + .ColWidth(c) + 2 * GAP
280             Next c
290             Printer.CurrentY = Printer.CurrentY + GAP

                ' Move to the next line.
300             PrintText vbCrLf

310             For lThisRow = lRowsPrinted To lRowsPerPage * lPrintPage

320                 If lThisRow < lNumRows Then
330                     If lThisRow > 0 Then Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)
340                     Printer.CurrentY = Printer.CurrentY + GAP

                        ' Print the entries on this row.
350                     X = xmin + GAP
360                     For c = 0 To .Cols - 1
370                         Printer.CurrentX = X
380                         PrintText BoundedText(Printer, .TextMatrix(lThisRow, c), .ColWidth(c)), "MS Sans Serif", 8
390                         X = X + .ColWidth(c) + 2 * GAP
400                     Next c
410                     Printer.CurrentY = Printer.CurrentY + GAP

                        ' Move to the next line.
420                     PrintText vbCrLf

430                     lRowsPrinted = lRowsPrinted + 1
440                 Else
450                     Exit Do
460                 End If
470             Next
480         End With
490     Loop While lRowsPrinted < lRowsPerPage * lPrintPage

500     ymax = Printer.CurrentY

        ' Draw a box around everything.
510     Printer.Line (xmin, ymin)-(xmax, ymax), , B

        ' Draw lines between the columns.
520     X = xmin
530     For c = 0 To grdReport.Cols - 2
540         X = X + grdReport.ColWidth(c) + 2 * GAP
550         Printer.Line (X, ymin)-(X, ymax)
560     Next c

570     Printer.EndDoc
580     lPrintPage = lPrintPage + 1

590 Loop While lRowsPrinted < lNumRows

600 Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

610 intEL = Erl
620 strES = Err.Description
630 LogError "frmCaseEventLog", "cmdPrint_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10  On Error GoTo Form_Activate_Error



20  Exit Sub

Form_Activate_Error:

    Dim strES As String
    Dim intEL As Integer

30  intEL = Erl
40  strES = Err.Description
50  LogError "frmCaseEventLog", "Form_Activate", intEL, strES

End Sub

Private Sub Form_DblClick()

If IsIDE Then
    If sysOptCurrentLanguage = "Russian" Then
        sysOptCurrentLanguage = "English"
    Else
        sysOptCurrentLanguage = "Russian"
    End If
    
    LoadLanguage sysOptCurrentLanguage
'    frmCaseEventLog_ChangeLanguage
End If

End Sub




Private Sub Form_Load()
10  On Error GoTo Form_Load_Error

20  InitializeGrid
30  lblLoggedIn = UserName
40  If SID <> "" Then
50      txtCaseId = SID
60      SID = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
70      RunReport
80  End If
90  If blnIsTestMode Then EnableTestMode Me
100 Exit Sub

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmCaseEventLog", "Form_Load", intEL, strES


End Sub



Private Sub grdReport_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

    Dim d1 As String
    Dim d2 As String

10  If Not IsDate(grdReport.TextMatrix(Row1, grdReport.Col)) Then
20      Cmp = 0
30      Exit Sub
40  End If

50  If Not IsDate(grdReport.TextMatrix(Row2, grdReport.Col)) Then
60      Cmp = 0
70      Exit Sub
80  End If

90  d1 = Format(grdReport.TextMatrix(Row1, grdReport.Col), "dd/mmm/yyyy hh:mm:ss")
100 d2 = Format(grdReport.TextMatrix(Row2, grdReport.Col), "dd/mmm/yyyy hh:mm:ss")

110 If SortOrder Then
120     Cmp = Sgn(DateDiff("s", d1, d2))
130 Else
140     Cmp = Sgn(DateDiff("s", d2, d1))
150 End If


End Sub

Private Sub grdReport_Click()

10  If grdReport.MouseRow = 0 Then
20      If grdReport.MouseCol = 0 Then
30          grdReport.Col = grdReport.MouseCol
40          grdReport.Sort = 9
50      Else
60          If SortOrder Then
70              grdReport.Sort = flexSortGenericAscending
80          Else
90              grdReport.Sort = flexSortGenericDescending
100         End If
110     End If

120     SortOrder = Not SortOrder
130     Exit Sub
140 End If

End Sub

Private Sub ClearGrid()

10  grdReport.Rows = 2
20  grdReport.AddItem ""
30  grdReport.RemoveItem 1

End Sub

Private Sub LoadEventLog()
    Dim rsRec As Recordset
    Dim sql As String
    Dim s As String
    Dim c As Integer


10  On Error GoTo LoadEventLog_Error

20  sql = "SELECT * from CaseEventLog WHERE CaseId = N'" & SID & "' "
30  Set rsRec = New Recordset
40  RecOpenClient 0, rsRec, sql

50  If Not rsRec.EOF Then
60      pbProgress.Max = rsRec.RecordCount + 1
70      grdReport.Visible = False
80      fraProgress.Visible = True
90  End If

100 Do While Not rsRec.EOF
110     pbProgress.Value = pbProgress.Value + 1
120     lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
130     lblProgress.Refresh

140     s = Format(rsRec!DateTimeOfRecord, "dd/mm/yy hh:mm:ss") & vbTab
150     If rsRec!Path & "" <> "" Then
160         s = s & Mid(rsRec!Path, InStr(rsRec!Path, "\") + 1) & " : "
170     End If
180     s = s & rsRec!EventDesc & " " & rsRec!Comments & vbTab
190     s = s & rsRec!UserName & "" & vbTab
200     If rsRec!Path & "" <> "" Then
210         s = s & Left(Mid(rsRec!Path, InStr(rsRec!Path, "\") + 1), 1)
220     End If
230     grdReport.AddItem s
240     For c = 0 To grdReport.Cols - 1
250         grdReport.Col = c
260         grdReport.Row = grdReport.Rows - 1
270         grdReport.CellBackColor = &H80FFFF
280     Next
290     rsRec.MoveNext
300 Loop

310 fraProgress.Visible = False
320 pbProgress.Value = 1

330 If grdReport.Rows > 2 And grdReport.TextMatrix(1, 0) = "" Then
340     grdReport.RemoveItem 1
350 End If


360 Exit Sub

LoadEventLog_Error:
370 grdReport.Visible = True
380 fraProgress.Visible = False
    Dim strES As String
    Dim intEL As Integer

390 intEL = Erl
400 strES = Err.Description
410 LogError "frmCaseEventLog", "LoadEventLog", intEL, strES, sql


End Sub
Private Sub LoadDemographicChanges()

    Dim tb As Recordset
    Dim sql As String
    Dim LatestValue(1 To 33) As String
    Dim dbName(1 To 33) As String
    Dim ShowName(1 To 33) As String
    Dim n As Integer
    Dim X As Integer
    Dim s As String



10  On Error GoTo LoadDemographicChanges_Error

20  For n = 1 To 33
30      dbName(n) = Choose(n, "CaseId", "NOPAS", "MRN", "AandENo", _
                           "FirstName", "Surname", "PatientName", _
                           "Address1", "Address2", "Address3", "Address4", "County", _
                           "DateOfBirth", "Age", "Sex", "Clinician", "GP", "Ward", _
                           "Phone", "Source", "Coroner", "NatureOfSpecimen", "SpecimenLabelled", _
                           "ClinicalHistory", "Comments", "AutopsyFor", "AutopsyRequestedBy", _
                           "DateOfDeath", "MothersName", "MothersDOB", "PaedType", "NoHistTaken", "Urgent", "Year")

40      ShowName(n) = Choose(n, "Case Id", "NOPAS Number", "MRN Number", "A&E Number", _
                             "First Name", "Surname", "Patient Name", _
                             "Address", "Address", "Address", "Address", "County", _
                             "Date Of Birth", "Age", "Sex", "Clinician", "GP", "Ward", _
                             "Phone", "Source", "Coroner", "Nature Of Specimen", "Specimen Labelled", _
                             "Clinical Details", "Comments", "Autopsy For", "Autopsy Requested By", _
                             "Date Of Death", "Mothers Name", "Mothers DOB", "PaedType", "No Histology Taken", "Urgent", "Year")
50  Next

60  sql = "SELECT " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(NOPAS)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NOPAS, '<BLANK>') END NOPAS, " & _
          "CASE LTRIM(RTRIM(MRN)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MRN, '<BLANK>') END MRN, " & _
          "CASE LTRIM(RTRIM(AandENo)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AandENo, '<BLANK>') END AandENo, " & _
          "CASE LTRIM(RTRIM(FirstName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(FirstName, '<BLANK>') END FirstName, " & _
          "CASE LTRIM(RTRIM(Surname)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Surname, '<BLANK>') END Surname, " & _
          "CASE LTRIM(RTRIM(PatientName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PatientName, '<BLANK>') END PatientName, " & _
          "CASE LTRIM(RTRIM(Address1)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address1, '<BLANK>') END Address1, " & _
          "CASE LTRIM(RTRIM(Address2)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address2, '<BLANK>') END Address2, " & _
          "CASE LTRIM(RTRIM(Address3)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address3, '<BLANK>') END Address3, " & _
          "CASE LTRIM(RTRIM(Address4)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address4, '<BLANK>') END Address4, " & _
          "CASE LTRIM(RTRIM(County)) WHEN '' THEN '<BLANK>' ELSE ISNULL(County, '<BLANK>') END County, " & _
          "CASE LTRIM(RTRIM(DateOfBirth)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfBirth, 103), '<BLANK>') END DateOfBirth, " & _
          "CASE LTRIM(RTRIM(Age)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Age, '<BLANK>') END Age, " & _
          "CASE LTRIM(RTRIM(Sex)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Sex, '<BLANK>') END Sex, " & _
          "CASE LTRIM(RTRIM(Clinician)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Clinician, '<BLANK>') END Clinician, " & _
          "CASE LTRIM(RTRIM(GP)) WHEN '' THEN '<BLANK>' ELSE ISNULL(GP, '<BLANK>') END GP, " & _
          "CASE LTRIM(RTRIM(Ward)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Ward, '<BLANK>') END Ward, " & _
          "CASE LTRIM(RTRIM(Phone)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Phone, '<BLANK>') END Phone, " & _
          "CASE LTRIM(RTRIM(Source)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Source, '<BLANK>') END Source, " & _
          "CASE LTRIM(RTRIM(Coroner)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Coroner, '<BLANK>') END Coroner, " & _
          "CASE LTRIM(RTRIM(NatureOfSpecimen)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NatureOfSpecimen, '<BLANK>') END NatureOfSpecimen, "

70  sql = sql & _
          "CASE LTRIM(RTRIM(SpecimenLabelled)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SpecimenLabelled, '<BLANK>') END SpecimenLabelled, " & _
          "CASE LTRIM(RTRIM(ClinicalHistory)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ClinicalHistory, '<BLANK>') END ClinicalHistory, " & _
          "CASE LTRIM(RTRIM(Comments)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Comments, '<BLANK>') END Comments, " & _
          "CASE LTRIM(RTRIM(AutopsyFor)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AutopsyFor, '<BLANK>') END AutopsyFor, " & _
          "CASE LTRIM(RTRIM(AutopsyRequestedBy)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AutopsyRequestedBy, '<BLANK>') END AutopsyRequestedBy, " & _
          "CASE LTRIM(RTRIM(DateOfDeath)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfDeath, 103), '<BLANK>') END DateOfDeath, " & _
          "CASE LTRIM(RTRIM(MothersName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MothersName, '<BLANK>') END MothersName, " & _
          "CASE LTRIM(RTRIM(MothersDOB)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), MothersDOB, 103), '<BLANK>') END MothersDOB, " & _
          "CASE LTRIM(RTRIM(PaedType)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PaedType, '<BLANK>') END PaedType, " & _
          "CASE NoHistTaken WHEN 1 THEN 'No Histology Taken' ELSE 'Histology Taken' END NoHistTaken, " & _
          "CASE Urgent WHEN 1 THEN 'Urgent' ELSE 'Not Urgent' END Urgent "
    '"CASE LTRIM(RTRIM(Year)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Year, '<BLANK>') END Year "

    '"CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
     '"CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

80  sql = sql & _
          "FROM Demographics WHERE " & _
          "CaseID = N'" & SID & "'"

90  Set tb = New Recordset
100 RecOpenClient 0, tb, sql
110 If Not tb.EOF Then
120     For n = 1 To 33
130         LatestValue(n) = tb(dbName(n)) & ""
140     Next
150 Else
160     For n = 1 To 33
170         LatestValue(n) = "<BLANK>"
180     Next
190 End If

200 sql = "SELECT ArchiveDateTime, " & _
          "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(NOPAS)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NOPAS, '<BLANK>') END NOPAS, " & _
          "CASE LTRIM(RTRIM(MRN)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MRN, '<BLANK>') END MRN, " & _
          "CASE LTRIM(RTRIM(AandENo)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AandENo, '<BLANK>') END AandENo, " & _
          "CASE LTRIM(RTRIM(FirstName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(FirstName, '<BLANK>') END FirstName, " & _
          "CASE LTRIM(RTRIM(Surname)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Surname, '<BLANK>') END Surname, " & _
          "CASE LTRIM(RTRIM(PatientName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PatientName, '<BLANK>') END PatientName, " & _
          "CASE LTRIM(RTRIM(Address1)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address1, '<BLANK>') END Address1, " & _
          "CASE LTRIM(RTRIM(Address2)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address2, '<BLANK>') END Address2, " & _
          "CASE LTRIM(RTRIM(Address3)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address3, '<BLANK>') END Address3, " & _
          "CASE LTRIM(RTRIM(Address4)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address4, '<BLANK>') END Address4, " & _
          "CASE LTRIM(RTRIM(County)) WHEN '' THEN '<BLANK>' ELSE ISNULL(County, '<BLANK>') END County, " & _
          "CASE LTRIM(RTRIM(DateOfBirth)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfBirth, 103), '<BLANK>') END DateOfBirth, " & _
          "CASE LTRIM(RTRIM(Age)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Age, '<BLANK>') END Age, " & _
          "CASE LTRIM(RTRIM(Sex)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Sex, '<BLANK>') END Sex, " & _
          "CASE LTRIM(RTRIM(Clinician)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Clinician, '<BLANK>') END Clinician, " & _
          "CASE LTRIM(RTRIM(GP)) WHEN '' THEN '<BLANK>' ELSE ISNULL(GP, '<BLANK>') END GP, " & _
          "CASE LTRIM(RTRIM(Ward)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Ward, '<BLANK>') END Ward, " & _
          "CASE LTRIM(RTRIM(Phone)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Phone, '<BLANK>') END Phone, " & _
          "CASE LTRIM(RTRIM(Source)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Source, '<BLANK>') END Source, " & _
          "CASE LTRIM(RTRIM(Coroner)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Coroner, '<BLANK>') END Coroner, " & _
          "CASE LTRIM(RTRIM(NatureOfSpecimen)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NatureOfSpecimen, '<BLANK>') END NatureOfSpecimen, "

210 sql = sql & _
          "CASE LTRIM(RTRIM(SpecimenLabelled)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SpecimenLabelled, '<BLANK>') END SpecimenLabelled, " & _
          "CASE LTRIM(RTRIM(ClinicalHistory)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ClinicalHistory, '<BLANK>') END ClinicalHistory, " & _
          "CASE LTRIM(RTRIM(Comments)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Comments, '<BLANK>') END Comments, " & _
          "CASE LTRIM(RTRIM(AutopsyFor)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AutopsyFor, '<BLANK>') END AutopsyFor, " & _
          "CASE LTRIM(RTRIM(AutopsyRequestedBy)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AutopsyRequestedBy, '<BLANK>') END AutopsyRequestedBy, " & _
          "CASE LTRIM(RTRIM(DateOfDeath)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfDeath, 103), '<BLANK>') END DateOfDeath, " & _
          "CASE LTRIM(RTRIM(MothersName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MothersName, '<BLANK>') END MothersName, " & _
          "CASE LTRIM(RTRIM(MothersDOB)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), MothersDOB, 103), '<BLANK>') END MothersDOB, " & _
          "CASE LTRIM(RTRIM(PaedType)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PaedType, '<BLANK>') END PaedType, " & _
          "CASE NoHistTaken WHEN 1 THEN 'No Histology Taken' ELSE 'Histology Taken' END NoHistTaken, " & _
          "CASE Urgent WHEN 1 THEN 'Urgent' ELSE 'Not Urgent' END Urgent "

220 sql = sql & _
          "FROM DemographicsAudit WHERE " & _
          "CaseID = N'" & SID & "' " & _
          "ORDER BY ArchiveDateTime DESC"

230 Set tb = New Recordset
240 RecOpenClient 0, tb, sql


250 If Not tb.EOF Then
260     pbProgress.Max = tb.RecordCount + 1
270     grdReport.Visible = False
280     fraProgress.Visible = True
290 End If

300 Do While Not tb.EOF
310     pbProgress.Value = pbProgress.Value + 1
320     lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
330     lblProgress.Refresh

340     For X = 1 To 33
350         If LatestValue(X) <> tb(dbName(X)) & "" Then
360             s = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss") & vbTab & _
                    ShowName(X) & " changed from " & tb.Fields(dbName(X)) & " to " & LatestValue(X) & vbTab & _
                    tb!ArchivedBy
370             grdReport.AddItem s

380         End If
390     Next
400     For X = 1 To 33
410         LatestValue(X) = tb(dbName(X))
420     Next

430     tb.MoveNext
440 Loop
450 fraProgress.Visible = False
460 pbProgress.Value = 1

470 If grdReport.Rows > 2 And grdReport.TextMatrix(1, 0) = "" Then
480     grdReport.RemoveItem 1
490 End If



500 Exit Sub

LoadDemographicChanges_Error:
510 grdReport.Visible = True
520 fraProgress.Visible = False
    Dim strES As String
    Dim intEL As Integer

530 intEL = Erl
540 strES = Err.Description
550 LogError "frmCaseEventLog", "LoadDemographicChanges", intEL, strES, sql


End Sub

Private Sub LoadCaseDetailChanges()

    Dim tb As Recordset
    Dim sql As String
    Dim LatestValue(1 To 6) As String
    Dim dbName(1 To 6) As String
    Dim ShowName(1 To 6) As String
    Dim n As Integer
    Dim X As Integer
    Dim s As String


10  On Error GoTo LoadCaseDetailChanges_Error

20  For n = 1 To 6
30      dbName(n) = Choose(n, "CaseId", "State", "SampleTaken", "SampleReceived", _
                           "LinkedCaseId", "Phase")

40      ShowName(n) = Choose(n, "CaseId", "State", "Sample Taken", "Sample Received", _
                             "Linked CaseId", "Phase")
50  Next

60  sql = "SELECT " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(State)) WHEN '' THEN '<BLANK>' ELSE ISNULL(State, '<BLANK>') END State, " & _
          "CASE LTRIM(RTRIM(SampleTaken)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), SampleTaken, 103) + ' ' + CONVERT(nvarchar(50), SampleTaken, 108), '<BLANK>') END SampleTaken, " & _
          "CASE LTRIM(RTRIM(SampleReceived)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), SampleReceived, 103) + ' ' + CONVERT(nvarchar(50), SampleReceived, 108), '<BLANK>') END SampleReceived, " & _
          "CASE LTRIM(RTRIM(LinkedCaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(LinkedCaseId, '<BLANK>') END LinkedCaseId, " & _
          "CASE LTRIM(RTRIM(Phase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Phase, '<BLANK>') END Phase "

70  sql = sql & _
          "FROM Cases WHERE " & _
          "CaseID = N'" & SID & "'"

80  Set tb = New Recordset
90  RecOpenClient 0, tb, sql
100 If Not tb.EOF Then
110     For n = 1 To 6
120         LatestValue(n) = tb(dbName(n)) & ""
130     Next
140 Else
150     For n = 1 To 6
160         LatestValue(n) = "<BLANK>"
170     Next
180 End If

190 sql = "SELECT ArchiveDateTime, " & _
          "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(State)) WHEN '' THEN '<BLANK>' ELSE ISNULL(State, '<BLANK>') END State, " & _
          "CASE LTRIM(RTRIM(SampleTaken)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), SampleTaken, 103) + ' ' + CONVERT(nvarchar(50), SampleTaken, 108), '<BLANK>') END SampleTaken, " & _
          "CASE LTRIM(RTRIM(SampleReceived)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), SampleReceived, 103) + ' ' + CONVERT(nvarchar(50), SampleReceived, 108), '<BLANK>') END SampleReceived, " & _
          "CASE LTRIM(RTRIM(LinkedCaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(LinkedCaseId, '<BLANK>') END LinkedCaseId, " & _
          "CASE LTRIM(RTRIM(Phase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Phase, '<BLANK>') END Phase "

200 sql = sql & _
          "FROM CasesAudit WHERE " & _
          "CaseID = N'" & SID & "' " & _
          "ORDER BY ArchiveDateTime DESC"

210 Set tb = New Recordset
220 RecOpenClient 0, tb, sql

230 If Not tb.EOF Then
240     pbProgress.Max = tb.RecordCount + 1
250     grdReport.Visible = False
260     fraProgress.Visible = True
270 End If

280 Do While Not tb.EOF
290     pbProgress.Value = pbProgress.Value + 1
300     lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
310     lblProgress.Refresh

320     For X = 1 To 6
330         If LatestValue(X) <> tb(dbName(X)) & "" Then
340             s = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss") & vbTab & _
                    ShowName(X) & " changed from " & tb.Fields(dbName(X)) & " to " & LatestValue(X) & vbTab & _
                    tb!ArchivedBy
350             grdReport.AddItem s
360         End If
370     Next
380     For X = 1 To 6
390         LatestValue(X) = tb(dbName(X))
400     Next

410     tb.MoveNext
420 Loop
430 fraProgress.Visible = False
440 pbProgress.Value = 1

450 If grdReport.Rows > 2 And grdReport.TextMatrix(1, 0) = "" Then
460     grdReport.RemoveItem 1
470 End If


480 Exit Sub

LoadCaseDetailChanges_Error:
490 grdReport.Visible = True
500 fraProgress.Visible = False
    Dim strES As String
    Dim intEL As Integer

510 intEL = Erl
520 strES = Err.Description
530 LogError "frmCaseEventLog", "LoadCaseDetailChanges", intEL, strES, sql


End Sub

Private Sub LoadDiscrepancyChanges()

    Dim tb As Recordset
    Dim sql As String
    Dim LatestValue(1 To 7) As String
    Dim dbName(1 To 7) As String
    Dim ShowName(1 To 7) As String
    Dim n As Integer
    Dim X As Integer
    Dim s As String


10  On Error GoTo LoadDiscrepancyChanges_Error

20  For n = 1 To 7
30      dbName(n) = Choose(n, "CaseId", "DateOfDiscrepancy", "DiscrepancyType", _
                           "NatureOfDiscrepancy", "PersonResponsible", "PersonDealingWith", _
                           "Resolution")

40      ShowName(n) = Choose(n, "CaseId", "Date Of Discrepancy", "Discrepancy Type", _
                             "Nature Of Discrepancy", "Person Responsible For Discrepancy", "Person Dealing With Discrepancy", _
                             "Discrepancy Resolution")
50  Next

60  sql = "SELECT " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(DateOfDiscrepancy)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfDiscrepancy, 103) + ' ' + CONVERT(nvarchar(50), DateOfDiscrepancy, 108), '<BLANK>') END DateOfDiscrepancy, " & _
          "CASE LTRIM(RTRIM(DiscrepancyType)) WHEN '' THEN '<BLANK>' ELSE ISNULL(DiscrepancyType, '<BLANK>') END DiscrepancyType, " & _
          "CASE LTRIM(RTRIM(NatureOfDiscrepancy)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NatureOfDiscrepancy, '<BLANK>') END NatureOfDiscrepancy, " & _
          "CASE LTRIM(RTRIM(PersonResponsible)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PersonResponsible, '<BLANK>') END PersonResponsible, " & _
          "CASE LTRIM(RTRIM(PersonDealingWith)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PersonDealingWith, '<BLANK>') END PersonDealingWith, " & _
          "CASE LTRIM(RTRIM(Resolution)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Resolution, '<BLANK>') END Resolution "

70  sql = sql & _
          "FROM Discrepancy WHERE " & _
          "CaseID = N'" & SID & "'"

80  Set tb = New Recordset
90  RecOpenClient 0, tb, sql
100 If Not tb.EOF Then
110     For n = 1 To 7
120         LatestValue(n) = tb(dbName(n)) & ""
130     Next
140 Else
150     For n = 1 To 7
160         LatestValue(n) = "<BLANK>"
170     Next
180 End If

190 sql = "SELECT ArchiveDateTime, " & _
          "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(DateOfDiscrepancy)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfDiscrepancy, 103) + ' ' + CONVERT(nvarchar(50), DateOfDiscrepancy, 108), '<BLANK>') END DateOfDiscrepancy, " & _
          "CASE LTRIM(RTRIM(DiscrepancyType)) WHEN '' THEN '<BLANK>' ELSE ISNULL(DiscrepancyType, '<BLANK>') END DiscrepancyType, " & _
          "CASE LTRIM(RTRIM(NatureOfDiscrepancy)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NatureOfDiscrepancy, '<BLANK>') END NatureOfDiscrepancy, " & _
          "CASE LTRIM(RTRIM(PersonResponsible)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PersonResponsible, '<BLANK>') END PersonResponsible, " & _
          "CASE LTRIM(RTRIM(PersonDealingWith)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PersonDealingWith, '<BLANK>') END PersonDealingWith, " & _
          "CASE LTRIM(RTRIM(Resolution)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Resolution, '<BLANK>') END Resolution "

200 sql = sql & _
          "FROM DiscrepancyAudit WHERE " & _
          "CaseID = N'" & SID & "' " & _
          "ORDER BY ArchiveDateTime DESC"

210 Set tb = New Recordset
220 RecOpenClient 0, tb, sql

230 If Not tb.EOF Then
240     pbProgress.Max = tb.RecordCount + 1
250     grdReport.Visible = False
260     fraProgress.Visible = True
270 End If

280 Do While Not tb.EOF
290     pbProgress.Value = pbProgress.Value + 1
300     lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
310     lblProgress.Refresh

320     For X = 1 To 7
330         If LatestValue(X) <> tb(dbName(X)) & "" Then
340             s = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss") & vbTab & _
                    ShowName(X) & " changed from " & tb.Fields(dbName(X)) & " to " & LatestValue(X) & vbTab & _
                    tb!ArchivedBy
350             grdReport.AddItem s
360         End If
370     Next
380     For X = 1 To 7
390         LatestValue(X) = tb(dbName(X))
400     Next

410     tb.MoveNext
420 Loop
430 fraProgress.Visible = False
440 pbProgress.Value = 1

450 If grdReport.Rows > 2 And grdReport.TextMatrix(1, 0) = "" Then
460     grdReport.RemoveItem 1
470 End If




480 Exit Sub

LoadDiscrepancyChanges_Error:
490 grdReport.Visible = True
500 fraProgress.Visible = False
    Dim strES As String
    Dim intEL As Integer

510 intEL = Erl
520 strES = Err.Description
530 LogError "frmCaseEventLog", "LoadDiscrepancyChanges", intEL, strES, sql


End Sub

Private Sub LoadChangesOtherCodes()
    Dim tb As Recordset
    Dim sql As String
    Dim s As String



10  On Error GoTo LoadChangesOtherCodes_Error

20  sql = "SELECT ArchiveDateTime, " & _
          "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
          "CASE LTRIM(RTRIM(ListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ListId, '<BLANK>') END ListId, " & _
          "CASE LTRIM(RTRIM(CaseListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseListId, '<BLANK>') END CaseListId, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(Type)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Type, '<BLANK>') END Type, " & _
          "CASE LTRIM(RTRIM(TissueTypeId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(TissueTypeId, '<BLANK>') END TissueTypeId "

30  sql = sql & _
          "FROM CaseListLinkAudit WHERE " & _
          "CaseID = N'" & SID & "' AND CaseListId <> '' " & _
          "ORDER BY ArchiveDateTime DESC"

40  Set tb = New Recordset
50  RecOpenClient 0, tb, sql

60  If Not tb.EOF Then
70      pbProgress.Max = tb.RecordCount + 1
80      grdReport.Visible = False
90      fraProgress.Visible = True
100 End If

110 Do While Not tb.EOF
120     pbProgress.Value = pbProgress.Value + 1
130     lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
140     lblProgress.Refresh

150     s = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss") & vbTab & _
            ListIdToCode(tb.Fields("ListId")) & " was deleted "
160     If tb.Fields("Type") & "" = "M" Then
170         s = s & "from " & TissueIdToCode(tb.Fields("TissueTypeId"))
180     End If
190     s = s & vbTab & tb!ArchivedBy
200     grdReport.AddItem s
        '
210     tb.MoveNext
220 Loop
230 fraProgress.Visible = False
240 pbProgress.Value = 1

250 Exit Sub

LoadChangesOtherCodes_Error:
260 grdReport.Visible = True
270 fraProgress.Visible = False
    Dim strES As String
    Dim intEL As Integer

280 intEL = Erl
290 strES = Err.Description
300 LogError "frmCaseEventLog", "LoadChangesOtherCodes", intEL, strES, sql


End Sub

Private Sub cmdSearch_Click()
10  SID = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
20  RunReport
End Sub

Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub lblAll_Click()
10  RunReport
End Sub

Private Sub lblChanges_Click()
10  ClearGrid
20  LoadDemographicChanges
30  LoadCaseDetailChanges
40  LoadChangesOtherCodes
50  LoadDiscrepancyChanges

60  grdReport.Visible = True
    'Sort by datetime
70  grdReport.Col = 0
80  grdReport.Sort = 9
End Sub

Private Sub lblEvents_Click()
10  ClearGrid
20  LoadEventLog

30  grdReport.Visible = True
    'Sort by datetime
40  grdReport.Col = 0
50  grdReport.Sort = 9
End Sub

Private Sub txtCaseId_KeyPress(KeyAscii As Integer)
    Dim lngSel As Long, lngLen As Long

10  If UCase(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
20      Call ValidateTullCaseId(KeyAscii, Me)
30  Else
40      Call ValidateLimCaseId(KeyAscii, Me)
50  End If
End Sub

Private Sub InitializeGrid()

10  With grdReport
20      .Rows = 2: .FixedRows = 1
30      .Cols = 3: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName

70      .ScrollBars = flexScrollBarBoth

80      .TextMatrix(0, 0) = "Date": .ColWidth(0) = 1900: .ColAlignment(0) = flexAlignLeftCenter
90      .TextMatrix(0, 1) = "Description": .ColWidth(1) = 10000: .ColAlignment(1) = flexAlignLeftCenter
100     .TextMatrix(0, 2) = "Logged In User": .ColWidth(2) = 2000: .ColAlignment(2) = flexAlignLeftCenter

110 End With
End Sub

Private Function TissueIdToCode(ListId As String) As String
    Dim tb As New Recordset
    Dim sql As String



10  On Error GoTo TissueIdToCode_Error

20  TissueIdToCode = "???"

30  sql = "SELECT LocationName FROM CaseTree WHERE LocationID = '" & ListId & "'"
40  Set tb = New Recordset
50  RecOpenClient 0, tb, sql
60  If Not tb.EOF Then
70      TissueIdToCode = tb!LocationName & ""
80  Else
90      TissueIdToCode = "???"
100 End If




110 Exit Function

TissueIdToCode_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmCaseEventLog", "TissueIdToCode", intEL, strES, sql


End Function


Private Function ListIdToCode(ListId As String) As String
    Dim tb As New Recordset
    Dim sql As String


10  On Error GoTo ListIdToCode_Error

20  ListIdToCode = "???"

30  sql = "SELECT Code, Description FROM Lists WHERE ListId = '" & ListId & "'"
40  Set tb = New Recordset
50  RecOpenClient 0, tb, sql
60  If Not tb.EOF Then
70      ListIdToCode = tb!Code & " - " & tb!Description & ""
80  Else
90      ListIdToCode = "???"
100 End If




110 Exit Function

ListIdToCode_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmCaseEventLog", "ListIdToCode", intEL, strES, sql


End Function

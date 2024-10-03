VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   9060
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   15240
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAudit 
      Caption         =   "&Audit"
      Height          =   855
      Left            =   3240
      Picture         =   "frmSearch.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame fraForename 
      Caption         =   "Forename"
      Height          =   615
      Left            =   2160
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox txtForename 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraSurname 
      Caption         =   "Surname"
      Height          =   615
      Left            =   240
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton bcopy 
      Appearance      =   0  'Flat
      Caption         =   "Copy to &Edit"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2160
      Picture         =   "frmSearch.frx":0F2D
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   5040
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
   Begin VB.Frame Frame4 
      Caption         =   "Search For"
      Height          =   1620
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   5715
      Begin VB.OptionButton optFor 
         Caption         =   "Surname + Forename"
         Height          =   435
         Index           =   5
         Left            =   1320
         TabIndex        =   20
         Top             =   900
         Width           =   1875
      End
      Begin VB.OptionButton optFor 
         Caption         =   "Case Id"
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   1755
      End
      Begin VB.OptionButton optFor 
         Caption         =   "Surname"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optFor 
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1275
      End
      Begin VB.OptionButton optFor 
         Caption         =   "DoB"
         Height          =   555
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optFor 
         Caption         =   "ForeName"
         Height          =   195
         Index           =   4
         Left            =   1320
         TabIndex        =   11
         Top             =   300
         Width           =   1755
      End
      Begin VB.Frame fraSearch 
         Caption         =   "How"
         Height          =   1620
         Left            =   3240
         TabIndex        =   7
         Top             =   0
         Width           =   2470
         Begin VB.CheckBox chkType 
            Caption         =   "As You Type"
            Height          =   225
            Left            =   120
            TabIndex        =   10
            Top             =   900
            Width           =   2175
         End
         Begin VB.OptionButton optExact 
            Caption         =   "Exact Match"
            Height          =   195
            Left            =   90
            TabIndex        =   9
            Top             =   300
            Width           =   2160
         End
         Begin VB.OptionButton optLeading 
            Caption         =   "Leading Characters"
            Height          =   195
            Left            =   90
            TabIndex        =   8
            Top             =   600
            Value           =   -1  'True
            Width           =   2160
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   1200
      Picture         =   "frmSearch.frx":1237
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   855
      Left            =   240
      Picture         =   "frmSearch.frx":1579
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   915
   End
   Begin VB.Frame fraText 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3585
      Begin VB.TextBox txtCaseId 
         Height          =   285
         Left            =   90
         MaxLength       =   20
         TabIndex        =   0
         Top             =   150
         Width           =   3375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6885
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   12144
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   16711680
      ForeColorSel    =   65280
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   4140
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
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
      TabIndex        =   26
      Top             =   8520
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Label lNoPrevious 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Previous Details"
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   9960
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit
Private NoPrevious As Boolean
Private mFromEdit As Boolean


Public Property Let FromEdit(ByVal X As Boolean)

10  On Error GoTo FromEdit_Error

20  mFromEdit = X



30  Exit Property

FromEdit_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmSearch", "FromEdit", intEL, strES


End Property

Private Sub InitializeGrid()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 21: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

        '120     .TextMatrix(0, 0) = "Status": .ColWidth(0) = 550: .ColAlignment(0) = flexAlignLeftCenter
        '130     .TextMatrix(0, 1) = "A": .ColWidth(1) = 200: .ColAlignment(1) = flexAlignLeftCenter
        '140     .TextMatrix(0, 2) = "H": .ColWidth(2) = 200: .ColAlignment(2) = flexAlignLeftCenter
        '150     .TextMatrix(0, 3) = "C": .ColWidth(3) = 200: .ColAlignment(3) = flexAlignLeftCenter
        '160     .TextMatrix(0, 4) = "Demographic Date": .ColWidth(4) = 1700: .ColAlignment(4) = flexAlignLeftCenter
        '170     .TextMatrix(0, 5) = "Case ID": .ColWidth(5) = 1150: .ColAlignment(5) = flexAlignLeftCenter
        '180     .TextMatrix(0, 6) = "Chart No": .ColWidth(6) = 1000: .ColAlignment(6) = flexAlignLeftCenter
        '190     .TextMatrix(0, 7) = "Name": .ColWidth(7) = 2000: .ColAlignment(7) = flexAlignLeftCenter
        '200     .TextMatrix(0, 8) = "DOB": .ColWidth(8) = 1000: .ColAlignment(8) = flexAlignLeftCenter
        '210     .TextMatrix(0, 9) = "Sex": .ColWidth(9) = 400: .ColAlignment(9) = flexAlignLeftCenter
        '220     .TextMatrix(0, 10) = "Address" & " 1": .ColWidth(10) = 1000: .ColAlignment(10) = flexAlignLeftCenter
        '230     .TextMatrix(0, 11) = "Address" & " 2": .ColWidth(11) = 1000: .ColAlignment(11) = flexAlignLeftCenter
        '240     .TextMatrix(0, 12) = "Ward": .ColWidth(12) = 1000: .ColAlignment(12) = flexAlignLeftCenter
        '250     .TextMatrix(0, 13) = "Clinician": .ColWidth(13) = 900: .ColAlignment(13) = flexAlignLeftCenter
        '260     .TextMatrix(0, 14) = "GP": .ColWidth(14) = 1000: .ColAlignment(14) = flexAlignLeftCenter
        '270     .TextMatrix(0, 15) = "Hospital": .ColWidth(15) = 1200: .ColAlignment(15) = flexAlignLeftCenter
        '280     .TextMatrix(0, 16) = "Address" & " 3": .ColWidth(16) = 0: .ColAlignment(16) = flexAlignLeftCenter
        '290     .TextMatrix(0, 17) = "Region": .ColWidth(17) = 0: .ColAlignment(17) = flexAlignLeftCenter
        '300     .TextMatrix(0, 18) = "ForeName": .ColWidth(18) = 0: .ColAlignment(18) = flexAlignLeftCenter
        '310     .TextMatrix(0, 19) = "Surname": .ColWidth(19) = 0: .ColAlignment(19) = flexAlignLeftCenter
320     .TextMatrix(0, 0) = "Dart Viewer": .ColWidth(0) = 1000: .ColAlignment(0) = flexAlignLeftCenter
120     .TextMatrix(0, 1) = "Status": .ColWidth(1) = 400: .ColAlignment(1) = flexAlignLeftCenter
130     .TextMatrix(0, 2) = "A": .ColWidth(2) = 200: .ColAlignment(2) = flexAlignLeftCenter
140     .TextMatrix(0, 3) = "H": .ColWidth(3) = 200: .ColAlignment(3) = flexAlignLeftCenter
150     .TextMatrix(0, 4) = "C": .ColWidth(4) = 200: .ColAlignment(4) = flexAlignLeftCenter
160     .TextMatrix(0, 5) = "Demographic Date": .ColWidth(5) = 1700: .ColAlignment(5) = flexAlignLeftCenter
170     .TextMatrix(0, 6) = "Case ID": .ColWidth(6) = 1150: .ColAlignment(6) = flexAlignLeftCenter
180     .TextMatrix(0, 7) = "Chart No": .ColWidth(7) = 1000: .ColAlignment(7) = flexAlignLeftCenter
190     .TextMatrix(0, 8) = "Name": .ColWidth(8) = 2000: .ColAlignment(8) = flexAlignLeftCenter
200     .TextMatrix(0, 9) = "DOB": .ColWidth(9) = 1000: .ColAlignment(9) = flexAlignLeftCenter
210     .TextMatrix(0, 10) = "Sex": .ColWidth(10) = 400: .ColAlignment(10) = flexAlignLeftCenter
220     .TextMatrix(0, 11) = "Address" & " 1": .ColWidth(11) = 1000: .ColAlignment(11) = flexAlignLeftCenter
230     .TextMatrix(0, 12) = "Address" & " 2": .ColWidth(12) = 1000: .ColAlignment(12) = flexAlignLeftCenter
240     .TextMatrix(0, 13) = "Ward": .ColWidth(13) = 1000: .ColAlignment(13) = flexAlignLeftCenter
250     .TextMatrix(0, 14) = "Clinician": .ColWidth(14) = 900: .ColAlignment(14) = flexAlignLeftCenter
260     .TextMatrix(0, 15) = "GP": .ColWidth(15) = 1000: .ColAlignment(15) = flexAlignLeftCenter
270     .TextMatrix(0, 16) = "Hospital": .ColWidth(16) = 1200: .ColAlignment(16) = flexAlignLeftCenter
280     .TextMatrix(0, 17) = "Address" & " 3": .ColWidth(17) = 0: .ColAlignment(17) = flexAlignLeftCenter
290     .TextMatrix(0, 18) = "Region": .ColWidth(18) = 0: .ColAlignment(18) = flexAlignLeftCenter
300     .TextMatrix(0, 19) = "ForeName": .ColWidth(19) = 0: .ColAlignment(19) = flexAlignLeftCenter
310     .TextMatrix(0, 20) = "Surname": .ColWidth(20) = 0: .ColAlignment(20) = flexAlignLeftCenter
        'ALI


        'ali

330 End With
End Sub

Private Sub bcopy_Click()

    Dim gRow As Long
    Dim strSex As String
    Dim strName As String

10  gRow = g.row


20  With frmDemographics
30      .txtChartNo = g.TextMatrix(gRow, 7)
40      .txtFirstName = g.TextMatrix(gRow, 19)
50      .txtSurname = g.TextMatrix(gRow, 20)
60      .txtDOB = g.TextMatrix(gRow, 9)
70      .txtAge = CalcAge(.txtDOB, Now)
80      strSex = g.TextMatrix(gRow, 10)
90      If strSex = "" Then
100         NameLostFocus strName, strSex
110     End If
120     If strSex = "M" Then
130         .txtSex = "Male"
140     ElseIf strSex = "F" Then
150         .txtSex = "Female"
160     End If
170     .txtAddress1 = initial2upper(g.TextMatrix(gRow, 11))
180     .txtAddress2 = initial2upper(g.TextMatrix(gRow, 12))
190     .txtAddress3 = initial2upper(g.TextMatrix(gRow, 17))
200     .txtCounty.MaxLength = 0
210     .txtCounty = initial2upper(g.TextMatrix(gRow, 18))

220 End With

230 Unload Me
End Sub



Private Sub cmdAudit_Click()

10  On Error GoTo cmdAudit_Click_Error

20  If g.row > 0 Then
30      With frmCaseEventLog
40          .txtCaseId = g.TextMatrix(g.row, 5)
50          .SID = Replace(g.TextMatrix(g.row, 5), " " & sysOptCaseIdSeperator(0) & " ", "")
60          .RunReport
70          .Show 1
80      End With
90  End If

100 Exit Sub

cmdAudit_Click_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmSearch", "cmdAudit_Click", intEL, strES

End Sub

Private Sub cmdExit_Click()
10  Unload Me
End Sub

Private Sub cmdSearch_Click()
10  FillG
End Sub



Private Sub Form_Load()

    ChangeFont Me, "Arial"

    'frmSearch_ChangeLanguage
    InitializeGrid
    lblLoggedIn = UserName

    Label9.Caption = Label9.Caption & UserName

    If UCase$(UserMemberOf) <> "LOOKUP" Then
        cmdAudit.Visible = True
    End If

End Sub

Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub FillG()

10  lNoPrevious.Visible = False

20  If optFor(5) Then
30      If Trim$(txtSurname) = "" And Trim$(txtForename) = "" Then
40          frmMsgBox.Msg "No Criteria Entered", mbOKOnly, "Histology", mbExclamation
50          If TimedOut Then Unload Me: Exit Sub
60          Exit Sub
70      End If
80  Else
90      If Trim$(txtCaseId) = "" Then
100         frmMsgBox.Msg "No Criteria Entered", mbOKOnly, "Histology", mbExclamation
110         If TimedOut Then Unload Me: Exit Sub
120         Exit Sub
130     End If
140 End If

150 ClearFGrid g

160 LocalFillG

170 With g
180     If .Rows > 2 Then
190         .RemoveItem 1
200         .row = 1
210         .col = 3
220         .ColSel = .Cols - 1
230         .RowSel = 1
240         .HighLight = flexHighlightAlways
250     End If
260 End With
270 bcopy.Enabled = mFromEdit

280 g.Visible = True

290 Screen.MousePointer = vbDefault

End Sub

Private Sub LocalFillG()

    Dim s As String
    Dim tb As New Recordset
    Dim sqlBase As String
    Dim Criteria As String
    Dim TempCaseId As String

10  On Error GoTo LocalFillG_Error

20  Criteria = ""

30  If InStr(txtCaseId, "%") > 0 Or InStr(txtCaseId, "_") > 0 Or txtCaseId = "'" Then
40      frmMsgBox.Msg "Invalid Search Criteria", mbOKOnly, "Histology", mbExclamation
50      If TimedOut Then Unload Me: Exit Sub
60      Exit Sub
70  End If

80  If optFor(0) Then
90      If optExact Then
100         Criteria = "D.Surname = N'" & AddTicks(txtCaseId) & "' "
110     ElseIf optLeading Then
120         Criteria = "D.Surname like N'" & AddTicks(txtCaseId) & "%' "
130     Else
140         Criteria = "D.Surname like N'%" & AddTicks(txtCaseId) & "' "
150     End If
160 ElseIf optFor(1) Then
170     Criteria = "D.MRN = N'" & UCase(AddTicks(txtCaseId)) & "' "
180 ElseIf optFor(2) Then
190     txtCaseId = Convert62Date(txtCaseId, BACKWARD)
200     If Not IsDate(txtCaseId) Then
210         Screen.MousePointer = vbDefault
220         iMsg "Invalid Date", vbExclamation, "Date of Birth Search"
230         If TimedOut Then Unload Me: Exit Sub
240         Exit Sub
250     End If
260     Criteria = "D.DateOfBirth = '" & Format$(txtCaseId, "yyyymmdd") & "'"
270 ElseIf optFor(3) Then
280     TempCaseId = Replace(UCase(AddTicks(txtCaseId)), sysOptCaseIdSeperator(0), "")
290     TempCaseId = Replace(TempCaseId, " ", "")
300     Criteria = "D.CaseId = N'" & Trim(TempCaseId) & "' "
310 ElseIf optFor(4) Then
320     If optExact Then
330         Criteria = "D.FirstName = N'" & AddTicks(txtCaseId) & "' "
340     ElseIf optLeading Then
350         Criteria = "D.FirstName like N'" & AddTicks(txtCaseId) & "%' "
360     Else
370         Criteria = "D.FirstName like N'%" & AddTicks(txtCaseId) & "' "
380     End If
390 ElseIf optFor(5) Then
400     If optExact Then
410         Criteria = "D.Surname = N'" & AddTicks(txtSurname) & "' " & _
                       "AND D.FirstName = N'" & AddTicks(txtForename) & "' "
420     ElseIf optLeading Then
430         Criteria = "D.Surname like N'" & AddTicks(txtSurname) & "%' " & _
                       "AND D.FirstName like N'" & AddTicks(txtForename) & "%' "
440     Else
450         Criteria = "D.Surname like N'%" & AddTicks(txtCaseId) & "' " & _
                       "AND D.FirstName like N'%" & AddTicks(txtForename) & "' "
460     End If
470 End If

480 sqlBase = "SELECT DISTINCT D.CaseId, " & _
              "ISNULL(D.[Year], YEAR(D.DateTimeOfRecord)) [Year], " & _
              "D.PatientName, D.SurName, D.FirstName, D.MRN, " & _
              "D.Clinician, " & _
              "D.Ward, D.Address1, D.Address2, D.Address3, D.County, " & _
              "D.Sex, D.DateOfBirth, D.DateTimeOfRecord, D.Source,D.GP, C.State, C.WithPathologist " & _
              "FROM Demographics AS D " & _
              "INNER JOIN Cases AS C " & _
              "ON D.CaseId = C.CaseId " & _
              "WHERE " & _
              Criteria & " "

490 If UCase(UserMemberOf) = "LOOKUP" Then
500     sqlBase = sqlBase & "AND SUBSTRING(D.CaseId,2,1) <> 'A' "
510 End If

520 sqlBase = sqlBase & "ORDER BY D.DateTimeOfRecord DESC, D.CaseId DESC"

530 NoPrevious = True

540 Set tb = New Recordset
550 RecOpenClient 0, tb, sqlBase
560 With tb
570     If Not .EOF Then
580         NoPrevious = False
590         pbProgress.Max = .RecordCount + 1
600         g.Visible = False
610         fraProgress.Visible = True
620     End If

630     Do While Not .EOF

640         pbProgress.Value = pbProgress.Value + 1
650         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
660         lblProgress.Refresh
670         If Mid(!CaseId & "", 2, 1) = "P" Or Mid(!CaseId & "", 2, 1) = "A" Then
680             TempCaseId = Left(!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(!CaseId, 2)
690         Else
700             TempCaseId = Left(!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(!CaseId, 2)
710         End If
720         s = vbTab & vbTab & vbTab & vbTab & vbTab
730         s = s & _
                IIf(IsNull(!DateTimeOfRecord), "", Format$(!DateTimeOfRecord, "dd/mm/yy")) & vbTab & _
                IIf(IsNull(Trim$(TempCaseId)), "", Trim$(TempCaseId)) & vbTab & _
                IIf(IsNull(!MRN), "", !MRN) & vbTab & _
                IIf(IsNull(!PatientName), "", IIf(!PatientName = Null, "", !PatientName)) & vbTab & _
                IIf(IsNull(!DateOfBirth), "", Format$(!DateOfBirth, "Short Date"))










740         s = s & vbTab & _
                IIf(IsNull(!Sex), "", !Sex) & vbTab & _
                IIf(IsNull(!Address1), "", !Address1) & vbTab & _
                IIf(IsNull(!Address2), "", !Address2) & vbTab & _
                IIf(IsNull(!Ward), "", !Ward) & vbTab & _
                IIf(IsNull(initial2upper(!Clinician & "")), "", initial2upper(!Clinician & "")) & vbTab & _
                IIf(IsNull(initial2upper(!GP & "")), "", initial2upper(!GP & "")) & vbTab & _
                IIf(IsNull(!Source), "", !Source)


750         s = s & vbTab & _
                IIf(IsNull(!Address3), "", !Address3) & vbTab & _
                IIf(IsNull(!County), "", !County) & vbTab & _
                IIf(IsNull(!FirstName), "", !FirstName) & vbTab & _
                IIf(IsNull(!Surname), "", !Surname)


760         g.AddItem s
770         g.row = g.Rows - 1

780         g.col = 1
790         If UCase(!State & "") = "IN HISTOLOGY" Then
800             g.CellForeColor = vbBlue
810             g.TextMatrix(g.row, 1) = "L"
820         ElseIf UCase(!State & "") = "WITH PATHOLOGIST" Then
830             g.CellForeColor = &H6495ED
840             g.TextMatrix(g.row, 1) = !WithPathologist & ""
850         ElseIf UCase(!State & "") = "AWAITING AUTHORISATION" Then
860             g.CellForeColor = vbRed
870             g.TextMatrix(g.row, 1) = "AA"
880         ElseIf UCase(!State & "") = "AUTHORISED" Then
890             g.CellForeColor = vbGreen
900             g.TextMatrix(g.row, 1) = "A"
910         End If

920         If Left(!CaseId & "", 1) = "H" Then
930             g.col = 3
940             g.CellBackColor = vbRed
950         ElseIf Left(!CaseId & "", 1) = "C" Then
960             g.col = 4
970             g.CellBackColor = vbRed
980         Else
990             g.col = 2
1000            g.CellBackColor = vbRed
1010        End If

1020        .MoveNext
1030    Loop
1040    fraProgress.Visible = False
1050    pbProgress.Value = 1
1060 End With

1070 If NoPrevious Then
1080    lNoPrevious.Visible = True
1090 End If
    '     g.Cols = g.Cols + 1
    '     g.TextMatrix(0, g.Cols - 1) = "Dart Viewer"
    '
    '     Dim i As Integer
    ''
    ''     'To change the color of dart viewer column to green
    ''     'Zyam
    '     For i = 1 To g.Rows - 1
    '        With g
    '            .row = i
    '            .col = g.Cols - 1
    '            .CellBackColor = vbGreen
    '        End With
    '     Next
    ''     'Zyam


    'ali
    Dim i As Integer
    For i = 1 To g.Rows - 1
        With g
            .row = i
            .col = 0
            .CellBackColor = vbGreen
        End With
    Next
    '------------
1100 Exit Sub

LocalFillG_Error:
1110 g.Visible = True
1120 fraProgress.Visible = False
    Dim strES As String
    Dim intEL As Integer

1130 intEL = Erl
1140 strES = Err.Description
1150 LogError "frmSearch", "LocalFillG", intEL, strES


End Sub

Private Sub Form_Unload(Cancel As Integer)
10  Unload frmWorkSheet
End Sub

Private Sub g_Click()

10  If g.col > 4 And mFromEdit Then
20      g.col = 0
30      g.ColSel = g.Cols - 1
40      g.RowSel = g.row
50      g.HighLight = flexHighlightAlways
60      bcopy.Enabled = True
70  ElseIf g.col = 2 Or g.col = 3 Or g.col = 4 Then
80      If g.CellBackColor = vbRed Then
90          CaseNo = Replace(g.TextMatrix(g.row, 6), " " & sysOptCaseIdSeperator(0) & " ", "")
100         If UCase$(UserMemberOf) = "LOOKUP" Then
110             If g.TextMatrix(g.row, 1) = "A" Then
120                 PrintHistology "", True
130                 With frmRichText
140                     .cmdPrint.Visible = False
150                     .cmdExit.Left = 0
160                     .rtb.SelStart = 0
170                     .Show 1
180                 End With
190             End If
200         Else
210             PrintHistology "", True
220             With frmRichText
230                 If g.TextMatrix(g.row, 1) = "A" Then
240                     .cmdPrint.Enabled = True
250                 Else
260                     .cmdPrint.Enabled = False
270                 End If
280                 .rtb.SelStart = 0
290                 .Show 1
300             End With
310         End If
320     End If
330 ElseIf g.col = 1 Then
340     If g.TextMatrix(g.row, 1) = "L" Then
350         With frmPhase
360             .Search = True
370             .Show 1
380         End With
390     End If
400 ElseIf g.MouseCol = 0 Then
        Dim CaseId As String
        Dim formatedCaseID As String
        
410     CaseId = g.TextMatrix(g.MouseRow, 6)
420     CaseId = Trim(CaseId)
430     formatedCaseID = Replace(CaseId, " ", "")
    
        DoEvents
        DoEvents
        Sleep (55)
450     If Dir("\\tdws08fs01.mhb.health.gov.ie\MRHP_WardEnquiry\MRHT\Netaquire\The Plumtree Group\DartViewer\DartViewer.exe") = "" Then
460         iMsg "Dart client not installed on this machine. Please contact you system administrator", vbInformation
470         Exit Sub
480     End If
        DoEvents
        DoEvents
        Sleep (55)
490     Shell "\\tdws08fs01.mhb.health.gov.ie\MRHP_WardEnquiry\MRHT\Netaquire\The Plumtree Group\DartViewer\DartViewer.exe " & formatedCaseID, vbNormalFocus
        DoEvents
        DoEvents
        Sleep (55)
        'Zyam
500 End If

End Sub



'Private Sub g_DblClick()
' 'Zyam
'10       On Error GoTo g_DblClick_Error
'
'20      If g.MouseCol = g.Cols Then
'            Dim CaseId As String
'30          CaseId = g.TextMatrix(g.MouseRow, 5)
'40          CaseId = Trim(CaseId)
'50          CaseId = Replace(CaseId, " ", "")
'60          CaseId = Replace(CaseId, "/", "")
'            'MsgBox (CaseId)
'70       If Dir("\\tdws08fs01.mhb.health.gov.ie\MRHP_WardEnquiry\MRHT\Netaquire\The Plumtree Group\DartViewer\DartViewer.exe") = "" Then
'80          iMsg "Dart client not installed on this machine. Please contact you system administrator", vbInformation
'90          Exit Sub
'100      End If
'110      Shell "\\tdws08fs01.mhb.health.gov.ie\MRHP_WardEnquiry\MRHT\Netaquire\The Plumtree Group\DartViewer\DartViewer.exe " & CaseId, vbNormalFocus
'120     End If
'    'Zyam
'g_DblClick_Error:
'       Dim strES As String
'       Dim intEL As Long
'
'130    intEL = Erl
'140    strES = Err.Description
'150    LogError "frmSearch", "g_DblClick", intEL, strES
'
'
'End Sub

Private Sub optFor_Click(Index As Integer)
10  If Index = 5 Then
20      fraText.Visible = False
30      fraSurname.Visible = True
40      fraForename.Visible = True
50      txtCaseId = ""
60      txtSurname.SetFocus
70  Else
80      txtSurname = ""
90      txtForename = ""
100     fraText.Visible = True
110     fraSurname.Visible = False
120     fraForename.Visible = False
130     txtCaseId.SetFocus
140 End If

End Sub

Private Sub txtCaseId_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim SearchForFound As Boolean



10  On Error GoTo txtCaseId_KeyPress_Error
    'Zyam 1-08-24
    If optFor(3).Value = True Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    'Zyam 1-08-24

20  If optFor(3) Then
30      If UCase(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
40          Call ValidateTullCaseId(KeyAscii, Me)
50      Else
60          Call ValidateLimCaseId(KeyAscii, Me)
70      End If
80  End If

90  For i = 0 To 4
100     If optFor(i) Then
110         SearchForFound = True
120         Exit For
130     End If
140 Next

150 If Not SearchForFound Then
160     frmMsgBox.Msg "Pleae select what you are searching for", mbOKOnly, "Histology", mbInformation
170     KeyAscii = 0
180 End If



txtCaseId_KeyPress_Error:
    Dim strES As String
    Dim intEL As Long

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmSearch", "txtCaseId_KeypPress", intEL, strES



End Sub




VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmNCRI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAquire - Cellular Pathology"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13785
   Icon            =   "frmNCRI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4800
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   12360
      Picture         =   "frmNCRI.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   975
      Left            =   11160
      Picture         =   "frmNCRI.frx":120C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1035
   End
   Begin Threed.SSPanel panSampleDates 
      Height          =   1665
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   2937
      _StockProps     =   15
      Caption         =   "Between Sample Dates"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin VB.CheckBox chkAuthSince 
         Caption         =   "Show Only Cases between above dates that were authorised since"
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   1200
         Width           =   5295
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "&Calculate"
         Default         =   -1  'True
         Height          =   930
         Left            =   6840
         Picture         =   "frmNCRI.frx":1356
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Search"
         Top             =   120
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker calfrom 
         Height          =   345
         Left            =   210
         TabIndex        =   2
         Top             =   600
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110100481
         CurrentDate     =   37753
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   345
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110100481
         CurrentDate     =   37753
      End
      Begin MSComCtl2.DTPicker dtAuthorisedSince 
         Height          =   345
         Left            =   5520
         TabIndex        =   16
         Top             =   1200
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
         Format          =   110100483
         CurrentDate     =   37753
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2745
         TabIndex        =   5
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   4
         Top             =   330
         Width           =   945
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdNCRI 
      Height          =   5775
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   10186
      _Version        =   393216
      RowHeightMin    =   300
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
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
      TabIndex        =   14
      Top             =   8040
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Image imgReport 
      Height          =   240
      Left            =   8520
      Picture         =   "frmNCRI.frx":173E
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmNCRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InitializeGrid()

10  With grdNCRI
20      .Rows = 2: .FixedRows = 1
30      .Cols = 11: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "": .ColWidth(0) = 400: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Case Id": .ColWidth(1) = 1050: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "Name": .ColWidth(2) = 2000: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Date Of Birth": .ColWidth(3) = 1500: .ColAlignment(3) = flexAlignLeftCenter
160     .TextMatrix(0, 4) = "Sex": .ColWidth(4) = 500: .ColAlignment(4) = flexAlignLeftCenter
170     .TextMatrix(0, 5) = "Address": .ColWidth(5) = 2000: .ColAlignment(5) = flexAlignLeftCenter
180     .TextMatrix(0, 6) = "Source": .ColWidth(6) = 1000: .ColAlignment(6) = flexAlignLeftCenter
190     .TextMatrix(0, 7) = "Chart": .ColWidth(7) = 800: .ColAlignment(7) = flexAlignLeftCenter
200     .TextMatrix(0, 8) = "GP": .ColWidth(8) = 1000: .ColAlignment(8) = flexAlignLeftCenter
210     .TextMatrix(0, 9) = "Clinican": .ColWidth(9) = 1000: .ColAlignment(9) = flexAlignLeftCenter
220     .TextMatrix(0, 10) = "Date Of Sample": .ColWidth(10) = 1500: .ColAlignment(10) = flexAlignLeftCenter

230 End With
End Sub



Private Sub chkAuthSince_Click()
10  If chkAuthSince = 1 Then
20      dtAuthorisedSince.Enabled = True
30      dtAuthorisedSince.CustomFormat = "dd/MM/yyyy"
40      dtAuthorisedSince = GetOptionSetting("NCRILastViewed", Format(Now, "dd/mm/yyyy"))
50  Else
60      dtAuthorisedSince.Enabled = False
70      dtAuthorisedSince.CustomFormat = " "
80  End If
End Sub

Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub cmdCalc_Click()
    Dim tb As Recordset
    Dim sql As String
    Dim s As String
    Dim TempCaseId As String
    Dim PrevCaseId As String
    Dim PrevTissueType As String

10  On Error GoTo cmdCalc_Click_Error

20  With grdNCRI
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60  End With

70  sql = "SELECT * FROM CaseListLink cl " & _
          "INNER JOIN Lists l on l.ListId = cl.ListId " & _
          "INNER JOIN Demographics d ON d.CaseId = cl.CaseId " & _
          "INNER JOIN Cases c ON d.CaseId = c.CaseId " & _
          "WHERE cl.Type = 'M' and l.Cancerous = 1 AND " & _
          "c.State = 'Authorised' AND " & _
          "c.SampleTaken between '" & Format(calFrom, "dd/MMM/yyyy") & " 00:00:00'" & _
        " AND '" & Format(calTo, "dd/MMM/yyyy") & " 23:59:59' "

80  If chkAuthSince = 1 Then
90      sql = sql & "AND c.ValReportDate BETWEEN '" & Format(dtAuthorisedSince, "dd/MMM/yyyy") & " 00:00:00' " & _
              "AND '" & Format(Now, "dd/MMM/yyyy hh:mm:ss") & "' "
100 End If

110 sql = sql & "ORDER BY cl.CaseId,cl.TissueTypeId"

120 Set tb = New Recordset
130 RecOpenClient 0, tb, sql
140 If Not tb.EOF Then
150     pbProgress.Max = tb.RecordCount + 1
160     grdNCRI.Visible = False
170     fraProgress.Visible = True
180     Do While Not tb.EOF
190         pbProgress.Value = pbProgress.Value + 1
200         lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
210         lblProgress.Refresh

220         If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
230             TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
240         Else
250             TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
260         End If

270         If PrevCaseId <> TempCaseId Or PrevTissueType <> tb!TissueTypeId & "" Then
280             s = vbTab & TempCaseId
290             s = s & vbTab & Trim(tb!PatientName) & vbTab
300             If IsDate(tb!DateOfBirth) Then s = s & Format(tb!DateOfBirth, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
310             s = s & tb!Sex & vbTab & tb!Address1 & vbTab & tb!Source & vbTab & tb!MRN & vbTab & tb!GP & vbTab & tb!Clinician & vbTab
320             If IsDate(tb!DateTimeOfRecord) Then s = s & Format(tb!DateTimeOfRecord, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
330             grdNCRI.AddItem s

340             grdNCRI.Row = grdNCRI.Rows - 1
350             grdNCRI.Col = 0
360             grdNCRI.CellPictureAlignment = flexAlignCenterCenter
370             Set grdNCRI.CellPicture = imgReport.Picture
380         End If

390         PrevCaseId = TempCaseId
400         PrevTissueType = tb!TissueTypeId & ""
410         tb.MoveNext
420     Loop
430     fraProgress.Visible = False
440     grdNCRI.Visible = True
450     pbProgress.Value = 1
460 End If


470 If grdNCRI.Rows > 2 And grdNCRI.TextMatrix(1, 0) = "" Then grdNCRI.RemoveItem 1

480 Call SaveOptionSetting("NCRILastViewed", Format(Now, "dd/mm/yyyy"))

490 Exit Sub



cmdCalc_Click_Error:
500 grdNCRI.Visible = True
510 fraProgress.Visible = False

    Dim strES As String
    Dim intEL As Integer

520 intEL = Erl
530 strES = Err.Description
540 LogError "frmNCRI", "cmdCalc_Click", intEL, strES, sql


End Sub

Private Sub cmdExit_Click()
10  Unload Me
End Sub


Private Sub cmdExport_Click()
10  On Error GoTo cmdExport_Click_Error

20  ExportFlexGrid grdNCRI, Me

30  Exit Sub

cmdExport_Click_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmNCRI", "cmdExport_Click", intEL, strES

End Sub

Private Sub Form_Load()


10    ChangeFont Me, "Arial"
'20    frmNCRI_ChangeLanguage
30    calTo = Format(Now, "dd/MMM/yyyy")
40    calFrom = Format(Now - 7, "dd/MMM/yyyy")

50    InitializeGrid

60    lblLoggedIn = UserName
70    If blnIsTestMode Then EnableTestMode Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
10  Unload frmWorkSheet
End Sub

Private Sub grdNCRI_Click()
10  If grdNCRI.Col = 0 Then
20      CaseNo = Replace(grdNCRI.TextMatrix(grdNCRI.Row, 1), " " & sysOptCaseIdSeperator(0) & " ", "")
30      PrintHistology "", True
40      With frmRichText
50          .rtb.SelStart = 0
60          .Show 1
70      End With
80  End If
End Sub


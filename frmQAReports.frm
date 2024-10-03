VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmQAReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histology Coding System - QA Reports"
   ClientHeight    =   10335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFilter 
      Height          =   2175
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "GP"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Exit"
      Height          =   1100
      Left            =   8880
      Picture         =   "frmQAReports.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9000
      Width           =   1200
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   1100
      Left            =   8880
      Picture         =   "frmQAReports.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid grdReport 
      Height          =   7575
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Height          =   1100
         Left            =   3960
         Picture         =   "frmQAReports.frx":6AE8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   1200
      End
      Begin VB.ComboBox cmbReports 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   360
         Left            =   2040
         TabIndex        =   1
         Top             =   735
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   80936961
         CurrentDate     =   37951
      End
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   735
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   80936961
         CurrentDate     =   37951
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmQAReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadReports()
10  cmbReports.AddItem "Hospital Workload"
20  cmbReports.AddItem "Total Blocks & Slides"
30  cmbReports.AddItem "IHC & Special Stains"
40  cmbReports.AddItem "Interinstitutional Consultation"
50  cmbReports.AddItem "Intradepartmental Consultation"
60  cmbReports.AddItem "Retrospective Review"
70  cmbReports.AddItem "MDT"
80  cmbReports.AddItem "Lab Based Incidents"
90  cmbReports.AddItem "TAT"
100 cmbReports.AddItem "Addendum Reports"
110 cmbReports.AddItem "Reports to Clinician"


120 cmbReports.ListIndex = 0

End Sub
Private Sub InitGridWL()
    Dim i As Integer
10  With grdReport
20      .Rows = 2: .FixedRows = 1
30      .Cols = 2: .FixedCols = 0
40      .Rows = 1
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 0) = "Case Type": .ColWidth(0) = 5000: .ColAlignment(0) = flexAlignLeftCenter
70      .TextMatrix(0, 1) = "No. of Cases": .ColWidth(1) = 825: .ColAlignment(1) = flexAlignLeftCenter
80      For i = 0 To .Cols - 1
90          If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
100             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
110         End If
120     Next i
130 End With
End Sub
Private Sub InitGridBS()
    Dim i As Integer
10  With grdReport
20      .Rows = 2: .FixedRows = 1
30      .Cols = 2: .FixedCols = 0
40      .Rows = 2
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 0) = "Total No. of Blocks": .ColWidth(0) = 2000: .ColAlignment(0) = flexAlignLeftCenter
70      .TextMatrix(0, 1) = "Total No. of Slides": .ColWidth(1) = 2000: .ColAlignment(1) = flexAlignLeftCenter
80      For i = 0 To .Cols - 1
90          If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
100             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
110         End If
120     Next i
130 End With
End Sub

Private Sub InitGridStains()
    Dim i As Integer
10  With grdReport
20      .Rows = 2: .FixedRows = 1
30      .Cols = 3: .FixedCols = 0
40      .Rows = 1
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 0) = "Stain Type": .ColWidth(0) = 4000: .ColAlignment(0) = flexAlignLeftCenter
70      .TextMatrix(0, 1) = "Total No. of Stains": .ColWidth(1) = 2000: .ColAlignment(1) = flexAlignLeftCenter
80      .TextMatrix(0, 2) = "Total No. of Cases": .ColWidth(2) = 2000: .ColAlignment(2) = flexAlignLeftCenter
90      For i = 0 To .Cols - 1
100         If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
110             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
120         End If
130     Next i
140 End With
End Sub

Private Sub InitGridInter()
    Dim i As Integer
10  With grdReport
20      .Rows = 2: .FixedRows = 1
30      .Cols = 3: .FixedCols = 0
40      .Rows = 1
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 0) = "Interinstitutional Consultation Type": .ColWidth(0) = 4000: .ColAlignment(0) = flexAlignLeftCenter
70      .TextMatrix(0, 1) = "No. of Cases": .ColWidth(1) = 2000: .ColAlignment(1) = flexAlignLeftCenter
80      .TextMatrix(0, 2) = "% Agreement": .ColWidth(2) = 2000: .ColAlignment(2) = flexAlignLeftCenter
90      For i = 0 To .Cols - 1
100         If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
110             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
120         End If
130     Next i
140 End With
End Sub

Private Sub InitGridIntra()
    Dim i As Integer
10  With grdReport
20      .Rows = 2: .FixedRows = 1
30      .Cols = 2: .FixedCols = 0
40      .Rows = 1
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 0) = "Case Type": .ColWidth(0) = 4000: .ColAlignment(0) = flexAlignLeftCenter
70      .TextMatrix(0, 1) = "% of Total Cases with Intradepartmental Consultation": .ColWidth(1) = 4000: .ColAlignment(1) = flexAlignLeftCenter
80      For i = 0 To .Cols - 1
90          If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
100             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
110         End If
120     Next i
130 End With
End Sub

Private Sub InitGridTAT()
    Dim i As Integer
10  With grdReport
20      .Rows = 3: .FixedRows = 2
30      .Cols = 13: .FixedCols = 0
40      .Rows = 2
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 8) = "% completed by": .ColWidth(8) = 1000: .ColAlignment(8) = flexAlignLeftCenter
70      .TextMatrix(1, 0) = "Case Type": .ColWidth(0) = 2500: .ColAlignment(0) = flexAlignLeftCenter
80      .TextMatrix(1, 1) = "Tissue Type": .ColWidth(1) = 2000: .ColAlignment(1) = flexAlignLeftCenter
90      .TextMatrix(1, 2) = "Day 1": .ColWidth(2) = 1000: .ColAlignment(2) = flexAlignLeftCenter
100     .TextMatrix(1, 3) = "Day 2": .ColWidth(3) = 1000: .ColAlignment(3) = flexAlignLeftCenter
110     .TextMatrix(1, 4) = "Day 3": .ColWidth(4) = 1000: .ColAlignment(4) = flexAlignLeftCenter
120     .TextMatrix(1, 5) = "Day 4": .ColWidth(5) = 1000: .ColAlignment(5) = flexAlignLeftCenter
130     .TextMatrix(1, 6) = "Day 5": .ColWidth(6) = 1000: .ColAlignment(6) = flexAlignLeftCenter
140     .TextMatrix(1, 7) = "Day 6": .ColWidth(7) = 1000: .ColAlignment(7) = flexAlignLeftCenter
150     .TextMatrix(1, 8) = "Day 7": .ColWidth(8) = 1000: .ColAlignment(8) = flexAlignLeftCenter
160     .TextMatrix(1, 9) = "Day 8": .ColWidth(9) = 1000: .ColAlignment(9) = flexAlignLeftCenter
170     .TextMatrix(1, 10) = "Day 9": .ColWidth(10) = 1000: .ColAlignment(10) = flexAlignLeftCenter
180     .TextMatrix(1, 11) = "Day 10": .ColWidth(11) = 1000: .ColAlignment(11) = flexAlignLeftCenter
190     .TextMatrix(1, 12) = "> 10 Days": .ColWidth(12) = 1000: .ColAlignment(12) = flexAlignLeftCenter
200     For i = 0 To .Cols - 1
210         If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
220             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
230         End If
240     Next i
250 End With
End Sub

Private Sub InitGridAddendum()
    Dim i As Integer
10  With grdReport
20      .Rows = 2: .FixedRows = 1
30      .Cols = 2: .FixedCols = 0
40      .Rows = 1
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 0) = "Addendum Report Type": .ColWidth(0) = 3000: .ColAlignment(0) = flexAlignLeftCenter
70      .TextMatrix(0, 1) = "Qty expressed as % of total first reports": .ColWidth(1) = 4000: .ColAlignment(1) = flexAlignLeftCenter
80      For i = 0 To .Cols - 1
90          If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
100             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
110         End If
120     Next i
130 End With
End Sub
Private Sub InitGridRepToClin()
    Dim i As Integer
10  With grdReport
20      .Rows = 2: .FixedRows = 1
30      .Cols = 1: .FixedCols = 0
40      .Rows = 1
50      .ScrollBars = flexScrollBarBoth

60      .TextMatrix(0, 0) = "% of total cases reported to clinician": .ColWidth(0) = 3000: .ColAlignment(0) = flexAlignLeftCenter
70      For i = 0 To .Cols - 1
80          If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
90              .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
100         End If
110     Next i
120 End With
End Sub

Private Sub cmdCalc_Click()
10  grdReport.Clear
    'ClearFGrid grdReport
20  Select Case cmbReports
    Case "Hospital Workload"
30      InitGridWL
40      FillGridWL
50  Case "Total Blocks & Slides"
60      InitGridBS
70      FillGridBS
80  Case "IHC & Special Stains"
90      InitGridStains
100     FillGridStains
110 Case "Interinstitutional Consultation"
120     InitGridInter
130     FillGridInter
140 Case "Intradepartmental Consultation"
150     InitGridIntra
160 Case "TAT"
170     InitGridTAT
180     FillGridTAT
190 Case "Addendum Reports"
200     InitGridAddendum
210     FillGridAddendum
220 Case "Reports to Clinician"
230     InitGridRepToClin
240     FillGridRepToClin
250 Case Else
260     grdReport.Cols = 0

270 End Select
    'FixG grdReport
End Sub

Private Sub cmdExport_Click()
    Dim s As String

10  s = "ST JOHN 'S HOSPITAL,LIMERICK TEL 061-462141" & vbCr & _
        cmbReports
    'If optClinician Then
    '  s = s & "Clinicians"
    'ElseIf optWard Then
    '  s = s & "Wards"
    'Else
    '  s = s & "GP's"
    'End If
20  s = s & vbCr & _
        "Between " & calFrom & " and " & calTo & vbCr

30  ExportFlexGrid grdReport, Me, s
End Sub

Private Sub Form_Resize()
10  Me.Top = 0
20  Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub FillGridWL()

    Dim tb As New Recordset
    Dim sn As New Recordset
    Dim sql As String
    Dim s As String

10  On Error GoTo FillGridWL_Error

20  sql = "select * from lists where listtype = 'P'"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  Do While Not tb.EOF

60      sql = "SELECT count(*) as tot FROM CaseListLink WHERE " & _
              "ListId = " & tb!ListId & " and datetimeofrecord between '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
              "and '" & Format(calTo, "yyyymmdd") & " 23:59:59'"
70      Set sn = New Recordset
80      RecOpenServer 0, sn, sql
90      If Not sn.EOF Then
            'If sn!Tot > 0 Then
100         s = tb!Description & vbTab & sn!tot
110         grdReport.AddItem s
            'End If
120     End If
130     tb.MoveNext
140 Loop

150 Exit Sub

FillGridWL_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmQAReports", "FillGridWL", intEL, strES, sql

End Sub
Private Sub FillGridBS()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo FillGridBS_Error

20  sql = "select block from casedetails group by caseid,block having COALESCE(block, '') <> ''"
30  Set tb = New Recordset
40  RecOpenClient 0, tb, sql


50  If Not tb.EOF Then
60      grdReport.TextMatrix(1, 0) = tb.RecordCount
70  End If

80  sql = "select slide from casedetails group by caseid,slide having COALESCE(slide, '') <> ''"
90  Set tb = New Recordset
100 RecOpenClient 0, tb, sql

110 If Not tb.EOF Then
120     grdReport.TextMatrix(1, 1) = tb.RecordCount
130 End If


140 Exit Sub

FillGridBS_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmQAReports", "FillGridBS", intEL, strES, sql

End Sub

Private Sub FillGridStains()

    Dim tb As New Recordset
    Dim sql As String
    Dim s As String




10  On Error GoTo FillGridStains_Error

20  sql = "SELECT Count(DISTINCT CaseId) as totCases, Count(DISTINCT Stain) as totStain, l.listtype " & _
          "FROM CaseDetails cl,lists l  " & _
          "WHERE l.description = cl.stain " & _
          "AND COALESCE(cl.stain, '') <> '' " & _
          "GROUP BY l.ListType"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  Do Until tb.EOF
60      If tb!ListType & "" = "SS" Then
70          s = "Special Stains" & vbTab & tb!totStain & vbTab & tb!totCases
80          grdReport.AddItem s
90      Else
100         s = "Immunohistochemical Stains" & vbTab & tb!totStain & vbTab & tb!totCases
110         grdReport.AddItem s
120     End If
130     tb.MoveNext
140 Loop




150 Exit Sub

FillGridStains_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmQAReports", "FillGridStains", intEL, strES, sql


End Sub

Private Sub FillGridInter()

    Dim tb As New Recordset
    Dim sn As New Recordset
    Dim rs As New Recordset
    Dim sql As String
    Dim s As String
    Dim tps As String
    Dim totAgreed As Long


10  On Error GoTo FillGridInter_Error



20  sql = "select * from lists where code in ('Q001','Q002','Q003')"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  Do While Not tb.EOF
60      sql = "SELECT count(*) as tot FROM CaseListLink WHERE " & _
              "ListId = " & tb!ListId & " and datetimeofrecord between '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
              "and '" & Format(calTo, "yyyymmdd") & " 23:59:59'"
70      Set sn = New Recordset
80      RecOpenServer 0, sn, sql
90      If Not sn.EOF Then
100         sql = "SELECT count(*) as totAgreed FROM CaseMovements WHERE Agreed = '1' And Code = '" & tb!Code & "'"
110         Set rs = New Recordset
120         RecOpenServer 0, rs, sql

130         If rs!totAgreed <> 0 Then
140             tps = Format$(sn!tot / rs!totAgreed, "##.00") * 100
150         End If
160         If tb!Code = "Q003" Or sn!tot = 0 Then
170             s = tb!Description & vbTab & sn!tot & vbTab & "N/A"
180         Else
190             s = tb!Description & vbTab & sn!tot & vbTab & tps
200         End If
210         grdReport.AddItem s
220     End If
230     tb.MoveNext
240 Loop


250 Exit Sub

FillGridInter_Error:

    Dim strES As String
    Dim intEL As Integer

260 intEL = Erl
270 strES = Err.Description
280 LogError "frmQAReports", "FillGridInter", intEL, strES, sql


End Sub

Private Sub FillGridTAT()

    Dim tb As New Recordset
    Dim sn As New Recordset
    Dim sql As String
    Dim s As String
    Dim tot As Long

10  sql = "select distinct l.code,l.description, cd.listid as tcodeid from lists l, caselistlink cl, casedetails cd " & _
          "where cl.listid = l.listid and cd.caseid = cl.caseid " & _
          "and l.listtype = 'p' and COALESCE(cd.tissuetype, '') <> '' " & _
          "and cd.caseid in " & _
          "(select caseid from cases where sampleTaken between '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
          "and '" & Format(calTo, "yyyymmdd") & " 23:59:59')"

20  Set tb = New Recordset
30  RecOpenServer 0, tb, sql

40  Do While Not tb.EOF
50      sql = "select l.description, "
60      sql = sql & "sum(case "
70      sql = sql & "When Datediff(hour,SampleTaken,OrigValDate) <= (24 * 1) then 1 "
80      sql = sql & "else 0 "
90      sql = sql & "end) as 'Day1', "
100     sql = sql & "sum(case "
110     sql = sql & "When Datediff(hour,SampleTaken,OrigValDate) > (24 * 1) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 2) then 1 "
120     sql = sql & "else 0 "
130     sql = sql & "end)as 'Day2', "
140     sql = sql & "sum(case "
150     sql = sql & "When Datediff(hour,SampleTaken,OrigValDate) > (24 * 2) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 3) then 1 "
160     sql = sql & "else 0 "
170     sql = sql & "end)as 'Day3', "
180     sql = sql & "sum(case "
190     sql = sql & "   When Datediff(hour,SampleTaken,OrigValDate) > (24 * 3) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 4) then 1 "
200     sql = sql & "   else 0 "
210     sql = sql & "end)as 'Day4', "
220     sql = sql & "sum(case "
230     sql = sql & "   When Datediff(hour,SampleTaken,OrigValDate) > (24 * 4) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 5) then 1 "
240     sql = sql & "   else 0 "
250     sql = sql & "end)as 'Day5', "
260     sql = sql & "sum(case "
270     sql = sql & "   When Datediff(hour,SampleTaken,OrigValDate) > (24 * 5) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 6) then 1 "
280     sql = sql & "   else 0 "
290     sql = sql & "end)as 'Day6', "
300     sql = sql & "sum(case "
310     sql = sql & "When Datediff(hour,SampleTaken,OrigValDate) > (24 * 6) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 7) then 1 "
320     sql = sql & "else 0 "
330     sql = sql & "end)as 'Day7', "
340     sql = sql & "sum(case "
350     sql = sql & "   When Datediff(hour,SampleTaken,OrigValDate) > (24 * 7) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 8) then 1 "
360     sql = sql & "   else 0 "
370     sql = sql & "end)as 'Day8', "
380     sql = sql & "sum(case "
390     sql = sql & "   When Datediff(hour,SampleTaken,OrigValDate) > (24 * 8) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 9) then 1 "
400     sql = sql & "   else 0 "
410     sql = sql & "end)as 'Day9', "
420     sql = sql & "sum(case "
430     sql = sql & "When Datediff(hour,SampleTaken,OrigValDate) > (24 * 9) AND Datediff(hour,SampleTaken,OrigValDate) <= (24 * 10) then 1 "
440     sql = sql & "else 0 "
450     sql = sql & "end)as 'Day10', "
460     sql = sql & "sum(case "
470     sql = sql & "   When Datediff(hour,SampleTaken,OrigValDate) > (24 * 10) then 1 "
480     sql = sql & "   else 0 "
490     sql = sql & "end)as 'Day11' "
500     sql = sql & "From cases c, lists l "
510     sql = sql & "Where OrigValDate Is Not Null and caseid in "
520     sql = sql & "(select caseid from casedetails cd , lists l where cd.listid = l.listid and l.listid = " & tb!tcodeid & ") "
530     sql = sql & "and l.listid = " & tb!tcodeid & " group by l.description"
540     Set sn = New Recordset
550     RecOpenServer 0, sn, sql
560     If Not sn.EOF Then
570         tot = sn!Day1 + sn!Day2 + sn!Day3 + sn!Day4 + sn!Day5 + sn!Day6 + sn!Day7 + sn!Day8 + sn!Day9 + sn!Day10 + sn!Day11

580         s = tb!Description & vbTab & sn!Description & vbTab & Format$(sn!Day1 / tot, "##.00") * 100 & vbTab & _
                Format$(sn!Day2 / tot, "##.00") * 100 & vbTab & Format$(sn!Day3 / tot, "##.00") * 100 & vbTab & Format$(sn!Day4 / tot, "##.00") * 100 & vbTab & Format$(sn!Day5 / tot, "##.00") * 100 & vbTab & _
                Format$(sn!Day6 / tot, "##.00") * 100 & vbTab & Format$(sn!Day7 / tot, "##.00") * 100 & vbTab & Format$(sn!Day8 / tot, "##.00") * 100 & vbTab & Format$(sn!Day9 / tot, "##.00") * 100 & vbTab & _
                Format$(sn!Day10 / tot, "##.00") * 100 & vbTab & Format$(sn!Day11 / tot, "##.00") * 100
590         grdReport.AddItem s
600     End If
610     tb.MoveNext
620 Loop


End Sub

Private Sub FillGridAddendum()

    Dim tb As New Recordset
    Dim sn As New Recordset
    Dim sql As String
    Dim s As String
    Dim tps As String
    Dim totCases As Long

10  On Error GoTo FillGridAddendum_Error

20  sql = "select count(*) as totCases from Cases"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  If Not tb.EOF Then
60      totCases = tb!totCases
70  End If

80  sql = "select * from lists where code in ('Q020','Q021','Q022')"
90  Set tb = New Recordset
100 RecOpenServer 0, tb, sql

110 Do While Not tb.EOF
120     sql = "SELECT count(*)as tot FROM CaseListLink WHERE " & _
              "ListId = " & tb!ListId & " and datetimeofrecord between '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
              "and '" & Format(calTo, "yyyymmdd") & " 23:59:59'"
130     Set sn = New Recordset
140     RecOpenServer 0, sn, sql
150     If Not sn.EOF Then
160         tps = Format$(sn!tot / totCases, "##.00") * 100
170         s = tb!Description & vbTab & tps
180         grdReport.AddItem s
190     End If
200     tb.MoveNext
210 Loop




220 Exit Sub

FillGridAddendum_Error:

    Dim strES As String
    Dim intEL As Integer

230 intEL = Erl
240 strES = Err.Description
250 LogError "frmQAReports", "FillGridAddendum", intEL, strES, sql


End Sub

Private Sub FillGridRepToClin()

    Dim tb As New Recordset
    Dim sn As New Recordset
    Dim sql As String
    Dim s As String
    Dim tps As String
    Dim totCases As Long


10  On Error GoTo FillGridRepToClin_Error

20  sql = "select count(*) as totCases from Cases"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  If Not tb.EOF Then
60      totCases = tb!totCases
70  End If

80  sql = "select * from lists where code in ('Q023')"
90  Set tb = New Recordset
100 RecOpenServer 0, tb, sql

110 Do While Not tb.EOF
120     sql = "SELECT count(*)as tot FROM CaseListLink WHERE " & _
              "ListId = " & tb!ListId & " and datetimeofrecord between '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
              "and '" & Format(calTo, "yyyymmdd") & " 23:59:59'"
130     Set sn = New Recordset
140     RecOpenServer 0, sn, sql
150     If Not sn.EOF Then
160         tps = Format$(sn!tot / totCases, "##.00") * 100
170         s = tps
180         grdReport.AddItem s
190     End If
200     tb.MoveNext
210 Loop


220 Exit Sub

FillGridRepToClin_Error:

    Dim strES As String
    Dim intEL As Integer

230 intEL = Erl
240 strES = Err.Description
250 LogError "frmQAReports", "FillGridRepToClin", intEL, strES, sql


End Sub


Private Sub cmdHide_Click()
10  Unload Me
End Sub

Private Sub Form_Load()
10  calFrom = DateAdd("m", -1, Now)
20  calTo = Now
30  LoadReports
40  If blnIsTestMode Then EnableTestMode Me
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmSnomedSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAquire - Cellular Pathology"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmSnomedSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   9360
      Picture         =   "frmSnomedSearch.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export to Excel"
      Height          =   975
      Left            =   6885
      Picture         =   "frmSnomedSearch.frx":120C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   480
      Width           =   1035
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   975
      Left            =   8100
      Picture         =   "frmSnomedSearch.frx":1356
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Search"
      Top             =   480
      Width           =   1035
   End
   Begin VB.Frame fraMCode 
      Caption         =   "M Code"
      Height          =   3015
      Left            =   5400
      TabIndex        =   9
      Top             =   1680
      Width           =   4935
      Begin VB.CommandButton cmdRemoveM 
         Height          =   375
         Left            =   4440
         Picture         =   "frmSnomedSearch.frx":173E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdMCode 
         Height          =   255
         Left            =   4440
         Picture         =   "frmSnomedSearch.frx":19DE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox txtMCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1200
      End
      Begin MSComctlLib.ListView lstMCode 
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CutUp"
            Object.Width           =   6703
         EndProperty
      End
   End
   Begin VB.Frame fraTCode 
      Caption         =   "T Code"
      Height          =   3015
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   4935
      Begin VB.CommandButton cmdRemoveT 
         Height          =   375
         Left            =   4440
         Picture         =   "frmSnomedSearch.frx":1B43
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdTCode 
         Height          =   255
         Left            =   4440
         Picture         =   "frmSnomedSearch.frx":1DE3
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtTDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox txtTCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1200
      End
      Begin MSComctlLib.ListView lstTCode 
         Height          =   2055
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CutUp"
            Object.Width           =   6703
         EndProperty
      End
   End
   Begin Threed.SSPanel panSampleDates 
      Height          =   1185
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   2090
      _StockProps     =   15
      Caption         =   "Between Dates"
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
      Begin MSComCtl2.DTPicker calfrom 
         Height          =   345
         Left            =   210
         TabIndex        =   1
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
         TabIndex        =   3
         Top             =   330
         Width           =   945
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4605
      Left            =   240
      TabIndex        =   13
      Top             =   5280
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   8123
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   300
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
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
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
      TabIndex        =   25
      Top             =   9960
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label lblDescriptionSearch 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   4800
      Width           =   10080
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgReport 
      Height          =   240
      Left            =   7200
      Picture         =   "frmSnomedSearch.frx":1F48
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblMCodeListId 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6240
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblTCodeListId 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1320
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmSnomedSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pSearchType As String
Private FormLoaded As Boolean

Private Sub cmdCalc_Click()


10  With g
20      .Rows = 2
30      .AddItem ""
40      .RemoveItem 1
50  End With

60  If pSearchType = "1" Then
70      CalculateGrid1
80  ElseIf pSearchType = "2" Then
90      CalculateGrid2
100 ElseIf pSearchType = "3" Then
110     CalculateGrid3
120 ElseIf pSearchType = "4" Then
130     CalculateGrid4
140 ElseIf pSearchType = "5" Then
150     CalculateGrid5
160 End If

170 If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1

End Sub

Private Sub CalculateGrid1()
    Dim sql As String
    Dim tb As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    Dim TempCaseId As String


10  On Error GoTo CalculateGrid1_Error

20  For i = 1 To lstTCode.ListItems.Count
30      If lstMCode.ListItems.Count > 0 Then
40          For j = 1 To lstMCode.ListItems.Count
50              sql = "SELECT * FROM CaseListLink CL " & _
                      "INNER JOIN Demographics D ON CL.Caseid = D.CaseId " & _
                      "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
                      "WHERE CL.ListId = '" & lstMCode.ListItems(j).Tag & "' " & _
                      "AND CL.TissueTypeListId = '" & lstTCode.ListItems(i).Tag & "' " & _
                      "AND C.Validated  = 1 " & _
                      "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
                      "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59'"

60              Set tb = New Recordset
70              RecOpenClient 0, tb, sql

80              Do While Not tb.EOF

90                  If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
100                     TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
110                 Else
120                     TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
130                 End If

140                 s = vbTab & TempCaseId & vbTab & tb!TissueTypeLetter & _
                        vbTab & tb!PatientName & vbTab & Left(lstTCode.ListItems(i).Text, InStr(lstTCode.ListItems(i).Text, ":") - 1) & _
                        vbTab & Left(lstMCode.ListItems(j).Text, InStr(lstMCode.ListItems(j).Text, ":") - 1) & ""
150                 g.AddItem s

160                 g.Row = g.Rows - 1
170                 g.Col = 0
180                 g.CellPictureAlignment = flexAlignCenterCenter
190                 Set g.CellPicture = imgReport.Picture
200                 tb.MoveNext
210             Loop
220         Next j
230     Else
240         sql = "SELECT * FROM CaseListLink CL " & _
                  "INNER JOIN Demographics D ON CL.Caseid = D.CaseId " & _
                  "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
                  "LEFT JOIN Lists L ON L.ListId = CL.ListId " & _
                  "WHERE CL.TissueTypeListId = '" & lstTCode.ListItems(i).Tag & "' " & _
                  "AND C.Validated  = 1 " & _
                  "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
                  "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59'"
250         Set tb = New Recordset
260         RecOpenClient 0, tb, sql

270         Do While Not tb.EOF

280             If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
290                 TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
300             Else
310                 TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
320             End If

330             s = vbTab & TempCaseId & vbTab & tb!TissueTypeLetter & _
                    vbTab & tb!PatientName & vbTab & Left(lstTCode.ListItems(i).Text, InStr(lstTCode.ListItems(i).Text, ":") - 1) & _
                    vbTab & tb!Code & ""
340             g.AddItem s

350             g.Row = g.Rows - 1
360             g.Col = 0
370             g.CellPictureAlignment = flexAlignCenterCenter
380             Set g.CellPicture = imgReport.Picture
390             tb.MoveNext


400         Loop

410     End If

420 Next i





430 Exit Sub

CalculateGrid1_Error:

    Dim strES As String
    Dim intEL As Integer

440 intEL = Erl
450 strES = Err.Description
460 LogError "frmSnomedSearch", "CalculateGrid1", intEL, strES, sql


End Sub

Private Sub CalculateGrid2()
    Dim sql As String
    Dim tb As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim s As String

10  On Error GoTo CalculateGrid2_Error

20  For i = 1 To lstTCode.ListItems.Count
30      For j = 1 To lstMCode.ListItems.Count
40          sql = "SELECT Count(*) AS TotCases FROM CaseListLink CL " & _
                  "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
                  "WHERE CL.ListId = '" & lstMCode.ListItems(j).Tag & "' " & _
                  "AND CL.TissueTypeListId = '" & lstTCode.ListItems(i).Tag & "' " & _
                  "AND C.Validated  = 1 " & _
                  "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
                  "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59'"

50          Set tb = New Recordset
60          RecOpenClient 0, tb, sql

70          If Not tb.EOF Then
80              If j = 1 Then
90                  s = Left(lstTCode.ListItems(i).Text, InStr(lstTCode.ListItems(i).Text, ":") - 1) & vbTab & _
                        Trim(Mid(lstTCode.ListItems(i).Text, InStr(lstTCode.ListItems(i).Text, ":") + 1)) & vbTab
100             Else
110                 s = vbTab & vbTab
120             End If

130             s = s & Left(lstMCode.ListItems(j).Text, InStr(lstMCode.ListItems(j).Text, ":") - 1) & vbTab & _
                    Trim(Mid(lstMCode.ListItems(j).Text, InStr(lstMCode.ListItems(j).Text, ":") + 1)) & vbTab & _
                    tb!totCases & ""
140             g.AddItem s
150         End If
160     Next j

170 Next i

180 Exit Sub

CalculateGrid2_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmSnomedSearch", "CalculateGrid2", intEL, strES, sql


End Sub

Private Sub CalculateGrid3()
    Dim sql As String
    Dim tb As Recordset
    Dim sn As Recordset
    Dim i As Integer
    Dim s As String


10  On Error GoTo CalculateGrid3_Error

20  For i = 1 To lstTCode.ListItems.Count
30      sql = "SELECT DISTINCT CL.ListId, l.code, l.Description FROM CaseListLink CL " & _
              "INNER JOIN Lists L ON CL.ListId = l.ListId " & _
              "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
              "WHERE CL.TissueTypeListId = '" & lstTCode.ListItems(i).Tag & "' " & _
              "AND C.Validated  = 1 " & _
              "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
              "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59'"

40      Set tb = New Recordset
50      RecOpenClient 0, tb, sql
60      If Not tb.EOF Then
70          s = Left(lstTCode.ListItems(i).Text, InStr(lstTCode.ListItems(i).Text, ":") - 1) & vbTab
80          Do While Not tb.EOF
90              sql = "SELECT Count(*) AS TotCases FROM CaseListLink CL " & _
                      "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
                      "WHERE CL.TissueTypeListId = '" & lstTCode.ListItems(i).Tag & "' " & _
                      "AND CL.ListId = '" & tb!ListId & "' " & _
                      "AND C.Validated  = 1 " & _
                      "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
                      "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59'"
100             Set sn = New Recordset
110             RecOpenClient 0, sn, sql

120             If Not sn.EOF Then
130                 If s = "" Then
140                     s = vbTab
150                 End If

160                 s = s & tb!Code & vbTab & _
                        tb!Description & vbTab & _
                        sn!totCases & ""
170                 g.AddItem s
180             End If
190             tb.MoveNext
200             s = ""
210         Loop





220     End If

230 Next i



240 Exit Sub

CalculateGrid3_Error:

    Dim strES As String
    Dim intEL As Integer

250 intEL = Erl
260 strES = Err.Description
270 LogError "frmSnomedSearch", "CalculateGrid3", intEL, strES, sql


End Sub

Private Sub CalculateGrid4()
    Dim sql As String
    Dim tb As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim s As String
    Dim TempCaseId As String




10  On Error GoTo CalculateGrid4_Error

20  For i = 1 To lstTCode.ListItems.Count
30      s = vbTab & Left(lstTCode.ListItems(i).Text, InStr(lstTCode.ListItems(i).Text, ":") - 1)
40      For j = 1 To lstMCode.ListItems.Count
50          sql = "SELECT * FROM CaseListLink CL " & _
                  "INNER JOIN Demographics D ON CL.Caseid = D.CaseId " & _
                  "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
                  "WHERE CL.ListId = '" & lstMCode.ListItems(j).Tag & "' " & _
                  "AND CL.TissueTypeListId = '" & lstTCode.ListItems(i).Tag & "' " & _
                  "AND C.Validated  = 1 " & _
                  "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
                  "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59'"

60          Set tb = New Recordset
70          RecOpenClient 0, tb, sql


80          Do While Not tb.EOF
90              If s = "" Then
100                 s = vbTab
110             End If

120             If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
130                 TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
140             Else
150                 TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
160             End If

170             s = s & vbTab & TempCaseId & vbTab & tb!TissueTypeLetter & _
                    vbTab & Trim(Mid(lstTCode.ListItems(i).Text, InStr(lstTCode.ListItems(i).Text, ":") + 1)) & _
                    vbTab & Left(lstMCode.ListItems(j).Text, InStr(lstMCode.ListItems(j).Text, ":") - 1) & _
                    vbTab & tb!PatientName & ""
180             g.AddItem s

190             g.Row = g.Rows - 1
200             g.Col = 0
210             g.CellPictureAlignment = flexAlignCenterCenter
220             Set g.CellPicture = imgReport.Picture

230             tb.MoveNext
240             s = ""
250         Loop
260     Next j

270 Next i



280 Exit Sub

CalculateGrid4_Error:

    Dim strES As String
    Dim intEL As Integer

290 intEL = Erl
300 strES = Err.Description
310 LogError "frmSnomedSearch", "CalculateGrid4", intEL, strES, sql


End Sub

Private Sub CalculateGrid5()
    Dim sql As String
    Dim tb As Recordset
    Dim sn As Recordset
    Dim rs As Recordset
    Dim j As Integer
    Dim s As String


10  sql = "SELECT * FROM Lists " & _
          "WHERE ListType = 'Source' "    '& _

20                                              Set tb = New Recordset
30  RecOpenClient 0, tb, sql


40  Do While Not tb.EOF
50      s = tb!Description
60      For j = 1 To lstMCode.ListItems.Count
70          sql = "SELECT DISTINCT(D.Clinician), D.GP FROM CaseListLink CL " & _
                  "INNER JOIN Demographics D ON CL.Caseid = D.CaseId " & _
                  "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
                  "WHERE CL.ListId = '" & lstMCode.ListItems(j).Tag & "' " & _
                  "AND D.Source = '" & tb!Description & "' " & _
                  "AND C.Validated  = 1 " & _
                  "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
                  "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59' "
80          Set sn = New Recordset
90          RecOpenClient 0, sn, sql



100         If Not sn.EOF Then

110             If s = vbTab Then
120                 s = s & Left(lstMCode.ListItems(j).Text, InStr(lstMCode.ListItems(j).Text, ":") - 1)
130             Else
140                 s = s & vbTab & Left(lstMCode.ListItems(j).Text, InStr(lstMCode.ListItems(j).Text, ":") - 1)
150             End If

160             Do While Not sn.EOF


170                 If sn!Clinician & "" <> "" Then
180                     sql = "SELECT COUNT(*) AS TotCases FROM CaseListLink CL " & _
                              "INNER JOIN Demographics D ON CL.Caseid = D.CaseId " & _
                              "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
                              "WHERE CL.ListId = '" & lstMCode.ListItems(j).Tag & "' " & _
                              "AND D.Source = '" & tb!Description & "' " & _
                              "AND D.Clinician = '" & sn!Clinician & "' " & _
                              "AND C.Validated  = 1 " & _
                              "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
                              "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59' "
190                     Set rs = New Recordset
200                     RecOpenClient 0, rs, sql


210                     s = s & _
                            vbTab & sn!Clinician & _
                            vbTab & rs!totCases
220                     g.AddItem s
230                     s = vbTab
240                 ElseIf sn!GP & "" <> "" Then
250                     sql = "SELECT COUNT(*) AS TotCases FROM CaseListLink CL " & _
                              "INNER JOIN Demographics D ON CL.Caseid = D.CaseId " & _
                              "INNER JOIN Cases C ON C.Caseid = CL.CaseId " & _
                              "WHERE CL.ListId = '" & lstMCode.ListItems(j).Tag & "' " & _
                              "AND D.Source = '" & tb!Description & "' " & _
                              "AND D.GP = '" & sn!GP & "' " & _
                              "AND C.Validated  = 1 " & _
                              "AND CL.DateTimeOfRecord BETWEEN '" & Format(calFrom, "yyyymmdd") & " 00:00:00' " & _
                              "AND '" & Format(calTo, "yyyymmdd") & " 23:59:59' "
260                     Set rs = New Recordset
270                     RecOpenClient 0, rs, sql


280                     s = s & _
                            vbTab & sn!GP & _
                            vbTab & rs!totCases
290                     g.AddItem s
300                     s = vbTab
310                 End If
320                 sn.MoveNext

330             Loop
340         End If
350     Next j
360     tb.MoveNext
370 Loop

    'Next i



End Sub


Private Sub cmdExit_Click()
Unload Me
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
60  LogError "frmSnomedSearch", "cmdExport_Click", intEL, strES


End Sub


Private Sub cmdRemoveM_Click()
    Dim i As Integer
10  On Error GoTo cmdRemoveM_Click_Error

20  For i = lstMCode.ListItems.Count To 1 Step -1
30      If lstMCode.ListItems(i).Selected Then
40          lstMCode.ListItems.Remove (lstMCode.ListItems(i).Index)
50      End If
60  Next

70  If pSearchType = "1" Then
80      InitializeGrid1
90  ElseIf pSearchType = "2" Then
100     InitializeGrid2
110 ElseIf pSearchType = "3" Then
120     InitializeGrid3
130 ElseIf pSearchType = "4" Then
140     InitializeGrid4
150 ElseIf pSearchType = "5" Then
160     InitializeGrid5
170 End If


180 Exit Sub

cmdRemoveM_Click_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmSnomedSearch", "cmdRemoveM_Click", intEL, strES

End Sub

Private Sub cmdRemoveT_Click()
    Dim i As Integer
10  On Error GoTo cmdRemoveT_Click_Error

20  For i = lstTCode.ListItems.Count To 1 Step -1
30      If lstTCode.ListItems(i).Selected Then
40          lstTCode.ListItems.Remove (lstTCode.ListItems(i).Index)
50      End If
60  Next

70  If pSearchType = "1" Then
80      InitializeGrid1
90  ElseIf pSearchType = "2" Then
100     InitializeGrid2
110 ElseIf pSearchType = "3" Then
120     InitializeGrid3
130 ElseIf pSearchType = "4" Then
140     InitializeGrid4
150 ElseIf pSearchType = "5" Then
160     InitializeGrid5
170 End If

180 Exit Sub

cmdRemoveT_Click_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmSnomedSearch", "cmdRemoveT_Click", intEL, strES

End Sub

Private Sub cmdTCode_Click()
    Dim i As Integer
    Dim found As Boolean

10  found = False
20  If lstTCode.ListItems.Count > 0 Then
30      For i = 1 To lstTCode.ListItems.Count
40          If Left(lstTCode.ListItems(i).Text, InStr(lstTCode.ListItems(i).Text, ":") - 1) = txtTCode Then
50              frmMsgBox.Msg "T Code already added to list", mbOKOnly, , mbInformation
60              found = True
70              txtTCode = ""
80              txtTDescription = ""
90              Exit Sub
100         End If
110     Next
120     If found = False Then
130         lstTCode.ListItems.Add lstTCode.ListItems.Count + 1, , txtTCode & " : " & txtTDescription
140         lstTCode.ListItems(lstTCode.ListItems.Count).Tag = lblTCodeListId
150     End If

160 Else
170     lstTCode.ListItems.Add lstTCode.ListItems.Count + 1, , txtTCode & " : " & txtTDescription
180     lstTCode.ListItems(lstTCode.ListItems.Count).Tag = lblTCodeListId
190 End If

200 Set lstTCode.SelectedItem = Nothing
210 txtTCode = ""
220 txtTDescription = ""
End Sub

Private Sub cmdMCode_Click()

    Dim i As Integer
    Dim found As Boolean

10  found = False

20  If lstMCode.ListItems.Count > 0 Then
30      For i = 1 To lstMCode.ListItems.Count
40          If lstMCode.ListItems(i).Text = txtMCode Then
50              frmMsgBox.Msg "M Code already added to list", mbOKOnly, , mbInformation
60              found = True
70              txtMCode = ""
80              txtMDescription = ""
90              Exit Sub
100         End If
110     Next
120     If found = False Then
130         lstMCode.ListItems.Add lstMCode.ListItems.Count + 1, , txtMCode & " : " & txtMDescription
140         lstMCode.ListItems(lstMCode.ListItems.Count).Tag = lblMCodeListId
150     End If

160 Else
170     lstMCode.ListItems.Add lstMCode.ListItems.Count + 1, , txtMCode & " : " & txtMDescription
180     lstMCode.ListItems(lstMCode.ListItems.Count).Tag = lblMCodeListId
190 End If

200 Set lstMCode.SelectedItem = Nothing
210 txtMCode = ""
220 txtMDescription = ""
End Sub

Private Sub Form_Activate()
10  If Not FormLoaded Then
20      If pSearchType = "1" Then
30          InitializeGrid1
40      ElseIf pSearchType = "2" Then
50          InitializeGrid2
60      ElseIf pSearchType = "3" Then
70          fraMCode.Enabled = False
80          InitializeGrid3
90      ElseIf pSearchType = "4" Then
100         InitializeGrid4
110     ElseIf pSearchType = "5" Then
120         fraTCode.Enabled = False
130         InitializeGrid5
140     End If
150     FormLoaded = True
160 End If
End Sub

Private Sub Form_Load()

10    ChangeFont Me, "Arial"
'20    frmSnomedSearch_ChangeLanguage

30    calTo = Format(Now, "dd/MMM/yyyy")
40    calFrom = Format(Now - 7, "dd/MMM/yyyy")

50    lblLoggedIn = UserName

60    Me.Caption = "NetAcquire - Cellular Pathology. Version " & strVersion
70    If blnIsTestMode Then EnableTestMode Me

End Sub

Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub InitializeGrid1()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 6: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "": .ColWidth(0) = 400: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Case Id": .ColWidth(1) = 1100: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "Specimen Letter": .ColWidth(2) = 1500: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Name": .ColWidth(3) = 2500: .ColAlignment(3) = flexAlignLeftCenter
160     .TextMatrix(0, 4) = "T Code": .ColWidth(4) = 2000: .ColAlignment(4) = flexAlignLeftCenter
170     .TextMatrix(0, 5) = "M Code": .ColWidth(5) = 2000: .ColAlignment(5) = flexAlignLeftCenter

180 End With
End Sub
Private Sub InitializeGrid2()

10  With g
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

120     .TextMatrix(0, 0) = "T Code": .ColWidth(0) = 1500: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Tissue Type": .ColWidth(1) = 2500: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "M Code": .ColWidth(2) = 1500: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Morphology": .ColWidth(3) = 2500: .ColAlignment(3) = flexAlignLeftCenter
160     .TextMatrix(0, 4) = "No. Of Cases": .ColWidth(4) = 1500: .ColAlignment(4) = flexAlignLeftCenter

170 End With
End Sub

Private Sub InitializeGrid3()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 4: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "T Code": .ColWidth(0) = 2000: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "M Code": .ColWidth(1) = 2000: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "Morphology": .ColWidth(2) = 2500: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "No. Of Cases": .ColWidth(3) = 1500: .ColAlignment(3) = flexAlignLeftCenter

160 End With
End Sub

Private Sub InitializeGrid4()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 7: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "": .ColWidth(0) = 400: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "T Code": .ColWidth(1) = 1250: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "Case Id": .ColWidth(2) = 1100: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Specimen Letter": .ColWidth(3) = 1500: .ColAlignment(3) = flexAlignLeftCenter
160     .TextMatrix(0, 4) = "T Code Expansion": .ColWidth(4) = 2200: .ColAlignment(4) = flexAlignLeftCenter
170     .TextMatrix(0, 5) = "M Code": .ColWidth(5) = 1250: .ColAlignment(5) = flexAlignLeftCenter
180     .TextMatrix(0, 6) = "Patient Name": .ColWidth(6) = 2000: .ColAlignment(6) = flexAlignLeftCenter

190 End With
End Sub


Private Sub InitializeGrid5()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 4: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "Hospital": .ColWidth(0) = 2000: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "M Code": .ColWidth(1) = 1250: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "Clinician": .ColWidth(2) = 2000: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "No. of Cases": .ColWidth(3) = 1000: .ColAlignment(3) = flexAlignLeftCenter

160 End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
10  frmWorklist.Enabled = True
20  FormLoaded = False
30  Unload frmList
40  Unload frmWorkSheet
End Sub

Private Sub g_Click()
10  If pSearchType = "1" Or pSearchType = "4" Then
20      If g.Col = 0 Then
30          If pSearchType = "1" Then
40              CaseNo = Replace(g.TextMatrix(g.Row, 1), " " & sysOptCaseIdSeperator(0) & " ", "")
50          ElseIf pSearchType = "4" Then
60              CaseNo = Replace(g.TextMatrix(g.Row, 2), " " & sysOptCaseIdSeperator(0) & " ", "")
70          End If

80          PrintHistology "", True
90          With frmRichText
100             .rtb.SelStart = 0
110             .Show 1
120         End With

130     End If
140 End If
End Sub

Private Sub txtTCode_KeyPress(KeyAscii As Integer)
10  frmList.PrevCode = txtTCode
20  frmList.PrevDesc = txtTDescription
30  Set frmList.txtCode = txtTCode
40  Set frmList.txtDescription = txtTDescription
50  Set frmList.tempListId = lblTCodeListId
60  frmList.SearchByCode = True
70  frmList.ListType = "T"
80  frmList.Show
90  frmList.Move Me.Left + fraTCode.Left + txtTCode.Left + 50, Me.Top + fraTCode.Top + txtTCode.Top + 625
100 If KeyAscii = 13 Then
110     KeyAscii = 0
120 End If
End Sub

Private Sub txtTDescription_KeyPress(KeyAscii As Integer)
10  frmList.PrevCode = txtTCode
20  frmList.PrevDesc = txtTDescription
30  Set frmList.txtCode = txtTCode
40  Set frmList.txtDescription = txtTDescription
50  Set frmList.tempListId = lblTCodeListId
60  frmList.SearchByCode = False
70  frmList.ListType = "T"
80  frmList.Show
90  frmList.Move Me.Left + fraTCode.Left + txtTCode.Left + 50, Me.Top + fraTCode.Top + txtTCode.Top + 625
100 If KeyAscii = 13 Then
110     KeyAscii = 0
120 End If
End Sub

Private Sub txtMCode_KeyPress(KeyAscii As Integer)
10  frmList.PrevCode = txtMCode
20  frmList.PrevDesc = txtMDescription
30  Set frmList.txtCode = txtMCode
40  Set frmList.txtDescription = txtMDescription
50  Set frmList.tempListId = lblMCodeListId
60  frmList.SearchByCode = True
70  frmList.ListType = "M"
80  frmList.Show
90  frmList.Move Me.Left + fraMCode.Left + txtMCode.Left + 50, Me.Top + fraMCode.Top + txtMCode.Top + 625
100 If KeyAscii = 13 Then
110     KeyAscii = 0
120 End If
End Sub

Private Sub txtMDescription_KeyPress(KeyAscii As Integer)
10  frmList.PrevCode = txtMCode
20  frmList.PrevDesc = txtMDescription
30  Set frmList.txtCode = txtMCode
40  Set frmList.txtDescription = txtMDescription
50  Set frmList.tempListId = lblMCodeListId
60  frmList.SearchByCode = False
70  frmList.ListType = "M"
80  frmList.Show
90  frmList.Move Me.Left + fraMCode.Left + txtMCode.Left + 50, Me.Top + fraMCode.Top + txtMCode.Top + 625
100 If KeyAscii = 13 Then
110     KeyAscii = 0
120 End If
End Sub

Public Property Let SearchType(ByVal sNewValue As String)

10  pSearchType = sNewValue

End Property

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCasesCurrentlyOpened 
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReloadList 
      Caption         =   "&Reload"
      Height          =   735
      Left            =   8175
      Picture         =   "frmCasesCurrentlyOpened.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   405
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   735
      Left            =   8175
      Picture         =   "frmCasesCurrentlyOpened.frx":0373
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2985
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   735
      Left            =   8175
      Picture         =   "frmCasesCurrentlyOpened.frx":0D75
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4035
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdReport 
      Height          =   4365
      Left            =   195
      TabIndex        =   0
      Top             =   405
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   7699
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   12648447
      ForeColor       =   -2147483625
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      AllowUserResizing=   1
      FormatString    =   "<Case Id                  |<User Name                            |<Compute Name                |<Date/Time                       "
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   195
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label Label 
      Caption         =   "List of Case Ids opened for editing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   225
      TabIndex        =   4
      Top             =   165
      Width           =   7800
   End
End
Attribute VB_Name = "frmCasesCurrentlyOpened"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
10  Unload Me
End Sub

Private Sub cmdReloadList_Click()

10  ClearGrid
20  LoadLockedForEditingCaseIds

End Sub

Private Sub cmdRemove_Click()
    Dim sql As String
    Dim Y As Integer

10  On Error GoTo cmdRemove_Click_Error

20  If iMsg("Are you sure you wish to unlock Case Id from editing mode?", vbQuestion + vbYesNo) = vbYes Then

30      For Y = 1 To grdReport.Rows - 1
40          grdReport.Row = Y
50          If grdReport.CellBackColor = vbYellow Then
60              Exit For
70          End If
80      Next

90      If grdReport.TextMatrix(Y, 0) <> "" Then
100         sql = "DELETE FROM CasesLocked WHERE " & _
                  "CaseId = '" & grdReport.TextMatrix(Y, 0) & "' AND Username = '" & AddTicks(grdReport.TextMatrix(Y, 1)) & "' AND MachineName = '" & grdReport.TextMatrix(Y, 2) & "'"
110         Cnxn(0).Execute sql
120     End If

130     ClearGrid
140     LoadLockedForEditingCaseIds
150 End If

160 Exit Sub

cmdRemove_Click_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmCasesCurrentlyOpened", "cmdRemove_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10    ChangeFont Me, "Arial"
'20    frmCasesCurrentlyOpened_ChangeLanguage
30    If UCase$(UserMemberOf) <> "MANAGER" Then
40      cmdRemove.Visible = False
50    End If

60    LoadLockedForEditingCaseIds
End Sub

Private Sub LoadLockedForEditingCaseIds()

    Dim tb As Recordset
    Dim sql As String
    Dim s As String

10  On Error GoTo LoadLockedForEditingCaseIds_Error

20  cmdRemove.Enabled = False

30  sql = "SELECT * FROM CasesLocked"

40  Set tb = New Recordset
50  RecOpenClient 0, tb, sql

60  Do While Not tb.EOF
70      s = tb!CaseId & "" & vbTab & tb!UserName & "" & vbTab & tb!MachineName & "" & vbTab & Format(tb!DateTimeOfRecord, "dd/mmm/yyyy hh:mm:ss")
80      grdReport.AddItem s
90      tb.MoveNext
100 Loop


110 If grdReport.Rows > 2 And grdReport.TextMatrix(1, 0) = "" Then
120     grdReport.RemoveItem 1
130 End If

140 Exit Sub

LoadLockedForEditingCaseIds_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmCasesCurrentlyOpened", "LoadLockedForEditingCaseIds", intEL, strES, sql

End Sub

Private Sub ClearGrid()

10  grdReport.Rows = 2
20  grdReport.AddItem "", 1
30  grdReport.RemoveItem 2

End Sub

Private Sub grdReport_Click()
    Dim ySave As Integer
    Dim X As Integer
    Dim Y As Integer

10  On Error GoTo grdReport_Click_Error

20  ySave = grdReport.Row

30  grdReport.Col = 0
40  If grdReport.TextMatrix(ySave, 0) <> "" Then
50      For Y = 1 To grdReport.Rows - 1
60          grdReport.Row = Y
70          If grdReport.CellBackColor = vbYellow Then
80              For X = 0 To grdReport.Cols - 1
90                  grdReport.Col = X
100                 grdReport.CellBackColor = 0
110             Next
120             Exit For
130         End If
140     Next
150     grdReport.Row = ySave

160     For X = 0 To grdReport.Cols - 1
170         grdReport.Col = X
180         grdReport.CellBackColor = vbYellow
190     Next

200     cmdRemove.Enabled = True
210 End If

220 Exit Sub

grdReport_Click_Error:

    Dim strES As String
    Dim intEL As Integer

230 intEL = Erl
240 strES = Err.Description
250 LogError "frmCasesCurrentlyOpened", "grdReport_Click", intEL, strES

End Sub

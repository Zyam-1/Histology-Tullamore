VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmListsDates 
   Caption         =   "Non-Working days"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   9855
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbYear 
      Height          =   360
      Left            =   60
      TabIndex        =   9
      Top             =   450
      Width           =   1020
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add non-working days between Monday to Friday"
      Height          =   1455
      Left            =   60
      TabIndex        =   3
      Top             =   945
      Width           =   8085
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1740
         MaxLength       =   100
         TabIndex        =   5
         Top             =   975
         Width           =   5085
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   855
         Left            =   7080
         Picture         =   "frmListsDates.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTDate 
         Height          =   390
         Left            =   135
         TabIndex        =   11
         Top             =   975
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   688
         _Version        =   393216
         Format          =   82378753
         CurrentDate     =   41177
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1770
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   720
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   8475
      Picture         =   "frmListsDates.frx":033E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6495
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   855
      Left            =   8475
      Picture         =   "frmListsDates.frx":0680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5385
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8490
      Picture         =   "frmListsDates.frx":0782
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2595
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4755
      Left            =   60
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2595
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   8387
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmListsDates.frx":0951
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "Work Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   210
      Width           =   915
   End
End
Attribute VB_Name = "frmListsDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbYear_Click()

10  DTDate = "01/01/" & cmbYear
20  FillG

End Sub

Private Sub cmbYear_KeyPress(KeyAscii As Integer)
10  KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    Dim n As Integer
    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo cmdAdd_Click_Error

20  txtText = Trim$(txtText)

30  If txtText = "" Then
40      txtText.SetFocus
50      Exit Sub
60  End If

70  If cmbYear <> Year(DTDate) Then
80      frmMsgBox.Msg "Date not in selected year", , , mbExclamation
90      If TimedOut Then Unload Me: Exit Sub
100     Exit Sub
110 End If

    'If Weekend date entered

120 If Weekday(DTDate) = 1 Or Weekday(DTDate) = 7 Then    'Saturaday or Sunday
130     frmMsgBox.Msg "Date selected is a weekend date", , , mbExclamation
140     If TimedOut Then Unload Me: Exit Sub
150     Exit Sub
160 End If

170 For n = 1 To g.Rows - 1
180     If g.TextMatrix(n, 0) = DTDate Then
190         frmMsgBox.Msg "Date already entered", , , mbExclamation
200         If TimedOut Then Unload Me: Exit Sub
210         Exit Sub
220     End If
230 Next

240 sql = "SELECT * FROM NonWorkDays WHERE WorkYear = '" & Year(DTDate) & "' " & _
          "AND NonWorkDate = '" & Format(DTDate, "dd/mmm/yyyy") & "'"
250 Set tb = New Recordset
260 RecOpenServer 0, tb, sql

270 If Not tb.EOF Then
280     frmMsgBox.Msg "This date is already entered", , , mbExclamation
290     If TimedOut Then Unload Me: Exit Sub
300     Exit Sub
310 End If


320 g.AddItem Format(DTDate, "dd/mm/yyyy") & vbTab & txtText
330 g.Row = g.Rows - 1

340 If g.Rows > 2 And g.TextMatrix(1, 1) = "" Then
350     g.RemoveItem 1
360 End If

370 txtText = ""

380 cmdSave.Visible = True


390 Exit Sub

cmdAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

400 intEL = Erl
410 strES = Err.Description
420 LogError "frmListsDates", "cmdAdd_Click", intEL, strES, sql

End Sub

Private Sub cmdDelete_Click()
    Dim Y As Integer
    Dim sql As String
    Dim s As String

10  On Error GoTo cmdDelete_Click_Error

20  g.Col = 0
30  For Y = 1 To g.Rows - 1
40      g.Row = Y
50      If g.CellBackColor = vbYellow Then
60          s = "Delete <" & g.TextMatrix(Y, 1) & ">" & "[" & g.TextMatrix(Y, 0) & "]" & " ?"

70          Answer = frmMsgBox.Msg(s, mbYesNo, , mbQuestion)

80          If TimedOut Then Unload Me: Exit Sub
90          If Answer = 1 Then
100             sql = "Delete from NonWorkDays where " & _
                      "NonWorkDate = '" & Format(g.TextMatrix(Y, 0), "dd/mmm/yyyy") & "' " & _
                      "AND DateDescription = '" & g.TextMatrix(Y, 1) & "' "
110             Cnxn(0).Execute sql

120         End If
130         Exit For
140     End If
150 Next

160 cmdDelete.Enabled = False
170 FillG

180 Exit Sub

cmdDelete_Click_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmListsDates", "cmdDelete_Click", intEL, strES, sql

End Sub

Private Sub cmdExit_Click()
10  Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim Y As Integer

10  For Y = 1 To g.Rows - 1
20      SaveDates Y
30  Next

40  FillG

50  txtText = ""

60  cmdSave.Visible = False
70  cmdDelete.Enabled = False

End Sub

Private Sub Form_Load()

10  cmbYear.AddItem "2010"
20  cmbYear.AddItem "2011"
30  cmbYear.AddItem "2012"
40  cmbYear.AddItem "2013"
50  cmbYear.AddItem "2014"
60  cmbYear.AddItem "2015"
70  cmbYear.AddItem "2016"
80  cmbYear.AddItem "2017"
90  cmbYear.AddItem "2018"
100 cmbYear.AddItem "2019"
110 cmbYear.AddItem "2020"

120 cmbYear = Year(Now)
130 DTDate = "01/01/" & cmbYear
140 FillG

End Sub


Private Sub SaveDates(Y As Integer)
    Dim tb As Recordset
    Dim sql As String


10  On Error GoTo SaveDates_Error

20  sql = "SELECT * FROM NonWorkDays WHERE " & _
          "WorkYear = '" & cmbYear & "' " & _
          "AND NonWorkDate = '" & Format(g.TextMatrix(Y, 0), "dd/mmm/yyyy") & "'"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  If tb.EOF Then
60      tb.AddNew
70  End If
80  tb!WorkYear = cmbYear
90  tb!NonWorkDate = Format(g.TextMatrix(Y, 0), "dd/mmm/yyyy")
100 tb!DateDescription = g.TextMatrix(Y, 1)
110 tb!ListOrder = Y
120 tb!UserName = UserName
130 tb.Update

140 Exit Sub

SaveDates_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmListsDates", "SaveDates", intEL, strES, sql

End Sub

Private Sub FillG()

10  g.Rows = 2
20  g.AddItem ""
30  g.RemoveItem 1
40  cmdDelete.Enabled = False

50  LoadDates

60  If g.Rows > 2 Then
70      g.RemoveItem 1
80  End If

End Sub


Private Sub LoadDates()
    Dim s As String
    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo LoadDates_Error

20  sql = "SELECT * FROM NonWorkDays WHERE " & _
          "WorkYear = '" & cmbYear & "' " & _
          "ORDER BY ListOrder"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql
50  Do While Not tb.EOF
60      s = Format(tb!NonWorkDate, "dd/mm/yyyy") & vbTab & _
            tb!DateDescription & ""
70      g.AddItem s
80      g.Row = g.Rows - 1

90      tb.MoveNext
100 Loop



110 Exit Sub

LoadDates_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmListsDates", "LoadDates", intEL, strES, sql

End Sub


Private Sub g_Click()

    Dim X As Integer
    Dim Y As Integer
    Dim ySave As Integer

10  On Error GoTo g_Click_Error

20  ySave = g.Row

30  g.Visible = False
40  g.Col = 0
50  For Y = 1 To g.Rows - 1
60      g.Row = Y
70      If g.CellBackColor = vbYellow Then
80          For X = 0 To g.Cols - 1
90              g.Col = X
100             g.CellBackColor = 0
110         Next
120         Exit For
130     End If
140 Next
150 g.Row = ySave
160 g.Visible = True

170 If Len(g) > 0 Then
180     cmdDelete.Enabled = True
190     For X = 0 To g.Cols - 1
200         g.Col = X
210         g.CellBackColor = vbYellow
220     Next
230 End If

240 Exit Sub

g_Click_Error:

    Dim strES As String
    Dim intEL As Integer

250 intEL = Erl
260 strES = Err.Description
270 LogError "frmListsDates", "g_Click", intEL, strES

End Sub

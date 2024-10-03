VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListDestinations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cellular Pathology - Destinations"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   855
      Left            =   7080
      Picture         =   "frmListDestinations.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8400
      Picture         =   "frmListDestinations.frx":033E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   8400
      Picture         =   "frmListDestinations.frx":050D
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   855
      Left            =   8400
      Picture         =   "frmListDestinations.frx":084F
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9360
      Top             =   3600
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9360
      Top             =   4380
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Destination"
      Height          =   1215
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   8085
      Begin VB.TextBox txtText 
         Height          =   315
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   3855
      End
      Begin VB.ComboBox cmbCodes 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2970
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   480
         Width           =   450
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6855
      Left            =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1710
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   12091
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmListDestinations.frx":0951
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
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
      TabIndex        =   12
      Top             =   8640
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   9900
      Picture         =   "frmListDestinations.frx":09F2
      Top             =   750
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   9900
      Picture         =   "frmListDestinations.frx":0CC8
      Top             =   270
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmListDestinations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FireCounter As Integer

Private pListType As String
Private pListTypeName As String    'Singular name



Private Sub cmdExit_Click()
10  Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim Y As Integer
    Dim tb As Recordset
    Dim sql As String



10  On Error GoTo cmdSave_Click_Error

20  For Y = 1 To g.Rows - 1
30      If g.TextMatrix(Y, 1) <> "" Then
40          sql = "SELECT * FROM Destinations WHERE " & _
                  "Destination = '" & g.TextMatrix(Y, 1) & "' " & _
                  "AND Code = '" & g.TextMatrix(Y, 0) & "'"
50          Set tb = New Recordset
60          RecOpenServer 0, tb, sql
70          If tb.EOF Then
80              tb.AddNew
90          End If
100         tb!Code = g.TextMatrix(Y, 0)
110         tb!ListType = pListType
120         tb!Destination = g.TextMatrix(Y, 1)
130         tb!UserName = UserName
140         tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")


150         tb.Update
160     End If
170 Next

180 FillG

190 cmbCodes = ""
200 txtText = ""
210 cmbCodes.SetFocus

220 cmdSave.Visible = False
230 cmdDelete.Enabled = False

240 Exit Sub

cmdSave_Click_Error:

    Dim strES As String
    Dim intEL As Integer

250 intEL = Erl
260 strES = Err.Description
270 LogError "frmListDestinations", "cmdSave_Click", intEL, strES, sql

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10  If cmdSave.Visible Then
20      Answer = frmMsgBox.Msg("Cancel without saving?", mbYesNo, , mbQuestion)
30      If TimedOut Then Unload Me: Exit Sub
40      If Answer = 2 Then
50          Cancel = True
60          Exit Sub
70      End If
80  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

10  pListType = ""
20  pListTypeName = ""

End Sub

Private Sub FillG()

    Dim s As String
    Dim tb As Recordset
    Dim sql As String


10  On Error GoTo FillG_Error

20  g.Rows = 2
30  g.AddItem ""
40  g.RemoveItem 1

50  sql = "SELECT * FROM Destinations WHERE " & _
          "ListType = '" & pListType & "' " & _
          "ORDER BY Code"
60  Set tb = New Recordset
70  RecOpenServer 0, tb, sql
80  Do While Not tb.EOF
90      s = tb!Code & "" & vbTab & _
            tb!Destination & ""
100     g.AddItem s
110     g.Row = g.Rows - 1



120     tb.MoveNext
130 Loop

140 If g.Rows > 2 Then
150     g.RemoveItem 1
160 End If



170 Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmListDestinations", "FillG", intEL, strES, sql


End Sub

Private Sub InitializeGrid()
    Dim i As Integer
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

120     .TextMatrix(0, 0) = "Code": .ColWidth(0) = 1200: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Destination": .ColWidth(1) = 4900: .ColAlignment(1) = flexAlignLeftCenter

140     For i = 0 To .Cols - 1
150         If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
160             .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
170         End If
180     Next i
190 End With
End Sub

Private Sub Form_Load()
10  LoadCombo
20  InitializeGrid
30  FillG
40  lblLoggedIn = UserName
50  If blnIsTestMode Then EnableTestMode Me
End Sub


Private Sub LoadCombo()
    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo LoadCombo_Error

20  If UCase(pListTypeName) = "STAINS" Then
30      sql = "SELECT * FROM Lists WHERE [External] = 1 AND (Listtype = 'IS' OR ListType = 'SS')"
40  Else
50      sql = "SELECT * FROM Lists WHERE [External] = 1 AND Listtype = 'Q'"
60  End If
70  Set tb = New Recordset
80  RecOpenServer 0, tb, sql

90  If Not tb.EOF Then
100     Do Until tb.EOF
110         cmbCodes.AddItem tb!Code & " - " & tb!Description
120         tb.MoveNext
130     Loop

140 End If

150 Exit Sub

LoadCombo_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmListDestinations", "LoadCombo", intEL, strES, sql


End Sub

Public Property Let ListTypeName(ByVal strNewValue As String)

10  pListTypeName = strNewValue

End Property

Private Sub cmdAdd_Click()

    Dim Code As String

10  On Error GoTo cmdAdd_Click_Error

20  Code = Left(cmbCodes, InStr(1, cmbCodes, " - "))
30  txtText = Trim$(txtText)

40  If cmbCodes = "" Then
50      cmbCodes.SetFocus
60      Exit Sub
70  End If

80  If txtText = "" Then
90      txtText.SetFocus
100     Exit Sub
110 End If


120 g.AddItem Code & vbTab & txtText

130 If g.Rows > 2 And g.TextMatrix(1, 1) = "" Then
140     g.RemoveItem 1
150 End If

160 cmbCodes = ""
170 txtText = ""


180 cmdSave.Visible = True

190 cmbCodes.SetFocus



200 Exit Sub

cmdAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

210 intEL = Erl
220 strES = Err.Description
230 LogError "frmListDestinations", "cmdAdd_Click", intEL, strES


End Sub


Private Sub tmrDown_Timer()

10  FireDown

End Sub


Private Sub tmrUp_Timer()

10  FireUp

End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)

10  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

End Sub


Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized And Me.Top >= 0 Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub FireDown()

    Dim n As Integer
    Dim s As String
    Dim X As Integer
    Dim VisibleRows As Integer

10  If g.Row = g.Rows - 1 Then Exit Sub
20  n = g.Row

30  FireCounter = FireCounter + 1
40  If FireCounter > 5 Then
50      tmrDown.Interval = 100
60  End If

70  VisibleRows = g.Height \ g.RowHeight(1) - 1

80  g.Visible = False

90  s = ""
100 For X = 0 To g.Cols - 1
110     s = s & g.TextMatrix(n, X) & vbTab
120 Next
130 s = Left$(s, Len(s) - 1)

140 g.RemoveItem n
150 If n < g.Rows Then
160     g.AddItem s, n + 1
170     g.Row = n + 1
180 Else
190     g.AddItem s
200     g.Row = g.Rows - 1
210 End If

220 For X = 0 To g.Cols - 1
230     g.Col = X
240     g.CellBackColor = vbYellow
250 Next

260 If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
270     If g.Row - VisibleRows + 1 > 0 Then
280         g.TopRow = g.Row - VisibleRows + 1
290     End If
300 End If

310 g.Visible = True

320 cmdSave.Visible = True

End Sub

Private Sub FireUp()

    Dim n As Integer
    Dim s As String
    Dim X As Integer

10  If g.Row = 1 Then Exit Sub

20  FireCounter = FireCounter + 1
30  If FireCounter > 5 Then
40      tmrUp.Interval = 100
50  End If

60  n = g.Row

70  g.Visible = False

80  s = ""
90  For X = 0 To g.Cols - 1
100     s = s & g.TextMatrix(n, X) & vbTab
110 Next
120 s = Left$(s, Len(s) - 1)

130 g.RemoveItem n
140 g.AddItem s, n - 1

150 g.Row = n - 1
160 For X = 0 To g.Cols - 1
170     g.Col = X
180     g.CellBackColor = vbYellow
190 Next

200 If Not g.RowIsVisible(g.Row) Then
210     g.TopRow = g.Row
220 End If

230 g.Visible = True

240 cmdSave.Visible = True

End Sub

Public Property Let ListType(ByVal Code As String)

10  pListType = Code

End Property

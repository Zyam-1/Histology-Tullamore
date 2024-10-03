VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmGps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cellular Pathology - G. P. Entry"
   ClientHeight    =   10140
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   14115
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10140
   ScaleWidth      =   14115
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   12840
      Picture         =   "frmGps.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   855
      Left            =   12840
      Picture         =   "frmGps.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   915
      Left            =   12840
      Picture         =   "frmGps.frx":0709
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   12840
      Picture         =   "frmGps.frx":0B24
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   13380
      Top             =   5340
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   13380
      Top             =   4770
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add GP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   180
      TabIndex        =   15
      Top             =   45
      Width           =   12285
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   9630
         MaxLength       =   5
         TabIndex        =   3
         Top             =   480
         Width           =   1635
      End
      Begin VB.TextBox txtCounty 
         Height          =   285
         Left            =   810
         TabIndex        =   7
         Top             =   1080
         Width           =   3525
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   855
         Left            =   11520
         Picture         =   "frmGps.frx":0C6E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtFAX 
         Height          =   285
         Left            =   9630
         TabIndex        =   8
         Top             =   1140
         Width           =   1635
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   9630
         TabIndex        =   6
         Top             =   810
         Width           =   1635
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   4350
         TabIndex        =   2
         Top             =   450
         Width           =   4155
      End
      Begin VB.TextBox txtForeName 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   450
         Width           =   2295
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   810
         TabIndex        =   0
         Top             =   450
         Width           =   1185
      End
      Begin VB.TextBox txtAddr1 
         Height          =   285
         Left            =   4350
         TabIndex        =   5
         Top             =   780
         Width           =   4155
      End
      Begin VB.TextBox txtAddr0 
         Height          =   285
         Left            =   810
         TabIndex        =   4
         Top             =   780
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9100
         TabIndex        =   28
         Top             =   510
         Width           =   375
      End
      Begin VB.Label lblGpid 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5040
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   8640
         TabIndex        =   26
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   4440
         TabIndex        =   25
         Top             =   1080
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "County"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9150
         TabIndex        =   21
         Top             =   1170
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9000
         TabIndex        =   20
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4380
         TabIndex        =   19
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Forename"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2070
         TabIndex        =   18
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   810
         TabIndex        =   17
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   810
         Width           =   570
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7395
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   13044
      _Version        =   393216
      Cols            =   14
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
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   225
      Left            =   150
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   30
      Top             =   9360
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   12690
      TabIndex        =   23
      Top             =   4050
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmGps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Dim strFullName As String

Private FireCounter As Long

Private Sub InitializeGrid()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 13: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth
        '^In Use |<Code         |<Text                                                                       |^Allow Delete
120     .TextMatrix(0, 0) = "": .ColWidth(0) = 400: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "In Use": .ColWidth(1) = 600: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "Code": .ColWidth(2) = 600: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "GP Name": .ColWidth(3) = 1800: .ColAlignment(3) = flexAlignLeftCenter
160     .TextMatrix(0, 4) = "Title": .ColWidth(4) = 600: .ColAlignment(4) = flexAlignLeftCenter
170     .TextMatrix(0, 5) = "FirstName": .ColWidth(5) = 800: .ColAlignment(5) = flexAlignLeftCenter
180     .TextMatrix(0, 6) = "Surname": .ColWidth(6) = 1000: .ColAlignment(6) = flexAlignLeftCenter
190     .TextMatrix(0, 7) = "Address 1": .ColWidth(7) = 2000: .ColAlignment(7) = flexAlignLeftCenter
200     .TextMatrix(0, 8) = "Address 2": .ColWidth(8) = 2000: .ColAlignment(8) = flexAlignLeftCenter
210     .TextMatrix(0, 9) = "County": .ColWidth(9) = 1200: .ColAlignment(9) = flexAlignLeftCenter
220     .TextMatrix(0, 10) = "Phone": .ColWidth(10) = 800: .ColAlignment(10) = flexAlignLeftCenter
230     .TextMatrix(0, 11) = "Fax": .ColWidth(11) = 800: .ColAlignment(11) = flexAlignLeftCenter
240     .TextMatrix(0, 12) = "GPid": .ColWidth(12) = 0: .ColAlignment(12) = flexAlignLeftCenter

250 End With
End Sub
Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then
20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub FillCountyNames()

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo FillCountyNames_Error

20  sql = "SELECT Description FROM Lists WHERE " & _
          "ListType = 'County' " & _
          "AND Code = '" & AddTicks(txtCounty) & "' " & _
          "AND Inuse = 1 ORDER BY Description"
30  Set tb = New Recordset
40  RecOpenClient 0, tb, sql

50  If Not tb.EOF Then
60      txtCounty.MaxLength = 0
70      txtCounty = tb!Description & ""
80  Else
90      txtCounty = ""
100 End If




110 Exit Sub

FillCountyNames_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmGps", "FillCountyNames", intEL, strES, sql


End Sub


Private Sub cmdExcel_Click()

10  ExportFlexGrid g, Me

End Sub

Private Sub FireDown()

    Dim n As Long
    Dim s As String
    Dim X As Long
    Dim VisibleRows As Long

10  On Error GoTo FireDown_Error

20  If g.Row = g.Rows - 1 Then Exit Sub
30  n = g.Row

40  VisibleRows = g.Height \ g.RowHeight(1) - 1

50  FireCounter = FireCounter + 1
60  If FireCounter > 5 Then
70      tmrDown.Interval = 100
80  End If

90  g.Visible = False

100 s = ""
110 For X = 0 To g.Cols - 1
120     s = s & g.TextMatrix(n, X) & vbTab
130 Next
140 s = Left$(s, Len(s) - 1)

150 g.RemoveItem n
160 If n < g.Rows Then
170     g.AddItem s, n + 1
180     g.Row = n + 1
190 Else
200     g.AddItem s
210     g.Row = g.Rows - 1
220 End If

230 For X = 0 To g.Cols - 1
240     g.Col = X
250     g.CellBackColor = vbYellow
260 Next

270 If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
280     If g.Row - VisibleRows + 1 > 0 Then
290         g.TopRow = g.Row - VisibleRows + 1
300     End If
310 End If

320 g.Visible = True

330 cmdSave.Visible = True

340 Exit Sub

FireDown_Error:

    Dim strES As String
    Dim intEL As Integer



350 intEL = Erl
360 strES = Err.Description
370 LogError "frmGps", "FireDown", intEL, strES


End Sub

Private Sub FireUp()

    Dim n As Long
    Dim s As String
    Dim X As Long

10  On Error GoTo FireUp_Error

20  If g.Row = 1 Then Exit Sub

30  FireCounter = FireCounter + 1
40  If FireCounter > 5 Then
50      tmrUp.Interval = 100
60  End If

70  n = g.Row

80  g.Visible = False

90  s = ""
100 For X = 0 To g.Cols - 1
110     s = s & g.TextMatrix(n, X) & vbTab
120 Next
130 s = Left$(s, Len(s) - 1)

140 g.RemoveItem n
150 g.AddItem s, n - 1

160 g.Row = n - 1
170 For X = 0 To g.Cols - 1
180     g.Col = X
190     g.CellBackColor = vbYellow
200 Next

210 If Not g.RowIsVisible(g.Row) Then
220     g.TopRow = g.Row
230 End If

240 g.Visible = True

250 cmdSave.Visible = True

260 Exit Sub

FireUp_Error:

    Dim strES As String
    Dim intEL As Integer



270 intEL = Erl
280 strES = Err.Description
290 LogError "frmGps", "FireUp", intEL, strES


End Sub

Private Sub cmdExit_Click()
10  Unload Me
End Sub



Private Sub cmdPrint_Click()

    Dim Y As Long

10  On Error GoTo cmdPrint_Click_Error

20  Printer.Print
30  Printer.Font.Name = "Courier New"
40  Printer.Font.Size = 12

50  Printer.Print "List of G. P.'s."

60  For Y = 0 To g.Rows - 1
70      g.Row = Y
80      g.Col = 0
90      Printer.Print g; Tab(5);
100     g.Col = 2
110     Printer.Print g; Tab(50);
120     g.Col = 8
130     Printer.Print g
140 Next

150 Printer.EndDoc



160 Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer



170 intEL = Erl
180 strES = Err.Description
190 LogError "frmGps", "cmdPrint_Click", intEL, strES


End Sub

Private Sub FillG()

    Dim tb As Recordset
    Dim sql As String
    Dim s As String

10  On Error GoTo FillG_Error

20  g.Visible = False
30  g.Rows = 2
40  g.AddItem ""
50  g.RemoveItem 1


60  sql = "SELECT GPid, " & _
          "CASE InUse WHEN 1 THEN 'Yes' ELSE 'No' END InUse, " & _
          "COALESCE(Code, '') Code, " & _
          "COALESCE(FirstName, '') FirstName, " & _
          "COALESCE(Surname, '') Surname, " & _
          "COALESCE(GPName, '') GPName, " & _
          "COALESCE(Title, '') Title, " & _
          "COALESCE(Address1, '') Address1, " & _
          "COALESCE(Address2, '') Address2, " & _
          "COALESCE(County, '') County, " & _
          "COALESCE(Phone, '') Phone, " & _
          "COALESCE(FAX, '') FAX, " & _
          "CASE HealthLink WHEN 1 THEN 'Yes' ELSE 'No' END HealthLink, " & _
          "COALESCE(MCNumber, '') MCNumber " & _
          "FROM GPs " & _
          "ORDER BY Surname "
70  Set tb = New Recordset
80  RecOpenServer 0, tb, sql

90  Do While Not tb.EOF
100     With tb
110         s = vbTab & !InUse & vbTab & _
                !Code & vbTab & _
                !GPName & vbTab & _
                !Title & vbTab & _
                !FirstName & vbTab & _
                !Surname & vbTab & _
                !Address1 & vbTab & _
                !Address2 & vbTab & _
                !County & vbTab & _
                !Phone & vbTab & _
                !FAX & vbTab & _
                !GpId
120         g.AddItem s
130     End With
140     tb.MoveNext
150 Loop

160 If g.Rows > 2 Then g.RemoveItem 1
170 g.Visible = True

180 Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmGps", "FillG", intEL, strES, sql
220 g.Visible = True

End Sub

Private Sub cmdSave_Click()

    Dim Y As Long
    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo cmdSave_Click_Error


20  pBar.Max = g.Rows - 1
30  pBar.Visible = True
40  cmdSave.Caption = "Saving..."

50  For Y = 1 To g.Rows - 1
60      pBar = Y
70      sql = "SELECT * from GPs WHERE " & _
              "GpId = '" & g.TextMatrix(Y, 12) & "' "
80      Set tb = New Recordset
90      RecOpenClient 0, tb, sql
100     If tb.EOF Then
110         tb.AddNew
120     End If
130     With tb
140         If g.TextMatrix(Y, 1) = "Yes" Then !InUse = 1 Else !InUse = 0
150         !Code = g.TextMatrix(Y, 2)
160         !GPName = g.TextMatrix(Y, 3)
170         !Address1 = initial2upper(g.TextMatrix(Y, 7))
180         !Address2 = initial2upper(g.TextMatrix(Y, 8))
190         !Title = initial2upper(g.TextMatrix(Y, 4))
200         !FirstName = initial2upper(g.TextMatrix(Y, 5))
210         !Surname = initial2upper(g.TextMatrix(Y, 6))
220         !County = initial2upper(g.TextMatrix(Y, 9))
230         !Phone = g.TextMatrix(Y, 10)
240         !FAX = g.TextMatrix(Y, 11)
250         .Update
260     End With
270 Next

280 pBar.Visible = False
290 cmdSave.Visible = False
300 cmdSave.Caption = "Save"

310 Exit Sub

cmdSave_Click_Error:

    Dim strES As String
    Dim intEL As Integer

320 intEL = Erl
330 strES = Err.Description
340 LogError "frmGps", "cmdsave_Click", intEL, strES, sql

End Sub



Private Sub cmdAdd_Click()

    Dim strSurName As String
    Dim s As String
    Dim sql As String

10  On Error GoTo cmdAdd_Click_Error

20  strSurName = Trim$(txtSurname)
30  If strSurName = "" Then
40      frmMsgBox.Msg "Enter Surname", mbOKOnly, "Histology", mbInformation
50      Exit Sub
60  End If


70  If Trim$(txtCounty) = "" Then
80      frmMsgBox.Msg "Enter County", mbOKOnly, "Histology", mbInformation
90      Exit Sub
100 End If


110 s = vbTab & "Yes" & vbTab & _
        txtCode & vbTab & _
        strFullName & vbTab & _
        txtTitle & vbTab & _
        txtForename & vbTab & _
        txtSurname & vbTab & _
        txtAddr0 & vbTab & _
        txtAddr1 & vbTab & _
        txtCounty & vbTab & _
        txtPhone & vbTab & _
        txtFAX & vbTab & _
        lblGpId

120 g.AddItem s


130 txtAddr0 = ""
140 txtAddr1 = ""
150 txtCounty = ""
160 txtTitle = ""
170 txtForename = ""
180 txtSurname = ""
190 txtPhone = ""
200 txtFAX = ""
210 lblGpId = ""
220 txtCode = ""

230 cmdSave.Visible = True

240 Exit Sub

cmdAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

250 intEL = Erl
260 strES = Err.Description
270 LogError "frmGps", "cmdAdd_Click", intEL, strES, sql

End Sub
Private Sub txtCode_LostFocus()

10  On Error GoTo txtCode_LostFocus_Error

20  txtCode = UCase$(Trim$(txtCode))

30  Exit Sub

txtCode_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer



40  intEL = Erl
50  strES = Err.Description
60  LogError "frmGps", "txtCode_LostFocus", intEL, strES


End Sub




Private Sub Form_Activate()

10  On Error GoTo Form_Activate_Error

20  If Activated Then
30      Exit Sub
40  End If

50  Activated = True

60  FillG

70  Exit Sub

Form_Activate_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmGps", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

    Dim sql As String

10  On Error GoTo Form_Load_Error

20  InitializeGrid
30  lblLoggedIn = UserName
40  If blnIsTestMode Then EnableTestMode Me
50  Exit Sub

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer



60  intEL = Erl
70  strES = Err.Description
80  LogError "frmGps", "Form_Load", intEL, strES, sql


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10  On Error GoTo Form_QueryUnload_Error

20  If cmdSave.Visible Then
30      Answer = frmMsgBox.Msg("Cancel without saving?", mbYesNo, , mbQuestion)
40      If TimedOut Then Unload Me: Exit Sub
50      If Answer = 2 Then
60          Cancel = True
70          Exit Sub
80      End If
90  End If

100 Exit Sub

Form_QueryUnload_Error:

    Dim strES As String
    Dim intEL As Integer



110 intEL = Erl
120 strES = Err.Description
130 LogError "frmGps", "Form_QueryUnload", intEL, strES


End Sub


Private Sub Form_Unload(Cancel As Integer)

10  Activated = False

End Sub


Private Sub g_Click()

    Static SortOrder As Boolean
    Dim X As Long
    Dim Y As Long
    Dim ySave As Long

10  On Error GoTo g_Click_Error

20  ySave = g.Row

30  If g.MouseRow = 0 Then
40      If SortOrder Then
50          g.Sort = flexSortGenericAscending
60      Else
70          g.Sort = flexSortGenericDescending
80      End If
90      SortOrder = Not SortOrder
100     cmdSave.Visible = True
110     Exit Sub
120 End If

130 If g.Col = 1 Then
140     g = IIf(g = "No", "Yes", "No")
150     cmdSave.Visible = True
160     Exit Sub
170 End If


180 If g.Col = 0 Then
190     g.Enabled = False
200     If frmMsgBox.Msg("Edit/Remove this line?", mbYesNo, "Histology", mbQuestion) = 1 Then

210         txtCode = g.TextMatrix(g.Row, 2)
220         txtTitle = g.TextMatrix(g.Row, 4)
230         txtForename = g.TextMatrix(g.Row, 5)
240         txtSurname = g.TextMatrix(g.Row, 6)
250         txtPhone = g.TextMatrix(g.Row, 10)
260         txtFAX = g.TextMatrix(g.Row, 11)
270         txtAddr0 = g.TextMatrix(g.Row, 7)
280         txtAddr1 = g.TextMatrix(g.Row, 8)
290         txtCounty = g.TextMatrix(g.Row, 9)
300         lblGpId = g.TextMatrix(g.Row, 12)
310         g.RemoveItem g.Row
320         cmdSave.Visible = True

330     End If
340     g.Enabled = True
350     Exit Sub
360 End If

370 If g.Col = 10 Then
380     g = iBOX("Enter Fax Number ", , g, False)
390     cmdSave.Visible = True
400     Exit Sub
410 End If

420 g.Visible = False
430 g.Col = 0
440 For Y = 1 To g.Rows - 1
450     g.Row = Y
460     If g.CellBackColor = vbYellow Then
470         For X = 0 To g.Cols - 1
480             g.Col = X
490             g.CellBackColor = 0
500         Next
510         Exit For
520     End If
530 Next
540 g.Row = ySave
550 g.Visible = True

560 For X = 0 To g.Cols - 1
570     g.Col = X
580     g.CellBackColor = vbYellow
590 Next


600 Exit Sub

g_Click_Error:

    Dim strES As String
    Dim intEL As Integer

610 intEL = Erl
620 strES = Err.Description
630 LogError "frmGps", "g_Click", intEL, strES

End Sub

Private Sub tmrDown_Timer()

10  On Error GoTo tmrDown_Timer_Error

20  FireDown

30  Exit Sub

tmrDown_Timer_Error:

    Dim strES As String
    Dim intEL As Integer



40  intEL = Erl
50  strES = Err.Description
60  LogError "frmGps", "tmrDown_Timer", intEL, strES


End Sub

Private Sub tmrUp_Timer()

10  On Error GoTo tmrUp_Timer_Error

20  FireUp

30  Exit Sub

tmrUp_Timer_Error:

    Dim strES As String
    Dim intEL As Integer



40  intEL = Erl
50  strES = Err.Description
60  LogError "frmGps", "tmrUp_Timer", intEL, strES


End Sub


Private Sub txtCounty_GotFocus()
10  txtCounty.MaxLength = 2
End Sub

Private Sub txtCounty_LostFocus()
10  FillCountyNames
End Sub

Private Sub txtForeName_Change()

10  On Error GoTo txtForeName_Change_Error

20  txtForename = Trim$(txtForename)

30  strFullName = txtTitle & " " & txtForename & " " & txtSurname

40  Exit Sub

txtForeName_Change_Error:

    Dim strES As String
    Dim intEL As Integer



50  intEL = Erl
60  strES = Err.Description
70  LogError "frmGps", "txtForeName_Change", intEL, strES


End Sub

Private Sub txtSurname_Change()

10  On Error GoTo txtSurname_Change_Error

20  txtSurname = Trim$(txtSurname)

30  strFullName = txtTitle & " " & txtForename & " " & txtSurname

40  Exit Sub

txtSurname_Change_Error:

    Dim strES As String
    Dim intEL As Integer



50  intEL = Erl
60  strES = Err.Description
70  LogError "frmGps", "txtSurname_Change", intEL, strES


End Sub


Private Sub txtTitle_Change()

10  On Error GoTo txtTitle_Change_Error

20  txtTitle = Trim$(txtTitle)

30  strFullName = txtTitle & " " & txtForename & " " & txtSurname

40  Exit Sub

txtTitle_Change_Error:

    Dim strES As String
    Dim intEL As Integer



50  intEL = Erl
60  strES = Err.Description
70  LogError "frmGps", "txtTitle_Change", intEL, strES


End Sub



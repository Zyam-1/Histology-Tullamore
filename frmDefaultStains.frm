VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDefaultStains 
   Caption         =   "Default Stains"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   5400
      Picture         =   "frmDefaultStains.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   855
      Left            =   4680
      Picture         =   "frmDefaultStains.frx":032A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6800
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   6000
      Picture         =   "frmDefaultStains.frx":0668
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   6000
      Picture         =   "frmDefaultStains.frx":093E
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblTissueType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4170
   End
End
Attribute VB_Name = "frmDefaultStains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pTissueCode As String
Private pTissueName As String
Private pTissueCodeListId As String


Private Sub InitializeGrid()

10  With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 3: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.Name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "Stain": .ColWidth(0) = 4300: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Add": .ColWidth(1) = 850: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "StainCode": .ColWidth(2) = 0: .ColAlignment(2) = flexAlignLeftCenter

150 End With
End Sub

Private Sub cmdAdd_Click()
    Dim sql As String
    Dim tb As New Recordset
    Dim r As Integer

10  sql = "DELETE FROM DefaultStains " & _
          "WHERE TissueCode = '" & pTissueCode & "'"
20  Cnxn(0).Execute sql

30  sql = "SELECT * FROM DefaultStains " & _
          "WHERE TissueCode = '" & pTissueCode & "'"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql


60  Do Until r = g.Rows - 1
70      g.Row = r
80      g.Col = 1
90      If g.CellPicture = imgSquareTick.Picture Then
100         tb.AddNew
110         tb!TissueCodeListId = pTissueCodeListId
120         tb!TissueCode = pTissueCode
130         tb!StainCode = g.TextMatrix(r, 2)
140         tb!UserName = UserName
150         tb.Update
160     End If

170     r = r + 1
180 Loop

190 Unload Me

End Sub

Private Sub cmdExit_Click()
10  Unload Me
End Sub

Private Sub Form_Load()

10  lblTissueType = pTissueCode & " " & pTissueName
20  InitializeGrid
30  FillGrid
40  If blnIsTestMode Then EnableTestMode Me
End Sub
Private Sub FillGrid()
    Dim sql As String
    Dim tb As New Recordset
    Dim s As String
    Dim r As Integer

10  sql = "SELECT * FROM Lists WHERE " & _
          "ListType = 'SS' OR ListType = 'IS' OR ListType = 'RS' " & _
          "ORDER BY listtype"



20  Set tb = New Recordset
30  RecOpenServer 0, tb, sql
40  Do While Not tb.EOF

50      s = tb!Description & ""

60      g.AddItem s
70      g.Row = g.Rows - 1


80      g.Col = 1
90      g.CellPictureAlignment = flexAlignCenterCenter
100     Set g.CellPicture = imgSquareCross.Picture

110     g.TextMatrix(g.Row, 2) = tb!Code & ""

120     tb.MoveNext
130 Loop

140 sql = "SELECT * FROM DefaultStains " & _
          "WHERE TissueCode = '" & pTissueCode & "'"

150 Set tb = New Recordset
160 RecOpenServer 0, tb, sql

170 Do While Not tb.EOF
180     Do Until r = g.Rows - 1
190         g.Row = r
200         g.Col = 1
210         g.CellPictureAlignment = flexAlignCenterCenter
220         If g.TextMatrix(r, 2) = tb!StainCode & "" Then
230             Set g.CellPicture = imgSquareTick.Picture
240             Exit Do
250         End If
260         r = r + 1
270     Loop
280     tb.MoveNext
290 Loop


300 If g.Rows > 2 Then
310     g.RemoveItem 1
320 End If
End Sub

Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized And Me.Top >= 0 Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Public Property Let TissueCode(ByVal strNewValue As String)

10  pTissueCode = strNewValue

End Property
Public Property Let TissueCodeListId(ByVal strNewValue As String)

10  pTissueCodeListId = strNewValue

End Property

Public Property Let TissueName(ByVal strNewValue As String)

10  pTissueName = strNewValue

End Property

Private Sub Form_Unload(Cancel As Integer)
    Dim sql As String
    Dim sn As Recordset

10  On Error GoTo Form_Unload_Error

20  sql = "SELECT Count(StainCode) as TotStains FROM DefaultStains WHERE TissueCodeListId = '" & pTissueCodeListId & "'"
30  Set sn = New Recordset
40  RecOpenServer 0, sn, sql

50  If Not sn.EOF Then
60      frmListsGeneric.g.TextMatrix(frmListsGeneric.g.Row, 4) = sn!TotStains & ""
70  End If

80  Exit Sub

Form_Unload_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmDefaultStains", "Form_Unload", intEL, strES, sql


End Sub

Private Sub g_Click()

10  If g.ColWidth(1) <> 0 Then
20      If g.Col = 1 Then
30          If g.CellPicture = imgSquareTick.Picture Then
40              Set g.CellPicture = imgSquareCross.Picture
50          Else
60              Set g.CellPicture = imgSquareTick.Picture
70          End If
80      End If
90  End If

End Sub

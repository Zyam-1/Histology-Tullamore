VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGPConflict 
   Caption         =   "GPs"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   975
      Width           =   4665
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblAddress2 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   780
         TabIndex        =   7
         Top             =   600
         Width           =   3675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   570
      End
      Begin VB.Label lblAddress1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   780
         TabIndex        =   4
         Top             =   270
         Width           =   3675
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   855
      Left            =   3480
      Picture         =   "frmGPConflict.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   4200
      Picture         =   "frmGPConflict.frx":033E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid grdGPConflict 
      Height          =   1155
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2037
      _Version        =   393216
      Cols            =   3
      RowHeightMin    =   315
      BackColorBkg    =   -2147483648
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   35
      ImageHeight     =   35
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGPConflict.frx":0668
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGPConflict.frx":09B6
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblGP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   2970
   End
End
Attribute VB_Name = "frmGPConflict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pGPName As String

Private Sub InitializeGrid()

  
10        With grdGPConflict
20            .Clear
30            .Rows = 2: .FixedRows = 1
40            .Cols = 3: .FixedCols = 0
50            .Rows = 1
60            .Font.Size = fgcFontSize
70            .Font.Name = fgcFontName
80            .ForeColor = fgcForeColor
90            .BackColor = fgcBackColor
100           .ForeColorFixed = fgcForeColorFixed
110           .BackColorFixed = fgcBackColorFixed
120           .ScrollBars = flexScrollBarBoth
              '<Patient Name                 |<Date Of Birth
130           .TextMatrix(0, 0) = "Address 1": .ColWidth(0) = 2250: .ColAlignment(0) = flexAlignLeftCenter
140           .TextMatrix(0, 1) = "Address 2": .ColWidth(1) = 2250: .ColAlignment(1) = flexAlignLeftCenter
150           .TextMatrix(0, 2) = "GPId": .ColWidth(2) = 0: .ColAlignment(2) = flexAlignLeftCenter

160       End With
End Sub

Private Sub cmdAdd_Click()
10    AddDemo

End Sub

Private Sub cmdExit_Click()
10    Unload Me
End Sub

Private Sub Form_Load()
10    InitializeGrid
20    FillGrid

End Sub

Private Sub Form_Resize()
10    If Me.WindowState <> vbMinimized Then
20        Me.Top = 0
30        Me.Left = Screen.Width / 2 - Me.Width / 2
40    End If
End Sub

Private Sub FillGrid()
      Dim sql As String
      Dim sn As New Recordset
      Dim s As String

10    sql = "SELECT Address1, Address2, GPId FROM GPs WHERE GPName = '" & AddTicks(pGPName) & "'"
20    Set sn = New Recordset
30    RecOpenServer 0, sn, sql

40    Do Until sn.EOF
50        s = sn!Address1 & vbTab & sn!Address2 & vbTab & sn!GpId
60        grdGPConflict.AddItem s
70        sn.MoveNext
80    Loop

End Sub

Private Sub AddDemo()


10    If lblId <> "" Then
20        With frmDemographics
30            .txtGP = lblGP
40            .lblGpId = lblId
50        End With
60    Else
70        frmMsgBox.Msg "Please select an address", mbOKOnly, , mbInformation
80        Exit Sub
90    End If
100   Unload Me
End Sub

Private Sub grdGPConflict_Click()
10    If grdGPConflict.Rows > 1 Then
20        Rada = grdGPConflict.MouseRow
30        lblAddress1 = grdGPConflict.TextMatrix(Rada, 0)
40        lblAddress2 = grdGPConflict.TextMatrix(Rada, 1)
50        lblId = grdGPConflict.TextMatrix(Rada, 2)
60    End If
End Sub

Private Sub grdGPConflict_KeyPress(KeyAscii As Integer)


10    If KeyAscii = 13 Then
20        If grdGPConflict.Rows > 1 Then
30            Rada = grdGPConflict.RowSel
40            lblAddress1 = grdGPConflict.TextMatrix(Rada, 0)
50            lblAddress2 = grdGPConflict.TextMatrix(Rada, 1)
60            lblId = grdGPConflict.TextMatrix(Rada, 2)
70        End If
80    End If
End Sub

Public Property Let GPName(ByVal Id As String)

10    pGPName = Id
End Property


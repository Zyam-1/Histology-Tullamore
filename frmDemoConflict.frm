VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmDemoConflict 
   Caption         =   "Demographics Conflict"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   4080
      Picture         =   "frmDemoConflict.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   855
      Left            =   3360
      Picture         =   "frmDemoConflict.frx":032A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Index           =   3
      Left            =   120
      TabIndex        =   0
      Top             =   975
      Width           =   4665
      Begin VB.Label lDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   780
         TabIndex        =   4
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label lName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   780
         TabIndex        =   3
         Top             =   270
         Width           =   3705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   2
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   1
         Top             =   300
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdDemoConflict 
      Height          =   1155
      Left            =   120
      TabIndex        =   6
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
      Left            =   4920
      Top             =   480
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
            Picture         =   "frmDemoConflict.frx":0668
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDemoConflict.frx":09B6
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
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
   Begin VB.Label lblNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   375
      TabIndex        =   5
      Top             =   360
      Width           =   2850
   End
End
Attribute VB_Name = "frmDemoConflict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pIDNumber As String

Private Sub InitializeGrid()

With grdDemoConflict
    .Clear
    .Rows = 2: .FixedRows = 1
    .Cols = 2: .FixedCols = 0
    .Rows = 1
    .Font.Size = fgcFontSize
    .Font.name = fgcFontName
    .ForeColor = fgcForeColor
    .BackColor = fgcBackColor
    .ForeColorFixed = fgcForeColorFixed
    .BackColorFixed = fgcBackColorFixed
    .ScrollBars = flexScrollBarBoth
    '<Patient Name                 |<Date Of Birth
    .TextMatrix(0, 0) = "Patient Name"
    .TextMatrix(0, 1) = "Date Of Birth"
    .ColWidth(0) = 750
    .ColWidth(1) = 750

End With
End Sub

Private Sub cmdAdd_Click()



AddDemo

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.Caption = "Demographics Conflict"
Me.Label4(3).Caption = ""
'Me.Label2(2).Caption = ""
Me.cmdAdd.Caption = "Add"
Me.cmdExit.Caption = "Exit"
End Sub

Private Sub Form_Load()

ChangeFont Me, "Arial"
InitializeGrid
FillGrid

End Sub
Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
    Me.Top = 0
    Me.Left = Screen.Width / 2 - Me.Width / 2
End If
End Sub


Private Sub FillGrid()
Dim sql As String
Dim sn As New Recordset
Dim s As String

sql = "SELECT DISTINCT PatientName, DateOfBirth FROM Demographics " & _
      "WHERE MRN = '" & pIDNumber & "' "
Set sn = New Recordset
RecOpenServer 0, sn, sql

Do Until sn.EOF
    s = IIf(sn!PatientName = Null, "", sn!PatientName) & vbTab & Format(sn!DateOfBirth, "dd/mm/yyyy")
    grdDemoConflict.AddItem s
    sn.MoveNext
Loop

End Sub
Public Property Let IDNumber(ByVal Id As String)

pIDNumber = Id
End Property

Private Sub AddDemo()
Dim sql As String
Dim tb As New Recordset

sql = "SELECT TOP 1 * FROM Demographics WHERE MRN = '" & pIDNumber & "' " & _
      "AND PatientName = '" & lName & "' " & _
      "AND DateOfBirth = '" & Format(lDoB, "mm/dd/yyyy") & "' " & _
      "ORDER BY DateTimeOfRecord DESC"

Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
    With frmDemographics
        .txtFirstName = tb!FirstName & ""
        .txtSurname = tb!Surname & ""
        .txtAddress1 = tb!Address1 & ""
        .txtAddress2 = tb!Address2 & ""
        .txtAddress3 = tb!Address3 & ""
        '.txtAddress4 = tb!Address4 & ""
        .txtCounty.MaxLength = 0
        .txtCounty = tb!County & ""
        If Not IsNull(tb!DateOfBirth) Then
            .txtDOB = Format(tb!DateOfBirth, "dd/mm/yyyy")
        End If
        .txtAge = tb!Age & ""
        Select Case Trim(tb!Sex) & ""
        Case "F": .txtSex = "Female"
        Case "M": .txtSex = "Male"
        Case Else: .txtSex = ""
        End Select
        .txtPhone = tb!Phone & ""

        .txtPatComments = tb!Comments & ""
        .chkUrgent.Value = 0
        .Link = False
    End With
End If
Unload Me
End Sub

Private Sub grdDemoConflict_Click()
If grdDemoConflict.Rows > 1 Then
    Rada = grdDemoConflict.MouseRow
    lName = grdDemoConflict.TextMatrix(Rada, 0)
    lDoB = grdDemoConflict.TextMatrix(Rada, 1)

End If
End Sub

Private Sub grdDemoConflict_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grdDemoConflict.Rows > 1 Then
        Rada = grdDemoConflict.RowSel
        lName = grdDemoConflict.TextMatrix(Rada, 0)
        lDoB = grdDemoConflict.TextMatrix(Rada, 1)
    End If
End If


End Sub

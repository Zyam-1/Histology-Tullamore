VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCopyTo 
   Caption         =   "Copy To Clinician/GP"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtClinician 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   870
      Left            =   7320
      Picture         =   "frmCopyTo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Height          =   870
      Left            =   7320
      Picture         =   "frmCopyTo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdTransferToPrimary 
      Height          =   555
      Left            =   3240
      Picture         =   "frmCopyTo.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   765
   End
   Begin VB.CommandButton cmdRemoveFromPrimary 
      Height          =   555
      Left            =   3240
      Picture         =   "frmCopyTo.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   765
   End
   Begin VB.Frame fraCopyTo 
      Caption         =   "Copy To"
      Height          =   3195
      Left            =   4320
      TabIndex        =   3
      Top             =   480
      Width           =   2715
      Begin VB.ListBox lstCopyTo 
         Height          =   2580
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   2385
      End
   End
   Begin VB.Frame fraConsultants 
      Caption         =   "Clinicians/GPs"
      Height          =   3195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2715
      Begin VB.ListBox lstConsultants 
         Height          =   2580
         Left            =   120
         TabIndex        =   2
         Top             =   330
         Width           =   2385
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
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
      TabIndex        =   11
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "frmCopyTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
      Dim Num As Long

10    For Num = 0 To lstCopyTo.ListCount - 1
20        If Num > 0 Then
30            frmDemographics.lblCopyTo = frmDemographics.lblCopyTo & vbCrLf & lstCopyTo.List(Num)
40        Else
50            frmDemographics.lblCopyTo = lstCopyTo.List(Num)
60        End If
70    Next
80    If lstCopyTo.ListCount = 0 Then
90        frmDemographics.lblCopyTo = ""
100   End If
110   Unload Me
End Sub

Private Sub cmdExit_Click()
10    Unload Me
End Sub

Private Sub cmdRemoveFromPrimary_Click()
      Dim Num As Long

10    For Num = 0 To lstCopyTo.ListCount - 1
20      If lstCopyTo.Selected(Num) Then
30        lstConsultants.AddItem lstCopyTo.List(Num)
40        lstCopyTo.RemoveItem Num
50        Exit For
60      End If
70    Next


End Sub

Private Sub cmdTransferToPrimary_Click()
       Dim Num As Long

10    For Num = 0 To lstConsultants.ListCount - 1
20      If lstConsultants.Selected(Num) Then
30        lstCopyTo.AddItem lstConsultants.List(Num)
40        lstConsultants.RemoveItem Num
50        Exit For
60      End If
70    Next


End Sub

Private Sub Form_Load()
10    FillConsultants

20    FillCopyTo
30    lblLoggedIn = UserName
End Sub

Private Sub FillConsultants()
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillConsultants_Error

20    sql = "SELECT * FROM SourceItemLists " & _
              "WHERE Listtype = 'Clinician' "
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    Do While Not tb.EOF
60        lstConsultants.AddItem tb!Description
70        tb.MoveNext
80    Loop

90    sql = "SELECT * FROM GPs "
100   Set tb = New Recordset
110   RecOpenClient 0, tb, sql
    
120   Do While Not tb.EOF
130       lstConsultants.AddItem tb!GPName
140       tb.MoveNext
150   Loop

160   Exit Sub

FillConsultants_Error:

Dim strES As String
Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmCopyTo", "FillConsultants", intEL, strES, sql

End Sub

Private Sub ReloadConsultants()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo ReloadConsultants_Error

20    lstConsultants.Clear

30    sql = "SELECT Description FROM SourceItemLists WHERE " & _
      "ListType = 'Clinician' " & _
      "AND Description LIKE '%" & AddTicks(txtClinician) & "%' "

40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql

60    Do While Not tb.EOF
70        lstConsultants.AddItem tb!Description & ""
80        tb.MoveNext
90    Loop

100   sql = "SELECT GPName FROM GPs WHERE GPName LIKE '%" & AddTicks(txtClinician) & "%' "
110   Set tb = New Recordset
120   RecOpenClient 0, tb, sql
    
130   Do While Not tb.EOF
140       lstConsultants.AddItem tb!GPName
150       tb.MoveNext
160   Loop

170   Exit Sub

ReloadConsultants_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmCopyTo", "ReloadConsultants", intEL, strES, sql


End Sub

Private Sub FillCopyTo()
      Dim strArray() As String
      Dim itm As Variant
      Dim Num As Integer

10    strArray = Split(frmDemographics.lblCopyTo, vbCrLf)

20    For Each itm In strArray
30        lstCopyTo.AddItem itm
40        For Num = 0 To lstConsultants.ListCount - 1
50          If lstConsultants.List(Num) = itm Then
60            lstConsultants.RemoveItem Num
70            Exit For
80          End If
90        Next
100   Next itm



End Sub
Private Sub Form_Resize()
10    If Me.WindowState <> vbMinimized Then
20        Me.Top = 0
30        Me.Left = Screen.Width / 2 - Me.Width / 2
40    End If
End Sub

Private Sub txtClinician_Change()
10    If txtClinician.Text <> "" Then
20        ReloadConsultants
30    Else
40        FillConsultants
50    End If
End Sub

Private Sub txtClinician_KeyUp(KeyCode As Integer, Shift As Integer)

10    If KeyCode = vbKeyDown Then
20        If lstConsultants.ListCount > 0 Then
30            lstConsultants.SetFocus
40            lstConsultants.Selected(0) = True
50        End If
60    End If
End Sub

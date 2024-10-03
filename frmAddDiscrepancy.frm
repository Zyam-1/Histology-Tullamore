VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAddDiscrepancy 
   Caption         =   "Discrepancy Log"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   615
      Left            =   1080
      Picture         =   "frmAddDiscrepancy.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   1920
      Picture         =   "frmAddDiscrepancy.frx":01CF
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   615
      Left            =   240
      Picture         =   "frmAddDiscrepancy.frx":0511
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   6825
      Left            =   225
      TabIndex        =   0
      Top             =   990
      Width           =   5925
      Begin VB.TextBox txtResolutionDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         MaxLength       =   10
         TabIndex        =   24
         Top             =   6270
         Width           =   1455
      End
      Begin VB.TextBox txtNatureDiscrepancy 
         Height          =   285
         Left            =   375
         TabIndex        =   5
         Top             =   3180
         Width           =   5295
      End
      Begin VB.ComboBox cmbResolution 
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Top             =   5640
         Width           =   5295
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   3975
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1665
      End
      Begin VB.ComboBox cmbOperator 
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   4800
         Width           =   5295
      End
      Begin VB.TextBox txtResponsiblePerson 
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   3960
         Width           =   5295
      End
      Begin VB.ComboBox cmbDiscrepancyType 
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   2400
         Width           =   5295
      End
      Begin VB.TextBox txtPatientName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox txtCaseId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Resolution Date"
         Height          =   285
         Left            =   375
         TabIndex        =   25
         Top             =   6015
         Width           =   1875
      End
      Begin VB.Label L 
         Caption         =   "Nature Of Discrepancy / Corrective Action"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   19
         Top             =   2880
         Width           =   3945
      End
      Begin VB.Label L 
         Caption         =   "Resolution"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   18
         Top             =   5400
         Width           =   2505
      End
      Begin VB.Label L 
         Caption         =   "Case Id"
         Height          =   255
         Index           =   0
         Left            =   350
         TabIndex        =   17
         Top             =   360
         Width           =   555
      End
      Begin VB.Label L 
         Caption         =   "Discrepancy Type"
         Height          =   255
         Index           =   1
         Left            =   345
         TabIndex        =   16
         Top             =   2115
         Width           =   1725
      End
      Begin VB.Label L 
         Caption         =   "Responsible Person"
         Height          =   255
         Index           =   2
         Left            =   345
         TabIndex        =   15
         Top             =   3660
         Width           =   1665
      End
      Begin VB.Label L 
         Caption         =   "Person Dealing with Discrepancy"
         Height          =   255
         Index           =   3
         Left            =   345
         TabIndex        =   14
         Top             =   4485
         Width           =   2505
      End
      Begin VB.Label Label1 
         Caption         =   "Date of Discrepancy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3990
         TabIndex        =   13
         Top             =   330
         Width           =   1845
      End
      Begin VB.Label L 
         Caption         =   "Patient Name"
         Height          =   255
         Index           =   4
         Left            =   350
         TabIndex        =   12
         Top             =   1260
         Width           =   1725
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblCaseLocked 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   2640
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   22
      Top             =   7440
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   7995
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddDiscrepancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pCaseId As String
Private pPatientName As String

Public Property Let CaseId(ByVal Id As String)

10    pCaseId = Id
End Property


Public Property Let PatientName(ByVal Id As String)

10    pPatientName = Id
End Property

Private Sub cmbResolution_Click()
10    If Len(cmbResolution) > 0 Then
20      txtResolutionDate.Enabled = True
30      If Len(txtResolutionDate) = 0 Then
40          txtResolutionDate = Format(Now, "dd/mm/yyyy")
50      End If
60    End If
End Sub

Private Sub cmdExit_Click()
10    Unload Me
End Sub

Private Sub cmdSave_Click()
10    If txtDate <> "" Then
20      SaveDiscrepancy
30      Unload Me
40    Else
50      iMsg "No Date set!"
60    End If

End Sub

Private Sub Form_Load()
10    FillOperator
20    FillLists
30    txtDate = Format(Now, "DD/MM/YYYY")
40    cmbOperator = UserName
50    LoadDiscrepancy
60    lblLoggedIn = UserName

70    If bLocked Then
80      DisableForm
90      cmdClear.Enabled = False
100     cmdSave.Enabled = False
110     lblCaseLocked.Visible = True
120     lblCaseLocked = "RECORD BEING EDITED BY " & sCaseLockedBy
130     lblCaseLocked.BackColor = &H8080FF
140   End If
150   If blnIsTestMode Then EnableTestMode Me
End Sub
Private Sub DisableForm()
10    txtCaseId.Enabled = False
20    txtDate.Enabled = False
30    txtPatientName.Enabled = False
40    cmbDiscrepancyType.Enabled = False
50    txtNatureDiscrepancy.Enabled = False
60    txtResponsiblePerson.Enabled = False
70    cmbOperator.Enabled = False
80    cmbResolution.Enabled = False
End Sub

Private Sub LoadDiscrepancy()
    Dim sql As String
    Dim tb As New Recordset

10    On Error GoTo LoadDiscrepancy_Error

20    txtCaseId = pCaseId
30    txtPatientName = pPatientName

40    sql = "SELECT * FROM Discrepancy WHERE CaseId = '" & CaseNo & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80      If Not IsNull(tb!DateOfDiscrepancy) Then
90          txtDate = Format(tb!DateOfDiscrepancy, "DD/MM/YYYY")
100     Else
110         txtDate = ""
120     End If
130     cmbDiscrepancyType = tb!DiscrepancyType & ""
140     txtNatureDiscrepancy = tb!NatureOfDiscrepancy & ""
150     txtResponsiblePerson = tb!PersonResponsible & ""
160     cmbOperator = tb!PersonDealingWith & ""
170     cmbResolution = tb!Resolution & ""
180     If Len(cmbResolution) > 0 Then
190         txtResolutionDate.Enabled = True
200     End If
210     If Not IsNull(tb!DateOfResolution) Then
220         txtResolutionDate = Format(tb!DateOfResolution, "dd/mm/yyyy")
230     End If
240   End If

250   Exit Sub

LoadDiscrepancy_Error:

    Dim strES As String
    Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmAddDiscrepancy", "LoadDiscrepancy", intEL, strES, sql

End Sub

Private Sub SaveDiscrepancy()
    Dim sql As String
    Dim tb As New Recordset

10    On Error GoTo SaveDiscrepancy_Error


20    sql = "SELECT * FROM Discrepancy WHERE CaseId = '" & CaseNo & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      tb.AddNew

70      CaseAddLogEvent CaseNo, DiscrepancyAdded
80    ElseIf tb!DateOfDiscrepancy <> txtDate Or _
           tb!DiscrepancyType <> cmbDiscrepancyType Or _
           tb!NatureOfDiscrepancy <> txtNatureDiscrepancy Or _
           tb!PersonResponsible <> txtResponsiblePerson Or _
           tb!PersonDealingWith <> cmbOperator Or _
           tb!Resolution <> cmbResolution Then
90      CaseAddLogEvent CaseNo, DiscrepancyEdited

100   End If

110   tb!CaseId = CaseNo
120   tb!DateOfDiscrepancy = Format(txtDate, "dd/mmm/yyyy")
130   tb!DiscrepancyType = cmbDiscrepancyType
140   tb!NatureOfDiscrepancy = txtNatureDiscrepancy
150   tb!PersonResponsible = txtResponsiblePerson
160   tb!PersonDealingWith = cmbOperator
170   If Len(cmbResolution) > 0 Then
180     tb!Resolution = cmbResolution
190     tb!DateOfResolution = Format(txtResolutionDate, "dd/mmm/yyyy")
200   End If
210   tb!UserName = UserName
220   tb.Update

230   Exit Sub

SaveDiscrepancy_Error:

    Dim strES As String
    Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmAddDiscrepancy", "SaveDiscrepancy", intEL, strES, sql


End Sub

Private Sub FillLists()
    Dim sql As String
    Dim tb As Recordset

10    On Error GoTo FillLists_Error

20    cmbDiscrepancyType.AddItem ""
30    sql = "SELECT * FROM Lists WHERE ListType = 'DiscrepType'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    Do While Not tb.EOF
70      cmbDiscrepancyType.AddItem tb!Description & ""
80      tb.MoveNext
90    Loop

100   cmbResolution.AddItem ""
110   sql = "SELECT * FROM Lists WHERE ListType = 'DiscrepRes'"
120   Set tb = New Recordset
130   RecOpenServer 0, tb, sql

140   Do While Not tb.EOF
150     cmbResolution.AddItem tb!Description & ""
160     tb.MoveNext
170   Loop

180   Exit Sub

FillLists_Error:

    Dim strES As String
    Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmAddDiscrepancy", "FillLists", intEL, strES, sql

End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
10    KeyAscii = VI(KeyAscii, NumericSlash)
End Sub

Private Sub txtDate_LostFocus()
10    txtDate = Convert62Date(txtDate, BACKWARD)

20    If Len(txtDate) = 8 And Not IsDate(txtDate) Then
30      txtDate = Left(txtDate, 2) & "/" & Mid(txtDate, 3, 2) & "/" & Right(txtDate, 4)
40    End If
50    If Not IsDate(txtDate) Then
60      txtDate = ""
70      Exit Sub
80    End If

90    If Format$(txtDate, "yyyymmdd") > Format$(Now, "yyyymmdd") Then
100     txtDate = ""
110     Exit Sub
120   End If

End Sub

Private Sub Form_Resize()
10    If Me.WindowState <> vbMinimized Then
20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40    End If
End Sub

Private Sub cmbOperator_KeyPress(KeyAscii As Integer)
10    KeyAscii = 0
End Sub
Private Sub FillOperator()

    Dim sql As String
    Dim tb As Recordset

10    On Error GoTo FillOperator_Error

20    sql = "SELECT Username FROM Users"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    Do While Not tb.EOF
60      cmbOperator.AddItem tb!UserName & ""
70      tb.MoveNext
80    Loop
90    tb.Close

100   Exit Sub

FillOperator_Error:

    Dim strES As String
    Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmAddDiscrepancy", "FillOperator", intEL, strES, sql

End Sub


Private Sub txtResolutionDate_KeyPress(KeyAscii As Integer)
10    KeyAscii = VI(KeyAscii, NumericSlash)
End Sub

Private Sub txtResolutionDate_LostFocus()

10    txtResolutionDate = Convert62Date(txtResolutionDate, BACKWARD)

20    If Len(txtResolutionDate) = 8 And Not IsDate(txtResolutionDate) Then
30      txtResolutionDate = Left(txtResolutionDate, 2) & "/" & Mid(txtResolutionDate, 3, 2) & "/" & Right(txtResolutionDate, 4)
40    End If
50    If Not IsDate(txtResolutionDate) Then
60      txtResolutionDate = ""
70      Exit Sub
80    End If

90    If Format$(txtResolutionDate, "yyyymmdd") > Format$(Now, "yyyymmdd") Then
100     txtResolutionDate = ""
110     Exit Sub
120   End If

End Sub

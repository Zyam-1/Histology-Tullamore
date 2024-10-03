VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmSystemManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   9375
   ClientLeft      =   2550
   ClientTop       =   390
   ClientWidth     =   8970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmSystemManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9375
   ScaleWidth      =   8970
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7875
      Picture         =   "frmSystemManager.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5790
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7815
      Picture         =   "frmSystemManager.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   7815
      Picture         =   "frmSystemManager.frx":1B5B
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8535
      Width           =   975
   End
   Begin VB.ComboBox cmbUserName 
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   2505
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3495
      Left            =   210
      TabIndex        =   14
      Top             =   5775
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   7
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Operator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   225
      TabIndex        =   11
      Top             =   3135
      Width           =   7410
      Begin VB.TextBox txtCode 
         DataField       =   "opcode"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5040
         MaxLength       =   5
         TabIndex        =   7
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txtMCRN 
         Height          =   285
         Left            =   5040
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   5040
         TabIndex        =   4
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtAutoLogOff 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1860
         TabIndex        =   10
         Text            =   "5"
         Top             =   2100
         Width           =   495
      End
      Begin VB.ComboBox cmbMemberOf 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1650
         Width           =   2205
      End
      Begin VB.TextBox txtConfirm 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1860
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtForename 
         DataField       =   "opname"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   3
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtPass 
         DataField       =   "oppass"
         DataSource      =   "Data1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1860
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   750
         Width           =   2175
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2355
         TabIndex        =   18
         Top             =   2100
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "txtAutoLogOff"
         BuddyDispid     =   196617
         OrigLeft        =   2400
         OrigTop         =   2130
         OrigRight       =   2895
         OrigBottom      =   2370
         Max             =   999
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
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
         Left            =   4545
         TabIndex        =   26
         Top             =   780
         Width           =   375
      End
      Begin VB.Label lblMCRN 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MCRN"
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
         Left            =   4425
         TabIndex        =   24
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   4275
         TabIndex        =   21
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
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
         Left            =   2850
         TabIndex        =   19
         Top             =   2160
         Width           =   1770
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Auto Log Off in"
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
         Left            =   105
         TabIndex        =   17
         Top             =   2130
         Width           =   1725
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Access Level"
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
         Left            =   870
         TabIndex        =   16
         Top             =   1710
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
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
         Left            =   105
         TabIndex        =   15
         Top             =   1140
         Width           =   1725
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   1125
         TabIndex        =   13
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   105
         TabIndex        =   12
         Top             =   780
         Width           =   1725
      End
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1200
      TabIndex        =   20
      Top             =   240
      Width           =   5250
      Begin VB.CommandButton cmdHide 
         Caption         =   "E&xit"
         Height          =   735
         Left            =   3000
         Picture         =   "frmSystemManager.frx":1E9D
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   885
         Width           =   2475
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Login"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1410
         Picture         =   "frmSystemManager.frx":21DF
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1530
         Width           =   1200
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "User Name"
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
         Left            =   600
         TabIndex        =   23
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   855
         TabIndex        =   22
         Top             =   930
         Width           =   915
      End
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   9360
      Top             =   690
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   9150
      Top             =   690
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmSystemManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private temp As String

Private mOperator As Boolean
Private mManager As Boolean
Private mAdministrator As Boolean
Private mSecretary As Boolean
Private mSysManager As Boolean
Private LoginCount As Long

Private AlphaOrderTechnicians As Boolean



Public Property Let Administrator(ByVal ShowAdministrator As Boolean)

On Error GoTo Administrator_Error

mAdministrator = ShowAdministrator

Exit Property

Administrator_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "Administrator", intEL, strES


End Property
Public Property Let SysManager(ByVal ShowSysManager As Boolean)

On Error GoTo SysManager_Error

mSysManager = ShowSysManager

Exit Property

SysManager_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "SysManager", intEL, strES


End Property



Private Sub UpdateLoggedOnUser()

    Dim tb As Recordset
    Dim sql As String
    Dim MachineName As String

On Error GoTo UpdateLoggedOnUser_Error

MachineName = UCase$(vbGetComputerName())

sql = "SELECT * FROM LoggedOnUsers WHERE " & _
    "MachineName = N'" & MachineName & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
End If
tb!MachineName = MachineName
tb!AppName = "Histology"
tb!UserName = UserName
tb.Update

Exit Sub

UpdateLoggedOnUser_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "UpdateLoggedOnUser", intEL, strES, sql

End Sub

Private Sub cmbMemberOf_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub cmbUserName_Click()

On Error GoTo cmbUserName_Click_Error

txtPassword = ""

Exit Sub

cmbUserName_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "cmbUserName_Click", intEL, strES


End Sub

Private Sub cmbUserName_LostFocus()

    Dim sql As String
    Dim tb As New Recordset

On Error GoTo cmbUserName_LostFocus_Error

If Trim(cmbUserName) = "" Then Exit Sub

cmbUserName = initial2upper(cmbUserName)

sql = "SELECT * FROM users WHERE username like N'" & AddTicks(cmbUserName) & "%'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  frmMsgBox.Msg "Username : " & cmbUserName & " is Incorrect", , , mbExclamation
  If TimedOut Then Unload Me: Exit Sub
  cmbUserName = ""
  cmbUserName.SetFocus
Else
  cmbUserName = tb!UserName
End If

txtPassword = ""
Exit Sub

cmbUserName_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "cmbUserName_LostFocus", intEL, strES, sql

End Sub

Private Sub cmdEdit_Click()
    Dim Y As Integer
    Dim tb As New Recordset
    Dim sql As String
    Dim AdminCount As Integer

If cmdEdit.Caption = "&Edit" Then

  For Y = 1 To g.Rows - 1
      g.row = Y
      If g.CellBackColor = vbYellow Then
          Exit For
      End If
  Next

  sql = "SELECT Forename, Surname, MCRN, PassWord, UserName, AccessLevel, LogOffDelay, PassDate, Code, UserId " & _
        "From Users where UserId = N'" & g.TextMatrix(Y, 6) & "'"

  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then
      txtForename = tb!ForeName & ""
      txtSurname = tb!Surname & ""
      txtCode = tb!Code & ""
      txtMCRN = tb!MCRN & ""
      cmbMemberOf = tb!AccessLevel & ""
      txtAutoLogOff = tb!LogOffDelay & ""
  End If

  If g.TextMatrix(Y, 1) = "Administrator" Then    'if changing "Administrator" users
      AdminCount = 0
      For Y = 1 To g.Rows - 1
          If g.TextMatrix(Y, 1) = "Administrator" Then
              AdminCount = AdminCount + 1
          End If
      Next
  End If

  If AdminCount = 1 Then
      cmbMemberOf.Enabled = False
  Else
      cmbMemberOf.Enabled = True
  End If

  txtForename.BackColor = &H80C0FF
  txtSurname.BackColor = &H80C0FF
  txtCode.BackColor = &H80C0FF
  txtMCRN.BackColor = &H80C0FF
  cmbMemberOf.BackColor = &H80C0FF
  txtAutoLogOff.BackColor = &H80C0FF
  txtPass.BackColor = &H80C0FF
  txtConfirm.BackColor = &H80C0FF
  Frame1.Caption = "Edit Operator"

  cmdEdit.Caption = "&Cancel Edit"

  txtCode.Enabled = False
Else    'Cancel Edit
  txtForename = ""
  txtSurname = ""
  txtCode = ""
  txtMCRN = ""
  cmbMemberOf.ListIndex = -1
  txtAutoLogOff = ""

  txtForename.BackColor = &H80000005
  txtSurname.BackColor = &H80000005
  txtCode.BackColor = &H80000005
  txtMCRN.BackColor = &H80000005
  cmbMemberOf.BackColor = &H80000005
  txtAutoLogOff.BackColor = &H80000005
  txtPass.BackColor = &H80000005
  txtConfirm.BackColor = &H80000005
  Frame1.Caption = "Add New Operator"

  cmdEdit.Caption = "&Edit"
  txtCode.Enabled = True
  cmbMemberOf.Enabled = True
End If

End Sub

Private Sub cmdExit_Click()

On Error GoTo cmdExit_Click_Error

cmbUserName.ListIndex = -1
txtPassword = ""
Me.Width = 7950
Me.Height = 3500

Exit Sub

cmdExit_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "cmdExit_Click", intEL, strES

End Sub

Private Sub cmdHide_Click()

Unload Me

End Sub

Private Sub cmdOK_Click()

          Dim tb As New Recordset
          Dim sql As String

10    On Error GoTo cmdOK_Click_Error

      'ENABLE THIS CODE FOR TEST SYSTEM.
      'If IsTestSystemExpired(1000) Then
      '    iMsg "System has expired, Please contact Custom Software."
      '    Exit Sub
      'End If

20    cmbUserName = Trim$(cmbUserName)
30    If cmbUserName = "" Then
40      cmbUserName.SetFocus
50      Exit Sub
60    End If

70    txtPassword = Trim$(txtPassword)
80    If txtPassword = "" Then
90      txtPassword.SetFocus
100     Exit Sub
110   End If

120   sql = "SELECT * FROM Users WHERE " & _
          "UserName = N'" & AddTicks(cmbUserName) & "' " & _
          "and Password = N'" & AddTicks(txtPassword) & "' "
130   Set tb = New Recordset
140   RecOpenServer 0, tb, sql
150   If Not tb.EOF Then
160     UserName = cmbUserName
170     UserMemberOf = tb!AccessLevel & ""
180     If DateDiff("d", tb!PassDate, Now) > Val(GetOptionSetting("PasswordExpirationDays", "45")) Then
190         If frmMsgBox.Msg("Your password has expired and must be changed", mbOKCancel, "Password Expired", mbExclamation) = 1 Then
200             With frmChangePass
210                 .Show 1
220                 If .Changed = False Then
230                     Exit Sub
240                 End If
250             End With
260         Else
270             Exit Sub
280         End If
290     Else
300         UserPass = UCase(tb!PassWord & "")
310     End If

320     If tb!AccessLevel = "Administrator" Then
330         Me.Width = 9060    '7950
340         Me.Height = 9765
350         cmbUserName = ""
360         txtPassword = ""
370         txtForename.SetFocus

380         UserName = ""
390         userCode = ""

400         Exit Sub
410     End If
420     UserName = Trim$(tb!UserName & "")
430     userCode = Trim$(tb!Code & "")
440     LogOffDelayMin = Val(tb!LogOffDelay & "")
450     LogOffDelaySecs = Val(tb!LogOffDelay & "") * 60
460     UpdateLoggedOnUser
470     Unload Me
        'MsgBox "After Unload)"
      '  Set Me = Nothing
480      'frmWorklist.Show
490     With frmWorklist
500         TimedOut = False
510         .Timer1.Enabled = True
            'MsgBox ("After Timer enabled")

520         .mnuLogOff.Enabled = True
530         .mnuAudit.Enabled = True
            'MsgBox ("Before show")
540         .Show
            'MsgBox ("After Show")
550     End With
      'frmWorklist.Timer1.Enabled = True

560   Else
570     LoginCount = LoginCount + 1
580     If LoginCount = 3 Then
590         frmMsgBox.Msg "3 Logins tried Program will now close" & vbCrLf & "Contact System Administrator", , , mbInformation
600         If TimedOut Then Unload Me: Exit Sub
610         End
620     End If
630     txtPassword = ""
640   End If

650   Exit Sub

cmdOK_Click_Error:

          Dim strES As String
          Dim intEL As Integer

660   intEL = Erl
670   strES = Err.Description
680   LogError "frmSystemManager", "cmdOK_Click", intEL, strES, sql

End Sub


Private Sub cmdSave_Click()

    Dim tb As New Recordset
    Dim sql As String

On Error GoTo cmdSave_Click_Error

txtForename = Trim$(txtForename)
txtCode = UCase$(Trim$(txtCode))
txtPass = UCase$(Trim$(txtPass))
txtConfirm = UCase$(Trim$(txtConfirm))

If Val(txtAutoLogOff) < 1 Then
  txtAutoLogOff = "5"
End If

If txtForename = "" Or txtSurname = "" Or txtCode = "" Or txtPass = "" Then
  frmMsgBox.Msg "Must have Name,Code and Password", , , mbCritical
  If TimedOut Then Unload Me: Exit Sub
  Exit Sub
End If

If cmbMemberOf = "" Then
  frmMsgBox.Msg "Member Of", , , mbCritical
  If TimedOut Then Unload Me: Exit Sub
  Exit Sub
End If

If UCase$(cmbMemberOf) = "CONSULTANT" And txtMCRN = "" Then
  frmMsgBox.Msg "You must enter a Medical Council Reference Number (MCRN)", , , mbCritical
  If TimedOut Then Unload Me: Exit Sub
  Exit Sub
End If

If txtPass <> txtConfirm Then
  txtPass = ""
  txtConfirm = ""
  frmMsgBox.Msg "Password Confirm dont match" & vbCrLf & _
                "Retype Password and Confirmation", , , mbExclamation
  If TimedOut Then Unload Me: Exit Sub
  Exit Sub
End If

If Frame1.Caption = "AddNewOperator" Then
  sql = "SELECT * FROM Users WHERE UserName = N'" & AddTicks(Trim$(txtForename)) & " " & AddTicks(Trim$(txtSurname)) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then
      frmMsgBox.Msg "UserName already used", , , mbExclamation
      If TimedOut Then Unload Me: Exit Sub
      txtForename = ""
      Exit Sub
  End If

  sql = "SELECT * from Users WHERE Code = N'" & Trim$(txtCode) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then
      frmMsgBox.Msg "Code already used", , , mbExclamation
      txtCode = ""
      Exit Sub
  End If


  sql = "SELECT * FROM Users WHERE Password = N'" & Trim$(txtPass) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then
      frmMsgBox.Msg "Password already used" & vbCrLf & "Type another Password", , , mbExclamation
      If TimedOut Then Unload Me: Exit Sub
      txtPass = ""
      txtConfirm = ""
      Exit Sub
  End If

  tb.AddNew
  tb!LogOffDelay = Val(txtAutoLogOff)
  tb!Code = txtCode
  tb!ForeName = Trim(Left(initial2upper(txtForename), 50))
  tb!Surname = Trim(Left(initial2upper(txtSurname), 50))
  tb!UserName = Trim(Left(initial2upper(txtForename), 50)) & " " & Trim(Left(initial2upper(txtSurname), 50))
  tb!PassWord = txtPass
  tb!AccessLevel = cmbMemberOf
  tb!MCRN = txtMCRN
  tb!PassDate = Format(Now - 51, "yyyy mm dd")
  tb.Update
Else    'Edit save
  sql = "SELECT * from Users WHERE Code = N'" & Trim$(txtCode) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then
      tb!LogOffDelay = Val(txtAutoLogOff)
      tb!Code = txtCode
      tb!ForeName = Trim(Left(initial2upper(txtForename), 50))
      tb!Surname = Trim(Left(initial2upper(txtSurname), 50))
      tb!UserName = Trim(Left(initial2upper(txtForename), 50)) & " " & Trim(Left(initial2upper(txtSurname), 50))
      tb!PassWord = txtPass
      tb!AccessLevel = cmbMemberOf
      tb!MCRN = txtMCRN
      tb!PassDate = Format(Now - 51, "yyyy mm dd")
      tb.Update
  End If

  txtForename = ""
  txtSurname = ""
  txtCode = ""
  txtMCRN = ""
  cmbMemberOf.ListIndex = -1
  txtAutoLogOff = ""

  txtForename.BackColor = &H80000005
  txtSurname.BackColor = &H80000005
  txtCode.BackColor = &H80000005
  txtMCRN.BackColor = &H80000005
  cmbMemberOf.BackColor = &H80000005
  txtAutoLogOff.BackColor = &H80000005
  txtPass.BackColor = &H80000005
  txtConfirm.BackColor = &H80000005
  Frame1.Caption = "Add New Operator"

  cmdEdit.Caption = "&Edit"
  cmdEdit.Visible = False
  txtCode.Enabled = True
  cmbMemberOf.Enabled = True
End If

FillG


txtForename = ""
txtSurname = ""
txtPass = ""
txtConfirm = ""
txtCode = ""
txtMCRN = ""
cmbMemberOf.ListIndex = -1
txtAutoLogOff = "5"

Exit Sub

cmdSave_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub FillG()

    Dim s As String
    Dim tb As New Recordset
    Dim sql As String

On Error GoTo FillG_Error

g.Visible = False
g.Rows = 2
g.AddItem ""
g.RemoveItem 1

cmbUserName.Clear

sql = "SELECT PassWord, UserName, AccessLevel, LogOffDelay, PassDate, Code, UserId From Users"
If AlphaOrderTechnicians Then
  sql = sql & "order by UserName Asc"
Else
End If
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  s = initial2upper(tb!UserName) & vbTab & _
      tb!AccessLevel & vbTab & _
      tb!LogOffDelay & vbTab & _
      "*****" & vbTab & _
      tb!PassWord & vbTab & _
      tb!Code & "" & vbTab & _
      tb!UserId & ""
  g.AddItem s

  cmbUserName.AddItem initial2upper(tb!UserName)

  tb.MoveNext
Loop

If g.Rows > 2 Then
  g.RemoveItem 1
End If
g.Visible = True

Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "FillG", intEL, strES, sql
g.Visible = True

End Sub

Private Sub Form_Activate()

      Dim Path As String
10    On Error GoTo Form_Activate_Error

20    If Not IsIDE Then
30      If SysOptChange(0) = True Then
40          Path = CheckNewEXE("Histology")    '<---Change this to your prog Name
50          If Path <> "" Then
60              Shell App.Path & "\CustomStart.exe Histology"    '<---Change this to your prog Name
70              End
80              Exit Sub
90          End If
100     End If
110   End If


      'frmSystemManager_ChangeLanguage
120   InitializeGrid
130   Me.Width = 9060
140   If UCase$(UserMemberOf) <> "ADMINISTRATOR" Then
150     g.Width = 7455
160     Me.Height = 3500
170   End If

180   txtPassword = ""

190   LoginCount = 0

200   FillG


220   Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmSystemManager", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

    Dim Con As String
    Dim ConBB As String

10    On Error GoTo Form_Load_Error

20    strVersion = App.Major & "." & App.Minor

30    Me.Caption = "NetAcquire - Cellular Pathology. Version " & strVersion

40    CheckIDE

50    strTestSystemColor = &HC0E0FF

60    If InStr(UCase$(App.Path), "TEST") Then
70      blnIsTestMode = True
80    Else
90      blnIsTestMode = False
100   End If

110   HospName(0) = GetcurrentConnectInfo(Con, ConBB)
'HospName(0) = "TULLAMORE"
120   If UCase(HospName(0)) = "TULLAMORE" Then
130     ConnectToDatabase
140   Else
150     connectDb
160   End If
170   CheckUpdateLanguage
180   LoadOptions
'LoadLanguage sysOptCurrentLanguage


190   ChangeFont Me, "Arial"



200   InitializeGrid
210   Me.Width = 6165
220   Me.Height = 3500

230   g.ColWidth(3) = 0

240   txtPassword = ""

250   FillG


260   blnVIstatus = GetOptionSetting("optValidateKeyAsciiChars", "0")
270   strDeptLetter4Histo = GetOptionSetting("optDeptLetter4Histo", "H")

280   With cmbMemberOf
290     .Clear
300     .AddItem "Lookup"
310     .AddItem "Clerical"
320     .AddItem "Scientist"
330     .AddItem "Manager"
340     .AddItem "IT Manager"
350     .AddItem "Consultant"
360     .AddItem "Specialist Registrar"
370     .AddItem "NCRI"
380     .AddItem "Administrator"
390   End With

400   If blnIsTestMode Then EnableTestMode Me

420   Exit Sub
      'Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "frmSystemManager", "Form_Load", intEL, strES

End Sub

Private Sub Form_Resize()
Me.Top = 0
Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Dim Form As Form
'For Each Form In Forms
'  Unload Form
'  Set Form = Nothing
'Next
End Sub

Private Sub g_Click()

    Static SortOrder As Boolean
    Dim LogOff As String
    Dim NewPassw As String
    Dim sql As String
    Dim tmpN As Integer
    Dim AdminCount As Integer
    Dim Y As Integer

    Dim X As Integer
    Dim ySave As Integer

On Error GoTo g_Click_Error

If g.MouseRow = 0 Then
  If SortOrder Then
      g.Sort = flexSortGenericAscending
  Else
      g.Sort = flexSortGenericDescending
  End If
  SortOrder = Not SortOrder
  Exit Sub
End If

ySave = g.row

g.col = 1
For Y = 1 To g.Rows - 1
  g.row = Y
  If g.CellBackColor = vbYellow Then
      For X = 1 To g.Cols - 1
          g.col = X
          g.CellBackColor = 0
      Next
      Exit For
  End If
Next

g.row = ySave
For X = 1 To g.Cols - 1
  g.col = X
  g.CellBackColor = vbYellow
Next


cmdEdit.Visible = True

    '110 AdminCount = 0
    '120 For Y = 1 To g.Rows - 1
    '130     If g.TextMatrix(Y, 1) = "Administrator" Then
    '140         AdminCount = AdminCount + 1
    '150     End If
    '160 Next

Select Case g.col

    Case 1:


  '180     Select Case g.TextMatrix(g.Row, 1)
  '        Case "Lookup":
  '190         g.TextMatrix(g.Row, 1) = "Clerical"
  '200     Case "Clerical":
  '210         g.TextMatrix(g.Row, 1) = "Scientist"
  '220     Case "Scientist":
  '230         g.TextMatrix(g.Row, 1) = "Manager"
  '240     Case "Manager":
  '250         g.TextMatrix(g.Row, 1) = "IT Manager"
  '260     Case "IT Manager":
  '270         g.TextMatrix(g.Row, 1) = "Consultant"
  '280     Case "Consultant":
  '290         g.TextMatrix(g.Row, 1) = "Specialist Registrar"
  '300     Case "Specialist Registrar":
  '310         g.TextMatrix(g.Row, 1) = "NCRI"
  '320     Case "NCRI":
  '330         g.TextMatrix(g.Row, 1) = "Administrator"
  '340     Case "Administrator":
  '350         If AdminCount > 1 Then
  '360             g.TextMatrix(g.Row, 1) = "Lookup"
  '370         End If
  '380     End Select
  '390     sql = "UPDATE Users " & _
   "SET AccessLevel = '" & g.TextMatrix(g.Row, 1) & "' " & _
   "WHERE userName = '" & AddTicks(g.TextMatrix(g.Row, 0)) & "'"
  '400     Cnxn(0).Execute sql




Case 2:    'Log Off Delay
  ' g.Enabled = False
  ' LogOff = g.TextMatrix(g.Row, 2)
  ' LogOff = iBOX("Log Off Delay", , LogOff)
  ' If TimedOut Then Unload Me: Exit Sub
  ' If LogOff = "" Then
  '     g.Enabled = True
  '     Exit Sub
  ' End If
  '  g.TextMatrix(g.Row, 2) = Format$(Val(LogOff))
  '  sql = "UPDATE Users " & _
     "SET LogOffDelay = " & Val(LogOff) & " " & _
     "WHERE userName = '" & g.TextMatrix(g.Row, 0) & "'"
  '  Cnxn(0).Execute sql
  '  g.Enabled = True

Case 3:
  ' g.Enabled = False
  ' NewPassw = iBOX("Reset Password" & vbCrLf & vbCrLf & g.TextMatrix(g.Row, 0), , , True)
  ' If TimedOut Then Unload Me: Exit Sub
  ' If NewPassw = "" Then
  '     g.Enabled = True
  '     Exit Sub
  ' End If
  ' sql = "UPDATE Users SET PassWord = '" & NewPassw & "', " & _
    '       "PassDate = '" & Format(Now - 50, "yyyymmdd") & "' " & _
    '       "WHERE password = '" & g.TextMatrix(g.Row, 4) & "'"
  ' Cnxn(0).Execute sql
  ' g.Enabled = True

End Select

    '660 tmpN = g.TopRow

    '670 FillG

    '680 g.TopRow = tmpN

Exit Sub

g_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "g_Click", intEL, strES, sql

End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim n As Long
    Static PrevY As Long

On Error GoTo g_MouseMove_Error

If g.MouseRow = 0 Then Exit Sub

If g.MouseCol = 3 Then
  g.ToolTipText = g.TextMatrix(g.MouseRow, 4)
  Exit Sub
ElseIf g.MouseCol = 0 Then
  Exit Sub
ElseIf g.MouseCol = 1 And Not AlphaOrderTechnicians Then
Else
  g.ToolTipText = ""
End If

    ' If Button = vbLeftButton And g.MouseRow > 0 And g.MouseCol = 1 Then
    '     If temp = "" Then
    '         PrevY = g.MouseRow
    '         For n = 0 To g.Cols - 1
    '             temp = temp & g.TextMatrix(g.Row, n) & vbTab
    '         Next
    '         temp = Left$(temp, Len(temp) - 1)
    '         Exit Sub
    '     Else
    '         If g.MouseRow <> PrevY Then
    '             g.RemoveItem PrevY
    '             If g.MouseRow <> PrevY Then
    '                 g.AddItem temp, g.MouseRow
    '                 PrevY = g.MouseRow
    '             Else
    '                 g.AddItem temp
    '                 PrevY = g.Rows - 1
    '             End If
    '         End If
    '     End If
    ' End If

Exit Sub

g_MouseMove_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "g_MouseMove", intEL, strES

End Sub

Public Property Let Manager(ByVal ShowManager As Boolean)

mManager = ShowManager

End Property



Public Property Let Operator(ByVal ShowOperator As Boolean)

On Error GoTo Operator_Error

mOperator = ShowOperator

Exit Property

Operator_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "Operator", intEL, strES


End Property

Public Property Let Secretary(ByVal ShowSecretary As Boolean)

On Error GoTo Secretary_Error

mSecretary = ShowSecretary

Exit Property

Secretary_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "Secretary", intEL, strES


End Property






Private Sub txtAutoLogOff_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)

KeyAscii = VI(KeyAscii, AlphaNumeric)

End Sub



Private Sub txtConfirm_KeyPress(KeyAscii As Integer)

KeyAscii = VI(KeyAscii, AlphaNumeric)

End Sub


Private Sub txtPass_KeyPress(KeyAscii As Integer)

KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

End Sub

Private Sub txtPassword_LostFocus()

On Error GoTo txtPassword_LostFocus_Error

txtPassword = UCase$(txtPassword)

Exit Sub

txtPassword_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "txtPassword_LostFocus", intEL, strES


End Sub


Private Sub InitializeGrid()
    Dim i As Integer
With g
  .Rows = 2: .FixedRows = 1
  .Cols = 7: .FixedCols = 1
  .Font.Size = fgcFontSize
  .Font.name = fgcFontName
  .ForeColor = fgcForeColor
  .BackColor = fgcBackColor
  .ForeColorFixed = fgcForeColorFixed
  .BackColorFixed = fgcBackColorFixed
  .ScrollBars = flexScrollBarBoth
  .TextMatrix(0, 0) = "username": .ColWidth(0) = 1875: .ColAlignment(0) = flexAlignLeftCenter
  .TextMatrix(0, 1) = "AccessRights": .ColWidth(1) = 1605: .ColAlignment(1) = flexAlignLeftCenter
  .TextMatrix(0, 2) = "LogOffDelayMinutes": .ColWidth(2) = 1500: .ColAlignment(2) = flexAlignLeftCenter
  .TextMatrix(0, 3) = "Password": .ColWidth(3) = 1065: .ColAlignment(3) = flexAlignLeftCenter
  .TextMatrix(0, 4) = "": .ColWidth(4) = 0: .ColAlignment(4) = flexAlignLeftCenter
  .TextMatrix(0, 5) = "Code": .ColWidth(5) = 1000: .ColAlignment(5) = flexAlignLeftCenter
  .TextMatrix(0, 6) = "": .ColWidth(6) = 0: .ColAlignment(6) = flexAlignLeftCenter

  For i = 0 To .Cols - 1
      If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
          .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + fgcExtraSpace
      End If
  Next i
End With
End Sub





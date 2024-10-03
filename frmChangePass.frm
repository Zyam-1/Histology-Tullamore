VERSION 5.00
Begin VB.Form frmChangePass 
   Caption         =   "Change Password"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   ControlBox      =   0   'False
   Icon            =   "frmChangePass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   855
      Left            =   3660
      Picture         =   "frmChangePass.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   675
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   4440
      Picture         =   "frmChangePass.frx":1208
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   675
   End
   Begin VB.TextBox txtConfirmPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2220
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2835
   End
   Begin VB.TextBox txtNewPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2220
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   2835
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Confirm New Password"
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   900
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter New Password"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   420
      Width           =   1485
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pChanged As Boolean


Private Sub cmdExit_Click()
pChanged = False
Unload Me

End Sub

Private Sub cmdUpdate_Click()
NewPass
End Sub

Private Sub NewPass()

Dim sql As String
Dim tb As New Recordset


On Error GoTo NewPass_Error


If txtNewPass <> txtConfirmPass Or Len(txtNewPass) < 6 Then
    If txtNewPass <> txtConfirmPass Then
        frmMsgBox.Msg "Passwords don't match!" & vbCrLf & "Enter New PassWord", mbOKOnly, "Change Password", mbExclamation
    Else
        frmMsgBox.Msg "Password must be at least six Characters." & vbCrLf & "Enter New PassWord", mbOKOnly, "Change Password", mbExclamation
    End If
    txtNewPass = ""
    txtConfirmPass = ""
    Exit Sub
End If

sql = "SELECT * from UsersAudit WHERE " & _
      "PassWord = N'" & txtNewPass & "' " & _
      "AND UserName = N'" & AddTicks(UserName) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
    frmMsgBox.Msg "Password previously used", mbOKOnly, "Change Password", mbExclamation
    txtNewPass = ""
    txtConfirmPass = ""
    Exit Sub
End If

sql = "SELECT * FROM Users WHERE " & _
      "PassWord = N'" & txtNewPass & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
    frmMsgBox.Msg "Password in use or been used!", mbOKOnly, "Change Password", mbExclamation
    txtNewPass = ""
    txtConfirmPass = ""
    Exit Sub
End If

sql = "IF EXISTS (SELECT * FROM Users WHERE " & _
      "           UserName = N'" & AddTicks(UserName) & "') " & _
      "  UPDATE Users " & _
      "  SET Password = N'" & UCase$(txtNewPass) & "', " & _
      "  Passdate = getdate() " & _
      "  WHERE UserName = N'" & AddTicks(UserName) & "'"
Cnxn(0).Execute sql

frmMsgBox.Msg UserName & " your Password is now Changed. Thank You!", mbOKOnly, "Password Changed", mbInformation

UserPass = UCase(txtNewPass)

pChanged = True

Unload Me

Exit Sub

NewPass_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemManager", "NewPass", intEL, strES, sql


End Sub

Public Property Get Changed() As Boolean


On Error GoTo Changed_Error

Changed = pChanged

Exit Property

Changed_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmChangePass", "Changed", intEL, strES


End Property

Private Sub Form_Load()
'frmChangePass_ChangeLanguage
End Sub

Private Sub Form_Resize()
Me.Top = 0
Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPhase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phase"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13605
   Icon            =   "frmPhase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMoveBack 
      Height          =   375
      Left            =   120
      Picture         =   "frmPhase.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdMoveForward 
      Height          =   375
      Left            =   12960
      Picture         =   "frmPhase.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame fraPhase 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   13335
      Begin MSComctlLib.ListView lstWorkPhase 
         Height          =   8175
         Index           =   4
         Left            =   8880
         TabIndex        =   1
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Special"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkPhase 
         Height          =   8175
         Index           =   3
         Left            =   6720
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Immuno"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkPhase 
         Height          =   8175
         Index           =   2
         Left            =   4560
         TabIndex        =   3
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cutting"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkPhase 
         Height          =   8175
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Embed"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkPhase 
         Height          =   8175
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CutUp"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkPhase 
         Height          =   8175
         Index           =   5
         Left            =   11040
         TabIndex        =   13
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Special"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLoggedIn 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   8760
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Logged In : "
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   8760
         Width           =   1335
      End
      Begin VB.Label lblWorkPhase 
         AutoSize        =   -1  'True
         Caption         =   "Cytology"
         Height          =   195
         Index           =   5
         Left            =   11040
         TabIndex        =   14
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblWorkPhase 
         AutoSize        =   -1  'True
         Caption         =   "Special"
         Height          =   195
         Index           =   4
         Left            =   8880
         TabIndex        =   10
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblWorkPhase 
         AutoSize        =   -1  'True
         Caption         =   "Immunohistochemical"
         Height          =   195
         Index           =   3
         Left            =   6720
         TabIndex        =   9
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lblWorkPhase 
         AutoSize        =   -1  'True
         Caption         =   "Cutting"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblWorkPhase 
         AutoSize        =   -1  'True
         Caption         =   "Embedding"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblWorkPhase 
         AutoSize        =   -1  'True
         Caption         =   "Cut-Up"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   720
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmPhase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCaseId As String
Public pSearch As Boolean


Public Property Let Search(ByVal bNewValue As Boolean)

10    pSearch = bNewValue

End Property

Private Sub FillLists()
      Dim sql As String
      Dim tb As Recordset
      Dim i As Integer

10    On Error GoTo FillLists_Error

20    sql = "SELECT DISTINCT C.CaseId,C.Phase FROM Cases C " & _
            "INNER JOIN BlockDetails B ON C.CaseId = B.CaseId " & _
            "WHERE C.State = N'" & "C" & "' " & _
            "UNION " & _
            "SELECT DISTINCT C.CaseId,C.Phase FROM Cases C " & _
            "WHERE LEFT(C.CaseId,1) = N'" & "C" & "' AND C.State = N'" & "In Histology" & "' "


30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    If Not tb.EOF Then
60        Do While Not tb.EOF
70            For i = 0 To 5
80                If UCase(tb!Phase) = UCase(lblWorkPhase(i)) Then
90                    If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
100                       lstWorkPhase(i).ListItems.Add , , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
110                   Else
120                       lstWorkPhase(i).ListItems.Add , , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
130                   End If
140               End If
150           Next
160           tb.MoveNext
170       Loop
180       MakeListsUnselected
190   End If


200   Exit Sub

FillLists_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmPhase", "FillLists", intEL, strES, sql


End Sub
Private Sub MakeListsUnselected()
      Dim i As Integer
      Dim j As Integer

10    On Error GoTo MakeListsUnselected_Error

20    For i = 0 To 5
30        For j = lstWorkPhase(i).ListItems.Count To 1 Step -1
40            lstWorkPhase(i).ListItems(j).Selected = False
50        Next
60    Next

70    Exit Sub

MakeListsUnselected_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmPhase", "MakeListsUnselected", intEL, strES

End Sub

Private Sub cmdMoveBack_Click()
      Dim i As Integer
      Dim s As Integer
      Dim SelectedListFound As Boolean

10    For i = 1 To 4
20        For s = lstWorkPhase(i).ListItems.Count To 1 Step -1
30            If lstWorkPhase(i).ListItems(s).Selected Then

40                lstWorkPhase(i - 1).ListItems.Add , , lstWorkPhase(i).ListItems(s).Text
50                UpdatePhase lstWorkPhase(i).ListItems(s).Text, lblWorkPhase(i - 1).Caption
60                lstWorkPhase(i - 1).ListItems(lstWorkPhase(i - 1).ListItems.Count).Selected = False
70                lstWorkPhase(i).ListItems.Remove (lstWorkPhase(i).ListItems(s).Index)
80                SelectedListFound = True
90            End If
100       Next s
110       If SelectedListFound Then
120           Exit Sub
130       End If
140   Next
End Sub

Private Sub cmdMoveForward_Click()
      Dim i As Integer
      Dim s As Integer
      Dim SelectedListFound As Boolean

10    For i = 0 To 3
20        For s = lstWorkPhase(i).ListItems.Count To 1 Step -1
30            If lstWorkPhase(i).ListItems(s).Selected Then

40                lstWorkPhase(i + 1).ListItems.Add , , lstWorkPhase(i).ListItems(s).Text
50                UpdatePhase lstWorkPhase(i).ListItems(s).Text, lblWorkPhase(i + 1).Caption
60                lstWorkPhase(i + 1).ListItems(lstWorkPhase(i + 1).ListItems.Count).Selected = False
70                lstWorkPhase(i).ListItems.Remove (lstWorkPhase(i).ListItems(s).Index)
80                SelectedListFound = True
90            End If
100       Next s
110       If SelectedListFound Then
120           Exit Sub
130       End If

140   Next
End Sub

Private Sub UpdatePhase(CaseId As String, Phase As String)
      Dim sql As String
      Dim tb As Recordset
      Dim TempCaseId As String

10    On Error GoTo UpdatePhase_Error

20    TempCaseId = Replace(CaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
30    sql = "SELECT * FROM Cases WHERE CaseId = N'" & TempCaseId & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then
70        tb!Phase = Phase
80        tb.Update
90    End If

100   Exit Sub

UpdatePhase_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmPhase", "UpdatePhase", intEL, strES, sql


End Sub

Private Sub Form_Load()
      Dim i As Integer
10    ChangeFont Me, "Arial"
'20    frmPhase_ChangeLanguage
30    If pSearch Then
40        cmdMoveForward.Enabled = False
50        cmdMoveBack.Enabled = False
60    End If
70    FillLists
80    For i = 0 To 5
90        Set lstWorkPhase(i).SelectedItem = Nothing
100   Next

110   If UCase$(UserMemberOf) = "CLERICAL" Or _
         UCase$(UserMemberOf) = "LOOKUP" Or _
         UCase$(UserMemberOf) = "NCRI" Then

120       cmdMoveForward.Enabled = False
130       cmdMoveBack.Enabled = False
140   End If

150   lblLoggedIn = UserName
160   If blnIsTestMode Then EnableTestMode Me
End Sub

Private Sub Form_Resize()
10    If Me.WindowState <> vbMinimized Then

20        Me.Top = 0
30        Me.Left = Screen.Width / 2 - Me.Width / 2
40    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
10    frmWorklist.Enabled = True
End Sub

Private Sub lstWorkPhase_DblClick(Index As Integer)
10    If UCase$(UserMemberOf) <> "LOOKUP" And _
         UCase$(UserMemberOf) <> "NCRI" Then
20        If lstWorkPhase(Index).SelectedItem Is Nothing Then
30            Exit Sub
40        Else
50            With frmWorkSheet
60                .txtCaseId = lstWorkPhase(Index).SelectedItem
70                .Show
80                .cmbPatientId.SetFocus
90            End With
100           Unload Me
110       End If
120   End If
End Sub

Private Sub lstWorkPhase_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      Dim i As Integer
10    If lstWorkPhase(Index).SelectedItem Is Nothing Then
20        For i = 0 To 5
30            Set lstWorkPhase(i).SelectedItem = Nothing
40        Next
50        sCaseId = ""
60        Exit Sub
70    Else
80        sCaseId = lstWorkPhase(Index).SelectedItem
90    End If
End Sub

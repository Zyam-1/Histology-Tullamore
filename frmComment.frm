VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComment 
   Caption         =   "Reason For Deleting"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComment 
      Height          =   2535
      Left            =   120
      MaxLength       =   4000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   870
      Left            =   5160
      Picture         =   "frmComment.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Height          =   870
      Left            =   5160
      Picture         =   "frmComment.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pNode As MSComctlLib.Node
Private pCommentType As String
Private pGeneralComment As String
Private pGridRow As Integer


Public Property Let Node(ByVal Value As MSComctlLib.Node)

10    Set pNode = Value

End Property

Public Property Let CommentType(ByVal Value As String)

10    pCommentType = Value

End Property

Public Property Let GeneralComment(ByVal Value As String)

10    pGeneralComment = Value

End Property

Public Property Let GridRow(ByVal Value As Integer)

10    pGridRow = Value

End Property


Private Sub cmdAdd_Click()


10    On Error GoTo cmdAdd_Click_Error

20    If UCase$(pCommentType) = "DELTREE" Then
30        SaveTreeComment
40    ElseIf UCase$(pCommentType) = "GENERAL" Then
50        AddGeneralComment
60        DataChanged = True
70    ElseIf UCase$(pCommentType) = "DISPOSAL" Then
80        AddDisposalComment
90        DataChanged = True
100   End If


110   Unload Me
120   Exit Sub

cmdAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmComment", "cmdAdd_Click", intEL, strES


End Sub

Private Sub AddGeneralComment()
10    frmWorkSheet.lblGeneralComments = txtComment

End Sub

Private Sub AddDisposalComment()
10    With frmHistDisposal.g
20          .TextMatrix(pGridRow, 5) = txtComment
30    End With
End Sub

Private Sub SaveTreeComment()
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo SaveTreeComment_Error

20    sql = "SELECT * FROM CaseTreeComment WHERE nodeID = '" & Right(pNode.Key, Len(pNode.Key) - 2) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If tb.EOF Then
60        tb.AddNew
70    End If
80    tb!NodeId = Right(pNode.Key, Len(pNode.Key) - 2)
90    tb!Comment = txtComment
100   tb!UserName = UserName
110   tb.Update

120   CaseUpdateLogEvent CaseNo, TreeNodeDeleted, "Reason : " & txtComment, pNode.FullPath

130   Exit Sub

SaveTreeComment_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmComment", "SaveTreeComment", intEL, strES, sql


End Sub

Private Sub cmdExit_Click()
10    Unload Me
End Sub

Private Sub Form_Activate()
10    If UCase$(pCommentType) = "GENERAL" Then
20        Me.Caption = "General Comments"
30        txtComment.Text = pGeneralComment
40    ElseIf UCase$(pCommentType) = "DELTREE" Then
50        Me.Caption = "Reason For Deleting"
60    ElseIf UCase$(pCommentType) = "DISPOSAL" Then
70        Me.Caption = "Comments"
80        txtComment.Text = pGeneralComment
90    End If

100   If bLocked Then
110       txtComment.Locked = True
120   End If
End Sub



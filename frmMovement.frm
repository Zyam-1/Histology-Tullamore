VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMovement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movement Tracker"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Height          =   870
      Left            =   5520
      Picture         =   "frmMovement.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   870
      Left            =   5520
      Picture         =   "frmMovement.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2610
      Width           =   975
   End
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
      Height          =   3375
      Index           =   3
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5145
      Begin VB.ComboBox cmbReferralReason 
         Height          =   315
         ItemData        =   "frmMovement.frx":0614
         Left            =   1680
         List            =   "frmMovement.frx":0616
         TabIndex        =   16
         Text            =   "cmbReferralReason"
         Top             =   2040
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtSent 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   2415
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Format          =   105906177
         CurrentDate     =   40269
      End
      Begin VB.ComboBox cmbDestination 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Text            =   "cmbDestination"
         Top             =   1665
         Width           =   3255
      End
      Begin MSMask.MaskEdBox txtTimeReceived 
         Height          =   315
         Left            =   4080
         TabIndex        =   9
         Top             =   2820
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtReceived 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   2820
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   " "
         Format          =   105775107
         CurrentDate     =   40269
      End
      Begin MSMask.MaskEdBox txtTimeSent 
         Height          =   315
         Left            =   4080
         TabIndex        =   13
         Top             =   2415
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   885
         Width           =   795
      End
      Begin VB.Label lblDescription 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1680
         TabIndex        =   18
         Top             =   840
         Width           =   3225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Reason for Referral"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2085
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date Received"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   2865
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date Sent"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   1260
         Width           =   3225
      End
      Begin VB.Label lblCode 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1680
         TabIndex        =   4
         Top             =   375
         Width           =   3225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Referred To"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1305
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   425
         Width           =   615
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   15
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
   Begin VB.Label lblSpecId 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6000
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pUpdate As Boolean
Private pMovementId As String
Private pCode As String
Private pDescription As String
Private pRefType As String
Private pSpecId As String






Private Sub cmdCancel_Click()
10  If pUpdate = False Then
20      If Left(pCode, 1) = "Q" Then
30          With frmWorkSheet.grdQCodes
40              If .Rows - .FixedRows = 1 Then
50                  .Rows = .Rows - 1
60              Else
70                  If .Rows > 1 Then
80                      .RemoveItem .Rows - 1
90                  End If
100             End If
110         End With
120     End If
130 End If

140 Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim s As String
    Dim CaseNode As MSComctlLib.Node
    Dim DateSent As String

10  If cmbDestination <> "" Then


20      With frmWorkSheet.grdTracker(pRefType)
30          If pUpdate Then

40              For i = 1 To .Rows - 1
50                  If .TextMatrix(i, 5) = pMovementId And .TextMatrix(i, 4) = pCode Then
60                      If dtReceived.CustomFormat <> " " Then
70                          If dtReceived < dtSent Then
80                              frmMsgBox.Msg "Date Received cannot be before Date Sent", , , mbExclamation
90                              Exit Sub
100                         ElseIf dtReceived = dtSent Then
110                             If txtTimeReceived.Text <> "" Then
120                                 If txtTimeReceived.FormattedText >= txtTimeSent.FormattedText Then
130                                     .TextMatrix(i, 3) = Format$(dtReceived, "dd/MM/yyyy") & " " & Format$(txtTimeReceived.FormattedText, "HH:mm")
140                                 Else
150                                     frmMsgBox.Msg "Date Received cannot be before Date Sent", , , mbExclamation
160                                     Exit Sub
170                                 End If
180                             Else
190                                 frmMsgBox.Msg "Please enter time received", , , mbExclamation
200                                 Exit Sub
210                             End If
220                         Else
230                             .TextMatrix(i, 3) = Format$(dtReceived, "dd/MM/yyyy") & " " & Format$(txtTimeReceived.FormattedText, "HH:mm")
240                         End If
250                     End If
260                     Exit For
270                 End If
280             Next
290         Else
300             DateSent = Format$(dtSent, "dd/MM/yyyy") & " " & Format$(txtTimeSent.FormattedText, "HH:mm")
310             s = lblDescription & vbTab & lblSpecId & vbTab & DateSent & vbTab & "" & vbTab & pCode & vbTab & pMovementId _
                  & vbTab & vbTab & lblType & vbTab & cmbDestination & vbTab & cmbReferralReason

320             .AddItem s
330             .Row = .Rows - 1
340             .Col = 6
350             .CellPictureAlignment = flexAlignCenterCenter

360             Set .CellPicture = frmWorkSheet.imgSquare.Picture

370             .Row = .Rows - 1
380             .Col = 10
390             .CellPictureAlignment = flexAlignCenterCenter

400             Set .CellPicture = frmWorkSheet.imgBlueInfo.Picture


410             With frmWorkSheet

420                 ChangeTabCaptionColour .SSTabMovement, .Picture1, vbRed, lblType, pRefType

430                 .SSTabMovement.Tab = pRefType
440             End With



450             Set CaseNode = frmWorkSheet.tvCaseDetails.SelectedItem
460             While Not CaseNode.Parent Is Nothing
470                 Set CaseNode = CaseNode.Parent
480             Wend

490             CaseNode.ForeColor = &H66FF
500         End If
510     End With
520     DataChanged = True
530     Unload Me
540 Else
550     frmMsgBox.Msg "Please enter destination", mbOKOnly, , mbInformation
560 End If
End Sub

Private Sub dtReceived_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10  dtReceived.CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub Form_Load()
    Dim Index As Integer
    Dim CaptionName As String

10  On Error GoTo Form_Load_Error

20  If pUpdate = True Then

30      Index = frmWorkSheet.SSTabMovement.Tab

40      With frmWorkSheet.grdTracker(Index)
50          lblCode = .TextMatrix(Rada, 4)
60          lblDescription = .TextMatrix(Rada, 0)
70          lblType = .TextMatrix(Rada, 7)
80          lblSpecId = .TextMatrix(Rada, 1)
90          FillLists
100         cmbDestination = .TextMatrix(Rada, 8)
110         cmbReferralReason = .TextMatrix(Rada, 9)
120         dtSent = Format$(Left$(.TextMatrix(Rada, 2), 10), "dd/MM/yyyy")
130         txtTimeSent = Format(Mid$(.TextMatrix(Rada, 2), 12), "HH:mm")
140         If .TextMatrix(Rada, 3) <> "" Then
150             dtReceived.CustomFormat = "dd/MM/yyyy"
160             dtReceived = Format$(Left$(.TextMatrix(Rada, 3), 10), "dd/MM/yyyy")
170             txtTimeReceived.BackColor = &H80000005
180             txtTimeReceived = Format(Mid$(.TextMatrix(Rada, 3), 12), "HH:mm")
190         Else
200             dtReceived.Value = Date

210         End If
220     End With
230 Else
240     Select Case pRefType
        Case 0
250         CaptionName = "Specimen"
260     Case 1
270         CaptionName = "Stain"
280     Case 2
290         CaptionName = "Case"
300     Case 3
310         CaptionName = "Block/Slide"
320     End Select

330     lblCode = pCode
340     lblDescription = pDescription
350     lblType = CaptionName
360     lblSpecId = pSpecId
370     FillLists
380     dtSent = Format$(Now, "dd/MM/yyyy")
390     txtTimeSent = Format$(Now, "HH:mm")
400     dtReceived.Enabled = False
410     txtTimeReceived.Enabled = False
420     txtTimeReceived.BackColor = &H8000000F
430 End If
440 If blnIsTestMode Then EnableTestMode Me
450 Exit Sub

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer

460 intEL = Erl
470 strES = Err.Description
480 LogError "frmMovement", "Form_Load", intEL, strES

End Sub

Public Property Let Update(ByVal Value As Boolean)

10  pUpdate = Value

End Property

Public Property Let MovementId(ByVal Id As String)

10  pMovementId = Id

End Property

Public Property Let Code(ByVal Id As String)

10  pCode = Id

End Property

Public Property Let Description(ByVal Id As String)

10  pDescription = Id

End Property
Public Property Let RefType(ByVal Id As Integer)

10  pRefType = Id

End Property

Public Property Let SpecId(ByVal Id As String)

10  pSpecId = Id

End Property



Private Sub FillLists()
    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo FillLists_Error

20  cmbReferralReason.Clear
30  cmbDestination.Clear

40  sql = "SELECT Description FROM Lists WHERE ListType = 'RefReason' order by description"
50  Set tb = New Recordset
60  RecOpenServer 0, tb, sql
70  If Not tb.EOF Then
80      While Not tb.EOF
90          cmbReferralReason.AddItem tb!Description
100         tb.MoveNext
110     Wend
120 End If

130 sql = "SELECT Description FROM Lists WHERE ListType = 'RefTo' order by description"
140 Set tb = New Recordset
150 RecOpenServer 0, tb, sql
160 If Not tb.EOF Then
170     While Not tb.EOF
180         cmbDestination.AddItem tb!Description
190         tb.MoveNext
200     Wend
210 End If


220 Exit Sub

FillLists_Error:

    Dim strES As String
    Dim intEL As Integer

230 intEL = Erl
240 strES = Err.Description
250 LogError "frmMovement", "FillLists", intEL, strES, sql


End Sub



Private Sub txtTimeReceived_GotFocus()
10  txtTimeReceived.SelStart = 0
20  txtTimeReceived.SelLength = Len(txtTimeReceived.FormattedText)
30  txtTimeReceived.SetFocus
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReferralDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Referral Details"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   ForeColor       =   &H00000000&
   Icon            =   "frmReferralDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   615
      Left            =   5340
      Picture         =   "frmReferralDetails.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8100
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   6120
      Picture         =   "frmReferralDetails.frx":1291
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8100
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Returned"
      Height          =   3075
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   6495
      Begin VB.TextBox txtComments 
         Height          =   915
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   1920
         Width           =   6015
      End
      Begin VB.CheckBox chkRpt 
         Height          =   315
         Left            =   4680
         TabIndex        =   29
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox cmbReceivedBy 
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtUnStainRet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5040
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtSpecialRet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtImmunoRet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtHERet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtBlocksRet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         Height          =   195
         Left            =   3720
         TabIndex        =   28
         Top             =   1125
         Width           =   480
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Received By"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1120
         Width           =   915
      End
      Begin VB.Label lblReturned 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unstained"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   4
         Left            =   5040
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblReturned 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Special"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   3
         Left            =   3840
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblReturned 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Immuno"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   2
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblReturned 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H&&E"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblReturned 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Blocks"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sent"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   6495
      Begin VB.ComboBox cmbSentBy 
         Height          =   315
         Left            =   4680
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cmbPermission 
         Height          =   315
         ItemData        =   "frmReferralDetails.frx":15D3
         Left            =   1440
         List            =   "frmReferralDetails.frx":15D5
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtUnStainSent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtSpecialSent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtImmunoSent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtHESent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtBlocksSent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Sent By"
         Height          =   195
         Left            =   3720
         TabIndex        =   24
         Top             =   1125
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Permission"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1120
         Width           =   750
      End
      Begin VB.Label lblSent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unstained"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   4
         Left            =   5040
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblSent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Special"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   3
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblSent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Immuno"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   2
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblSent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H&&E"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblSent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Blocks"
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   360
      TabIndex        =   47
      Top             =   60
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblDateSent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2880
      TabIndex        =   42
      Top             =   2460
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Sent"
      Height          =   195
      Index           =   1
      Left            =   540
      TabIndex        =   41
      Top             =   2460
      Width           =   720
   End
   Begin VB.Label lblRefReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2880
      TabIndex        =   40
      Top             =   2100
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Referral"
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   39
      Top             =   2100
      Width           =   1380
   End
   Begin VB.Label lblRefTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2880
      TabIndex        =   38
      Top             =   1740
      Width           =   45
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2880
      TabIndex        =   37
      Top             =   1380
      Width           =   45
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2880
      TabIndex        =   36
      Top             =   1020
      Width           =   45
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2880
      TabIndex        =   35
      Top             =   660
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referred To"
      Height          =   195
      Left            =   540
      TabIndex        =   34
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   540
      TabIndex        =   33
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   540
      TabIndex        =   32
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      Height          =   195
      Left            =   540
      TabIndex        =   31
      Top             =   660
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2355
      Left            =   345
      Top             =   480
      Width           =   6270
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referral Details"
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   345
      TabIndex        =   30
      Top             =   240
      Width           =   6270
   End
End
Attribute VB_Name = "frmReferralDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Index As Integer


Private Sub cmdExit_Click()
10  Unload Me
End Sub


Private Sub cmdSave_Click()
    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo cmdSave_Click_Error


20  sql = "SELECT * FROM CaseMovementDetails " & _
          "WHERE CaseListId = '" & frmWorkSheet.grdTracker(Index).TextMatrix(Rada, 5) & "'"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql
50  If tb.EOF Then tb.AddNew

60  tb!CaseId = CaseNo
70  tb!CaseListId = frmWorkSheet.grdTracker(Index).TextMatrix(Rada, 5)
80  tb!BlocksSent = txtBlocksSent
90  tb!HESent = txtHESent
100 tb!ImmunoSent = txtImmunoSent
110 tb!SpecialSent = txtSpecialSent
120 tb!UnstainedSent = txtUnStainSent
130 tb!SentPermission = cmbPermission
140 tb!SentBy = cmbSentBy
150 tb!BlocksRet = txtBlocksRet
160 tb!HERet = txtHERet
170 tb!ImmunoRet = txtImmunoRet
180 tb!SpecialRet = txtSpecialRet
190 tb!UnstainedRet = txtUnStainRet
200 tb!ReceivedBy = cmbReceivedBy
210 tb!ReportRet = chkRpt.Value
220 tb!Comments = txtComments
230 tb!UserName = UserName
240 tb.Update

250 With frmWorkSheet.grdTracker(Index)
260     .Row = Rada
270     .Col = 10
280     .CellPictureAlignment = flexAlignCenterCenter
290     If CheckReferralDiscrep(.TextMatrix(Rada, 5)) Then
300         Set .CellPicture = frmWorkSheet.imgRedInfo.Picture
310     Else
320         Set .CellPicture = frmWorkSheet.imgBlueInfo.Picture
330     End If
340 End With

350 Unload Me


360 Exit Sub

cmdSave_Click_Error:

    Dim strES As String
    Dim intEL As Integer

370 intEL = Erl
380 strES = Err.Description
390 LogError "frmReferralDetails", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

10  Index = frmWorkSheet.SSTabMovement.Tab

20  With frmWorkSheet.grdTracker(Index)
30      lblCode = .TextMatrix(Rada, 4)
40      lblDescription = .TextMatrix(Rada, 0)
50      lblType = .TextMatrix(Rada, 7)
60      lblRefTo = .TextMatrix(Rada, 8)
70      lblRefReason = .TextMatrix(Rada, 9)
80      lblDateSent = Format$(Left$(.TextMatrix(Rada, 2), 16), "dd/MM/yyyy HH:mm")
90  End With
100 FillLists
110 LoadReferralDetails
120 If blnIsTestMode Then EnableTestMode Me

End Sub
Private Sub LoadReferralDetails()
    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo LoadReferralDetails_Error

20  sql = "SELECT * FROM CaseMovementDetails " & _
          "WHERE CaseListId = '" & frmWorkSheet.grdTracker(Index).TextMatrix(Rada, 5) & "'"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  If Not tb.EOF Then
60      txtBlocksSent = tb!BlocksSent & ""
70      txtHESent = tb!HESent & ""
80      txtImmunoSent = tb!ImmunoSent & ""
90      txtSpecialSent = tb!SpecialSent & ""
100     txtUnStainSent = tb!UnstainedSent & ""
110     cmbPermission = tb!SentPermission & ""
120     cmbSentBy = tb!SentBy & ""
130     txtBlocksRet = tb!BlocksRet & ""
140     txtHERet = tb!HERet & ""
150     txtImmunoRet = tb!ImmunoRet & ""
160     txtSpecialRet = tb!SpecialRet & ""
170     txtUnStainRet = tb!UnstainedRet & ""
180     cmbReceivedBy = tb!ReceivedBy & ""
190     If tb!ReportRet = True Then
200         chkRpt.Value = vbChecked
210     Else
220         chkRpt.Value = vbUnchecked
230     End If
240     txtComments = tb!Comments & ""

250 End If

260 Exit Sub

LoadReferralDetails_Error:

    Dim strES As String
    Dim intEL As Integer

270 intEL = Erl
280 strES = Err.Description
290 LogError "frmReferralDetails", "LoadReferralDetails", intEL, strES, sql

End Sub

Private Sub FillLists()
    Dim sql As String
    Dim tb As Recordset



10  On Error GoTo FillLists_Error

20  sql = "SELECT * FROM Users " & _
          "WHERE AccessLevel = 'Consultant' " & _
          "OR AccessLevel = 'Scientist' " & _
          "OR AccessLevel = 'Manager' "

30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  cmbSentBy.Clear
60  cmbPermission.Clear
70  cmbReceivedBy.Clear

80  cmbSentBy.AddItem ""
90  cmbPermission.AddItem ""
100 cmbReceivedBy.AddItem ""

110 Do While Not tb.EOF
120     cmbSentBy.AddItem tb!UserName & ""
130     cmbSentBy.ItemData(cmbSentBy.NewIndex) = tb!UserId & ""
140     cmbPermission.AddItem tb!UserName & ""
150     cmbPermission.ItemData(cmbPermission.NewIndex) = tb!UserId & ""
160     cmbReceivedBy.AddItem tb!UserName & ""
170     cmbReceivedBy.ItemData(cmbReceivedBy.NewIndex) = tb!UserId & ""
180     tb.MoveNext
190 Loop

200 cmbSentBy.ListIndex = -1
210 cmbPermission.ListIndex = -1
220 cmbReceivedBy.ListIndex = -1



230 Exit Sub

FillLists_Error:

    Dim strES As String
    Dim intEL As Integer

240 intEL = Erl
250 strES = Err.Description
260 LogError "frmReferralDetails", "FillLists", intEL, strES, sql


End Sub


Private Sub txtBlocksRet_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtBlocksSent_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtHERet_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtHESent_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub


Private Sub txtImmunoRet_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub


Private Sub txtImmunoSent_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub


Private Sub txtSpecialRet_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub


Private Sub txtSpecialSent_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub


Private Sub txtUnStainRet_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtUnStainSent_KeyPress(KeyAscii As Integer)
10  KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

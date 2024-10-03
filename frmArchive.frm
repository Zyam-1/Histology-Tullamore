VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audit Changes"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   Icon            =   "frmArchive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Changes"
      Height          =   975
      Left            =   2880
      TabIndex        =   47
      Top             =   120
      Width           =   2655
      Begin VB.Label lblMovement 
         AutoSize        =   -1  'True
         Caption         =   "Movement"
         Height          =   195
         Left            =   1680
         TabIndex        =   51
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblCodes 
         AutoSize        =   -1  'True
         Caption         =   "Codes"
         Height          =   195
         Left            =   1680
         TabIndex        =   50
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblTreeChanges 
         AutoSize        =   -1  'True
         Caption         =   "Tree"
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   600
         Width           =   330
      End
      Begin VB.Label lblDemographics 
         AutoSize        =   -1  'True
         Caption         =   "Demographics"
         Height          =   195
         Left            =   360
         TabIndex        =   48
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Case ID"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdStart 
         Height          =   375
         Left            =   1440
         Picture         =   "frmArchive.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   435
      End
      Begin VB.TextBox txtCaseId 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   11
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Demographics"
      TabPicture(0)   =   "frmArchive.frx":1F4C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtbDemo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tree"
      TabPicture(1)   =   "frmArchive.frx":1F68
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbArcDateTime(1)"
      Tab(1).Control(1)=   "cmbArcDateTime(0)"
      Tab(1).Control(2)=   "lstImages"
      Tab(1).Control(3)=   "tvArcTree(1)"
      Tab(1).Control(4)=   "tvArcTree(0)"
      Tab(1).Control(5)=   "Label2"
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(7)=   "lblAuditByTree(0)"
      Tab(1).Control(8)=   "lblAuditByTree(1)"
      Tab(1).Control(9)=   "lblTree(0)"
      Tab(1).Control(10)=   "lblTree(1)"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Codes"
      TabPicture(2)   =   "frmArchive.frx":1F84
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "rtbCodes"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Movement "
      TabPicture(3)   =   "frmArchive.frx":1FA0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "rtbMovement"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Amendments"
      TabPicture(4)   =   "frmArchive.frx":1FBC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblAmend(1)"
      Tab(4).Control(1)=   "lblAmend(0)"
      Tab(4).Control(2)=   "Label4"
      Tab(4).Control(3)=   "lblAuditBy(1)"
      Tab(4).Control(4)=   "lblAuditBy(0)"
      Tab(4).Control(5)=   "Label1"
      Tab(4).Control(6)=   "txtAmendments(1)"
      Tab(4).Control(7)=   "txtAmendments(0)"
      Tab(4).Control(8)=   "cmbArcAmendments(1)"
      Tab(4).Control(9)=   "cmbArcAmendments(0)"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Micro"
      TabPicture(5)   =   "frmArchive.frx":1FD8
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblMicro(1)"
      Tab(5).Control(1)=   "lblMicro(0)"
      Tab(5).Control(2)=   "lblAuditByMicro(0)"
      Tab(5).Control(3)=   "lblAuditByMicro(1)"
      Tab(5).Control(4)=   "Label12"
      Tab(5).Control(5)=   "Label9"
      Tab(5).Control(6)=   "txtMicro(1)"
      Tab(5).Control(7)=   "txtMicro(0)"
      Tab(5).Control(8)=   "cmbArcMicro(0)"
      Tab(5).Control(9)=   "cmbArcMicro(1)"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Gross"
      TabPicture(6)   =   "frmArchive.frx":1FF4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblGross(1)"
      Tab(6).Control(1)=   "lblGross(0)"
      Tab(6).Control(2)=   "lblAuditByGross(1)"
      Tab(6).Control(3)=   "lblAuditByGross(0)"
      Tab(6).Control(4)=   "Label16"
      Tab(6).Control(5)=   "Label13"
      Tab(6).Control(6)=   "txtGross(1)"
      Tab(6).Control(7)=   "txtGross(0)"
      Tab(6).Control(8)=   "cmbArcGross(1)"
      Tab(6).Control(9)=   "cmbArcGross(0)"
      Tab(6).ControlCount=   10
      Begin VB.ComboBox cmbArcDateTime 
         Height          =   315
         Index           =   1
         Left            =   -69480
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.ComboBox cmbArcDateTime 
         Height          =   315
         Index           =   0
         Left            =   -74160
         TabIndex        =   11
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ComboBox cmbArcAmendments 
         Height          =   315
         Index           =   0
         Left            =   -74160
         TabIndex        =   10
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbArcAmendments 
         Height          =   315
         Index           =   1
         Left            =   -69240
         TabIndex        =   9
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbArcMicro 
         Height          =   315
         Index           =   1
         Left            =   -69240
         TabIndex        =   8
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbArcMicro 
         Height          =   315
         Index           =   0
         Left            =   -74160
         TabIndex        =   7
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbArcGross 
         Height          =   315
         Index           =   0
         Left            =   -74160
         TabIndex        =   6
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox cmbArcGross 
         Height          =   315
         Index           =   1
         Left            =   -69240
         TabIndex        =   5
         Top             =   1200
         Width           =   4215
      End
      Begin RichTextLib.RichTextBox txtGross 
         Height          =   4815
         Index           =   0
         Left            =   -74160
         TabIndex        =   4
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmArchive.frx":2010
      End
      Begin MSComctlLib.ImageList lstImages 
         Left            =   -72240
         Top             =   4080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchive.frx":2092
               Key             =   "two"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchive.frx":2419
               Key             =   "one"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvArcTree 
         Height          =   5295
         Index           =   1
         Left            =   -69480
         TabIndex        =   12
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9340
         _Version        =   393217
         Indentation     =   88
         Style           =   7
         ImageList       =   "lstImages"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvArcTree 
         Height          =   5295
         Index           =   0
         Left            =   -74160
         TabIndex        =   13
         Top             =   1560
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   9340
         _Version        =   393217
         Indentation     =   88
         Style           =   7
         ImageList       =   "lstImages"
         Appearance      =   1
      End
      Begin RichTextLib.RichTextBox rtbDemo 
         Height          =   6225
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   10980
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmArchive.frx":27A0
      End
      Begin RichTextLib.RichTextBox rtbCodes 
         Height          =   6345
         Left            =   -74520
         TabIndex        =   15
         Top             =   840
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   11192
         _Version        =   393217
         TextRTF         =   $"frmArchive.frx":2822
      End
      Begin RichTextLib.RichTextBox rtbMovement 
         Height          =   6225
         Left            =   -74520
         TabIndex        =   16
         Top             =   840
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   10980
         _Version        =   393217
         TextRTF         =   $"frmArchive.frx":28A4
      End
      Begin RichTextLib.RichTextBox txtGross 
         Height          =   4815
         Index           =   1
         Left            =   -69240
         TabIndex        =   17
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmArchive.frx":2926
      End
      Begin RichTextLib.RichTextBox txtMicro 
         Height          =   4815
         Index           =   0
         Left            =   -74160
         TabIndex        =   18
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmArchive.frx":29A8
      End
      Begin RichTextLib.RichTextBox txtMicro 
         Height          =   4815
         Index           =   1
         Left            =   -69240
         TabIndex        =   19
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmArchive.frx":2A2A
      End
      Begin RichTextLib.RichTextBox txtAmendments 
         Height          =   4815
         Index           =   0
         Left            =   -74160
         TabIndex        =   20
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmArchive.frx":2AAC
      End
      Begin RichTextLib.RichTextBox txtAmendments 
         Height          =   4815
         Index           =   1
         Left            =   -69240
         TabIndex        =   21
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmArchive.frx":2B2E
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Previous Version"
         Height          =   195
         Left            =   -69480
         TabIndex        =   45
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date"
         Height          =   195
         Left            =   -74160
         TabIndex        =   44
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lblAuditBy 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   43
         Top             =   6720
         Width           =   45
      End
      Begin VB.Label lblAuditBy 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   -68160
         TabIndex        =   42
         Top             =   6720
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date"
         Height          =   195
         Left            =   -69240
         TabIndex        =   41
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date"
         Height          =   195
         Left            =   -74160
         TabIndex        =   40
         Top             =   840
         Width           =   750
      End
      Begin VB.Label lblAuditByTree 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   39
         Top             =   7080
         Width           =   165
      End
      Begin VB.Label lblAuditByTree 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   -68400
         TabIndex        =   38
         Top             =   7080
         Width           =   75
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date"
         Height          =   195
         Left            =   -69240
         TabIndex        =   37
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date"
         Height          =   195
         Left            =   -74160
         TabIndex        =   36
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date"
         Height          =   195
         Left            =   -74160
         TabIndex        =   35
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Audit Date"
         Height          =   195
         Left            =   -69240
         TabIndex        =   34
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lblAuditByMicro 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   -68160
         TabIndex        =   33
         Top             =   6720
         Width           =   165
      End
      Begin VB.Label lblAuditByMicro 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   32
         Top             =   6720
         Width           =   165
      End
      Begin VB.Label lblAuditByGross 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   31
         Top             =   6720
         Width           =   165
      End
      Begin VB.Label lblAuditByGross 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   -68160
         TabIndex        =   30
         Top             =   6720
         Width           =   165
      End
      Begin VB.Label lblTree 
         AutoSize        =   -1  'True
         Caption         =   "Created By:"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   29
         Top             =   7080
         Width           =   825
      End
      Begin VB.Label lblTree 
         AutoSize        =   -1  'True
         Caption         =   "Created By:"
         Height          =   195
         Index           =   1
         Left            =   -69480
         TabIndex        =   28
         Top             =   7080
         Width           =   825
      End
      Begin VB.Label lblAmend 
         AutoSize        =   -1  'True
         Caption         =   "Created By:"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   27
         Top             =   6720
         Width           =   825
      End
      Begin VB.Label lblAmend 
         AutoSize        =   -1  'True
         Caption         =   "Created By:"
         Height          =   195
         Index           =   1
         Left            =   -69240
         TabIndex        =   26
         Top             =   6720
         Width           =   825
      End
      Begin VB.Label lblMicro 
         AutoSize        =   -1  'True
         Caption         =   "Created By:"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   25
         Top             =   6720
         Width           =   825
      End
      Begin VB.Label lblMicro 
         AutoSize        =   -1  'True
         Caption         =   "Created By:"
         Height          =   195
         Index           =   1
         Left            =   -69240
         TabIndex        =   24
         Top             =   6720
         Width           =   825
      End
      Begin VB.Label lblGross 
         AutoSize        =   -1  'True
         Caption         =   "Created By:"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   23
         Top             =   6720
         Width           =   825
      End
      Begin VB.Label lblGross 
         AutoSize        =   -1  'True
         Caption         =   "Created By:"
         Height          =   195
         Index           =   1
         Left            =   -69240
         TabIndex        =   22
         Top             =   6720
         Width           =   825
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SID As String
Dim CompareWords() As String


Private Sub cmbArcAmendments_Click(Index As Integer)
10  FillArcAmendments (Index)
End Sub

Private Sub cmbArcDateTime_Click(Index As Integer)
10  If cmbArcDateTime(0).ListIndex < cmbArcDateTime(1).ListCount - 1 Then
20      cmbArcDateTime(1).ListIndex = cmbArcDateTime(0).ListIndex + 1
30      FillArcTree (0)
40      FillArcTree (1)
50  Else
60      FillArcTree (0)
70      tvArcTree(1).Nodes.Clear
80      lblAuditByTree(1) = ""
90  End If
End Sub



Private Sub cmbArcGross_Click(Index As Integer)
10  FillArcGross (Index)
End Sub

Private Sub cmbArcMicro_Click(Index As Integer)
10  FillArcMicro (Index)
End Sub

Private Sub Form_Load()
10  SSTab1.TabVisible(4) = False
20  SSTab1.TabVisible(5) = False
30  SSTab1.TabVisible(6) = False
40  If blnIsTestMode Then EnableTestMode Me
End Sub

Private Sub Form_Resize()
10  If Me.WindowState <> vbMinimized Then

20      Me.Top = 0
30      Me.Left = Screen.Width / 2 - Me.Width / 2
40  End If
End Sub

Private Sub cmdStart_Click()
10  ClearFields
20  SID = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
30  DoDemographics
40  FillTreeCombo
50  DoPCodes
60  DoOtherCodes
70  DoMovement

80  If rtbDemo.Text <> "" Then
90      lblDemographics.ForeColor = vbRed
100 Else
110     lblDemographics.ForeColor = vbBlack
120 End If

130 If tvArcTree(1).Nodes.Count > 0 Then
140     lblTreeChanges.ForeColor = vbRed
150 Else
160     lblTreeChanges.ForeColor = vbBlack
170 End If

180 If rtbCodes.Text <> "" Then
190     lblCodes.ForeColor = vbRed
200 Else
210     lblCodes.ForeColor = vbBlack
220 End If

230 If rtbMovement.Text <> "" Then
240     lblMovement.ForeColor = vbRed
250 Else
260     lblMovement.ForeColor = vbBlack
270 End If



    'FillAmendCombo
    'FillMicroCombo
    'FillGrossCombo

End Sub

Private Sub ClearFields()
    Dim i As Integer
10  rtbDemo = ""
20  For i = 0 To 1
30      cmbArcDateTime(i).Clear
        '    cmbArcAmendments(i).Clear
        '    cmbArcMicro(i).Clear
        '    cmbArcGross(i).Clear
40      tvArcTree(i).Nodes.Clear
50  Next
60  rtbCodes = ""
70  rtbMovement = ""
End Sub

Private Sub FillTreeCombo()
    Dim sql As String
    Dim tb As New Recordset
    Dim i As Integer



10  On Error GoTo FillTreeCombo_Error

20  sql = "SELECT Distinct convert(varchar,ArchiveDateTime,21) as ArchiveDateTime FROM CaseDetailsAudit " & _
          "WHERE CaseId = '" & SID & "' ORDER BY ArchiveDateTime DESC"

30  Set tb = New Recordset
40  RecOpenClient 0, tb, sql
50  For i = 0 To 1
60      cmbArcDateTime(i).AddItem "Current"
70  Next
80  Do While Not tb.EOF
90      For i = 0 To 1
100         cmbArcDateTime(i).AddItem tb!ArchiveDateTime
110     Next
120     tb.MoveNext
130 Loop

140 For i = 0 To 1
150     cmbArcDateTime(i).ListIndex = 0
160 Next


170 Exit Sub

FillTreeCombo_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmArchive", "FillTreeCombo", intEL, strES, sql


End Sub

'Private Sub FillAmendCombo()
'    Dim sql As String
'    Dim tb As New Recordset
'    Dim i As Integer


'10  On Error GoTo FillAmendCombo_Error

'20  sql = "SELECT Distinct convert(varchar,ArchiveDateTime,21) as ArchiveDateTime FROM CaseAmendmentsAudit " & _
          "WHERE CaseId = '" & SID & "' ORDER BY ArchiveDateTime DESC"

'30  Set tb = New Recordset
'40  RecOpenClient 0, tb, sql
'50  For i = 0 To 1
'60      cmbArcAmendments(i).AddItem "Current"
'70  Next
'80  Do While Not tb.EOF
'90      For i = 0 To 1
'100         cmbArcAmendments(i).AddItem tb!ArchiveDateTime
'110     Next
'120     tb.MoveNext
'130 Loop

'140 For i = 0 To 1
'150     cmbArcAmendments(i).ListIndex = 0
'160 Next

'170 Exit Sub

'FillAmendCombo_Error:

'    Dim strES As String
'    Dim intEL As Integer

'180 intEL = Erl
'190 strES = Err.Description
'200 LogError "frmArchive", "FillAmendCombo", intEL, strES, sql


'End Sub

'Private Sub FillMicroCombo()
'    Dim sql As String
'    Dim tb As New Recordset
'    Dim i As Integer


'10  On Error GoTo FillMicroCombo_Error

'20  sql = "SELECT Distinct convert(varchar,ArchiveDateTime,21) as ArchiveDateTime FROM CasesAudit " & _
          "WHERE CaseId = '" & SID & "' AND Micro IS NOT NULL ORDER BY ArchiveDateTime DESC"

'30  Set tb = New Recordset
'40  RecOpenClient 0, tb, sql
'50  For i = 0 To 1
'60      cmbArcMicro(i).AddItem "Current"
'70  Next
'80  Do While Not tb.EOF
'90      For i = 0 To 1
'100         cmbArcMicro(i).AddItem tb!ArchiveDateTime
'110     Next
'120     tb.MoveNext
'130 Loop

'140 For i = 0 To 1
'150     cmbArcMicro(i).ListIndex = 0
'160 Next

'170 Exit Sub

'FillMicroCombo_Error:

'    Dim strES As String
'    Dim intEL As Integer

'180 intEL = Erl
'190 strES = Err.Description
'200 LogError "frmArchive", "FillMicroCombo", intEL, strES, sql

'End Sub

'Private Sub FillGrossCombo()
'    Dim sql As String
'    Dim tb As New Recordset
'    Dim i As Integer


'10  On Error GoTo FillGrossCombo_Error

'20  sql = "SELECT Distinct convert(varchar,ArchiveDateTime,21) as ArchiveDateTime FROM CasesAudit " & _
          "WHERE CaseId = '" & SID & "' AND Gross IS NOT NULL ORDER BY ArchiveDateTime DESC"

'30  Set tb = New Recordset
'40  RecOpenClient 0, tb, sql
'50  For i = 0 To 1
'60      cmbArcGross(i).AddItem "Current"
'70  Next
'80  Do While Not tb.EOF
'90      For i = 0 To 1
'100         cmbArcGross(i).AddItem tb!ArchiveDateTime
'110     Next
'120     tb.MoveNext
'130 Loop

'140 For i = 0 To 1
'150     cmbArcGross(i).ListIndex = 0
'160 Next

'170 Exit Sub

'FillGrossCombo_Error:

'    Dim strES As String
'    Dim intEL As Integer

'180 intEL = Erl
'190 strES = Err.Description
'200 LogError "frmArchive", "FillGrossCombo", intEL, strES, sql


'End Sub


Private Sub FillArcTree(Index As Integer)
    Dim tb As New Recordset
    Dim sql As String
    Dim TissueType As MSComctlLib.Node
    Dim Block As MSComctlLib.Node
    Dim Slide As MSComctlLib.Node
    Dim CaseId As MSComctlLib.Node
    Dim Stain As MSComctlLib.Node

10  On Error GoTo FillArcTree_Error

20  If cmbArcDateTime(Index) = "Current" Then
30      sql = "SELECT * FROM CaseDetails " & _
              "WHERE CaseId = '" & SID & "' "
40  Else
50      sql = "SELECT * FROM CaseDetailsAudit " & _
              "WHERE CaseId = '" & SID & "' AND ArchiveDateTime = '" & cmbArcDateTime(Index) & "' "
60  End If

    'sql = sql & "ORDER BY CD.CaseId, CD.TissueType, CD.Block, CD.Slide"


70  Set tb = New Recordset
80  RecOpenClient 0, tb, sql
90  tvArcTree(Index).Nodes.Clear
100 If Not tb.EOF Then

110     Do While Not tb.EOF


120         If Not CaseId Is Nothing Then
130             If (CaseId.Key <> "L1" & tb!CaseId) Then
140                 Set CaseId = tvArcTree(Index).Nodes.Add(, , "L1" & tb!CaseId, tb!CaseId & "", 1, 2)
150                 CaseId.Expanded = True
160             End If
170         Else
180             Set CaseId = tvArcTree(Index).Nodes.Add(, , "L1" & tb!CaseId, tb!CaseId & "", 1, 2)
190             CaseId.Expanded = True
200         End If

210         If tb!TissueType <> "" Then
220             If Not TissueType Is Nothing Then
230                 If (TissueType.Key <> "L2" & tb!CaseId & tb!TissueType) Then
240                     Set TissueType = tvArcTree(Index).Nodes.Add("L1" & tb!CaseId, tvwChild, "L2" & tb!CaseId & tb!TissueType, tb!TissueType & "", 1, 2)
250                     TissueType.Expanded = True
260                 End If
270             Else

280                 Set TissueType = tvArcTree(Index).Nodes.Add("L1" & tb!CaseId, tvwChild, "L2" & tb!CaseId & tb!TissueType, tb!TissueType & "", 1, 2)
290                 TissueType.Expanded = True
300             End If
310         End If

320         If tb!Block <> "" Then
330             If Not Block Is Nothing Then
340                 If (Block.Key <> "L3" & tb!CaseId & tb!TissueType & tb!Block) Then
350                     Set Block = tvArcTree(Index).Nodes.Add("L2" & tb!CaseId & tb!TissueType, tvwChild, "L3" & tb!CaseId & tb!TissueType & tb!Block, tb!Block, 1, 2)
360                     Block.Expanded = True
370                 End If
380             Else
390                 Set Block = tvArcTree(Index).Nodes.Add("L2" & tb!CaseId & tb!TissueType, tvwChild, "L3" & tb!CaseId & tb!TissueType & tb!Block, tb!Block, 1, 2)
400                 Block.Expanded = True
410             End If
420         End If

430         If tb!Slide <> "" Then
440             If Not Slide Is Nothing Then
450                 If (Slide.Key <> "L4" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide) Then
460                     If tb!Stain <> "" Then
470                         Set Slide = tvArcTree(Index).Nodes.Add("L3" & tb!CaseId & tb!TissueType & tb!Block, tvwChild, "L4" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide, "**" & tb!Slide & "", 1, 2)
480                         Slide.Expanded = True
                            'CaseId.ForeColor = &H66FF
490                     Else
500                         Set Slide = tvArcTree(Index).Nodes.Add("L3" & tb!CaseId & tb!TissueType & tb!Block, tvwChild, "L4" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide, tb!Slide & "", 1, 2)
510                         Slide.Expanded = True
520                     End If
530                 End If
540             Else
550                 If tb!Stain <> "" Then
560                     Set Slide = tvArcTree(Index).Nodes.Add("L3" & tb!CaseId & tb!TissueType & tb!Block, tvwChild, "L4" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide, "**" & tb!Slide & "", 1, 2)
570                     Slide.Expanded = True
                        'CaseId.ForeColor = &H66FF
580                 Else
590                     Set Slide = tvArcTree(Index).Nodes.Add("L3" & tb!CaseId & tb!TissueType & tb!Block, tvwChild, "L4" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide, tb!Slide, 1, 2)
600                     Slide.Expanded = True
610                 End If
620             End If
630         End If

640         If tb!Stain <> "" Then
650             If Not Stain Is Nothing Then
660                 If (Stain.Key <> "L5" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide & tb!Stain) Then
670                     Set Stain = tvArcTree(Index).Nodes.Add("L4" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide, tvwChild, "L5" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide & tb!Stain, tb!Stain, 1, 2)
680                     Stain.Expanded = True
690                 End If
700             Else
710                 Set Stain = tvArcTree(Index).Nodes.Add("L4" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide, tvwChild, "L5" & tb!CaseId & tb!TissueType & IIf(tb!Block = "Block ", "Block A", tb!Block) & tb!Slide & tb!Stain, tb!Stain, 1, 2)
720                 Stain.Expanded = True
730             End If
740         End If

            '            If cmbArcDateTime(Index) = "Current" Then
750         lblAuditByTree(Index) = tb!UserName & ""
760         lblTree(Index) = "Created By:"
            '            Else
            '                If tb!ArchivedBy <> "" Then
            '                    lblAuditByTree(Index) = tb!ArchivedBy
            '
            '                Else
            '                   lblAuditByTree(Index) = "???"
            '                End If
            '                lblTree(Index) = "Changed By:"
            '            End If
            '
770         tb.MoveNext

780     Loop



790 End If
800 HighlightChanges


810 Exit Sub

FillArcTree_Error:

    Dim strES As String
    Dim intEL As Integer

820 intEL = Erl
830 strES = Err.Description
840 LogError "frmArchive", "FillArcTree", intEL, strES, sql


End Sub


Private Sub FillArcAmendments(Index As Integer)
    Dim tb As New Recordset
    Dim sql As String
    Dim i As Integer
    Dim X As Integer
    Dim lngPos As Long
    Dim lngSelStart As Long
    Dim lngSelLength As Long


10  On Error GoTo FillArcAmendments_Error

20  If cmbArcAmendments(Index) = "Current" Then
30      sql = "SELECT * FROM CaseAmendments " & _
              "WHERE CaseId = '" & SID & "' "
40  Else
50      sql = "SELECT * FROM CaseAmendmentsAudit " & _
              "WHERE CaseId = '" & SID & "' AND ArchiveDateTime = '" & cmbArcAmendments(Index) & "' "
60  End If

70  Set tb = New Recordset
80  RecOpenClient 0, tb, sql

90  If Not tb.EOF Then


100     txtAmendments(Index).Text = tb!Comment & ""


110     For i = 0 To 1
120         txtAmendments(i).SelStart = 0
130         txtAmendments(i).SelLength = Len(txtAmendments(i).Text)
140         txtAmendments(i).SelColor = vbRed
150         If i = 0 Then
160             CompareWords() = Split(Replace(txtAmendments(1).Text, vbCrLf, " "), " ")
170         Else
180             CompareWords() = Split(Replace(txtAmendments(0).Text, vbCrLf, " "), " ")
190         End If
            'CompareWords() = Split(tb!Comment & "", " ")


200         lngSelStart = txtAmendments(i).SelStart

210         lngSelLength = txtAmendments(i).SelLength

220         For X = 0 To UBound(CompareWords)
230             lngPos = txtAmendments(i).Find(CompareWords(X), 0)

240             Do
250                 If lngPos <> -1 Then

260                     If txtAmendments(i).SelColor = vbBlack Then
270                         lngPos = txtAmendments(i).Find(CompareWords(X), lngPos + txtAmendments(i).SelLength)
280                     Else
290                         txtAmendments(i).SelColor = vbBlack
300                         lngPos = 0
310                     End If
320                 End If


330             Loop While lngPos > 0

340             txtAmendments(i).SelStart = lngSelStart

350             txtAmendments(i).SelLength = lngSelLength
360         Next

370     Next



380     If cmbArcAmendments(Index) = "Current" Then
390         lblAuditBy(Index) = tb!UserName & ""
400         lblAmend(Index) = "Created By:"
410     Else
420         If tb!ArchivedBy <> "" Then
430             lblAuditBy(Index) = tb!ArchivedBy
440         Else
450             lblAuditBy(Index) = "???"
460         End If
470         lblAmend(Index) = "Changed By:"
480     End If

490 End If

500 Exit Sub

FillArcAmendments_Error:

    Dim strES As String
    Dim intEL As Integer

510 intEL = Erl
520 strES = Err.Description
530 LogError "frmArchive", "FillArcAmendments", intEL, strES, sql


End Sub

Private Sub FillArcMicro(Index As Integer)
    Dim tb As New Recordset
    Dim sql As String
    Dim i As Integer
    Dim X As Integer
    Dim lngPos As Long
    Dim lngSelStart As Long
    Dim lngSelLength As Long


10  On Error GoTo FillArcMicro_Error

20  If cmbArcMicro(Index) = "Current" Then
30      sql = "SELECT * FROM Cases " & _
              "WHERE CaseId = '" & SID & "' "
40  Else
50      sql = "SELECT * FROM CasesAudit " & _
              "WHERE CaseId = '" & SID & "' AND ArchiveDateTime = '" & cmbArcMicro(Index) & "' "
60  End If

70  Set tb = New Recordset
80  RecOpenClient 0, tb, sql

90  If Not tb.EOF Then

100     txtMicro(Index).Text = tb!Micro & ""

110     For i = 0 To 1
120         txtMicro(i).SelStart = 0
130         txtMicro(i).SelLength = Len(txtAmendments(i).Text)
140         txtMicro(i).SelColor = vbRed
150         If i = 0 Then
160             CompareWords() = Split(Replace(txtMicro(1).Text, vbCrLf, " "), " ")
170         Else
180             CompareWords() = Split(Replace(txtMicro(0).Text, vbCrLf, " "), " ")
190         End If
            'CompareWords() = Split(tb!Comment & "", " ")


200         lngSelStart = txtMicro(i).SelStart

210         lngSelLength = txtMicro(i).SelLength

220         For X = 0 To UBound(CompareWords)
230             lngPos = txtMicro(i).Find(CompareWords(X), 0)

240             Do
250                 If lngPos <> -1 Then

260                     If txtMicro(i).SelColor = vbBlack Then
270                         lngPos = txtMicro(i).Find(CompareWords(X), lngPos + txtMicro(i).SelLength)
280                     Else
290                         txtMicro(i).SelColor = vbBlack
300                         lngPos = 0
310                     End If
320                 End If


330             Loop While lngPos > 0

340             txtMicro(i).SelStart = lngSelStart

350             txtMicro(i).SelLength = lngSelLength
360         Next

370     Next

380     If cmbArcMicro(Index) = "Current" Then
390         lblAuditByMicro(Index) = tb!UserName & ""
400         lblMicro(Index) = "Created By:"
410     Else
420         If tb!ArchivedBy <> "" Then
430             lblAuditByMicro(Index) = tb!ArchivedBy
440         Else
450             lblAuditByMicro(Index) = "???"
460         End If
470         lblMicro(Index) = "Changed By:"
480     End If

490 End If

500 Exit Sub

FillArcMicro_Error:

    Dim strES As String
    Dim intEL As Integer

510 intEL = Erl
520 strES = Err.Description
530 LogError "frmArchive", "FillArcMicro", intEL, strES, sql


End Sub

Private Sub FillArcGross(Index As Integer)
    Dim tb As New Recordset
    Dim sql As String
    Dim i As Integer
    Dim X As Integer
    Dim lngPos As Long
    Dim lngSelStart As Long
    Dim lngSelLength As Long


10  On Error GoTo FillArcGross_Error

20  If cmbArcGross(Index) = "Current" Then
30      sql = "SELECT * FROM Cases " & _
              "WHERE CaseId = '" & SID & "' "
40  Else
50      sql = "SELECT * FROM CasesAudit " & _
              "WHERE CaseId = '" & SID & "' AND ArchiveDateTime = '" & cmbArcGross(Index) & "' "
60  End If

70  Set tb = New Recordset
80  RecOpenClient 0, tb, sql

90  If Not tb.EOF Then

100     txtGross(Index).Text = tb!Gross & ""

110     For i = 0 To 1
120         txtGross(i).SelStart = 0
130         txtGross(i).SelLength = Len(txtAmendments(i).Text)
140         txtGross(i).SelColor = vbRed
150         If i = 0 Then
160             CompareWords() = Split(Replace(txtGross(1).Text, vbCrLf, " "), " ")
170         Else
180             CompareWords() = Split(Replace(txtGross(0).Text, vbCrLf, " "), " ")
190         End If
            'CompareWords() = Split(tb!Comment & "", " ")


200         lngSelStart = txtGross(i).SelStart

210         lngSelLength = txtGross(i).SelLength

220         For X = 0 To UBound(CompareWords)
230             lngPos = txtGross(i).Find(CompareWords(X), 0)

240             Do
250                 If lngPos <> -1 Then

260                     If txtGross(i).SelColor = vbBlack Then
270                         lngPos = txtGross(i).Find(CompareWords(X), lngPos + txtGross(i).SelLength)
280                     Else
290                         txtGross(i).SelColor = vbBlack
300                         lngPos = 0
310                     End If
320                 End If


330             Loop While lngPos > 0

340             txtGross(i).SelStart = lngSelStart

350             txtGross(i).SelLength = lngSelLength
360         Next

370     Next

380     If cmbArcGross(Index) = "Current" Then
390         lblAuditByGross(Index) = tb!UserName & ""
400         lblGross(Index) = "Created By:"
410     Else
420         If tb!ArchivedBy <> "" Then
430             lblAuditByGross(Index) = tb!ArchivedBy
440         Else
450             lblAuditByGross(Index) = "???"
460         End If
470         lblGross(Index) = "Changed By:"
480     End If

490 End If

500 Exit Sub

FillArcGross_Error:

    Dim strES As String
    Dim intEL As Integer

510 intEL = Erl
520 strES = Err.Description
530 LogError "frmArchive", "FillArcGross", intEL, strES, sql


End Sub


Private Sub DoDemographics()

    Dim tb As Recordset
    Dim sql As String
    Dim LatestValue(1 To 33) As String
    Dim dbName(1 To 33) As String
    Dim ShowName(1 To 33) As String
    Dim n As Integer
    Dim X As Integer

10  On Error GoTo DoDemographics_Error

20  For n = 1 To 33
30      dbName(n) = Choose(n, "CaseId", "NOPAS", "MRN", "AandENo", _
                           "FirstName", "Surname", "PatientName", _
                           "Address1", "Address2", "Address3", "Address4", "County", _
                           "DateOfBirth", "Age", "Sex", "Clinician", "GP", "Ward", _
                           "Phone", "Source", "Coroner", "NatureOfSpecimen", "SpecimenLabelled", _
                           "ClinicalHistory", "Comments", "AutopsyFor", "AutopsyRequestedBy", _
                           "DateOfDeath", "MothersName", "MothersDOB", "PaedType", "NoHistTaken", "Urgent", "Year")

40      ShowName(n) = Choose(n, "Case Id", "NOPAS Number", "MRN Number", "A&E Number", _
                             "First Name", "Surname", "Patient Name", _
                             "Address", "Address", "Address", "Address", "County", _
                             "Date Of Birth", "Age", "Sex", "Clinician", "GP", "Ward", _
                             "Phone", "Source", "Coroner", "Nature Of Specimen", "Specimen Labelled", _
                             "Clinical Details", "Comments", "Autopsy For", "Autopsy Requested By", _
                             "Date Of Death", "Mothers Name", "Mothers DOB", "PaedType", "No Histology Taken", "Urgent", "Year")
50  Next

60  sql = "SELECT " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(NOPAS)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NOPAS, '<BLANK>') END NOPAS, " & _
          "CASE LTRIM(RTRIM(MRN)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MRN, '<BLANK>') END MRN, " & _
          "CASE LTRIM(RTRIM(AandENo)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AandENo, '<BLANK>') END AandENo, " & _
          "CASE LTRIM(RTRIM(FirstName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(FirstName, '<BLANK>') END FirstName, " & _
          "CASE LTRIM(RTRIM(Surname)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Surname, '<BLANK>') END Surname, " & _
          "CASE LTRIM(RTRIM(PatientName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PatientName, '<BLANK>') END PatientName, " & _
          "CASE LTRIM(RTRIM(Address1)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address1, '<BLANK>') END Address1, " & _
          "CASE LTRIM(RTRIM(Address2)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address2, '<BLANK>') END Address2, " & _
          "CASE LTRIM(RTRIM(Address3)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address3, '<BLANK>') END Address3, " & _
          "CASE LTRIM(RTRIM(Address4)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address4, '<BLANK>') END Address4, " & _
          "CASE LTRIM(RTRIM(County)) WHEN '' THEN '<BLANK>' ELSE ISNULL(County, '<BLANK>') END County, " & _
          "CASE LTRIM(RTRIM(DateOfBirth)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfBirth, 103), '<BLANK>') END DateOfBirth, " & _
          "CASE LTRIM(RTRIM(Age)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Age, '<BLANK>') END Age, " & _
          "CASE LTRIM(RTRIM(Sex)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Sex, '<BLANK>') END Sex, " & _
          "CASE LTRIM(RTRIM(Clinician)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Clinician, '<BLANK>') END Clinician, " & _
          "CASE LTRIM(RTRIM(GP)) WHEN '' THEN '<BLANK>' ELSE ISNULL(GP, '<BLANK>') END GP, " & _
          "CASE LTRIM(RTRIM(Ward)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Ward, '<BLANK>') END Ward, " & _
          "CASE LTRIM(RTRIM(Phone)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Phone, '<BLANK>') END Phone, " & _
          "CASE LTRIM(RTRIM(Source)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Source, '<BLANK>') END Source, " & _
          "CASE LTRIM(RTRIM(Coroner)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Coroner, '<BLANK>') END Coroner, " & _
          "CASE LTRIM(RTRIM(NatureOfSpecimen)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NatureOfSpecimen, '<BLANK>') END NatureOfSpecimen, "

70  sql = sql & _
          "CASE LTRIM(RTRIM(SpecimenLabelled)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SpecimenLabelled, '<BLANK>') END SpecimenLabelled, " & _
          "CASE LTRIM(RTRIM(ClinicalHistory)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ClinicalHistory, '<BLANK>') END ClinicalHistory, " & _
          "CASE LTRIM(RTRIM(Comments)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Comments, '<BLANK>') END Comments, " & _
          "CASE LTRIM(RTRIM(AutopsyFor)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AutopsyFor, '<BLANK>') END AutopsyFor, " & _
          "CASE LTRIM(RTRIM(AutopsyRequestedBy)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AutopsyRequestedBy, '<BLANK>') END AutopsyRequestedBy, " & _
          "CASE LTRIM(RTRIM(DateOfDeath)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfDeath, 103), '<BLANK>') END DateOfDeath, " & _
          "CASE LTRIM(RTRIM(MothersName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MothersName, '<BLANK>') END MothersName, " & _
          "CASE LTRIM(RTRIM(MothersDOB)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), MothersDOB, 103), '<BLANK>') END MothersDOB, " & _
          "CASE LTRIM(RTRIM(PaedType)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PaedType, '<BLANK>') END PaedType, " & _
          "CASE NoHistTaken WHEN 1 THEN 'No Histology Taken' ELSE 'Histology Taken' END NoHistTaken, " & _
          "CASE Urgent WHEN 1 THEN 'Urgent' ELSE 'Not Urgent' END Urgent "
    '"CASE LTRIM(RTRIM(Year)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Year, '<BLANK>') END Year "

    '"CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
     '"CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

80  sql = sql & _
          "FROM Demographics WHERE " & _
          "CaseID = '" & SID & "'"

90  Set tb = New Recordset
100 RecOpenServer 0, tb, sql
110 If Not tb.EOF Then
120     For n = 1 To 33
130         LatestValue(n) = tb(dbName(n)) & ""
140     Next
150 Else
160     For n = 1 To 33
170         LatestValue(n) = "<BLANK>"
180     Next
190 End If

200 sql = "SELECT ArchiveDateTime, " & _
          "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(NOPAS)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NOPAS, '<BLANK>') END NOPAS, " & _
          "CASE LTRIM(RTRIM(MRN)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MRN, '<BLANK>') END MRN, " & _
          "CASE LTRIM(RTRIM(AandENo)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AandENo, '<BLANK>') END AandENo, " & _
          "CASE LTRIM(RTRIM(FirstName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(FirstName, '<BLANK>') END FirstName, " & _
          "CASE LTRIM(RTRIM(Surname)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Surname, '<BLANK>') END Surname, " & _
          "CASE LTRIM(RTRIM(PatientName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PatientName, '<BLANK>') END PatientName, " & _
          "CASE LTRIM(RTRIM(Address1)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address1, '<BLANK>') END Address1, " & _
          "CASE LTRIM(RTRIM(Address2)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address2, '<BLANK>') END Address2, " & _
          "CASE LTRIM(RTRIM(Address3)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address3, '<BLANK>') END Address3, " & _
          "CASE LTRIM(RTRIM(Address4)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Address4, '<BLANK>') END Address4, " & _
          "CASE LTRIM(RTRIM(County)) WHEN '' THEN '<BLANK>' ELSE ISNULL(County, '<BLANK>') END County, " & _
          "CASE LTRIM(RTRIM(DateOfBirth)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfBirth, 103), '<BLANK>') END DateOfBirth, " & _
          "CASE LTRIM(RTRIM(Age)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Age, '<BLANK>') END Age, " & _
          "CASE LTRIM(RTRIM(Sex)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Sex, '<BLANK>') END Sex, " & _
          "CASE LTRIM(RTRIM(Clinician)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Clinician, '<BLANK>') END Clinician, " & _
          "CASE LTRIM(RTRIM(GP)) WHEN '' THEN '<BLANK>' ELSE ISNULL(GP, '<BLANK>') END GP, " & _
          "CASE LTRIM(RTRIM(Ward)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Ward, '<BLANK>') END Ward, " & _
          "CASE LTRIM(RTRIM(Phone)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Phone, '<BLANK>') END Phone, " & _
          "CASE LTRIM(RTRIM(Source)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Source, '<BLANK>') END Source, " & _
          "CASE LTRIM(RTRIM(Coroner)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Coroner, '<BLANK>') END Coroner, " & _
          "CASE LTRIM(RTRIM(NatureOfSpecimen)) WHEN '' THEN '<BLANK>' ELSE ISNULL(NatureOfSpecimen, '<BLANK>') END NatureOfSpecimen, "

210 sql = sql & _
          "CASE LTRIM(RTRIM(SpecimenLabelled)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SpecimenLabelled, '<BLANK>') END SpecimenLabelled, " & _
          "CASE LTRIM(RTRIM(ClinicalHistory)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ClinicalHistory, '<BLANK>') END ClinicalHistory, " & _
          "CASE LTRIM(RTRIM(Comments)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Comments, '<BLANK>') END Comments, " & _
          "CASE LTRIM(RTRIM(AutopsyFor)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AutopsyFor, '<BLANK>') END AutopsyFor, " & _
          "CASE LTRIM(RTRIM(AutopsyRequestedBy)) WHEN '' THEN '<BLANK>' ELSE ISNULL(AutopsyRequestedBy, '<BLANK>') END AutopsyRequestedBy, " & _
          "CASE LTRIM(RTRIM(DateOfDeath)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateOfDeath, 103), '<BLANK>') END DateOfDeath, " & _
          "CASE LTRIM(RTRIM(MothersName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MothersName, '<BLANK>') END MothersName, " & _
          "CASE LTRIM(RTRIM(MothersDOB)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), MothersDOB, 103), '<BLANK>') END MothersDOB, " & _
          "CASE LTRIM(RTRIM(PaedType)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PaedType, '<BLANK>') END PaedType, " & _
          "CASE NoHistTaken WHEN 1 THEN 'No Histology Taken' ELSE 'Histology Taken' END NoHistTaken, " & _
          "CASE Urgent WHEN 1 THEN 'Urgent' ELSE 'Not Urgent' END Urgent "
    '"CASE LTRIM(RTRIM(Year)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Year, '<BLANK>') END Year "

    '"CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
     '"CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

220 sql = sql & _
          "FROM DemographicsAudit WHERE " & _
          "CaseID = '" & SID & "' " & _
          "ORDER BY ArchiveDateTime DESC"

230 Set tb = New Recordset
240 RecOpenServer 0, tb, sql
250 Do While Not tb.EOF
260     rtbDemo.SelColor = vbBlack
270     rtbDemo.SelBold = True
280     rtbDemo.SelText = "Demographics"
290     rtbDemo.SelText = vbCrLf
300     rtbDemo.SelBold = False
310     If Not IsNull(tb!ArchiveDateTime) Then
320         rtbDemo.SelText = "Archive Date/Time: " & Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
330     Else
340         rtbDemo.SelText = "Archive Time not known."
350     End If
360     rtbDemo.SelText = vbCrLf
370     For X = 1 To 33
380         If LatestValue(X) <> tb(dbName(X)) & "" Then
390             rtbDemo.SelText = ShowName(X)
400             rtbDemo.SelText = " changed from "
410             rtbDemo.SelColor = vbRed
420             rtbDemo.SelText = tb.Fields(dbName(X))
430             rtbDemo.SelColor = vbBlack
440             rtbDemo.SelText = " to "
450             rtbDemo.SelColor = vbBlue
460             rtbDemo.SelText = LatestValue(X)
470             rtbDemo.SelColor = vbBlack
480             rtbDemo.SelText = " by "
                'rtbDemo.SelColor = vbGreen
490             rtbDemo.SelText = tb!ArchivedBy
500             rtbDemo.SelColor = vbBlack
510             rtbDemo.SelText = vbCrLf
520         End If
530     Next
540     For X = 1 To 33
550         LatestValue(X) = tb(dbName(X))
560     Next
570     rtbDemo.SelText = vbCrLf
580     tb.MoveNext
590 Loop

600 DoOtherCaseDetails

610 Exit Sub

DoDemographics_Error:

    Dim strES As String
    Dim intEL As Integer

620 intEL = Erl
630 strES = Err.Description
640 LogError "frmTreeArchive", "DoDemographics", intEL, strES, sql


End Sub

Private Sub DoOtherCaseDetails()

    Dim tb As Recordset
    Dim sql As String
    Dim LatestValue(1 To 11) As String
    Dim dbName(1 To 11) As String
    Dim ShowName(1 To 11) As String
    Dim n As Integer
    Dim X As Integer



10  On Error GoTo DoOtherCaseDetails_Error

20  For n = 1 To 11
30      dbName(n) = Choose(n, "CaseId", "State", "SampleTaken", "SampleReceived", "PreReportDate", "ValReportDate", _
                           "Preliminary", "Validated", "LinkedCaseId", "Phase", "WithPathologist")

40      ShowName(n) = Choose(n, "CaseId", "State", "SampleTaken", "SampleReceived", "PreReportDate", "ValReportDate", _
                             "Preliminary", "Validated", "Linked CaseId", "Phase", "With Pathologist")
50  Next

60  sql = "SELECT " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(State)) WHEN '' THEN '<BLANK>' ELSE ISNULL(State, '<BLANK>') END State, " & _
          "CASE LTRIM(RTRIM(SampleTaken)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), SampleTaken, 103) + ' ' + CONVERT(nvarchar(50), SampleTaken, 108), '<BLANK>') END SampleTaken, " & _
          "CASE LTRIM(RTRIM(SampleReceived)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), SampleReceived, 103) + ' ' + CONVERT(nvarchar(50), SampleReceived, 108), '<BLANK>') END SampleReceived, " & _
          "CASE LTRIM(RTRIM(PreReportDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), PreReportDate, 103) + ' ' + CONVERT(nvarchar(50), PreReportDate, 108), '<BLANK>') END PreReportDate, " & _
          "CASE LTRIM(RTRIM(ValReportDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), ValReportDate, 103) + ' ' + CONVERT(nvarchar(50), ValReportDate, 108), '<BLANK>') END ValReportDate, " & _
          "CASE Preliminary WHEN 1 THEN 'Preliminary' ELSE 'Not Preliminary' END Preliminary, " & _
          "CASE Validated WHEN 1 THEN 'Authorised' ELSE 'Not Authorised' END Validated, " & _
          "CASE LTRIM(RTRIM(LinkedCaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(LinkedCaseId, '<BLANK>') END LinkedCaseId, " & _
          "CASE LTRIM(RTRIM(Phase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Phase, '<BLANK>') END Phase, " & _
          "CASE LTRIM(RTRIM(WithPathologist)) WHEN '' THEN '<BLANK>' ELSE ISNULL(WithPathologist, '<BLANK>') END WithPathologist "

    '"CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
     '"CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

70  sql = sql & _
          "FROM Cases WHERE " & _
          "CaseID = '" & SID & "'"

80  Set tb = New Recordset
90  RecOpenServer 0, tb, sql
100 If Not tb.EOF Then
110     For n = 1 To 11
120         LatestValue(n) = tb(dbName(n)) & ""
130     Next
140 Else
150     For n = 1 To 11
160         LatestValue(n) = "<BLANK>"
170     Next
180 End If

190 sql = "SELECT ArchiveDateTime, " & _
          "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(State)) WHEN '' THEN '<BLANK>' ELSE ISNULL(State, '<BLANK>') END State, " & _
          "CASE LTRIM(RTRIM(SampleTaken)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), SampleTaken, 103) + ' ' + CONVERT(nvarchar(50), SampleTaken, 108), '<BLANK>') END SampleTaken, " & _
          "CASE LTRIM(RTRIM(SampleReceived)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), SampleReceived, 103) + ' ' + CONVERT(nvarchar(50), SampleReceived, 108), '<BLANK>') END SampleReceived, " & _
          "CASE LTRIM(RTRIM(PreReportDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), PreReportDate, 103) + ' ' + CONVERT(nvarchar(50), PreReportDate, 108), '<BLANK>') END PreReportDate, " & _
          "CASE LTRIM(RTRIM(ValReportDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), ValReportDate, 103) + ' ' + CONVERT(nvarchar(50), ValReportDate, 108), '<BLANK>') END ValReportDate, " & _
          "CASE Preliminary WHEN 1 THEN 'Preliminary' ELSE 'Not Preliminary' END Preliminary, " & _
          "CASE Validated WHEN 1 THEN 'Authorised' ELSE 'Not Authorised' END Validated, " & _
          "CASE LTRIM(RTRIM(LinkedCaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(LinkedCaseId, '<BLANK>') END LinkedCaseId, " & _
          "CASE LTRIM(RTRIM(Phase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Phase, '<BLANK>') END Phase, " & _
          "CASE LTRIM(RTRIM(WithPathologist)) WHEN '' THEN '<BLANK>' ELSE ISNULL(WithPathologist, '<BLANK>') END WithPathologist "

    '"CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
     '"CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

200 sql = sql & _
          "FROM CasesAudit WHERE " & _
          "CaseID = '" & SID & "' " & _
          "ORDER BY ArchiveDateTime DESC"

210 Set tb = New Recordset
220 RecOpenServer 0, tb, sql
230 Do While Not tb.EOF
240     For X = 1 To 11
250         If LatestValue(X) <> tb(dbName(X)) & "" Then
260             rtbDemo.SelColor = vbBlack
270             rtbDemo.SelBold = True
280             rtbDemo.SelText = "Case Details"
290             rtbDemo.SelText = vbCrLf
300             rtbDemo.SelBold = False
310             If Not IsNull(tb!ArchiveDateTime) Then
320                 rtbDemo.SelText = "Archive Date/Time: " & Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
330             Else
340                 rtbDemo.SelText = "Archive Time not known."
350             End If
360             rtbDemo.SelText = vbCrLf
370             Exit For
380         End If
390     Next

400     For X = 1 To 11
410         If LatestValue(X) <> tb(dbName(X)) & "" Then
420             rtbDemo.SelText = ShowName(X)
430             rtbDemo.SelText = " changed from "
440             rtbDemo.SelColor = vbRed
450             rtbDemo.SelText = tb.Fields(dbName(X))
460             rtbDemo.SelColor = vbBlack
470             rtbDemo.SelText = " to "
480             rtbDemo.SelColor = vbBlue
490             rtbDemo.SelText = LatestValue(X)
500             rtbDemo.SelColor = vbBlack
510             rtbDemo.SelText = " by "
                'rtbDemo.SelColor = vbGreen
520             rtbDemo.SelText = tb!ArchivedBy
530             rtbDemo.SelColor = vbBlack
540             rtbDemo.SelText = vbCrLf
550         End If
560     Next
570     For X = 1 To 11
580         LatestValue(X) = tb(dbName(X))
590     Next
        'rtbDemo.SelText = vbCrLf
600     tb.MoveNext
610 Loop

620 Exit Sub

DoOtherCaseDetails_Error:

    Dim strES As String
    Dim intEL As Integer

630 intEL = Erl
640 strES = Err.Description
650 LogError "frmArchive", "DoOtherCaseDetails", intEL, strES, sql

End Sub

Private Sub DoPCodes()
    Dim tb As Recordset
    Dim sql As String
    Dim LatestValue(1 To 4) As String
    Dim dbName(1 To 4) As String
    Dim ShowName(1 To 4) As String
    Dim n As Integer
    Dim X As Integer

10  On Error GoTo DoCodes_Error

20  For n = 1 To 4
30      dbName(n) = Choose(n, "ListId", "CaseListId", "CaseId", "Type")

40      ShowName(n) = Choose(n, "Code", "CaseListId", "CaseId", "Type")
50  Next

60  sql = "SELECT " & _
          "CASE LTRIM(RTRIM(ListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ListId, '<BLANK>') END ListId, " & _
          "CASE LTRIM(RTRIM(CaseListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseListId, '<BLANK>') END CaseListId, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(Type)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Type, '<BLANK>') END Type "
    '    "CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
         '    "CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

70  sql = sql & _
          "FROM CaseListLink WHERE " & _
          "CaseID = '" & SID & "' AND CaseListId = '' "

80  Set tb = New Recordset
90  RecOpenServer 0, tb, sql
100 If Not tb.EOF Then
110     For n = 1 To 4
120         LatestValue(n) = tb(dbName(n)) & ""
130     Next
140 Else
150     For n = 1 To 4
160         LatestValue(n) = "<BLANK>"
170     Next
180 End If

190 sql = "SELECT ArchiveDateTime, " & _
          "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
          "CASE LTRIM(RTRIM(ListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ListId, '<BLANK>') END ListId, " & _
          "CASE LTRIM(RTRIM(CaseListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseListId, '<BLANK>') END CaseListId, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(Type)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Type, '<BLANK>') END Type "
    '        "CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
             '        "CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

200 sql = sql & _
          "FROM CaseListLinkAudit WHERE " & _
          "CaseID = '" & SID & "' AND CaseListId = '' " & _
          "ORDER BY ArchiveDateTime DESC"

210 Set tb = New Recordset
220 RecOpenServer 0, tb, sql
230 Do While Not tb.EOF
240     rtbCodes.SelColor = vbBlack
250     If Not IsNull(tb!ArchiveDateTime) Then
260         rtbCodes.SelText = "Archive Date/Time: " & Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
270     Else
280         rtbCodes.SelText = "Archive Time not known."
290     End If
300     rtbCodes.SelText = vbCrLf
310     For X = 1 To 4
320         If LatestValue(X) <> tb(dbName(X)) & "" Then
330             rtbCodes.SelText = ShowName(X)
340             rtbCodes.SelText = " changed from "
350             rtbCodes.SelColor = vbRed
360             If ShowName(X) = "Code" Then
370                 rtbCodes.SelText = " " & ListIdToCode(tb.Fields(dbName(X)))
380             Else
390                 rtbCodes.SelText = tb.Fields(dbName(X))
400             End If
410             rtbCodes.SelColor = vbBlack
420             rtbCodes.SelText = " to "
430             rtbCodes.SelColor = vbBlue
440             If ShowName(X) = "Code" And LatestValue(X) <> "<BLANK>" Then
450                 rtbCodes.SelText = " " & ListIdToCode(LatestValue(X))
460             Else
470                 rtbCodes.SelText = LatestValue(X)
480             End If
490             rtbCodes.SelColor = vbBlack
500             rtbCodes.SelText = " by "
                'rtbCodes.SelColor = vbGreen
510             rtbCodes.SelText = tb!ArchivedBy
520             rtbCodes.SelColor = vbBlack
530             rtbCodes.SelText = vbCrLf
540         End If
550     Next
560     For X = 1 To 4
570         LatestValue(X) = tb(dbName(X))
580     Next
590     rtbCodes.SelText = vbCrLf
600     tb.MoveNext
610 Loop
    'rsR.MoveNext
    'Loop

620 Exit Sub

DoCodes_Error:

    Dim strES As String
    Dim intEL As Integer

630 intEL = Erl
640 strES = Err.Description
650 LogError "frmArchive", "DoPCodes", intEL, strES, sql


End Sub

Private Sub DoOtherCodes()
    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo DoOtherCodes_Error

20  sql = "SELECT ArchiveDateTime, " & _
          "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
          "CASE LTRIM(RTRIM(ListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ListId, '<BLANK>') END ListId, " & _
          "CASE LTRIM(RTRIM(CaseListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseListId, '<BLANK>') END CaseListId, " & _
          "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
          "CASE LTRIM(RTRIM(Type)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Type, '<BLANK>') END Type, " & _
          "CASE LTRIM(RTRIM(TissueTypeId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(TissueTypeId, '<BLANK>') END TissueTypeId "

    '        "CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
             '        "CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

30  sql = sql & _
          "FROM CaseListLinkAudit WHERE " & _
          "CaseID = '" & SID & "' AND CaseListId <> '' " & _
          "ORDER BY ArchiveDateTime DESC"

40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql
60  Do While Not tb.EOF
70      rtbCodes.SelColor = vbBlack
80      If Not IsNull(tb!ArchiveDateTime) Then
90          rtbCodes.SelText = "Archive Date/Time: " & Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
100     Else
110         rtbCodes.SelText = "Archive Time not known."
120     End If
130     rtbCodes.SelText = vbCrLf
        '        For x = 1 To 4
        'If LatestValue(x) <> tb(dbName(x)) & "" Then
140     rtbCodes.SelColor = vbRed
150     rtbCodes.SelText = ListIdToCode(tb.Fields("ListId"))
160     rtbCodes.SelColor = vbBlack
170     rtbCodes.SelText = " was deleted "

        '            If ShowName(x) = "Code" Then
        '               rtbCodes.SelText = " " & ListIdToCode(tb.Fields(dbName(x)))
        '            Else
        '              rtbCodes.SelText = tb.Fields(dbName(x))
        '            End If
        '            rtbCodes.SelColor = vbBlack
        '            rtbCodes.SelText = " to "
        '            rtbCodes.SelColor = vbBlue
        '            If ShowName(x) = "Code" Then
        '               rtbCodes.SelText = " " & ListIdToCode(LatestValue(x))
        '            Else
        '              rtbCodes.SelText = LatestValue(x)
        '            End If
        'rtbCodes.SelColor = vbBlack

180     If tb.Fields("Type") & "" = "M" Then
190         rtbCodes.SelText = " from "
200         rtbCodes.SelColor = vbRed
210         rtbCodes.SelText = ListIdToCode(tb.Fields("TissueTypeId"))
220         rtbCodes.SelColor = vbBlack
230     End If

240     rtbCodes.SelText = " by "
        'rtbCodes.SelColor = vbGreen
250     rtbCodes.SelText = tb!ArchivedBy
260     rtbCodes.SelText = vbCrLf
        '          End If
        '        Next
        '        For x = 1 To 4
        '          LatestValue(x) = tb(dbName(x))
        '        Next
270     rtbCodes.SelText = vbCrLf
280     tb.MoveNext
290 Loop

300 Exit Sub

DoOtherCodes_Error:

    Dim strES As String
    Dim intEL As Integer

310 intEL = Erl
320 strES = Err.Description
330 LogError "frmArchive", "DoOtherCodes", intEL, strES, sql

End Sub


Private Sub DoMovement()
    Dim tb As Recordset
    Dim sql As String
    Dim LatestValue(1 To 8) As String
    Dim dbName(1 To 8) As String
    Dim ShowName(1 To 8) As String
    Dim n As Integer
    Dim X As Integer
    Dim SqlT As String
    Dim rsR As Recordset

10  On Error GoTo DoMovement_Error

20  For n = 1 To 8
30      dbName(n) = Choose(n, "CaseId", "CaseListId", "Code", "Type", "Destination", "DateSent", "DateReceived", "Agreed")

40      ShowName(n) = Choose(n, "CaseId", "CaseListId", "Code", "Type", "Destination", "Date Sent", "Date Received", "Agreed")
50  Next

60  SqlT = "select distinct caselistId from CaseMovements where Caseid = '" & SID & "'"
70  Set rsR = New Recordset
80  RecOpenServer 0, rsR, SqlT
90  Do While Not rsR.EOF

100     sql = "SELECT " & _
              "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
              "CASE LTRIM(RTRIM(CaseListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseListId, '<BLANK>') END CaseListId, " & _
              "CASE LTRIM(RTRIM(Code)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Code, '<BLANK>') END Code, " & _
              "CASE LTRIM(RTRIM(Type)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Type, '<BLANK>') END Type, " & _
              "CASE LTRIM(RTRIM(Destination)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Destination, '<BLANK>') END Destination, " & _
              "CASE LTRIM(RTRIM(DateSent)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateSent, 103) + ' ' + CONVERT(nvarchar(50), DateSent, 108), '<BLANK>') END DateSent, " & _
              "CASE LTRIM(RTRIM(DateReceived)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateReceived, 103) + ' ' + CONVERT(nvarchar(50), DateReceived, 108), '<BLANK>') END DateReceived, " & _
              "CASE Agreed WHEN '1' THEN 'Agreed' ELSE 'Not Agreed' END Agreed "
        '    "CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
             '    "CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "
        '
110     sql = sql & _
              "FROM CaseMovements WHERE " & _
              "CaseID = '" & SID & "' AND CaseListId = '" & rsR!CaseListId & "' "

120     Set tb = New Recordset
130     RecOpenServer 0, tb, sql
140     If Not tb.EOF Then
150         For n = 1 To 8
160             LatestValue(n) = tb(dbName(n)) & ""
170         Next
180     Else
190         For n = 1 To 8
200             LatestValue(n) = "<BLANK>"
210         Next
220     End If

230     sql = "SELECT ArchiveDateTime, " & _
              "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
              "CASE LTRIM(RTRIM(CaseId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseId, '<BLANK>') END CaseId, " & _
              "CASE LTRIM(RTRIM(CaseListId)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CaseListId, '<BLANK>') END CaseListId, " & _
              "CASE LTRIM(RTRIM(Code)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Code, '<BLANK>') END Code, " & _
              "CASE LTRIM(RTRIM(Type)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Type, '<BLANK>') END Type, " & _
              "CASE LTRIM(RTRIM(Destination)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Destination, '<BLANK>') END Destination, " & _
              "CASE LTRIM(RTRIM(DateSent)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateSent, 103) + ' ' + CONVERT(nvarchar(50), DateSent, 108), '<BLANK>') END DateSent, " & _
              "CASE LTRIM(RTRIM(DateReceived)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateReceived, 103) + ' ' + CONVERT(nvarchar(50), DateReceived, 108), '<BLANK>') END DateReceived, " & _
              "CASE Agreed WHEN '1' THEN 'Agreed' ELSE 'Not Agreed' END Agreed "
        '        "CASE LTRIM(RTRIM(Username)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Username, '<BLANK>') END Username, " & _
                 '        "CASE LTRIM(RTRIM(DateTimeOfRecord)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DateTimeOfRecord, 103), '<BLANK>') END DateTimeOfRecord "

240     sql = sql & _
              "FROM CaseMovementsAudit WHERE " & _
              "CaseID = '" & SID & "' AND CaseListId = '" & rsR!CaseListId & "' " & _
              "ORDER BY ArchiveDateTime DESC"

250     Set tb = New Recordset
260     RecOpenServer 0, tb, sql
270     Do While Not tb.EOF
280         rtbMovement.SelColor = vbBlack
290         If Not IsNull(tb!ArchiveDateTime) Then
300             rtbMovement.SelText = "Archive Date/Time: " & Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
310         Else
320             rtbMovement.SelText = "Archive Time not known."
330         End If
340         rtbMovement.SelText = vbCrLf
350         For X = 1 To 8
360             If LatestValue(X) <> tb(dbName(X)) & "" Then
370                 rtbMovement.SelText = ShowName(X)
380                 rtbMovement.SelText = " changed from "
390                 rtbMovement.SelColor = vbRed
400                 rtbMovement.SelText = tb.Fields(dbName(X))
410                 rtbMovement.SelColor = vbBlack
420                 rtbMovement.SelText = " to "
430                 rtbMovement.SelColor = vbBlue
440                 rtbMovement.SelText = LatestValue(X)
450                 rtbMovement.SelColor = vbBlack
460                 rtbMovement.SelText = " by "
                    'rtbMovement.SelColor = vbGreen
470                 rtbMovement.SelText = tb!ArchivedBy
480                 rtbMovement.SelColor = vbBlack
490                 rtbMovement.SelText = vbCrLf
500             End If
510         Next
520         For X = 1 To 8
530             LatestValue(X) = tb(dbName(X))
540         Next
550         rtbMovement.SelText = vbCrLf
560         tb.MoveNext
570     Loop
580     rsR.MoveNext
590 Loop

600 Exit Sub

DoMovement_Error:

    Dim strES As String
    Dim intEL As Integer

610 intEL = Erl
620 strES = Err.Description
630 LogError "frmArchive", "DoMovement", intEL, strES, sql

End Sub

Private Function ListIdToCode(ListId As String) As String
    Dim tb As New Recordset
    Dim sql As String


10  On Error GoTo ListIdToCode_Error

20  ListIdToCode = "???"

30  sql = "SELECT Code, Description FROM Lists WHERE ListId = '" & ListId & "'"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql
60  If Not tb.EOF Then
70      ListIdToCode = tb!Code & " - " & tb!Description & ""
80  Else
90      ListIdToCode = "???"
100 End If

110 Exit Function

ListIdToCode_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmArchive", "ListIdToCode", intEL, strES, sql


End Function

Private Sub HighlightChanges()

    Dim node0 As MSComctlLib.Node
    Dim node1 As MSComctlLib.Node
    Dim found As Boolean

10  For Each node0 In tvArcTree(0).Nodes

20      For Each node1 In tvArcTree(1).Nodes
30          If node0.Key = node1.Key Then
40              found = True
50              Exit For
60          End If
70      Next
80      If found = False Then
90          node0.BackColor = vbRed
100     Else
110         node0.BackColor = vbWhite
120     End If

130     found = False
140 Next

150 For Each node1 In tvArcTree(1).Nodes

160     For Each node0 In tvArcTree(0).Nodes
170         If node0.Key = node1.Key Then
180             found = True
190             Exit For
200         End If
210     Next
220     If found = False Then
230         node1.BackColor = vbRed
240     Else
250         node1.BackColor = vbWhite
260     End If

270     found = False
280 Next

End Sub



Private Sub txtCaseId_Change()
10  If Len(txtCaseId.Text) = txtCaseId.MaxLength Then cmdStart.SetFocus
End Sub


Private Sub txtCaseId_KeyPress(KeyAscii As Integer)

10  If UCase(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
20      Call ValidateTullCaseId(KeyAscii, Me)
30  Else
40      Call ValidateLimCaseId(KeyAscii, Me)
50  End If
End Sub

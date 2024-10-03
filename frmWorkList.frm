VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmWorklist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13710
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   13710
   Begin VB.CommandButton Command 
      Height          =   735
      Left            =   14160
      TabIndex        =   26
      Top             =   2640
      Width           =   375
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   15
      Top             =   3000
   End
   Begin VB.ComboBox cmbWithPathologist 
      Height          =   315
      ItemData        =   "frmWorkList.frx":0ECA
      Left            =   10920
      List            =   "frmWorkList.frx":0ECC
      TabIndex        =   0
      Text            =   "cmbWithPathologist"
      Top             =   0
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13320
      Top             =   720
   End
   Begin VB.Frame fraWorkList 
      Height          =   9615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   13335
      Begin VB.CommandButton cmdChangePath 
         Height          =   315
         Left            =   1920
         Picture         =   "frmWorkList.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh Lists"
         Height          =   615
         Left            =   6105
         Picture         =   "frmWorkList.frx":33939
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8880
         Width           =   1695
      End
      Begin VB.PictureBox fraProgress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3135
         ScaleHeight     =   585
         ScaleWidth      =   2760
         TabIndex        =   22
         Top             =   8880
         Visible         =   0   'False
         Width           =   2790
         Begin MSComctlLib.ProgressBar pbProgress 
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   300
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Min             =   1
         End
         Begin VB.Label lblProgress 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Refreshing ..."
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   60
            Width           =   2775
         End
      End
      Begin VB.CommandButton cmdExtensiveSearch 
         Caption         =   "Search"
         DownPicture     =   "frmWorkList.frx":34023
         Height          =   615
         Left            =   7920
         Picture         =   "frmWorkList.frx":34125
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   8880
         Width           =   1695
      End
      Begin VB.CommandButton cmdInHistology 
         Caption         =   "In Histology"
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddDemo 
         Caption         =   "Add Demographics"
         DownPicture     =   "frmWorkList.frx":344A3
         Height          =   615
         Left            =   9720
         Picture         =   "frmWorkList.frx":345A5
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   8880
         Width           =   1695
      End
      Begin MSComctlLib.ListView lstWorkList 
         Height          =   8175
         Index           =   5
         Left            =   11040
         TabIndex        =   16
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
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
            Text            =   "ExternalEvents"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkList 
         Height          =   8175
         Index           =   4
         Left            =   8880
         TabIndex        =   15
         Top             =   585
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
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
            Text            =   "SpecialStain"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkList 
         Height          =   8175
         Index           =   3
         Left            =   6720
         TabIndex        =   14
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
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
            Text            =   "AuthNotPrinted"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkList 
         Height          =   8175
         Index           =   2
         Left            =   4560
         TabIndex        =   13
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
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
            Text            =   "AwaitAuthorisation"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkList 
         Height          =   8175
         Index           =   1
         Left            =   2400
         TabIndex        =   12
         Top             =   600
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
            Text            =   "WithPathologist"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkList 
         Height          =   8175
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "InHistology"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdWorkSheet 
         Caption         =   "Go To Worksheet"
         Height          =   615
         Left            =   11520
         Picture         =   "frmWorkList.frx":3467D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8880
         Width           =   1575
      End
      Begin VB.Image imgRemoveExtraRequests 
         Height          =   225
         Left            =   10710
         Picture         =   "frmWorkList.frx":34BAF
         Top             =   375
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblLoggedIn 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   9120
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Logged In : "
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   9120
         Width           =   1335
      End
      Begin VB.Label lblWorkList 
         AutoSize        =   -1  'True
         Caption         =   "In Histology"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblWorkList 
         AutoSize        =   -1  'True
         Caption         =   "With Pathologist"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblWorkList 
         AutoSize        =   -1  'True
         Caption         =   "Awaiting Authorisation"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label lblWorkList 
         AutoSize        =   -1  'True
         Caption         =   "Authorised - Not Printed"
         Height          =   195
         Index           =   3
         Left            =   6720
         TabIndex        =   4
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label lblWorkList 
         AutoSize        =   -1  'True
         Caption         =   "Extra Requests"
         Height          =   195
         Index           =   4
         Left            =   8880
         TabIndex        =   3
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblWorkList 
         AutoSize        =   -1  'True
         Caption         =   "External Events - Out"
         Height          =   195
         Index           =   5
         Left            =   11040
         TabIndex        =   2
         Top             =   360
         Width           =   1500
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   5040
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "St. Johns Histology System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   435
      Left            =   4200
      TabIndex        =   9
      Top             =   960
      Width           =   4545
   End
   Begin VB.Image imgMain 
      Height          =   1395
      Left            =   5040
      Picture         =   "frmWorkList.frx":34E85
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log &Off"
      End
      Begin VB.Menu mnuAudit 
         Caption         =   "&Audit"
      End
      Begin VB.Menu mnuNull 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLists 
      Caption         =   "&Lists"
      Begin VB.Menu mnuPCodes 
         Caption         =   "&P Codes"
      End
      Begin VB.Menu mnuMCodes 
         Caption         =   "&M Codes"
      End
      Begin VB.Menu mnuQCodes 
         Caption         =   "&Q Codes"
      End
      Begin VB.Menu mnuTCodes 
         Caption         =   "&T Codes"
      End
      Begin VB.Menu mnuChangeCodes 
         Caption         =   "&Change P/M/Q Codes"
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReferrals 
         Caption         =   "&Referrals"
         Begin VB.Menu mnuReferredTo 
            Caption         =   "Referred To"
         End
         Begin VB.Menu mnuReasonReferral 
            Caption         =   "Reason For Referral"
         End
      End
      Begin VB.Menu mnuDestStains 
         Caption         =   "Stains"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDestCodes 
         Caption         =   "Codes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuStains 
         Caption         =   "&Stains"
         Begin VB.Menu mnuRoutineStainList 
            Caption         =   "Routine"
         End
         Begin VB.Menu mnuSpecialStainList 
            Caption         =   "Special"
         End
         Begin VB.Menu mnuImmunoStainList 
            Caption         =   "Immunohistochemical"
         End
      End
      Begin VB.Menu mnuOther 
         Caption         =   "&Other"
         Begin VB.Menu mnuWards 
            Caption         =   "Wards"
         End
         Begin VB.Menu mnuCoroners 
            Caption         =   "Coroners"
         End
         Begin VB.Menu mnuClinicians 
            Caption         =   "Clinicians"
         End
         Begin VB.Menu mnuGPs 
            Caption         =   "GP's"
         End
         Begin VB.Menu mnuCounty 
            Caption         =   "County"
         End
         Begin VB.Menu mnuSource 
            Caption         =   "Source"
         End
         Begin VB.Menu mnuOrientation 
            Caption         =   "Orientation"
         End
         Begin VB.Menu mnuProcessor 
            Caption         =   "Processor"
         End
         Begin VB.Menu mnuDiscrepancy 
            Caption         =   "&Discrepancy"
            Begin VB.Menu mnuDiscrepancyType 
               Caption         =   "&Type"
            End
            Begin VB.Menu mnuDiscrepancyResolution 
               Caption         =   "&Resolution"
            End
         End
         Begin VB.Menu mnuAccreditationSettings 
            Caption         =   "Accreditation Settings"
         End
         Begin VB.Menu mnuNonWorkingDays 
            Caption         =   "Non Working Days"
         End
      End
   End
   Begin VB.Menu mnuWorkLogs 
      Caption         =   "&WorkLogs"
      Begin VB.Menu mnuCutUpSheet 
         Caption         =   "&Cut-Up"
      End
      Begin VB.Menu mnuEmbedSheet 
         Caption         =   "&Embedding"
      End
      Begin VB.Menu mnuCutting 
         Caption         =   "&Cutting"
      End
      Begin VB.Menu mnuDisposal 
         Caption         =   "&Disposal"
         Begin VB.Menu mnuHistDisposal 
            Caption         =   "&Histology"
         End
         Begin VB.Menu mnuCytDisposal 
            Caption         =   "&Cytology"
         End
         Begin VB.Menu mnuAutDisposal 
            Caption         =   "&Autopsy"
         End
      End
      Begin VB.Menu mnuReferral 
         Caption         =   "&Referral"
      End
      Begin VB.Menu mnuLocked4Editing 
         Caption         =   "&Locked for Editing"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuNcri 
         Caption         =   "&NCRI"
      End
      Begin VB.Menu mnuSnomed 
         Caption         =   "&Snomed Search"
         Begin VB.Menu mnuSearch1 
            Caption         =   "&By Tissue Type "
         End
         Begin VB.Menu mnuSearch2 
            Caption         =   "&Diagnosis Specific Totals "
         End
         Begin VB.Menu mnuSearch3 
            Caption         =   "&Diagnosis Range Search "
         End
         Begin VB.Menu mnuSearch4 
            Caption         =   "&Grouped Tissue Search "
         End
         Begin VB.Menu mnuSearch5 
            Caption         =   "&Location Specific Search "
         End
      End
      Begin VB.Menu mnuStats 
         Caption         =   "&Statistics"
         Begin VB.Menu mnuNumerical 
            Caption         =   "&Numerical"
         End
         Begin VB.Menu mnuTAT 
            Caption         =   "&TAT"
            Begin VB.Menu mnuTatPCodes 
               Caption         =   "P Codes"
            End
            Begin VB.Menu mnuTatTCodes 
               Caption         =   "T Codes"
            End
            Begin VB.Menu mnuTATPCodeCaseIds 
               Caption         =   "P Codes Case Ids"
            End
            Begin VB.Menu mnuTATTCodeCaseIds 
               Caption         =   "T Codes Case Ids"
            End
         End
      End
      Begin VB.Menu mnuRptDiscrep 
         Caption         =   "&Discrepancy"
      End
      Begin VB.Menu mnuAuthorisedReports 
         Caption         =   "Authorised Reports"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "Language"
      Visible         =   0   'False
      Begin VB.Menu mnuLanguageEnglish 
         Caption         =   "English"
      End
      Begin VB.Menu mnuLanguageRussian 
         Caption         =   "Russian"
      End
      Begin VB.Menu mnuLanguagePortuguese 
         Caption         =   "Portuguese"
      End
   End
End
Attribute VB_Name = "frmWorklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sCaseId As String
Public PathologistCode As String
Public PathologistName As String
Public blnWorklistLoadedAlready As Boolean
    Dim LoadFlag As Boolean
Private Sub cmbWithPathologist_Change()



    'FillList PathologistCode
End Sub

Private Sub cmbWithPathologist_Click()
    Dim sql As String
    Dim tb As Recordset
    Dim userCode As String

10  On Error GoTo cmbWithPathologist_Click_Error

20  If cmbWithPathologist <> "All" Then
30      sql = "SELECT * FROM Users WHERE UserId = " & cmbWithPathologist.ItemData(cmbWithPathologist.ListIndex)
40      Set tb = New Recordset
50      RecOpenServer 0, tb, sql
60      If Not tb.EOF Then
70          PathologistCode = tb!Code & ""
80          PathologistName = tb!UserName & ""
90      End If
100 Else
110     PathologistCode = ""
120     PathologistName = ""
130     imgRemoveExtraRequests.Visible = False
140 End If

150 blnWorklistLoadedAlready = False
160 'tmrRefresh.Enabled = True

    '170   If cmbWithPathologist.Text <> "All" Then
180 sql = "SELECT Code FROM Users WHERE UserName = '" & AddTicks(Trim(cmbWithPathologist.Text)) & "'"

190 Set tb = New Recordset
200 RecOpenServer 0, tb, sql

210 If Not tb Is Nothing Then
220     If Not tb.EOF Then
230         userCode = tb!Code
240     End If
250 End If

    ''260       sql = "SELECT CaseID FROM Cases WHERE State = 'With Pathologist' AND WithPathologist = '" & Trim(UserCode) & "'"
    ''270       Set tb = New Recordset
    ''280       RecOpenServer 0, tb, sql
    ''290       lstWorkList(1).ListItems.Clear
    ''300       If Not tb Is Nothing Then
    ''310         Do While Not tb.EOF
    ''320              If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
    ''330                     lstWorkList(1).ListItems.Add 1, , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
    ''340              Else
    ''350                     lstWorkList(1).ListItems.Add 1, , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
    ''360              End If
    ''370         tb.MoveNext
    ''380         Loop
    ''
    ''390       End If
    ''
    ''
    ''400       sql = "SELECT CaseID FROM Cases WHERE State = 'Awaiting Authorisation' AND WithPathologist = '" & Trim(UserCode) & "'"
    ''410       Set tb = New Recordset
    ''420       RecOpenServer 0, tb, sql
    ''430       lstWorkList(2).ListItems.Clear
    ''440       If Not tb Is Nothing Then
    ''450         Do While Not tb.EOF
    ''460              If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
    ''470                     lstWorkList(2).ListItems.Add 1, , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
    ''480              Else
    ''490                     lstWorkList(2).ListItems.Add 1, , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
    ''500              End If
    ''510          tb.MoveNext
    ''520         Loop
    ''
    ''530       End If
    ''540   Else
550 FillList userCode
    '560   End If

570 FillExternal userCode
580 FillExtraRequests userCode
590 FillValNotPrint AddTicks(Trim(cmbWithPathologist.Text))


600 Exit Sub

cmbWithPathologist_Click_Error:

    Dim strES As String
    Dim intEL As Integer

610 intEL = Erl
620 strES = Err.Description
630 LogError "frmWorkList", "cmbWithPathologist_Click", intEL, strES, sql


End Sub

Private Sub cmdAddDemo_Click()
10  With frmDemographics
20      .AddNew = True
30      .Link = False
40      .Show 1
50  End With
End Sub

Private Sub cmdExtensiveSearch_Click()
10  With frmSearch
20      .FromEdit = False
30      .Show 1
40  End With
End Sub

Private Sub cmdInHistology_Click()
10  With frmPhase
20      .Show 1
30  End With
40  'frmWorklist.Enabled = False
End Sub


Private Sub cmdRefresh_Click()

10  blnWorklistLoadedAlready = False
20  tmrRefresh.Enabled = True

End Sub

Private Sub cmdWorkSheet_Click()

10  With frmWorkSheet
20      .txtCaseId = sCaseId
30      .Show
40      .cmbPatientId.SetFocus
50  End With
60  'frmWorklist.Enabled = False

End Sub

'Private Sub Command_Click(Index As Integer)
'Dim Sql As String
'Dim tb As New Recordset
'Dim Sql2 As String

'ITS 819211

'If Index = 0 Then
'    Sql = "select C.CaseId, C.LinkedCaseId from Cases as C, CaseListLink as CLK where C.CaseId = clk.CaseId and C.CaseId like 'H%' " & _
     "and clk.Type = 'P' and clk.ListId = '2722' and C.LinkedCaseId in (Select CaseId from CaseListLink " & _
     "where CaseId =  C.LinkedCaseId and ListId = '1932')"

'    Set tb = New Recordset
'    RecOpenClient 0, tb, Sql

'    Do While Not tb.EOF
'        Sql2 = "Update CaseListLink SET ListId = '3218' WHERE (CaseId = '" & Trim$(tb!CaseId) & "" & "') AND (Type = 'P') AND (ListId = '2722')"
'        Cnxn(0).Execute Sql2

'        tb.MoveNext
'    Loop
'Else
'    Sql = "select C.CaseId, C.LinkedCaseId from Cases as C, CaseListLink as CLK where C.CaseId = clk.CaseId and C.CaseId like 'H%' " & _
     "and clk.Type = 'P' and clk.ListId = '2722' and C.LinkedCaseId in (Select CaseId from CaseListLink " & _
     "where CaseId =  C.LinkedCaseId and ListId = '1933')"



Private Sub Command_Click()
    With FrmCaseStatus
        .Show
    End With
End Sub

Private Sub cmdChangePath_Click()
    Dim sql As String
    Dim tb As Recordset
    Dim i As Integer
    Dim CaseId As String
    Dim TempCaseId As String
    Dim TempCaseIddub As String
    Dim CaseCollection As New Collection

    On Error GoTo cmdChangePath_Click_Error



    If cmdChangePath.Left = 1920 Then
        With frmChange
            .WithPath = False
        End With

        For i = 1 To lstWorkList(0).ListItems.Count
            If lstWorkList(0).ListItems(i).Selected Then
                CaseId = Replace(lstWorkList(0).ListItems(i).Text, "/", "")
                CaseId = Replace(CaseId, " ", "")
                sql = "SELECT C.SampleTaken, D.PatientName FROM Cases C INNER JOIN Demographics D ON C.CaseID = D.CaseID WHERE C.CaseID ='" & Trim(CaseId) & "'"
                Set tb = New Recordset
                RecOpenServer 0, tb, sql
                If Not tb Is Nothing Then
                    If Not tb.EOF Then
                        Dim CaseInfo(2) As String
                        If Mid(CaseId & "", 2, 1) = "P" Or Mid(CaseId & "", 2, 1) = "A" Then
                            TempCaseId = Left(CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(CaseId, 2)
                        Else
                            TempCaseId = Left(CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(CaseId, 2)
                        End If
                        CaseInfo(0) = TempCaseId
                        CaseInfo(1) = Trim(tb!PatientName)
                        CaseInfo(2) = Format$(CDate(tb!SampleTaken), "YYYY-MM-DD HH:MM")
                        CaseCollection.Add CaseInfo
                    End If
                End If
            End If
        Next i

    ElseIf cmdChangePath.Left = 4020 Then
        With frmChange
            .WithPath = True
        End With

        For i = 1 To lstWorkList(1).ListItems.Count
            If lstWorkList(1).ListItems(i).Selected Then
                CaseId = Replace(lstWorkList(1).ListItems(i).Text, "/", "")
                CaseId = Replace(CaseId, " ", "")
                sql = "SELECT d.PatientName, u.UserName, c.SampleTaken FROM Cases c JOIN Demographics d ON c.CaseID = d.CaseID JOIN Users u ON c.WithPathologist = u.Code WHERE d.CaseID = '" & Trim(CaseId) & "'"
                Set tb = New Recordset
                RecOpenServer 0, tb, sql
                If Not tb Is Nothing Then
                    If Not tb.EOF Then
                        Dim CaseInfoid(3) As String
                        CaseInfoid(0) = CaseId
                        CaseInfoid(1) = Trim(tb!PatientName)
                        CaseInfoid(2) = Trim(tb!UserName)
                        CaseInfoid(3) = Format$(CDate(tb!SampleTaken), "YYYY-MM-DD HH:MM")
                        CaseCollection.Add CaseInfoid
                    End If
                End If
            End If
        Next i



    End If
    With frmChange
        Set .CInfo = CaseCollection
        .Show 1
    End With

    Exit Sub

cmdChangePath_Click_Error:
    Dim strES As String
    Dim intEL As Long

    intEL = Erl
    strES = Err.Description
    LogError "frmEditAll", "LoadAllDetails", intEL, strES
    Resume Next


End Sub

'    Set tb = New Recordset
'    RecOpenClient 0, tb, Sql

'    Do While Not tb.EOF
'        Sql2 = "Update CaseListLink SET ListId = '3219' WHERE (CaseId = '" & Trim$(tb!CaseId) & "" & "') AND (Type = 'P') AND (ListId = '2722')"
'        Cnxn(0).Execute Sql2

'        tb.MoveNext
'    Loop
'End If

'End Sub


Private Sub Form_Activate()
'      'FillList
'      'DoEvents
'      frmWorklist.Enabled = True
'10    If blnWorklistLoadedAlready = False Then
'
    If Not LoadFlag Then
201     FillList
301     FillExternal
401     FillExtraRequests
501     FillValNotPrint
        LoadFlag = True
    End If
    tmrRefresh.Enabled = False
    '        blnWorklistLoadedAlready = True
    '60    End If
    '      cmdChangePath.Visible = False
    '      'Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
    '
    '
    ''10    frmWorklist_ChangeLanguage
    '          'tmrRefresh.Enabled = True
End Sub

Private Sub Form_Load()




'
    Dim i As Integer
    '
    '

10  On Error GoTo Form_Load_Error
    LoadFlag = False
20  blnWorklistLoadedAlready = False
30  pBar.Visible = False
    '
    '
40  ChangeFont Me, "Arial"
    fillcombo
60  Select Case UCase$(UserMemberOf)
    Case "LOOKUP"
70      DisableLookup
80  Case "CLERICAL"
90      DisableClerical
100 Case "SCIENTIST"
110     DisableScientist
120 Case "MANAGER"
        'DisableManager
130 Case "CONSULTANT"
140     DisableConsultant
150 Case "SPECIALIST REGISTRAR"
160     DisableConsultant
170 Case "NCRI"
180     DisableNCRI
190 End Select

200 lblLoggedIn = UserName
    'Comment
210 If Val(GetOptionSetting("DemographicEntry", "0")) = 0 Then
220     cmdAddDemo.Visible = False
230     cmdExtensiveSearch.Left = 9720
240 End If

250 For i = 0 To 5
260     Set lstWorkList(i).SelectedItem = Nothing
270 Next

280 Me.Caption = "NetAcquire - " & "Cellular Pathology" & " . Version " & App.Major & "." & App.Minor

330 If blnIsTestMode Then EnableTestMode Me



Form_Load_Error:
    Dim strES As String
    Dim intEL As Long

340 intEL = Erl
350 strES = Err.Description
360 LogError "frmWorkList", "Form_Load", intEL, strES



End Sub

Private Sub DisableNCRI()
10  cmdAddDemo.Visible = False
20  cmdWorkSheet.Visible = False
30  cmdExtensiveSearch.Left = 11520
40  mnuWorkLogs.Visible = False
50  mnuLists.Visible = False
60  mnuAudit.Visible = False
End Sub

Private Sub DisableLookup()

10  cmdAddDemo.Visible = False
20  cmdWorkSheet.Visible = False
30  cmdExtensiveSearch.Left = 11520
40  mnuReports.Visible = False
50  mnuWorkLogs.Visible = False
60  mnuLists.Visible = False
70  mnuAudit.Visible = False

End Sub

Private Sub DisableConsultant()
10  cmdAddDemo.Enabled = False
20  mnuAudit.Visible = False
30  mnuLists.Visible = False
End Sub



Private Sub DisableScientist()
10  mnuAudit.Visible = False
20  mnuLists.Visible = False
End Sub

Private Sub DisableClerical()
10  mnuReports.Visible = False
20  mnuWorkLogs.Visible = False
30  mnuLists.Visible = False
40  mnuAudit.Visible = False
End Sub

Private Sub fillcombo()
    Dim sql As String
    Dim tb As Recordset
    Dim i As Integer
    Dim bInFound As Boolean

10  On Error GoTo FillCombo_Error

20  sql = "SELECT UserName FROM Users WHERE AccessLevel = 'Consultant' AND InUse = 1"

30  Set tb = New Recordset

40  RecOpenServer 0, tb, sql
    'tb.MoveFirst
    cmbWithPathologist.Clear
50  cmbWithPathologist.AddItem "All"
60  cmbWithPathologist.Text = "All"

70  If Not tb Is Nothing Then
        cmbWithPathologist.Visible = False
80      Do While Not tb.EOF
90          cmbWithPathologist.AddItem tb!UserName & ""
100         tb.MoveNext
110     Loop
120 End If
130 tb.Close
140 Set tb = Nothing
    cmbWithPathologist.Visible = True
150 Exit Sub

FillCombo_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmWorkList", "FillCombo", intEL, strES, sql

End Sub

Private Sub Form_Resize()

    On Error Resume Next
20  Me.Top = 0
30  Me.Left = Screen.Width / 2 - Me.Width / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
10  UserName = ""
20  UpdateLoggedOnUsers
30  Do While Forms.Count > 1
40      Unload Forms(Forms.Count - 1)
50  Loop
End Sub



Private Sub imgRemoveExtraRequests_Click()
    Dim intL As Integer
    Dim blnItemSelected As Boolean
    Dim strCaseId As String
    Dim strReason As String

10  For intL = 1 To lstWorkList(4).ListItems.Count    'Find selected CaseId
20      If lstWorkList(4).ListItems(intL).Selected Then
30          blnItemSelected = True
40          Exit For
50      End If
60  Next

70  strCaseId = Replace(Trim$(lstWorkList(4).ListItems(intL)), " ", "")    'Remove
80  strCaseId = Replace(strCaseId, "/", "")                             'space and /

90  If blnItemSelected Then
100     If iMsg("Remove" & " " & Trim$(lstWorkList(4).ListItems(intL)), vbQuestion + vbYesNo) = vbYes Then    'ask user
110         strReason = iBOX("Enter reason for removal:", , "", False)      'Enter Reason
120         CaseAddLogEvent strCaseId, ExtraRequestsRemoved, strReason  'Log Event
130         ClearExtraRequestForCaseId Trim$(lstWorkList(4).ListItems(intL))    'Remove Extras for Case Id
140         lstWorkList(4).ListItems.Clear    'Clear Extra Requests list
150         FillExtraRequests PathologistCode    'Fill Extra Requests again
160         imgRemoveExtraRequests.Visible = False
170     End If
180 End If

End Sub

Private Sub ClearExtraRequestForCaseId(ByVal strCaseId As String)
    Dim sql As String

    'Remove space and /
10  On Error GoTo ClearExtraRequestForCaseId_Error

20  strCaseId = Replace(strCaseId, " ", "")
30  strCaseId = Replace(strCaseId, "/", "")

40  sql = "UPDATE CaseTree " & _
          "SET ExtraRequests = '0' " & _
          "WHERE CaseId = N'" & strCaseId & "'"
50  Cnxn(0).Execute sql

60  Exit Sub

ClearExtraRequestForCaseId_Error:

    Dim strES As String
    Dim intEL As Integer

70  intEL = Erl
80  strES = Err.Description
90  LogError "frmWorklist", "ClearExtraRequestForCaseId", intEL, strES, sql

End Sub

Private Sub lstWorkList_Click(Index As Integer)
    Dim CaseIds() As String

10  On Error GoTo lstWorkList_Click_Error

20  If Index = 4 Then    'for Extra Request list only
30      If UCase$(UserMemberOf) = "MANAGER" Then    'only for Manager users
40          If lstWorkList(4).ListItems.Count > 0 Then    'Only if there are case Ids present in list
50              imgRemoveExtraRequests.Visible = True
60          End If
70      End If
80  Else
90      imgRemoveExtraRequests.Visible = False
100 End If

110 If (UCase$(UserMemberOf) = "SCIENTISTS") Or (UCase$(UserMemberOf) = "SCIENTIST") Or (UCase$(UserMemberOf) = "MANAGER") Then
120     If Index = 0 Then
130         cmdChangePath.Visible = True
            cmdChangePath.Left = 1920
140     ElseIf Index = 1 Then
150         cmdChangePath.Visible = True
160         cmdChangePath.Left = 4020
        Else
            cmdChangePath.Visible = False
170     End If




180 End If



lstWorkList_Click_Error:
    Dim strES As String
    Dim intEL As Long

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmWorklist", "lstWorkList_Click", intEL, strES


End Sub

Private Sub lstWorkList_DblClick(Index As Integer)
10  If UCase$(UserMemberOf) <> "LOOKUP" And _
       UCase$(UserMemberOf) <> "NCRI" Then
20      If lstWorkList(Index).SelectedItem Is Nothing Then
30          Exit Sub
40      Else
50          With frmWorkSheet
60              .txtCaseId = lstWorkList(Index).SelectedItem
70              .Show
80              .cmbPatientId.SetFocus
90          End With
100         frmWorklist.Enabled = False
110     End If
120 End If
End Sub


Private Sub lstWorkList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
10  If lstWorkList(Index).SelectedItem Is Nothing Then
20      For i = 0 To 5
30          Set lstWorkList(i).SelectedItem = Nothing
40      Next
50      sCaseId = ""
60      Exit Sub
70  Else
80      sCaseId = lstWorkList(Index).SelectedItem
90  End If
End Sub

Private Sub mnuAbout_Clifrmck()
10  frmAbout.Show 1
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuAccreditationSettings_Click()
10  With frmAccreditation
20      .Show 1
30  End With
End Sub

Private Sub mnuAudit_Click()
10  With frmCaseEventLog
20      .Show 1
30  End With
End Sub

Private Sub mnuAutDisposal_Click()
10  With frmHistDisposal
20      .DisposalType = "A"
30      .Show 1
40  End With
End Sub

Private Sub mnuAuthorisedReports_Click()

10  On Error GoTo mnuAuthorisedReports_Click_Error

20  With frmAuthorisedReports
30      .Show 1
40  End With

50  Exit Sub

mnuAuthorisedReports_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmWorklist", "mnuAuthorisedReports_Click", intEL, strES

End Sub

Private Sub mnuChangeCodes_Click()
10  frmChangePMQ.Show 1
End Sub

Private Sub mnuClinicians_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "Clinician"
30      .ListTypeName = "Clinician"
40      .ListTypeNames = "Clinicians"
50      .ListGenericName = "Clinicians"
60      .FrameAdd.Caption = " Add Clinicians"
70      .Show 1
80  End With
End Sub

Private Sub mnuCoroners_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "Coroner"
30      .ListTypeName = "Coroner"
40      .ListTypeNames = "Coroners"
50      .ListGenericName = "Coroners"
60      .FrameAdd.Caption = "Add Coroners"
70      .Show 1
80  End With
End Sub

Private Sub mnuCounty_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "County"
30      .ListTypeName = "County"
40      .ListTypeNames = "Counties"
50      .ListGenericName = "Counties"
60      .FrameAdd.Caption = "Add Counties"
70      .Show 1
80  End With
End Sub

Private Sub mnuCutting_Click()
10  With frmCutUpEmbed
20      .SingleEdit = False
30      .Phase = "Cutting"
40      .Show 1
50  End With
End Sub

Private Sub mnuCutUpSheet_Click()
10  With frmCutUpEmbed
20      .SingleEdit = False
30      .Phase = "Cut-Up"
40      .Show 1
50  End With
End Sub

Private Sub mnuCytDisposal_Click()
10  With frmHistDisposal
20      .DisposalType = "C"
30      .Show 1
40  End With
End Sub

Private Sub mnuDiscrepancyResolution_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "DiscrepRes"
30      .ListTypeName = "Discrepancy Resolution"
40      .ListTypeNames = "Discrepancy Resolution"
50      .ListGenericName = "Discrepancy Resolutions"
60      .FrameAdd.Caption = " Add Discrepancy Resolutions"
70      .Show 1
80  End With
End Sub

Private Sub mnuDiscrepancyType_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "DiscrepType"
30      .ListTypeName = "Discrepancy Type"
40      .ListTypeNames = "Discrepancy Types"
50      .ListGenericName = "Discrepancy Types"
60      .FrameAdd.Caption = "Add Discrepancy Types"
70      .Show 1
80  End With
End Sub

Private Sub mnuEmbedSheet_Click()
10  With frmCutUpEmbed
20      .SingleEdit = False
30      .Phase = "Embedding"
40      .Show 1
50  End With
End Sub

Private Sub mnuExit_Click()
10  Unload Me
End Sub

Private Sub mnuDestCodes_Click()
10  With frmListDestinations
20      .ListType = "Code"
30      .ListTypeName = "Codes"
40      .Show 1
50  End With
End Sub


Private Sub mnuDestStains_Click()
10  With frmListDestinations
20      .ListType = "Stain"
30      .ListTypeName = "Stains"
40      .Show 1
50  End With
End Sub

Private Sub mnuGPs_Click()
10  frmGps.Show 1
End Sub

Private Sub mnuHistDisposal_Click()
10  With frmHistDisposal
20      .DisposalType = GetOptionSetting("HistologyLeadingCaseIdCharacter", "H")
30      .Show 1
40  End With
End Sub

Private Sub mnuLanguageEnglish_Click()
10  sysOptCurrentLanguage = "English"
20  LoadLanguage sysOptCurrentLanguage
End Sub

Private Sub mnuLanguagePortuguese_Click()
10  sysOptCurrentLanguage = "Portuguese"
20  LoadLanguage sysOptCurrentLanguage
End Sub

Private Sub mnuLanguageRussian_Click()
10  sysOptCurrentLanguage = "Russian"
20  LoadLanguage sysOptCurrentLanguage
End Sub

Private Sub mnuLocked4Editing_Click()

10  frmCasesCurrentlyOpened.Show 1

End Sub

Private Sub mnuNcri_Click()
10  With frmNCRI
20      .Show 1
30  End With
End Sub

Private Sub mnuNonWorkingDays_Click()
10  frmListsDates.Show 1
End Sub

Private Sub mnuNumerical_Click()
10  With frmNumericalStats
20      .Show 1
30  End With
End Sub

Private Sub mnuOrientation_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "Orientation"
30      .ListTypeName = "Orientation"
40      .ListTypeNames = "Orientation"
50      .ListGenericName = "Orientation"
60      .FrameAdd.Caption = "Add Orientation"
70      .Show 1
80  End With
End Sub

Private Sub mnuPCodes_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "P"
30      .ListTypeName = "P Code"
40      .ListTypeNames = "P Codes"
50      .ListGenericName = "P Codes"
60      .FrameAdd.Caption = "Add P Codes"
70      .Show 1
80  End With

End Sub

Private Sub mnuMCodes_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "M"
30      .ListTypeName = "M Code"
40      .ListTypeNames = "M Codes"
50      .ListGenericName = "M Codes"
60      .FrameAdd.Caption = "Add M Codes"
70      .Show 1
80  End With

End Sub

Private Sub mnuProcessor_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "Processor"
30      .ListTypeName = "Processor"
40      .ListTypeNames = "Processor"
50      .ListGenericName = "Processor"
60      .FrameAdd.Caption = "Add Processor"
70      .Show 1
80  End With
End Sub

Private Sub mnuQCodes_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "Q"
30      .ListTypeName = "Q Code"
40      .ListTypeNames = "Q Codes"
50      .ListGenericName = "Q Codes"
60      .FrameAdd.Caption = "Add Q Codes"
70      .Show 1
80  End With

End Sub



Private Sub mnuReasonReferral_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "RefReason"
30      .ListTypeName = "Reason For Referral"
40      .ListTypeNames = "Reasons For Referral"
50      .ListGenericName = "Reasons For Referral"
60      .FrameAdd.Caption = "Add Reason For Referral"
70      .Show 1
80  End With
End Sub

Private Sub mnuReferral_Click()
10  With frmReferralLog
20      .Show 1
30  End With
End Sub

Private Sub mnuReferredTo_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "RefTo"
30      .ListTypeName = "Referred To"
40      .ListTypeNames = "Referred To"
50      .ListGenericName = "Referred To"
60      .FrameAdd.Caption = "Add Refered To"
70      .Show 1
80  End With
End Sub

Private Sub mnuRoutineStainList_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "RS"
30      .ListTypeName = "Routine Stain"
40      .ListTypeNames = "Routine Stains"
50      .ListGenericName = "Routine Stains"
60      .FrameAdd.Caption = "Add Routine Stains"
70      .Show 1
80  End With
End Sub

Private Sub mnuRptDiscrep_Click()
10  With frmDiscrepancyReport
20      .Show 1
30  End With
End Sub

Private Sub mnuSearch1_Click()
10  With frmSnomedSearch
20      .lblDescriptionSearch = "List of specimens that have T Code and M Code"
30      .SearchType = "1"
40      .Show
50  End With
60  frmWorklist.Enabled = False
End Sub

Private Sub mnuSearch2_Click()
10  With frmSnomedSearch
20      .lblDescriptionSearch = "Number of Cases that have T Code and M Code"
30      .SearchType = "2"
40      .Show
50  End With
60  frmWorklist.Enabled = False
End Sub

Private Sub mnuSearch3_Click()
10  With frmSnomedSearch
20      .lblDescriptionSearch = "Number of Cases"
30      .SearchType = "3"
40      .Show
50  End With
60  frmWorklist.Enabled = False
End Sub

Private Sub mnuSearch4_Click()

10  With frmSnomedSearch
20      .lblDescriptionSearch = "List of specimens that are within T Code Group ? and have an M Code ?"
30      .SearchType = "4"
40      .Show
50  End With
60  frmWorklist.Enabled = False

End Sub

Private Sub mnuSearch5_Click()
10  With frmSnomedSearch
20      .lblDescriptionSearch = "No. of Cases filterd by Hospital and clinician with an M Code ?"
30      .SearchType = "5"
40      .Show
50  End With
60  frmWorklist.Enabled = False
End Sub

Private Sub mnuSource_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "Source"
30      .ListTypeName = "Source"
40      .ListTypeNames = "Source"
50      .ListGenericName = "Source"
60      .FrameAdd.Caption = "Add Source"
70      .Show 1
80  End With
End Sub


Private Sub mnuTATPCodeCaseIds_Click()
10  With frmTATcases
20      .Code = "P"
30      .Show 1
40  End With
End Sub

Private Sub mnuTatPCodes_Click()
10  With frmTAT
20      .Code = "P"
30      .Show 1
40  End With
End Sub

Private Sub mnuTATTCodeCaseIds_Click()
10  With frmTATcases
20      .Code = "T"
30      .Show 1
40  End With
End Sub

Private Sub mnuTatTCodes_Click()
10  With frmTAT
20      .Code = "T"
30      .Show 1
40  End With
End Sub

Private Sub mnuTCodes_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "T"
30      .ListTypeName = "T Code"
40      .ListTypeNames = "T Codes"
50      .ListGenericName = "T Codes"
60      .FrameAdd.Caption = "Add T Codes"
70      .Show 1
80  End With

End Sub

Private Sub mnuSpecialStainList_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "SS"
30      .ListTypeName = "Special Stain"
40      .ListTypeNames = "Special Stains"
50      .ListGenericName = "Special Stains"
60      .FrameAdd.Caption = "Add Special Stains"
70      .Show 1
80  End With

End Sub

Private Sub mnuImmunoStainList_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "IS"
30      .ListTypeName = "Immunohistochemical Stain"
40      .ListTypeNames = "Immunohistochemical Stains"
50      .ListGenericName = "Immunohistochemical Stains"
60      .FrameAdd.Caption = "Add Immunohistochemical Stains"
70      .Show 1
80  End With

End Sub

Private Sub mnuLogOff_Click()

10  Unload Me
20  With frmSystemManager
30      .Show
40  End With

50  UserName = ""
60  userCode = ""

70  UpdateLoggedOnUsers
End Sub



Private Sub mnuWards_Click()
'Line number 60 added by Ibrahim 29-7-24
10  With frmListsGeneric
20      .ListType = "Ward"
30      .ListTypeName = "Ward"
40      .ListTypeNames = "Wards"
50      .ListGenericName = "Wards"
60      .FrameAdd.Caption = "Add Wards"
70      .Show 1
80  End With
End Sub


Private Sub Timer1_Timer()

    Static X As Long
    Static Y As Long
    Dim TempX As Long
    Dim TempY As Long
    Dim h As Long

10  On Error GoTo Timer1_Timer_Error

20  If TimedOut Then
30      Debug.Print Forms.Count
40      If Forms.Count > 1 Then
50          Do While Forms.Count > 1
60              Unload Forms(Forms.Count - 1)
70          Loop
80      Else
90          TimedOut = False
100     End If
110 End If

120 If TopMostWindow() = Screen.ActiveForm.Caption Then

130     h = Screen.ActiveForm.hwnd

140     TempX = MouseX(h)
150     TempY = MouseY(h)
160     If X <> TempX Or Y <> TempY Then
170         If TempX > 0 And _
               TempY > -30 And _
               TempX * Screen.TwipsPerPixelX < Screen.ActiveForm.Width And _
               TempY * Screen.TwipsPerPixelY < Screen.ActiveForm.Height - 320 Then
180             X = TempX
190             Y = TempY
200             Screen.ActiveForm.Controls("pBar").Value = 0
210         End If
220     End If

230     If KB() Then
240         Screen.ActiveForm.Controls("pBar").Value = 0
250     End If

260 End If

270 With Screen.ActiveForm.Controls("pBar")
280     If LogOffDelaySecs <> 0 Then
290         .Max = LogOffDelaySecs
300     Else
310         .Max = 30
320     End If
330     .Value = .Value + 1
340     If .Value = .Max Then
350         .Value = 0
360         TimedOut = True



            Dim iLoop As Integer
            Dim iHighestForm As Integer
370         iHighestForm = Forms.Count - 1

380         For iLoop = iHighestForm To 0 Step -1
390             Unload Forms(iLoop)
400         Next iLoop



410         With frmSystemManager
420             .Show
430         End With

440         UserName = ""

450         UpdateLoggedOnUsers
460     End If
470 End With



480 Exit Sub

Timer1_Timer_Error:

    Dim strES As String
    Dim intEL As Integer

490 intEL = Erl
500 strES = Err.Description
510 LogError "frmWorkList", "Timer1_Timer", intEL, strES


End Sub


Public Sub FillList(Optional Code As String)
    Dim tb As New Recordset
    Dim sql As String
    Dim i As Integer

10  On Error GoTo FillList_Error

20  For i = 0 To 5
30      lstWorkList(i).ListItems.Clear
40  Next



50  sql = "SELECT DISTINCT C.CaseID, C.State, D.Urgent " & _
          "FROM Cases C " & _
          "INNER JOIN Demographics D ON C.CaseId = D.CaseId "


60  If Code <> "" Then
70      sql = sql & "AND ((C.State = N'" & "With Pathologist" & "' "
80      sql = sql & "AND C.WithPathologist = '" & Code & "') "
90      sql = sql & "OR (C.State = N'" & "Awaiting Authorisation" & "' "
100     sql = sql & "AND C.AAPathologist = '" & Code & "')) "
110 End If

120 sql = sql & "ORDER BY C.CaseID, C.State"


130 Set tb = New Recordset
140 RecOpenClient 0, tb, sql

150 If Not tb.EOF Then
160     pbProgress.Max = tb.RecordCount + 1
170     fraProgress.Visible = True

180     Do While Not tb.EOF
            '        If Right(tb!CaseId, 2) = "14" Then
            '            A = b
            '        End If


190         For i = 0 To 2
200             If Trim(tb!State) = Trim(lblWorkList(i)) Then
210                 pbProgress.Value = pbProgress.Value + 1
220                 lblProgress = "Refreshing " & lblWorkList(i) & " ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
230                 lblProgress.Refresh

240                 If tb!Urgent = 1 Then
250                     If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
260                         lstWorkList(i).ListItems.Add 1, , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
270                     Else
280                         lstWorkList(i).ListItems.Add 1, , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
290                     End If
300                     lstWorkList(i).ListItems(1).ForeColor = vbRed
310                 ElseIf i = 2 And CaseLinked2AuthCytoCase(tb!CaseId, i) Then
320                     lstWorkList(i).ListItems.Add 1, , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
330                     lstWorkList(i).ListItems(1).ForeColor = vbGreen
340                 Else
350                     If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
360                         lstWorkList(i).ListItems.Add , , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
370                     Else
380                         lstWorkList(i).ListItems.Add , , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
390                     End If

400                 End If
410                 Exit For
420             End If
430         Next
440         tb.MoveNext
450     Loop

460     fraProgress.Visible = False
470     pbProgress.Value = 1

480 End If



490 Exit Sub

FillList_Error:

    Dim strES As String
    Dim intEL As Integer

500 intEL = Erl
510 strES = Err.Description
520 LogError "frmWorkList", "FillList", intEL, strES, sql

End Sub


Public Sub FillExternal(Optional Code As String)
    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo FillExternal_Error

20  sql = "Select DISTINCT CM.CaseId, D.Urgent From CaseMovements CM " & _
          "INNER JOIN Cases C ON CM.CaseId = C.CaseId " & _
          "INNER JOIN Lists L ON CM.ReferralReason = L.Description " & _
          "INNER JOIN Demographics D ON CM.CaseId = D.CaseId " & _
          "WHERE (CM.DateReceived IS NULL OR CM.DateReceived = '') " & _
          "AND L.ShowWorkList = 1"
    '
30  If Code <> "" Then
40      sql = sql & "AND C.State = N'" & "With Pathologist" & "' "
50      sql = sql & "AND C.WithPathologist = '" & Code & "' "
60  End If


70  Set tb = New Recordset
80  RecOpenClient 0, tb, sql

90  If Not tb.EOF Then
100     pbProgress.Max = tb.RecordCount + 1
110     fraProgress.Visible = True

120     Do While Not tb.EOF
130         pbProgress.Value = pbProgress.Value + 1
140         lblProgress = "Refreshing " & lblWorkList(5) & " ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
150         lblProgress.Refresh

160         If tb!Urgent = 1 Then
170             If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
180                 lstWorkList(5).ListItems.Add 1, , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
190             Else
200                 lstWorkList(5).ListItems.Add 1, , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
210             End If
220             lstWorkList(5).ListItems(1).ForeColor = vbRed
230         Else
240             If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
250                 lstWorkList(5).ListItems.Add , , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
260             Else
270                 lstWorkList(5).ListItems.Add , , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
280             End If
290         End If

300         tb.MoveNext
310     Loop
320     fraProgress.Visible = False
330     pbProgress.Value = 1

340 End If


350 Exit Sub

FillExternal_Error:

    Dim strES As String
    Dim intEL As Integer

360 intEL = Erl
370 strES = Err.Description
380 LogError "frmWorkList", "FillExternal", intEL, strES, sql


End Sub

Public Sub FillExtraRequests(Optional Code As String)
    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo FillExtraRequests_Error

20  sql = "SELECT DISTINCT CaseId FROM CaseTree " & _
          "WHERE ExtraRequests <> '0' " & _
          "AND ExtraRequests IS NOT NULL " & _
          "AND ExtraRequests <> '' "


30  Set tb = New Recordset
40  RecOpenClient 0, tb, sql

50  If Not tb.EOF Then

60      pbProgress.Max = tb.RecordCount + 1
70      fraProgress.Visible = True

80      Do While Not tb.EOF
90          pbProgress.Value = pbProgress.Value + 1
100         lblProgress = "Refreshing " & lblWorkList(4) & " ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
110         lblProgress.Refresh



120         If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
130             lstWorkList(4).ListItems.Add , , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
140         Else
150             lstWorkList(4).ListItems.Add , , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
160         End If
170         lstWorkList(4).ListItems(lstWorkList(4).ListItems.Count).ForeColor = vbBlue    '&H66FF

180         tb.MoveNext
190     Loop

200     fraProgress.Visible = False
210     pbProgress.Value = 1
220 End If


230 Exit Sub

FillExtraRequests_Error:

    Dim strES As String
    Dim intEL As Integer

240 intEL = Erl
250 strES = Err.Description
260 LogError "frmWorkList", "FillExtraRequests", intEL, strES, sql


End Sub

Public Sub FillValNotPrint(Optional Code As String)
    Dim tb As New Recordset
    Dim sql As String



10  On Error GoTo FillValNotPrint_Error


20  sql = "SELECT DISTINCT C.CaseId, D.Urgent FROM Cases C " & _
          "INNER JOIN Demographics D ON C.CaseId = D.CaseId " & _
          "WHERE C.Validated = 1 AND C.PrintedVal = 0 "

30  If Code <> "" Then
40      sql = sql & "AND C.ValidatedBy = N'" & AddTicks(Code) & "' "
50  End If

60  sql = sql & "ORDER BY C.CaseID"

70  Set tb = New Recordset
80  RecOpenClient 0, tb, sql

90  If Not tb.EOF Then

100     pbProgress.Max = tb.RecordCount + 1
110     fraProgress.Visible = True

120     Do While Not tb.EOF
130         pbProgress.Value = pbProgress.Value + 1
140         lblProgress = "Refreshing " & lblWorkList(3) & " ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
150         lblProgress.Refresh

160         If tb!Urgent = 1 Then
170             If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
180                 lstWorkList(3).ListItems.Add 1, , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
190             Else
200                 lstWorkList(3).ListItems.Add 1, , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
210             End If
220             lstWorkList(3).ListItems(1).ForeColor = vbRed
230         Else
240             If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
250                 lstWorkList(3).ListItems.Add , , Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
260             Else
270                 lstWorkList(3).ListItems.Add , , Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
280             End If

290         End If

300         tb.MoveNext
310     Loop
320     fraProgress.Visible = False
330     pbProgress.Value = 1
340 End If

350 Exit Sub

FillValNotPrint_Error:

    Dim strES As String
    Dim intEL As Integer

360 intEL = Erl
370 strES = Err.Description
380 LogError "frmWorkList", "FillValNotPrint", intEL, strES, sql


End Sub

Private Sub tmrRefresh_Timer()
    Dim i As Integer

10  If Not blnWorklistLoadedAlready Or UCase$(UserMemberOf) = "CONSULTANT" Then
20      imgRemoveExtraRequests.Visible = False
30      blnWorklistLoadedAlready = True

40      FillList PathologistCode
50      FillExternal PathologistCode
60      FillExtraRequests PathologistCode
70      FillValNotPrint PathologistName

80      For i = 0 To 5
90          Set lstWorkList(i).SelectedItem = Nothing
100     Next

110     fraProgress.Visible = False
120     pbProgress.Value = 1
130 End If
    blnWorklistLoadedAlready = True
140 tmrRefresh.Enabled = False


End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13350
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCaseId 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MaxLength       =   12
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   1100
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Exit"
      CausesValidation=   0   'False
      Height          =   1100
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1200
   End
   Begin MSComctlLib.TreeView tvCaseDetails 
      Height          =   5355
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   9446
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "lstImages"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList lstImages 
      Left            =   9720
      Top             =   1680
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
            Picture         =   "frmTest.frx":0ECA
            Key             =   "two"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1251
            Key             =   "one"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   2880
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblLocationPath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   690
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   10350
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Location Path "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label lblParentLocationID 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8160
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Menu mnuPopUpLevel1 
      Caption         =   "PopUpLevel1"
      Visible         =   0   'False
      Begin VB.Menu mnuAddTissueType 
         Caption         =   "Add Tissue Type"
      End
   End
   Begin VB.Menu mnuPopUpLevel2 
      Caption         =   "PopUpLevel2"
      Visible         =   0   'False
      Begin VB.Menu mnuSingleBlock 
         Caption         =   "Add Single Block"
      End
   End
   Begin VB.Menu mnuPopUpLevel3 
      Caption         =   "PopUpLevel3"
      Visible         =   0   'False
      Begin VB.Menu mnuSingleSlide 
         Caption         =   "Add Single Slide"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedNode As MSComctlLib.Node
Public MyOpt As String

Private Function FillTree() As Boolean

      Dim tb As New Recordset
      Dim sql As String
      Dim nod As MSComctlLib.Node

10    On Error GoTo FillTree_Error

20    sql = "Select * From CaseTree Where CaseId = '" & CaseNo & "' Order By LocationParentID"
30    RecOpenClient 0, tb, sql

40    If Not tb.EOF Then
50        With tvCaseDetails
60        .Nodes.Clear
70        While Not tb.EOF
80            If tb!LocationParentID = 0 Then
90                 Set nod = .Nodes.Add(, , "L" & tb!LocationLevel & tb!LocationID, tb!LocationName, 1, 2)
100                nod.Bold = True
 
110           Else
120               .SingleSel = True
130               Set nod = .Nodes.Add("L" & Val(tb!LocationLevel) - 1 & tb!LocationParentID, tvwChild, "L" & tb!LocationLevel & tb!LocationID, tb!LocationName, 1, 2)
140           End If
150           nod.ForeColor = &H80000003
  
160           tb.MoveNext
170       Wend
180       .SingleSel = False
190       End With
200       FillTree = True
210   Else
220       FillTree = False
230   End If

      '210   tvCaseDetails.Nodes(1).Expanded = True
      '220   tvCaseDetails.Nodes(1).Selected = True
      '230   Set SelectedNode = tvCaseDetails.Nodes(1)
      '240   lblParentLocationID.Caption = Right(SelectedNode.Key, Len(SelectedNode.Key) - 1)
      '250   lblLocationPath.Caption = SelectedNode.FullPath
240   Exit Function

FillTree_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmSetLocations", "FillTree", intEL, strES, sql

End Function

Private Sub cmdCancel_Click()



10    On Error GoTo cmdCancel_Click_Error


20    Unload Me

30    Exit Sub

cmdCancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmSetLocations", "cmdCancel_Click", intEL, strES

End Sub


Private Sub cmdSave_Click()
      Dim tb As New Recordset
      Dim sql As String
      Dim n As Integer
      Dim i As Integer

10    On Error GoTo cmdSave_Click_Error


      'If SelectedNode Is Nothing Then
      '    iMsg "PleaseSelectParentLocationFirst", vbInformation
      '    If TimedOut Then Unload Me: Exit Sub
      '    Exit Sub
      'End If

20    For i = 1 To tvCaseDetails.Nodes.Count

30        sql = "Select Max(LocationID) MaxID From CaseTree"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql
    
60        If (tb.EOF And tb.BOF) Or IsNull(tb!MaxID) Then
70            n = 10000
80        Else
90            n = tb!MaxID + 1
100       End If
    
110       Set tb = New Recordset
120       sql = "Select * From CaseTree Where 1=0"
130       RecOpenServer 0, tb, sql
    
140       tb.AddNew
150       tb!CaseId = CaseNo
160       tb!LocationID = Right(tvCaseDetails.Nodes(i).Key, Len(tvCaseDetails.Nodes(i).Key) - 2)
170       tb!LocationName = tvCaseDetails.Nodes(i).Text
180       If Left(tvCaseDetails.Nodes(i).Key, 2) = "L0" Then
190           tb!LocationParentID = 0
200       Else
              'tb!LocationParentID = Val(lblParentLocationID)
210           tb!LocationParentID = Right(tvCaseDetails.Nodes(i).Parent.Key, Len(tvCaseDetails.Nodes(i).Parent.Key) - 2)
220       End If
230       If Left(tvCaseDetails.Nodes(i).Key, 2) = "L0" Then
240           tb!LocationLevel = 0
250       Else
260           tb!LocationLevel = GetNodeLevel(Right(tvCaseDetails.Nodes(i).Parent.Key, Len(tvCaseDetails.Nodes(i).Parent.Key) - 2)) + 1
270       End If
280       tb!LocationPath = SelectedNode.FullPath & "\" & tvCaseDetails.Nodes(i).Text
290       Debug.Print tb!LocationPath
300       tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")
310       tb!UserName = UserName
    
320       tb.Update

330   Next i


      Dim nod As MSComctlLib.Node
340   Set nod = SelectedNode
350   FillTree
360   lblParentLocationID = ""

370   Set SelectedNode = FindNodeByKey(nod.Key)
380   SelectedNode.Selected = True
390   SelectedNode.Expanded = True
400   lblParentLocationID = Right(SelectedNode.Key, Len(SelectedNode.Key) - 1)
410   lblLocationPath = SelectedNode.FullPath


420   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "frmSetLocations", "cmdSave_Click", intEL, strES, sql


End Sub









Private Function FindNodeByKey(Key As String) As MSComctlLib.Node

      Dim n As Integer
10    On Error GoTo FindNodeByKey_Error

20    For n = 1 To tvCaseDetails.Nodes.Count
30        If UCase(tvCaseDetails.Nodes(n).Key) = UCase(Key) Then
              'tvCaseDetails.Nodes(n).Selected = True
              'tvCaseDetails.Nodes(n).Parent.Expanded = True
40            Set FindNodeByKey = tvCaseDetails.Nodes(n)
              'exit for
50        End If
60    Next n

70    Exit Function

FindNodeByKey_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmLocations", "FindNodeByKey", intEL, strES

End Function





Private Sub Form_Load()



10    On Error GoTo Form_Load_Error



      'CreateFirstLocation
20    FillTree

30    Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmSetLocations", "Form_Load", intEL, strES


End Sub


Private Sub CallTreePopupMenu(tvtemp As MSComctlLib.TreeView, X As Single, Y As Single)

10    If tvtemp.SelectedItem Is Nothing Then
20        Exit Sub
30    Else
40        Set tvtemp.SelectedItem = tvtemp.HitTest(X, Y)
50        Select Case Left(tvtemp.SelectedItem.Key, 2)
          Case "L0"
60            PopupMenu mnuPopupLevel2
70        Case "L2"
80            PopupMenu mnuPopupLevel2
90        Case "L3"
100               PopupMenu mnuPopupLevel3
110       End Select
120   End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
10    On Error GoTo Form_Unload_Error

20    Me.MyOpt = ""

30    Exit Sub

Form_Unload_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmSetLocations", "Form_Unload", intEL, strES

End Sub



Private Sub mnuAddTissueType_Click()

10    With tvCaseDetails.SelectedItem
20        With frmAddCodeTree
30            Set .tvtemp = tvCaseDetails
40            .Update = False
50            .ListType = "T"
60            .Move frmWorkSheet.Left + tvCaseDetails.Left, frmWorkSheet.Top + tvCaseDetails.Top - 750
70            .Show
80        End With
90    End With
End Sub

Private Sub mnuSingleBlock_Click()
      Dim iBlock As Integer
      Dim tnode As MSComctlLib.Node
      Dim NoOfBlocks As Integer
      Dim UniqueId As Integer

10    If InStr(1, tvCaseDetails.SelectedItem.Text, "Frozen Section") Then
20        NoOfBlocks = GetBlockNumber(tvCaseDetails.SelectedItem.Parent)
30    Else
40        NoOfBlocks = GetBlockNumber(tvCaseDetails.SelectedItem)
50    End If
60    With tvCaseDetails.SelectedItem
70        UniqueId = GetUniqueID
80        'SaveUniqueID (UniqueId)
90        If sysOptBlockNumberingFormat(0) = "1" Then
  
              'saveuniqueid(
100           iBlock = NoOfBlocks + 1
110           Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L1" & UniqueId, "Block " & iBlock, 1, 2)
120           If UCase(Left(Trim(tvCaseDetails.SelectedItem.Text), 14)) <> "FROZEN SECTION" Then
130               AddDefaultStains "L1" & UniqueId, tvCaseDetails.SelectedItem
140           End If
150       Else
160           If NoOfBlocks = 1 Then
170               .Child.Text = "Block A"
180           End If
190           iBlock = NoOfBlocks + 65
200           Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L1" & UniqueId, "Block " & IIf(Chr(iBlock) = "A", "", Chr(iBlock)), 1, 2)
210           If UCase(Left(Trim(tvCaseDetails.SelectedItem.Text), 14)) <> "FROZEN SECTION" Then
220               AddDefaultStains "L1" & UniqueId, tvCaseDetails.SelectedItem
230           End If
240       End If
250       tnode.Expanded = True
260       tnode.Selected = True
270   End With
280   DataChanged = True
290   TreeChanged = True
End Sub

Private Sub mnuSingleSlide_Click()
      Dim iSlide As Integer
      Dim tnode As MSComctlLib.Node
      Dim NoOfSlides As Integer
      Dim UniqueId As Integer

10    NoOfSlides = GetSlideNumber(tvCaseDetails.SelectedItem)
20    With tvCaseDetails.SelectedItem
30        UniqueId = GetUniqueID
40        'SaveUniqueID (UniqueId)
50        If sysOptSlideNumberingFormat(0) = "1" Then
60            iSlide = NoOfSlides + 1
70            Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L2" & UniqueId, "Slide " & iSlide, 1, 2)
80        Else
90            If NoOfSlides = 1 Then
100               .Child.Text = "Slide A"
110           End If
120           iSlide = NoOfSlides + 65
130           Set tnode = tvCaseDetails.Nodes.Add(.Key, tvwChild, "L2" & UniqueId, "Slide " & IIf(Chr(iSlide) = "A", "", Chr(iSlide)), 1, 2)
140       End If
150       .Expanded = True
160       tnode.Selected = True
170   End With
180   DataChanged = True
190   TreeChanged = True
End Sub

Private Sub tvCaseDetails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    If Button = vbRightButton Then
20        CallTreePopupMenu tvCaseDetails, X, Y
30    End If
End Sub

Private Sub tvCaseDetails_NodeClick(ByVal Node As MSComctlLib.Node)
10    On Error GoTo tvCaseDetails_NodeClick_Error

20    lblParentLocationID.Caption = Right$(Node.Key, Len(Node.Key) - 2)
30    Set SelectedNode = Node
40    lblLocationPath = Node.FullPath

50    Exit Sub

tvCaseDetails_NodeClick_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmSetLocations", "tvCaseDetails_NodeClick", intEL, strES

End Sub

Private Function GetNodeLevel(LocationID As Integer) As Integer

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo GetNodeLevel_Error

20    sql = "Select LocationLevel From CaseTree Where LocationID = " & LocationID
30    RecOpenClient 0, tb, sql

40    If tb.EOF Then
50        GetNodeLevel = -1
60    Else
70        GetNodeLevel = tb!LocationLevel
80    End If

90    Exit Function

GetNodeLevel_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmSetLocations", "GetNodeLevel", intEL, strES, sql

End Function

Private Sub txtCaseId_KeyPress(KeyAscii As Integer)

10    On Error GoTo txtCaseId_KeyPress_Error


20    If DataChanged = False Then
30        If UCase(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
40            Call ValidateTullCaseId(KeyAscii, Me)
50        Else
60            Call ValidateLimCaseId(KeyAscii, Me)
70        End If


80    Else
90        If frmMsgBox.Msg("Alert!!  You have not saved your changes.  Do you want to Exit without saving?", mbYesNo, , mbCritical) = 1 Then
100           If UCase(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
110               Call ValidateTullCaseId(KeyAscii, Me)
120           Else
130               Call ValidateLimCaseId(KeyAscii, Me)
140           End If
  
150               tvCaseDetails.Nodes.Clear

160       Else
170           KeyAscii = 0
180       End If
190   End If


200   Exit Sub

txtCaseId_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmWorkSheet", "txtCaseId_KeyPress", intEL, strES

End Sub

Private Sub txtCaseId_LostFocus()
      Dim sql As String
      Dim UniqueId As Integer

10    On Error GoTo txtCaseId_LostFocus_Error
  
20        If IsValidCaseNo(txtCaseId) Then
30            CaseNo = Replace(txtCaseId, " " & sysOptCaseIdSeperator(0) & " ", "")
40                UniqueId = GetUniqueID
50                'SaveUniqueID (UniqueId)
60            If FillTree Then
  
70            Else
80                If txtCaseId = "" Then
90                    tvCaseDetails.Nodes.Clear
100               Else
110                   tvCaseDetails.Nodes.Clear
120                   If Mid(CaseNo, 2, 1) = "P" Or Mid(CaseNo & "", 2, 1) = "A" Then
130                       tvCaseDetails.Nodes.Add , , "L0" & UniqueId, Left(CaseNo, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(CaseNo, 2), 1, 2
140                   Else
150                       tvCaseDetails.Nodes.Add , , "L0" & UniqueId, Left(CaseNo, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(CaseNo, 2), 1, 2
160                   End If
170                   tvCaseDetails.Nodes(1).Selected = True
180               End If
    
190           End If
200       End If
210   Exit Sub

txtCaseId_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmWorkSheet", "txtCaseId_LostFocus", intEL, strES, sql


End Sub

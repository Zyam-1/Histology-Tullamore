VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmWithPathologist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmWithPathologist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCheckedBy 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   660
      Width           =   2895
   End
   Begin VB.ComboBox cmbWithPathologist 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdAddPathologist 
      Height          =   285
      Left            =   4860
      Picture         =   "frmWithPathologist.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   315
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Checked By"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1710
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "With Pathologist"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   1695
   End
End
Attribute VB_Name = "frmWithPathologist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pName As String
Private pCheckedBy

Public Property Let PathologistName(ByVal strValue As String)

10    pName = strValue

End Property
Public Property Let CheckedBy(ByVal strValue As String)

10    pCheckedBy = strValue

End Property

Private Sub cmdAddPathologist_Click()
      Dim sql As String
      Dim tb As Recordset



10    On Error GoTo cmdAddPathologist_Click_Error

20    If cmbWithPathologist <> "" And cmbCheckedBy <> "" Then
30        sql = "SELECT * FROM Users WHERE UserId = " & cmbWithPathologist.ItemData(cmbWithPathologist.ListIndex)
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            frmWorkSheet.lblWithPathologist = tb!Code & ""
80            frmWorkSheet.lblWithPathologistName = tb!UserName & ""
90        End If
    
100       frmWorkSheet.lblCheckedBy = cmbCheckedBy
    
110       Unload Me
120   Else
130       frmMsgBox.Msg "Please Select Pathologist and Checked By", mbOKOnly, "Histology", mbExclamation
140   End If

150   Exit Sub

cmdAddPathologist_Click_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmWithPathologist", "cmdAddPathologist_Click", intEL, strES, sql


End Sub

Private Sub Form_Activate()
'frmWithPathologist_ChangeLanguage
End Sub

Private Sub Form_Load()
Dim i As Integer
FillPathologistList
For i = 0 To cmbWithPathologist.ListCount - 1
    If UCase$(cmbWithPathologist.List(i)) = UCase$(pName) Then
        cmbWithPathologist.ListIndex = i
        Exit For
    End If
Next
FillCheckedByList
For i = 0 To cmbCheckedBy.ListCount - 1
    If UCase$(cmbCheckedBy.List(i)) = UCase$(pCheckedBy) Then
        cmbCheckedBy.ListIndex = i
        Exit For
    End If
Next
End Sub

Private Sub FillPathologistList()
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillList_Error

20    sql = "SELECT * FROM Users WHERE AccessLevel = 'Consultant'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    cmbWithPathologist.AddItem ""
60    Do While Not tb.EOF
70        cmbWithPathologist.AddItem tb!UserName & ""
80        cmbWithPathologist.ItemData(cmbWithPathologist.NewIndex) = tb!UserId & ""
90        tb.MoveNext
100   Loop
110   cmbWithPathologist.ListIndex = -1

120   Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmWithPathologist", "FillList", intEL, strES, sql

End Sub

Private Sub FillCheckedByList()
      Dim sql As String
      Dim tb As Recordset


10    On Error GoTo FillCheckedByList_Error

20    sql = "SELECT * FROM Users WHERE AccessLevel = 'Consultant' " & _
              "OR AccessLevel = 'Manager' " & _
              "OR AccessLevel = 'Scientist' "
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    cmbCheckedBy.AddItem ""
60    Do While Not tb.EOF
70        cmbCheckedBy.AddItem tb!UserName & ""
80        cmbCheckedBy.ItemData(cmbCheckedBy.NewIndex) = tb!UserId & ""
90        tb.MoveNext
100   Loop
110   cmbCheckedBy.ListIndex = -1


120   Exit Sub

FillCheckedByList_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmWithPathologist", "FillCheckedByList", intEL, strES, sql


End Sub



VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmChange 
   Caption         =   "Change ""With Pathologist"" "
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmChange.frx":0000
      Left            =   3360
      List            =   "frmChange.frx":0002
      TabIndex        =   6
      Top             =   4410
      Width           =   4215
   End
   Begin VB.ComboBox cmbCheckedBy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   4
      Top             =   5040
      Width           =   4215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   975
      Left            =   9720
      Picture         =   "frmChange.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Index           =   1
      Left            =   9720
      Picture         =   "frmChange.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5953
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Checked by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   5100
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Change To Pathologist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   4440
      Width           =   2535
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CInfo As New Collection
Public WithPath As Boolean


Private Sub InitializeGrid()

10    With g
20      .Rows = 2: .FixedRows = 1
30      .Cols = 4: .FixedCols = 0
40      .Rows = 1
50      .Font.Size = fgcFontSize
60      .Font.name = fgcFontName
70      .ForeColor = fgcForeColor
80      .BackColor = fgcBackColor
90      .ForeColorFixed = fgcForeColorFixed
100     .BackColorFixed = fgcBackColorFixed
110     .ScrollBars = flexScrollBarBoth

120     .TextMatrix(0, 0) = "Case Id": .ColWidth(0) = 1350: .ColAlignment(0) = flexAlignLeftCenter
130     .TextMatrix(0, 1) = "Patient Name": .ColWidth(1) = 3700: .ColAlignment(1) = flexAlignLeftCenter
140     .TextMatrix(0, 2) = "With Pathologist": .ColWidth(2) = 3000: .ColAlignment(2) = flexAlignLeftCenter
150     .TextMatrix(0, 3) = "Sample Date/Time": .ColWidth(3) = 2000: .ColAlignment(3) = flexAlignLeftCenter
160
170   End With
End Sub

Private Function retrieveCode(name As String) As String
          Dim sql As String
          Dim tb As Recordset
          Dim userCode As String
10        sql = "SELECT code FROM users WHERE username = '" & Trim(name) & "'"
20        Set tb = New Recordset
30        RecOpenServer 0, tb, sql
40        If Not tb.EOF Then
50            userCode = tb!code
60        End If
70        tb.Close
80        Set tb = Nothing
90        retrieveCode = userCode
End Function

Private Function check(caseid As String) As Boolean
          Dim sql As String
          Dim tb As Recordset

10        sql = "SELECT COUNT(*) AS RecordCount " & _
                "FROM Cases c " & _
                "JOIN Demographics d ON c.CaseID = d.CaseID " & _
                "JOIN Users u ON c.WithPathologist = u.Code " & _
                "WHERE d.CaseID = '" & Trim(caseid) & "'"
20        Set tb = New Recordset
30        RecOpenServer 0, tb, sql
40        If Not tb.EOF And tb!RecordCount > 0 Then
50            check = True
60        Else
70            check = False
80        End If
90        tb.Close
100       Set tb = Nothing
          
End Function


Private Sub cmbPath_Click()
If cmbPath.Text <> "" Then
cmdUpdate.Enabled = True
End If
End Sub

Private Sub cmdExit_Click(Index As Integer)
Me.Hide
End Sub

Private Sub cmdUpdate_Click()
          Dim code As String
          Dim i As Integer
          Dim caseid As String
10       On Error GoTo cmdUpdate_Click_Error
          

20        code = retrieveCode(cmbPath.Text)
30        For i = 1 To g.Rows - 1
40        caseid = Replace(g.TextMatrix(i, 0), "/", "")
50                caseid = Replace(caseid, " ", "")
60             sql = "update cases set WithPathologist= '" & Trim(code) & "',CheckedBy = '" & Trim(cmbCheckedBy.Text) & "', State = 'With Pathologist' Where CaseId = '" & Trim(caseid) & "'"
70           Set tb = New Recordset
80    Cnxn(0).Execute sql
90    If check(Trim(caseid)) Then
100   WithPath = True
110   End If
120       Next
130   cmdUpdate.Enabled = False
150   cmbPath.Text = ""
160   cmbCheckedBy.Text = ""
170   FillG
      Exit Sub
cmdUpdate_Click_Error:
             Dim strES As String
             Dim intEL As Long

180          intEL = Erl
190          strES = Err.Description
200          LogError "frmChange", "cmdUpdate_Click", intEL, strES

          
End Sub


Private Sub Form_Load()
InitializeGrid
FillG
fillcmb

If g.TextMatrix(1, 2) = "" Then

Label2.Visible = True
cmbCheckedBy.Visible = True

FillCheckedByList
Else

Label2.Visible = False
cmbCheckedBy.Visible = False

End If

End Sub

Private Sub fillcmb()
    On Error GoTo fillcmb_Error

    sql = "SELECT * FROM Users WHERE AccessLevel = 'Consultant' And InUse='1' ORDER BY UserID"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    Do While Not tb.EOF
        cmbPath.AddItem tb!UserName
        tb.MoveNext
    Loop

    Exit Sub

fillcmb_Error:
    Dim strES As String
    Dim intEL As Long

    intEL = Erl
    strES = Err.Description
    LogError "frmChange", "fillcmb", intEL, strES
End Sub


Private Sub FillG()
          Dim s As String
          Dim caseid As String
          Dim dcaseid As String

10       On Error GoTo FillG_Error
        g.Clear
        InitializeGrid
        
20        If WithPath = False Then
30            For Each Item In CInfo
40            s = Item(0) & vbTab & Item(1) & vbTab & vbTab & Item(2)
50            g.AddItem s
              
60    Next
70    End If
80        If WithPath = True Then
90            For Each Items In CInfo
100               caseid = Items(0)
110              caseid = Replace(Items(0), "/", "")
120              caseid = Replace(caseid, " ", "")
                 
130              sql = "SELECT d.PatientName, u.UserName, c.SampleTaken " & _
                       "FROM Cases c " & _
                       "JOIN Demographics d ON c.CaseID = d.CaseID " & _
                       "JOIN Users u ON c.WithPathologist = u.Code " & _
                       "WHERE d.CaseID = '" & Trim(caseid) & "'"

140              Set tb = New Recordset
150              RecOpenServer 0, tb, sql

160              Do While Not tb.EOF
170                  If Mid(caseid, 2, 1) = "P" Or Mid(caseid, 2, 1) = "A" Then
180                      dcaseid = Left(caseid, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(caseid, 2)
190                  Else
200                      dcaseid = Left(caseid, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(caseid, 2)
210                  End If

220                  s = dcaseid & vbTab & tb!PatientName & vbTab & tb!UserName & vbTab & Format$(CDate(tb!SampleTaken), "YYYY-MM-DD HH:MM")
230                  g.AddItem s

240                  tb.MoveNext
250              Loop

260              tb.Close
270              Set tb = Nothing
280           Next
290       End If
300       Exit Sub

FillG_Error:
             Dim strES As String
             Dim intEL As Long

310          intEL = Erl
320          strES = Err.Description
330          LogError "frmChange", "FillG", intEL, strES

          
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

60    Do While Not tb.EOF
70        cmbCheckedBy.AddItem tb!UserName & ""
90        tb.MoveNext
100   Loop


120   Exit Sub

FillCheckedByList_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmWithPathologist", "FillCheckedByList", intEL, strES, sql


End Sub

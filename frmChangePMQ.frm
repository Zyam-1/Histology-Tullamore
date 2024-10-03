VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmChangePMQ 
   Caption         =   "Change P / M / Q Codes"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   795
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   9870
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   8760
      Picture         =   "frmChangePMQ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   855
      Left            =   8640
      Picture         =   "frmChangePMQ.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   360
      Width           =   915
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8760
      Picture         =   "frmChangePMQ.frx":06C0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   795
   End
   Begin VB.Frame frCom 
      Caption         =   "Event Comment"
      Height          =   1335
      Left            =   480
      TabIndex        =   11
      Top             =   6960
      Width           =   7935
      Begin VB.TextBox txtCom 
         Height          =   855
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   7335
      End
   End
   Begin VB.Frame frChangeTo 
      Caption         =   "Change To"
      Height          =   1455
      Left            =   480
      TabIndex        =   8
      Top             =   5280
      Width           =   7935
      Begin VB.ComboBox cmbCodesDes 
         Height          =   330
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label2 
         Caption         =   "Code - Description"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3255
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
   End
   Begin VB.Frame frCodeRadio 
      Height          =   1575
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton opt 
         Caption         =   "M Codes"
         Height          =   375
         Index           =   1
         Left            =   315
         TabIndex        =   6
         Top             =   610
         Width           =   975
      End
      Begin VB.OptionButton opt 
         Caption         =   "Q Codes"
         Height          =   210
         Index           =   2
         Left            =   315
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton opt 
         Caption         =   "P Codes"
         Height          =   210
         Index           =   0
         Left            =   315
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frCaseID 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3375
      Begin VB.TextBox txtCaseID 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Case ID"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   340
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmChangePMQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colorFlag As Boolean

Private Sub cmdExit_Click()
Me.Hide
End Sub

Private Sub cmdSearch_Click()
          
          Dim sql As String
          Dim tb As Recordset
          Dim CaseId As String
10        CaseId = Replace(Trim(txtCaseID.Text), "/", "")
20        CaseId = Replace(Trim(CaseId), " ", "")


30        g.Rows = 1
40        InitilizeGrid
50        Set tb = New Recordset
          
            
          
60        On Error GoTo cmdSearch_Click_Error
70        If CaseId <> "" Then
80            If opt(0).Value = True Then
90                sql = "SELECT CL.DateTimeOfRecord, L.Code, L.Description FROM CaseListLink CL INNER JOIN Lists L ON L.ListId = CL.ListId WHERE  CL.CaseId" _
                      & " = '" & Trim(CaseId) & "' and cl.Type = 'P'"


100               RecOpenServer 0, tb, sql
110               If Not tb Is Nothing Then
120                   Do While Not tb.EOF
130                       g.AddItem tb!Code & vbTab & tb!Description & vbTab & tb!DateTimeOfRecord
140                       tb.MoveNext
150                   Loop
160               End If
                  FillComboBox (0)

170           ElseIf opt(1).Value = True Then
180               sql = "SELECT CL.TissueTypeLetter, CL.DateTimeOfRecord, L.Code, L.Description FROM CaseListLink CL INNER JOIN Lists L ON L.ListId = CL.ListId WHERE  CL.CaseId" _
                      & " = '" & Trim(CaseId) & "' and cl.Type = 'M'"


190               RecOpenServer 0, tb, sql
200               If Not tb Is Nothing Then
210                   Do While Not tb.EOF
220                       g.AddItem tb!Code & vbTab & tb!Description & vbTab & tb!TissueTypeLetter & vbTab & tb!DateTimeOfRecord
230                       tb.MoveNext
240                   Loop
250               End If
                  FillComboBox (1)
260           ElseIf opt(2).Value = True Then
270               sql = "SELECT CL.DateTimeOfRecord, L.Code, L.Description FROM CaseListLink CL INNER JOIN Lists L ON L.ListId = CL.ListId WHERE  CL.CaseId" _
                      & " = '" & Trim(CaseId) & "' and cl.Type = 'Q'"


280               RecOpenServer 0, tb, sql
290               If Not tb Is Nothing Then
300                   Do While Not tb.EOF
310                       g.AddItem tb!Code & vbTab & tb!Description & vbTab & tb!DateTimeOfRecord
320                       tb.MoveNext
330                   Loop
340               End If
                FillComboBox (2)

350           End If

360       End If
370       Set tb = Nothing
          

380       Exit Sub

cmdSearch_Click_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmChangePMQ", "cmdSearch_Click", intEL, strES

End Sub

Private Sub cmdUpdate_Click()

         Dim sql As String
         Dim tb As Recordset
         Dim i As Integer
         Dim j As Integer
         Dim row As String, col As String
         Dim Code As String
         Dim newCode As String
         Dim CaseId As String
         Dim OldListID As String
         Dim NewListID As String
         Dim NewCodeArr() As String
         
         
10       On Error GoTo cmdUpdate_Click_Error
20       CaseId = Replace(Trim(txtCaseID.Text), "/", "")
30       CaseId = Replace(CaseId, " ", "")
40       For i = 1 To g.Rows - 1
50          g.row = i
            
60          For j = 0 To g.Cols - 1
70             g.col = j
80             If g.CellBackColor = vbYellow Then
90                row = i
100               Exit For
110            End If
120         Next
130      Next
         
140      Code = g.TextMatrix(row, 0)
150      sql = "SELECT ListID FROM Lists WHERE Code = '" & Code & "'"
160      Set tb = New Recordset
170      RecOpenServer 0, tb, sql
             
180      If Not tb Is Nothing Then
190         If Not tb.EOF Then
200            OldListID = tb!ListId & ""
210         End If
             
220      End If
             
230      sql = "SELECT "
240      If Trim(cmbCodesDes.Text) <> "" Then
250         NewCodeArr = Split(cmbCodesDes.Text, "-")
260         newCode = Trim(NewCodeArr(0))
270         sql = "SELECT ListID FROM Lists WHERE Code = '" & newCode & "'"
280         Set tb = New Recordset
290         RecOpenServer 0, tb, sql
                
300         If Not tb Is Nothing Then
310             If Not tb.EOF Then
320                 NewListID = tb!ListId & ""
330             End If
340         End If
350         sql = "UPDATE CaseListLink SET ListID = '" & NewListID & "', DateTimeOfRecord = GETDATE() WHERE CaseID = '" & CaseId & "' AND ListID = '" & OldListID & "'"
                
360         Cnxn(0).Execute (sql)
             
370      End If

380      cmdSearch_Click
390      cmdUpdate.Enabled = False
cmdUpdate_Click_Error:
             Dim strES As String
             Dim intEL As Long

400          intEL = Erl
410          strES = Err.Description
420          LogError "frmEditAll", "LoadAllDetails", intEL, strES

          
End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error
          ChangeFont Me, "Arial"
20        cmbCodesDes.Enabled = False
          opt(0).Value = True
          InitilizeGrid
          colorFlag = False
          'If blnIsTestMode Then EnableTestMode Me
30        Exit Sub
          Permissions

Form_Load_Error:


          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmChangePMQ", "Form_Load", intEL, strES

End Sub

Private Sub Permissions()


   On Error GoTo Permissions_Error
    If UCase(UserMemberOf) = "MANAGER" Or UCase(UserMemberOf) = "MANAGERS" Then
        cmdUpdate.Visible = False
    End If
    
Permissions_Error:
       Dim strES As String
       Dim intEL As Long

       intEL = Erl
       strES = Err.Description
       LogError "frmChangePMQ", "Permissions", intEL, strES

    
End Sub

Private Sub InitilizeGrid()
10        On Error GoTo InitilizeGrid_Error
20        If opt(0).Value = True Then
30            g.Clear
40            g.Cols = 3
50            g.TextMatrix(0, 0) = "P Code"
60            g.TextMatrix(0, 1) = "Description"
70            g.TextMatrix(0, 2) = "Date/Time"
80            g.ColWidth(0) = 1000
90            g.ColWidth(1) = 4000
100           g.ColWidth(2) = 2000
110       ElseIf opt(1).Value = True Then
120           g.Clear
130           g.Cols = 4
140           g.TextMatrix(0, 0) = "M Code"
150           g.TextMatrix(0, 1) = "Description"
160           g.TextMatrix(0, 2) = "Tissue"
170           g.TextMatrix(0, 3) = "Date/Time"
180           g.ColWidth(0) = 1000
190           g.ColWidth(1) = 3000
200           g.ColWidth(2) = 500
210           g.ColWidth(3) = 2000
220       ElseIf opt(2).Value = True Then
230           g.Clear
240           g.Cols = 3
250           g.TextMatrix(0, 0) = "Q Code"
260           g.TextMatrix(0, 1) = "Description"
270           g.TextMatrix(0, 2) = "Date/Time"
280           g.ColWidth(0) = 1000
290           g.ColWidth(1) = 4000
300           g.ColWidth(2) = 2000
310       End If
320       Exit Sub

InitilizeGrid_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmChangePMQ", "InitilizeGrid", intEL, strES

End Sub


Private Sub g_Click()
          Dim rowIndex As Integer
          Dim i As Integer
          Dim defaultBackColor As Long
          Dim defaultForeColor As Long
          Dim targetColor As Long
          Dim j As Integer
         

10        On Error GoTo g_Click_Error

        
20        defaultBackColor = vbWindowBackground
30        defaultForeColor = vbWindowText
40        For i = 1 To g.Rows - 1
50        g.row = i
60            For j = 0 To g.Cols - 1
70               g.col = j
80               If g.CellBackColor = vbYellow Then
90                  cmdSearch_Click
100              End If
110           Next
120       Next

130       If g.MouseRow > 0 Then
              
140           rowIndex = g.MouseRow
150           g.row = rowIndex

              
160           g.col = 0
170           If g.CellBackColor = vbYellow Then
180               targetColor = defaultBackColor
190               cmbCodesDes.Enabled = False
                  
                  
200           Else
210               targetColor = vbYellow
220               cmbCodesDes.Enabled = True
230               cmdUpdate.Enabled = True
240           End If

250           For i = 0 To g.Cols - 1
260               g.col = i
270               g.CellBackColor = targetColor
280           Next
290       End If

300       Exit Sub
g_Click_Error:
          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmChangePMQ", "g_Click_Error", intEL, strES

End Sub

Private Sub opt_Click(Index As Integer)

10        On Error GoTo InitilizeGrid_Error
          cmbCodesDes.Enabled = False
          cmbCodesDes.Clear
20        If Index = 0 Then
30            g.Clear
40            g.Cols = 3
50            g.TextMatrix(0, 0) = "P Code"
60            g.TextMatrix(0, 1) = "Description"
70            g.TextMatrix(0, 2) = "Date/Time"
80            g.ColWidth(0) = 1000
90            g.ColWidth(1) = 4000
100           g.ColWidth(2) = 1500
110       ElseIf Index = 1 Then
120           g.Clear
130           g.Cols = 4
140           g.TextMatrix(0, 0) = "M Code"
150           g.TextMatrix(0, 1) = "Description"
160           g.TextMatrix(0, 2) = "Tissue"
170           g.TextMatrix(0, 3) = "Date/Time"
180           g.ColWidth(0) = 1000
190           g.ColWidth(1) = 3000
200           g.ColWidth(2) = 500
210           g.ColWidth(3) = 2000
220       ElseIf Index = 2 Then
230           g.Clear
240           g.Cols = 3
250           g.TextMatrix(0, 0) = "Q Code"
260           g.TextMatrix(0, 1) = "Description"
270           g.TextMatrix(0, 2) = "Date/Time"
280           g.ColWidth(0) = 1000
290           g.ColWidth(1) = 4000
300           g.ColWidth(2) = 1500
310       End If
320       Exit Sub

InitilizeGrid_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmChangePMQ", "InitilizeGrid", intEL, strES

          
End Sub

Private Sub txtCaseId_KeyPress(KeyAscii As Integer)
          Dim Val As String
10        Val = "Tullamore"
           KeyAscii = Asc(UCase(Chr(KeyAscii)))
20        If UCase(Val) = "TULLAMORE" Then
30            Call ValidateTullCaseId(KeyAscii, Me)
40        Else
50            Call ValidateLimCaseId(KeyAscii, Me)
60        End If

70        Exit Sub
End Sub

Private Sub FillComboBox(ByVal Index As Integer)
          Dim sql As String
          Dim tb As Recordset
10        Set tb = New Recordset
20        On Error GoTo FillComboBox_Error
30        cmbCodesDes.Clear
40        If Index = 0 Then
50            sql = "SELECT Code, Description FROM Lists WHERE InUse = 1 and ListType = 'P'"
60            RecOpenServer 0, tb, sql
            
70            If Not tb Is Nothing Then
80                Do While Not tb.EOF
90                    cmbCodesDes.AddItem (tb!Code & "-" & tb!Description)
100                   tb.MoveNext
110               Loop
120           End If
         
130       ElseIf Index = 1 Then
140           sql = "SELECT Code, Description FROM Lists WHERE InUse = 1 and ListType = 'M'"
150           RecOpenServer 0, tb, sql
            
160           If Not tb Is Nothing Then
170               Do While Not tb.EOF
180                   cmbCodesDes.AddItem (tb!Code & "-" & tb!Description)
190                   tb.MoveNext
200               Loop
210           End If
         
220       ElseIf Index = 2 Then
230           sql = "SELECT Code, Description FROM Lists WHERE InUse = 1 and ListType = 'Q'"
240           RecOpenServer 0, tb, sql
            
250           If Not tb Is Nothing Then
260               Do While Not tb.EOF
270                   cmbCodesDes.AddItem (tb!Code & "-" & tb!Description)
280                   tb.MoveNext
290               Loop
300           End If
         
310       End If
     
320       Exit Sub

FillComboBox_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmChangePMQ", "InitilizeGrid", intEL, strES

End Sub



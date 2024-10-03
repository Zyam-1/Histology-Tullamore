VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmCaseStatus 
   Caption         =   "FrmCaseStatus"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16140
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   16140
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCheck msgUpdate 
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   975
      Left            =   11160
      Picture         =   "FrmCaseStatus.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Index           =   0
      Left            =   13680
      Picture         =   "FrmCaseStatus.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   975
      Index           =   0
      Left            =   12360
      Picture         =   "FrmCaseStatus.frx":048C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   840
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5895
      Left            =   300
      TabIndex        =   13
      Top             =   3030
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   10398
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   16711680
      ForeColorSel    =   65280
      GridLines       =   3
      GridLinesFixed  =   3
      MouseIcon       =   "FrmCaseStatus.frx":08A7
   End
   Begin VB.ComboBox cmbMessageType 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmCaseStatus.frx":1981
      Left            =   5880
      List            =   "FrmCaseStatus.frx":1988
      TabIndex        =   12
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Frame Frame 
      Caption         =   "Between Dates"
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   10215
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   855
         Left            =   1920
         Picture         =   "FrmCaseStatus.frx":19CC
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   915
      End
      Begin VB.ComboBox cmbStatusType 
         Height          =   315
         Index           =   0
         ItemData        =   "FrmCaseStatus.frx":1D4A
         Left            =   5520
         List            =   "FrmCaseStatus.frx":1D51
         TabIndex        =   11
         Top             =   1080
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   375
         Index           =   1
         Left            =   7800
         TabIndex        =   8
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119537665
         CurrentDate     =   45501
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   7
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119537665
         CurrentDate     =   45501
      End
      Begin VB.TextBox txtCaseId 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label 
         Caption         =   "Message Type"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Status Type"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   7440
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "Case Id"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Label lblLoggedIn 
      Caption         =   "Logged In: "
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   13680
      X2              =   13680
      Y1              =   2640
      Y2              =   2880
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   13200
      X2              =   13680
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Mark as read"
      Height          =   255
      Index           =   4
      Left            =   12120
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "From"
      Height          =   615
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "FrmCaseStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chkArray() As CheckBox
Private Sub InitilizeGrid()

10        On Error GoTo InitilizeGrid_Error
20            g.Clear
30            g.Cols = 7
40            g.TextMatrix(0, 0) = "Case Id"
50            g.TextMatrix(0, 1) = "Status Type"
60            g.TextMatrix(0, 2) = "Message Type"
70            g.TextMatrix(0, 3) = "Status Message"
80            g.TextMatrix(0, 4) = "Created By"
90            g.TextMatrix(0, 5) = "Date/Time"
100           g.TextMatrix(0, 6) = ""
110           g.ColWidth(0) = 1500
120           g.ColWidth(1) = 1500
130           g.ColWidth(2) = 1500
140           g.ColWidth(3) = 5500
150           g.ColWidth(4) = 1500
160           g.ColWidth(5) = 1800
170           g.ColWidth(6) = 500
180          Exit Sub

InitilizeGrid_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmChangePMQ", "InitilizeGrid", intEL, strES

End Sub



Private Sub cmdExit_Click(Index As Integer)
Me.Hide
End Sub

Private Sub cmdExport_Click()
    ExportFlexGrid g, Me
End Sub

Public Sub cmdSearch_Click()
    Dim sql As String
    Dim whereClause As String
    Dim tb As Recordset
    Dim CaseId As String
10  On Error GoTo cmdSearch_Click_Error
    CaseId = Replace(txtCaseId.Text, "/", "")
    CaseId = Replace(CaseId, " ", "")

20  sql = "SELECT * FROM VentanaStatusUpdate WHERE 1=1"


30  If Trim(CaseId) <> "" Then
40      sql = sql & " AND CaseID = '" & Trim(CaseId) & "'"
50  End If

60  If Trim(cmbStatusType(0).Text) <> "" Then
70      sql = sql & " AND StatusType = '" & Trim(cmbStatusType(0).Text) & "'"
80  End If

90  If Trim(cmbMessageType(1).Text) <> "" Then
100     sql = sql & " AND StatusMessageType = '" & Trim(cmbMessageType(1).Text) & "'"
110 End If

120 If IsDate(dtFrom(0).Value) And IsDate(dtTo(1).Value) Then
130     sql = sql & " AND CreatedDateTime BETWEEN '" & Format(dtFrom(0).Value, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(dtTo(1).Value, "yyyy-mm-dd 23:59:59") & "'"
140 End If


150 Set tb = New Recordset
160 RecOpenServer 0, tb, sql


170 g.Clear
180 InitilizeGrid
190 g.Rows = 1

200 If Not tb.EOF Then
210     tb.MoveFirst
        Dim i As Integer
220     i = 1
230     Do While Not tb.EOF

240         g.AddItem tb!CaseId & vbTab & tb!StatusType & vbTab & tb!StatusMessageType & vbTab & tb!StatusMessage & vbTab & tb!CreatedBy & vbTab & tb!CreatedDateTime
250         tb.MoveNext
260     Loop
270 End If
        

280 tb.Close
290 Set tb = Nothing
    AdjustGridSize


300 Exit Sub

cmdSearch_Click_Error:
    Dim strES As String
    Dim intEL As Integer

310 intEL = Erl
320 strES = Err.Description
330 LogError "frmSearch", "cmdSearch_Click", intEL, strES, sql


End Sub




Private Sub Form_Activate()
      dtFrom(0).Value = Now
      dtTo(1).Value = Now
      If Trim(txtCaseId.Text) <> "" Then
        cmdSearch_Click
      End If
End Sub

Private Sub Form_Load()
10    ChangeFont Me, "Arial"
20    InitilizeGrid
30    FillComboStatusType
40    FillComboStatusMessageType
50    MergeCells
60    If blnIsTestMode Then EnableTestMode Me
      lblLoggedIn.Caption = lblLoggedIn.Caption & UserName
      'FillCheckBox


End Sub

Private Sub MergeCells()

    Dim i As Integer
    Dim sameData As String
    With g
        .MergeCells = flexMergeFree
        .MergeCol(0) = True ' Enable merging for the first column
        
        ' Loop through the rows starting from the second row
        For i = 1 To .Rows - 1
            ' Check if the current cell value is the same as the previous row's cell value
            If Trim(.TextMatrix(i, 0)) = Trim(.TextMatrix(i - 1, 0)) Then
                ' Set the current cell value to a placeholder value (or keep it the same)
                .TextMatrix(i, 0) = Trim(.TextMatrix(i - 1, 0))
            End If
        Next i
    End With
End Sub

Private Sub AdjustGridSize()
    


   On Error GoTo AdjustGridSize_Error
   Dim i As Integer
   
   For i = 1 To g.Rows - 1
    g.RowHeight(i) = 500
   
   Next
   g.FontWidth = 3.5
    

 

AdjustGridSize_Error:
       Dim strES As String
       Dim intEL As Long

       intEL = Erl
       strES = Err.Description
       LogError "frmCaseStatus", "AdjustGridSize", intEL, strES

    
End Sub




Private Sub FillComboStatusType()
    Dim sql As String
    Dim tb As Recordset

10    On Error GoTo FillComboStatusType_Error

20    sql = "Select Distinct StatusType from VentanaStatusUpdate"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

    
50    Do While Not tb.EOF
60        cmbStatusType(0).AddItem tb!StatusType & ""
70        tb.MoveNext
80    Loop

90    Exit Sub

FillComboStatusType_Error:
    Dim strES As String
    Dim intEL As Integer

100    intEL = Erl
120    strES = Err.Description
130    LogError "frmCaseStatus", "FillComboStatusType", intEL, strES, sql
End Sub



Private Sub FillComboStatusMessageType()
    Dim sql As String
    Dim tb As Recordset

10    On Error GoTo FillComboStatusMessageType_Error

20    sql = "Select Distinct StatusMessageType from VentanaStatusUpdate"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

    
50    Do While Not tb.EOF
60        cmbMessageType(1).AddItem tb!StatusMessageType & ""
70        tb.MoveNext
80    Loop

90    Exit Sub

FillComboStatusMessageType_Error:
    Dim strES As String
    Dim intEL As Integer

100    intEL = Erl
120    strES = Err.Description
130    LogError "frmCaseStatus", "FillComboStatusMessageType", intEL, strES, sql
End Sub

Private Sub Picture1_Click()

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


End Sub

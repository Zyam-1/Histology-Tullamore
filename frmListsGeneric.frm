VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmListsGeneric 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cellular Pathology- List of Generic"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8520
      Picture         =   "frmListsGeneric.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   855
      Left            =   8520
      Picture         =   "frmListsGeneric.frx":01CF
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   8520
      Picture         =   "frmListsGeneric.frx":02D1
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   1095
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Generic"
      Height          =   1455
      Left            =   100
      TabIndex        =   3
      Top             =   360
      Width           =   8085
      Begin VB.ComboBox cmbSource 
         Height          =   315
         Left            =   4800
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   855
         Left            =   7080
         Picture         =   "frmListsGeneric.frx":0613
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   975
         Width           =   1545
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1740
         MaxLength       =   400
         TabIndex        =   1
         Top             =   975
         Width           =   5085
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4830
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1770
         TabIndex        =   4
         Top             =   720
         Width           =   390
      End
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8700
      Top             =   4350
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8700
      Top             =   3570
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6855
      Left            =   105
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   12091
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmListsGeneric.frx":0951
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   60
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1080
      TabIndex        =   15
      Top             =   8880
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   9240
      Picture         =   "frmListsGeneric.frx":09F4
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   9240
      Picture         =   "frmListsGeneric.frx":0CCA
      Top             =   720
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmListsGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer

Private pListType As String
Private pListTypeName As String    'Singular name
Private pListTypeNames As String    'Plural name
Private pListGenericName As String    'List
Private pExternal As Boolean    'External events associated with list item

Dim bDoNotEdit As Boolean


Private Sub FillG()



On Error GoTo FillG_Error

g.Rows = 2
g.AddItem ""
g.RemoveItem 1

Select Case UCase$(pListTypeName)
    Case "CLINICIAN", "WARD", "CORONER"
  LoadSourceItemLists
Case Else
  LoadLists
End Select


If g.Rows > 2 Then
  g.RemoveItem 1
End If

Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "FillG", intEL, strES


End Sub
Private Sub LoadLists()
    Dim s As String
    Dim tb As Recordset
    Dim sql As String
    Dim sn As Recordset

On Error GoTo LoadLists_Error

sql = "SELECT * FROM Lists WHERE " & _
    "ListType = '" & pListType & "' " & _
    "ORDER BY ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  s = IIf(tb!InUse & 0, "Yes", "No") & vbTab & _
      tb!Code & "" & vbTab & _
      tb!Description & ""
  g.AddItem s
  g.row = g.Rows - 1
  If g.ColWidth(3) <> 0 Then
      g.col = 3
      g.CellPictureAlignment = flexAlignCenterCenter
      If Not IsNull(tb!External) Then
          If tb!External Then
              Set g.CellPicture = imgSquareTick.Picture
          Else
              Set g.CellPicture = imgSquareCross.Picture
          End If
      Else
          Set g.CellPicture = imgSquareCross.Picture
      End If
  End If
  If g.ColWidth(4) <> 0 Then
      sql = "SELECT Count(StainCode) as TotStains FROM DefaultStains WHERE TissueCodeListId = '" & tb!ListId & "'"
      Set sn = New Recordset
      RecOpenServer 0, sn, sql

      If Not sn.EOF Then
          g.TextMatrix(g.row, 4) = sn!TotStains & ""
      End If
  End If
  g.TextMatrix(g.row, 5) = tb!Levels & ""

  g.TextMatrix(g.row, 6) = tb!ListId & ""

  If g.ColWidth(8) <> 0 Then
      g.col = 8
      g.CellPictureAlignment = flexAlignCenterCenter
      If Not IsNull(tb!Cancerous) Then
          If tb!Cancerous Then
              Set g.CellPicture = imgSquareTick.Picture
          Else
              Set g.CellPicture = imgSquareCross.Picture
          End If
      Else
          Set g.CellPicture = imgSquareCross.Picture
      End If
  End If

  If g.ColWidth(9) <> 0 Then
      g.col = 9
      g.CellPictureAlignment = flexAlignCenterCenter
      If Not IsNull(tb!ShowWorkList) Then
          If tb!ShowWorkList Then
              Set g.CellPicture = imgSquareTick.Picture
          Else
              Set g.CellPicture = imgSquareCross.Picture
          End If
      Else
          Set g.CellPicture = imgSquareCross.Picture
      End If
  End If

  tb.MoveNext
Loop

Exit Sub

LoadLists_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "LoadLists", intEL, strES, sql


End Sub
Private Sub LoadSourceItemLists()
    Dim s As String
    Dim tb As Recordset
    Dim sql As String

On Error GoTo LoadSourceItemLists_Error

sql = "SELECT * FROM SourceItemLists WHERE " & _
    "ListType = '" & pListType & "' " & _
    "ORDER BY ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  s = IIf(tb!InUse & 0, "Yes", "No") & vbTab & _
      tb!Code & "" & vbTab & _
      tb!Description & ""
  g.AddItem s
  g.row = g.Rows - 1

  g.TextMatrix(g.row, 6) = tb!ListId & ""

  g.TextMatrix(g.row, 7) = ListDescriptionFor("Source", tb!Source & "")

  tb.MoveNext
Loop

Exit Sub

LoadSourceItemLists_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "LoadSourceItemLists", intEL, strES, sql

End Sub
Private Sub FireDown()

    Dim n As Integer
    Dim s As String
    Dim X As Integer
    Dim VisibleRows As Integer

If g.row = g.Rows - 1 Then Exit Sub
n = g.row

FireCounter = FireCounter + 1
If FireCounter > 5 Then
  tmrDown.Interval = 100
End If

VisibleRows = g.Height \ g.RowHeight(1) - 1

g.Visible = False

s = ""
For X = 0 To g.Cols - 1
  s = s & g.TextMatrix(n, X) & vbTab
Next
s = Left$(s, Len(s) - 1)

g.RemoveItem n
If n < g.Rows Then
  g.AddItem s, n + 1
  g.row = n + 1
Else
  g.AddItem s
  g.row = g.Rows - 1
End If

For X = 0 To g.Cols - 1
  g.col = X
  g.CellBackColor = vbYellow
Next

If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
  If g.row - VisibleRows + 1 > 0 Then
      g.TopRow = g.row - VisibleRows + 1
  End If
End If

g.Visible = True

cmdSave.Visible = True

End Sub

Private Sub FireUp()

    Dim n As Integer
    Dim s As String
    Dim X As Integer

If g.row = 1 Then Exit Sub

FireCounter = FireCounter + 1
If FireCounter > 5 Then
  tmrUp.Interval = 100
End If

n = g.row

g.Visible = False

s = ""
For X = 0 To g.Cols - 1
  s = s & g.TextMatrix(n, X) & vbTab
Next
s = Left$(s, Len(s) - 1)

g.RemoveItem n
g.AddItem s, n - 1

g.row = n - 1
For X = 0 To g.Cols - 1
  g.col = X
  g.CellBackColor = vbYellow
Next

If Not g.RowIsVisible(g.row) Then
  g.TopRow = g.row
End If

g.Visible = True

cmdSave.Visible = True

End Sub





Private Sub cmdAdd_Click()

    Dim n As Integer
    Dim sql As String
    Dim tb As Recordset

On Error GoTo cmdAdd_Click_Error

txtCode = Trim$(UCase$(txtCode))
txtText = Trim$(txtText)

If txtCode = "" Then
  txtCode.SetFocus
  Exit Sub
End If

If txtText = "" Then
  txtText.SetFocus
  Exit Sub
End If

For n = 1 To g.Rows - 1
  If g.TextMatrix(n, 1) = txtCode Then
      frmMsgBox.Msg "Code already used, Please try another", , , mbExclamation
      If TimedOut Then Unload Me: Exit Sub
      txtCode = ""
      txtCode.SetFocus
      Exit Sub
  End If
Next
Select Case pListType
    Case "IS", "RS", "SS"
  sql = "SELECT * FROM Lists WHERE Code = '" & txtCode & "' " & _
        "AND (ListType = 'IS' OR ListType = 'RS' OR ListType = 'SS')"
  Set tb = New Recordset
  RecOpenServer 0, tb, sql

  If Not tb.EOF Then
      frmMsgBox.Msg "This code is already used for Immunohistochemical, Routine or Special Stain, Please try another", , , mbExclamation
      If TimedOut Then Unload Me: Exit Sub
      txtCode = ""
      txtCode.SetFocus
      Exit Sub
  End If
End Select

g.AddItem "Yes" & vbTab & txtCode & vbTab & txtText
g.row = g.Rows - 1
If g.ColWidth(3) <> 0 Then
  g.col = 3
  g.CellPictureAlignment = flexAlignCenterCenter
  Set g.CellPicture = imgSquareCross.Picture
End If

If cmbSource.Visible = True Then
  g.TextMatrix(g.row, 7) = cmbSource
End If



If g.Rows > 2 And g.TextMatrix(1, 1) = "" Then
  g.RemoveItem 1
End If

txtCode = ""
txtText = ""


cmdSave.Visible = True

txtCode.SetFocus

Exit Sub

cmdAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "cmdadd_Click", intEL, strES

End Sub





Private Sub cmdDelete_Click()

    Dim Y As Integer
    Dim sql As String
    Dim s As String

On Error GoTo cmdDelete_Click_Error

g.col = 0
For Y = 1 To g.Rows - 1
  g.row = Y
  If g.CellBackColor = vbYellow Then
      s = "Delete " & g.TextMatrix(Y, 2) & vbCrLf & _
        " From " & pListTypeNames & " ?"

      Answer = frmMsgBox.Msg(s, mbYesNo, , mbQuestion)

      If TimedOut Then Unload Me: Exit Sub
      If Answer = 1 Then
          Select Case UCase$(pListTypeName)
          Case "CLINICIAN", "WARD", "CORONER"
              sql = "Delete from SourceItemLists where " & _
                    "ListType = '" & pListType & "' " & _
                    "AND Code = '" & g.TextMatrix(Y, 1) & "' " & _
                    "AND Source = '" & ListCodeFor("Source", g.TextMatrix(Y, 7)) & "'"
              Cnxn(0).Execute sql
          Case Else
              sql = "Delete from Lists where " & _
                    "ListType = '" & pListType & "' " & _
                    "and Code = '" & g.TextMatrix(Y, 1) & "'"
              Cnxn(0).Execute sql
          End Select

      End If
      Exit For
  End If
Next

cmdDelete.Enabled = False
FillG

Exit Sub

cmdDelete_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "cmdDelete_Click", intEL, strES, sql


End Sub




Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

    Dim Y As Integer


On Error GoTo cmdSave_Click_Error

For Y = 1 To g.Rows - 1
  If g.TextMatrix(Y, 1) <> "" Then
      Select Case UCase$(pListTypeName)
      Case "CLINICIAN", "WARD", "CORONER"
          SaveSourceItemLists Y
      Case Else
          SaveLists Y
      End Select
  End If
Next

FillG

txtCode = ""
txtText = ""
txtCode.SetFocus

cmdSave.Visible = False
cmdDelete.Enabled = False

Exit Sub

cmdSave_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "cmdSave_Click", intEL, strES


End Sub
Private Sub SaveSourceItemLists(Y As Integer)
    Dim HospSourceCode As String
    Dim tb As Recordset
    Dim sql As String

On Error GoTo SaveSourceItemLists_Error

HospSourceCode = ListCodeFor("Source", g.TextMatrix(Y, 7))

sql = "SELECT * FROM SourceItemLists WHERE " & _
    "ListType = '" & pListType & "' " & _
    "AND Code = '" & g.TextMatrix(Y, 1) & "' " & _
    "AND Source = '" & HospSourceCode & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
End If
tb!InUse = IIf(g.TextMatrix(Y, 0) = "Yes", 1, 0)
tb!Code = g.TextMatrix(Y, 1)
tb!ListType = pListType
tb!Description = g.TextMatrix(Y, 2)
tb!ListOrder = Y
tb!Source = HospSourceCode
tb!UserName = UserName
tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")


tb.Update

Exit Sub

SaveSourceItemLists_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "SaveSourceItemLists", intEL, strES, sql

End Sub
Private Sub SaveLists(Y As Integer)
    Dim tb As Recordset
    Dim sql As String

On Error GoTo SaveLists_Error

sql = "SELECT * FROM Lists WHERE " & _
    "ListType = '" & pListType & "' " & _
    "AND Code = '" & g.TextMatrix(Y, 1) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  tb.AddNew
End If
tb!InUse = IIf(g.TextMatrix(Y, 0) = "Yes", 1, 0)
tb!Code = g.TextMatrix(Y, 1)
tb!ListType = pListType
tb!Description = g.TextMatrix(Y, 2)
tb!ListOrder = Y
g.row = Y
g.col = 3
tb!External = IIf(g.CellPicture = imgSquareTick.Picture, 1, 0)
tb!Levels = g.TextMatrix(Y, 5)
g.col = 8
tb!Cancerous = IIf(g.CellPicture = imgSquareTick.Picture, 1, 0)
g.col = 9
tb!ShowWorkList = IIf(g.CellPicture = imgSquareTick.Picture, 1, 0)
tb!UserName = UserName
tb!DateTimeOfRecord = Format(Now, "dd/MMM/yyyy hh:mm:ss")


tb.Update

Exit Sub

SaveLists_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "SaveLists", intEL, strES, sql

End Sub


Private Sub Form_Load()


10    ChangeFont Me, "Arial"
20    g.Font.Bold = True

30    If pListType = "" Then
40      MsgBox "pListType not set"
50    End If
60    If pListTypeName = "" Then
70      MsgBox "pListTypeName not set"
80    End If
90    If pListTypeNames = "" Then
100     MsgBox "pListTypeNames not set"
110   End If

      'frmListsGeneric_ChangeLanguage pListTypeName, pListTypeNames
      'FrameAdd.Caption = "Add"
      'FrameAdd.Caption = FrameAdd.Caption & " " & pListGenericName
      'Me.Caption = "Histology --- Lists"
      'Me.Caption = Me.Caption & " (" & pListGenericName & ")"
120   lblLoggedIn = UserName

130   If pListType = "P" Then
140       Me.Caption = "Histology --- Lists ( P Codes)"
150   ElseIf pListType = "M" Then
160       Me.Caption = "Histology --- Lists ( M Codes)"
170   ElseIf pListType = "Q" Then
180       Me.Caption = "Histology --- Lists ( Q Codes)"
190   ElseIf pListType = "T" Then
200       Me.Caption = "Histology --- Lists ( T Codes)"
210   ElseIf pListType = "RefReason" Then
220       Me.Caption = "Histology --- Lists ( Reasons For Referral)"
230   ElseIf pListType = "RefTo" Then
240       Me.Caption = "Histology --- Lists ( Referred To)"
250   ElseIf pListType = "RS" Then
260       Me.Caption = "Histology --- Lists ( Routine Stains)"
270   ElseIf pListType = "SS" Then
280       Me.Caption = "Histology --- Lists ( Special Stains)"
290   ElseIf pListType = "IS" Then
300       Me.Caption = "Histology --- Lists ( Immunohistochemical Stains)"
310   ElseIf pListType = "Ward" Then
320       Me.Caption = "Histology --- Lists ( Wards)"
330   ElseIf pListType = "Coroner" Then
340       Me.Caption = "Histology --- Lists ( Coroners)"
350   ElseIf pListType = "Clinician" Then
360       Me.Caption = "Histology --- Lists ( Clinicians)"
370   ElseIf pListType = "County" Then
380       Me.Caption = "Histology --- Lists ( Counties)"
390   ElseIf pListType = "Source" Then
400       Me.Caption = "Histology --- Lists ( Source)"
410   ElseIf pListType = "Orientation" Then
420       Me.Caption = "Histology --- Lists ( Orientation)"
430   ElseIf pListType = "Processor" Then
440       Me.Caption = "Histology --- Lists ( Processor)"
450   ElseIf pListType = "DiscrepType" Then
460       Me.Caption = "Histology --- Lists ( Discrepancy Types)"
470   ElseIf pListType = "DiscrepRes" Then
480       Me.Caption = "Histology --- Lists ( Discrepancy Resolutions)"
490   End If

500   InitializeGrid

510   If pListType <> "Q" And pListType <> "IS" And pListType <> "SS" Then
520     g.ColWidth(3) = 0
530   End If

540   If pListType <> "T" Then
550     g.ColWidth(4) = 0
560     g.ColWidth(5) = 0
570   Else
580     g.ColWidth(2) = 3400
590   End If

600   If pListType <> "M" Then
610     g.ColWidth(8) = 0
620   End If

630   If UCase(pListType) <> "REFREASON" Then
640     g.ColWidth(9) = 0
650   End If

660   Select Case UCase$(pListTypeName)
          Case "CLINICIAN", "WARD", "CORONER"
670     lblSource.Visible = True
680     cmbSource.Visible = True
690     FillSource
700   Case Else
710     lblSource.Visible = False
720     cmbSource.Visible = False
730     g.ColWidth(7) = 0
740   End Select


750   txtEdit = ""
760   bDoNotEdit = False

770   FillG
780   If UCase(pListType) = "T" Then
790     FillDefaultBlocks
      g.ColAlignment(g.Cols - 1) = flexAlignLeftCenter
800   End If
810   If blnIsTestMode Then EnableTestMode Me
End Sub

Private Sub FillDefaultBlocks()
        

10       On Error GoTo FillDefaultBlocks_Error
         Dim sql As String
         Dim tb As Recordset

20       g.Cols = g.Cols + 1
         'g.ColWidth(1) = 1500
         g.ColWidth(2) = 1700
         Dim Code As String
         Dim i As Integer
30       i = 1
40       g.TextMatrix(0, g.Cols - 1) = "Default Blocks"
         g.ColWidth(g.Cols - 1) = 1100
50       For i = 1 To g.Rows - 1
60          Code = Trim(g.TextMatrix(i, 1))
70          sql = "SELECT DefaultBlocks FROM Lists WHERE Code = '" & Code & "'"
80          Set tb = New Recordset
90          RecOpenServer 0, tb, sql
100         If Not tb Is Nothing Then
110            If Not tb.EOF Then
120               g.TextMatrix(i, g.Cols - 1) = tb!DefaultBlocks & ""
130            End If
140         End If
150      Next
         
FillDefaultBlocks_Error:
             Dim strES As String
             Dim intEL As Long

160          intEL = Erl
170          strES = Err.Description
180          LogError "frmListsGeneric", "FillDefaultBlocks", intEL, strES

          
End Sub

Private Sub FillSource()
    Dim sql As String
    Dim tb As Recordset

On Error GoTo FillSource_Error

cmbSource.AddItem ""
sql = "SELECT * FROM Lists WHERE ListType = 'Source'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

Do While Not tb.EOF
  cmbSource.AddItem tb!Description & ""
  tb.MoveNext
Loop
cmbSource.ListIndex = -1

Exit Sub

FillSource_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmListsGeneric", "FillSource", intEL, strES, sql

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdSave.Visible Then
  Answer = frmMsgBox.Msg("Cancel without saving?", mbYesNo, , mbQuestion)
  If TimedOut Then Unload Me: Exit Sub
  If Answer = 2 Then
      Cancel = True
      Exit Sub
  End If
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

pListType = ""
pListTypeName = ""
pListTypeNames = ""
pExternal = False

End Sub

Private Sub g_Click()

    Static SortOrder As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim ySave As Integer

If g.MouseRow = 0 Then
  If SortOrder Then
      g.Sort = flexSortGenericAscending
  Else
      g.Sort = flexSortGenericDescending
  End If
  SortOrder = Not SortOrder
  Exit Sub
End If

cmdDelete.Enabled = True

If g.col = 0 Then
  g.TextMatrix(g.row, 0) = IIf(g.TextMatrix(g.row, 0) = "Yes", "No", "Yes")
  cmdSave.Visible = True
End If

If g.ColWidth(3) <> 0 Then
  If g.col = 3 Then
      If g.CellPicture = imgSquareTick.Picture Then
          Set g.CellPicture = imgSquareCross.Picture
      Else
          Set g.CellPicture = imgSquareTick.Picture
      End If
      cmdSave.Visible = True
      cmdDelete.Enabled = False
  End If
End If

If g.ColWidth(4) <> 0 Then
  If g.col = 4 Then
      With frmDefaultStains
          .TissueCodeListId = g.TextMatrix(g.row, 6)
          .TissueCode = g.TextMatrix(g.row, 1)
          .TissueName = g.TextMatrix(g.row, 2)
          .Show 1
      End With
  End If
End If

If g.ColWidth(5) <> 0 Then
  If g.col = 5 Then
      txtEdit = g.TextMatrix(g.row, 5)
      Call pEditGrid(32)
      cmdSave.Visible = True
  End If
End If

If g.ColWidth(8) <> 0 Then
  If g.col = 8 Then
      If g.CellPicture = imgSquareTick.Picture Then
          Set g.CellPicture = imgSquareCross.Picture
      Else
          Set g.CellPicture = imgSquareTick.Picture
      End If
      cmdSave.Visible = True
      cmdDelete.Enabled = False
  End If
End If

If g.ColWidth(9) <> 0 Then
  If g.col = 9 Then
      If g.CellPicture = imgSquareTick.Picture Then
          Set g.CellPicture = imgSquareCross.Picture
      Else
          Set g.CellPicture = imgSquareTick.Picture
      End If
      cmdSave.Visible = True
      cmdDelete.Enabled = False
  End If
End If

ySave = g.row

g.Visible = False
g.col = 0
For Y = 1 To g.Rows - 1
  g.row = Y
  If g.CellBackColor = vbYellow Then
      For X = 0 To g.Cols - 1
          g.col = X
          g.CellBackColor = 0
      Next
      Exit For
  End If
Next
g.row = ySave
g.Visible = True

For X = 0 To g.Cols - 1
  g.col = X
  g.CellBackColor = vbYellow
Next




End Sub



Private Sub pEditGrid(KeyAscii As Integer)
'
' Populate the textbox and position it.
'
With txtEdit
  Select Case KeyAscii
  Case 0 To 32
      '
      ' Edit the current text.
      '
      .Text = g
      .SelStart = 0
      .SelLength = 1000

  Case 8, 46, 48 To 57
      '
      ' Replace the current text but only
      ' if the user entered a number.
      '
      .Text = Chr(KeyAscii)
      .SelStart = 1
  Case Else
      '
      ' If an alpha character was entered,
      ' use a zero instead.
      '
      .Text = "0"
  End Select
End With
    '
    ' Show the textbox at the right place.
    '
With g
  If .CellWidth < 0 Then Exit Sub
  txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
  '
  ' NOTE:
  '   Depending on the style of the Grid Lines that you set, you
  '   may need to adjust the textbox position slightly. For example
  '   if you use raised grid lines use the following:
  '
  'txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
End With

txtEdit.Visible = True
txtEdit.SetFocus
End Sub

Private Sub g_KeyPress(KeyAscii As Integer)
'
' Display the textbox.
'
Call pEditGrid(KeyAscii)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'
' See what key was pressed in the textbox.
'
With g
  Select Case KeyCode
  Case 13   'ENTER
      .SetFocus
  Case 27   'ESC
      txtEdit.Visible = False
      .SetFocus
  Case 38   'Up arrow
      .SetFocus
      DoEvents
      If .row > .FixedRows Then
          bDoNotEdit = True
          .row = .row - 1
          bDoNotEdit = False
      End If
  Case 40   'Down arrow
      .SetFocus
      DoEvents
      If .row < .Rows - 1 Then
          bDoNotEdit = True
          .row = .row + 1
          bDoNotEdit = False
      End If
  End Select
End With
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'
' Delete carriage returns to get rid of beep
' and only allow numbers.
'
Select Case KeyAscii
    Case Asc(vbCr)
  KeyAscii = 0
Case 8, 46
Case 48 To 57
Case Else
  KeyAscii = 0
End Select
End Sub

Private Sub pSetCellValue()
'
' NOTE:
'       This code should be called anytime
'       the grid loses focus and the grid's
'       contents may change.  Otherwise, the
'       cell's new value may be lost and the
'       textbox may not line up correctly.
'
If txtEdit.Visible Then
  g.TextMatrix(g.row, 5) = txtEdit.Text
  txtEdit.Visible = False
End If
End Sub

Private Sub tmrDown_Timer()

FireDown

End Sub


Private Sub tmrUp_Timer()

FireUp

End Sub



Public Property Let ListType(ByVal Code As String)

pListType = Code

End Property
Public Property Let ListTypeName(ByVal strNewValue As String)

pListTypeName = strNewValue

End Property

Public Property Let ListGenericName(ByVal strNewValue As String)

pListGenericName = strNewValue

End Property

Public Property Let ListTypeNames(ByVal strNewValue As String)

pListTypeNames = strNewValue

End Property

Private Sub txtCode_KeyPress(KeyAscii As Integer)

KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

End Sub


Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
Call pSetCellValue
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)

KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

End Sub

Private Sub InitializeGrid()

With g
  .Rows = 2: .FixedRows = 1
  .Cols = 10: .FixedCols = 0
  .Rows = 1
  .Font.Size = fgcFontSize
  .Font.name = fgcFontName
  .ForeColor = fgcForeColor
  .BackColor = fgcBackColor
  .ForeColorFixed = fgcForeColorFixed
  .BackColorFixed = fgcBackColorFixed
  .ScrollBars = flexScrollBarBoth

  .TextMatrix(0, 0) = "In Use": .ColWidth(0) = 800: .ColAlignment(0) = flexAlignLeftCenter
  .TextMatrix(0, 1) = "Code": .ColWidth(1) = 1000: .ColAlignment(1) = flexAlignLeftCenter
  .TextMatrix(0, 2) = "Description": .ColWidth(2) = 4600: .ColAlignment(2) = flexAlignLeftCenter
  .TextMatrix(0, 3) = "External": .ColWidth(3) = 850: .ColAlignment(3) = flexAlignLeftCenter
  .TextMatrix(0, 4) = "Default Stains": .ColWidth(4) = 1300: .ColAlignment(4) = flexAlignLeftCenter
  .TextMatrix(0, 5) = "Levels": .ColWidth(5) = 900: .ColAlignment(5) = flexAlignLeftCenter
  .TextMatrix(0, 6) = "ListId": .ColWidth(6) = 0: .ColAlignment(6) = flexAlignLeftCenter
  .TextMatrix(0, 7) = "Source": .ColWidth(7) = 1200: .ColAlignment(7) = flexAlignLeftCenter
  .TextMatrix(0, 8) = "NCRI": .ColWidth(8) = 1100: .ColAlignment(8) = flexAlignLeftCenter
  .TextMatrix(0, 9) = "Show Worklist": .ColWidth(9) = 1500: .ColAlignment(9) = flexAlignLeftCenter
End With
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And Me.Top >= 0 Then

  Me.Top = 0
  Me.Left = Screen.Width / 2 - Me.Width / 2
End If
End Sub



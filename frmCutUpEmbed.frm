VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmCutUpEmbed 
   Appearance      =   0  'Flat
   Caption         =   "Cut-Up/Embedding WorkLog"
   ClientHeight    =   10425
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   13740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   855
      Left            =   9960
      Picture         =   "frmCutUpEmbed.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   960
      Width           =   765
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   12840
      Picture         =   "frmCutUpEmbed.frx":041B
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   765
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Update"
      Height          =   855
      Left            =   4560
      Picture         =   "frmCutUpEmbed.frx":075D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   765
   End
   Begin VB.CommandButton cmdReloadList 
      Caption         =   "&Reload"
      Height          =   855
      Left            =   10920
      Picture         =   "frmCutUpEmbed.frx":0A9B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   750
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   855
      Left            =   11880
      Picture         =   "frmCutUpEmbed.frx":0E0E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame fraData 
      Height          =   1695
      Left            =   4560
      TabIndex        =   9
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtPiecesAfterCutUp 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cmbOrientation 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pieces At Cut-Up"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Orientation"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1020
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7935
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   13996
      _Version        =   393216
      HighLight       =   0
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Frame fraCutUp 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ComboBox cmbProcessor 
         Height          =   315
         Left            =   1800
         TabIndex        =   23
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox cmbAssistedBy 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cmbCutUpBy 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblProcessor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Processor"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cut-Up By"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assisted By"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   1440
      End
   End
   Begin VB.Frame fraCutting 
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ComboBox cmbCutBy 
         Height          =   315
         Left            =   1440
         TabIndex        =   21
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cut By"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   765
         Width           =   1185
      End
   End
   Begin VB.Frame fraEmbed 
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtPiecesAfterEmbedding 
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox cmbEmbeddedBy 
         Height          =   315
         Left            =   2040
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pieces At Embedding"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1020
         Width           =   1770
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Embedded By"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   540
         Width           =   1770
      End
   End
   Begin VB.Label lblCaseId 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9960
      TabIndex        =   29
      Top             =   240
      Width           =   60
   End
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   28
      Top             =   10080
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   10080
      Width           =   1335
   End
End
Attribute VB_Name = "frmCutUpEmbed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" _
                                          (ByVal hwnd As Long) _
                                          As Long

Private m_booKeyCtrl As Boolean
Private m_booKeyShift As Boolean
Private m_booProcessSelected As Boolean
Private m_UpKey As Boolean
Private m_DownKey As Boolean
Private pPhase As String
Private pSingleEdit As Boolean
Private HideUnselected As Boolean
Private GridChanged As Boolean


Public Property Let Phase(ByVal strNewValue As String)

pPhase = strNewValue

End Property

Public Property Let SingleEdit(ByVal bNewValue As Boolean)

pSingleEdit = bNewValue

End Property



Private Sub cmbAssistedBy_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbCutBy_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub cmbCutUpBy_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub cmbEmbeddedBy_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub





Private Sub cmbProcessor_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    Dim i As Integer
    Dim s As Integer


On Error GoTo cmdAdd_Click_Error

cmdSave.Visible = True

With g
  m_booProcessSelected = True
  LockWindowUpdate .hwnd
  Dim lngColSave As Long: lngColSave = .col
  Dim lngRowSave As Long: lngRowSave = .row
  Dim lngRow As Long
  For lngRow = .FixedRows To .Rows - 1
      .row = lngRow
      If .CellBackColor = .BackColorSel Then
          If cmbCutUpBy <> "" Then
              .TextMatrix(.row, 9) = cmbCutUpBy
              GridChanged = True
          End If
          If cmbAssistedBy <> "" Then
              .TextMatrix(.row, 10) = cmbAssistedBy
              GridChanged = True
          End If
          If txtPiecesAfterCutUp <> "" Then
              .TextMatrix(.row, 5) = txtPiecesAfterCutUp    ' & " <- Selected"
              GridChanged = True
          End If
          If cmbOrientation <> "" Then
              .TextMatrix(.row, 6) = cmbOrientation
              GridChanged = True
          End If

          If txtPiecesAfterEmbedding <> "" Then
              .TextMatrix(.row, 7) = txtPiecesAfterEmbedding
              GridChanged = True
          End If
          If cmbEmbeddedBy <> "" Then
              .TextMatrix(.row, 8) = cmbEmbeddedBy
              GridChanged = True
          End If
          If cmbCutBy <> "" Then
              .TextMatrix(.row, 11) = cmbCutBy
              GridChanged = True
          End If
          If cmbProcessor <> "" Then
              .TextMatrix(.row, 13) = cmbProcessor
              GridChanged = True
          End If


      Else
          If HideUnselected Then
              .RowHeight(.row) = 0
          End If
      End If
  Next lngRow

  If pPhase = "Cut-Up" Then
      If Not pSingleEdit Then
          fraData.Visible = True
          cmdAdd.Left = 9000
      End If
  End If

  For lngRow = .FixedRows To .Rows - 1
      .row = lngRow
      ClearFields
      If .CellBackColor = .BackColorSel Then
          For s = .FixedCols To .Cols - 1
              .col = s
              .CellBackColor = .BackColor
              .CellForeColor = .ForeColor
          Next s
          If Not HideUnselected Then
              For i = .row + 1 To .Rows - 1
                  If .RowHeight(i) <> 0 Then
                      .row = i
                      For s = .FixedCols To .Cols - 1
                          .col = s
                          .CellBackColor = .BackColorSel
                          .CellForeColor = .ForeColorSel
                      Next s

                      Select Case UCase(pPhase)
                      Case "Cut-Up"

                          txtPiecesAfterCutUp.SetFocus
                      Case "Embedding"

                          txtPiecesAfterEmbedding.SetFocus
                      Case "Cutting"
                          cmbCutBy.SetFocus
                      End Select

                      Exit For

                  End If
              Next i
              Exit For
          End If
      End If
  Next lngRow


  HideUnselected = False
  .col = lngColSave
  .row = lngRowSave
  LockWindowUpdate 0
  m_booProcessSelected = False
End With

Exit Sub

cmdAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmCutUpEmbed", "cmdAdd_Click", intEL, strES

End Sub



Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub InitializeGrid()
      'Location Path column width changed from 0 to 1500 by Ibrahim 28-7-24
      'Location Path Column widht again changed to 0 by Ibrahim 29-7-4
      'Column name cut up by changed to cut by at two locations by ibrahim 29-7-24
10    With g
20      .Rows = 2
30      .FixedRows = 1
40      .Cols = 15
50      .FixedCols = 1
60      .Font.Size = fgcFontSize
70      .Font.Name = fgcFontName
80      .ForeColor = fgcForeColor
90      .BackColor = fgcBackColor
100     .ForeColorFixed = fgcForeColorFixed
110     .BackColorFixed = fgcBackColorFixed
120     .ScrollBars = flexScrollBarBoth
        
130     .TextMatrix(0, 0) = ""
140     .ColWidth(0) = 400
150     .ColAlignment(0) = flexAlignLeftCenter
        
160     .TextMatrix(0, 1) = "Caseid"
170     .ColWidth(1) = 1100
180     .ColAlignment(1) = flexAlignLeftCenter
        
190     .TextMatrix(0, 2) = "TissueID"
200     .ColWidth(2) = 0
210     .ColAlignment(2) = flexAlignLeftCenter
        
220     .TextMatrix(0, 3) = "Tissue Type"
230     .ColWidth(3) = 2700
240     .ColAlignment(3) = flexAlignLeftCenter
        
250     .TextMatrix(0, 4) = "Block"
260     .ColWidth(4) = 850
270     .ColAlignment(4) = flexAlignLeftCenter
        
280     .TextMatrix(0, 5) = "Pieces At Cut Up"
290     .ColWidth(5) = 1500
300     .ColAlignment(5) = flexAlignLeftCenter
        
310     .TextMatrix(0, 6) = "Orientation"
320     .ColWidth(6) = 1900
330     .ColAlignment(6) = flexAlignLeftCenter
        
340     .TextMatrix(0, 7) = "Pieces At Embedding"
350     .ColWidth(7) = 1600
360     .ColAlignment(7) = flexAlignLeftCenter
        
370     .TextMatrix(0, 8) = "Embedded By"
380     .ColWidth(8) = 1500
390     .ColAlignment(8) = flexAlignLeftCenter
        
400     .TextMatrix(0, 9) = "Cut-Up By"
410     .ColWidth(9) = 1500
420     .ColAlignment(9) = flexAlignLeftCenter
        
430     .TextMatrix(0, 10) = "Assisted By"
440     .ColWidth(10) = 1500
450     .ColAlignment(10) = flexAlignLeftCenter
        
460     .TextMatrix(0, 11) = "Cut-Up By"
470     .ColWidth(11) = 2000
480     .ColAlignment(11) = flexAlignLeftCenter
        
490     .TextMatrix(0, 12) = ""
500     .ColWidth(12) = 0
510     .ColAlignment(12) = flexAlignLeftCenter
        
520     .TextMatrix(0, 13) = "Processor"
530     .ColWidth(13) = 1000
540     .ColAlignment(13) = flexAlignLeftCenter
        
550     .TextMatrix(0, 14) = "AE"
560     .ColWidth(14) = 400
570     .ColAlignment(14) = flexAlignLeftCenter
        
580   End With

End Sub

Private Sub cmdPrint_Click()
    Const GAP = 60

    Dim xmax As Single
    Dim ymax As Single
    Dim xmin As Single
    Dim ymin As Single
    Dim X As Single
    Dim c As Integer


On Error GoTo cmdPrint_Click_Error

xmin = 1440
ymin = 1660

    Dim lRowsPrinted As Long, lRowsPerPage As Long
    Dim lThisRow As Long, lNumRows As Long
    Dim lPrinterPageHeight As Long
    Dim lPrintPage As Long
    Dim lNoOfPages As Long

g.TopRow = 1
lNumRows = g.Rows - 1
lPrinterPageHeight = Printer.Height

lRowsPrinted = 1



xmax = xmin + GAP
For c = 0 To g.Cols - 1
  If g.ColWidth(c) <> 0 Then
      xmax = xmax + g.ColWidth(c) + 2 * GAP
  End If
Next c

lPrintPage = 1

Do

  If UCase(pPhase) = "Cut-Up" Then
      Printer.Orientation = 2
      lRowsPerPage = 29
      lNoOfPages = Int(lNumRows / lRowsPerPage) + 1
  ElseIf UCase(pPhase) = "Embedding" Then
      Printer.Orientation = 2
      lRowsPerPage = 29
      lNoOfPages = Int(lNumRows / lRowsPerPage) + 1
  ElseIf UCase(pPhase) = "Cutting" Then
      Printer.Orientation = 1
      lRowsPerPage = 58
      lNoOfPages = Int(lNumRows / lRowsPerPage) + 1
  End If

  Do

      With g

          PrintHeadingWorkLog "Page " & lPrintPage & " of " & lNoOfPages, "Worklog : " & pPhase

          ' Print each row.
          Printer.CurrentY = ymin

          Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)

          Printer.CurrentY = Printer.CurrentY + GAP

          X = xmin + GAP
          For c = 0 To .Cols - 1
              Printer.CurrentX = X
              If .ColWidth(c) <> 0 Then
                  PrintText BoundedText(Printer, .TextMatrix(0, c), .ColWidth(c)), "MS Sans Serif", , True
                  X = X + .ColWidth(c) + 2 * GAP
              End If

          Next c
          Printer.CurrentY = Printer.CurrentY + GAP

          ' Move to the next line.
          PrintText vbCrLf

          For lThisRow = lRowsPrinted To lRowsPerPage * lPrintPage

              If (lThisRow - 1) < lNumRows Then
                  If lThisRow > 0 Then Printer.Line (xmin, Printer.CurrentY)-(xmax, Printer.CurrentY)
                  Printer.CurrentY = Printer.CurrentY + GAP

                  ' Print the entries on this row.
                  X = xmin + GAP
                  For c = 0 To .Cols - 1
                      Printer.CurrentX = X
                      If .ColWidth(c) <> 0 Then
                          PrintText BoundedText(Printer, .TextMatrix(lThisRow, c), .ColWidth(c)), "MS Sans Serif", 8
                          X = X + .ColWidth(c) + 2 * GAP
                      End If

                  Next c
                  Printer.CurrentY = Printer.CurrentY + GAP

                  ' Move to the next line.
                  PrintText vbCrLf

                  lRowsPrinted = lRowsPrinted + 1
              Else
                  Exit Do
              End If
          Next
      End With
  Loop While (lRowsPrinted - 1) < lRowsPerPage * lPrintPage

  ymax = Printer.CurrentY

  ' Draw a box around everything.
  Printer.Line (xmin, ymin)-(xmax, ymax), , B

  ' Draw lines between the columns.
  X = xmin
  For c = 0 To g.Cols - 2
      If g.ColWidth(c) <> 0 Then
          X = X + g.ColWidth(c) + 2 * GAP
          Printer.Line (X, ymin)-(X, ymax)
      End If
  Next c

  Printer.EndDoc
  lPrintPage = lPrintPage + 1

Loop While (lRowsPrinted - 1) < lNumRows

Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmCutUpEmbed", "cmdPrint_Click", intEL, strES


End Sub

Private Sub cmdReloadList_Click()

cmdSave.Visible = False
If pSingleEdit Then
  HideUnselected = False
Else
  HideUnselected = True
  fraData.Visible = False
  cmdAdd.Left = 4560
End If

ClearFields
ClearFGrid g
GridChanged = False
FillGrid
End Sub
Private Sub ClearFields()
cmbCutUpBy = ""
cmbAssistedBy = ""
cmbEmbeddedBy = ""
txtPiecesAfterEmbedding = ""
cmbOrientation = ""
txtPiecesAfterCutUp = ""
cmbCutBy = ""
cmbProcessor = ""
End Sub


Private Sub cmdSave_Click()
    Dim sql As String
    Dim tb As Recordset
    Dim r As Integer
    Dim SID As String

On Error GoTo cmdSave_Click_Error


r = 1
Do Until r = g.Rows
  If g.RowHeight(r) <> 0 Then

      SID = Replace(g.TextMatrix(r, 1), " " & sysOptCaseIdSeperator(0) & " ", "")
      sql = "SELECT * FROM BlockDetails B " & _
            "LEFT JOIN CaseTree CT ON B.Tissuelistid = CT.LocationSpecimenId " & _
            "WHERE B.CaseId = N'" & SID & "' " & _
            "AND B.TissueListId = N'" & g.TextMatrix(r, 2) & "' " & _
            "AND B.UniqueValue = N'" & Left(g.TextMatrix(r, 3), 1) & "' " & _
            "AND B.Block = N'" & g.TextMatrix(r, 4) & "' "
      Set tb = New Recordset
      RecOpenServer 0, tb, sql

      If Not tb.EOF Then

          tb!PiecesAfterCutUp = g.TextMatrix(r, 5)
          If g.TextMatrix(r, 5) <> "" Then
              CaseUpdateLogEvent SID, PiecesCutUp, g.TextMatrix(r, 5), g.TextMatrix(r, 12)
          End If

          tb!Orientation = g.TextMatrix(r, 6)

          tb!PiecesAfterEmbedding = g.TextMatrix(r, 7)    'Pieces after embedding
          If g.TextMatrix(r, 7) <> "" Then
              CaseUpdateLogEvent SID, PiecesEmbedding, g.TextMatrix(r, 7), g.TextMatrix(r, 12)
          End If

          tb!EmbeddedBy = g.TextMatrix(r, 8)    'Embedded By
          If g.TextMatrix(r, 8) <> "" Then
              CaseUpdateLogEvent SID, EmbeddedEvent, g.TextMatrix(r, 8), g.TextMatrix(r, 12)
          End If

          tb!CutUpBy = g.TextMatrix(r, 9)
          If g.TextMatrix(r, 9) <> "" Then    'Cut Up by
              CaseAddLogEvent SID, CutUpEvent, g.TextMatrix(r, 9), g.TextMatrix(r, 12)
          End If

          tb!AssistedBy = g.TextMatrix(r, 10)    'Assisted by
          If g.TextMatrix(r, 10) <> "" Then
              CaseUpdateLogEvent SID, AssistedBy, g.TextMatrix(r, 10), g.TextMatrix(r, 12)

          End If

          tb!CuttingBy = g.TextMatrix(r, 11)    'Cutting By
          If g.TextMatrix(r, 11) <> "" Then
              CaseAddLogEvent SID, CuttingBy, g.TextMatrix(r, 11), g.TextMatrix(r, 12)
          End If

          tb!Processor = g.TextMatrix(r, 13)    'Processor
          If g.TextMatrix(r, 13) <> "" Then
              CaseUpdateLogEvent SID, Processor, g.TextMatrix(r, 13), g.TextMatrix(r, 12)
          End If

          tb.Update
      End If

  End If
  r = r + 1
Loop

GridChanged = False

Exit Sub

cmdSave_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmCutUpEmbed", "cmdSave_Click", intEL, strES, sql

End Sub





Private Sub Form_Activate()
'frmCutUpEmbed_ChangeLanguage
End Sub

Private Sub Form_Load()

On Error GoTo Form_Load_Error
ChangeFont Me, "Arial"
Me.Caption = pPhase & " " & "Work logs"
LoadOrientaion
LoadProcessor
LoadUserLists
fraData.Visible = False
If pSingleEdit Then
  HideUnselected = False

  Me.Caption = pPhase & " " & "Work logs" & " - " & frmWorkSheet.txtCaseId
Else
  HideUnselected = True
End If
cmdAdd.Left = 4560
InitializeGrid
ClearFGrid g
GridChanged = False
With g
  .AllowUserResizing = flexResizeColumns
  .AllowBigSelection = False

  .FillStyle = flexFillRepeat
  .FocusRect = flexFocusNone

  .SelectionMode = flexSelectionFree
  If UCase(pPhase) = UCase("Cut-Up") Then
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      .ColWidth(11) = 0
      fraCutUp.Visible = True
      If pSingleEdit Then
          fraData.Visible = True
          cmdAdd.Left = 9000
      End If
  ElseIf UCase(pPhase) = UCase("Embedding") Then

      .ColWidth(9) = 0
      .ColWidth(10) = 0
      .ColWidth(11) = 0
      .ColWidth(13) = 0
      .ColWidth(14) = 0
      fraEmbed.Visible = True
      cmdPrint.Left = 8500
      cmdReloadList.Left = 9460
      cmdSave.Left = 10420
      cmdExit.Left = 11380
      g.Width = 11995
      Me.Width = 12400
  ElseIf UCase(pPhase) = UCase("Cutting") Then
      .ColWidth(5) = 0
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      .ColWidth(9) = 0
      .ColWidth(10) = 0
      .ColWidth(13) = 0
      .ColWidth(14) = 0
      fraCutting.Visible = True
      cmdPrint.Left = 6140
      cmdReloadList.Left = 7080
      cmdSave.Left = 8040
      cmdExit.Left = 9000
      g.Width = 9615
      Me.Width = 10020

  End If

End With
FillGrid
lblLoggedIn = UserName
If blnIsTestMode Then EnableTestMode Me

Exit Sub

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmCutUpEmbed", "Form_Load", intEL, strES

End Sub
Private Sub LoadOrientaion()
    Dim sql As String
    Dim tb As Recordset

On Error GoTo LoadOrientaion_Error

cmbOrientation.AddItem ""
sql = "SELECT * FROM Lists WHERE ListType = 'Orientation'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

Do While Not tb.EOF
  cmbOrientation.AddItem tb!Description & ""
  tb.MoveNext
Loop
cmbOrientation.ListIndex = -1

Exit Sub

LoadOrientaion_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmCutUpEmbed", "LoadOrientaion", intEL, strES, sql

End Sub

Private Sub LoadProcessor()
    Dim sql As String
    Dim tb As Recordset


On Error GoTo LoadProcessor_Error

cmbProcessor.AddItem ""
sql = "SELECT * FROM Lists WHERE ListType = 'Processor'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

Do While Not tb.EOF
  cmbProcessor.AddItem tb!Description & ""
  tb.MoveNext
Loop
cmbProcessor.ListIndex = -1



Exit Sub

LoadProcessor_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmCutUpEmbed", "LoadProcessor", intEL, strES, sql


End Sub

Private Sub LoadUserLists()
    Dim sql As String
    Dim tb As Recordset

On Error GoTo LoadUserLists_Error

cmbEmbeddedBy.AddItem ""
cmbCutUpBy.AddItem ""
cmbAssistedBy.AddItem ""
cmbCutBy.AddItem ""
sql = "SELECT * FROM USERS " & _
    "WHERE (AccessLevel = 'Scientist' " & _
    "OR AccessLevel = 'Manager' " & _
    "OR AccessLevel = 'Consultant') AND UserName <> 'Custom Software' "
Set tb = New Recordset
RecOpenServer 0, tb, sql

Do While Not tb.EOF
  cmbEmbeddedBy.AddItem tb!UserName & ""
  cmbCutUpBy.AddItem tb!UserName & ""
  cmbAssistedBy.AddItem tb!UserName & ""
  cmbCutBy.AddItem tb!UserName & ""
  tb.MoveNext
Loop
cmbEmbeddedBy.ListIndex = -1
cmbCutUpBy.ListIndex = -1
cmbAssistedBy.ListIndex = -1
cmbCutBy.ListIndex = -1

Exit Sub

LoadUserLists_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmCutUpEmbed", "LoadUserLists", intEL, strES, sql

End Sub

Private Sub FillGrid()
    Dim sql As String
    Dim tb As New Recordset
    Dim sn As Recordset
    Dim s As String
    Dim Block As String
    Dim TempCaseId As String


On Error GoTo FillGrid_Error

sql = "SELECT *, C.Phase FROM BlockDetails B " & _
    "INNER JOIN Cases C ON B.CaseId = C.CaseId " & _
    "WHERE C.Phase = N'" & pPhase & "' " & _
    "AND C.State = N'" & "In Histology" & "' "
If pSingleEdit = True Then

  TempCaseId = Replace(frmWorkSheet.tvCaseDetails.SelectedItem, " " & sysOptCaseIdSeperator(0) & " ", "")

  sql = sql & "AND B.CaseId = N'" & TempCaseId & "' "

End If
sql = sql & "ORDER BY B.CaseId, B.UniqueValue, B.TissueListid, B.BlockNumber"

Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF

  If Mid(tb!CaseId & "", 2, 1) = "P" Or Mid(tb!CaseId & "", 2, 1) = "A" Then
      TempCaseId = Left(tb!CaseId, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
  Else
      TempCaseId = Left(tb!CaseId, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(tb!CaseId, 2)
  End If

  sql = "SELECT * FROM CaseTree WHERE LocationID = N'" & tb!Tissuelistid & "' "
  Set sn = New Recordset
  RecOpenServer 0, sn, sql

  If tb!Block & "" <> "" Then
      s = vbTab & TempCaseId & vbTab & tb!Tissuelistid & vbTab _
        & sn!LocationName & vbTab & tb!Block _
        & vbTab & tb!PiecesAfterCutUp & "" & vbTab & tb!Orientation & "" & vbTab _
        & tb!PiecesAfterEmbedding & "" & vbTab & tb!EmbeddedBy & "" & vbTab _
        & tb!CutUpBy & "" & vbTab & tb!AssistedBy & "" & vbTab & tb!CuttingBy & "" _
        & vbTab & tb!LocationPath & "" & vbTab & tb!Processor & ""

      g.AddItem s
      g.row = g.Rows - 1
  End If

  tb.MoveNext
Loop

g.Visible = True
If g.Rows > 2 Then
  g.RemoveItem 1
End If

Exit Sub

FillGrid_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmCutUpEmbed", "FillGrid", intEL, strES, sql

End Sub

Private Sub Form_Unload(Cancel As Integer)
If TimedOut Then Unload Me: Exit Sub
If GridChanged = True Then
  If frmMsgBox.Msg("Alert!! Do you want to save your changes?", mbYesNo, , mbQuestion) = 1 Then
      cmdSave_Click
  End If
End If
End Sub

Private Sub g_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift = vbShiftMask _
 Then
  m_booKeyShift = True
End If
If Shift = vbCtrlMask _
 Then
  m_booKeyCtrl = True
End If

End Sub

Private Sub g_Keyup(KeyCode As Integer, Shift As Integer)
With g
  If KeyCode = 38 Then
      m_UpKey = True
      m_DownKey = False
  ElseIf KeyCode = 40 Then
      m_DownKey = True
      m_UpKey = False
  Else
      m_booKeyCtrl = False
      m_booKeyShift = False
  End If

End With
End Sub

Private Sub g_RowColChange()

With g
  Static booBusy As Boolean
  If m_booKeyShift _
     Or m_booProcessSelected _
     Or booBusy _
     Then
      Exit Sub
  End If
  booBusy = True

  Dim intThisCol As Integer
  Dim intThisRow As Integer
  intThisCol = .col
  intThisRow = .row

  LockWindowUpdate .hwnd

  If m_booKeyCtrl _
     Then
      .col = 1
      .row = intThisRow
      .ColSel = .Cols - 1
      .RowSel = intThisRow
      If .CellBackColor = .BackColorSel _
         Then
          .CellBackColor = .BackColor
          .CellForeColor = .ForeColor
      Else
          .CellBackColor = .BackColorSel
          .CellForeColor = .ForeColorSel
      End If
  Else
      .col = 1
      .row = 1
      .ColSel = .Cols - 1
      .RowSel = .Rows - 1
      .FillStyle = flexFillRepeat
      .CellBackColor = .BackColor
      .CellForeColor = .ForeColor
      .col = 1
      .row = intThisRow
      .ColSel = .Cols - 1
      .RowSel = intThisRow
      .CellBackColor = .BackColorSel
      .CellForeColor = .ForeColorSel
  End If

  .col = intThisCol
  .row = intThisRow

  LockWindowUpdate 0&
  booBusy = False

End With
End Sub

Private Sub g_SelChange()
With g

  If Not m_booKeyShift _
     Then
      Exit Sub
  End If

  Dim intThisCol As Integer
  Dim intThisRow As Integer
  intThisCol = .col
  intThisRow = .row

  Dim intNextCol As Integer
  Dim intNextRow As Integer
  intNextCol = .ColSel
  intNextRow = .RowSel

  LockWindowUpdate .hwnd

  ' Clear Screen
  .col = 1
  .row = intNextRow
  .ColSel = .Cols - 1
  .RowSel = intThisRow
  .FillStyle = flexFillRepeat
  If .CellBackColor = .BackColorSel _
     Then

      .CellBackColor = .BackColor
      .CellForeColor = .ForeColor
  Else

      .CellBackColor = .BackColorSel
      .CellForeColor = .ForeColorSel
  End If
  LockWindowUpdate 0&

Tag900:

  .col = intNextCol
  .row = intNextRow

End With
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then

  Me.Top = 0
  Me.Left = Screen.Width / 2 - Me.Width / 2
End If
End Sub




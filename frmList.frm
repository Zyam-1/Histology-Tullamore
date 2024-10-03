VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmList 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   3960
   ClientTop       =   3555
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   2500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   4419
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   0
      RowHeight       =   15
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Description"
         Caption         =   "Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******Set these variable before calling this form********
Public ListType As String
Public ListId As Integer
Public External As Boolean
Public SearchByCode As Boolean
Public txtCode As TextBox
Public txtDescription As TextBox
Public tempListId As Label
Public grdTemp As MSFlexGrid
'***************************************
Private CodeOrDesc As String
Private SearchString As String
Private pPrevCode As String
Private pPrevDesc As String
'Private Ado As New AdoDB.Connection
Private AdoRs As Recordset

Public Property Let PrevCode(ByVal Value As String)

10    pPrevCode = Value

End Property

Public Property Let PrevDesc(ByVal Value As String)

10    pPrevDesc = Value

End Property


Private Sub Form_Activate()
10    If SearchByCode Then
20        CodeOrDesc = "Code"
30        SearchString = txtCode
40    Else
50        CodeOrDesc = "Description"
60        SearchString = txtDescription
70    End If


80    GridRequery "Select ListId, Code, Description, [External] From Lists " & _
                          "Where " & CodeOrDesc & " Like N'%" & SearchString & "%' " & _
                          "And ListType = '" & ListType & "' AND InUse = 1 " & _
                          "Order By Rank Desc"

90    SetListSize
End Sub

Private Sub Form_Deactivate()
10    txtCode = pPrevCode
20    txtDescription = pPrevDesc
30    Unload Me
End Sub

Private Sub GridRequery(Optional sql As String = "")

10    On Error GoTo GridRequery_Error

20    If sql = "" Then
30        sql = "Select listid, Code, Description, [External] From Lists WHERE InUse = 1 order by rank desc"
40    End If

50    Set AdoRs = New Recordset

60    RecOpenClient 0, AdoRs, sql

70    Set Grid.DataSource = AdoRs
80    Grid.Refresh
90    Grid.MarqueeStyle = dbgHighlightRow


100   Exit Sub

GridRequery_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmList", "GridRequery", intEL, strES, sql


End Sub

Private Sub Form_Load()

10    GridRequery

End Sub

Private Sub Form_Unload(Cancel As Integer)
10    ListType = ""
20    SearchByCode = False
30    Set grdTemp = Nothing
End Sub


Private Sub Grid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
10    Cancel = True
End Sub

Private Sub Grid_DblClick()
      On Error Resume Next
10    If AdoRs.RecordCount = 0 Then Exit Sub
20            ListId = AdoRs.Fields("ListId")
30            If Not tempListId Is Nothing Then
40                tempListIds = AdoRs.Fields("ListId")
50            End If
60            External = AdoRs.Fields("External")
70            txtCode = AdoRs.Fields("Code")
80            txtDescription = AdoRs.Fields("Description")

  
90            If Not grdTemp Is Nothing Then
100               For i = 1 To grdTemp.Rows - 1
110                   If UCase(grdTemp.TextMatrix(i, 0)) = UCase(txtCode) Then
120                       MsgBox "Item already exists in the list, Please choose different item", vbInformation
130                       txtCode.Visible = False
140                       txtDescription.Visible = False
150                       Exit Sub
160                   End If
170               Next i
180               grdTemp.TextMatrix(grdTemp.row, 0) = txtCode
190               grdTemp.TextMatrix(grdTemp.row, 1) = txtDescription
200               txtCode.Visible = False
210               txtDescription.Visible = False
220           Else
230               If SearchByCode Then
240                   txtCode.SetFocus
250               Else
260                   txtDescription.SetFocus
270               End If
280           End If
290           Unload Me
End Sub

Private Sub Grid_Error(ByVal DataError As Integer, Response As Integer)
10    Response = vbDataErrContinue
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
10    If KeyCode = vbKeyEscape Then

20        txtCode = pPrevCode
30        txtDescription = pPrevDesc

40        KeyCode = 0
50        Unload Me
60    End If

End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)

          Dim i As Integer

10        If KeyAscii = 13 Then
20            If AdoRs.RecordCount = 0 Then Exit Sub
30            ListId = AdoRs.Fields("ListId")
40            If Not tempListId Is Nothing Then
50                tempListId = AdoRs.Fields("ListId")
60            End If
70            External = AdoRs.Fields("External")
80            txtCode = AdoRs.Fields("Code")
90            txtDescription = AdoRs.Fields("Description")

100           If Not grdTemp Is Nothing Then
110               For i = 1 To grdTemp.Rows - 1
120                   If UCase(grdTemp.TextMatrix(i, 0)) = UCase(txtCode) Then
130                       MsgBox "Item already exists in the list, Please choose different item", vbInformation
140                       txtCode.Visible = False
150                       txtDescription.Visible = False
160                       Exit Sub
170                   End If
180               Next i
190               grdTemp.TextMatrix(grdTemp.row, 0) = txtCode
200               grdTemp.TextMatrix(grdTemp.row, 1) = txtDescription
210               txtCode.Visible = False
220               txtDescription.Visible = False
230           Else
240               If SearchByCode Then
250                   txtCode.SetFocus
260               Else
270                   txtDescription.SetFocus
280               End If
290           End If
300           KeyAscii = 0
310           Unload Me


320       Else
330           If KeyAscii = 27 Then
340               Exit Sub
350           ElseIf KeyAscii = 8 Then
360               If SearchByCode Then
370                   If Len(txtCode) > 0 Then
380                       SearchString = Mid(txtCode, 1, Len(txtCode) - 1)
390                       txtCode = SearchString
400                       GridRequery "Select ListId, Code, Description, [External] From Lists " & _
                                             "Where " & CodeOrDesc & " Like '%" & SearchString & "%' " & _
                                             "And ListType = '" & ListType & "' AND InUse = 1 " & _
                                             "Order By Rank Desc"
410                       Grid.SetFocus
420                   End If
430               Else
440                   If Len(txtDescription) > 0 Then
450                       SearchString = Mid(txtDescription, 1, Len(txtDescription) - 1)
460                       txtDescription = SearchString
470                       GridRequery "Select ListId, Code, Description, [External] From Lists " & _
                                             "Where " & CodeOrDesc & " Like '%" & SearchString & "%' " & _
                                             "And ListType = '" & ListType & "' AND InUse = 1 " & _
                                             "Order By Rank Desc"
480                       Grid.SetFocus
490                   End If
500               End If
510           Else
520               If SearchByCode Then
530                   SearchString = Me.txtCode & Chr(KeyAscii)
540                   txtCode = SearchString
550               Else
560                   SearchString = Me.txtDescription & Chr(KeyAscii)
570                   txtDescription = SearchString
580               End If
590               GridRequery "Select ListId, Code, Description, External From Lists " & _
                                     "Where " & CodeOrDesc & " Like '%" & SearchString & "%' " & _
                                     "And ListType = '" & ListType & "' AND InUse = 1 " & _
                                     "Order By Rank Desc"
600               Grid.SetFocus
610           End If
620           SetListSize
630           KeyAscii = 0
640       End If

End Sub


Private Sub SetListSize()
      Dim RecordHeight As Long

10    On Error GoTo SetListSize_Error

20    RecordHeight = (AdoRs.RecordCount * 225) + 50
30    Grid.Height = (Grid.RowHeight * 10) + 50

40    If Grid.Height > RecordHeight Then
50        Grid.Height = RecordHeight
60    End If
70    Me.Height = Grid.Height
80    Grid.Width = txtCode.Width + txtDescription.Width
90    Me.Width = Grid.Width



100   Grid.Columns(0).Width = 1200

110   If RecordHeight > (Grid.RowHeight * 10) Then
120           Grid.Columns(1).Width = Grid.Width - 1435
130   Else
140           Grid.Columns(1).Width = Grid.Width - 1250
150   End If

160   Exit Sub

SetListSize_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmList", "SetListSize", intEL, strES


End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAuthorisedReports 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid gData 
      Height          =   4680
      Left            =   180
      TabIndex        =   14
      Top             =   4995
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8255
      _Version        =   393216
      Cols            =   5
      FormatString    =   "Consultant  |    |    |Case ID|Authorise Date"
   End
   Begin Threed.SSPanel panSampleDates 
      Height          =   4500
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   10995
      _Version        =   65536
      _ExtentX        =   19394
      _ExtentY        =   7937
      _StockProps     =   15
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin VB.ComboBox cmbSource 
         Height          =   315
         Left            =   1080
         TabIndex        =   23
         Top             =   1200
         Width           =   4215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   900
         Left            =   9855
         Picture         =   "frmAuthorisedReports.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3300
         Width           =   1000
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Export to Excel"
         Height          =   900
         Left            =   8655
         Picture         =   "frmAuthorisedReports.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3300
         Width           =   1000
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   900
         Left            =   7380
         Picture         =   "frmAuthorisedReports.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3300
         Width           =   1000
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   900
         Left            =   8370
         Picture         =   "frmAuthorisedReports.frx":08A7
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   690
         Width           =   1000
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6390
         TabIndex        =   11
         Top             =   1845
         Width           =   495
      End
      Begin VB.ComboBox cmbTCode 
         Height          =   315
         ItemData        =   "frmAuthorisedReports.frx":0C25
         Left            =   3240
         List            =   "frmAuthorisedReports.frx":0C27
         TabIndex        =   10
         Top             =   1905
         Width           =   1000
      End
      Begin VB.ComboBox cmbQCode 
         Height          =   315
         Left            =   5250
         TabIndex        =   9
         Top             =   1905
         Width           =   1000
      End
      Begin VB.ComboBox cmbMCode 
         Height          =   315
         Left            =   4260
         TabIndex        =   8
         Top             =   1905
         Width           =   1000
      End
      Begin VB.ComboBox cmbPCode 
         Height          =   315
         ItemData        =   "frmAuthorisedReports.frx":0C29
         Left            =   2235
         List            =   "frmAuthorisedReports.frx":0C2B
         TabIndex        =   7
         Top             =   1905
         Width           =   1000
      End
      Begin VB.ComboBox cmbConsultant 
         Height          =   315
         ItemData        =   "frmAuthorisedReports.frx":0C2D
         Left            =   240
         List            =   "frmAuthorisedReports.frx":0C2F
         TabIndex        =   6
         Top             =   1905
         Width           =   1995
      End
      Begin MSComCtl2.DTPicker calfrom 
         Height          =   345
         Left            =   210
         TabIndex        =   1
         Top             =   720
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   251789313
         CurrentDate     =   37753
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   345
         Left            =   2910
         TabIndex        =   2
         Top             =   720
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   251789313
         CurrentDate     =   37753
      End
      Begin MSFlexGridLib.MSFlexGrid gFilter 
         Height          =   1995
         Left            =   210
         TabIndex        =   5
         Top             =   2205
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   3519
         _Version        =   393216
         Cols            =   6
         FormatString    =   "Consultant             |P Code    |T Code    |M Code    |Q Code    |    "
      End
      Begin VB.Label Label 
         Caption         =   "Source"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Advanced Filter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   13
         Top             =   1530
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Filter By Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   660
         TabIndex        =   12
         Top             =   180
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   450
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2895
         TabIndex        =   3
         Top             =   450
         Width           =   945
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   285
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gDetail 
      Height          =   3900
      Left            =   6420
      TabIndex        =   19
      Top             =   5760
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6879
      _Version        =   393216
      FormatString    =   "Code|Description"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Click on Case ID to get details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6967
      TabIndex        =   20
      Top             =   5400
      Width           =   3660
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmAuthorisedReports.frx":0C31
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmAuthorisedReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit









Private Sub cmbMCode_Click()
      Dim intP As Integer
10    If Len(cmbMCode) > 0 Then
20        intP = InStr(cmbMCode, " - ")
30        If intP > 0 Then
40            cmbMCode = Left(cmbMCode, intP - 1)
50        End If
60    End If
End Sub

Private Sub cmbPCode_Click()
      Dim intP As Integer
10    If Len(cmbPCode) > 0 Then
20        intP = InStr(cmbPCode, " - ")
30        If intP > 0 Then
40            cmbPCode = Left(cmbPCode, intP - 1)
50        End If
60    End If
End Sub

Private Sub cmbQCode_Click()
      Dim intP As Integer

10    If Len(cmbQCode) > 0 Then
20        intP = InStr(cmbQCode, " - ")
30        If intP > 0 Then
40            cmbQCode = Left(cmbQCode, intP - 1)
50        End If
60    End If
End Sub

Private Sub cmbTCode_Click()
      Dim intP As Integer
10    If Len(cmbTCode) > 0 Then
20        intP = InStr(cmbTCode, " - ")
30        If intP > 0 Then
40            cmbTCode = Left(cmbTCode, intP - 1)
50        End If
60    End If
End Sub

Private Sub cmdAdd_Click()

      Dim s As String
      Dim i As Integer

10    On Error GoTo cmdAdd_Click_Error

20    If cmbConsultant = "" And cmbPCode = "" And cmbTCode = "" And cmbMCode = "" And cmbQCode = "" Then
30        iMsg "Please select criteria first", vbInformation
40        Exit Sub
50    End If

60    With gFilter
70        For i = 1 To gFilter.Rows - 1
80            If UCase$(cmbConsultant & cmbPCode & cmbTCode & cmbMCode & cmbQCode) = _
                      UCase$(.TextMatrix(i, 0) & .TextMatrix(i, 1) & .TextMatrix(i, 2) & .TextMatrix(i, 3) & .TextMatrix(i, 4)) Then
90                iMsg "Criteria already exists", vbInformation
100               Exit Sub
110           End If
120       Next i
130   End With




      'every thing is ok. add criteria here

140   With gFilter

150       s = cmbConsultant & vbTab & _
                  cmbPCode & vbTab & _
                  cmbTCode & vbTab & _
                  cmbMCode & vbTab & _
                  cmbQCode
160       .AddItem s
170       .row = .Rows - 1
180       .col = 5
190       Set .CellPicture = imgSquareCross

200   End With

210   cmbConsultant.ListIndex = -1
220   cmbPCode = ""
230   cmbTCode = ""
240   cmbMCode = ""
250   cmbQCode = ""


260   Exit Sub

cmdAdd_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmAuthorisedReports", "cmdAdd_Click", intEL, strES

End Sub

Private Sub cmdExcel_Click()

      Dim Title As String

10    On Error GoTo cmdExcel_Click_Error

20    If gData.Rows = 1 Then
30        iMsg "Nothing to export", vbInformation
40        Exit Sub
50    End If

60    Title = "Cellular Pathology - Authorised Cases" & vbCr
70    Title = Title & "From " & calFrom.Value & " To " & calTo & vbCr
80    ExportFlexGrid gData, Me, Title


90    Exit Sub

cmdExcel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmAuthorisedReports", "cmdExcel_Click", intEL, strES

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

      Dim i As Integer

10    On Error GoTo cmdPrint_Click_Error

20    If gData.Rows = 1 Then
30        iMsg "Nothing to print", vbInformation
40        Exit Sub
50    End If

60    PrintText vbCr
70    PrintText FormatString("Cellular Pathology - Authorised Cases", 100, , AlignCenter) & vbCr, , 10, True
80    PrintText FormatString("From " & calFrom.Value & " To " & calTo, 100, , AlignCenter) & vbCr, , 10, True
90    PrintText vbCr

      'PrintText String(105, "-") & vbCr
100   PrintText String(241, "-") & vbCr, , 4, True
110   PrintText FormatString("Consultant", 20, "|", AlignCenter), , , True
120   PrintText FormatString("Case ID", 15, "|", AlignCenter), , , True
130   PrintText FormatString("Authorised Date", 16, "|", AlignCenter) & vbCr, , , True
140   PrintText String(241, "-") & vbCr, , 4, True

150   For i = 1 To gData.Rows - 1
160       If InStr(1, gData.TextMatrix(i, 0), "Total") > 0 Then
170           PrintText String(40, "-") & vbCr, , 4, True
180           PrintText "Total Cases " & gData.TextMatrix(i, 1) & vbCr, , , True
190           PrintText String(40, "-") & vbCr, , 4, True
200       Else
210           PrintText FormatString(gData.TextMatrix(i, 0), 20, "|", Alignleft)
220           PrintText FormatString(gData.TextMatrix(i, 3), 15, "|", Alignleft)
230           PrintText FormatString(gData.TextMatrix(i, 4), 16, "|", Alignleft) & vbCr
240       End If
    
    
    
    
250   Next i
    


260   Printer.EndDoc

      'With frmRichText
      '    PrintTextRTB .rtb, FormatString("Cellular Pathology - Authorised Cases", 100, , AlignCenter) & vbCr
      '    PrintTextRTB .rtb, FormatString("From " & calfrom.Value & " To " & calTo, 100, , AlignCenter) & vbCr
      '    PrintTextRTB .rtb, vbCr
      '
      '    PrintTextRTB .rtb, FormatString("Consultant", 25, "|", AlignCenter)
      '    PrintTextRTB .rtb, FormatString("Case ID", 15, "|", AlignCenter)
      '    PrintTextRTB .rtb, FormatString("P Code", 10, "|", AlignCenter)
      '    PrintTextRTB .rtb, FormatString("Description", 32, "|", AlignCenter)
      '    PrintTextRTB .rtb, FormatString("Authorised Date", 18, "|", AlignCenter) & vbCr
      '
      '    PrintTextRTB .rtb, String(100, "-") & vbCr
      '
      '    For i = 1 To gData.Rows - 2
      '        PrintTextRTB .rtb, FormatString(gData.TextMatrix(i, 0), 25, "|", AlignLeft)
      '        PrintTextRTB .rtb, FormatString(gData.TextMatrix(i, 3), 15, "|", AlignLeft)
      '        PrintTextRTB .rtb, FormatString(gData.TextMatrix(i, 4), 10, "|", AlignLeft)
      '        PrintTextRTB .rtb, FormatString(gData.TextMatrix(i, 5), 32, "|", AlignLeft)
      '        PrintTextRTB .rtb, FormatString(gData.TextMatrix(i, 6), 18, "|", AlignLeft) & vbCr
      '    Next i
      '
      '    PrintTextRTB .rtb, String(100, "-") & vbCr
      '    PrintTextRTB .rtb, "Total Caases " & gData.TextMatrix(gData.Rows - 1, 1)
      '
      '    .Show 1
      'End With

270   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmAuthorisedReports", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdSearch_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim subSql As String
      Dim CodeList As String
      Dim i As Integer
      Dim OldConsultant As String
      Dim NewConsultant As String
      Dim CaseCount As Integer


10    On Error GoTo cmdSearch_Click_Error



20    subSql = ""
30    InitializeGridgData

40    If Trim(cmbSource.Text) = "" Then

50        If gFilter.Rows = 1 Then
60            sql = "SELECT DISTINCT C.OrigValBy, C.OrigValDate, C.CaseId FROM CaseListLink CLL " & _
                      "LEFT JOIN Lists L on CAST(CLL.ListId   as nvarchar(50))  = L.ListId " & _
                      "INNER JOIN Cases C on CLL.CaseId = C.CaseId " & _
                      "WHERE C.OrigValDate BETWEEN '" & Format(calFrom.Value, "MM/dd/yyyy 00:00:01") & "' AND '" & Format(calTo, "MM/dd/yyyy 23:59:59") & "' "
70        Else
80            With gFilter
90                For i = 1 To .Rows - 1
    
100                   CodeList = ""
110                   If i > 1 Then sql = sql & " Union "
    
120                   subSql = "SELECT DISTINCT C.OrigValBy, C.OrigValDate, C.CaseId FROM CaseListLink CLL " & _
                              "LEFT JOIN Lists L on CAST(CLL.ListId   as nvarchar(50))  = L.ListId " & _
                              "INNER JOIN Cases C on CLL.CaseId = C.CaseId " & _
                              "WHERE C.OrigValDate BETWEEN '" & Format(calFrom.Value, "MM/dd/yyyy 00:00:01") & "' AND '" & Format(calTo, "MM/dd/yyyy 23:59:59") & "' "
    
    
130                   If .TextMatrix(i, 0) <> "" Then
140                       subSql = subSql & "AND C.OrigValBy = '" & AddTicks(.TextMatrix(i, 0)) & "' "
150                   End If
    
                      'create code list
                      'PCode
    
160                   If .TextMatrix(i, 1) <> "" Then
170                       subSql = subSql & "AND EXISTS(SELECT 1 FROM CaseListLink INNER JOIN Lists ON CaseListLink.ListId = Lists.ListId " & _
                                  "WHERE Lists.Code = '" & .TextMatrix(i, 1) & "' AND CaseId = C.CaseId AND CaseListLink.Type = 'P') "
          '                subSql = subSql & "AND CLL.Type = 'P' "
180                       CodeList = CodeList & "'" & .TextMatrix(i, 1) & "'"
        
190                   End If
                      'T Code
200                   If .TextMatrix(i, 2) <> "" Then
210                       subSql = subSql & "AND EXISTS(SELECT 1 FROM CaseListLink INNER JOIN Lists ON CaseListLink.TissueTypeListId = Lists.ListId " & _
                                  "WHERE Lists.Code = '" & .TextMatrix(i, 2) & "' AND CaseId = C.CaseId ) "
220                       If CodeList <> "" Then CodeList = CodeList & ","
230                       CodeList = CodeList & "'" & .TextMatrix(i, 2) & "'"
240                   End If
                      'M Code
250                   If .TextMatrix(i, 3) <> "" Then
260                       subSql = subSql & "AND EXISTS(SELECT 1 FROM CaseListLink INNER JOIN Lists ON CaseListLink.ListId = Lists.ListId " & _
                                  "WHERE Lists.Code = '" & .TextMatrix(i, 3) & "' AND CaseId = C.CaseId AND CaseListLink.Type = 'M') "
          '                subSql = subSql & "AND CLL.Type = 'M' "
270                       If CodeList <> "" Then CodeList = CodeList & ","
280                       CodeList = CodeList & "'" & .TextMatrix(i, 3) & "'"
    
290                   End If
                      'Q Code
300                   If .TextMatrix(i, 4) <> "" Then
310                       subSql = subSql & "AND EXISTS(SELECT 1 FROM CaseListLink INNER JOIN Lists ON CaseListLink.ListId = Lists.ListId " & _
                                  "WHERE Lists.Code = '" & .TextMatrix(i, 4) & "' AND CaseId = C.CaseId AND CaseListLink.Type = 'Q') "
          '                subSql = subSql & "AND CLL.Type = 'Q' "
320                       If CodeList <> "" Then CodeList = CodeList & ","
330                       CodeList = CodeList & "'" & .TextMatrix(i, 4) & "'"
340                   End If
    
350                   sql = sql & subSql
360                   If CodeList <> "" Then
370                       sql = sql & "AND L.Code IN (" & CodeList & ") "
380                   End If
    
390               Next i
400           End With
410       End If
420     Else
430       If gFilter.Rows = 1 Then
440           sql = "SELECT DISTINCT C.OrigValBy, C.OrigValDate, C.CaseId FROM CaseListLink CLL " & _
                      "LEFT JOIN Lists L on CAST(CLL.ListId   as nvarchar(50))  = L.ListId " & _
                      "INNER JOIN Cases C on CLL.CaseId = C.CaseId INNER JOIN Demographics D ON D.CaseID = CLL.CaseID " & _
                      "WHERE D.Source = '" & Trim(cmbSource.Text) & "' AND C.OrigValDate BETWEEN '" & Format(calFrom.Value, "MM/dd/yyyy 00:00:01") & "' AND '" & Format(calTo, "MM/dd/yyyy 23:59:59") & "' "
450       Else
460           With gFilter
470               For i = 1 To .Rows - 1
    
480                   CodeList = ""
490                   If i > 1 Then sql = sql & " Union "
    
500                   subSql = "SELECT DISTINCT C.OrigValBy, C.OrigValDate, C.CaseId FROM CaseListLink CLL " & _
                              "LEFT JOIN Lists L on CAST(CLL.ListId   as nvarchar(50))  = L.ListId " & _
                              "INNER JOIN Cases C on CLL.CaseId = C.CaseId INNER JOIN Demographics D ON D.CaseID = CLL.CaseID " & _
                              "WHERE D.Source = '" & Trim(cmbSource.Text) & "' AND C.OrigValDate BETWEEN '" & Format(calFrom.Value, "MM/dd/yyyy 00:00:01") & "' AND '" & Format(calTo, "MM/dd/yyyy 23:59:59") & "' "
    
    
510                   If .TextMatrix(i, 0) <> "" Then
520                       subSql = subSql & "AND C.OrigValBy = '" & AddTicks(.TextMatrix(i, 0)) & "' "
530                   End If
    
                      'create code list
                      'PCode
    
540                   If .TextMatrix(i, 1) <> "" Then
550                       subSql = subSql & "AND EXISTS(SELECT 1 FROM CaseListLink INNER JOIN Lists ON CaseListLink.ListId = Lists.ListId " & _
                                  "WHERE Lists.Code = '" & .TextMatrix(i, 1) & "' AND CaseId = C.CaseId AND CaseListLink.Type = 'P') "
          '                subSql = subSql & "AND CLL.Type = 'P' "
560                       CodeList = CodeList & "'" & .TextMatrix(i, 1) & "'"
        
570                   End If
                      'T Code
580                   If .TextMatrix(i, 2) <> "" Then
590                       subSql = subSql & "AND EXISTS(SELECT 1 FROM CaseListLink INNER JOIN Lists ON CaseListLink.TissueTypeListId = Lists.ListId " & _
                                  "WHERE Lists.Code = '" & .TextMatrix(i, 2) & "' AND CaseId = C.CaseId ) "
600                       If CodeList <> "" Then CodeList = CodeList & ","
610                       CodeList = CodeList & "'" & .TextMatrix(i, 2) & "'"
620                   End If
                      'M Code
630                   If .TextMatrix(i, 3) <> "" Then
640                       subSql = subSql & "AND EXISTS(SELECT 1 FROM CaseListLink INNER JOIN Lists ON CaseListLink.ListId = Lists.ListId " & _
                                  "WHERE Lists.Code = '" & .TextMatrix(i, 3) & "' AND CaseId = C.CaseId AND CaseListLink.Type = 'M') "
          '                subSql = subSql & "AND CLL.Type = 'M' "
650                       If CodeList <> "" Then CodeList = CodeList & ","
660                       CodeList = CodeList & "'" & .TextMatrix(i, 3) & "'"
    
670                   End If
                      'Q Code
680                   If .TextMatrix(i, 4) <> "" Then
690                       subSql = subSql & "AND EXISTS(SELECT 1 FROM CaseListLink INNER JOIN Lists ON CaseListLink.ListId = Lists.ListId " & _
                                  "WHERE Lists.Code = '" & .TextMatrix(i, 4) & "' AND CaseId = C.CaseId AND CaseListLink.Type = 'Q') "
          '                subSql = subSql & "AND CLL.Type = 'Q' "
700                       If CodeList <> "" Then CodeList = CodeList & ","
710                       CodeList = CodeList & "'" & .TextMatrix(i, 4) & "'"
720                   End If
    
730                   sql = sql & subSql
740                   If CodeList <> "" Then
750                       sql = sql & "AND L.Code IN (" & CodeList & ") "
760                   End If
    
770               Next i
780           End With
            
        End If
            
790     End If

800   sql = sql & "ORDER BY C.OrigValBy"
      'sql is created, time to fetch some results
810   Set tb = New Recordset
820   RecOpenClient 0, tb, sql
830   If Not tb.EOF Then

840       gData.Visible = False

850       NewConsultant = tb!OrigValBy & ""
860       While Not tb.EOF
870           OldConsultant = NewConsultant
880           NewConsultant = tb!OrigValBy & ""
890           If UCase(OldConsultant) = UCase(NewConsultant) Then
900               CaseCount = CaseCount + 1
910           Else
920               gData.AddItem "Total Cases" & vbTab & CaseCount
930               MarkGridRow gData, gData.Rows - 1, &HCCCCCC, , True
940               CaseCount = 1

950           End If
960           gData.AddItem tb!OrigValBy & "" & vbTab & _
                      vbTab & vbTab & _
                      tb!CaseID & "" & vbTab & _
                      tb!OrigValDate & ""

970           gData.row = gData.Rows - 1
980           gData.col = 2
990           gData.CellBackColor = vbRed

              'Add count to consultant cases
              '        If UCase(OldConsultant) <> UCase(NewConsultant) Then
              '            gData.AddItem "Total Cases" & vbTab & CaseCount
              '        End If
1000          tb.MoveNext

1010      Wend
1020      gData.AddItem "Total Cases" & vbTab & CaseCount
1030      MarkGridRow gData, gData.Rows - 1, &HCCCCCC, , True

1040      gData.Visible = True
1050  End If


1060  Exit Sub

cmdSearch_Click_Error:

1070  gData.Visible = True
      Dim strES As String
      Dim intEL As Integer

1080  intEL = Erl
1090  strES = Err.Description
1100  LogError "frmAuthorisedReports", "cmdSearch_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()
'frmAuthorisedReports_ChangeLanguage

10    calFrom = Date - 7
20    calTo = Date
30    InitializeGridgFilter
40    InitializeGridgData
50    InitializeGridgDetail
60    FillPathologistList
70    PopulateGenericList_2Cols "P", cmbPCode, "Code"
80    PopulateGenericList_2Cols "T", cmbTCode, "Code"
90    PopulateGenericList_2Cols "M", cmbMCode, "Code"
100   PopulateGenericList_2Cols "Q", cmbQCode, "Code"
110   FillSource

120   FixComboWidth cmbConsultant
130   FixComboWidth cmbPCode
140   FixComboWidth cmbTCode
150   FixComboWidth cmbMCode
160   FixComboWidth cmbQCode

170   If blnIsTestMode Then EnableTestMode Me

End Sub


Private Sub FillSource()
          Dim sql As String
          Dim tb As Recordset
         
10       On Error GoTo FillSource_Error
20       sql = "SELECT Description FROM Lists WHERE ListType = 'Source' AND InUse = 1"
         
30       Set tb = New Recordset
40       RecOpenServer 0, tb, sql
50       cmbSource.Clear
60       cmbSource.AddItem ""
70       If Not tb Is Nothing Then
80          Do While Not tb.EOF
90              cmbSource.AddItem tb!Description & ""
                tb.MoveNext
100         Loop
         
110      End If
FillSource_Error:
             Dim strES As String
             Dim intEL As Long

120          intEL = Erl
130          strES = Err.Description
140          LogError "frmAuthorisedReport", "fillSource", intEL, strES

          
End Sub






Private Sub InitializeGridgFilter()

10    With gFilter
20        .Rows = 2: .FixedRows = 1
30        .Cols = 6: .FixedCols = 0
40        .Rows = 1
50        .Font.Size = fgcFontSize
60        .Font.Name = fgcFontName

70        .ScrollBars = flexScrollBarBoth

80        .TextMatrix(0, 0) = "Consultant": .ColWidth(0) = 2000: .ColAlignment(0) = flexAlignLeftCenter
90        .TextMatrix(0, 1) = "P Code": .ColWidth(1) = 1000: .ColAlignment(1) = flexAlignLeftCenter
100       .TextMatrix(0, 2) = "T Code": .ColWidth(2) = 1000: .ColAlignment(2) = flexAlignLeftCenter
110       .TextMatrix(0, 3) = "M Code": .ColWidth(3) = 1000: .ColAlignment(2) = flexAlignLeftCenter
120       .TextMatrix(0, 4) = "Q Code": .ColWidth(4) = 1000: .ColAlignment(2) = flexAlignLeftCenter
130       .TextMatrix(0, 5) = "": .ColWidth(5) = 250: .ColAlignment(2) = flexAlignCenterCenter
140   End With
End Sub

Private Sub InitializeGridgData()

10    With gData
20        .Rows = 2: .FixedRows = 1
30        .Cols = 5: .FixedCols = 0
40        .Rows = 1
50        .Font.Size = fgcFontSize
60        .Font.Name = fgcFontName

70        .ScrollBars = flexScrollBarBoth

80        .TextMatrix(0, 0) = "Consultant": .ColWidth(0) = 2000: .ColAlignment(0) = flexAlignLeftCenter
90        .TextMatrix(0, 1) = "": .ColWidth(1) = 500: .ColAlignment(1) = flexAlignLeftCenter
100       .TextMatrix(0, 2) = "": .ColWidth(2) = 250: .ColAlignment(2) = flexAlignLeftCenter
110       .TextMatrix(0, 3) = "Case No": .ColWidth(3) = 1200: .ColAlignment(3) = flexAlignLeftCenter
          '.TextMatrix(0, 4) = "Code": .ColWidth(4) = 1000: .ColAlignment(4) = flexAlignLeftCenter
          '.TextMatrix(0, 5) = "Description": .ColWidth(5) = 3800: .ColAlignment(5) = flexAlignCenterCenter
120       .TextMatrix(0, 4) = "Authorise Date": .ColWidth(4) = 1800: .ColAlignment(4) = flexAlignCenterCenter

130   End With
End Sub

Private Sub InitializeGridgDetail()
10    With gDetail
20        .Rows = 2: .FixedRows = 1
30        .Cols = 2: .FixedCols = 0
40        .Rows = 1
50        .Font.Size = fgcFontSize
60        .Font.Name = fgcFontName

70        .ScrollBars = flexScrollBarBoth

80        .TextMatrix(0, 0) = "Code": .ColWidth(0) = 1000: .ColAlignment(0) = flexAlignLeftCenter
90        .TextMatrix(0, 1) = "Description": .ColWidth(1) = 3350: .ColAlignment(1) = flexAlignLeftCenter
    
100   End With
End Sub

Private Sub FillPathologistList()
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillList_Error

20    sql = "SELECT * FROM Users WHERE AccessLevel = 'Consultant'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    cmbConsultant.AddItem ""
60    Do While Not tb.EOF
70        cmbConsultant.AddItem tb!UserName & ""
          '80      cmbConsultant.ItemData(cmbConsultant.NewIndex) = tb!UserId & ""
80        tb.MoveNext
90    Loop
100   cmbConsultant.ListIndex = -1

110   Exit Sub

FillList_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmWithPathologist", "FillList", intEL, strES, sql

End Sub

Private Sub gData_Click()


      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo gData_Click_Error

20    If gData.row = 0 Then Exit Sub

30    With gData
40        If .col = 2 Then
50            basShared.CaseNo = .TextMatrix(.row, 3)
60            PrintHistology "", True
70            With frmRichText
80                .cmdPrint.Visible = False
90                .cmdExit.Left = 0
100               .rtb.SelStart = 0
110               .Show 1
120           End With
      '        CaseNo = .TextMatrix(.Row, 3)
      '        With frmReportViewer
      '            .SampleID = CaseNo
      '            .Year = 2000 + Val(Right(CaseNo, 2))
      '            .Show 1
      '        End With
130       ElseIf .col = 3 Then
140           InitializeGridgDetail
150           sql = "SELECT CLL.CaseID, L.Code, L.Description FROM CaseListLink CLL " & _
                      "INNER JOIN Lists L ON CLL.ListID = L.ListID " & _
                      "WHERE CaseID = '" & .TextMatrix(.row, 3) & "' " & _
                      "ORDER BY DateTimeCreated Desc"
160           Set tb = New Recordset
170           RecOpenServer 0, tb, sql
180           If Not tb.EOF Then
190               While Not tb.EOF
200                   gDetail.AddItem tb!Code & "" & vbTab & _
                              tb!Description & ""
210                   tb.MoveNext
220               Wend
230           End If
240       End If

250   End With



260   Exit Sub

gData_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmAuthorisedReports", "gData_Click", intEL, strES

End Sub


Private Sub gData_EnterCell()

10    On Error GoTo gData_EnterCell_Error

20    With gData
30        If .col = 3 Then
40            .row = .row
50            .col = .col
60            .CellBackColor = vbYellow
70        End If
80    End With

90    Exit Sub

gData_EnterCell_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmAuthorisedReports", "gData_EnterCell", intEL, strES

End Sub

Private Sub gData_LeaveCell()

10    On Error GoTo gData_LeaveCell_Error


20    With gData
30        If .col = 3 Then
40            .row = .row
50            .col = .col
60            .CellBackColor = vbWhite
70        End If
80    End With


90    Exit Sub

gData_LeaveCell_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmAuthorisedReports", "gData_LeaveCell", intEL, strES

End Sub

Private Sub gFilter_Click()
10    On Error GoTo gFilter_Click_Error


20    With gFilter
30        If .row = 0 Or .MouseRow = 0 Then Exit Sub
40        Select Case .col
              Case 5:
50                If iMsg("Do you want to remove selected criteria", vbQuestion + vbYesNo) = vbYes Then
60                    If .Rows = 2 Then
70                        InitializeGridgFilter
80                    Else
90                        .RemoveItem .row
100                   End If
110               End If
120       End Select

130   End With

140   Exit Sub

gFilter_Click_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmAuthorisedReports", "gFilter_Click", intEL, strES

End Sub

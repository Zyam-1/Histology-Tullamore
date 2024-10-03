VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAccreditation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire -- Accreditation Settings"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Height          =   870
      Left            =   9525
      Picture         =   "frmAccreditation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5190
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   870
      Left            =   8295
      Picture         =   "frmAccreditation.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5190
      Width           =   975
   End
   Begin VB.TextBox txtDocAutopsy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7080
      TabIndex        =   10
      Top             =   3960
      Width           =   3435
   End
   Begin VB.TextBox txtDocCytology 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7080
      TabIndex        =   9
      Top             =   2565
      Width           =   3435
   End
   Begin VB.TextBox txtDocHistology 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7080
      TabIndex        =   8
      Top             =   1170
      Width           =   3435
   End
   Begin VB.TextBox txtAccAutopsy 
      Height          =   975
      Left            =   1200
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3690
      Width           =   5235
   End
   Begin VB.TextBox txtAccCytology 
      Height          =   975
      Left            =   1200
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2295
      Width           =   5235
   End
   Begin VB.TextBox txtAccHistology 
      Height          =   975
      Left            =   1200
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   900
      Width           =   5235
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   240
      TabIndex        =   13
      Top             =   6300
      Visible         =   0   'False
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000003&
      X1              =   240
      X2              =   10620
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000003&
      X1              =   240
      X2              =   10620
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      X1              =   240
      X2              =   10620
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Document Control Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7080
      TabIndex        =   7
      Top             =   300
      Width           =   3195
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      X1              =   7080
      X2              =   10680
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Autopsy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   3690
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cytology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   2
      Top             =   2295
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Histology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   1
      Top             =   900
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Accreditation Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   300
      Width           =   2190
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   1200
      X2              =   4800
      Y1              =   660
      Y2              =   660
   End
End
Attribute VB_Name = "frmAccreditation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

10  On Error GoTo cmdAdd_Click_Error

20  SaveOptionSetting "HistologyAccreditationText", txtAccHistology
30  SaveOptionSetting "CytologyAccreditationText", txtAccCytology
40  SaveOptionSetting "AutopsyAccreditationText", txtAccAutopsy

50  SaveOptionSetting "HistologyDocumentNo", txtDocHistology
60  SaveOptionSetting "CytologyDocumentNo", txtDocCytology
70  SaveOptionSetting "AutopsyDocumentNo", txtDocAutopsy

80  iMsg "Settings saved!"

90  Exit Sub

cmdAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmAccreditation", "cmdAdd_Click", intEL, strES

End Sub

Private Sub cmdExit_Click()
10  On Error GoTo cmdExit_Click_Error

20  Unload Me

30  Exit Sub

cmdExit_Click_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmAccreditation", "cmdExit_Click", intEL, strES

End Sub

Private Sub Form_Load()

10  On Error GoTo Form_Load_Error

20  txtAccHistology = GetOptionSetting("HistologyAccreditationText", "")
30  txtAccCytology = GetOptionSetting("CytologyAccreditationText", "")
40  txtAccAutopsy = GetOptionSetting("AutopsyAccreditationText", "")

50  txtDocHistology = GetOptionSetting("HistologyDocumentNo", "")
60  txtDocCytology = GetOptionSetting("CytologyDocumentNo", "")
70  txtDocAutopsy = GetOptionSetting("AutopsyDocumentNo", "")

80  If blnIsTestMode Then EnableTestMode Me

90  Exit Sub

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmAccreditation", "Form_Load", intEL, strES

End Sub

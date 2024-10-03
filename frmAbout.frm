VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5100
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   3390
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   3
      Top             =   420
      Width           =   3315
   End
   Begin VB.Label lblDescription 
      Caption         =   "Histology Coding System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   585
      Left            =   165
      TabIndex        =   2
      Top             =   855
      Width           =   4650
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":0ECA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   480
      TabIndex        =   1
      Top             =   1590
      Width           =   4035
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

40    Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmAbout", "Form_Load", intEL, strES

End Sub

Private Sub Form_Resize()
10    If Me.WindowState <> vbMinimized Then

20        Me.Top = 0
30        Me.Left = Screen.Width / 2 - Me.Width / 2
40    End If
End Sub


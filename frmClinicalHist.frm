VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmClinicalHist 
   Caption         =   "Clinical History"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   6300
   Begin VB.CommandButton cmdAdd 
      Height          =   870
      Left            =   5160
      Picture         =   "frmClinicalHist.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Height          =   870
      Left            =   5160
      Picture         =   "frmClinicalHist.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   570
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtAmendment 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmClinicalHist.frx":0614
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmClinicalHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pDescription As String

Private Sub cmdAdd_Click()
10    frmWorkSheet.lblClinicalHist = txtAmendment.Text
20    DataChanged = True
30    Unload Me
End Sub

Private Sub cmdExit_Click()
10    Unload Me
End Sub

Private Sub Form_Load()
10    txtAmendment.Text = pDescription
20    If bLocked Then
30        txtAmendment.Locked = True
40    End If
End Sub
Public Property Let Description(ByVal Desc As String)

10    pDescription = Desc

End Property


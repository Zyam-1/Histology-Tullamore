VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRichText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Cellular Pathology"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   Icon            =   "frmRichText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Height          =   435
      Left            =   480
      Picture         =   "frmRichText.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   435
      Left            =   0
      Picture         =   "frmRichText.frx":120C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin RichTextLib.RichTextBox tempRtb 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRichText.frx":1627
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   9435
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   16642
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmRichText.frx":16A9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   960
      TabIndex        =   4
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
   Begin VB.Label lblLoggedIn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Logged In : "
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmRichText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pRtbTextBox As String

Private Sub cmdExit_Click()
10    Unload Me
End Sub

Private Sub cmdPrint_Click()
10    Unload Me
20    PrintHistology

End Sub

Private Sub Form_Activate()
10    If pRtbTextBox = "GROSS" Then
20        rtb = frmWorkSheet.txtGross
30        rtb.SelStart = Len(rtb.Text)
40        cmdPrint.Visible = False
50        cmdExit.Visible = False
60    ElseIf pRtbTextBox = "MICRO" Then
70        rtb = frmWorkSheet.txtMicro
80        rtb.SelStart = Len(rtb.Text)
90        cmdPrint.Visible = False
100       cmdExit.Visible = False
110   ElseIf pRtbTextBox = "AMENDMENTS" Then
120       rtb = frmAmendments.txtAmendment
130       rtb.SelStart = Len(rtb.Text)
140       cmdPrint.Visible = False
150       cmdExit.Visible = False
160   Else
170       rtb.Locked = True
180   End If
End Sub

Private Sub Form_Load()

10    lblLoggedIn = UserName
20    If blnIsTestMode Then EnableTestMode Me

End Sub

Private Sub Form_Resize()
10    If Me.WindowState <> vbMinimized Then

20        Me.Top = 0
30        Me.Left = Screen.Width / 2 - Me.Width / 2
40    End If
End Sub
Public Property Let rtbTextBox(ByVal strNewValue As String)

10    pRtbTextBox = strNewValue

End Property

Private Sub Form_Unload(Cancel As Integer)
10    If pRtbTextBox = "GROSS" Then
20        If frmWorkSheet.txtGross <> rtb Then
30            DataChanged = True
40        End If
50        frmWorkSheet.txtGross = rtb
60        frmWorkSheet.txtGross.SelStart = Len(frmWorkSheet.txtGross.Text)
70        pRtbTextBox = ""
80    ElseIf pRtbTextBox = "MICRO" Then
90        If frmWorkSheet.txtMicro <> rtb Then
100           DataChanged = True
110       End If
120       frmWorkSheet.txtMicro = rtb
130       frmWorkSheet.txtMicro.SelStart = Len(frmWorkSheet.txtMicro.Text)
140       pRtbTextBox = ""
150   ElseIf pRtbTextBox = "AMENDMENTS" Then
160       If frmAmendments.txtAmendment <> rtb Then
170           DataChanged = True
180       End If
190       frmAmendments.txtAmendment = rtb
200       frmAmendments.txtAmendment.SelStart = Len(frmAmendments.txtAmendment.Text)
210       pRtbTextBox = ""
220   End If
End Sub

Private Sub rtb_KeyDown(KeyCode As Integer, Shift As Integer)

10 If UCase$(UserMemberOf) = "LOOKUP" Then
20  If KeyCode = vbKeyC And Shift = vbCtrlMask Then
30      MsgBox "Copy text disabled!", vbExclamation
40      KeyCode = 0
50  End If
60 End If

End Sub

Private Sub rtb_KeyUp(KeyCode As Integer, Shift As Integer)

10    If KeyCode = 44 Then 'Print Screen
20        Clipboard.Clear
30    End If

End Sub

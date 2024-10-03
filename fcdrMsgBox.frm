VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fcdrMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FcdrMsgBox"
   ClientHeight    =   2115
   ClientLeft      =   1920
   ClientTop       =   2505
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton b 
      Caption         =   "&Ignore"
      Height          =   525
      Index           =   5
      Left            =   4410
      TabIndex        =   7
      Top             =   1350
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&Retry"
      Height          =   525
      Index           =   4
      Left            =   4410
      TabIndex        =   6
      Top             =   750
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&Abort"
      Height          =   525
      Index           =   3
      Left            =   4410
      TabIndex        =   5
      Top             =   150
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&Yes"
      Height          =   525
      Index           =   6
      Left            =   4410
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&Cancel"
      Height          =   525
      Index           =   2
      Left            =   4410
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "O. K."
      Height          =   525
      Index           =   1
      Left            =   4410
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton b 
      Caption         =   "&No"
      Height          =   525
      Index           =   7
      Left            =   4410
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   48
      Left            =   180
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   32
      Left            =   180
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   64
      Left            =   180
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   16
      Left            =   180
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   1695
      Left            =   840
      TabIndex        =   1
      Top             =   150
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fcdrMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ReturnValue As Integer

Private mDefaultButton As Integer
Private mMsgFontSize As Long

Private mButtons As Long

Private mIcon As Long
Private mMessage As String

Private Sub b_Click(Index As Integer)

10    On Error GoTo b_Click_Error

20    ReturnValue = Index
30    Unload Me

40    Exit Sub

b_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    Screen.MousePointer = 0

60    intEL = Erl
70    strES = Err.Description
80    LogError "fcdrMsgBox", "b_Click", intEL, strES


End Sub

Public Property Let DefaultButton(ByVal lngButton As Integer)

10    On Error GoTo DefaultButton_Error

20    mDefaultButton = lngButton

30    Exit Property

DefaultButton_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "fcdrMsgBox", "DefaultButton", intEL, strES


End Property

Public Property Let DisplayButtons(ByVal intButtons As Long)

10    On Error GoTo DisplayButtons_Error

20    mButtons = intButtons

30    Exit Property

DisplayButtons_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "fcdrMsgBox", "DisplayButtons", intEL, strES


End Property

Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error

20    If mDefaultButton > 0 Then
30      b(mDefaultButton).Default = True
40    End If

50    If mMsgFontSize <> 0 Then
60      lblMessage.Font.Size = mMsgFontSize
70    End If
  
80    Select Case mIcon
        Case 16, 32, 48, 64: i(mIcon).Visible = True
  Case Else:
90    End Select
  
100   lblMessage = mMessage
  
110   Select Case mButtons
        Case 0: 'MB_OK 0 Display OK button only.
120       b(1).Visible = True
130       b(1).Cancel = True
  Case 1: 'MB_OKCANCEL 1 Display OK and Cancel buttons.
140       b(1).Visible = True
150       b(2).Visible = True
160       b(2).Cancel = True
          'SELECT Case DefaultButton
          '  Case 0: .b(1).Default = True
          '  Case 256: .b(2).Default = True
          'End SELECT
  
  Case 2: 'MB_ABORTRETRYIGNORE 2 Display Abort, Retry, and Ignore buttons.
170       b(3).Visible = True
180       b(4).Visible = True
190       b(5).Visible = True
      '      SELECT Case DefaultButton
      '        Case 0: .b(3).Default = True
      '        Case 256: .b(4).Default = True
      '        Case 512: .b(5).Default = True
      '      End SELECT
  
  Case 3: 'MB_YESNOCANCEL  3 Display Yes, No, and Cancel buttons.
200       b(6).Visible = True
210       b(7).Visible = True
220       b(2).Visible = True
230       b(2).Cancel = True
          'SELECT Case DefaultButton
          '  Case 0: .b(6).Default = True
          '  Case 256: .b(7).Default = True
          '  Case 512: .b(2).Default = True
          'End SELECT
  
  Case 4: 'MB_YESNO  4 Display Yes and No buttons.
240       b(6).Visible = True
250       b(7).Visible = True
          'SELECT Case DefaultButton
          '  Case 0: .b(6).Default = True
          '  Case 256: .b(7).Default = True
          'End SELECT
  
  Case 5: 'MB_RETRYCANCEL  5 Display Retry and Cancel buttons.
260       b(4).Visible = True
270       b(2).Visible = True
280       b(2).Cancel = True
          'SELECT Case DefaultButton
          '  Case 0: .b(4).Default = True
          '  Case 256: .b(2).Default = True
          'End SELECT

290   End Select

300   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

310   Screen.MousePointer = 0

320   intEL = Erl
330   strES = Err.Description
340   LogError "fcdrMsgBox", "Form_Activate", intEL, strES


End Sub

Public Property Let Message(ByVal strMessage As String)

10    On Error GoTo Message_Error

20    mMessage = strMessage

30    Exit Property

Message_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "fcdrMsgBox", "Message", intEL, strES


End Property

Public Property Let MsgFontSize(ByVal FntSize As Long)

10    On Error GoTo MsgFontSize_Error

20    mMsgFontSize = FntSize

30    Exit Property

MsgFontSize_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "fcdrMsgBox", "MsgFontSize", intEL, strES


End Property

Public Property Get RetVal() As Long

10    On Error GoTo Retval_Error

20    RetVal = ReturnValue

30    Exit Property

Retval_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "fcdrMsgBox", "Retval", intEL, strES


End Property

Public Property Let ShowIcon(ByVal intIcon As Long)

10    On Error GoTo ShowIcon_Error

20    mIcon = intIcon

30    Exit Property

ShowIcon_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "fcdrMsgBox", "ShowIcon", intEL, strES


End Property

Private Sub Form_Load()
10    If blnIsTestMode Then EnableTestMode Me
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3210
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3210
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAction 
      Height          =   375
      Index           =   0
      Left            =   345
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox chkDontAsk 
      Caption         =   "Don't ask me again"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblMsg 
      Caption         =   "Msg"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum eMsgIcon
    mbNone
    mbExclamation
    mbInformation
    mbCritical
    mbQuestion
End Enum

Public Enum StandardIconEnum
    IDI_ASTERISK = 32516&
    IDI_EXCLAMATION = 32515&
    IDI_HAND = 32513&
    IDI_QUESTION = 32514&
End Enum

Public Enum eBtns
    mbOKOnly
    mbOKCancel
    mbAbortRetryIgnore
    mbYesNoCancel
    mbYesNo
    mbRetryCancel
End Enum

Private Declare Function LoadStandardIcon Lib "user32" Alias "LoadIconA" _
                                        (ByVal hInstance As Long, _
                                         ByVal lpIconNum As StandardIconEnum) As Long
    
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, _
                                                ByVal X As Long, _
                                                ByVal Y As Long, _
                                                ByVal hIcon As Long) As Long

Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Private Const MB_ICONEXCLAMATION = 49
Private Const MB_ICONHAND = 17
Private Const MB_ICONINFORMATION = 65

'varables used to return the users chossen options
Public g_lBtnClicked As Long
Public g_bDontAsk As Boolean

Public Function Msg(ByRef Promt As String, _
                    Optional ByRef Buttons As eBtns = mbOKOnly, _
                    Optional ByRef Title As String, _
                    Optional ByRef MsgIcon As eMsgIcon = mbNone, _
                    Optional ByRef DefaultBtn As Long = 1, _
                    Optional ByRef ShowDontAsk As Boolean = False) As Long
      'If Title is left blank, App.Path is used
      'If you don't want any title just set Title = " "
      'Btn captions are made automatically
          Dim sBtnText() As String
    
10        SetTitleBar Title
20        LoadIcon MsgIcon
30        SetLabelWidth Promt, MsgIcon
    
40        Select Case Buttons
              Case mbOKOnly
50                sBtnText = Split("Ok")
60            Case mbOKCancel
70                sBtnText = Split("Ok|Cancel", "|")
80            Case mbAbortRetryIgnore
90                sBtnText = Split("Abort|Retry|Ignore", "|")
100           Case mbYesNoCancel
110               sBtnText = Split("Yes|No|Cancel", "|")
120           Case mbYesNo
130               sBtnText = Split("Yes|No", "|")
140           Case mbRetryCancel
150               sBtnText = Split("Retry|Cancel", "|")
160           Case Else
170               sBtnText = Split("Ok")
180       End Select
190       SetBtns sBtnText, MsgIcon
200       SetDontAsk ShowDontAsk
210       PositionForm DefaultBtn 'ParentForm, Me.Width
220       Msg = g_lBtnClicked
End Function

Public Function MsgCstm(ByRef Promt As String, _
                        ByRef Title As String, _
                        ByRef MsgIcon As eMsgIcon, _
                        ByRef DefaultBtn As Long, _
                        ByRef ShowDontAsk As Boolean, _
                        ParamArray btnText()) As Long
      'sets the user msg, Title and number and text on the buttons
      'If Title is left blank, App.Path is used
          Dim lB As Long
          Dim sBtnText() As String

10        SetTitleBar Title
20        LoadIcon MsgIcon
30        SetLabelWidth Promt, MsgIcon
    
40        ReDim sBtnText(UBound(btnText))
50        For lB = 0 To UBound(btnText)
60            sBtnText(lB) = CStr(btnText(lB))
70        Next
80        SetBtns sBtnText, MsgIcon
90        SetDontAsk ShowDontAsk
100       PositionForm DefaultBtn 'ParentForm, Me.Width
110       MsgCstm = g_lBtnClicked
End Function

Private Sub btnAction_Click(Index As Integer)
          'add one to the Btn#, 0 indicates they closed without hitting any btn
10        g_lBtnClicked = Index + 1
20        g_bDontAsk = CBool(chkDontAsk.Value)
30        Unload Me
End Sub

Private Sub Form_Load()
10        g_lBtnClicked = 0
20        g_bDontAsk = False
30        Me.Font = btnAction(0).Font
40        lblMsg.Font = Me.Font
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
          Dim obj As Object

10        On Error Resume Next
20        Me.Cls
      '    For Each obj In Me 'frm
      '        Unload obj
      '        Set obj = Nothing
      '    Next
30        Unload Me
40        On Error GoTo 0
End Sub

Private Sub LoadIcon(ByRef MsgIcon As eMsgIcon)
          Dim hIcon As Long

10        If MsgIcon Then
              'show the icon and play the right sound
20            Select Case MsgIcon
                  Case mbExclamation
30                    hIcon = LoadStandardIcon(0&, IDI_EXCLAMATION)
                      'MessageBeep MB_ICONEXCLAMATION '49
40                Case mbInformation
50                    hIcon = LoadStandardIcon(0&, IDI_ASTERISK)
                      'MessageBeep MB_ICONINFORMATION '65
60                Case mbCritical
70                    hIcon = LoadStandardIcon(0&, IDI_HAND)
                      'MessageBeep MB_ICONHAND '17
80                Case mbQuestion
90                    hIcon = LoadStandardIcon(0&, IDI_QUESTION)
100           End Select
110           Call DrawIcon(Me.hDC, 9&, 10&, hIcon)
120       Else
130           Me.Cls
140       End If
End Sub

Private Sub PositionForm(ByRef DefaultBtn As Long)

10        If DefaultBtn > btnAction.Count Then
20            DefaultBtn = btnAction.Count
30        ElseIf DefaultBtn < 1 Then
40            DefaultBtn = 1
50        End If
60        btnAction(DefaultBtn - 1).tabIndex = 0
70        Me.Show vbModal
End Sub

Private Sub SetBtns(ByRef btnText() As String, _
                    ByRef MsgIcon As eMsgIcon)
          Dim lX As Long
          Dim lUb As Long
          Dim lWidth As Long
          Dim lRightMost As Long
          Dim lRowWidth() As Long
          Dim lBtnsInRow() As Long
          Dim lR As Long
          Dim lCnt As Long
          Dim lBtnTop As Long
          Dim lMaxWidth As Long
    
10        lBtnTop = lblMsg.Height + (lblMsg.Top * 2)
20        If lBtnTop < 900 Then
30            If MsgIcon Then
40                lBtnTop = 900
50            End If
60        End If
70        btnAction(0).Top = lBtnTop
    
80        Select Case Me.Width
              Case Is < Screen.Width / 4
90                lMaxWidth = Screen.Width / 4
100           Case Is < Screen.Width / 2
110               lMaxWidth = Screen.Width / 2
120           Case Else ' Is < Screen.Width / 4
130               lMaxWidth = Screen.Width * 0.75
140       End Select
    
150       ReDim lRowWidth(0)
160       ReDim lBtnsInRow(0)
170       lUb = UBound(btnText)
180       For lX = 0 To lUb
190           With btnAction(lX)
200               If lX Then
                      'dynamically load the needed buttons
210                   Load btnAction(lX)
220                   .Top = btnAction(lX - 1).Top
230                   .Left = btnAction(lX - 1).Left + btnAction(lX - 1).Width + 120
240               End If
                  'set the button width and text
250               .Width = Me.TextWidth(btnText(lX) & "WW") '"WW" is a buffer to make extra room on the button
260               .Caption = btnText(lX)
270               .Visible = True
                  'wrap the buttons if needed
280               If .Width + .Left + 120 > lMaxWidth Then
290                   lR = lR + 1
300                   ReDim Preserve lRowWidth(lR)
310                   ReDim Preserve lBtnsInRow(lR)
320                   .Left = btnAction(0).Left
330                   .Top = btnAction(lX - 1).Top + btnAction(lX - 1).Height + 120
340                   If btnAction(lX - 1).Left + btnAction(lX - 1).Width > _
                          btnAction(lRightMost).Left + btnAction(lRightMost).Width Then
350                       lRightMost = lX - 1
360                   End If
370               End If
380               lRowWidth(lR) = lRowWidth(lR) + .Width + 120
390               lBtnsInRow(lR) = lBtnsInRow(lR) + 1
400           End With
410       Next
    
          'adjust the width of the msg box
420       lWidth = Me.Width
430       If lRightMost = 0 Then
440           lRightMost = lUb
450       End If
460       If btnAction(lRightMost).Left + btnAction(lRightMost).Width + btnAction(0).Left > lWidth Then
470           lWidth = btnAction(lRightMost).Left + btnAction(lRightMost).Width + btnAction(0).Left
480       End If
490       Me.Width = lWidth
    
          'center the button rows
500       For lUb = 0 To UBound(lRowWidth)
510           lWidth = lRowWidth(lUb) - 120
520           lWidth = ((Me.Width - lWidth) / 2) - 30
530           For lR = 0 To lBtnsInRow(lUb) - 1
540               If lR = 0 Then
550                   btnAction(lCnt).Left = lWidth
560               Else
570                   btnAction(lCnt).Left = btnAction(lCnt - 1).Left + btnAction(lCnt - 1).Width + 120
580               End If
590               lCnt = lCnt + 1
600           Next
610       Next
End Sub

Private Sub SetDontAsk(ByRef ShowDontAsk As Boolean)
          Dim lUb As Long
    
10        lUb = btnAction.Count - 1
          'set the height of the form
20        If ShowDontAsk Then
30            chkDontAsk.Value = 0
40            chkDontAsk.Top = btnAction(lUb).Top + btnAction(lUb).Height + 120
50            chkDontAsk.Visible = True
60            Me.Height = chkDontAsk.Top + chkDontAsk.Height + 630 '585 '645
70        Else
80            chkDontAsk.Visible = False
90            Me.Height = btnAction(lUb).Top + btnAction(lUb).Height + 630 '585 '645
100       End If
End Sub

Private Sub SetLabelWidth(ByRef Promt As String, ByRef MsgIcon As eMsgIcon)
      'Make sure that the Promt Label doesn't cause the form to be wider that the screen.
    
10        lblMsg.Caption = Promt
20        lblMsg.Width = Me.TextWidth(Promt)
30        If MsgIcon Then
40            lblMsg.Left = 780
50        Else
60            lblMsg.Left = 180
70        End If
80        Me.Width = lblMsg.Left + lblMsg.Width + 240 '120
90        If Me.Width < 3330 Then Me.Width = 3330
100       lblMsg.Height = Me.TextHeight(Promt)
110       If lblMsg.Left + lblMsg.Width + 240 > Screen.Width * 0.75 Then
120           lblMsg.AutoSize = True
130           lblMsg.WordWrap = True
140           lblMsg.Width = (Screen.Width * 0.75) - (lblMsg.Left + 120)
150           Me.Width = Screen.Width
160       Else
170           lblMsg.WordWrap = False
180       End If
End Sub

Private Sub SetTitleBar(ByRef Title As String)
10        If Len(Title) Then
20            Me.Caption = Title
30        Else
40            Me.Caption = App.Title 'Path
50        End If
End Sub


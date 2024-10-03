VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmInputNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAquire - Cellular Pathology"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmInputNo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Top             =   270
      Width           =   735
   End
   Begin VB.CommandButton cmdCode 
      Height          =   285
      Left            =   4800
      Picture         =   "frmInputNo.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   270
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Please Enter Number of Blocks"
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   270
      Width           =   3645
   End
End
Attribute VB_Name = "frmInputNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pInputType As String
Private pLabel As String

Public Property Let InputType(ByVal Value As String)

10    pInputType = Value

End Property
Public Property Let Label(ByVal strNewValue As String)

10    pLabel = strNewValue

End Property

Private Sub cmdCode_Click()
      Dim iInputNo As Integer
      Dim iNo As Integer
      Dim iIncrement As Integer
      Dim i As Integer
      Dim tnode As MSComctlLib.Node
      Dim NoOfBlocks As Integer
      Dim NoOfSlides As Integer
      Dim UniqueId As String

10    On Error GoTo cmdCode_Click_Error

20    With frmWorkSheet.tvCaseDetails.SelectedItem
30        iIncrement = 65
40        If txtInput <> "" Then
50            iInputNo = CInt(txtInput)
60            If iInputNo = 0 Then
70                Unload Me
80                Exit Sub
90            End If
100           If pInputType = "B" Then
                  'ZcsFrozenSection
110               If InStr(1, frmWorkSheet.tvCaseDetails.SelectedItem.Text, "Frozen Section") Then
120                   NoOfBlocks = GetBlockNumber(frmWorkSheet.tvCaseDetails.SelectedItem.Parent)
130               Else
140                   NoOfBlocks = GetBlockNumber(frmWorkSheet.tvCaseDetails.SelectedItem)
150               End If
160               If NoOfBlocks + iInputNo > 200 Then
170                   frmMsgBox.Msg "Cannot enter more than 200 blocks. " & .Children & " Blocks already added.", , , mbExclamation
180                   Exit Sub
190               End If
200               For i = 0 To iInputNo - 1
210                   If InStr(1, frmWorkSheet.tvCaseDetails.SelectedItem.Text, "Frozen Section") Then
220                       NoOfBlocks = GetBlockNumber(frmWorkSheet.tvCaseDetails.SelectedItem.Parent)
230                   Else
240                       NoOfBlocks = GetBlockNumber(frmWorkSheet.tvCaseDetails.SelectedItem)
250                   End If
260                   UniqueId = GetUniqueID
          
270                   If sysOptBlockNumberingFormat(0) = "1" Then
280                       iNo = NoOfBlocks + 1
290                       Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(.Key, tvwChild, "L2" & UniqueId, "Block" & " " & iNo, 1, 2)
300                       If InStr(1, frmWorkSheet.tvCaseDetails.SelectedItem.Text, "Frozen Section") = 0 Then
310                           AddDefaultStains "L2" & UniqueId, frmWorkSheet.tvCaseDetails.SelectedItem
320                       End If
330                   Else
340                       iNo = NoOfBlocks
350                       If NoOfBlocks = 1 Then
360                           .Child.Text = "Block" & " A"
370                       End If
        
380                       Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(.Key, tvwChild, "L2" & UniqueId, "Block" & " " & IIf(AddLetter(iNo) = "A", "", AddLetter(iNo)), 1, 2)
390                       If InStr(1, frmWorkSheet.tvCaseDetails.SelectedItem.Text, "Frozen Section") Then
400                           AddDefaultStains "L2" & UniqueId, frmWorkSheet.tvCaseDetails.SelectedItem
410                       End If
420                   End If
430               Next
440               tnode.Selected = True
450           ElseIf pInputType = "S" Then
460               NoOfSlides = GetSlideNumber(frmWorkSheet.tvCaseDetails.SelectedItem)
470               If NoOfSlides + iInputNo > 100 Then
480                   frmMsgBox.Msg "Cannot enter more than 100 slides. " & .Children & " Slides already added.", , , mbExclamation
490                   Exit Sub
500               End If
510               For i = 0 To iInputNo - 1
520                   NoOfSlides = GetSlideNumber(frmWorkSheet.tvCaseDetails.SelectedItem)
          
530                   UniqueId = GetUniqueID
          
540                   If sysOptSlideNumberingFormat(0) = "1" Then
550                       iNo = NoOfSlides + 1
560                       Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(.Key, tvwChild, "L3" & UniqueId, "Slide" & " " & iNo, 1, 2)
570                   Else
580                       iNo = NoOfSlides
590                       If NoOfSlides = 1 Then
600                           .Child.Text = "Slide" & " A"
610                       End If
        
620                       Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(.Key, tvwChild, "L3" & UniqueId, "Slide" & " " & IIf(AddLetter(iNo) = "A", "", AddLetter(iNo)), 1, 2)
630                   End If
640               Next
650               tnode.Selected = True
660           ElseIf pInputType = "L" Then
670                   .Tag = iInputNo
680                   .ForeColor = vbBlue
                
690           ElseIf pInputType = "C" Then
700               .Tag = iInputNo
710           End If

        
        
        
720           .Expanded = True
730           DataChanged = True
740           TreeChanged = True
750           Unload Me
760       End If
770   End With

780   Exit Sub

cmdCode_Click_Error:

      Dim strES As String
      Dim intEL As Integer

790   intEL = Erl
800   strES = Err.Description
810   LogError "frmInputNo", "cmdCode_Click", intEL, strES

End Sub

Private Sub Form_Load()
With Me
    .Caption = "Histology"
    Select Case pInputType
        Case "L", "C"
            .lblCaption.Caption = "Please Enter Number Of Levels"
        Case "B"
            .lblCaption.Caption = "Please Enter Number Of Blocks"
        Case "S"
            .lblCaption = "Please Enter Number Of Slides"
        Case Else
            .lblCaption = pLabel
    End Select
'
End With


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End If
End Sub

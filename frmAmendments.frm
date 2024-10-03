VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAmendments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amendments"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Height          =   870
      Left            =   5160
      Picture         =   "frmAmendments.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1410
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   870
      Left            =   5160
      Picture         =   "frmAmendments.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
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
      ScrollBars      =   3
      TextRTF         =   $"frmAmendments.frx":0614
   End
End
Attribute VB_Name = "frmAmendments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pAmendId As String
Private pCode As String
Private pDescription As String
Private pUpdate As Boolean

Private Sub cmdAdd_Click()
      Dim i As Integer
      Dim Location As Integer
10    Location = 0
20    For i = 0 To frmWorkSheet.grdAmendments.Rows - 1
          'if an amendment is being edited then need to get the location in the grid from where its is being amended
30        If frmWorkSheet.grdAmendments.TextMatrix(i, 2) = pAmendId And frmWorkSheet.grdAmendments.TextMatrix(i, 3) = pCode Then
40            Location = i
50        End If
60    Next
70    If Location > 0 Then
80        frmWorkSheet.grdAmendments.TextMatrix(Location, 1) = Replace(txtAmendment.Text, vbTab, "<<tab>>")
90    Else
          'if location is 0 then adding a new amendment
   
100       frmWorkSheet.grdAmendments.AddItem Format(Now, "dd/mm/yy hh:mm") & vbTab & Replace(txtAmendment.Text, vbTab, "<<tab>>") & vbTab & pAmendId & vbTab & pCode
110       For i = 1 To frmWorkSheet.grdAmendments.Rows - 1
              'If Q021 then mark it red
120          If frmWorkSheet.grdAmendments.TextMatrix(i, 3) = "Q021" Then
130                  With frmWorkSheet.grdAmendments
140                      .row = i
150                      .col = 0
160                      .CellForeColor = vbRed
170                      .col = 1
180                      .CellForeColor = vbRed
190                  End With
200          End If
210       Next
220   End If
230   DataChanged = True
240   Unload Me
End Sub

Private Sub cmdExit_Click()
10    If pUpdate = False Then
20        If frmWorkSheet.grdQCodes.Rows - frmWorkSheet.grdQCodes.FixedRows = 1 Then
30            frmWorkSheet.grdQCodes.Rows = frmWorkSheet.grdQCodes.Rows - 1
40        Else
50            frmWorkSheet.grdQCodes.RemoveItem frmWorkSheet.grdQCodes.Rows - 1
60        End If
70    End If

80    Unload Me
End Sub

Public Property Let AmendId(ByVal Id As String)

10    pAmendId = Id

End Property

Public Property Let Code(ByVal Id As String)

10    pCode = Id

End Property
Public Property Let Description(ByVal Desc As String)

10    pDescription = Desc

End Property
Public Property Let Update(ByVal Value As Boolean)

10    pUpdate = Value

End Property


Private Sub Form_Load()
          Dim s As String
          Dim tb As Recordset
          Dim sql As String
          Dim ValidAmend As Boolean
    
    
    
10        If pUpdate = True Then
    
20            sql = "SELECT * FROM CaseAmendments WHERE CaseListId = " & pAmendId
30            Set tb = New Recordset
40            RecOpenServer 0, tb, sql
  
50            If Not tb.EOF Then
60                ValidAmend = IIf(IsNull(tb!Valid), 0, tb!Valid)
70            Else
80                ValidAmend = False
90            End If
  
100           s = frmWorkSheet.grdAmendments.TextMatrix(Rada, 1)
110           txtAmendment.Text = Replace(s, "<<tab>>", vbTab)
  
120           If ValidAmend Then
130               txtAmendment.Locked = True
140           Else
150               txtAmendment.Locked = False
160           End If
  
170       End If

    
End Sub

Private Sub txtAmendment_DblClick()
10    With frmRichText

20        .rtbTextBox = "AMENDMENTS"
30        .Show 1
40    End With
End Sub

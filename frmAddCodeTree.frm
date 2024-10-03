VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAddCodeTree 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAquire - Cellular Pathology"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "frmAddCodeTree.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCode 
      Height          =   285
      Left            =   5160
      Picture         =   "frmAddCodeTree.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1200
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3720
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmAddCodeTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ReturnValue As String
Public tvtemp As MSComctlLib.TreeView
'Private tnode As Node
Private tnode As MSComctlLib.Node
Private pListType As String
Private pListTypeName As String    'Singular name
Private pListTypeNames As String    'Plural name
Private pLevel As String
Private pUpdate As Boolean

'Orignal Code
'Private Sub cmdCode_Click()
'      Dim i As Integer
'      Dim UniqueId As String
'      Dim sql As String
'      Dim tb As New Recordset
'      Dim CaseNode As MSComctlLib.Node
'      Dim tempListId As String
'      Dim strPrevId As String
'      Dim strNewId As String
'      Dim strLetter As String
'
'10    On Error GoTo cmdCode_Click_Error
'
'20    If txtCode <> "" Then
'
'30        With tvtemp.SelectedItem
'
'              'If adding a T Code
'40            If pListType = "T" Then
'                  'Check the options table to see if numbering format is numbers or letters (ie. "A : TY4100 : Abdominal Cavity" or "1 : TY4100 : Abdominal Cavity")
'50                If sysOptTissueTypeNumberingFormat(0) = "1" Then
'60                    i = .Children + 1
'70                Else
'80                    i = .Children + 65
'90                End If
'                  'If updating an existing T Code in the tree then update the the text
'                  'else add to the tree using correct numbering format
'100               If pUpdate Then
'110                   ReturnValue = Left(tvtemp.SelectedItem.Text, 1) & " : " & txtCode & " : " & txtDescription
'120                   .Text = ReturnValue
'130                   strPrevId = .Tag    'Previous T Code
'
'                      'set the tag to be the T Code list id from list table
'140                   .Tag = CStr(frmList.ListId)
'150                   strNewId = .Tag    'New T Code
'160                   strLetter = Left(tvtemp.SelectedItem.Text, 1)    'Tissue Letter
'
'                      'Update the Tissue Code in the grids grdTempMCode grdMCodes
'                      'This is because these are used in the Save routines
'170                   Call UpdateTissueTypeListId(strPrevId, strNewId, strLetter)
'
'180               Else
'190                   UniqueId = GetUniqueID
'
'200                   If sysOptTissueTypeNumberingFormat(0) = "1" Then
'210                       ReturnValue = i & " : " & txtCode & " : " & txtDescription
'220                   Else
'230                       ReturnValue = Chr(i) & " : " & txtCode & " : " & txtDescription
'240                   End If
'                      'add node at Level 1
'250                   Set tnode = tvtemp.Nodes.Add(.Key, tvwChild, "L1" & UniqueId, ReturnValue, 1, 2)
'260                   tnode.Tag = CStr(frmList.ListId)
'270               End If
'
'
'
'                  'update the rank of the T Code so that it appears
'                  'higher in list the next time T Code is searched
'280               UpdateListRank GetListID(txtCode, pListType)
'290           Else
'                  'adding a stain
'300               ReturnValue = txtDescription
'
'310               If .Parent.Tag <> "" Then
'320                   tempListId = .Parent.Tag
'330               ElseIf .Parent.Parent.Tag <> "" Then
'340                   tempListId = .Parent.Parent.Tag
'350               ElseIf .Parent.Parent.Parent.Tag <> "" Then
'360                   tempListId = .Parent.Parent.Parent.Tag
'370               End If
'
'
'380               sql = "SELECT Code FROM Lists WHERE ListId = '" & tempListId & "' "
'
'390               Set tb = New Recordset
'400               RecOpenClient 0, tb, sql
'
'
'                  'If its level 2 and a Block
'410               If pLevel = "L2" And InStr(1, tvtemp.SelectedItem.Text, "Block") Then
'420                   Set tnode = tvtemp.SelectedItem
'430                   UniqueId = AddBlockLevelStain(tvtemp.SelectedItem.Key, tnode, ReturnValue)
'440               Else
'                      'else adding a stain to a slide
'450                   UniqueId = GetUniqueID
'
'460                   Set tnode = tvtemp.Nodes.Add(.Key, tvwChild, "L4" & UniqueId, ReturnValue, 1, 2)
'470                   tnode.Tag = UniqueId
'480                   If (UCase$(UserMemberOf) = "CONSULTANT" Or _
'                          UCase$(UserMemberOf) = "SPECIALIST REGISTRAR") Then
'                          'if a consultant adds it then its an extra request
'490                       tnode.ForeColor = vbBlue
'500                   End If
'
'510               End If
'
'520               Set CaseNode = tvtemp.SelectedItem
'530               While Not Left(CaseNode.Key, 2) = "L0"
'540                   Set CaseNode = CaseNode.Parent
'550               Wend
'
'                  'if stain is set ot external in the list table then it is always sent away for observation
'560               If frmList.External Then
'
'570                   With frmMovement
'580                       .Update = False
'590                       .MovementId = UniqueId
'600                       .Description = txtDescription
'610                       .Code = txtCode
'620                       .RefType = 1
'630                       .Move frmWorkSheet.Left + frmWorkSheet.fraWorkSheet.Left + frmWorkSheet.SSTabMovement.Left - .Width, frmWorkSheet.Top + frmWorkSheet.SSTabMovement.Top - frmWorkSheet.SSTabMovement.Height
'640                       .Show vbModal
'650                   End With
'660               End If
'
'670           End If
'
'              'expand the tree
'680           .Expanded = True
'
'690           If Not pUpdate Then
'                  'mark the correct node to be selected
'700               Select Case pListType
'                  Case "T"
'710                   tnode.Parent.Selected = True
'720               Case "RS", "SS", "IS"
'730                   tnode.Parent.Parent.Selected = True
'740               Case Else
'750                   tnode.Selected = True
'760               End Select
'770           End If
'
'780       End With
'
'790       DataChanged = True
'800       TreeChanged = True
'
'810       Select Case pListType
'          Case "T"
'820           txtCode = ""
'830           txtDescription = ""
'840           txtDescription.SetFocus
'850       Case "RS", "SS", "IS"
'860           If pLevel = "L2" And InStr(1, tvtemp.SelectedItem.Text, "Block") Then
'870               txtCode = ""
'880               txtDescription = ""
'890               txtDescription.SetFocus
'900           Else
'910               Unload Me
'920           End If
'930       Case Else
'940           Unload Me
'950       End Select
'
'
'960   End If
'
'970   Exit Sub

'cmdCode_Click_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'980   intEL = Erl
'990   strES = Err.Description
'1000  LogError "frmAddCodeTree", "cmdCode_Click", intEL, strES, sql
'
'End Sub

Private Sub cmdCode_Click()
          Dim i As Integer
          Dim UniqueId As String
          Dim sql As String
          Dim tb As New Recordset
          Dim CaseNode As MSComctlLib.Node
          Dim tempListId As String
          Dim strPrevId As String
          Dim strNewId As String
          Dim strLetter As String
          Dim tnode As MSComctlLib.Node
          Dim Tcode As String
          Dim DefaultBlocks As String
          Dim tbStain As Recordset
          
          Tcode = Trim(txtCode.Text)

10        On Error GoTo cmdCode_Click_Error

20        If txtCode <> "" Then

30            With tvtemp.SelectedItem

                  'If adding a T Code
40                If pListType = "T" Then
                      'Check the options table to see if numbering format is numbers or letters (ie. "A : TY4100 : Abdominal Cavity" or "1 : TY4100 : Abdominal Cavity")
50                    If sysOptTissueTypeNumberingFormat(0) = "1" Then
60                        i = .Children + 1
70                    Else
80                        i = .Children + 65
90                    End If

                      'If updating an existing T Code in the tree then update the text
                      'else add to the tree using correct numbering format
100                   If pUpdate Then
110                       ReturnValue = Left(tvtemp.SelectedItem.Text, 1) & " : " & txtCode & " : " & txtDescription
120                       .Text = ReturnValue
130                       strPrevId = .Tag    'Previous T Code

                          'set the tag to be the T Code list id from list table
140                       .Tag = CStr(frmList.ListId)
150                       strNewId = .Tag    'New T Code
160                       strLetter = Left(tvtemp.SelectedItem.Text, 1)    'Tissue Letter

                          'Update the Tissue Code in the grids grdTempMCode grdMCodes
                          'This is because these are used in the Save routines
170                       Call UpdateTissueTypeListId(strPrevId, strNewId, strLetter)

180                   Else
190                       UniqueId = GetUniqueID

200                       If sysOptTissueTypeNumberingFormat(0) = "1" Then
210                           ReturnValue = i & " : " & txtCode & " : " & txtDescription
220                       Else
230                           ReturnValue = Chr(i) & " : " & txtCode & " : " & txtDescription
240                       End If

                          ' Add node at Level 1
250                       Set tnode = tvtemp.Nodes.Add(.Key, tvwChild, "L1" & UniqueId, ReturnValue, 1, 2)
                          tnode.Expanded = True
260                       tnode.Tag = CStr(frmList.ListId)

                          sql = "SELECT DefaultBlocks FROM Lists WHERE Code = '" & Tcode & "'"
                          
                          Set tb = New Recordset
                          RecOpenServer 0, tb, sql
                          
                          If Not tb Is Nothing Then
                             If Not tb.EOF Then
                                DefaultBlocks = tb!DefaultBlocks & ""
                             
                             End If
                          End If
                          
                          Dim k As Integer
                          Dim tbCodes As Recordset
                          k = 1
                          sql = "SELECT StainCode FROM DefaultStains WHERE TissueCode = '" & Tcode & "'"
                          ' Add 10 child nodes of text "Block" under the new L1 node
                          Dim j As Integer
                          Dim L As Integer, n As Integer
                          Dim blockNode As MSComctlLib.Node
                          Dim slideNode As MSComctlLib.Node
                          Dim stainNode As MSComctlLib.Node
                          Dim stainCodes() As String
                          L = 0
275                       Set tbCodes = New Recordset
276                       RecOpenServer 0, tbCodes, sql
                          If Not tbCodes Is Nothing Then
                            Do While Not tbCodes.EOF
                                
                                L = L + 1
                                tbCodes.MoveNext
                            Loop
                          End If
                          ReDim stainCodes(L)
                          L = 0
                          tbCodes.MoveFirst
                          If Not tbCodes Is Nothing Then
                            Do While Not tbCodes.EOF
                                stainCodes(L) = tbCodes!StainCode
                                L = L + 1
                                tbCodes.MoveNext
                            Loop
                          End If
270                       For j = 1 To Val(DefaultBlocks)
271                            UniqueId = GetUniqueID
272                            Set blockNode = tvtemp.Nodes.Add(tnode.Key, tvwChild, "L2" & UniqueId, "Block " & j, 1, 2)
273                            blockNode.Tag = "L2"
274                            blockNode.Expanded = True

'277                            If Not tbCodes Is Nothing Then
'278                                Do While Not tbCodes.EOF
'279                                    UniqueId = GetUniqueID
'280                                    Set slideNode = tvtemp.Nodes.Add(blockNode.Key, tvwChild, "L3" & UniqueId, "Slide " & k, 1, 2)
'281                                    slideNode.Tag = "L3"
'282                                    slideNode.Expanded = True
'283                                    UniqueId = GetUniqueID
'284                                    sql = "SELECT Description FROM Lists WHERE Code = '" & Trim(tbCodes!StainCode) & "'"
'285                                    Set tbStain = New Recordset
'286                                    RecOpenServer 0, tbStain, sql
'287                                    'Set stainNode = tvtemp.Nodes.Add(slideNode.Key, tvwChild, "L4" & UniqueId, "Haematoxylin and Eosin", 1, 2)
'288                                    Set stainNode = tvtemp.Nodes.Add(slideNode.Key, tvwChild, "L4" & UniqueId, Trim(tbStain!Description), 1, 2)
'289                                    stainNode.Tag = "L4"
'299                                    k = k + 1
'291                                    tbCodes.MoveNext
'292                                Loop
'293                            End If
                               k = 1
277                            For n = 0 To UBound(stainCodes) - 1
279                                    UniqueId = GetUniqueID
280                                    Set slideNode = tvtemp.Nodes.Add(blockNode.Key, tvwChild, "L3" & UniqueId, "Slide " & k, 1, 2)
281                                    slideNode.Tag = "L3"
282                                    slideNode.Expanded = True
283                                    UniqueId = GetUniqueID
284                                    sql = "SELECT Description FROM Lists WHERE Code = '" & Trim(stainCodes(n)) & "'"
285                                    Set tbStain = New Recordset
286                                    RecOpenServer 0, tbStain, sql
287                                    'Set stainNode = tvtemp.Nodes.Add(slideNode.Key, tvwChild, "L4" & UniqueId, "Haematoxylin and Eosin", 1, 2)
288                                    Set stainNode = tvtemp.Nodes.Add(slideNode.Key, tvwChild, "L4" & UniqueId, Trim(tbStain!Description), 1, 2)
289                                    stainNode.Tag = "L4"
299                                    k = k + 1
                               Next
                       
294
295
290                       Next j
300                   End If

                      ' Update the rank of the T Code so that it appears
                      ' higher in list the next time T Code is searched
310                   UpdateListRank GetListID(txtCode, pListType)
320               Else
                      ' Adding a stain
330                   ReturnValue = txtDescription

340                   If .Parent.Tag <> "" Then
350                       tempListId = .Parent.Tag
360                   ElseIf .Parent.Parent.Tag <> "" Then
370                       tempListId = .Parent.Parent.Tag
380                   ElseIf .Parent.Parent.Parent.Tag <> "" Then
390                       tempListId = .Parent.Parent.Parent.Tag
400                   End If

410                   sql = "SELECT Code FROM Lists WHERE ListId = '" & tempListId & "' "

420                   Set tb = New Recordset
430                   RecOpenClient 0, tb, sql

                      ' If its level 2 and a Block
440                   If pLevel = "L2" And InStr(1, tvtemp.SelectedItem.Text, "Block") Then
450                       Set tnode = tvtemp.SelectedItem
460                       UniqueId = AddBlockLevelStain(tvtemp.SelectedItem.Key, tnode, ReturnValue)
470                   Else
                          ' Else adding a stain to a slide
480                       UniqueId = GetUniqueID

490                       Set tnode = tvtemp.Nodes.Add(.Key, tvwChild, "L4" & UniqueId, ReturnValue, 1, 2)
500                       tnode.Tag = UniqueId
510                       If (UCase$(UserMemberOf) = "CONSULTANT" Or _
                              UCase$(UserMemberOf) = "SPECIALIST REGISTRAR") Then
                              ' If a consultant adds it then it's an extra request
520                           tnode.ForeColor = vbBlue
530                       End If
540                   End If

550                   Set CaseNode = tvtemp.SelectedItem
560                   While Not Left(CaseNode.Key, 2) = "L0"
570                       Set CaseNode = CaseNode.Parent
580                   Wend

                      ' If stain is set to external in the list table then it is always sent away for observation
590                   If frmList.External Then

600                       With frmMovement
610                           .Update = False
620                           .MovementId = UniqueId
630                           .Description = txtDescription
640                           .Code = txtCode
650                           .RefType = 1
660                           .Move frmWorkSheet.Left + frmWorkSheet.fraWorkSheet.Left + frmWorkSheet.SSTabMovement.Left - .Width, frmWorkSheet.Top + frmWorkSheet.SSTabMovement.Top - frmWorkSheet.SSTabMovement.Height
670                           .Show vbModal
680                       End With
690                   End If

700               End If

                  ' Expand the tree
710               .Expanded = True

720               If Not pUpdate Then
                      ' Mark the correct node to be selected
730                   Select Case pListType
                      Case "T"
740                       tnode.Parent.Selected = True
750                   Case "RS", "SS", "IS"
760                       tnode.Parent.Parent.Selected = True
770                   Case Else
780                       tnode.Selected = True
790                   End Select
800               End If

810           End With

820           DataChanged = True
830           TreeChanged = True

840           Select Case pListType
              Case "T"
850               txtCode = ""
860               txtDescription = ""
870               txtDescription.SetFocus
880           Case "RS", "SS", "IS"
890               If pLevel = "L2" And InStr(1, tvtemp.SelectedItem.Text, "Block") Then
900                   txtCode = ""
910                   txtDescription = ""
920                   txtDescription.SetFocus
930               Else
940                   Unload Me
950               End If
960           Case Else
970               Unload Me
980           End Select

990       End If
          Exit Sub
cmdCode_Click_Error:
      Dim strES As String
      Dim intEL As Integer

1000  intEL = Erl
1010  strES = Err.Description
1020  LogError "frmAddCodeTree", "cmdCode_Click", intEL, strES
End Sub

Private Sub UpdateTissueTypeListId(ByVal strPrev As String, ByVal strNew As String, ByVal strLetter As String)
Dim intGR As Integer

On Error GoTo UpdateTissueTypeListId_Error

'To find and replace the new Tissue code id you need to search for
'Tissue Letter and previous Tissue code id.
'this uniquely identifies the Tissue type in the grids.

For intGR = 1 To frmWorkSheet.grdTempMCode.Rows - 1
    'tissue Letter  PreviousTissueListId
    If UCase(Trim$(frmWorkSheet.grdTempMCode.TextMatrix(intGR, 4))) = UCase(strLetter) And Trim$(frmWorkSheet.grdTempMCode.TextMatrix(intGR, 6)) = Trim$(strPrev) Then
        frmWorkSheet.grdTempMCode.TextMatrix(intGR, 6) = strNew
    End If
Next

'M Code Grid
For intGR = 1 To frmWorkSheet.grdMCodes.Rows - 1
    'tissue Letter  PreviousTissueListId
    If UCase(Trim$(frmWorkSheet.grdMCodes.TextMatrix(intGR, 4))) = UCase(strLetter) And Trim$(frmWorkSheet.grdMCodes.TextMatrix(intGR, 6)) = Trim$(strPrev) Then
        frmWorkSheet.grdMCodes.TextMatrix(intGR, 6) = strNew
    End If
Next

Exit Sub

UpdateTissueTypeListId_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmAddCodeTree", "UpdateTissueTypeListId", intEL, strES

End Sub



Private Sub Form_Load()
frmWorkSheet.Enabled = False

With frmAddCodeTree
    Select Case pListType
        Case "T"
            .Caption = "Histology" & " - " & "Add Tissue Type"
        Case "IS"
            .Caption = "Histology" & " - " & "Add Immuno histo chemical Stain"
        Case "RS"
            .Caption = "Histology" & " - " & "Add Routine Stain"
        Case "SS"
            .Caption = "Histology" & " - " & "Add Special Stain"
        Case Else
            .Caption = "Histology" & " - " & pListTypeName
    End Select
    
End With


'Me.Caption = "Histology --- Add " & pListTypeName

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmWorkSheet.Enabled = True
Unload frmList
Unload Me
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
frmList.PrevCode = txtCode
frmList.PrevDesc = txtDescription
Set frmList.txtCode = txtCode
Set frmList.txtDescription = txtDescription
frmList.SearchByCode = True
frmList.ListType = pListType

frmList.Show
frmList.Move Me.Left + txtCode.Left + 50, Me.Top + txtCode.Top + 625
If KeyAscii = 13 Then
    KeyAscii = 0
End If

End Sub


Private Sub txtDescription_KeyPress(KeyAscii As Integer)
frmList.PrevCode = txtCode
frmList.PrevDesc = txtDescription
Set frmList.txtCode = txtCode
Set frmList.txtDescription = txtDescription
frmList.SearchByCode = False
frmList.ListType = pListType

frmList.Show
frmList.Move Me.Left + txtCode.Left + 50, Me.Top + txtCode.Top + 625
If KeyAscii = 13 Then
    KeyAscii = 0
End If

End Sub

Public Property Let ListType(ByVal Code As String)

pListType = Code

End Property

Public Property Let ListTypeNames(ByVal strNewValue As String)

pListTypeNames = strNewValue

End Property

Public Property Let ListTypeName(ByVal strNewValue As String)

pListTypeName = strNewValue

End Property

Public Property Let Update(ByVal bNewValue As Boolean)

pUpdate = bNewValue

End Property

Public Property Let Level(ByVal strNewValue As String)

pLevel = strNewValue

End Property

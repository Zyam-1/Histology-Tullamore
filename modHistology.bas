Attribute VB_Name = "modHistology"
Option Explicit
Public lngMaxDigits As Long
Public ClearExtraRequests As Boolean
Public sCaseLockedBy As String
Public bLocked As Boolean


Public Sub ValidateLimCaseId(KeyAscii As Integer, ByVal f As Form)
Dim lngSel As Long, lngLen As Long


On Error GoTo ValidateLimCaseId_Error

lngSel = f.txtCaseId.SelStart
lngLen = f.txtCaseId.SelLength
If lngSel = 0 Then
    Select Case KeyAscii
    Case 106, 74
        KeyAscii = 74
    Case 8, 127
    Case Else
        'MsgBox "Format JXXXXXYY, J - Character, XXXXX - Numeric, YY - Year"
        KeyAscii = 0
    End Select
ElseIf lngSel = 1 Then
    Select Case KeyAscii
    Case 112, 80
        KeyAscii = 80
        lngMaxDigits = 12
    Case 48 To 57
        lngMaxDigits = 11
    Case 8, 127
    Case Else
        KeyAscii = 0
    End Select
ElseIf lngSel < lngMaxDigits Then
    Select Case KeyAscii

    Case 48 To 57
        If lngMaxDigits = 12 Then
            If lngSel = 7 Or lngSel = 8 Or lngSel = 9 Then
                f.txtCaseId.Text = Left(f.txtCaseId.Text, 7) & " - "
                lngSel = 10
            End If
        Else

            If lngSel = 6 Or lngSel = 7 Or lngSel = 8 Then
                f.txtCaseId.Text = Left(f.txtCaseId.Text, 6) & " - "
                lngSel = 9
            End If
        End If
    Case 8, 127
    Case Else
        'MsgBox "Format JXXXXXYY, J - Character, XXXXX - Numeric, YY - Year"
        KeyAscii = 0
    End Select
ElseIf KeyAscii <> 8 And KeyAscii <> 127 Then
    KeyAscii = 0
End If

f.txtCaseId.SelStart = lngSel
f.txtCaseId.SelLength = lngLen




Exit Sub

ValidateLimCaseId_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "ValidateLimCaseId", intEL, strES


End Sub

Public Sub ValidateTullCaseId(KeyAscii As Integer, ByVal f As Form)
Dim lngSel As Long, lngLen As Long
Dim TypedChar As String


On Error GoTo ValidateTullCaseId_Error

TypedChar = Chr(KeyAscii)

With f.txtCaseId
    lngSel = .SelStart
    lngLen = .SelLength

    'MsgBox "Format HXXXXX/YY, H - Character, XXXXX - Numeric, YY - Year"
    '"Format CXXXXX/YY, C - Character, XXXXX - Numeric, YY - Year"
    '****Autopsies*****
    '"Format PAXXXXX/YY, PA - Character, XXXXX - Numeric, YY - Year"
    '"Format MAXXXXX/YY, MA - Character, XXXXX - Numeric, YY - Year"
    '"Format TAXXXXX/YY, TA - Character, XXXXX - Numeric, YY - Year"
    If lngSel = 0 Then
'        If Not (UCase(TypedChar) = LS(csHforHistology) Or _
'                UCase(TypedChar) = LS(csCforCytology) Or _
'                UCase(TypedChar) = LS(csAforAutopsy)) Then
'
'            KeyAscii = 0
'        End If
        '        If sysOptCurrentLanguage = "English" Then
        '            Select Case KeyAscii
        '            Case 104, 72
        '                KeyAscii = 72
        '            Case 99, 67
        '                KeyAscii = 67
        '            Case 112, 80
        '                KeyAscii = 80
        '            Case 109, 77
        '                KeyAscii = 77
        '            Case 116, 84
        '                KeyAscii = 84
        '            Case 8, 127
        '            Case Else
        '                KeyAscii = 0
        '            End Select
        '        ElseIf sysOptCurrentLanguage = "Russian" Then
        '            Select Case KeyAscii
        '            Case 227, 195
        '                KeyAscii = 195
        '            Case 246, 214
        '                KeyAscii = 214
        '            Case 226, 194
        '                KeyAscii = 194
        '            Case Else
        '                KeyAscii = 0
        '            End Select
        '        End If
    ElseIf lngSel = 1 Then
        If sysOptCurrentLanguage = "English" Then
            If Left(f.txtCaseId, 1) = "M" Or Left(f.txtCaseId, 1) = "T" Or Left(f.txtCaseId, 1) = "P" Then
                Select Case KeyAscii
                Case 97, 65
                    KeyAscii = 65
                    lngMaxDigits = 12
                Case 8, 127
                Case Else
                    KeyAscii = 0
                End Select
            Else
                Select Case KeyAscii
                Case 48 To 57
                    lngMaxDigits = 11
                Case 8, 127
                Case Else
                    KeyAscii = 0
                End Select
            End If
        End If
    ElseIf lngSel < lngMaxDigits Then
        Select Case KeyAscii
        Case 32
            If lngMaxDigits = 12 Then
                If lngSel = 2 Or lngSel = 3 Or lngSel = 4 _
                   Or lngSel = 5 Or lngSel = 6 Then
                    .Text = Left(.Text, 2) & formatLeadingZero(Int(Val(Mid(.Text, 3, 5))), 5) & " /"
                    lngSel = 10
                End If
            Else

                If lngSel = 1 Or lngSel = 2 Or lngSel = 3 _
                   Or lngSel = 4 Or lngSel = 5 Then
                    .Text = Left(.Text, 1) & formatLeadingZero(Int(Val(Mid(.Text, 2, 5))), 5) & " /"
                    lngSel = 9
                End If
            End If

        Case 48 To 57
            If lngMaxDigits = 12 Then
                If lngSel = 7 Or lngSel = 8 Or lngSel = 9 Then
                    .Text = Left(.Text, 7) & " / "
                    lngSel = 10
                End If
            Else

                If lngSel = 6 Or lngSel = 7 Or lngSel = 8 Then
                    .Text = Left(.Text, 6) & " / "
                    lngSel = 9
                End If
            End If
        Case 8, 127
        Case Else
            KeyAscii = 0
        End Select
    ElseIf KeyAscii <> 8 And KeyAscii <> 127 Then
        KeyAscii = 0
    End If
    .SelStart = lngSel
    .SelLength = lngLen
End With




Exit Sub

ValidateTullCaseId_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "ValidateTullCaseId", intEL, strES


End Sub

Public Function IsValidLimCaseNo(SID As String) As Boolean

Dim Year As String

IsValidLimCaseNo = True
Year = Mid$(SID, lngMaxDigits - 1, 2)

If Not IsNumeric(Year) And Len(Year) < 2 Then
    IsValidLimCaseNo = False
    Exit Function
End If

If Left$(SID, 1) <> "J" Then
    IsValidLimCaseNo = False
    Exit Function
End If

If Mid$(SID, 2, 1) <> "P" And Len(SID) = 12 Then
    IsValidLimCaseNo = False
    Exit Function
End If

If Mid$(SID, 2, 1) = "P" And Len(SID) = 11 Then
    IsValidLimCaseNo = False
    Exit Function
End If

If Len(Mid$(SID, 2, 5)) < 5 Then
    IsValidLimCaseNo = False
    Exit Function
End If

End Function

Public Function IsValidCaseNo(SID As String) As Boolean

IsValidCaseNo = True
If UCase$(sysOptCaseIdValidation(0)) = "TULLAMORE" Then
    IsValidCaseNo = IsValidTullCaseNo(Trim(SID))
Else
    IsValidCaseNo = IsValidLimCaseNo(SID)
End If

End Function

Public Function IsValidTullCaseNo(SID As String) As Boolean

Dim Year As String

IsValidTullCaseNo = True
Year = Mid$(SID, lngMaxDigits - 1, 2)

If Len(Trim$(Year)) <> 2 Then
    IsValidTullCaseNo = False
    Exit Function
End If

If Not IsNumeric(Year) And Len(Year) < 2 Then
    IsValidTullCaseNo = False
    Exit Function
End If

If Val(Year) > Val(Format(Now, "YY")) Then
    IsValidTullCaseNo = False
    Exit Function
End If
'Zyam commented this 6-6-24
If UCase(Left$(SID, 1)) <> "H" And UCase(Left$(SID, 1)) <> "C" _
   And Left$(SID, 1) <> "P" And Left$(SID, 1) <> "M" _
   And Left$(SID, 1) <> "T" Then
    IsValidTullCaseNo = False
    Exit Function
End If
'
If (Left$(SID, 1) = "H" Or Left$(SID, 1) = "C") _
   And Mid$(SID, 2, 1) = "A" Then
    IsValidTullCaseNo = False
    Exit Function
End If
'Zyam

If UCase(Mid$(SID, 2, 1)) <> "A" And Len(SID) = 12 Then
    IsValidTullCaseNo = False
    Exit Function
End If

If UCase(Mid$(SID, 2, 1)) = "A" And Len(SID) = 11 Then
    IsValidTullCaseNo = False
    Exit Function
End If

If Len(Mid$(SID, 2, 5)) < 5 Then
    IsValidTullCaseNo = False
    Exit Function
End If

End Function

Public Function CalcAge(ByVal DoB As String, ByVal SampleDate As Date) As String
Dim diff As Long
Dim DobYr As Single

On Error GoTo CalcAge_Error

DoB = Format$(DoB, "dd/mm/yyyy")
If IsDate(DoB) Then

    diff = DateDiff("d", (DoB), SampleDate)

    DobYr = diff / 365.25
    If DobYr > 1 Then
        CalcAge = Format$(Int(DobYr), "###\Yr")
    ElseIf diff < 30.43 Then
        CalcAge = Format$(diff, "##\D")
    Else
        CalcAge = Format$(Int(diff / 30.43), "##\M")
    End If
Else
    CalcAge = ""
End If

Exit Function

CalcAge_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "CalcAge", intEL, strES


End Function

Public Function formatLeadingZero(intNumber, numDigits)
Dim strFormattedNumber As String
Dim X As Integer

numDigits = numDigits - Len(intNumber)

For X = 1 To numDigits
    strFormattedNumber = strFormattedNumber & "0"
Next

strFormattedNumber = strFormattedNumber & intNumber

formatLeadingZero = strFormattedNumber
End Function

Public Function GetBlockNumber(oNode As MSComctlLib.Node) As Integer
Dim oChildNode As MSComctlLib.Node


If Not (oNode Is Nothing) Then
    If oNode.Children Then
        Set oNode = oNode.Child
        Do Until oNode Is Nothing

            If InStr(1, oNode.Text, "Block") Then
                If sysOptBlockNumberingFormat(0) = "1" Then
                    If CInt(Mid(oNode.Text, InStr(oNode.Text, " "))) > GetBlockNumber Then
                        GetBlockNumber = CInt(Mid(oNode.Text, InStr(oNode.Text, " ")))
                    End If
                Else
                    GetBlockNumber = GetBlockNumber + 1
                End If
            ElseIf InStr(1, oNode.Text, "Frozen Section") Then
                Set oChildNode = oNode.Child
                Do Until oChildNode Is Nothing
                    If InStr(oNode.Text, "Block") Then
                        If sysOptBlockNumberingFormat(0) = "1" Then
                            If CInt(Mid(oChildNode.Text, InStr(oChildNode.Text, " "))) > GetBlockNumber Then
                                GetBlockNumber = CInt(Mid(oChildNode.Text, InStr(oChildNode.Text, " ")))
                            End If
                        Else
                            GetBlockNumber = GetBlockNumber + 1
                        End If
                    End If
                    Set oChildNode = oChildNode.Next
                Loop
            End If

            Set oNode = oNode.Next
        Loop
    Else
        GetBlockNumber = 0
    End If
End If



End Function

Public Function GetFrozenSectionNumber(oNode As MSComctlLib.Node) As Integer
If Not (oNode Is Nothing) Then
    If oNode.Children Then
        Set oNode = oNode.Child
        Do Until oNode Is Nothing

            If InStr(1, oNode.Text, "Frozen Section") Then
                If CInt(Mid(oNode.Text, InStrRev(oNode.Text, " "))) > GetFrozenSectionNumber Then
                    GetFrozenSectionNumber = CInt(Mid(oNode.Text, InStrRev(oNode.Text, " ")))
                End If
            End If

            Set oNode = oNode.Next
        Loop
    Else
        GetFrozenSectionNumber = 0
    End If
End If
End Function

Public Function GetSlideNumber(oNode As MSComctlLib.Node) As Integer
If Not (oNode Is Nothing) Then
    If oNode.Children Then
        Set oNode = oNode.Child
        Do Until oNode Is Nothing
            If InStr(1, oNode.Text, "Slide") Then
                GetSlideNumber = GetSlideNumber + 1
            End If

            Set oNode = oNode.Next
        Loop
    Else
        GetSlideNumber = 0
    End If
End If
End Function

Public Sub AddDefaultStains(Key As String, tempNode As MSComctlLib.Node)
Dim sql As String
Dim tb As New Recordset
Dim tnode As MSComctlLib.Node
Dim i As Integer
Dim UniqueId As String
Dim TempId As String




On Error GoTo AddDefaultStains_Error

sql = "SELECT l.Description, d.StainCode, l.Levels FROM DefaultStains d " & _
      "INNER JOIN Lists l ON d.staincode = l.Code " & _
      "WHERE d.TissueCodeListId = '" & tempNode.Tag & "' " & _
      "AND (l.ListType = 'IS' OR l.ListType = 'RS' OR l.ListType = 'SS')"

Set tb = New Recordset
RecOpenServer 0, tb, sql



If Not tb.EOF Then
    i = 1
    With tempNode
        Do Until tb.EOF
            UniqueId = GetUniqueID
            If sysOptSlideNumberingFormat(0) = "1" Then
                Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(Key, tvwChild, "L3" & UniqueId, "Slide" & " " & i, 1, 2)
                tnode.Expanded = True
                TempId = GetUniqueID
                Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add("L3" & UniqueId, tvwChild, "L4" & TempId, tb!Description & "", 1, 2)
            Else
                Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(Key, tvwChild, "L3" & UniqueId, "Slide" & " " & AddLetter(i), 1, 2)
                tnode.Expanded = True
                TempId = GetUniqueID
                Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add("L3" & UniqueId, tvwChild, "L4" & TempId, tb!Description & "", 1, 2)
            End If
            tnode.Expanded = True
            tb.MoveNext
            i = i + 1
        Loop
    End With
    AddLevels Key, tempNode, i
End If



Exit Sub

AddDefaultStains_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "AddDefaultStains", intEL, strES, sql


End Sub
Public Sub AddLevels(Key As String, tempNode As MSComctlLib.Node, i As Integer)

Dim tnode As MSComctlLib.Node
Dim j As Integer
Dim UniqueId As String
Dim sql As String
Dim sn As Recordset
Dim tb As Recordset
Dim Levels As Integer
Dim TempId As String

On Error GoTo AddLevels_Error

sql = "SELECT * FROM Lists WHERE ListId = '" & tempNode.Tag & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
    Levels = Val(tb!Levels & "")
    sql = "SELECT * FROM Lists WHERE Code = 'H&E' AND ListType = 'RS'"
    Set sn = New Recordset
    RecOpenServer 0, sn, sql

    If Not sn.EOF Then
        If Levels > 0 Then
            For j = 1 To Levels
                UniqueId = GetUniqueID

                If sysOptSlideNumberingFormat(0) = "1" Then
                    Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(Key, tvwChild, "L3" & UniqueId, "Slide" & " " & i, 1, 2)
                    tnode.Expanded = True
                    TempId = GetUniqueID
                    Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add("L3" & UniqueId, tvwChild, "L4" & TempId, sn!Description & "", 1, 2)
                Else
                    Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(Key, tvwChild, "L3" & UniqueId, "Slide" & " " & AddLetter(i), 1, 2)
                    tnode.Expanded = True
                    TempId = GetUniqueID
                    Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add("L3" & UniqueId, tvwChild, "L4" & TempId, sn!Description & "", 1, 2)
                End If
                i = i + 1
            Next j
        End If
    End If
End If


Exit Sub

AddLevels_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "AddLevels", intEL, strES, sql


End Sub
Public Function AddBlockLevelStain(Key As String, ByRef tnode As MSComctlLib.Node, Description As String) As String
Dim i As Integer
Dim UniqueId As String
Dim TempId As String

i = GetSlideNumber(frmWorkSheet.tvCaseDetails.SelectedItem)

UniqueId = GetUniqueID

If sysOptSlideNumberingFormat(0) = "1" Then
    i = i + 1
    Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(Key, tvwChild, "L3" & UniqueId, "Slide" & " " & i, 1, 2)
    tnode.Expanded = True
    If (UCase$(UserMemberOf) = "CONSULTANT" Or _
        UCase$(UserMemberOf) = "SPECIALIST REGISTRAR") Then
        tnode.ForeColor = vbBlue
    End If
    TempId = GetUniqueID
    Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add("L3" & UniqueId, tvwChild, "L4" & TempId, Description, 1, 2)
    tnode.Tag = TempId
    If (UCase$(UserMemberOf) = "CONSULTANT" Or _
        UCase$(UserMemberOf) = "SPECIALIST REGISTRAR") Then
        tnode.ForeColor = vbBlue
    End If
Else
    Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add(Key, tvwChild, "L3" & UniqueId, "Slide" & " " & AddLetter(i), 1, 2)
    tnode.Expanded = True
    If (UCase$(UserMemberOf) = "CONSULTANT" Or _
        UCase$(UserMemberOf) = "SPECIALIST REGISTRAR") Then
        tnode.ForeColor = vbBlue
    End If
    TempId = GetUniqueID
    Set tnode = frmWorkSheet.tvCaseDetails.Nodes.Add("L3" & UniqueId, tvwChild, "L4" & TempId, Description, 1, 2)
    tnode.Tag = TempId
    If (UCase$(UserMemberOf) = "CONSULTANT" Or _
        UCase$(UserMemberOf) = "SPECIALIST REGISTRAR") Then
        tnode.ForeColor = vbBlue
    End If
End If
AddBlockLevelStain = TempId
tnode.Expanded = True
End Function

Public Function AddLetter(ByVal Number As Integer) As String
'65-90
Dim s As String


On Error GoTo AddLetter_Error

If Number >= 26 And Number < 52 Then
    Number = Number - 26
    s = Chr$(65) & Chr$(Number + 65)
ElseIf Number >= 52 And Number < 78 Then
    Number = Number - 52
    s = Chr$(66) & Chr$(Number + 65)
ElseIf Number >= 78 And Number < 104 Then
    Number = Number - 78
    s = Chr$(67) & Chr$(Number + 65)
ElseIf Number >= 104 And Number < 130 Then
    Number = Number - 104
    s = Chr$(68) & Chr$(Number + 65)
ElseIf Number >= 130 And Number < 156 Then
    Number = Number - 130
    s = Chr$(69) & Chr$(Number + 65)
ElseIf Number >= 156 And Number < 182 Then
    Number = Number - 156
    s = Chr$(70) & Chr$(Number + 65)
ElseIf Number >= 182 And Number < 200 Then
    Number = Number - 182
    s = Chr$(71) & Chr$(Number + 65)
Else
    s = Chr$(Number + 65)
End If


AddLetter = s



Exit Function

AddLetter_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "AddLetter", intEL, strES


End Function

Public Sub ChangeTabCaptionColour(T As SSTab, pic As PictureBox, _
                                  colour As Long, Caption As String, _
                                  ByVal tabIndex As Integer)
On Error GoTo ChangeTabCaptionColour_Error

pic.AutoRedraw = True
pic.BorderStyle = 0
pic.Width = pic.TextWidth(Caption)
pic.Height = pic.TextHeight(Caption) * 2
pic.Cls
pic.ForeColor = colour
pic.CurrentX = 0
pic.CurrentY = pic.TextHeight(Caption) / 2
pic.Print Caption
T.TabCaption(tabIndex) = ""
T.TabPicture(tabIndex) = pic.Image

Exit Sub

ChangeTabCaptionColour_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "ChangeTabCaptionColour", intEL, strES

End Sub

Public Function CheckReferralDiscrep(RecordId As String) As Boolean
Dim sql As String
Dim tb As Recordset

On Error GoTo CheckReferralDiscrep_Error

sql = "SELECT * FROM CaseMovementDetails " & _
      "WHERE CaseListId = N'" & RecordId & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
    If tb!BlocksSent & "" <> tb!BlocksRet & "" Or _
       tb!HESent & "" <> tb!HERet & "" Or _
       tb!ImmunoSent & "" <> tb!ImmunoRet & "" Or _
       tb!SpecialSent & "" <> tb!SpecialRet & "" Or _
       tb!UnstainedSent & "" <> tb!UnstainedRet & "" Then

        CheckReferralDiscrep = True
    Else
        CheckReferralDiscrep = False
    End If
Else
    CheckReferralDiscrep = False
End If

Exit Function

CheckReferralDiscrep_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "CheckReferralDiscrep", intEL, strES, sql


End Function

Public Sub LockCase(sCaseId As String)

Dim sql As String
Dim MachineName As String


On Error GoTo LockCase_Error

MachineName = UCase$(vbGetComputerName())

sql = "Delete FROM CasesLocked WHERE username = N'" & AddTicks(UserName) & "' " & _
      "and MachineName = N'" & MachineName & "'"
Cnxn(0).Execute sql

sql = "INSERT INTO CasesLocked " & _
      "(CaseId, Username, MachineName) VALUES " & _
      "(N'" & sCaseId & "', " & _
      " N'" & AddTicks(UserName) & "', " & _
      " N'" & MachineName & "')"

Cnxn(0).Execute sql



Exit Sub

LockCase_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "LockCase", intEL, strES, sql


End Sub

Public Sub UnlockCase()

Dim sql As String
Dim MachineName As String

On Error GoTo UnlockCase_Error

MachineName = UCase$(vbGetComputerName())

sql = "Delete FROM CasesLocked WHERE username = N'" & AddTicks(UserName) & "' " & _
      "and MachineName = N'" & MachineName & "'"

Cnxn(0).Execute sql


Exit Sub

UnlockCase_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "UnlockCase", intEL, strES, sql


End Sub

Public Function CaseLockedBy(sCaseId As String) As String

Dim sql As String
Dim tb As Recordset

On Error GoTo CaseLockedBy_Error

sql = "SELECT * FROM CasesLocked " & _
      "WHERE CaseId = N'" & sCaseId & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If tb.EOF Then
    CaseLockedBy = ""
Else
    CaseLockedBy = tb!UserName & ""
End If

Exit Function

CaseLockedBy_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "CaseLockedBy", intEL, strES, sql


End Function

' WorkDays
' returns the number of working days between two dates
Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long

Dim dtFirstSunday As Date
Dim dtLastSaturday As Date
Dim lngWorkDays As Long

' get first sunday in range
dtFirstSunday = dtBegin + ((8 - Weekday(dtBegin)) Mod 7)

' get last saturday in range
dtLastSaturday = dtEnd - (Weekday(dtEnd) Mod 7)

' get work days between first sunday and last saturday
lngWorkDays = (((dtLastSaturday - dtFirstSunday) + 1) / 7) * 5

' if first sunday is not begin date
If dtFirstSunday <> dtBegin Then

    ' assume first sunday is after begin date
    ' add workdays from begin date to first sunday
    lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

End If

' if last saturday is not end date
If dtLastSaturday <> dtEnd Then

    ' assume last saturday is before end date
    ' add workdays from last saturday to end date
    lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

End If

' return working days
WorkDays = lngWorkDays

End Function



Public Function AuthorisedOrAddendumAdded(sCaseId As String) As String

Dim sql As String
Dim tb As Recordset

On Error GoTo AuthorisedOrAddendumAdded_Error

AuthorisedOrAddendumAdded = ""

sql = "SELECT Validated, OrigValDate, AddendumAdded FROM Cases " & _
      "WHERE CaseId = N'" & sCaseId & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
    'Not Authorised
    If ((tb!Validated = False) Or (tb!OrigValDate & "" = "")) And tb!AddendumAdded = False Then
        AuthorisedOrAddendumAdded = "1"
        'Authorised and No addendum added
    ElseIf ((tb!Validated = True) Or (tb!OrigValDate & "" <> "")) And tb!AddendumAdded = False Then
        AuthorisedOrAddendumAdded = "2"
        'Authorised and Addendum added
    ElseIf ((tb!Validated = True) Or (tb!OrigValDate & "" <> "")) And tb!AddendumAdded = True Then
        AuthorisedOrAddendumAdded = "3"
    End If
End If

Exit Function

AuthorisedOrAddendumAdded_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "AuthorisedOrAddendumAdded", intEL, strES, sql

End Function

Public Function DaysHavePassedSinceAuthorisation(ByVal strCaseId As String, ByVal intDays As Integer) As Boolean

Dim sql As String
Dim tb As Recordset
Dim dateNow As Date
Dim dateAuth As Date

On Error GoTo DaysHavePassedSinceAuthorisation_Error

DaysHavePassedSinceAuthorisation = False

strCaseId = UCase(Replace(Replace(strCaseId, " " & sysOptCaseIdSeperator(0) & " ", ""), " ", ""))

sql = "SELECT OrigValDate FROM Cases " & _
      "WHERE CaseId = N'" & strCaseId & "' AND OrigValDate is not null"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
    dateAuth = Format(tb!OrigValDate, "dd/mmm/yyyy")

    dateNow = Format(Now, "dd/mmm/yyyy")

    If DateDiff("d", dateAuth, dateNow) > intDays Then
        DaysHavePassedSinceAuthorisation = True
    End If
End If


Exit Function

DaysHavePassedSinceAuthorisation_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "DaysHavePassedSinceAuthorisation", intEL, strES, sql

End Function



Public Function CaseAllDisposed(ByVal strCaseId As String) As Boolean

Dim sql As String
Dim tb As Recordset

On Error GoTo CaseAllDisposed_Error

CaseAllDisposed = False

strCaseId = UCase(Replace(Replace(strCaseId, " " & sysOptCaseIdSeperator(0) & " ", ""), " ", ""))

sql = "SELECT Disposal FROM CaseTree " & _
      "WHERE CaseId = N'" & strCaseId & "' and tissueTypeListId is not null and (Disposal <> 'D' or Disposal is null)"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If tb.EOF Then
    CaseAllDisposed = True
End If

Exit Function

CaseAllDisposed_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "CaseAllDisposed", intEL, strES, sql

End Function


Public Function CaseLinked2AuthCytoCase(ByVal strCaseId As String, ByVal strState As String) As Boolean

Dim sql As String
Dim tb As Recordset

On Error GoTo CaseLinked2AuthCytoCase_Error

CaseLinked2AuthCytoCase = False
If UCase(Left(strCaseId, 1)) = strDeptLetter4Histo And strState = 2 Then
    sql = "SELECT LinkedCaseId FROM Cases " & _
          "WHERE CaseId = N'" & strCaseId & "' and LinkedCaseId like N'" & "C" & "%' "
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    If Not tb.EOF Then
        If IsLinkedCytoAuthorised(tb!LinkedCaseId) Then    'Is corresponding Cyto Authorised
            CaseLinked2AuthCytoCase = True
        End If
    End If

End If

Exit Function

CaseLinked2AuthCytoCase_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "CaseLinked2AuthCytoCase", intEL, strES, sql

End Function

Public Function IsLinkedCytoAuthorised(ByVal strCaseId As String) As Boolean

Dim sql As String
Dim tb As Recordset

On Error GoTo IsLinkedCytoAuthorised_Error

IsLinkedCytoAuthorised = False

sql = "SELECT Validated FROM Cases " & _
      "WHERE CaseId = N'" & strCaseId & "' and Validated = 1"
Set tb = New Recordset
RecOpenServer 0, tb, sql

If Not tb.EOF Then
    IsLinkedCytoAuthorised = True
End If

Exit Function

IsLinkedCytoAuthorised_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modHistology", "IsLinkedCytoAuthorised", intEL, strES, sql

End Function


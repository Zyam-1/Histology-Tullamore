Attribute VB_Name = "modPrintHistology"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const EM_GETLINECOUNT = &HBA
Private Validated As Boolean
Private LinesAllowed As Integer
Private pStrValidatedBy As String
Private pStrValReportDate As String
Private pStrFinalValidatedBy As String
Private pStrFinalValReportDate As String

Public Function GetPrintDetails()

    Dim sql As String
    Dim tb As Recordset
    Dim sn As Recordset
    Dim ca As Recordset
    Dim rs As Recordset
    Dim cl As Recordset
    Dim Amendments As String
    Dim Gross As String
    Dim Micro As String
    Dim ClinicalHistory As String
    Dim NOS As String
    Dim ContainerLabel As String
    Dim Heading As String
    Dim Snomed As String
    Dim PrevLetter As String
    Dim strReportTypeQ As String
    Dim strAuthorisedDateTimePostQcodeAdded As String
    Dim strAuthorisedUsernamePostQcodeAdded As String
                
10    On Error GoTo GetPrintDetails_Error

20    pStrValidatedBy = ""
30    pStrValReportDate = ""
40    pStrFinalValidatedBy = ""
50    pStrFinalValReportDate = ""

60    sql = "SELECT * " & _
          "FROM Demographics " & _
          "WHERE CaseId = '" & CaseNo & "' "
70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql

90    If Not tb.EOF Then
100     If tb!ClinicalHistory & "" <> "" Then
110         ClinicalHistory = "CLINICAL DETAILS:" & vbCr & Replace(tb!ClinicalHistory, vbCrLf, vbCr)    '& vbCr
120     End If
130     If tb!NatureOfSpecimen & "" <> "" Then
140         NOS = "NATURE OF SPECIMEN:" & vbCr & Replace(tb!NatureOfSpecimen, vbCrLf, vbCr)    '& vbCr
150     End If
160     If tb!SpecimenLabelled & "" <> "" Then
170         ContainerLabel = "CONTAINER LABELLED:" & vbCr & Replace(tb!SpecimenLabelled, vbCrLf, vbCr)    '& vbCr
180     End If
190   Else
200     ClinicalHistory = ""
210     NOS = ""
220     ContainerLabel = ""
230   End If

240   LinesAllowed = 44

250   sql = "SELECT * FROM Cases WHERE " & _
          "CaseId = '" & CaseNo & "' "
260   Set sn = New Recordset
270   RecOpenClient 0, sn, sql

280   If Not sn.EOF Then
290     Validated = sn!Validated

300     If Not Validated Then
310         Heading = "***** THIS IS A PRELIMINARY REPORT ONLY *****"    '& vbCr
320     End If

330     frmRichText.tempRtb.TextRTF = sn!Gross & ""
340     frmRichText.tempRtb.SelStart = 0
350     frmRichText.tempRtb.SelLength = Len(frmRichText.tempRtb.Text)
360     frmRichText.tempRtb.SelFontName = "Courier New"
370     frmRichText.tempRtb.SelFontSize = 10
380     If Trim(frmRichText.tempRtb.Text) <> "" Then
390         If Left(CaseNo, 1) = "C" Then
400             Gross = "GROSS APPEARANCE:" & vbCr & Replace(frmRichText.tempRtb.Text, vbCrLf, vbCr)    '& vbCr
410         ElseIf Mid(CaseNo, 2, 1) = "P" Or Mid(CaseNo, 2, 1) = "A" Then
420             Gross = "AUTOPSY REPORT:" & vbCr & Replace(frmRichText.tempRtb.Text, vbCrLf, vbCr)
430         Else
440             Gross = "GROSS EXAMINATION:" & vbCr & Replace(frmRichText.tempRtb.Text, vbCrLf, vbCr)    '& vbCr
450         End If
460     Else
470         Gross = ""
480     End If
490     frmRichText.tempRtb.TextRTF = sn!Micro & ""
500     frmRichText.tempRtb.SelStart = 0
510     frmRichText.tempRtb.SelLength = Len(frmRichText.tempRtb.Text)
520     frmRichText.tempRtb.SelFontName = "Courier New"
530     frmRichText.tempRtb.SelFontSize = 10
540     If frmRichText.tempRtb.Text <> "" Then
550         Micro = "MICROSCOPIC REPORT:" & vbCr & Replace(frmRichText.tempRtb.Text, vbCrLf, vbCr)    '& vbCr
560     Else
570         Micro = ""
580     End If
590     pStrValidatedBy = sn!OrigValBy & ""
600     pStrValReportDate = sn!OrigValDate & ""
610     If sn!OrigValDate & "" <> sn!ValReportDate & "" Then    'If the same then only Authorised once
620         pStrFinalValidatedBy = sn!ValidatedBy & ""       'So don't display
630         pStrFinalValReportDate = sn!ValReportDate & ""
640     End If
650   Else
660     Gross = ""
670     Micro = ""
680   End If

690   sql = "SELECT c.comment,c.username,c.datetimeofrecord,l.description FROM CaseAmendments c " & _
          "INNER JOIN CaseListLink cl ON c.caselistid = cl.caselistid " & _
          "INNER JOIN Lists l ON l.Listid = cl.listid " & _
          "WHERE c.CaseId = '" & CaseNo & "' order by c.datetimeofrecord "    'see ITS 818948 (1)
700   Set ca = New Recordset
710   RecOpenClient 0, ca, sql

720   If Not ca.EOF Then
730     Heading = "***** THIS REPORT HAS BEEN UPDATED (SEE BELOW) *****"    '& vbCr
740     Amendments = Amendments & "UPDATED REPORT:"    '& vbCr
750     Do Until ca.EOF

760         Amendments = Amendments & vbCr
770         Amendments = Amendments & ca!Description & ""
780         Amendments = Amendments & vbCr
790         strReportTypeQ = ""
800         Select Case UCase(Trim$(ca!Description & ""))
            Case "CORRECTED REPORTS": strReportTypeQ = "Corrected"
810         Case "SUPPLEMENTARY REPORTS": strReportTypeQ = "Supplementary"
820         Case "AMENDED REPORTS": strReportTypeQ = "Updated"
830         End Select
            
            'ITS: 819267
            'get DateTime and user who authorised Case post adding of Q codes
840         strAuthorisedDateTimePostQcodeAdded = "" 'default
850         strAuthorisedUsernamePostQcodeAdded = "" 'default
                'Caseid , DateTime Ammendement made, Authorising user, Authorised Date
860         Call getQcodeAuthorisationDetails(CaseNo, Format(ca!DateTimeOfRecord, "dd/mmm/yyyy hh:mm"), _
                strAuthorisedUsernamePostQcodeAdded, strAuthorisedDateTimePostQcodeAdded)
            
870         Amendments = Amendments & FormatString(strReportTypeQ & " Report Authorised Date: ", 38)
880         Amendments = Amendments & FormatString(Format(strAuthorisedDateTimePostQcodeAdded, "dd/mm/yyyy hh:mm"), 16)
890         Amendments = Amendments & vbCr
900         Amendments = Amendments & FormatString(strReportTypeQ & " Report Authorised By: ", 38)
910         Amendments = Amendments & Trim$(strAuthorisedUsernamePostQcodeAdded)

920         Amendments = Amendments & vbCr
930         Amendments = Amendments & Trim(Replace(ca!Comment, "<<tab>>", vbTab))
940         Amendments = Amendments & vbCr
950         Amendments = Amendments & "PRINTDASH"
960         ca.MoveNext
970     Loop
980   Else
990     Amendments = ""
1000  End If

'Zyam Commented this remove TissueType from Printing 04-08-24
'1010  sql = "SELECT cl.TissueTypeId AS Tid,cl.TissueTypeLetter AS TLetter, " & _
'           "(SELECT a.Code FROM Lists a WHERE cl.TissueTypeListId = a.ListId) AS TCode, " & _
'           "(SELECT a.Description FROM Lists a WHERE cl.TissueTypeListId = a.ListId) AS TCodeDescription, " & _
'           "l.Code as MCode, l.Description as MCodeDescription " & _
'           "FROM CaseListLink cl " & _
'           "INNER JOIN Lists l on l.ListId = cl.ListId " & _
'           "AND cl.CaseId = '" & CaseNo & "' " & _
'           "AND (cl.Type = 'M') " & _
'           "ORDER BY cl.TissueTypeLetter"
'
'1020  Set cl = New Recordset
'1030  RecOpenClient 0, cl, sql
'
'1040  If Not cl.EOF Then
'1050    Snomed = Snomed & "SNOMED:"
'1060    Do Until cl.EOF
'
'1070        If PrevLetter <> cl!TLetter Then
'1080            PrevLetter = cl!TLetter
'1090            Snomed = Snomed & vbCr
'1100            Snomed = Snomed & cl!TLetter & " - " & cl!Tcode & " - " & cl!TCodeDescription
'1110        End If
'
'1120        Snomed = Snomed & vbCr
'1130        Snomed = Snomed & Space(10)
'1140        Snomed = Snomed & cl!MCode & " - " & cl!MCodeDescription
'1150        cl.MoveNext
'1160    Loop
'1170  Else
'1180    Snomed = ""
'1190  End If
'Zyam Commented this remove TissueType from Printing 04-08-24


1200  ClearUdtHeading
1210  With udtHeading
1220    .SampleID = CaseNo
1230    .Dept = "Histology"
1240    .Name = tb!PatientName & ""
1250    .Ward = tb!Ward & ""
1260    .DoB = tb!DateOfBirth
1270    .Chart = tb!MRN & ""
1280    .Clinician = tb!Clinician & ""
1290    .Address0 = tb!Address1 & ""
1300    .Address1 = tb!Address2 & ""
1310    .County = tb!County & ""
1320    .GP = tb!GP & ""
1330    If tb!GpId & "" <> "" Then
1340        sql = "SELECT Address1 AS GPAd1,Address2 AS GPAd2, " & _
                  "County AS GPCo FROM GPs " & _
                  "WHERE GPId = '" & tb!GpId & "' "
1350        Set rs = New Recordset
1360        RecOpenClient 0, rs, sql

1370        If Not rs.EOF Then
1380            .GPAddress1 = Trim(rs!GPAd1 & "")
1390            .GPAddress2 = Trim(rs!GPAd2 & "")
1400            .GPCounty = Trim(rs!GPCo & "")
1410        End If
1420    End If
1430    .Hospital = tb!Source & ""
'commented by ali
'1440    .Sex = tb!Sex & ""
'ali
If Trim(tb!Sex) = "M" Then
    .Sex = "MALE"
ElseIf Trim(tb!Sex) = "F" Then
    .Sex = "FEMALE"
Else
    .Sex = ""
End If
'---------
1450    .SampleDate = sn!SampleTaken & ""
1460    .RecDate = sn!SampleReceived & ""
1470    .SampleType = ""
1480    .DateOfDeath = tb!DateOfDeath & ""
1490    .Coroner = tb!AutopsyRequestedBy & ""

1500    If Left(CaseNo, 1) = strDeptLetter4Histo Then    'Histology Case
1510        .DocomentNo = GetOptionSetting("HistologyDocumentNo", "")
1520        .AccreditationText = GetOptionSetting("HistologyAccreditationText", "")
1530    ElseIf Left(CaseNo, 1) = "C" Then    'Cytology Case
1540        .DocomentNo = GetOptionSetting("CytologyDocumentNo", "")
1550        .AccreditationText = GetOptionSetting("CytologyAccreditationText", "")
1560    ElseIf Mid(CaseNo, 2, 1) = "A" Then    'Autopsy Case
1570        .DocomentNo = GetOptionSetting("AutopsyDocumentNo", "")
1580        .AccreditationText = GetOptionSetting("AutopsyAccreditationText", "")
1590    End If
1600

1610  End With



1620  GetPrintDetails = Trim(IIf(Trim(Heading = ""), "", Heading & vbCr) & _
                            IIf(Trim(NOS = ""), "", NOS & vbCr) & _
                            IIf(Trim(ContainerLabel = ""), "", ContainerLabel & vbCr) & _
                            IIf(Trim(ClinicalHistory = ""), "", ClinicalHistory & vbCr) & _
                            IIf(Trim(Gross = ""), "", Gross & vbCr) & _
                            IIf(Trim(Micro = ""), "", Micro & vbCr) & _
                            IIf(Trim(Amendments = ""), "", Amendments & vbCr) & _
                            IIf(Trim(Snomed = ""), "", Snomed))


1630  Exit Function

GetPrintDetails_Error:

    Dim strES As String
    Dim intEL As Integer

1640  intEL = Erl
1650  strES = Err.Description
1660  LogError "modPrintHistology", "GetPrintDetails", intEL, strES, sql

End Function

Private Sub getQcodeAuthorisationDetails(ByVal strCaseNo As String, ByVal strQcodePostDate As String, _
    ByRef strAuthorisedUsernamePostQcodeAdded As String, ByRef strAuthorisedDateTimePostQcodeAdded As String)

      Dim sql As String
      Dim tb As Recordset

10       On Error GoTo getQcodeAuthorisationDetails_Error

20    strAuthorisedUsernamePostQcodeAdded = ""
30    strAuthorisedDateTimePostQcodeAdded = ""

40    sql = "SELECT TOP (1) UserName, DateTimeOfRecord From CaseEventLog WHERE " & _
      "CaseId = '" & strCaseNo & "' AND DateTimeOfRecord > '" & strQcodePostDate & "' AND EventDesc = 'Authorised' " & _
      "ORDER BY DateTimeOfRecord"

50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql

70    If Not tb.EOF Then
80        strAuthorisedUsernamePostQcodeAdded = tb!UserName & ""
90        strAuthorisedDateTimePostQcodeAdded = Format(tb!DateTimeOfRecord, "dd/mm/yyyy hh:mm")
100   End If
110   tb.Close

120      Exit Sub

getQcodeAuthorisationDetails_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "modPrintHistology", "getQcodeAuthorisationDetails", intEL, strES, sql

End Sub

Public Sub PrintHistology(Optional SendCopyTo As String, _
                          Optional ViewOnly As Boolean, _
                          Optional CopyNo As String = "")
    Dim sql As String

10    On Error GoTo PrintHistology_Error

20    ReDim pl(1 To 1) As String
    Dim plCounter As Long
    Dim crPos As Integer
    Dim HR As String

    Dim TotalPages As Integer
    Dim ThisPage As Integer
    Dim TopLine As Integer
    Dim BottomLine As Integer
    Dim crlfFound As Boolean
    Dim n As Integer
    Dim PrintTime As String
    Dim CurrentTime As String
    Dim rtfTemp As String
    Dim YYYY As String

30    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")


40    If GetPrintDetails = "" Then
50      Exit Sub
60    Else
70      HR = GetPrintDetails
80    End If
90    crlfFound = True
100   Do While crlfFound
110     HR = RTrim(HR)
120     crlfFound = False
130     If Right(HR, 1) = vbCr Or Right(HR, 1) = vbLf Then
140         HR = Left(HR, Len(HR) - 1)
150         crlfFound = True
160     End If
170   Loop

180   plCounter = 0
190   Do While Len(HR) > 0
200     crPos = InStr(HR, vbCr)
210     If crPos > 0 And crPos < 81 Then
220         plCounter = plCounter + 1
230         ReDim Preserve pl(1 To plCounter)
240         pl(plCounter) = Left(HR, crPos - 1)
250         HR = Mid(HR, crPos + 1)
260     Else
270         If Len(HR) > 81 Then
280             For n = 81 To 1 Step -1
290                 If Mid(HR, n, 1) = " " Then
300                     Exit For
310                 End If
320             Next
330             plCounter = plCounter + 1
340             ReDim Preserve pl(1 To plCounter)
350             If n = 0 Then
360                 n = 81
370             End If
380             pl(plCounter) = Left(HR, n)
390             HR = Mid(HR, n + 1)
400         Else
410             plCounter = plCounter + 1
420             ReDim Preserve pl(1 To plCounter)
430             pl(plCounter) = HR
440             Exit Do
450         End If
460     End If
470   Loop



480   CurrentTime = Format(Now, "ddMMyyyyhhmmss")

490   TotalPages = Int((plCounter - 1) / LinesAllowed) + 1
500   If TotalPages = 0 Then TotalPages = 1

510   YYYY = 2000 + Val(Right(CaseNo, 2))
520   For ThisPage = 1 To TotalPages

530     PrintHeadingRTB SendCopyTo, ViewOnly
540     PrintHeadingHistologyPreview SendCopyTo

550     TopLine = (ThisPage - 1) * LinesAllowed + 1
560     BottomLine = (ThisPage - 1) * LinesAllowed + LinesAllowed
570     If BottomLine > (plCounter) Then
580         BottomLine = plCounter
590     End If
600     For n = TopLine To BottomLine
610         rtfTemp = ""

620         If Left(pl(n), 18) = "GROSS EXAMINATION:" Or _
               Left(pl(n), 19) = "MICROSCOPIC REPORT:" Or _
               Left(pl(n), 17) = "CLINICAL DETAILS:" Or _
               Left(pl(n), 7) = "SNOMED:" Or _
               Left(pl(n), 17) = "GROSS APPEARANCE:" Or _
               Left(pl(n), 19) = "NATURE OF SPECIMEN:" Or _
               Left(pl(n), 19) = "CONTAINER LABELLED:" Or _
               Left(pl(n), 5) = "*****" Or _
               Left(pl(n), 15) = "UPDATED REPORT:" Then

630             PrintTextRTB frmRichText.rtb, vbCrLf, 10
640             PrintTextRTB frmRichText.rtb, Space(3)
650             PrintTextRTB frmRichText.rtb, pl(n), 10, True
660             PrintTextRTB frmRichText.rtb, vbCrLf, 10
670             PrintTextPreview vbCrLf, 10
680             PrintTextPreview Space(3)
690             PrintTextPreview pl(n), 10, True
700             PrintTextPreview vbCrLf, 10
710             CrCnt = CrCnt + 1
720         ElseIf UCase(Left(Replace(pl(n), vbLf, ""), 9)) = "PRINTDASH" Then
730             PrintTextRTB frmRichText.rtb, Space(3)
740             PrintTextRTB frmRichText.rtb, String(100, "-")
750             PrintTextRTB frmRichText.rtb, vbCrLf, 10
760             PrintTextPreview Space(3)
770             PrintTextPreview String(100, "-")
780             PrintTextPreview vbCrLf, 10
790         Else
800             PrintTextRTB frmRichText.rtb, Space(3)
810             If Left(pl(n), 6) = "{\rtf1" Then
820                 Do Until pl(n) = "\par }"
830                     rtfTemp = rtfTemp & pl(n)
840                     n = n + 1
850                 Loop
860             End If
870             PrintTextRTB frmRichText.rtb, rtfTemp & pl(n), 10


880             PrintTextRTB frmRichText.rtb, vbCrLf, 10
890             PrintTextPreview Space(3)
900             PrintTextPreview pl(n), 10
910             PrintTextPreview vbCrLf, 10
920             CrCnt = CrCnt + 1
930         End If

940     Next

'intBlanks = 0

950     Do While frmWorkSheet.PreviewPrint.CurrentY < (Printer.Height - 1100)
960         PrintTextPreview vbCrLf, 10
970         PrintTextRTB frmRichText.rtb, vbCrLf, 10
            'intBlanks = intBlanks + 1
            'Debug.Print frmWorkSheet.PreviewPrint.CurrentY
980     Loop
'Debug.Print intBlanks

990     PrintTextRTB frmRichText.rtb, FormatString(String(30, "_") & " Page " & ThisPage & " of " & TotalPages & " " & String(30, "_"), 110, , AlignCenter), , True
1000    PrintTextRTB frmRichText.rtb, vbCrLf
1010    PrintTextPreview FormatString(String(30, "_") & " Page " & ThisPage & " of " & TotalPages & " " & String(30, "_"), 110, , AlignCenter), , True
1020    PrintTextPreview vbCrLf

1030    If ThisPage <> TotalPages Then
1040        If ViewOnly = False Then
1050            frmWorkSheet.PreviewPrint.Cls
1060            frmRichText.rtb.SelStart = 0
1070            frmRichText.rtb.SelPrint Printer.hDC
1080            SaveRTF CaseNo, YYYY, ThisPage, UserName, CurrentTime & CopyNo
1090        Else
1100            frmWorkSheet.PreviewPrint.Cls
1110            PrintTextRTB frmRichText.rtb, vbCrLf, 32
1120        End If
1130    End If


1140  Next

    '#819017
    'Footer to read
    'Initial Report Authorised by:
    'Final Report Authorised by:

1150  PrintFooterRTB pStrValidatedBy, pStrValReportDate, pStrFinalValidatedBy, pStrFinalValReportDate, Validated
1160  PrintFooterHistologyPreview pStrValidatedBy, pStrValReportDate, pStrFinalValidatedBy, pStrFinalValReportDate, Validated

1170  If ViewOnly = False Then
1180    SaveRTF CaseNo, YYYY, ThisPage - 1, UserName, CurrentTime & CopyNo
1190    frmRichText.rtb.SelStart = 0
1200    frmRichText.rtb.SelPrint Printer.hDC
1210    If SendCopyTo <> "" Then
1220        CaseAddLogEvent CaseNo, ReportPrinted, "For " & SendCopyTo & " (" & vbGetComputerName() & " : " & Printer.DeviceName & ")"
1230    Else
1240        CaseAddLogEvent CaseNo, ReportPrinted, "(" & vbGetComputerName() & " : " & Printer.DeviceName & ")"
1250    End If
1260    Unload frmRichText
1270  End If



1280  frmWorkSheet.PreviewPrint.Cls


1290  Exit Sub

PrintHistology_Error:

    Dim strES As String
    Dim intEL As Integer

1300  intEL = Erl
1310  strES = Err.Description
1320  LogError "modPrintHistology", "PrintHistology", intEL, strES, sql

End Sub

Public Function FormatString(strDestString As String, _
                             intNumChars As Integer, _
                             Optional strSeperator As String = "", _
                             Optional intAlign As PrintAlignContants = Alignleft) As String

'**************intAlign = 0 --> Left Align
'**************intAlign = 1 --> Center Align
'**************intAlign = 2 --> Right Align
    Dim intPadding As Integer

10    On Error GoTo FormatString_Error

20    intPadding = 0

30    If Len(strDestString) > intNumChars Then
40      FormatString = Mid(strDestString, 1, intNumChars) & strSeperator
50    ElseIf Len(strDestString) < intNumChars Then
        Dim intStringLength As String
60      intStringLength = Len(strDestString)
70      intPadding = intNumChars - intStringLength

80      If intAlign = PrintAlignContants.Alignleft Then
90          strDestString = strDestString & String(intPadding, " ")  '& " "
100     ElseIf intAlign = PrintAlignContants.AlignCenter Then
110         If (intPadding Mod 2) = 0 Then
120             strDestString = String(intPadding / 2, " ") & strDestString & String(intPadding / 2, " ")
130         Else
140             strDestString = String((intPadding - 1) / 2, " ") & strDestString & String((intPadding - 1) / 2 + 1, " ")
150         End If
160     ElseIf intAlign = PrintAlignContants.AlignRight Then
170         strDestString = String(intPadding, " ") & strDestString
180     End If

190     strDestString = strDestString & strSeperator
200     FormatString = strDestString
210   Else
220     strDestString = strDestString & strSeperator
230     FormatString = strDestString
240   End If

250   Exit Function

FormatString_Error:

    Dim strES As String
    Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "Other", "FormatString", intEL, strES

End Function

Public Function PrintTextRTB(rtb As RichTextBox, ByVal Text As String, _
                             Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                             Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                             Optional FontColor As ColorConstants = vbBlack)

'---------------------------------------------------------------------------------------
' Procedure : PrintText
' DateTime  : 05/06/2008 11:40
' Author    : Babar Shahzad
' Note      : Printer object needs to be set first before calling this function.
'             Portrait mode (width X height) = 11800 X 16500
'---------------------------------------------------------------------------------------
10    On Error GoTo PrintTextRTB_Error

20    With rtb

30      .SelFontSize = FontSize
40      .SelBold = FontBold
50      .SelItalic = FontItalic
60      .SelUnderline = FontUnderLine
70      .SelColor = FontColor
80      If Left(Text, 6) = "{\rtf1" Then
90          .SelRTF = Text
100     Else
110         .SelText = Text
120     End If

130   End With

140   Exit Function

PrintTextRTB_Error:

    Dim strES As String
    Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "Other", "PrintTextRTB", intEL, strES

End Function

Public Function PrintTextPreview(ByVal Text As String, _
                                 Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                                 Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                                 Optional FontColor As ColorConstants = vbBlack)

'---------------------------------------------------------------------------------------
' Procedure : PrintText
' DateTime  : 05/06/2008 11:40
' Author    : Babar Shahzad
' Note      : Printer object needs to be set first before calling this function.
'             Portrait mode (width X height) = 11800 X 16500
'---------------------------------------------------------------------------------------



10    On Error GoTo PrintText_Error

20    With frmWorkSheet.PreviewPrint
30      .Font.Name = "Courier New"
40      .Font.Size = FontSize
50      .Font.Bold = FontBold
60      .Font.Italic = FontItalic
70      .Font.Underline = FontUnderLine
80      .ForeColor = FontColor
90      frmWorkSheet.PreviewPrint.Print Text;

100   End With


110   Exit Function

PrintText_Error:

    Dim strES As String
    Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modPrintHistology", "PrintText", intEL, strES


End Function

Public Function PrintText(ByVal Text As String, _
                          Optional FontName As String = "Courier New", _
                          Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                          Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                          Optional FontColor As ColorConstants = vbBlack)

'---------------------------------------------------------------------------------------
' Procedure : PrintText
' DateTime  : 05/06/2008 11:40
' Author    : Babar Shahzad
' Note      : Printer object needs to be set first before calling this function.
'             Portrait mode (width X height) = 11800 X 16500
'---------------------------------------------------------------------------------------



10    On Error GoTo PrintText_Error

20    With Printer
30      .Font.Name = FontName
40      .Font.Size = FontSize
50      .Font.Bold = FontBold
60      .Font.Italic = FontItalic
70      .Font.Underline = FontUnderLine
80      .ForeColor = FontColor

90      Printer.Print Text;

100   End With


110   Exit Function

PrintText_Error:

    Dim strES As String
    Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modPrintHistology", "PrintText", intEL, strES


End Function



' Truncate the string so it fits within the width.
Public Function BoundedText(ByVal ptr As Object, ByVal txt As String, ByVal max_wid As Single) As String
10    Do While ptr.TextWidth(txt) > max_wid
20      txt = Left$(txt, Len(txt) - 1)
30    Loop
40    BoundedText = txt
End Function
Public Sub PrintFlexGrid(ByVal ptr As Object, ByVal flx As MSFlexGrid, ByVal xmin As Single, ByVal ymin As Single, Optional lRowsPerPage As Long = 20)
    Const GAP = 60

    Dim xmax As Single
    Dim ymax As Single
    Dim X As Single
    Dim c As Integer


    Dim lRowsPrinted As Long
    Dim lThisRow As Long, lNumRows As Long
    Dim lPrinterPageHeight As Long
    Dim lPrintPage As Long

10    flx.TopRow = 1
20    lNumRows = flx.Rows - 1
30    lPrinterPageHeight = Printer.Height
40    lRowsPrinted = 1

    'Setup printer
50    ptr.Orientation = 2

60    xmax = xmin + GAP
70    For c = 0 To flx.Cols - 1
80      xmax = xmax + flx.ColWidth(c) + 2 * GAP
90    Next c

100   lPrintPage = 1
110   Do

120     Do

130         With flx

                ' Print each row.
140             ptr.CurrentY = ymin

150             ptr.Line (xmin, ptr.CurrentY)-(xmax, ptr.CurrentY)

160             ptr.CurrentY = ptr.CurrentY + GAP

170             X = xmin + GAP
180             For c = 0 To .Cols - 1
190                 ptr.CurrentX = X
200                 ptr.Print BoundedText(ptr, .TextMatrix(0, c), .ColWidth(c));
210                 X = X + .ColWidth(c) + 2 * GAP
220             Next c
230             ptr.CurrentY = ptr.CurrentY + GAP

                ' Move to the next line.
240             ptr.Print

250             For lThisRow = lRowsPrinted To lRowsPerPage * lPrintPage

260                 If lThisRow < lNumRows Then
270                     If lThisRow > 0 Then ptr.Line (xmin, ptr.CurrentY)-(xmax, ptr.CurrentY)
280                     ptr.CurrentY = ptr.CurrentY + GAP

                        ' Print the entries on this row.
290                     X = xmin + GAP
300                     For c = 0 To .Cols - 1
310                         ptr.CurrentX = X
320                         ptr.Print BoundedText(ptr, .TextMatrix(lThisRow, c), .ColWidth(c));
330                         X = X + .ColWidth(c) + 2 * GAP
340                     Next c
350                     ptr.CurrentY = ptr.CurrentY + GAP

                        ' Move to the next line.
360                     ptr.Print

370                     lRowsPrinted = lRowsPrinted + 1
380                 Else
390                     Exit Do
400                 End If
410             Next
420         End With
430     Loop While lRowsPrinted < lRowsPerPage * lPrintPage

440     ymax = ptr.CurrentY

        ' Draw a box around everything.
450     ptr.Line (xmin, ymin)-(xmax, ymax), , B

        ' Draw lines between the columns.
460     X = xmin
470     For c = 0 To flx.Cols - 2
480         X = X + flx.ColWidth(c) + 2 * GAP
490         ptr.Line (X, ymin)-(X, ymax)
500     Next c

510     ptr.EndDoc
520     lPrintPage = lPrintPage + 1

530   Loop While lRowsPrinted < lNumRows



End Sub

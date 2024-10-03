Attribute VB_Name = "modHeadFoot"
Option Explicit
Public CrCnt As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const EM_GETLINECOUNT = &HBA





Private Function getUserMCRN(ByVal strUserName As String) As String
          Dim sql As String
          Dim tb As Recordset



10        On Error GoTo getUserMCRN_Error

20        getUserMCRN = ""
30        sql = "SELECT MCRN From Users WHERE UserName = '" & AddTicks(Trim$(strUserName)) & "' "

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        If Not tb.EOF Then
70            getUserMCRN = tb!MCRN & ""
80        End If

90        Exit Function

getUserMCRN_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "modHeadFoot", "getUserMCRN", intEL, strES, sql


End Function

Public Function getUserCode(ByVal strUserName As String) As String
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo getUserCode_Error

20        getUserCode = ""
30        sql = "SELECT Code From Users WHERE UserName = '" & AddTicks(strUserName) & "' "

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        If Not tb.EOF Then
70            getUserCode = tb!Code & ""
80        End If

90        Exit Function

getUserCode_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "modHeadFoot", "getUserCode", intEL, strES, sql


End Function


Public Sub PrintHeadingRTB(Optional SendCopyTo As String, Optional ViewOnly As Boolean)

          Dim SampleID As String
          Dim Dept As String
          Dim name As String
          Dim Ward As String
          Dim DoB As String
          Dim Chart As String
          Dim Clinician As String
          Dim Address0 As String
          Dim Address1 As String
          Dim County As String
          Dim GP As String
          Dim GPAddress1 As String
          Dim GPAddress2 As String
          Dim GPCounty As String
          Dim Sex As String
          Dim Hospital As String
          Dim SampleDate As String
          Dim RecDate As String
          Dim Rundate As String
          Dim GpClin As String
          Dim SampleType As String
          Dim DateOfDeath As String
          Dim Coroner As String
          Dim TempCaseId As String

10        On Error GoTo PrintHeadingRTB_Error

20        CrCnt = 0

30        With udtHeading
40            SampleID = .SampleID
50            Dept = .Dept
60            name = .name
70            Ward = .Ward
80            DoB = .DoB
90            Chart = .Chart
100           Clinician = .Clinician
110           Address0 = .Address0
120           Address1 = .Address1
130           County = .County
140           GP = .GP
150           GPAddress1 = .GPAddress1
160           GPAddress2 = .GPAddress2
170           GPCounty = .GPCounty
180           Sex = .Sex
190           Hospital = .Hospital
200           SampleDate = .SampleDate
210           RecDate = .RecDate
220           Rundate = .Rundate
230           GpClin = .GpClin
240           SampleType = .SampleType
250           DateOfDeath = .DateOfDeath
260           Coroner = .Coroner
270       End With

280       With frmRichText

              '.rtb.Text = ""
290           If Mid(SampleID & "", 2, 1) = "P" Or Mid(SampleID & "", 2, 1) = "A" Then
300               TempCaseId = Left(SampleID, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(SampleID, 2)
310           Else
320               TempCaseId = Left(SampleID, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(SampleID, 2)
330           End If

340           PrintTextRTB .rtb, FormatString(GetOptionSetting("ReportHeader", "MIDLANDS REGIONAL HOSPITAL @ TULLAMORE"), 60, , AlignCenter), 16, True, , , vbBlack
350           PrintTextRTB .rtb, vbCrLf, 16
360           CrCnt = CrCnt + 1

370           Select Case Left(Dept, 4)
              Case "Hist"
380               PrintTextRTB .rtb, FormatString("Histopathology Dept.", 70, , AlignCenter), 14, True, , , vbBlack
390               PrintTextRTB .rtb, vbCrLf, 14
400               PrintTextRTB .rtb, FormatString("               Laboratory Phone : " & GetOptionSetting("ReportLabPhoneNo" & "  ", "057 - 9358338"), 50, , AlignCenter), 12, True, , , vbBlack
410               PrintTextRTB .rtb, FormatString("            " & udtHeading.DocomentNo, 40, , Alignleft), 8
             
420               If ViewOnly = False Then
430                   PrintTextRTB .rtb, vbCrLf, 12
440                   PrintTextRTB .rtb, FormatString("Printed On " & Format(Now, "dd/mm/yyyy hh:mm"), 95, , AlignCenter), 10
450               End If
460           Case Else
470               PrintTextRTB .rtb, FormatString("               Laboratory Phone : " & GetOptionSetting("ReportLabPhoneNo" & "  ", "057 - 9358338"), 50, , AlignCenter), 12, True, , , vbBlack
480               PrintTextRTB .rtb, FormatString("            " & udtHeading.DocomentNo, 40, , Alignleft), 8
490           End Select

500           PrintTextRTB .rtb, vbCrLf, 10
510           CrCnt = CrCnt + 1

              '********if accreditation text exists then print
520           If udtHeading.AccreditationText <> "" Then
                  'accreditation text on the whole line (left + right)
530               PrintTextRTB .rtb, Space(3)
540               PrintTextRTB .rtb, FormatString(udtHeading.AccreditationText, 108, , Alignleft), 8
550               PrintTextRTB .rtb, vbCrLf, 8
560               CrCnt = CrCnt + 1
570           End If
              
              'seperator line
580           PrintTextRTB .rtb, Space(3)
590           PrintTextRTB .rtb, String(100, "_"), 9

600           PrintTextRTB .rtb, vbCrLf
610           CrCnt = CrCnt + 1

              'name on the left side
620           PrintTextRTB .rtb, Space(3)
630           PrintTextRTB .rtb, FormatString("NAME:", 10, , Alignleft), 10
640           PrintTextRTB .rtb, FormatString(initial2upper(Left(name, 28)), 30, , Alignleft), 10, True

              'consultant on the right side
650           If Len(TempCaseId) = 12 Then
660               PrintTextRTB .rtb, FormatString("CORONER:", 16, , Alignleft), 10
670               PrintTextRTB .rtb, FormatString(initial2upper(Coroner), 30, , Alignleft), 10, True
680           Else
690               PrintTextRTB .rtb, FormatString("CONSULTANT:", 16, , Alignleft), 10
700               PrintTextRTB .rtb, FormatString(initial2upper(Clinician), 30, , Alignleft), 10, True
710           End If
720           PrintTextRTB .rtb, vbCrLf, 10
730           CrCnt = CrCnt + 1

              'lab number on the left side
740           PrintTextRTB .rtb, Space(3)
750           PrintTextRTB .rtb, FormatString("LAB NO:", 10, , Alignleft), 10
760           PrintTextRTB .rtb, FormatString(Trim(TempCaseId), 30, , Alignleft), 10, True

              'ward on the right side
770           PrintTextRTB .rtb, FormatString("WARD:", 16, , Alignleft), 10
780           PrintTextRTB .rtb, FormatString(UCase(Ward), 30, , Alignleft), 10, True

790           PrintTextRTB .rtb, vbCrLf, 10
800           CrCnt = CrCnt + 1

              'dob on the left side
810           PrintTextRTB .rtb, Space(3)
820           PrintTextRTB .rtb, FormatString("DOB:", 10, , Alignleft), 10
830           PrintTextRTB .rtb, FormatString(Format(DoB, "dd/mm/yyyy"), 30, , Alignleft), 10, True

              'chart number on the right side
840           PrintTextRTB .rtb, FormatString("CHART #:", 16, , Alignleft), 10
850           PrintTextRTB .rtb, FormatString(Trim(Chart), 30, , Alignleft), 10, True

860           PrintTextRTB .rtb, vbCrLf, 10
870           CrCnt = CrCnt + 1

              'sex on the left side
880           PrintTextRTB .rtb, Space(3)
890           PrintTextRTB .rtb, FormatString("SEX:", 10, , Alignleft), 10
900           Select Case Left(UCase(Trim(Sex)), 1)
              Case "Male": Sex = "Male"
910           Case "Female": Sex = "Female"
920           End Select
930           PrintTextRTB .rtb, FormatString(Sex, 30, , Alignleft), 10, True

              'gp on the right side
940           PrintTextRTB .rtb, FormatString("GP:", 16, , Alignleft), 10
950           PrintTextRTB .rtb, FormatString(UCase(GP), 30, , Alignleft), 10, True

960           PrintTextRTB .rtb, vbCrLf, 10
970           CrCnt = CrCnt + 1

              'address on the left side
980           PrintTextRTB .rtb, Space(3)
990           PrintTextRTB .rtb, FormatString("ADDRESS:", 10, , Alignleft), 10
1000          PrintTextRTB .rtb, FormatString(Left(UCase(Trim(Address0)), 21), 30, , Alignleft), 10, True

              'gp address on the right side
1010          PrintTextRTB .rtb, FormatString("GP ADDRESS:", 16, , Alignleft), 10
1020          PrintTextRTB .rtb, FormatString(Trim(GPAddress1), 30, , Alignleft), 10, True

1030          PrintTextRTB .rtb, vbCrLf, 10
1040          CrCnt = CrCnt + 1

              'address1 on the left side
1050          PrintTextRTB .rtb, Space(3)
1060          PrintTextRTB .rtb, Space(10), 10
1070          PrintTextRTB .rtb, FormatString(UCase(Trim(Address1)), 30, , Alignleft), 10, True

              'address2 on the right side
1080          PrintTextRTB .rtb, Space(3)
1090          PrintTextRTB .rtb, Space(13), 10
1100          PrintTextRTB .rtb, FormatString(UCase(Trim(GPAddress2)), 30, , Alignleft), 10, True

1110          PrintTextRTB .rtb, vbCrLf, 10
1120          CrCnt = CrCnt + 1

              'county on the left side
1130          PrintTextRTB .rtb, Space(3)
1140          PrintTextRTB .rtb, Space(10), 10
1150          PrintTextRTB .rtb, FormatString(UCase(Trim(County)), 30, , Alignleft), 10, True

              'gp county on the right side
1160          PrintTextRTB .rtb, Space(3)
1170          PrintTextRTB .rtb, Space(13), 10
1180          PrintTextRTB .rtb, FormatString(UCase(Trim(GPCounty)), 30, , Alignleft), 10, True

1190          PrintTextRTB .rtb, vbCrLf, 10
1200          CrCnt = CrCnt + 1

              'source on the left side
1210          PrintTextRTB .rtb, Space(3)
1220          PrintTextRTB .rtb, FormatString("SOURCE:", 10, , Alignleft), 10
1230          PrintTextRTB .rtb, FormatString(Trim(Hospital), 30, , Alignleft), 10, True

              'death date on the right side
1240          If Len(TempCaseId) = 12 Then
1250              PrintTextRTB .rtb, FormatString("DATE OF DEATH:", 16, , Alignleft), 10
1260              PrintTextRTB .rtb, FormatString(Format(DateOfDeath, "dd/mm/yyyy"), 30, , Alignleft), 10, True
1270          End If

1280          PrintTextRTB .rtb, vbCrLf, 10
1290          CrCnt = CrCnt + 1

1300          If SendCopyTo <> "" Then
                  'send copy to on the whole line (left + right)
1310              PrintTextRTB .rtb, Space(3)
1320              PrintTextRTB .rtb, FormatString("This is a COPY Report for the Attention of " & SendCopyTo, 95, , Alignleft), 10, True

1330              PrintTextRTB .rtb, vbCrLf, 10
1340              CrCnt = CrCnt + 1
1350          End If

              'seperator line (full line)
1360          PrintTextRTB .rtb, Space(3)
1370          PrintTextRTB .rtb, String(100, "_")

1380          PrintTextRTB .rtb, vbCrLf
1390          CrCnt = CrCnt + 1

              'sample date on the left side
1400          PrintTextRTB .rtb, Space(3)
1410          PrintTextRTB .rtb, FormatString("Sample Date :", 13, , Alignleft), 10
1420          PrintTextRTB .rtb, FormatString(Format(SampleDate, "dd/mm/yyyy"), 30, , Alignleft), 10

1430          PrintTextRTB .rtb, FormatString("Received :", 13, , Alignleft), 10
1440          PrintTextRTB .rtb, FormatString(Format(RecDate, "dd/MM/yyyy hh:mm"), 30, , Alignleft), 10

1450          PrintTextRTB .rtb, vbCrLf, 10
1460          CrCnt = CrCnt + 1

              'seperator line  (full line)
1470          PrintTextRTB .rtb, Space(3)
1480          PrintTextRTB .rtb, String(100, "_")

1490          PrintTextRTB .rtb, vbCrLf
1500          CrCnt = CrCnt + 1
1510      End With

1520      Exit Sub

PrintHeadingRTB_Error:

          Dim strES As String
          Dim intEL As Integer

1530      intEL = Erl
1540      strES = Err.Description
1550      LogError "modHeadFoot", "PrintHeadingRTB", intEL, strES

End Sub

Public Sub PrintHeadingHistology()

          Dim SampleID As String
          Dim Dept As String
          Dim name As String
          Dim Ward As String
          Dim DoB As String
          Dim Chart As String
          Dim Clinician As String
          Dim Address0 As String
          Dim Address1 As String
          Dim County As String
          Dim GP As String
          Dim GPAddress1 As String
          Dim GPAddress2 As String
          Dim GPCounty As String
          Dim Sex As String
          Dim Hospital As String
          Dim SampleDate As String
          Dim RecDate As String
          Dim Rundate As String
          Dim GpClin As String
          Dim SampleType As String
          Dim DateOfDeath As String
          Dim Coroner As String
          Dim TempCaseId As String

10        On Error GoTo PrintHeadingHistology_Error

20        With udtHeading
30            SampleID = .SampleID
40            Dept = .Dept
50            name = .name
60            Ward = .Ward
70            DoB = .DoB
80            Chart = .Chart
90            Clinician = .Clinician
100           Address0 = .Address0
110           Address1 = .Address1
120           County = .County
130           GP = .GP
140           GPAddress1 = .GPAddress1
150           GPAddress2 = .GPAddress2
160           GPCounty = .GPCounty
170           Sex = .Sex
180           Hospital = .Hospital
190           SampleDate = .SampleDate
200           RecDate = .RecDate
210           Rundate = .Rundate
220           GpClin = .GpClin
230           SampleType = .SampleType
240           DateOfDeath = .DateOfDeath
250           Coroner = .Coroner
260       End With

270       If Mid(SampleID & "", 2, 1) = "P" Or Mid(SampleID & "", 2, 1) = "A" Then
280           TempCaseId = Left(SampleID, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(SampleID, 2)
290       Else
300           TempCaseId = Left(SampleID, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(SampleID, 2)
310       End If

          'PrintText FormatString("MIDLANDS REGIONAL HOSPITAL @ TULLAMORE", 60, , AlignCenter), , 16, True, , , vbBlack
320       PrintText FormatString(GetOptionSetting("ReportHeader", "MIDLANDS REGIONAL HOSPITAL @ TULLAMORE"), 60, , AlignCenter), , 16, True, , , vbBlack
330       PrintText vbCrLf, , 16

340       Select Case Left(Dept, 4)

          Case "Hist"
350           PrintText FormatString("Histopathology Dept.", 70, , AlignCenter), , 14, True, , , vbBlack
360           PrintText vbCrLf, , 14
370           PrintText FormatString("Laboratory Phone : " & GetOptionSetting("ReportLabPhoneNo", "057 - 9358338"), 80, , AlignCenter), , 12, True, , , vbBlack
380           PrintText vbCrLf, , 12
390           PrintText FormatString("Printed On " & Format(Now, "dd/mm/yyyy hh:mm"), 95, , AlignCenter), , 10

400       Case Else
410           PrintText "Laboratory Phone : " & GetOptionSetting("ReportLabPhoneNo", "057 - 9358338")
420       End Select

430       PrintText vbCrLf, , 10
440       PrintText Space(3)
          'PrintTextPreview String(420, "-"),, 2
450       PrintText String(100, "_")

460       PrintText vbCrLf


470       PrintText Space(3)
480       PrintText FormatString("NAME:", 10, , Alignleft), , 10

490       PrintText FormatString(initial2upper(Left(name, 28)), 30, , Alignleft), , 10, True

500       If Len(TempCaseId) = 12 Then
510           PrintText FormatString("CORONER:", 16, , Alignleft), , 10
520           PrintText FormatString(initial2upper(Coroner), 30, , Alignleft), , 10, True
530       Else
540           PrintText FormatString("CONSULTANT:", 16, , Alignleft), , 10
550           PrintText FormatString(initial2upper(Clinician), 30, , Alignleft), , 10, True
560       End If
570       PrintText vbCrLf, , 10

580       PrintText Space(3)
590       PrintText FormatString("LAB NO:", 10, , Alignleft), , 10
600       PrintText FormatString(Trim(TempCaseId), 30, , Alignleft), , 10, True

610       PrintText FormatString("WARD:", 16, , Alignleft), , 10
620       PrintText FormatString(UCase(Ward), 30, , Alignleft), , 10, True

630       PrintText vbCrLf, , 10

640       PrintText Space(3)


650       PrintText FormatString("DOB:", 10, , Alignleft), , 10
660       PrintText FormatString(Format(DoB, "dd/mm/yyyy"), 30, , Alignleft), , 10, True

670       PrintText FormatString("CHART #:", 16, , Alignleft), , 10
680       PrintText FormatString(Trim(Chart), 30, , Alignleft), , 10, True

690       PrintText vbCrLf, , 10

700       PrintText Space(3)

710       PrintText FormatString("SEX:", 10, , Alignleft), , 10


720       Select Case Left(UCase(Trim(Sex)), 1)
          Case "M": Sex = "Male"
730       Case "F": Sex = "Female"
740       End Select

750       PrintText FormatString(Sex, 30, , Alignleft), , 10, True

760       PrintText FormatString("GP:", 16, , Alignleft), , 10
770       PrintText FormatString(UCase(GP), 30, , Alignleft), , 10, True

780       PrintText vbCrLf, , 10


790       PrintText Space(3)
800       PrintText FormatString("ADDRESS:", 10, , Alignleft), , 10
810       PrintText FormatString(Left(UCase(Trim(Address0)), 21), 22, , Alignleft), , 10, True

820       PrintText FormatString("GP ADDRESS:", 16, , Alignleft), , 10
830       PrintText FormatString(Trim(GPAddress1), 30, , Alignleft), , 10, True

840       PrintText vbCrLf, , 10


850       PrintText Space(3)
860       PrintText Space(10), , 10
870       PrintText FormatString(UCase(Trim(Address1)), 22, , Alignleft), , 10, True

880       PrintText Space(3)
890       PrintText Space(14), , 10
900       PrintText FormatString(UCase(Trim(GPAddress2)), 30, , Alignleft), , 10, True


910       PrintText vbCrLf, , 10

920       PrintText Space(3)
930       PrintText Space(10), , 10
940       PrintText FormatString(UCase(Trim(County)), 30, , Alignleft), , 10, True

950       PrintText Space(3)
960       PrintText Space(14), , 10
970       PrintText FormatString(UCase(Trim(GPCounty)), 30, , Alignleft), , 10, True

980       PrintText vbCrLf, , 10

990       PrintText Space(3)
1000      PrintText FormatString("SOURCE:", 10, , Alignleft), , 10
1010      PrintText FormatString(Trim(Hospital), 30, , Alignleft), , 10, True

1020      If Len(TempCaseId) = 12 Then
1030          PrintText FormatString("DATE OF DEATH:", 16, , Alignleft), , 10
1040          PrintText FormatString(Format(DateOfDeath, "dd/mm/yyyy"), 30, , Alignleft), , 10, True
1050      End If

          '    If SendCopyTo <> "" Then
          '        PrintText vbCrLf, , 10
          '        PrintText Space(3)
          '        PrintText FormatString("This is a COPY Report for the Attention of " & SendCopyTo, 95, , AlignCenter), , 10
          '    End If

1060      PrintText vbCrLf, , 10
1070      PrintText Space(3)
          'PrintText String(420, "-"),, 2
1080      PrintText String(100, "_")

1090      PrintText vbCrLf
1100      PrintText Space(3)
1110      PrintText FormatString("Sample Date :", 14, , Alignleft), , 10
1120      PrintText FormatString(Format(SampleDate, "dd/mm/yyyy hh:mm"), 30, , Alignleft), , 10

1130      PrintText FormatString("Received :", 11, , Alignleft), , 10
1140      PrintText FormatString(Format(RecDate, "dd/MM/yyyy hh:mm"), 30, , Alignleft), , 10

1150      PrintText vbCrLf, , 10

1160      PrintText Space(3)
1170      PrintText String(100, "_")
1180      PrintText vbCrLf


1190      Exit Sub

PrintHeadingHistology_Error:

          Dim strES As String
          Dim intEL As Integer

1200      intEL = Erl
1210      strES = Err.Description
1220      LogError "modHeadFoot", "PrintHeadingHistology", intEL, strES


End Sub

Public Sub PrintHeadingHistologyPreview(Optional SendCopyTo As String)

          Dim SampleID As String
          Dim Dept As String
          Dim name As String
          Dim Ward As String
          Dim DoB As String
          Dim Chart As String
          Dim Clinician As String
          Dim Address0 As String
          Dim Address1 As String
          Dim County As String
          Dim GP As String
          Dim GPAddress1 As String
          Dim GPAddress2 As String
          Dim GPCounty As String
          Dim Sex As String
          Dim Hospital As String
          Dim SampleDate As String
          Dim RecDate As String
          Dim Rundate As String
          Dim GpClin As String
          Dim SampleType As String
          Dim DateOfDeath As String
          Dim Coroner As String
          Dim TempCaseId As String


10        On Error GoTo PrintHeadingHistologyPreview_Error

20        With udtHeading
30            SampleID = .SampleID
40            Dept = .Dept
50            name = .name
60            Ward = .Ward
70            DoB = .DoB
80            Chart = .Chart
90            Clinician = .Clinician
100           Address0 = .Address0
110           Address1 = .Address1
120           County = .County
130           GP = .GP
140           GPAddress1 = .GPAddress1
150           GPAddress2 = .GPAddress2
160           GPCounty = .GPCounty
170           Sex = .Sex
180           Hospital = .Hospital
190           SampleDate = .SampleDate
200           RecDate = .RecDate
210           Rundate = .Rundate
220           GpClin = .GpClin
230           SampleType = .SampleType
240           DateOfDeath = .DateOfDeath
250           Coroner = .Coroner
260       End With

270       If Mid(SampleID & "", 2, 1) = "P" Or Mid(SampleID & "", 2, 1) = "A" Then
280           TempCaseId = Left(SampleID, 7) & " " & sysOptCaseIdSeperator(0) & " " & Right(SampleID, 2)
290       Else
300           TempCaseId = Left(SampleID, 6) & " " & sysOptCaseIdSeperator(0) & " " & Right(SampleID, 2)
310       End If

320       PrintTextPreview FormatString(GetOptionSetting("ReportHeader", "MIDLANDS REGIONAL HOSPITAL @ TULLAMORE"), 60, , AlignCenter), 16, True, , , vbBlack

330       PrintTextPreview vbCrLf, 16

340       Select Case Left(Dept, 4)

          Case "Hist"
350           PrintTextPreview FormatString("Histopathology Dept.", 70, , AlignCenter), 14, True, , , vbBlack
360           PrintTextPreview vbCrLf, 14
370           PrintTextPreview FormatString("               Laboratory Phone : " & GetOptionSetting("ReportLabPhoneNo", "057 - 9358338"), 80, , AlignCenter), 12, True, , , vbBlack
380           PrintTextPreview FormatString("            " & udtHeading.DocomentNo, 40, , Alignleft), 8
390           PrintTextPreview vbCrLf, 12
400           PrintTextPreview FormatString("Printed On " & Format(Now, "dd/mm/yyyy hh:mm"), 95, , AlignCenter), 10

410       Case Else
420           PrintTextPreview FormatString("               Laboratory Phone : " & GetOptionSetting("ReportLabPhoneNo", "057 - 9358338"), 80, , AlignCenter), 12, True, , , vbBlack
430           PrintTextPreview FormatString("            " & udtHeading.DocomentNo, 40, , Alignleft), 8
440       End Select

          '********if accreditation text exists then print
450       If udtHeading.AccreditationText <> "" Then
460           PrintTextPreview vbCrLf, 8
470           PrintTextPreview Space(3)
480           PrintTextPreview FormatString(udtHeading.AccreditationText, 108, , AlignCenter), 8
490       End If

          'PrintTextPreview vbCrLf, 10
500       PrintTextPreview Space(3)
510       PrintTextPreview String(100, "_")

520       PrintTextPreview vbCrLf


530       PrintTextPreview Space(3)
540       PrintTextPreview FormatString("NAME:", 10, , Alignleft), 10

550       PrintTextPreview FormatString(initial2upper(Left(name, 28)), 30, , Alignleft), 10, True

560       If Len(TempCaseId) = 12 Then
570           PrintTextPreview FormatString("CORONER:", 16, , Alignleft), 10
580           PrintTextPreview FormatString(initial2upper(Coroner), 30, , Alignleft), 10, True
590       Else
600           PrintTextPreview FormatString("CONSULTANT:", 16, , Alignleft), 10
610           PrintTextPreview FormatString(initial2upper(Clinician), 30, , Alignleft), 10, True
620       End If


630       PrintTextPreview vbCrLf, 10

640       PrintTextPreview Space(3)
650       PrintTextPreview FormatString("LAB NO:", 10, , Alignleft), 10
660       PrintTextPreview FormatString(Trim(TempCaseId), 30, , Alignleft), 10, True

670       PrintTextPreview FormatString("WARD:", 16, , Alignleft), 10
680       PrintTextPreview FormatString(UCase(Ward), 30, , Alignleft), 10, True

690       PrintTextPreview vbCrLf, 10

700       PrintTextPreview Space(3)


710       PrintTextPreview FormatString("DOB:", 10, , Alignleft), 10
720       PrintTextPreview FormatString(Format(DoB, "dd/mm/yyyy"), 30, , Alignleft), 10, True

730       PrintTextPreview FormatString("CHART #:", 16, , Alignleft), 10
740       PrintTextPreview FormatString(Trim(Chart), 30, , Alignleft), 10, True

750       PrintTextPreview vbCrLf, 10

760       PrintTextPreview Space(3)

770       PrintTextPreview FormatString("SEX:", 10, , Alignleft), 10


780       Select Case Left(UCase(Trim(Sex)), 1)
          Case "M": Sex = "Male"
790       Case "F": Sex = "Female"
800       End Select

810       PrintTextPreview FormatString(Sex, 30, , Alignleft), 10, True

820       PrintTextPreview FormatString("GP:", 16, , Alignleft), 10
830       PrintTextPreview FormatString(UCase(GP), 30, , Alignleft), 10, True

840       PrintTextPreview vbCrLf, 10


850       PrintTextPreview Space(3)
860       PrintTextPreview FormatString("ADDRESS:", 10, , Alignleft), 10
870       PrintTextPreview FormatString(Left(UCase(Trim(Address0)), 21), 22, , Alignleft), 10, True

880       PrintTextPreview FormatString("GP ADDRESS:", 16, , Alignleft), 10
890       PrintTextPreview FormatString(Trim(GPAddress1), 30, , Alignleft), 10, True

900       PrintTextPreview vbCrLf, 10


910       PrintTextPreview Space(3)
920       PrintTextPreview Space(10), 10
930       PrintTextPreview FormatString(UCase(Trim(Address1)), 22, , Alignleft), 10, True

940       PrintTextPreview Space(3)
950       PrintTextPreview Space(14), 10
960       PrintTextPreview FormatString(UCase(Trim(GPAddress2)), 30, , Alignleft), 10, True


970       PrintTextPreview vbCrLf, 10

980       PrintTextPreview Space(3)
990       PrintTextPreview Space(10), 10
1000      PrintTextPreview FormatString(UCase(Trim(County)), 30, , Alignleft), 10, True

1010      PrintTextPreview Space(3)
1020      PrintTextPreview Space(14), 10
1030      PrintTextPreview FormatString(UCase(Trim(GPCounty)), 30, , Alignleft), 10, True

1040      PrintTextPreview vbCrLf, 10

1050      PrintTextPreview Space(3)
1060      PrintTextPreview FormatString("SOURCE:", 10, , Alignleft), 10
1070      PrintTextPreview FormatString(Trim(Hospital), 30, , Alignleft), 10, True


1080      If Len(TempCaseId) = 12 Then
1090          PrintTextPreview FormatString("DATE OF DEATH:", 16, , Alignleft), 10
1100          PrintTextPreview FormatString(Format(DateOfDeath, "dd/mm/yyyy"), 30, , Alignleft), 10, True
1110      End If

1120      If SendCopyTo <> "" Then
1130          PrintTextPreview vbCrLf, 10
1140          PrintTextPreview Space(3)
1150          PrintTextPreview FormatString("This is a COPY Report for the Attention of " & SendCopyTo, 95, , AlignCenter), 10
1160      End If

1170      PrintTextPreview vbCrLf, 10
1180      PrintTextPreview Space(3)
          'PrintText String(420, "-"), 2
1190      PrintTextPreview String(100, "_")

1200      PrintTextPreview vbCrLf
1210      PrintTextPreview Space(3)
1220      PrintTextPreview FormatString("Sample Date :", 14, , Alignleft), 10
1230      PrintTextPreview FormatString(Format(SampleDate, "dd/mm/yyyy hh:mm"), 30, , Alignleft), 10

1240      PrintTextPreview FormatString("Received :", 11, , Alignleft), 10
1250      PrintTextPreview FormatString(Format(RecDate, "dd/MM/yyyy hh:mm"), 30, , Alignleft), 10

1260      PrintTextPreview vbCrLf, 10

1270      PrintTextPreview Space(3)
1280      PrintTextPreview String(100, "_")
1290      PrintTextPreview vbCrLf




1300      Exit Sub

PrintHeadingHistologyPreview_Error:

          Dim strES As String
          Dim intEL As Integer

1310      intEL = Erl
1320      strES = Err.Description
1330      LogError "modHeadFoot", "PrintHeadingHistologyPreview", intEL, strES


End Sub

Public Sub PrintHeadingAuditTrail(CaseId As String, PageNumber As String)




10        On Error GoTo PrintHeadingAuditTrail_Error

20        PrintText vbCrLf, , 16
30        PrintText Space(13)

40        PrintText FormatString(GetOptionSetting("ReportHeader", "MIDLANDS REGIONAL HOSPITAL @ TULLAMORE"), 60, , Alignleft), , 16, True, , , vbBlack

50        PrintText vbCrLf, , 16

60        PrintText Space(13)
70        PrintText FormatString("Histopathology Dept.", 70, , Alignleft), , 14, True, , , vbBlack
80        PrintText vbCrLf, , 14
90        PrintText Space(13)
100       PrintText FormatString("Laboratory Phone : " & GetOptionSetting("ReportLabPhoneNo", "057 - 9358338"), 80, , Alignleft), , 12, True, , , vbBlack
110       PrintText vbCrLf, , 12
120       PrintText Space(13)
130       PrintText FormatString("Audit Trail For Case ID : " & CaseId, 70, , Alignleft), , 10
140       PrintText Space(40)
150       PrintText FormatString(PageNumber, 50, , Alignleft), , 10
160       PrintText vbCrLf, , 10
170       PrintText Space(3)
          'PrintText String(420, "-"), 2
180       PrintText vbCrLf





190       Exit Sub

PrintHeadingAuditTrail_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "modHeadFoot", "PrintHeadingAuditTrail", intEL, strES


End Sub

Public Sub PrintHeadingWorkLog(PageNumber As String, WorkLogType As String, Optional iSpace As Integer = 13)


10        PrintText vbCrLf, , 16
20        PrintText Space(iSpace)

30        PrintText FormatString(GetOptionSetting("ReportHeader", "MIDLANDS REGIONAL HOSPITAL @ TULLAMORE"), 60, , Alignleft), , 16, True, , , vbBlack

40        PrintText vbCrLf, , 16

50        PrintText Space(iSpace)
60        PrintText FormatString("Histopathology Dept.", 70, , Alignleft), , 14, True, , , vbBlack
70        PrintText vbCrLf, , 14
80        PrintText Space(iSpace)
90        PrintText FormatString("Laboratory Phone : " & GetOptionSetting("ReportLabPhoneNo", "057 - 9358338"), 80, , Alignleft), , 12, True, , , vbBlack
100       PrintText vbCrLf, , 12
110       PrintText Space(iSpace)
120       PrintText FormatString(WorkLogType & " (" & PageNumber & ")", 70, , Alignleft), , 10
130       PrintText vbCrLf, , 10
140       PrintText Space(iSpace)
150       PrintText FormatString("Date : " & Format(Now, "dd/mm/yyyy hh:mm"), 70, , Alignleft), , 10
160       PrintText vbCrLf, , 10
          '170   PrintText Space(3)
          '      'PrintText String(420, "-"), 2
          '180   PrintText vbCrLf


End Sub

Public Sub PrintFooterRTB(ByVal strFirstInitiator As String, _
                          ByVal strFirstValidatedDate As String, _
                          ByVal strFinalValidatedBy As String, _
                          ByVal strFinalValReportDate As String, _
                          ByVal Valid As Boolean)

          Dim sql As String
          Dim Y As Long


10        On Error GoTo PrintFooterRTB_Error

20        Y = Printer.Height

30        With frmRichText


40            PrintTextRTB .rtb, vbCrLf, 10
50            PrintTextRTB .rtb, Space(3)
60            If Valid Then
70                If strFinalValReportDate = "" Then    '1st Authorisation date
80                    PrintTextRTB .rtb, FormatString("Report Authorised by:", 22)
90                    PrintTextRTB .rtb, FormatString(strFirstInitiator, 26)
100                   PrintTextRTB .rtb, FormatString("MCRN:", 5)
110                   PrintTextRTB .rtb, FormatString(getUserMCRN(strFirstInitiator), 8)
120                   PrintTextRTB .rtb, FormatString("Date:", 5, , Alignleft)
130                   PrintTextRTB .rtb, FormatString(Format(strFirstValidatedDate, "dd/MM/yyyy hh:mm"), 16, , Alignleft)
140               Else    'subsequent Authorisations
150                   PrintTextRTB .rtb, FormatString("Initial Report Authorised by:", 31)
160                   PrintTextRTB .rtb, FormatString(strFirstInitiator, 26)
170                   PrintTextRTB .rtb, FormatString("MCRN:", 5)
180                   PrintTextRTB .rtb, FormatString(getUserMCRN(strFirstInitiator), 8)
190                   PrintTextRTB .rtb, FormatString("Date:", 5, , Alignleft)
200                   PrintTextRTB .rtb, FormatString(Format(strFirstValidatedDate, "dd/MM/yyyy hh:mm"), 16, , Alignleft)
210                   PrintTextRTB .rtb, vbCrLf, 10
220                   PrintTextRTB .rtb, Space(3)
230                   PrintTextRTB .rtb, FormatString("Final Report Authorised by:", 31)
240                   PrintTextRTB .rtb, FormatString(strFinalValidatedBy, 26)
250                   PrintTextRTB .rtb, FormatString("MCRN:", 5)
260                   PrintTextRTB .rtb, FormatString(getUserMCRN(strFinalValidatedBy), 8)
270                   PrintTextRTB .rtb, FormatString("Date:", 5, , Alignleft)
280                   PrintTextRTB .rtb, FormatString(Format(strFinalValReportDate, "dd/MM/yyyy hh:mm"), 16, , Alignleft)
290               End If

300           Else
310               PrintTextRTB .rtb, FormatString("Preliminary :", 14)
320               PrintTextRTB .rtb, FormatString(strFirstInitiator, 50)
330               PrintTextRTB .rtb, FormatString("Preliminary Date:", 18, , Alignleft)
340           End If
350           PrintTextRTB .rtb, vbCrLf, 12
360       End With



370       Exit Sub

PrintFooterRTB_Error:

          Dim strES As String
          Dim intEL As Integer

380       intEL = Erl
390       strES = Err.Description
400       LogError "modHeadFoot", "PrintFooterRTB", intEL, strES, sql


End Sub

Public Sub PrintFooterHistology(ByVal Initiator As String, _
                                ByVal SampleDate As String, _
                                ByVal Valid As Boolean)

          Dim sql As String
          Dim Y As Long

10        On Error GoTo PrintFooterHistology_Error

20        Y = Printer.Height


30        PrintText vbCrLf, , 10
40        PrintText Space(3)

50        If Valid Then
60            PrintText FormatString("Authorised By:", 14)
70            PrintText FormatString(Initiator, 26)
80            PrintText FormatString("MCRN:", 5)
90            PrintText FormatString(getUserMCRN(Initiator), 18)
100           PrintText FormatString("Authorised Date:", 16, , Alignleft)
110           PrintText FormatString(Format(SampleDate, "dd/MM/yyyy hh:mm"), 16, , Alignleft)    '& Format(Rundate, "dd/MMM/yyyy hh:mm")
120       Else
130           PrintText FormatString("Preliminary :", 14)
140           PrintText FormatString(Initiator, 50)
150           PrintText FormatString("Preliminary Date:", 18, , Alignleft)

160       End If
170       PrintText vbCrLf, , 12

180       Exit Sub

PrintFooterHistology_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "modHeadFoot", "PrintFooterHistology", intEL, strES, sql


End Sub

Public Sub PrintFooterHistologyPreview(ByVal strFirstInitiator As String, _
                                       ByVal strFirstValidatedDate As String, _
                                       ByVal strFinalValidatedBy As String, _
                                       ByVal strFinalValReportDate As String, _
                                       ByVal Valid As Boolean)


          Dim Y As Long


10        On Error GoTo PrintFooterHistologyPreview_Error

20        Y = Printer.Height

30        PrintTextPreview vbCrLf, 10
40        PrintTextPreview Space(3)

50        If Valid Then
60            If strFinalValReportDate = "" Then    '1st Authorisation date
70                PrintTextPreview FormatString("Report Authorised by:", 22)
80                PrintTextPreview FormatString(strFirstInitiator, 26)
90                PrintTextPreview FormatString("MCRN:", 5)
100               PrintTextPreview FormatString(getUserMCRN(strFirstInitiator), 8)
110               PrintTextPreview FormatString("Date:", 5, , Alignleft)
120               PrintTextPreview FormatString(Format(strFirstValidatedDate, "dd/MM/yyyy hh:mm"), 16, , Alignleft)
130           Else 'Subsequent Authorisation
140               PrintTextPreview FormatString("Initial Report Authorised by:", 31)
150               PrintTextPreview FormatString(strFirstInitiator, 26)
160               PrintTextPreview FormatString("MCRN:", 5)
170               PrintTextPreview FormatString(getUserMCRN(strFirstInitiator), 8)
180               PrintTextPreview FormatString("Date:", 5, , Alignleft)
190               PrintTextPreview FormatString(Format(strFirstValidatedDate, "dd/MM/yyyy hh:mm"), 16, , Alignleft)

200               PrintTextPreview vbCrLf, 10
210               PrintTextPreview Space(3)

220               PrintTextPreview FormatString("Final Report Authorised by:", 31)
230               PrintTextPreview FormatString(strFinalValidatedBy, 26)
240               PrintTextPreview FormatString("MCRN:", 5)
250               PrintTextPreview FormatString(getUserMCRN(strFinalValidatedBy), 8)
260               PrintTextPreview FormatString("Date:", 5, , Alignleft)
270               PrintTextPreview FormatString(Format(strFinalValReportDate, "dd/MM/yyyy hh:mm"), 16, , Alignleft)
280           End If
290       Else
300           PrintTextPreview FormatString("Preliminary :", 14)
310           PrintTextPreview FormatString(strFirstInitiator, 50)
320           PrintTextPreview FormatString("Preliminary Date:", 18, , Alignleft)
330       End If
340       PrintTextPreview vbCrLf, 12

350       Exit Sub

PrintFooterHistologyPreview_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "modHeadFoot", "PrintFooterHistologyPreview", intEL, strES


End Sub



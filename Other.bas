Attribute VB_Name = "Other"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const MaxAgeToDays As Long = 43830

Public Const gVALID = 1
Public Const gNOTVALID = 2
Public Const gPRINTED = 1
Public Const gNOTPRINTED = 2
Public Const gDONTCARE = 0

Public Cn As Integer

Public gData(1 To 365, 1 To 3) As Variant '(n,1)=rundate, (n,2)=INR, (n,3)=Warfarin

Public LatestSampleID As String

Public LatestINR As String

Public CurrentDose As String
Public pLatest As String
Public pEarliest As String

Public pLowerTarget As String
Public pUpperTarget As String
Public pCondition As String

Public pForcePrintTo As String

'Public UserName As String

Public Type PrintLine
  Analyte As String * 16
  Result As String * 6
  Flag As String * 3
  Units As String * 7
  NormalRange As String * 11
  Fasting As String * 9
  Reason As String * 23
End Type

Public Type ReportToPrint
  Department As String
  SampleID As String
  Initiator As String
  Ward As String
  Clinician As String
  GP As String
  FAXNumber As String
  UsePrinter As String
  Year As String
End Type

Public rp As ReportToPrint

Private Type udtHead
  SampleID As String
  Dept As String
  Name As String
  Ward As String
  DoB As String
  Chart As String
  Clinician As String
  Address0 As String
  Address1 As String
  County As String
  GP As String
  GPAddress1 As String
  GPAddress2 As String
  GPCounty As String
  Sex As String
  Hospital As String
  SampleDate As String
  RecDate As String
  Rundate As String
  Tn As Integer
  GpClin As String
  SampleType As String
  DateOfDeath As String
  Coroner As String
  DocomentNo As String
  AccreditationText As String
End Type
Public udtHeading As udtHead

Public Enum PrintAlignContants
    AlignLeft = 0
    AlignCenter = 1
    AlignRight = 2
End Enum


Public Sub FillCommentLines(ByVal FullComment As String, _
                             ByVal NumberOfLines As Integer, _
                             ByRef Comments() As String, _
                             Optional ByVal MaxLen As Integer)

      Dim n As Integer
      Dim CurrentLine As Integer
      Dim X As Integer
      Dim ThisLine As String
      Dim SpaceFound As Boolean

10    On Error Resume Next

20    For n = 1 To UBound(Comments)
30      Comments(n) = ""
40    Next

50    CurrentLine = 0
60    FullComment = Trim(FullComment)
70    n = Len(FullComment)

80    For X = n - 1 To 1 Step -1
90      If Mid(FullComment, X, 1) = vbCr Or Mid(FullComment, X, 1) = vbLf Or Mid(FullComment, X, 1) = vbTab Then
100       Mid(FullComment, X, 1) = " "
110     End If
120   Next

130   For X = n - 3 To 1 Step -1
140     If Mid(FullComment, X, 2) = "  " Then
150       FullComment = Left(FullComment, X) & Mid(FullComment, X + 2)
160     End If
170   Next
180   n = Len(FullComment)

190   Do While n > MaxLen
200     SpaceFound = False
210     For X = MaxLen To 1 Step -1
220       If Mid(FullComment, X, 1) = " " Then
230         ThisLine = Left(FullComment, X - 1)
240         FullComment = Mid(FullComment, X + 1)

250         CurrentLine = CurrentLine + 1
260         If CurrentLine <= NumberOfLines Then
270           Comments(CurrentLine) = ThisLine
280         End If
290         SpaceFound = True
300         Exit For
310       End If
320     Next
330     If Not SpaceFound Then
340       ThisLine = Left(FullComment, MaxLen)
350       FullComment = Mid(FullComment, MaxLen + 1)
    
360       CurrentLine = CurrentLine + 1
370       If CurrentLine <= NumberOfLines Then
380         Comments(CurrentLine) = ThisLine
390       End If
400     End If
410     n = Len(FullComment)
420   Loop

430   CurrentLine = CurrentLine + 1
440   If CurrentLine <= NumberOfLines Then
450     Comments(CurrentLine) = FullComment
460   End If

End Sub






Public Sub ClearUdtHeading()

10    With udtHeading
20        .SampleID = ""
30        .Dept = ""
40        .Name = ""
50        .Ward = ""
60        .DoB = ""
70        .Chart = ""
80        .Clinician = ""
90        .Address0 = ""
100       .Address1 = ""
110       .County = ""
120       .GP = ""
130       .GPAddress1 = ""
140       .GPAddress2 = ""
150       .GPCounty = ""
160       .Sex = ""
170       .Hospital = ""
180       .SampleDate = ""
190       .RecDate = ""
200       .Rundate = ""
210       .Tn = 0
220       .GpClin = ""
230       .SampleType = ""
240       .DateOfDeath = ""
250       .Coroner = ""
260   End With

End Sub


Public Function IsPrinterA5Landscape(ByVal strPrinter As String) As Boolean
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo IsPrinterA5Landscape_Error

20    IsPrinterA5Landscape = False

30    sql = "Select * from Printers where " & _
            "PrinterName = '" & strPrinter & "'" & _
            " and Orientation = 'A5 LANDSCAPE'"

40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
    
60    If Not tb.EOF Then
70        IsPrinterA5Landscape = True
80    End If

90    Exit Function

IsPrinterA5Landscape_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "Other", "IsPrinterA5Landscape", intEL, strES, sql


End Function


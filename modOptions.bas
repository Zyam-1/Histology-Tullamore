Attribute VB_Name = "modOptions"
Option Explicit

Public sysOptCaseIdSeperator() As String
Public sysOptCaseIdValidation() As String
Public sysOptTissueTypeNumberingFormat() As String
Public sysOptBlockNumberingFormat() As String
Public sysOptSlideNumberingFormat() As String
Public SysOptChange() As Boolean
Public sysOptCurrentLanguage As String




Public Sub LoadOptions()

Dim tb As New Recordset
Dim sql As String
Dim n As Long

On Error GoTo LoadOptions_Error

ReDimOptions

For n = 0 To intOtherHospitalsInGroup

  sql = "SELECT * from Options "
  
  Set tb = New Recordset
  RecOpenServer n, tb, sql
  Do While Not tb.EOF
    Select Case UCase$(Trim$(tb!Description & ""))
        Case "CASEIDSEPERATOR": sysOptCaseIdSeperator(n) = Trim(tb!Contents & "")
        Case "CASEIDVALIDATION": sysOptCaseIdValidation(n) = Trim(tb!Contents & "")
        Case "TISSUETYPENUMBERINGFORMAT": sysOptTissueTypeNumberingFormat(n) = Trim(tb!Contents & "")
        Case "BLOCKNUMBERINGFORMAT": sysOptBlockNumberingFormat(n) = Trim(tb!Contents & "")
        Case "SLIDENUMBERINGFORMAT": sysOptSlideNumberingFormat(n) = Trim(tb!Contents & "")
        Case "CHANGE": SysOptChange(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        

    End Select
    tb.MoveNext
  Loop
Next
sysOptCurrentLanguage = GetOptionSetting("CurrentLanguage", "English")

Exit Sub

LoadOptions_Error:

Dim strES As String
Dim intEL As Integer



intEL = Erl
strES = Err.Description
LogError "modOptions", "LoadOptions", intEL, strES, sql

End Sub
Private Sub ReDimOptions()
10    ReDim sysOptCaseIdSeperator(0 To intOtherHospitalsInGroup) As String
20    ReDim sysOptCaseIdValidation(0 To intOtherHospitalsInGroup) As String
30    ReDim sysOptTissueTypeNumberingFormat(0 To intOtherHospitalsInGroup) As String
40    ReDim sysOptBlockNumberingFormat(0 To intOtherHospitalsInGroup) As String
50    ReDim sysOptSlideNumberingFormat(0 To intOtherHospitalsInGroup) As String
60    ReDim SysOptChange(0 To intOtherHospitalsInGroup) As Boolean
      'User Options


End Sub
Public Function GetOptionSetting(ByVal Description As String, _
                                 ByVal Default As String) As String
   
      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo GetOptionSetting_Error

20    sql = "SELECT Contents FROM Options WHERE " & _
            "Description = '" & Description & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      RetVal = Default
70    ElseIf Trim$(tb!Contents & "") = "" Then
80      RetVal = Default
90    Else
100     RetVal = tb!Contents
110   End If

120   GetOptionSetting = RetVal

130   Exit Function

GetOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "modOptions", "GetOptionSetting", intEL, strES, sql

End Function

Public Sub SaveOptionSetting(ByVal Description As String, _
                             ByVal Contents As String)
   
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveOptionSetting_Error

20    sql = "SELECT * FROM Options WHERE " & _
            "Description = '" & Description & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60      tb.AddNew
70    End If
80    tb!Description = Description
90    tb!Contents = Contents
100   tb.Update

110   Exit Sub

SaveOptionSetting_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modOptions", "SaveOptionSetting", intEL, strES, sql

End Sub


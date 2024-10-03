Attribute VB_Name = "modEventShared"
Public Enum CaseEvents
    CutUpEvent = 1  '
    EmbeddedEvent = 2 '"Embedded"
    PiecesCutUp = 3 '
    CuttingBy = 4 '
    AssistedBy = 5 '
    PiecesEmbedding = 6 '
    WithPathologist = 7
    InHistology = 8
    AwaitingAuthorisation = 9
    TreeNodeAdded = 10
    TreeNodeDeleted = 11
    DemographicsAdded = 12
    DemographicsEdited = 13
    GrossEdited = 14
    MicroEdited = 15
    PCodeEdited = 16
    MCodeAdded = 17
    QCodeAdded = 18
    CodeDeleted = 19
    Authorised = 20
    UnAuthorised = 21
    DiscrepancyAdded = 22
    DiscrepancyEdited = 23
    ReportPrinted = 24
    Processor = 25
    Disposal = 26
    TreeNodeEdited = 27
    ExtraRequestsRemoved = 28
End Enum

Public Function CaseAddLogEvent(CaseId As String, Evnt As CaseEvents, _
                                    Optional Comments As String = "", _
                                    Optional Path As String = "") As Boolean

      Dim tb As Recordset
      Dim sql As String



10    On Error GoTo CaseAddLogEvent_Error

20    sql = "SELECT * FROM CaseEventLog WHERE 1 = 0"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    tb.AddNew
60    UpdateEvent tb, sql, CaseId, Evnt, Comments, Path

70    CaseAddLogEvent = True

80    tb.Close



90    Exit Function

100 CaseAddLogEvent_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modEventShared", "CaseAddLogEvent", intEL, strES, sql


End Function

Public Function CaseUpdateLogEvent(CaseId As String, Evnt As CaseEvents, _
                              Optional Comments As String = "", _
                              Optional Path As String = "") As Boolean

      Dim tb As Recordset
      Dim sql As String



10    On Error GoTo CaseUpdateLogEvent_Error

20    sql = "SELECT * FROM CaseEventLog WHERE EventId = '" & Evnt & "' " & _
        "AND CaseId = '" & CaseId & "' "

30    If Path <> "" Then
40        sql = sql & "AND Path = '" & AddTicks(Path) & "'"
50    End If
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql


80    If Not tb.EOF Then
90        If tb!Comments <> Comments Then

100           UpdateEvent tb, sql, CaseId, Evnt, Comments, Path
110       End If
120   Else
130       tb.AddNew
140       UpdateEvent tb, sql, CaseId, Evnt, Comments, Path

150   End If


160   CaseUpdateLogEvent = True

170   tb.Close



180   Exit Function

190 CaseUpdateLogEvent_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "modEventShared", "CaseUpdateLogEvent", intEL, strES, sql


End Function

Private Sub UpdateEvent(ByRef tb As Recordset, ByVal sql As String, _
                            CaseId As String, _
                            Evnt As CaseEvents, _
                            Optional Comments As String = "", _
                            Optional Path As String = "")

10    On Error GoTo UpdateEvent_Error

20    tb!EventId = Evnt
30    tb!CaseId = CaseId
40    Select Case Evnt
          Case CaseEvents.CutUpEvent
50            tb!EventDesc = "Cut-Up By"
60        Case CaseEvents.EmbeddedEvent
70            tb!EventDesc = "Embedded By"
80        Case CaseEvents.PiecesCutUp
90            tb!EventDesc = "Pieces After Cut-Up"
100       Case CaseEvents.PiecesEmbedding
110           tb!EventDesc = "Pieces After Embedding"
120       Case CaseEvents.AssistedBy
130           tb!EventDesc = "Assisted By"
140       Case CaseEvents.CuttingBy
150           tb!EventDesc = "Cutting By"
160       Case CaseEvents.InHistology
170           tb!EventDesc = "Case is set to In Histology"
180       Case CaseEvents.WithPathologist
190           tb!EventDesc = "Case is With Pathologist"
200       Case CaseEvents.AwaitingAuthorisation
210           tb!EventDesc = "Case is Awaiting Authorisation"
220       Case CaseEvents.TreeNodeAdded
230           tb!EventDesc = "Added"
240       Case CaseEvents.TreeNodeDeleted
250           tb!EventDesc = "Deleted"
260       Case CaseEvents.DemographicsAdded
270           tb!EventDesc = "Demographics Added"
280       Case CaseEvents.DemographicsEdited
290           tb!EventDesc = "Demographics Edited"
300       Case CaseEvents.GrossEdited
310           tb!EventDesc = "Gross Edited"
320       Case CaseEvents.MicroEdited
330           tb!EventDesc = "Micro Edited"
340       Case CaseEvents.PCodeEdited
350           tb!EventDesc = "P Code Edited"
360       Case CaseEvents.MCodeAdded
370           tb!EventDesc = "M Code Added"
380       Case CaseEvents.QCodeAdded
390           tb!EventDesc = "Q Code Added"
400       Case CaseEvents.CodeDeleted
410           tb!EventDesc = "Code Deleted"
420       Case CaseEvents.Authorised
430           tb!EventDesc = "Authorised"
440       Case CaseEvents.UnAuthorised
450           tb!EventDesc = "UnAuthorised"
460       Case CaseEvents.DiscrepancyAdded
470           tb!EventDesc = "Discrepancy Added"
480       Case CaseEvents.DiscrepancyEdited
490           tb!EventDesc = "Discrepancy Edited"
500       Case CaseEvents.ReportPrinted
510           tb!EventDesc = "Report Printed"
520       Case CaseEvents.Processor
530           tb!EventDesc = "Processor"
540       Case CaseEvents.Disposal
550           tb!EventDesc = "Specimen"
560       Case CaseEvents.TreeNodeEdited
570           tb!EventDesc = "Edited"
          Case CaseEvents.ExtraRequestsRemoved
              tb!EventDesc = "Extra Requests Removed  Reason:"

580   End Select
590   tb!Path = Path
600   tb!Comments = Comments
610   tb!DateTimeOfRecord = Format$(Now, "dd/MM/yyyy hh:mm:ss")
620   tb!UserName = UserName
630   tb.Update

640   Exit Sub

UpdateEvent_Error:

      Dim strES As String
      Dim intEL As Integer

650   intEL = Erl
660   strES = Err.Description
670   LogError "modEventShared", "UpdateEvent", intEL, strES, sql

End Sub




Attribute VB_Name = "modTree"
Option Explicit

' Expand all the nodes in a TreeView
' Takes in a reference to a TreeView
' and a boolean of True (Expand) or
' False (Collapse)
' The Boolean parameter is optional
' and it assumes True (Expand)
Public Sub ExpandAll(tvwExpand As MSComctlLib.TreeView, _
        Optional ByVal blnExpand As Boolean = True)
          Dim nodExpand As MSComctlLib.Node ' declare iteration node
          ' Iterate through all the nodes in the
          ' given TreeView
10        For Each nodExpand In tvwExpand.Nodes
20            nodExpand.Expanded = blnExpand
30        Next
End Sub

Public Function FindNodeByKey(tvtemp As MSComctlLib.TreeView, Key As String) As MSComctlLib.Node

      Dim n As Integer


10    On Error GoTo FindNodeByKey_Error

20    For n = 1 To tvtemp.Nodes.Count
30        If UCase(tvtemp.Nodes(n).Key) = UCase(Key) Then
40            Set FindNodeByKey = tvtemp.Nodes(n)
              'exit for
50        End If
60    Next n



70    Exit Function

FindNodeByKey_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "modTree", "FindNodeByKey", intEL, strES


End Function


Public Function GetNodeLevel(LocationID As String) As Integer

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo GetNodeLevel_Error

20    sql = "Select LocationLevel From CaseTree Where LocationID = " & LocationID
30    RecOpenClient 0, tb, sql

40    If tb.EOF Then
50        GetNodeLevel = -1
60    Else
70        GetNodeLevel = tb!LocationLevel
80    End If



90    Exit Function

GetNodeLevel_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modTree", "GetNodeLevel", intEL, strES


End Function




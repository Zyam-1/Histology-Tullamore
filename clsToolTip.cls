VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'cslToolTip
'By Dipak Auddy
' Mail : auddy@gawab.com
'MODIFICATIONS
'by Roger Gilchrist
' Mail : rojagilkrist@hotmail.com
'Released with Dipak's permission
'
' Defaults:
' DelayInitial   = 500  (1/2 sec)
' DelayAutoPopup = 5000  (5 secs)
' DelayReshow    = 100 (1/10 sec)
' MaxTipWidth    = 0
' all Margins    = 0
'MODIFICATIONS
'moved all the support Types and declates into the class
'makes it much more portable
Private Type NMHDR
  hwndFrom                        As Long
  idFrom                          As Long
  Code                            As Long
End Type
Private Type POINTAPI
  X                               As Long
  Y                               As Long
End Type
Private Type RECT
  Left                            As Long
  Top                             As Long
  Right                           As Long
  Bottom                          As Long
End Type
Private Const WM_USER             As Long = &H400
Private Const TOOLTIPS_CLASS      As String = "tooltips_class32"
Private Const TTS_ALWAYSTIP       As Long = &H1
''Private Const TTS_NOPREFIX       As Long = &H2
#Const WIN32_IE = &H400
Private Type TOOLINFO
  cbSize                          As Long
  uFlags                          As TT_Flags
  hwnd                            As Long
  uId                             As Long
  RECT                            As RECT
  hinst                           As Long
  lpszText                        As String
#If (WIN32_IE >= &H300) Then
  lParam                          As Long
#End If
End Type
Public Enum TT_Flags
  TTF_IDISHWND = &H1
  TTF_CENTERTIP = &H2
  TTF_RTLREADING = &H4
  TTF_SUBCLASS = &H10
#If (WIN32_IE >= &H300) Then
  TTF_TRACK = &H20
  TTF_ABSOLUTE = &H80
  TTF_TRANSPARENT = &H100
  TTF_DI_SETITEM = &H8000&
#End If
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TTF_IDISHWND, TTF_CENTERTIP, TTF_RTLREADING, TTF_SUBCLASS, TTF_TRACK, TTF_ABSOLUTE, TTF_TRANSPARENT
Private TTF_DI_SETITEM
#End If
Public Enum TT_DelayTime
  TTDT_AUTOMATIC = 0
  TTDT_RESHOW = 1
  TTDT_AUTOPOP = 2
  TTDT_INITIAL = 3
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TTDT_AUTOMATIC, TTDT_RESHOW, TTDT_AUTOPOP, TTDT_INITIAL
#End If
Public Enum ttDelayTimeConstants
  ttDelayDefault = TTDT_AUTOMATIC '= 0
  ttDelayInitial = TTDT_INITIAL '= 3
  ttDelayShow = TTDT_AUTOPOP '= 2
  ttDelayReshow = TTDT_RESHOW '= 1
  ttDelayMask = 3
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private ttDelayDefault, ttDelayInitial, ttDelayShow, ttDelayReshow, ttDelayMask
#End If
Public Enum ttMarginConstants
  ttMarginLeft = 0
  ttMarginTop = 1
  ttMarginRight = 2
  ttMarginBottom = 3
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private ttMarginLeft, ttMarginTop, ttMarginRight, ttMarginBottom
#End If
Private Type TTHITTESTINFO
  hwnd                            As Long
  pt                              As POINTAPI
  ti                              As TOOLINFO
End Type
Public Enum TT_Msgs
  TTM_ACTIVATE = (WM_USER + 1)
  TTM_SETDELAYTIME = (WM_USER + 3)
  TTM_RELAYEVENT = (WM_USER + 7)
  TTM_GETTOOLCOUNT = (WM_USER + 13)
  TTM_WINDOWFROMPOINT = (WM_USER + 16)
#If UNICODE Then
  TTM_ADDTOOL = (WM_USER + 50)
  TTM_DELTOOL = (WM_USER + 51)
  TTM_NEWTOOLRECT = (WM_USER + 52)
  TTM_GETTOOLINFO = (WM_USER + 53)
  TTM_SETTOOLINFO = (WM_USER + 54)
  TTM_HITTEST = (WM_USER + 55)
  TTM_GETTEXT = (WM_USER + 56)
  TTM_UPDATETIPTEXT = (WM_USER + 57)
  TTM_ENUMTOOLS = (WM_USER + 58)
  TTM_GETCURRENTTOOL = (WM_USER + 59)
#Else
  TTM_ADDTOOL = (WM_USER + 4)
  TTM_DELTOOL = (WM_USER + 5)
  TTM_NEWTOOLRECT = (WM_USER + 6)
  TTM_GETTOOLINFO = (WM_USER + 8)
  TTM_SETTOOLINFO = (WM_USER + 9)
  TTM_HITTEST = (WM_USER + 10)
  TTM_GETTEXT = (WM_USER + 11)
  TTM_UPDATETIPTEXT = (WM_USER + 12)
  TTM_ENUMTOOLS = (WM_USER + 14)
  TTM_GETCURRENTTOOL = (WM_USER + 15)
#End If
#If (WIN32_IE >= &H300) Then
  TTM_TRACKACTIVATE = (WM_USER + 17)
  TTM_TRACKPOSITION = (WM_USER + 18)
  TTM_SETTIPBKCOLOR = (WM_USER + 19)
  TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
  TTM_GETDELAYTIME = (WM_USER + 21)
  TTM_GETTIPBKCOLOR = (WM_USER + 22)
  TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
  TTM_SETMAXTIPWIDTH = (WM_USER + 24)
  TTM_GETMAXTIPWIDTH = (WM_USER + 25)
  TTM_SETMARGIN = (WM_USER + 26)
  TTM_GETMARGIN = (WM_USER + 27)
  TTM_POP = (WM_USER + 28)
#End If
#If (WIN32_IE >= &H400) Then
  TTM_UPDATE = (WM_USER + 29)
#End If
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TTM_ACTIVATE, TTM_SETDELAYTIME, TTM_RELAYEVENT, TTM_GETTOOLCOUNT, TTM_WINDOWFROMPOINT
Private TTM_ADDTOOL, TTM_DELTOOL, TTM_NEWTOOLRECT, TTM_GETTOOLINFO, TTM_SETTOOLINFO, TTM_HITTEST
Private TTM_GETTEXT, TTM_UPDATETIPTEXT, TTM_ENUMTOOLS, TTM_GETCURRENTTOOL, TTM_TRACKACTIVATE
Private TTM_TRACKPOSITION, TTM_SETTIPBKCOLOR, TTM_SETTIPTEXTCOLOR, TTM_GETDELAYTIME, TTM_GETTIPBKCOLOR
Private TTM_GETTIPTEXTCOLOR, TTM_SETMAXTIPWIDTH, TTM_GETMAXTIPWIDTH, TTM_SETMARGIN, TTM_GETMARGIN
Private TTM_POP, TTM_UPDATE
#End If
Public Enum TT_Notifications
  TTN_FIRST = -520&
  TTN_LAST = -549&
#If UNICODE Then
  TTN_NEEDTEXT = (TTN_FIRST - 10)
#Else
  TTN_NEEDTEXT = (TTN_FIRST - 0)
#End If
  TTN_SHOW = (TTN_FIRST - 1)
  TTN_POP = (TTN_FIRST - 2)
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TTN_FIRST, TTN_LAST, TTN_NEEDTEXT, TTN_SHOW, TTN_POP
#End If
Private Type NMTTDISPINFO
  hdr                             As NMHDR
  lpszText                        As Long
#If UNICODE Then
  szText                          As String * 160
#Else
  szText                          As String * 80
#End If
  hinst                           As Long
  uFlags                          As Long
#If (WIN32_IE >= &H300) Then
  lParam                          As Long
#End If
End Type
'
' Exported by Comctl32.dll >= v4.00.950
' Ensures that the common control dynamic
' link library (DLL) is loaded.
'
' NOTE: API replaced by InitCommonControlsEx
Private mnlgHwndTT                As Long
Private mnlgMaxTip                As Long
Private m_TTHeader                As String
Private m_ShowHeader              As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SendMessageT Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                         ByVal wMsg As Long, _
                                                                         ByVal wParam As Long, _
                                                                         lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                                                              ByVal lpClassName As String, _
                                                                              ByVal lpWindowName As String, _
                                                                              ByVal dwStyle As Long, _
                                                                              ByVal X As Long, _
                                                                              ByVal Y As Long, _
                                                                              ByVal nWidth As Long, _
                                                                              ByVal nHeight As Long, _
                                                                              ByVal hwndParent As Long, _
                                                                              ByVal hMenu As Long, _
                                                                              ByVal hInstance As Long, _
                                                                              lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal dwLength As Long)

Public Function AddTool(ByRef ctrl As Control, _
                        Optional ByVal strText As String) As Boolean

        Dim ti As TOOLINFO

      'MODIFICATION
      'if a control has no hWnd to use for tooltip
10      If Not fIsWindow(ctrl) Then
      ''MsgBox ctrl.Name & " cannot have a tool tip", vbCritical
20        Exit Function
30      End If
40      If (mnlgHwndTT = 0) Then
50        Exit Function
60      End If
70      If (fGetToolInfo(ctrl.hwnd, ti) = False) Then
80        With ti
90          .cbSize = Len(ti)
      '
      ' TTF_IDISHWND must be set to tell the tooltip
      ' control to retrieve the control's rect from
      ' it's hWnd specified in uId.
      '
100         .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
110         .hwnd = ctrl.Container.hwnd
120         .uId = ctrl.hwnd
130         If Len(strText) > 0 Then
      'MODIFICATION
      'sending a blank string to ToolText deletes the control from the collection
      'MODIFICATION
      'add ToolTipHeader to strText if requested
140           If m_ShowHeader Then
150             If LenB(m_TTHeader) Then
160               strText = m_TTHeader & vbNewLine & strText
170             End If
180           End If
190           .lpszText = strText
      '            Else
      '                .lpszText = "Tool" & ToolCount + 1
200         End If
      '
      ' Maintain the maximun tip text
      ' length for fGetToolInfo.
      '
210         mnlgMaxTip = fMax(mnlgMaxTip, Len(.lpszText) + 1)
220       End With
      '
      ' Returns 1 on success, 0 on failure
      '
230       AddTool = SendMessageT(mnlgHwndTT, TTM_ADDTOOL, 0, ti)
240     End If

End Function

Public Property Get BackColor() As OLE_COLOR

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
      '
      ' OLE_COLOR is defined in stdole2.tlb
      '
40      BackColor = SendMessageT(mnlgHwndTT, TTM_GETTIPBKCOLOR, 0, 0)

End Property

Public Property Let BackColor(clr As OLE_COLOR)

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      Call SendMessageT(mnlgHwndTT, TTM_SETTIPBKCOLOR, clr, 0)

End Property

Private Sub Class_Initialize()

10      m_ShowHeader = False

End Sub

Private Sub Class_Terminate()

10      If mnlgHwndTT > 0 Then
20        Call DestroyWindow(mnlgHwndTT)
30      End If

End Sub

Public Function Create(ByRef frm As Form) As Boolean

10      If (mnlgHwndTT = 0) Then
20        Call InitCommonControls
      '
      ' The hwndParent param lets the tooltip window
      ' be owned by the specified form and be destroyed
      ' along with it. We'll cleanup in Class_Terminate anyway.
      ' No WS_EX_TOPMOST or TTS_ALWAYSTIP per Win95 UI rules.
      '
30        mnlgHwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP, 0, 0, 0, 0, frm.hwnd, 0, App.hInstance, ByVal 0)
40      End If
50      Create = CBool(mnlgHwndTT)

End Function

Public Property Get DelayTime(dwType As ttDelayTimeConstants) As Long

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      DelayTime = SendMessageT(mnlgHwndTT, TTM_GETDELAYTIME, (dwType And ttDelayMask), 0&)

End Property

Public Property Let DelayTime(dwType As ttDelayTimeConstants, _
                              dwMilliSecs As Long)

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      Call SendMessageT(mnlgHwndTT, TTM_SETDELAYTIME, (dwType And ttDelayMask), ByVal dwMilliSecs)
      ' no rtn val

End Property

Private Function fGetStrFromBuffer(ByVal strValue As String) As String

10      If InStr(strValue, vbNullChar) Then
20        fGetStrFromBuffer = Left$(strValue, InStr(strValue, vbNullChar) - 1)
30       Else
      '
      ' If strValue had no null char, the Left$ function
      ' above would rtn a zero length string ("").
      '
40        fGetStrFromBuffer = strValue
50      End If

End Function

Private Function fGetToolInfo(ByVal lnghwndTool As Long, _
                              ti As TOOLINFO, _
                              Optional ByVal fGetText As Boolean = False) As Boolean

        Dim nItems As Long
        Dim i      As Long

10      ti.cbSize = Len(ti)
20      If fGetText Then
30        ti.lpszText = String$(mnlgMaxTip, 0)
40      End If
50      nItems = ToolCount
60      For i = 0 To nItems - 1
      '
      ' Returns 1 on success, 0 on failure.
      '
70        If SendMessageT(mnlgHwndTT, TTM_ENUMTOOLS, (i), ti) Then
80          If (lnghwndTool = ti.uId) Then
90            fGetToolInfo = True
100           Exit Function
110         End If
120       End If
130     Next i

End Function

Public Function fIsWindow(ByRef ctrl As Control) As Boolean

10      On Error GoTo ErrorHandler
20      fIsWindow = CBool(ctrl.hwnd)
ErrorHandler:

End Function

Private Function fLowWord(ByVal lngValue As Long) As Long

      '
      ' Returns the low-order word from a 32-bit value.
      '

10      Call MoveMemory(fLowWord, lngValue, 2)

End Function

Private Function fMax(ByVal lngParm1 As Long, _
                      ByVal lngParm2 As Long) As Long

      '
      ' Returns the larger of the two values.
      '

10      If lngParm1 > lngParm2 Then
20        fMax = lngParm1
30       Else
40        fMax = lngParm2
50      End If

End Function

Public Property Get ForeColor() As OLE_COLOR

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      ForeColor = SendMessageT(mnlgHwndTT, TTM_SETTIPTEXTCOLOR, 0, 0)

End Property

Public Property Let ForeColor(clr As OLE_COLOR)

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      Call SendMessageT(mnlgHwndTT, TTM_SETTIPTEXTCOLOR, clr, 0)   ' no rtn val

End Property

Public Function HasToolTip(ByRef ctrl As Control) As Boolean

      'MODIFICATION
      'Test if control already has aToolTip
      'See Command1_Click on demo from for why

        Dim ti As TOOLINFO

10      HasToolTip = fGetToolInfo(ctrl.hwnd, ti)

End Function

Public Property Get hwnd() As Long

10      hwnd = mnlgHwndTT

End Property

Public Property Get Margin(dwType As ttMarginConstants) As Long

        Dim rc As RECT

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      Call SendMessageT(mnlgHwndTT, TTM_GETMARGIN, 0, rc)
50      Select Case dwType
         Case ttMarginLeft
60        Margin = rc.Left
70       Case ttMarginTop
80        Margin = rc.Top
90       Case ttMarginRight
100       Margin = rc.Right
110      Case ttMarginBottom
120       Margin = rc.Bottom
130     End Select

End Property

Public Property Let Margin(dwType As ttMarginConstants, _
                           ByVal cPixels As Long)

        Dim rc As RECT

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      Call SendMessageT(mnlgHwndTT, TTM_GETMARGIN, 0, rc)
50      Select Case dwType
         Case ttMarginLeft
60        rc.Left = cPixels
70       Case ttMarginTop
80        rc.Top = cPixels
90       Case ttMarginRight
100       rc.Right = cPixels
110      Case ttMarginBottom
120       rc.Bottom = cPixels
130     End Select
140     Call SendMessageT(mnlgHwndTT, TTM_SETMARGIN, 0, rc)

End Property

Public Property Get MaxTipWidth() As Long

      '
      ' If MaxTipWidth is -1, there is no word wrapping and
      ' text control characters are printed and not
      ' evaluated (i.e. a vbCrLf shows up as "||")
      '

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      MaxTipWidth = SendMessageT(mnlgHwndTT, TTM_GETMAXTIPWIDTH, 0, ByVal 0&)
        'MaxTipWidth = fLowWord(SendMessageT(mnlgHwndTT, TTM_GETMAXTIPWIDTH, 0, 0))

End Property

Public Property Let MaxTipWidth(ByVal lngWidth As Long)

      '
      ' If MaxTipWidth is -1, there is no word wrapping and
      ' text control characters are printed and not
      ' evaluated (i.e. a vbCrLf shows up as "||")
      '

10      If mnlgHwndTT = 0 Then
20        Exit Property
30      End If
40      If lngWidth < 1 Then
50        lngWidth = -1
60      End If
70      Call SendMessageT(mnlgHwndTT, TTM_SETMAXTIPWIDTH, 0, ByVal lngWidth)

End Property

Public Function RemoveTool(ByRef ctrl As Control) As Boolean

        Dim ti As TOOLINFO

10      If (mnlgHwndTT = 0) Then
20        Exit Function
30      End If
40      If fGetToolInfo(ctrl.hwnd, ti) Then
50        Call SendMessageT(mnlgHwndTT, TTM_DELTOOL, 0, ti)
60        RemoveTool = True
70      End If

End Function

Public Property Get ToolCount() As Long

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      ToolCount = SendMessageT(mnlgHwndTT, TTM_GETTOOLCOUNT, 0, 0)

End Property

Public Property Get ToolText(ByRef ctrl As Control) As String

        Dim ti As TOOLINFO

10      If (mnlgHwndTT = 0) Then
20        Exit Property
30      End If
40      If fGetToolInfo(ctrl.hwnd, ti, True) Then
50        ToolText = ToolText & fGetStrFromBuffer(ti.lpszText)
60      End If

End Property

Public Property Let ToolText(ByRef ctrl As Control, _
                             ByVal strText As String)

          Dim ti As TOOLINFO

10        If (mnlgHwndTT = 0) Then
20            Exit Property
30        End If
          'MODIFICATION
          'sending a blank string to ToolText deletes the control from the collection
          'If LenB(strText) = 0 Then
          'RemoveTool ctrl
          'Exit Property
          'End If
          'MODIFICATION
          'add ToolTipHeader to strText if requested
40        If m_ShowHeader Then
50            If LenB(m_TTHeader) Then
60                strText = m_TTHeader & vbNewLine & strText
70            End If
80        End If
90        If fGetToolInfo(ctrl.hwnd, ti) Then
100           ti.lpszText = strText
110           mnlgMaxTip = fMax(mnlgMaxTip, Len(strText) + 1)
              '
              ' The tooltip won't appear for the control
              ' if lpszText is an empty string
              '
120           Call SendMessageT(mnlgHwndTT, TTM_UPDATETIPTEXT, 0, ti)
130       End If

End Property

Public Property Get ToolTipHandle() As Long

10      ToolTipHandle = mnlgHwndTT

End Property

Public Property Get ToolTipHeader() As String

10      ToolTipHeader = m_TTHeader

End Property

Public Property Let ToolTipHeader(ByVal vNewValue As String)

      'MODIFICATION
      ' allows you to set a standard opening header for all tooltips

10      m_TTHeader = vNewValue
      'if you supply a Header then set it to default On
20      m_ShowHeader = True

End Property

Public Property Get ToolTipHeaderShow() As Boolean

10      ToolTipHeaderShow = m_ShowHeader

End Property

Public Property Let ToolTipHeaderShow(ByVal vNewValue As Boolean)

      'MODIFICATION
      ' turn the ToolTipHeader on/off on a per control basis

10      m_ShowHeader = vNewValue

End Property

':)Code Fixer V2.1.8 (9/06/2004 2:59:58 PM) 198 + 413 = 611 Lines Thanks Ulli for inspiration and lots of code.


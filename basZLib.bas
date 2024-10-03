Attribute VB_Name = "basZLib"
'*****************************************************
'* ZLib.bas                                          *
'* By: W-Buffer (Carlos Daniel Ruvalcaba Valenzuela) *
'*     Iridium Studios.
'* Web: http://istudios.virtualave.net               *
'* Mail: chadruva@hotmail.com                        *
'* Thanks to: the ZLib.dll guys! :)                  *
'*                                                   *
'* NOTES: - You need to have ZLib.dll in             *
'*        your System Folder.                        *
'*        - You need to have the ZLib.dll            *
'*        Version 1.1.3.1                            *
'*        - Fell Free to do with this bas whatever   *
'*        you want (Steal, Copy, etc.)               *
'*****************************************************
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function compress2 Lib "zlib2.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal Level As Long) As Long
Public Declare Function uncompress Lib "zlib2.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Public Sub CompressBytes(Bytes() As Byte)

      Dim BuffSize As Long
      Dim TBuff() As Byte

10    On Error GoTo CompressBytes_Error

20    BuffSize = UBound(Bytes) + 1
30    BuffSize = BuffSize + (BuffSize * 1.01) + 12
40    ReDim TBuff(BuffSize)

50    compress2 TBuff(0), BuffSize, Bytes(0), UBound(Bytes) + 1, 9

60    ReDim Bytes(BuffSize - 1)

70    CopyMemory Bytes(0), TBuff(0), BuffSize

80    Exit Sub
 
CompressBytes_Error:

      Dim strES As String
      Dim intEL As Integer


90    intEL = Erl
100   strES = Err.Description
110   LogError "basZLib", "CompressBytes", intEL, strES


End Sub

Public Sub UnCompressBytes(Bytes() As Byte, OriginalSize As Long)

      Dim BuffSize As Long
      Dim TBuff() As Byte

10    On Error GoTo UnCompressBytes_Error

20    BuffSize = OriginalSize
30    BuffSize = BuffSize + (BuffSize * 1.01) + 12
40    ReDim TBuff(BuffSize)

50    uncompress TBuff(0), BuffSize, Bytes(0), UBound(Bytes) + 1

60    ReDim Bytes(BuffSize - 1)

70    CopyMemory Bytes(0), TBuff(0), BuffSize


80    Exit Sub

UnCompressBytes_Error:

      Dim strES As String
      Dim intEL As Integer


90    intEL = Erl
100   strES = Err.Description
110   LogError "basZLib", "UnCompressBytes", intEL, strES


End Sub

Public Function CompressFile(strSourceFile As String) As String
    
      Dim s As String

10    On Error GoTo CompressFile_Error

20    Open strSourceFile For Binary Access Read As 1

      Dim OriginalSize As Long
30    OriginalSize = LOF(1)
40    ReDim buffer(OriginalSize - 1) As Byte
50    Get 1, , buffer
    
60    CompressBytes buffer
    
70    s = buffer
80    s = Format(OriginalSize, "000000") & s

90    CompressFile = s
100   Close

110   Exit Function

CompressFile_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "basZLib", "CompressFile", intEL, strES

End Function

Public Sub UnCompressToFile(ByVal src As String, Optional strFileName As String = "")

      Dim OriginalSize As Long
      Dim f As Long

10    On Error GoTo UnCompressToFile_Error

20    If Trim(src) = "" Then
30        Exit Sub
40    End If

50    OriginalSize = Val(Left(src, 6))

60    src = Mid(src, 7)

70    ReDim buff(0 To Len(src) - 1) As Byte

80    buff = src
    
90    UnCompressBytes buff, OriginalSize
    
100   f = FreeFile()
110   If strFileName = "" Then
120       Open "C:\UncompressedImage.bmp" For Binary Access Write As #f
130   Else
140       Open strFileName For Binary Access Write As #f
150   End If
160   Put #f, , buff
170   Close #f
180   f = 0

190   Exit Sub

UnCompressToFile_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "basZLib", "UnCompressToFile", intEL, strES

End Sub

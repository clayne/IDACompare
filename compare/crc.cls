VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'the C dll CRC routines actually arent that much faster, but they are DEP safe
'(1200 function idb went from 4.8secs to 4.2), unicode version 4.0
'native VB routines were horribly slow (like 13 seconds)

Private Declare Function crc32 Lib "crc32.dll" (ByRef b As Byte, ByVal sz As Long) As Long
Private Declare Function unicode_crc32 Lib "crc32.dll" (ByVal b As Long, ByVal sz As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private hLib As Long

Function crc(buf As String) As String
Attribute crc.VB_UserMemId = 0

    Dim v As Long
    Dim l As Long
    
    'Dim b() As Byte
    'b() = StrConv(buf, vbFromUnicode, &H409)
    'v = crc32(b(0), UBound(b) + 1)
    
    l = Len(buf)
    If l = 0 Then Exit Function
    v = unicode_crc32(StrPtr(buf), l)
    crc = Hex(v)
    
End Function

Public Function init() As Boolean
    If hLib = 0 Then Class_Initialize
    init = Not (hLib = 0)
End Function

Private Sub Class_Initialize()
    
    Dim p As String
    
    p = App.path & "\crc32.dll"
    If Not FileExists(p) Then p = App.path & "\..\crc32.dll"
    If Not FileExists(p) Then p = App.path & "\..\..\crc32.dll"
    hLib = LoadLibrary(p)
    
    If hLib = 0 Then
        MsgBox "crc32.dll not found?!" & p
    End If
    
End Sub

Private Sub Class_Terminate()
    If hLib <> 0 Then FreeLibrary hLib
End Sub


' CRC Routines are a direct translation from:
' RFC 1952 - GZIP file format specification version 4.3
' Copyright (c) 1996 L. Peter Deutsch
'
' Permission is granted to copy and distribute this document for any
' purpose and without charge, including translations into other
' languages and incorporation into compilations, provided that the
' copyright notice and this notice are preserved, and that any
' substantive changes or deletions from the original are clearly
' marked.
'
' 12.12.05 - CRC sample code translated to VB6
'            Author: david@idefense.com

'Private Declare Function CallAsm Lib "user32" Alias "CallWindowProcA" (ByRef lpBytes As Any, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
'Dim crc_table(256) As Long
'Dim crc_table_computed As Boolean
'
'Private shr_b() As Byte
'Private shrx_b() As Byte
'
'Private Sub make_crc_table()
'    Dim c As Long, n As Long, k As Long
'
'    For n = 0 To 255
'          c = n
'          For k = 0 To 7
'                If c And 1 Then
'                     c = &HEDB88320 Xor Shr(c)
'                Else
'                    c = Shr(c)
'                End If
'           Next
'          crc_table(n) = c
'    Next
'
'    crc_table_computed = True
'
'End Sub
'
'
'Private Sub Class_Initialize()
'
'    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
'    'D1E8           SHR EAX,1
'    'C2 10 00       RETN 10h
'    Const sr As String = "8B 45 0C D1 E8 C2 10 00"
'    shr_b() = toBytes(sr)
'
'    '8B45 0C        MOV EAX,DWORD PTR SS:[EBP+12]
'    'C1E8 02        SHR EAX,2
'    'C2 10 00       RETN 10h
'    Const sr2 As String = "8B 45 0C C1 E8 02 C2 10 00"
'    shrx_b = toBytes(sr2)
'
'    make_crc_table
'
'End Sub
'
'Function Shr(x As Long) As Long
'    Shr = CallAsm(shr_b(0), x, 0, 0, 0)
'End Function
'
'Function ShrX(x As Long, y As Byte) As Long
'    shrx_b(5) = y
'    ShrX = CallAsm(shrx_b(0), x, 0, 0, 0)
'End Function
'
'Function crc(buf As String) As String
'    Dim c As Long, n As Long, x As Long
'
'    c = -1
'    If Not crc_table_computed Then Class_Initialize
'
'    For n = 1 To Len(buf)
'        x = Asc(Mid(buf, n, 1))
'        c = crc_table((c Xor x) And &HFF) Xor ShrX(c, 8)
'    Next
'
'    crc = Hex(c Xor &HFFFFFFFF)
'
'End Function
'
'Private Function toBytes(x As String) As Byte()
'    Dim tmp() As String
'    Dim fx() As Byte
'    Dim i As Long
'
'    tmp = Split(x, " ")
'    ReDim fx(UBound(tmp))
'
'    For i = 0 To UBound(tmp)
'        fx(i) = CInt("&h" & tmp(i))
'    Next
'
'    toBytes = fx()
'
'End Function
'
'



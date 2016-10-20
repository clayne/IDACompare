VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function crc32 Lib "crc32.dll" (ByRef b As Byte, ByVal sz As Long) As Long
Private Declare Function unicode_crc32 Lib "crc32.dll" (ByVal b As Long, ByVal sz As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private hLib As Long

Private Sub Form_Unload(Cancel As Integer)
    FreeLibrary hLib
End Sub

Private Function init() As Boolean
    Dim p As String
    
    p = App.path & "\crc32.dll"
    If Not FileExists(p) Then p = App.path & "\..\crc32.dll"
    If Not FileExists(p) Then p = App.path & "\..\..\crc32.dll"
    hLib = LoadLibrary(p)
    
    If hLib = 0 Then
        MsgBox "crc32.dll not found?!" & p
    Else
        init = True
    End If
    
End Function

Private Sub Form_Load()
        
        Dim b() As Byte
        Dim v As Long
        Dim t As String
        t = "test"
        
        If Not init Then Exit Sub
        
        b() = StrConv(t, vbFromUnicode, &H409)
        v = crc32(b(0), UBound(b) + 1)
        Me.Caption = "crc32(test) = " & Hex(v)
        
        v2 = unicode_crc32(StrPtr(t), Len(t))
        Me.Caption = Me.Caption & " =? " & Hex(v2)
        
        
End Sub


Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


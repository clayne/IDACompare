VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Author: david@idefense.com <david@idefense.com, dzzie@yahoo.com>
'
'License: Copyright (C) 2005 iDefense.com, A Verisign Company
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

Public Length As Long
Public Calls As Long
Public matched As Boolean
Public Name As String
Public autoid As Long
Public index As Long
Public Jumps As Long
Public Pushs As Long
Public mCRC As String
Public esp As Long
Public MatchMethod As String
Public fxCalls As New Collection   'of functions by name
Public Constants As New Collection
Public labels As New Collection
Public Standardized As String
Public strings As New Collection


Function StandardizeAsm(ByVal asm As String) As String
    
    Dim tmp() As String
    Dim cmd() As String
    Dim x As Long
    Dim j As Long
    Dim i As Long
    Dim k
    Dim t As String
    Dim a As Long
    Dim str As String
    
    tmp = Split(asm, vbCrLf)
    
    'quick profile of function elements
    'hit counts for: calls, pushs, jxxs, esp size, and apis called
    For j = 0 To UBound(tmp)
            
            'If InStr(tmp(j), "28955B88h") > 0 Then Stop
            
            Select Case left(tmp(j), 3)
                Case "cmp", "and", "xor", "pus", "mov", "lea", "sub": CheckConstant tmp(j)
            End Select
            
            If InStr(tmp(j), "call") > 0 Then
                Calls = Calls + 1
                AddCall tmp(j)
            ElseIf InStr(tmp(j), "push") > 0 Then
                Pushs = Pushs + 1
            ElseIf left(tmp(j), 1) = "j" Then
                Jumps = Jumps + 1
            ElseIf CStr(tmp(j)) Like "sub*esp,*" Then
                k = trim(Mid(tmp(j), InStr(tmp(j), "esp,") + 4))
                If InStr(k, "h") > 0 Then
                    'a = FirstOccurance(k, "+,-,*", str)
                    'If a > 0 Then
                    '    str = Mid(k, a + 1)
                    'if its a mem offset we dont know what the real value is...ignore
                    If InStr(k, "[") < 1 Then esp = CLng("&h" & Replace(k, "h", ""))
                Else
                    If IsNumeric(k) Then esp = CLng(k)
                End If
            ElseIf right(trim(tmp(j)), 1) = ":" Then
                labels.Add Mid(trim(tmp(j)), 1, Len(trim(tmp(j))) - 1)
            End If
            
            a = InStr(tmp(j), ";")
            If a > 0 And a < Len(tmp(j)) Then
                str = trim(Mid(tmp(j), a + 1))
                AddUniqueString Replace(str, """", Empty)
            End If
                
    Next
        
    Standardize tmp
    'tmp = sort.AlphaSort(tmp) 'this would wack natural ordering
    
    Standardized = Join(tmp, vbCrLf)
    Standardized = Replace(Standardized, vbCrLf & vbCrLf, vbCrLf) 'strip empty lines (where labels were)
    mCRC = crc(Standardized)
    StandardizeAsm = mCRC
    
    'If strings.Count > 0 Then Debug.Print Name & " added " & strings.Count & " strings"
    
    'should we set standardized as empty now to free up memory? I dont think we use it again..
    
End Function

Function StringExists(s, Optional exact As Boolean = True)
        On Error Resume Next
        Dim v
        If exact Then
            v = strings(s)
            StringExists = (Err.Number = 0)
        Else
            If Len(s) < 2 Then Exit Function 'to short for good results...
            For Each v In strings
                If InStr(1, v, s, vbTextCompare) > 0 Then
                    If Abs(Len(v) - Len(s)) < 5 Then StringExists = True
                End If
            Next
        End If
        'If StringExists Then Stop
End Function

Function DumpStrings()
    Dim tmp() As String
    Dim v
    For Each v In strings
        push tmp, v
    Next
    DumpStrings = Join(tmp, vbCrLf)
End Function

Private Function AddUniqueString(s As String)
    On Error Resume Next
    If Len(s) = 0 Then Exit Function
    strings.Add s, s
    'If Err.Number = 0 Then Debug.Print Name & ">" & s
End Function

Function ReHash(ByVal asm)
    Dim tmp() As String
    If IsNull(asm) Then Exit Function
    tmp = Split(asm, vbCrLf)
    Standardize tmp, , False
    ReHash = crc(Join(tmp, vbCrLf))
    mCRC = ReHash
End Function

Private Sub AddCall(asm As String)
    
    Dim a As Long
    Dim t As String
    
    'disabled 7.16.17
    'If InStr(asm, "sub_") > 1 Then Exit Sub 'its an internal function reference no use for diffing..offsets change.
    
    If InStr(asm, "ds:") > 0 Then 'is api style call
        t = Replace(asm, "ds:", "")
        t = Replace(t, "call ", "")
        If InStr(t, ";") > 0 Then t = Mid(t, 1, InStr(t, ";") - 1)
        If Len(trim(t)) > 3 Then 'ignore call reg32
            fxCalls.Add trim(t)
        End If
    ElseIf InStr(asm, ";") Then 'there is a comment its probably an "call edi ;Sleep" format
        t = trim(Mid(asm, InStr(asm, ";") + 1))
        fxCalls.Add trim(t)
    Else
        a = InStr(asm, "call ")
        If a > 0 Then
            t = Mid(asm, a + 5)
            fxCalls.Add trim(t)
        End If
        
    End If
               
                
End Sub

Private Sub Standardize(tmp() As String, Optional stdJmps As Boolean = True, Optional stripArgs As Boolean = True)
    Dim i As Long, x, cmd
    
    For i = 0 To UBound(tmp)
        If right(tmp(i), 1) = ":" Then
            'its a label strip it
            tmp(i) = Empty
        Else
            tmp(i) = Replace(tmp(i), vbTab, " ")
            tmp(i) = Replace(tmp(i), "  ", " ")
            If Len(tmp(i)) > 0 Then
                cmd = Split(tmp(i), " ")
                
                If stdJmps Then
                    If VBA.left(cmd(0), 1) = "j" And cmd(0) <> "jmp" Then
                        cmd(0) = "jxx" 'standardize all conditional jumps
                    End If
                    If Not stripArgs Then tmp(i) = Join(cmd, " ")
                End If
                
                If stripArgs Then tmp(i) = cmd(0)
                
                x = InStr(tmp(i), ";")
                If x > 0 Then tmp(i) = Mid(tmp(i), 1, x - 1)
                tmp(i) = trim(tmp(i))
            End If
        End If
    Next
    
End Sub

Private Sub CheckConstant(ByVal asm)
    Dim tmp() As String
    Dim i As Long
    Dim a As Long
    
    a = InStr(asm, ";")
    If a > 0 Then asm = Mid(asm, 1, a)
    
    If right(asm, 1) = "]" Then
        asm = Mid(asm, 1, Len(asm) - 1)
        asm = Replace(asm, "+", " ")
        asm = Replace(asm, "-", " ")
        asm = Replace(asm, "*", " ")
    End If
    
    asm = Replace(asm, vbTab, " ")
    
    While InStr(asm, "  ") > 0
        asm = Replace(asm, "  ", " ")
    Wend

    asm = trim(asm)
    tmp = Split(asm, " ")
    
    'If InStr(asm, ",") Then i = 2 Else i = 1
    'If UBound(tmp) < i Then Exit Sub
   
    tmp(i) = trim(tmp(UBound(tmp)))
    If right(tmp(i), 1) = "h" And Len(tmp(i)) > 3 Then
        If KeyExistsInCollection(Constants, "key:" & tmp(i)) Then
            Constants.Add tmp(i)
        Else
           Constants.Add tmp(i), "key:" & tmp(i)
        End If
    End If
    
End Sub

Function ConstantExists(Key) As Boolean
    ConstantExists = KeyExistsInCollection(Constants, "key:" & Key)
End Function

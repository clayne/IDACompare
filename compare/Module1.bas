Attribute VB_Name = "Module1"
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



Global crc As New clsCrc
Global Sort As New CAlphaSort
Global dlg As New clsCmnDlg
Global fso As New CFileSystem2
Global wHash As New CWinHash
Global HighLightRunning As Boolean

Const LANG_US = 1049



Sub rtfHighlightDecompile(c_src As String, tb As RichTextBox)
    
    On Error Resume Next
    
    HighLightRunning = True
    
    Dim tmp() As String
    Dim x, i As Long
    
    tb.Text = " "
    tb.selStart = 0
    tb.selLength = 1
    tb.SelColor = vbBlack
    tb.SelBold = False
    
    tb.Text = c_src
    tmp() = Split(c_src, vbCrLf)
    
    rtf.SetWindowUpdate tb
    
    Dim curPos As Long
    Dim a As Long
    
   'color code comments..
    For i = 0 To UBound(tmp)
        x = trim(tmp(i))
        
        a = InStr(tmp(i), "//")
        If a > 0 Then 'comment
            tb.selStart = curPos + a
            tb.selLength = Len(tmp(i)) - a
            tb.SelColor = &H8000&
        End If
        
        curPos = curPos + Len(tmp(i)) + 2 'for crlf
    Next
            
            
    'now we search for and highlight some C keywords in the function..
    Dim k
    Dim eol As Long
    Dim nextSpace As Long
    Dim keywords() As String
    
    keywords = Split("return,int,char,struct,HANDLE,if,else,{,},while,do,break", ",")
    
    For Each k In keywords
        a = 0
        Do
            a = tb.Find(k, a, , rtfWholeWord)
            If a > -1 Then
                eol = InStr(a, tb.Text, vbCrLf)
                nextSpace = InStr(a + 1, tb.Text, " ")
                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
                nextSpace = InStr(a + 1, tb.Text, "(")
                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
                nextSpace = InStr(a + 1, tb.Text, "{")
                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
                nextSpace = InStr(a + 1, tb.Text, "{")
                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
                nextSpace = InStr(a + 1, tb.Text, ";")
                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
                tb.selStart = a
                tb.selLength = eol - a
                tb.SelColor = vbBlue
                tb.SelBold = True
                a = a + tb.selLength
            End If
        Loop While a > 0
    Next
      
    tb.selStart = 0
    rtf.SetWindowUpdate tb, False
    
    HighLightRunning = False
    
End Sub


Sub rtfHighlightAsm(asm As String, c As CFunction, tb As RichTextBox)
    
    On Error Resume Next
    
    HighLightRunning = True
    
    Dim tmp() As String
    Dim x, i As Long
    Const indentLen = 2
    
    'remove all old formatting
    tb.Text = " "
    tb.selStart = 0
    tb.selLength = 1
    tb.SelColor = vbBlack
    tb.SelBold = False
    
    If c Is Nothing Then 'functions coming from lvExact dont have the class
        Set c = New CFunction
        c.StandardizeAsm asm
    End If
    
    tmp() = Split(asm, vbCrLf)
    
    'first we add line breaks for comments and indents for code..
    For i = 0 To UBound(tmp)
        x = trim(tmp(i))
        If right(x, 1) = ":" Then 'label
            tmp(i) = vbCrLf & x
        Else
            tmp(i) = Space(indentLen) & x
        End If
    Next
    
    tb.Text = Join(tmp, vbCrLf) 'save to textbox..
    
    rtf.SetWindowUpdate tb
    
    Dim curPos As Long
    Dim a As Long
    
   'now we highlight
    For i = 0 To UBound(tmp)
        x = trim(tmp(i))
        
        If left(x, 1) = "j" Then 'isjump
            tb.selStart = curPos
            tb.selLength = Len(tmp(i))
            tb.SelColor = vbRed
        ElseIf left(x, 4) = "call" Then 'iscall
            tb.selStart = curPos
            tb.selLength = Len(tmp(i))
            tb.SelColor = vbBlue
        End If
        
        a = InStr(tmp(i), ";")
        If a > 0 Then 'comment
            tb.selStart = curPos + a
            tb.selLength = Len(tmp(i)) - a
            tb.SelColor = &H8000&
        End If
        
        If right(x, 1) = ":" Then 'is label
            tb.selStart = curPos
            tb.selLength = Len(tmp(i))
            tb.SelColor = &H404000
            tb.SelBold = True
        End If
        
        curPos = curPos + Len(tmp(i)) + 2 'for crlf
    Next
            
            
    'now we search for and highlight all constants from the function..
    Dim k
    Dim eol As Long
    Dim nextSpace As Long
    
    For Each k In c.Constants
        a = 0
        Do
            a = tb.Find(k, a)
            If a > 0 Then
                eol = InStr(a, tb.Text, vbCrLf)
                nextSpace = InStr(a + 1, tb.Text, " ")
                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
                tb.selStart = a
                tb.selLength = eol - a
                tb.SelBold = True
                a = a + tb.selLength
            End If
        Loop While a > 0
    Next
          
'    For Each k In c.labels 'they are already red we dont need them bold to, to much processing
'        a = 0
'        Do
'            a = tb.Find(k, a)
'            If a > 0 Then
'                eol = InStr(a, tb.Text, vbCrLf)
'                nextSpace = InStr(a + 1, tb.Text, " ")
'                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
'                tb.SelStart = a
'                tb.SelLength = eol - a
'                tb.SelBold = True
'                a = a + tb.SelLength
'            End If
'        Loop While a > 0
'    Next
    
    tb.selStart = 0
    
    rtf.SetWindowUpdate tb, False
    
    HighLightRunning = False
    
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
 
Function pad(v, Optional l As Long = 4, Optional lpad As Boolean = True)
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        If lpad Then
            pad = String(l - x, " ") & v
        Else
            pad = v & String(l - x, " ")
        End If
    Else
hell:
        pad = v
    End If
End Function

Public Sub LV_ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
     On Error Resume Next
    With ListViewControl
       If .SortKey <> Column.index - 1 Then
             .SortKey = Column.index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .sorted = -1
    End With
End Sub

Sub FormPos(fform As Form, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.Name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.Name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Function RandomNum() As Long
    Dim tmp As Long
    On Error GoTo hell
hell:
    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
End Function

Function isWithin(cnt As Integer, v1, v2, Optional Min As Integer = 0) As Boolean
    
    Dim low As Long
    Dim high As Long
    
    If v1 <= Min Or v2 <= Min Then Exit Function
    
    If v1 = v2 Then
        isWithin = True
        Exit Function
    End If
    
    low = IIf(v1 < v2, v1, v2)
    
    high = v1
    If low = v1 Then high = v2
    
    If low + cnt >= high Then isWithin = True
    
    
End Function


Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    If Len(val) = 0 Then Exit Function
    If IsObject(c(val)) Then
        KeyExistsInCollection = True
    Else
        t = c(val)
        KeyExistsInCollection = True
    End If
 Exit Function
nope: KeyExistsInCollection = False
End Function

 



Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function ado(sql) As Recordset
    Set ado = Form1.cn.Execute(sql)
End Function


Sub ResetPB(newMax, Optional caption)
    With Form1
        If newMax < 1 Then newMax = 1
        .pb.Max = newMax
        .pb.value = 0
        If Len(caption) > 0 Then
            .Label1.caption = caption
            .Label1.Refresh
        End If
    End With
End Sub

Function GetAllElements(lv As ListView, Optional selOnly As Boolean = False) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem

    For i = 1 To lv.ColumnHeaders.count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next

    push ret, tmp
    push ret, String(50, "-")

    For Each li In lv.ListItems
        If selOnly Then
            If li.Selected Then tmp = li.Text & vbTab
        Else
            tmp = li.Text & vbTab
        End If
        For i = 1 To lv.ColumnHeaders.count - 1
            If selOnly Then
                If li.Selected Then tmp = tmp & li.SubItems(i) & vbTab
            Else
                tmp = tmp & li.SubItems(i) & vbTab
            End If
        Next
        push ret, tmp
    Next

    GetAllElements = Join(ret, vbCrLf)

End Function

Function ReadFile(filename)
  Dim f, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Function WriteFile(path, it) As Boolean 'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte, f As Long
    If FileExists(path) Then Kill path
    f = FreeFile
    b() = StrConv(it, vbFromUnicode, LANG_US)
    Open path For Binary As #f
    Put f, , b()
    Close f
    WriteFile = True
    Exit Function
hell: WriteFile = False
End Function

Function lowest(ParamArray values()) As Long
   Dim Item
   On Error Resume Next
   lowest = &H7FFFFFFF
   For Each Item In values
      lowest = IIf(lowest > Item, Item, lowest)
   Next
   If lowest = &H7FFFFFFF Then lowest = -1
End Function

Function GetCompileTime(Optional ByVal exe As String) As String
    
    Dim f As Long, i As Integer
    Dim stamp As Long, e_lfanew As Long
    Dim base As Date, compiled As Date

    On Error GoTo errExit
    
    If Len(exe) = 0 Then
        exe = App.path & "\" & App.EXEName & ".exe"
    End If
    
    FileLen exe 'throw error if not exist
    
    f = FreeFile
    Open exe For Binary Access Read As f
    Get f, , i
    
    If i <> &H5A4D Then GoTo errExit 'MZ check
     
    Get f, 60 + 1, e_lfanew
    Get f, e_lfanew + 1, i
    
    If i <> &H4550 Then GoTo errExit 'PE check
    
    Get f, e_lfanew + 9, stamp
    Close f
    
    base = DateSerial(1970, 1, 1)
    compiled = DateAdd("s", stamp, base)
    GetCompileTime = Format(compiled, "ddd, mmm d yyyy, h:nn:ss ")
    
    Exit Function
errExit:
    Close f
        
End Function

'Function FirstOccurance(it, ByVal csvFind As String, ByRef outFoundVal) As Long
'    If Len(csvFind) = 0 Then Exit Function
'
'    Dim find() As String, x, lowestOffset As Long, lowestIndex As Long, i As Long, a As Long
'
'    outFoundVal = Empty
'    lowestOffset = MAX_LONG
'    find = Split(csvFind, ",")
'
'    For i = 0 To UBound(find)
'        If Len(find(i)) = 0 Then find(i) = ","
'        a = InStr(1, it, find(i), vbTextCompare)
'        If a > 0 And a < lowestOffset Then
'            lowestOffset = a
'            lowestIndex = i
'        End If
'    Next
'
'    If lowestOffset = MAX_LONG Then Exit Function
'
'    outFoundVal = find(lowestIndex)
'    FirstOccurance = lowestOffset
'
'End Function

'
'Sub ExactCrcMatch()
'
'    Dim li As ListItem
'    Dim lit As ListItem
'    Dim t, u
'    Dim x, i1, i2
'
'    'For Each x In c1
'    For Each lit In lv1.ListItems
'
'        If KeyExistsInCollection(c2, CStr(x)) Then 'exact crc match
'            Set c = c1(x)
'            i1 = c.li.index
'            'i1 = c1i(x)
'            i2 = c2(x)
'
'            Set li = lvExact.ListItems.Add
'            'li.Tag = t1(x) & "," & t2(x)
'
'            li.Text = lv1.ListItems(i1).SubItems(2)
'            li.SubItems(1) = lv2.ListItems(i2).SubItems(2)
'            li.SubItems(4) = "Exact CRC"
'
'            t = lv1.ListItems(i1).SubItems(1)
'            u = lv2.ListItems(i2).SubItems(1)
'
'            lv1.ListItems(i1).Tag = "delete"
'            lv2.ListItems(i2).Tag = "delete"
'            a.Remove lv1.ListItems(i1).SubItems(3)
'            b.Remove lv2.ListItems(i2).SubItems(3)
'
'            If t = u Then
'                li.SubItems(2) = "yes"
'            Else
'                li.SubItems(2) = t & "," & u
'            End If
'
'            li.SubItems(3) = lv1.ListItems(i1).SubItems(3)
'
'        End If
'
'        pb.value = pb.value + 1
'    Next
'
'     Dim i As Long
'
'    'remove matchs from top two list views
'    For i = lv1.ListItems.Count To 1 Step -1
'        Set li = lv1.ListItems(i)
'        If li.Tag = "delete" Then lv1.ListItems.Remove i
'    Next
'
'    For i = lv2.ListItems.Count To 1 Step -1
'        Set li = lv2.ListItems(i)
'        If li.Tag = "delete" Then lv2.ListItems.Remove i
'    Next
'
'End Sub


'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
'
'Public Function Clone(c As Collection) As Collection
'
'    Dim cc As New Collection
'    Dim i As Long
'
'    For i = 1 To c.Count
'        'If IsObject(c(i)) Then Err.Raise 256, , "CollectionEx.Clone: Can not clone a collection with object references"
'        cc.Add c(i), keyForIndex(c, i)
'    Next
'
'    Set Clone = cc
'End Function
'
'Public Function keyForIndex(c As Collection, index As Long) As String
'    ' Get a key based on its index value.  Must be in range, or error.
'    Dim i     As Long
'    Dim Ptr   As Long
'    Dim sKey  As String
'    '
'    If index < 1 Or index > c.Count Then
'        Err.Raise 9
'        Exit Function
'    End If
'    '
'    If index <= c.Count / 2 Then                                ' Start from front.
'        CopyMemory Ptr, ByVal ObjPtr(c) + &H18, 4               ' First item pointer of collection header.
'        For i = 2 To index
'            CopyMemory Ptr, ByVal Ptr + &H18, 4                 ' Next item pointer of collection item.
'        Next i
'    Else                                                        ' Start from end and go back.
'        CopyMemory Ptr, ByVal ObjPtr(c) + &H1C, 4               ' Last item pointer of collection header.
'        For i = c.Count - 1 To index Step -1
'            CopyMemory Ptr, ByVal Ptr + &H14, 4                 ' Previous item pointer of collection item.
'        Next i
'    End If
'    '
'    i = StrPtr(sKey)                                            ' Save string pointer because we're going to borrow the string.
'    CopyMemory ByVal VarPtr(sKey), ByVal Ptr + &H10, 4          ' Key string of collection item.
'    keyForIndex = sKey                                          ' Move key into property's return.
'    CopyMemory ByVal VarPtr(sKey), i, 4                         ' Put string pointer back to keep memory straight.
'End Function



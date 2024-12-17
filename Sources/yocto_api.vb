'/********************************************************************
'*
'* $Id: yocto_api.vb 63704 2024-12-16 10:05:02Z seb $
'*
'* High-level programming interface, common to all modules
'*
'* - - - - - - - - - License information: - - - - - - - - -
'*
'*  Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'*
'*  Yoctopuce Sarl (hereafter Licensor) grants to you a perpetual
'*  non-exclusive license to use, modify, copy and integrate this
'*  file into your software for the sole purpose of interfacing
'*  with Yoctopuce products.
'*
'*  You may reproduce and distribute copies of this file in
'*  source or object form, as long as the sole purpose of this
'*  code is to interface with Yoctopuce products. You must retain
'*  this notice in the distributed source file.
'*
'*  You should refer to Yoctopuce General Terms and Conditions
'*  for additional information regarding your rights and
'*  obligations.
'*
'*  THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'*  WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING
'*  WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS
'*  FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'*  EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'*  INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA,
'*  COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR
'*  SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT
'*  LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'*  CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'*  BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'*  WARRANTY, OR OTHERWISE.
'*
'/********************************************************************/

Imports YHANDLE = System.Int32
Imports YRETCODE = System.Int32
Imports ys8 = System.SByte
Imports ys16 = System.Int16
Imports ys32 = System.Int32
Imports ys64 = System.Int64
Imports yu8 = System.Byte
Imports yu16 = System.UInt16
Imports yu32 = System.UInt32
Imports yu64 = System.UInt64


Imports YDEV_DESCR = System.Int32      REM yStrRef of serial number
Imports YFUN_DESCR = System.Int32      REM yStrRef of serial + (ystrRef of funcId << 16)
Imports yTime = System.UInt32          REM measured in milliseconds
Imports yHash = System.Int16
Imports yBlkHdl = System.Char          REM (yHash << 1) + [0,1]
Imports yStrRef = System.Int16
Imports yUrlRef = System.Int16

Imports System.Runtime.InteropServices
Imports System.Buffer
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Threading

Module yocto_api
  Friend Class yMemoryStream
    Inherits MemoryStream
    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(capacity As Integer)
      MyBase.New(capacity)
    End Sub

    Public Sub Append(c As Char)
      Dim tmp As Char() = {c}
      Dim b As Byte() = YAPI.DefaultEncoding.GetBytes(tmp)
      Write(b, 0, b.Length)
    End Sub

    Public Sub Append(bin As Byte())
      Write(bin, 0, bin.Length)
    End Sub

    Public Sub Append(s As String)
      Dim b As Byte() = YAPI.DefaultEncoding.GetBytes(s)
      Write(b, 0, b.Length)
    End Sub
  End Class

  Public MustInherit Class YJSONContent
    Friend _data As String
    Friend _data_start As Integer
    Protected _data_len As Integer
    Friend _data_boundary As Integer
    Protected _type As YJSONType

    Public Enum YJSONType
      [STRING]
      NUMBER
      ARRAY
      [OBJECT]
    End Enum

    Protected Enum Tjstate
      JSTART
      JWAITFORNAME
      JWAITFORENDOFNAME
      JWAITFORCOLON
      JWAITFORDATA
      JWAITFORNEXTSTRUCTMEMBER
      JWAITFORNEXTARRAYITEM
      JWAITFORSTRINGVALUE
      JWAITFORSTRINGVALUE_ESC
      JWAITFORINTVALUE
      JWAITFORBOOLVALUE
    End Enum

    Public Shared Function ParseJson(data As String, start As Integer, [stop] As Integer) As YJSONContent
      Dim cur_pos As Integer = SkipGarbage(data, start, [stop])
      Dim res As YJSONContent
      If data(cur_pos) = "["c Then
        res = New YJSONArray(data, start, [stop])
      ElseIf data(cur_pos) = "{"c Then
        res = New YJSONObject(data, start, [stop])
      ElseIf data(cur_pos) = """"c Then
        res = New YJSONString(data, start, [stop])
      Else
        res = New YJSONNumber(data, start, [stop])
      End If
      res.parse()
      Return res
    End Function

    Protected Sub New(data As String, start As Integer, [stop] As Integer, type As YJSONType)
      _data = data
      _data_start = start
      _data_boundary = [stop]
      _type = type
    End Sub

    Protected Sub New(type As YJSONType)
      _data = Nothing
    End Sub

    Public Function getJSONType() As YJSONType
      Return _type
    End Function
    Public MustOverride Function parse() As Integer

    Protected Shared Function SkipGarbage(data As String, start As Integer, [stop] As Integer) As Integer
      If data.Length <= start Then
        Return start
      End If
      Dim sti As Char = data(start)
      While start < [stop] AndAlso (sti = ControlChars.Lf OrElse sti = ControlChars.Cr OrElse sti = " "c)
        start += 1
        sti = data(start)
      End While
      Return start
    End Function

    Protected Function FormatError(errmsg As String, cur_pos As Integer) As String
      Dim ststart As Integer = cur_pos - 10
      Dim stend As Integer = cur_pos + 10
      If ststart < 0 Then
        ststart = 0
      End If
      If stend > _data_boundary Then
        stend = _data_boundary
      End If
      If _data Is Nothing Then
        Return errmsg
      End If
      Return (errmsg & Convert.ToString(" near ")) + _data.Substring(ststart, cur_pos - ststart) + _data.Substring(cur_pos, stend - cur_pos)
    End Function

    Public MustOverride Function toJSON() As Byte()
  End Class

  Friend Class YJSONArray
    Inherits YJSONContent
    Private _arrayValue As New List(Of YJSONContent)()

    Public Sub New(data As String, start As Integer, [stop] As Integer)
      MyBase.New(data, start, [stop], YJSONType.ARRAY)
    End Sub

    Public Sub New(data As String)
      Me.New(data, 0, data.Length)
    End Sub

    Public Sub New()
      MyBase.New(YJSONType.ARRAY)
    End Sub

    Public ReadOnly Property Length() As Integer
      Get
        Return _arrayValue.Count
      End Get
    End Property

    Public Overrides Function parse() As Integer
      Dim cur_pos As Integer = SkipGarbage(_data, _data_start, _data_boundary)

      If cur_pos >= _data_boundary OrElse _data(cur_pos) <> "["c Then
        Throw New System.Exception(FormatError("Opening braces was expected", cur_pos))
      End If
      cur_pos += 1
      Dim state As Tjstate = Tjstate.JWAITFORDATA

      While cur_pos < _data_boundary
        Dim sti As Char = _data(cur_pos)
        Dim inc_pos As Boolean = True
        Select Case state
          Case Tjstate.JWAITFORDATA
            If sti = "{"c Then
              Dim jobj As New YJSONObject(_data, cur_pos, _data_boundary)
              Dim len As Integer = jobj.parse()
              cur_pos += len
              _arrayValue.Add(jobj)
              state = Tjstate.JWAITFORNEXTARRAYITEM
              'cur_pos is already incremented
              inc_pos = False
              Exit Select
            ElseIf sti = "["c Then
              Dim jobj As New YJSONArray(_data, cur_pos, _data_boundary)
              Dim len As Integer = jobj.parse()
              cur_pos += len
              _arrayValue.Add(jobj)
              state = Tjstate.JWAITFORNEXTARRAYITEM
              'cur_pos is already incremented
              inc_pos = False
              Exit Select
            ElseIf sti = """"c Then
              Dim jobj As New YJSONString(_data, cur_pos, _data_boundary)
              Dim len As Integer = jobj.parse()
              cur_pos += len
              _arrayValue.Add(jobj)
              state = Tjstate.JWAITFORNEXTARRAYITEM
              'cur_pos is already incremented
              inc_pos = False
              Exit Select
            ElseIf sti = "-"c OrElse (sti >= "0"c AndAlso sti <= "9"c) Then
              Dim jobj As New YJSONNumber(_data, cur_pos, _data_boundary)
              Dim len As Integer = jobj.parse()
              cur_pos += len
              _arrayValue.Add(jobj)
              state = Tjstate.JWAITFORNEXTARRAYITEM
              'cur_pos is already incremented
              inc_pos = False
              Exit Select
            ElseIf sti = "]"c Then
              _data_len = cur_pos + 1 - _data_start
              Return _data_len
            ElseIf sti <> " "c AndAlso sti <> ControlChars.Lf AndAlso sti <> ControlChars.Cr Then
              Throw New System.Exception(FormatError("invalid char: was expecting  "",0..9,t or f", cur_pos))
            End If
            Exit Select
          Case Tjstate.JWAITFORNEXTARRAYITEM
            If sti = ","c Then
              state = Tjstate.JWAITFORDATA
            ElseIf sti = "]"c Then
              _data_len = cur_pos + 1 - _data_start
              Return _data_len
            Else
              If sti <> " "c AndAlso sti <> ControlChars.Lf AndAlso sti <> ControlChars.Cr Then
                Throw New System.Exception(FormatError("invalid char: was expecting ,", cur_pos))
              End If
            End If
            Exit Select
          Case Else
            Throw New System.Exception(FormatError("invalid state for YJSONObject", cur_pos))
        End Select
        If inc_pos Then cur_pos += 1
      End While
      Throw New System.Exception(FormatError("unexpected end of data", cur_pos))
    End Function

    Public Function getYJSONObject(i As Integer) As YJSONObject
      Return DirectCast(_arrayValue(i), YJSONObject)
    End Function

    Public Function getString(i As Integer) As String
      Dim ystr As YJSONString = DirectCast(_arrayValue(i), YJSONString)
      Return ystr.getString()
    End Function

    Public Function [get](i As Integer) As YJSONContent
      Return _arrayValue(i)
    End Function

    Public Function getYJSONArray(i As Integer) As YJSONArray
      Return DirectCast(_arrayValue(i), YJSONArray)
    End Function

    Public Function getInt(i As Integer) As Integer
      Dim ystr As YJSONNumber = DirectCast(_arrayValue(i), YJSONNumber)
      Return ystr.getInt()
    End Function

    Public Function getLong(i As Integer) As Long
      Dim ystr As YJSONNumber = DirectCast(_arrayValue(i), YJSONNumber)
      Return ystr.getLong()
    End Function

    Public Sub put(flatAttr As String)
      Dim strobj As New YJSONString()
      strobj.setContent(flatAttr)
      _arrayValue.Add(strobj)
    End Sub

    Public Overrides Function toJSON() As Byte()
      Dim res As New yMemoryStream()
      res.Append("["c)
      Dim sep As String = ""
      For Each yjsonContent As YJSONContent In _arrayValue
        Dim subres As Byte() = yjsonContent.toJSON()
        res.Append(sep)
        res.Append(subres)
        sep = ","
      Next
      res.Append("]"c)
      Return res.ToArray()
    End Function

    Public Overrides Function ToString() As String
      Dim res As New StringBuilder()
      res.Append("["c)
      Dim sep As String = ""
      For Each yjsonContent As YJSONContent In _arrayValue
        Dim subres As String = yjsonContent.ToString()
        res.Append(sep)
        res.Append(subres)
        sep = ","
      Next
      res.Append("]"c)
      Return res.ToString()
    End Function
  End Class

  Friend Class YJSONString
    Inherits YJSONContent
    Private _stringValue As String

    Public Sub New(data As String, start As Integer, [stop] As Integer)
      MyBase.New(data, start, [stop], YJSONType.[STRING])
    End Sub

    Public Sub New(data As String)
      Me.New(data, 0, data.Length)
    End Sub

    Public Sub New()
      MyBase.New(YJSONType.[STRING])
    End Sub

    Public Overrides Function parse() As Integer
      Dim value As String = ""
      Dim cur_pos As Integer = SkipGarbage(_data, _data_start, _data_boundary)

      If cur_pos >= _data_boundary OrElse _data(cur_pos) <> """"c Then
        Throw New System.Exception(FormatError("double quote was expected", cur_pos))
      End If
      cur_pos += 1
      Dim str_start As Integer = cur_pos
      Dim state As Tjstate = Tjstate.JWAITFORSTRINGVALUE

      While cur_pos < _data_boundary
        Dim sti As Char = _data(cur_pos)
        Select Case state
          Case Tjstate.JWAITFORSTRINGVALUE
            If sti = "\"c Then
              value += _data.Substring(str_start, cur_pos - str_start)
              str_start = cur_pos
              state = Tjstate.JWAITFORSTRINGVALUE_ESC
            ElseIf sti = """"c Then
              value += _data.Substring(str_start, cur_pos - str_start)
              _stringValue = value
              _data_len = (cur_pos + 1) - _data_start
              Return _data_len
            ElseIf Asc(sti) < 32 Then
              Throw New System.Exception(FormatError("invalid char: was expecting string value", cur_pos))
            End If
            Exit Select
          Case Tjstate.JWAITFORSTRINGVALUE_ESC
            value += sti
            state = Tjstate.JWAITFORSTRINGVALUE
            str_start = cur_pos + 1
            Exit Select
          Case Else
            Throw New System.Exception(FormatError("invalid state for YJSONObject", cur_pos))
        End Select
        cur_pos += 1
      End While
      Throw New System.Exception(FormatError("unexpected end of data", cur_pos))
    End Function

    Public Overrides Function toJSON() As Byte()
      Dim res As New yMemoryStream(_stringValue.Length * 2)
      res.Append(""""c)
      For Each c As Char In _stringValue
        Select Case c
          Case """"c
            res.Append("\""")
            Exit Select
          Case "\"c
            res.Append("\\")
            Exit Select
          Case "/"c
            res.Append("\/")
            Exit Select
          Case ControlChars.Back
            res.Append("\b")
            Exit Select
          Case ControlChars.FormFeed
            res.Append("\f")
            Exit Select
          Case ControlChars.Lf
            res.Append("\n")
            Exit Select
          Case ControlChars.Cr
            res.Append("\r")
            Exit Select
          Case ControlChars.Tab
            res.Append("\t")
            Exit Select
          Case Else
            res.Append(c)
            Exit Select
        End Select
      Next
      res.Append(""""c)
      Return res.ToArray()
    End Function

    Public Function getString() As String
      Return _stringValue
    End Function

    Public Overrides Function ToString() As String
      Return _stringValue
    End Function

    Public Sub setContent(value As String)
      _stringValue = value
    End Sub
  End Class


  Friend Class YJSONNumber
    Inherits YJSONContent
    Private _intValue As Long = 0
    Private _doubleValue As Double = 0
    Private _isFloat As Boolean = False

    Public Sub New(data As String, start As Integer, [stop] As Integer)
      MyBase.New(data, start, [stop], YJSONType.NUMBER)
    End Sub

    Public Overrides Function parse() As Integer

      Dim neg As Boolean = False
      Dim start As Integer, dotPos As Integer
      Dim sti As Char
      Dim cur_pos As Integer = SkipGarbage(_data, _data_start, _data_boundary)
      sti = _data(cur_pos)
      If sti = "-"c Then
        neg = True
        cur_pos += 1
      End If
      start = cur_pos
      dotPos = start
      While cur_pos < _data_boundary
        sti = _data(cur_pos)
        If sti = "."c AndAlso _isFloat = False Then
          Dim int_part As String = _data.Substring(start, cur_pos - start)
          _intValue = Convert.ToInt64(int_part)
          _isFloat = True
        ElseIf sti < "0"c OrElse sti > "9"c Then
          Return getnumber(cur_pos, start, neg)
        End If
        cur_pos += 1
      End While
      Return getnumber(cur_pos, start, neg)
    End Function

    Private Function getnumber(cur_pos As Integer, start As Integer, neg As Boolean) As Integer

      Dim numberpart As String = _data.Substring(start, cur_pos - start)
      If _isFloat Then
        _doubleValue = Convert.ToDouble(numberpart)
      Else
        _intValue = Convert.ToInt64(numberpart)
      End If
      If neg Then
        _doubleValue = 0 - _doubleValue
        _intValue = 0 - _intValue
      End If
      Return cur_pos - _data_start
    End Function

    Public Overrides Function toJSON() As Byte()
      If _isFloat Then
        Return YAPI.DefaultEncoding.GetBytes(_doubleValue.ToString())
      Else
        Return YAPI.DefaultEncoding.GetBytes(_intValue.ToString())
      End If
    End Function

    Public Function getLong() As Long
      If _isFloat Then
        Return CLng(_doubleValue)
      Else
        Return _intValue
      End If
    End Function

    Public Function getInt() As Integer
      If _isFloat Then
        Return CInt(_doubleValue)
      Else
        Return CInt(_intValue)
      End If
    End Function

    Public Function getDouble() As Double
      If _isFloat Then
        Return _doubleValue
      Else
        Return _intValue
      End If
    End Function

    Public Overrides Function ToString() As String
      If _isFloat Then
        Return _doubleValue.ToString()
      Else
        Return _intValue.ToString()
      End If
    End Function
  End Class


  Public Class YJSONObject
    Inherits YJSONContent
    ReadOnly parsed As New Dictionary(Of String, YJSONContent)()
    ReadOnly _keys As New List(Of String)(16)

    Public Sub New(data As String)
      MyBase.New(data, 0, data.Length, YJSONType.[OBJECT])
    End Sub

    Public Sub New(data As String, start As Integer, len As Integer)
      MyBase.New(data, start, len, YJSONType.[OBJECT])
    End Sub

    Public Overrides Function parse() As Integer
      Dim current_name As String = ""
      Dim name_start As Integer = _data_start
      Dim cur_pos As Integer = SkipGarbage(_data, _data_start, _data_boundary)

      If _data.Length <= cur_pos OrElse cur_pos >= _data_boundary OrElse _data(cur_pos) <> "{"c Then
        Throw New System.Exception(FormatError("Opening braces was expected", cur_pos))
      End If
      cur_pos += 1
      Dim state As Tjstate = Tjstate.JWAITFORNAME

      While cur_pos < _data_boundary
        Dim sti As Char = _data(cur_pos)
        Dim inc_pos As Boolean = True
        Select Case state
          Case Tjstate.JWAITFORNAME
            If sti = """"c Then
              state = Tjstate.JWAITFORENDOFNAME
              name_start = cur_pos + 1
            ElseIf sti = "}"c Then
              _data_len = cur_pos + 1 - _data_start
              Return _data_len
            Else
              If sti <> " "c AndAlso sti <> ControlChars.Lf AndAlso sti <> ControlChars.Cr Then
                Throw New System.Exception(FormatError("invalid char: was expecting """, cur_pos))
              End If
            End If
            Exit Select
          Case Tjstate.JWAITFORENDOFNAME
            If sti = """"c Then
              current_name = _data.Substring(name_start, cur_pos - name_start)

              state = Tjstate.JWAITFORCOLON
            Else
              If Asc(sti) < 32 Then
                Throw New System.Exception(FormatError("invalid char: was expecting an identifier compliant char", cur_pos))
              End If
            End If
            Exit Select
          Case Tjstate.JWAITFORCOLON
            If sti = ":"c Then
              state = Tjstate.JWAITFORDATA
            Else
              If sti <> " "c AndAlso sti <> ControlChars.Lf AndAlso sti <> ControlChars.Cr Then
                Throw New System.Exception(FormatError("invalid char: was expecting """, cur_pos))
              End If
            End If
            Exit Select
          Case Tjstate.JWAITFORDATA
            If sti = "{"c Then
              Dim jobj As New YJSONObject(_data, cur_pos, _data_boundary)
              Dim len As Integer = jobj.parse()
              cur_pos += len
              parsed.Add(current_name, jobj)
              _keys.Add(current_name)
              state = Tjstate.JWAITFORNEXTSTRUCTMEMBER
              'cur_pos is already incremented
              inc_pos = False
              Exit Select
            ElseIf sti = "["c Then
              Dim jobj As New YJSONArray(_data, cur_pos, _data_boundary)
              Dim len As Integer = jobj.parse()
              cur_pos += len
              parsed.Add(current_name, jobj)
              _keys.Add(current_name)
              state = Tjstate.JWAITFORNEXTSTRUCTMEMBER
              'cur_pos is already incremented
              inc_pos = False
              Exit Select
            ElseIf sti = """"c Then
              Dim jobj As New YJSONString(_data, cur_pos, _data_boundary)
              Dim len As Integer = jobj.parse()
              cur_pos += len
              parsed.Add(current_name, jobj)
              _keys.Add(current_name)
              state = Tjstate.JWAITFORNEXTSTRUCTMEMBER
              'cur_pos is already incremented
              inc_pos = False
              Exit Select
            ElseIf sti = "-"c OrElse (sti >= "0"c AndAlso sti <= "9"c) Then
              Dim jobj As New YJSONNumber(_data, cur_pos, _data_boundary)
              Dim len As Integer = jobj.parse()
              cur_pos += len
              parsed.Add(current_name, jobj)
              _keys.Add(current_name)
              state = Tjstate.JWAITFORNEXTSTRUCTMEMBER
              'cur_pos is already incremented
              inc_pos = False
              Exit Select
            ElseIf sti <> " "c AndAlso sti <> ControlChars.Lf AndAlso sti <> ControlChars.Cr Then
              Throw New System.Exception(FormatError("invalid char: was expecting  "",0..9,t or f", cur_pos))
            End If
            Exit Select
          Case Tjstate.JWAITFORNEXTSTRUCTMEMBER
            If sti = ","c Then
              state = Tjstate.JWAITFORNAME
              name_start = cur_pos + 1
            ElseIf sti = "}"c Then
              _data_len = cur_pos + 1 - _data_start
              Return _data_len
            Else
              If sti <> " "c AndAlso sti <> ControlChars.Lf AndAlso sti <> ControlChars.Cr Then
                Throw New System.Exception(FormatError("invalid char: was expecting ,", cur_pos))
              End If
            End If
            Exit Select
          Case Tjstate.JWAITFORNEXTARRAYITEM, Tjstate.JWAITFORSTRINGVALUE, Tjstate.JWAITFORINTVALUE, Tjstate.JWAITFORBOOLVALUE
            Throw New System.Exception(FormatError("invalid state for YJSONObject", cur_pos))
        End Select
        If inc_pos Then cur_pos += 1
      End While
      Throw New System.Exception(FormatError("unexpected end of data", cur_pos))
    End Function

    Public Function has(key As String) As Boolean
      Return parsed.ContainsKey(key)
    End Function

    Public Function getYJSONObject(key As String) As YJSONObject
      Return DirectCast(parsed(key), YJSONObject)
    End Function

    Friend Function getYJSONString(key As String) As YJSONString
      Return DirectCast(parsed(key), YJSONString)
    End Function

    Friend Function getYJSONArray(key As String) As YJSONArray
      Return DirectCast(parsed(key), YJSONArray)
    End Function

    Public Function keys() As List(Of String)
      Return parsed.Keys.ToList()
    End Function

    Friend Function getYJSONNumber(key As String) As YJSONNumber
      Return DirectCast(parsed(key), YJSONNumber)
    End Function

    Public Sub remove(key As String)
      parsed.Remove(key)
    End Sub

    Public Function getString(key As String) As String
      If parsed(key).getJSONType() = YJSONType.NUMBER Then
        Dim ynumber As YJSONNumber = DirectCast(parsed(key), YJSONNumber)
        Return ynumber.getInt().ToString()
      ElseIf parsed(key).getJSONType() = YJSONType.STRING Then
        Dim ystr As YJSONString = DirectCast(parsed(key), YJSONString)
        Return ystr.getString()
      Else
        Return "<JSON_getString_error>"
      End If
    End Function

    Public Function getInt(key As String) As Integer
      Dim yint As YJSONNumber = DirectCast(parsed(key), YJSONNumber)
      Return yint.getInt()
    End Function

    Public Function [get](key As String) As YJSONContent
      Return parsed(key)
    End Function

    Public Function getLong(key As String) As Long
      Dim yint As YJSONNumber = DirectCast(parsed(key), YJSONNumber)
      Return yint.getLong()
    End Function

    Public Function getDouble(key As String) As Double
      Dim yint As YJSONNumber = DirectCast(parsed(key), YJSONNumber)
      Return yint.getDouble()
    End Function

    Public Overrides Function toJSON() As Byte()
      Dim res As New yMemoryStream()
      res.Append("{"c)
      Dim sep As String = ""
      For Each key As String In parsed.Keys.ToArray()
        Dim subContent As YJSONContent = parsed(key)
        Dim subres As Byte() = subContent.toJSON()
        res.Append(sep)
        res.Append(""""c)
        res.Append(key)
        res.Append(""":")
        res.Append(subres)
        sep = ","
      Next
      res.Append("}"c)
      Return res.ToArray()
    End Function

    Public Overrides Function ToString() As String
      Dim res As New StringBuilder()
      res.Append("{"c)
      Dim sep As String = ""
      For Each key As String In parsed.Keys.ToArray()
        Dim subContent As YJSONContent = parsed(key)
        Dim subres As String = subContent.ToString()
        res.Append(sep)
        res.Append(key)
        res.Append("=>")
        res.Append(subres)
        sep = ","
      Next
      res.Append("}"c)
      Return res.ToString()
    End Function


    Public Sub parseWithRef(reference As YJSONObject)
      If reference IsNot Nothing Then
        Try
          Dim yzon As New YJSONArray(_data, _data_start, _data_boundary)
          yzon.parse()
          convert(reference, yzon)
          Return

        Catch generatedExceptionName As Exception
        End Try
      End If
      Me.parse()
    End Sub

    Private Sub convert(reference As YJSONObject, newArray As YJSONArray)
      Dim length As Integer = newArray.Length
      For i As Integer = 0 To length - 1
        Dim key As String = reference.getKeyFromIdx(i)
        Dim new_item As YJSONContent = newArray.[get](i)
        Dim reference_item As YJSONContent = reference.[get](key)

        If new_item.getJSONType() = reference_item.getJSONType() Then
          parsed.Add(key, new_item)
          _keys.Add(key)
        ElseIf new_item.getJSONType() = YJSONType.ARRAY AndAlso reference_item.getJSONType() = YJSONType.[OBJECT] Then
          Dim jobj As New YJSONObject(new_item._data, new_item._data_start, reference_item._data_boundary)
          jobj.convert(DirectCast(reference_item, YJSONObject), DirectCast(new_item, YJSONArray))
          parsed.Add(key, jobj)
          _keys.Add(key)
        Else

          Throw New System.Exception("Unable to convert " + new_item.getJSONType().ToString() + " to " + reference.getJSONType().ToString())
        End If
      Next
    End Sub

    Private Function getKeyFromIdx(i As Integer) As String
      Return _keys(i)
    End Function
  End Class


  '=======================================================
  'Service provided by Telerik (www.telerik.com)
  'Conversion powered by NRefactory.
  'Twitter: @telerik
  'Facebook: facebook.com/telerik
  '=======================================================


  Public Const YOCTO_API_VERSION_STR As String = "2.0"
  Public Const YOCTO_API_VERSION_BCD As Integer = &H200
  Public Const YOCTO_API_BUILD_NO As String = "63744"

  Public Const YOCTO_DEFAULT_PORT As Integer = 4444
  Public Const YOCTO_VENDORID As Integer = &H24E0
  Public Const YOCTO_DEVID_FACTORYBOOT As Integer = 1
  Public Const YOCTO_DEVID_BOOTLOADER As Integer = 2

  Public Const YOCTO_ERRMSG_LEN As Integer = 256
  Public Const YOCTO_MANUFACTURER_LEN As Integer = 20
  Public Const YOCTO_SERIAL_LEN As Integer = 20
  Public Const YOCTO_BASE_SERIAL_LEN As Integer = 8
  Public Const YOCTO_PRODUCTNAME_LEN As Integer = 28
  Public Const YOCTO_FIRMWARE_LEN As Integer = 22
  Public Const YOCTO_LOGICAL_LEN As Integer = 20
  Public Const YOCTO_FUNCTION_LEN As Integer = 20
  Public Const YOCTO_PUBVAL_SIZE As Integer = 6 ' Size of the data (can be non null terminated)
  Public Const YOCTO_PUBVAL_LEN As Integer = 16 ' Temporary storage, > YOCTO_PUBVAL_SIZE
  Public Const YOCTO_PASS_LEN As Integer = 20
  Public Const YOCTO_REALM_LEN As Integer = 20
  Public Const INVALID_YHANDLE As Integer = 0
  Public Const yUnknowSize As Integer = 1024

  Public Const YOCTO_CALIB_TYPE_OFS As Integer = 30

  REM Global definitions for YRelay,YDatalogger an YWatchdog class
  Public Const Y_STATE_A As Integer = 0
  Public Const Y_STATE_B As Integer = 1
  Public Const Y_STATE_INVALID As Integer = -1
  Public Const Y_OUTPUT_OFF As Integer = 0
  Public Const Y_OUTPUT_ON As Integer = 1
  Public Const Y_OUTPUT_INVALID As Integer = -1
  Public Const Y_AUTOSTART_OFF As Integer = 0
  Public Const Y_AUTOSTART_ON As Integer = 1
  Public Const Y_AUTOSTART_INVALID As Integer = -1
  Public Const Y_RUNNING_OFF As Integer = 0
  Public Const Y_RUNNING_ON As Integer = 1
  Public Const Y_RUNNING_INVALID As Integer = -1





  REM --- (generated code: YHub return codes)
    REM --- (end of generated code: YHub return codes)
  REM --- (generated code: YHub dlldef)
    REM --- (end of generated code: YHub dlldef)
  REM --- (generated code: YHub yapiwrapper)
   REM --- (end of generated code: YHub yapiwrapper)
  REM --- (generated code: YHub globals)

  REM --- (end of generated code: YHub globals)

  REM --- (generated code: YHub class start)

  Public Class YHub
    REM --- (end of generated code: YHub class start)

    REM --- (generated code: YHub definitions)
    REM --- (end of generated code: YHub definitions)

    REM --- (generated code: YHub attributes declaration)
    Protected _ctx As YAPIContext
    Protected _hubref As Integer
    Protected _userData As Object
    REM --- (end of generated code: YHub attributes declaration)

    Public Sub New(ByVal yctx As YAPIContext, hubref As Integer)
      MyBase.New()
      REM --- (generated code: YHub attributes initialization)
      _hubref = 0
      _userData = Nothing
      REM --- (end of generated code: YHub attributes initialization)
      _ctx = yctx
      _hubref = hubref
    End Sub

    REM --- (generated code: YHub private methods declaration)

    REM --- (end of generated code: YHub private methods declaration)

    REM --- (generated code: YHub public methods declaration)
    Public Overridable Function _getStrAttr(attrName As String) As String
      Dim val As StringBuilder = New StringBuilder(1024)
      Dim res As Integer = 0
      Dim fullsize As Integer
      fullsize = 0
      res = _yapiGetHubStrAttr(Me._hubref, New StringBuilder(attrName), val, 1024, fullsize)
      If (res > 0) Then
        Return val.ToString()
      End If
      Return ""
    End Function

    Public Overridable Function _getIntAttr(attrName As String) As Integer
      Return _yapiGetHubIntAttr(Me._hubref, New StringBuilder(attrName))
    End Function

    Public Overridable Sub _setIntAttr(attrName As String, value As Integer)
      _yapiSetHubIntAttr(Me._hubref, New StringBuilder(attrName), value)
    End Sub

    '''*
    ''' <summary>
    '''   Returns the URL that has been used first to register this hub.
    ''' <para>
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function get_registeredUrl() As String
      Return Me._getStrAttr("registeredUrl")
    End Function

    '''*
    ''' <summary>
    '''   Returns all known URLs that have been used to register this hub.
    ''' <para>
    '''   URLs are pointing to the same hub when the devices connected
    '''   are sharing the same serial number.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function get_knownUrls() As List(Of String)
      Dim smallbuff As StringBuilder = New StringBuilder(1024)
      Dim bigbuff As StringBuilder
      Dim buffsize As Integer = 0
      Dim fullsize As Integer
      Dim yapi_res As Integer = 0
      Dim urls_packed As String
      Dim known_url_val As String
      Dim url_list As List(Of String) = New List(Of String)()

      fullsize = 0
      known_url_val = "knownUrls"
      yapi_res = _yapiGetHubStrAttr(Me._hubref, New StringBuilder(known_url_val), smallbuff, 1024, fullsize)
      If (yapi_res < 0) Then
        Return url_list
      End If
      If (fullsize <= 1024) Then
        urls_packed = smallbuff.ToString()
      Else
        buffsize = fullsize
        bigbuff = New StringBuilder(buffsize)
        yapi_res = _yapiGetHubStrAttr(Me._hubref, New StringBuilder(known_url_val), bigbuff, buffsize, fullsize)
        If (yapi_res < 0) Then
          bigbuff = Nothing
          Return url_list
        Else
          urls_packed = bigbuff.ToString()
        End If
        bigbuff = Nothing
      End If
      If (Not (urls_packed = "")) Then
        url_list = New List(Of String)(urls_packed.Split(New Char() {"?"c}))
      End If
      Return url_list
    End Function

    '''*
    ''' <summary>
    '''   Returns the URL currently in use to communicate with this hub.
    ''' <para>
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function get_connectionUrl() As String
      Return Me._getStrAttr("connectionUrl")
    End Function

    '''*
    ''' <summary>
    '''   Returns the hub serial number, if the hub was already connected once.
    ''' <para>
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function get_serialNumber() As String
      Return Me._getStrAttr("serialNumber")
    End Function

    '''*
    ''' <summary>
    '''   Tells if this hub is still registered within the API.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>true</c> if the hub has not been unregistered.
    ''' </returns>
    '''/
    Public Overridable Function isInUse() As Boolean
      Return Me._getIntAttr("isInUse") > 0
    End Function

    '''*
    ''' <summary>
    '''   Tells if there is an active communication channel with this hub.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>true</c> if the hub is currently connected.
    ''' </returns>
    '''/
    Public Overridable Function isOnline() As Boolean
      Return Me._getIntAttr("isOnline") > 0
    End Function

    '''*
    ''' <summary>
    '''   Tells if write access on this hub is blocked.
    ''' <para>
    '''   Return <c>true</c> if it
    '''   is not possible to change attributes on this hub
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>true</c> if it is not possible to change attributes on this hub.
    ''' </returns>
    '''/
    Public Overridable Function isReadOnly() As Boolean
      Return Me._getIntAttr("isReadOnly") > 0
    End Function

    '''*
    ''' <summary>
    '''   Modifies tthe network connection delay for this hub.
    ''' <para>
    '''   The default value is inherited from <c>ySetNetworkTimeout</c>
    '''   at the time when the hub is registered, but it can be updated
    '''   afterward for each specific hub if necessary.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="networkMsTimeout">
    '''   the network connection delay in milliseconds.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Sub set_networkTimeout(networkMsTimeout As Integer)
      Me._setIntAttr("networkTimeout",networkMsTimeout)
    End Sub

    '''*
    ''' <summary>
    '''   Returns the network connection delay for this hub.
    ''' <para>
    '''   The default value is inherited from <c>ySetNetworkTimeout</c>
    '''   at the time when the hub is registered, but it can be updated
    '''   afterward for each specific hub if necessary.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the network connection delay in milliseconds.
    ''' </returns>
    '''/
    Public Overridable Function get_networkTimeout() As Integer
      Return Me._getIntAttr("networkTimeout")
    End Function

    '''*
    ''' <summary>
    '''   Returns the numerical error code of the latest error with the hub.
    ''' <para>
    '''   This method is mostly useful when using the Yoctopuce library with
    '''   exceptions disabled.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a number corresponding to the code of the latest error that occurred while
    '''   using the hub object
    ''' </returns>
    '''/
    Public Overridable Function get_errorType() As Integer
      Return Me._getIntAttr("errorType")
    End Function

    '''*
    ''' <summary>
    '''   Returns the error message of the latest error with the hub.
    ''' <para>
    '''   This method is mostly useful when using the Yoctopuce library with
    '''   exceptions disabled.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest error message that occured while
    '''   using the hub object
    ''' </returns>
    '''/
    Public Overridable Function get_errorMessage() As String
      Return Me._getStrAttr("errorMessage")
    End Function

    '''*
    ''' <summary>
    '''   Returns the value of the userData attribute, as previously stored
    '''   using method <c>set_userData</c>.
    ''' <para>
    '''   This attribute is never touched directly by the API, and is at
    '''   disposal of the caller to store a context.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the object stored previously by the caller.
    ''' </returns>
    '''/
    Public Overridable Function get_userData() As Object
      Return Me._userData
    End Function

    '''*
    ''' <summary>
    '''   Stores a user context provided as argument in the userData
    '''   attribute of the function.
    ''' <para>
    '''   This attribute is never touched by the API, and is at
    '''   disposal of the caller to store a context.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="data">
    '''   any kind of object to be stored
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Sub set_userData(data As Object)
      Me._userData = data
    End Sub

    '''*
    ''' <summary>
    '''   Starts the enumeration of hubs currently in use by the API.
    ''' <para>
    '''   Use the method <c>YHub.nextHubInUse()</c> to iterate on the
    '''   next hubs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHub</c> object, corresponding to
    '''   the first hub currently in use by the API, or a
    '''   <c>Nothing</c> pointer if none has been registered.
    ''' </returns>
    '''/
    Public Shared Function FirstHubInUse() As YHub
      Return YAPI.nextHubInUseInternal(-1)
    End Function

    '''*
    ''' <summary>
    '''   Continues the module enumeration started using <c>YHub.FirstHubInUse()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the order of returned hubs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YHub</c> object, corresponding to
    '''   the next hub currenlty in use, or a <c>Nothing</c> pointer
    '''   if there are no more hubs to enumerate.
    ''' </returns>
    '''/
    Public Overridable Function nextHubInUse() As YHub
      Return Me._ctx.nextHubInUseInternal(Me._hubref)
    End Function



    REM --- (end of generated code: YHub public methods declaration)

  End Class

  REM --- (generated code: YHub functions)


  REM --- (end of generated code: YHub functions)


  REM --- (generated code: YAPIContext return codes)
    REM --- (end of generated code: YAPIContext return codes)
  REM --- (generated code: YAPIContext dlldef)
    REM --- (end of generated code: YAPIContext dlldef)
  REM --- (generated code: YAPIContext globals)

  REM --- (end of generated code: YAPIContext globals)

  REM --- (generated code: YAPIContext class start)

  Public Class YAPIContext
    REM --- (end of generated code: YAPIContext class start)

    REM --- (generated code: YAPIContext definitions)
    REM --- (end of generated code: YAPIContext definitions)

    REM --- (generated code: YAPIContext attributes declaration)
    Protected _defaultCacheValidity As Long
    REM --- (end of generated code: YAPIContext attributes declaration)

    Public Sub New()
      REM --- (generated code: YAPIContext attributes initialization)
      _defaultCacheValidity = 5
      REM --- (end of generated code: YAPIContext attributes initialization)
    End Sub

    REM --- (generated code: YAPIContext private methods declaration)

    REM --- (end of generated code: YAPIContext private methods declaration)

    REM --- (generated code: YAPIContext public methods declaration)
    '''*
    ''' <summary>
    '''   Modifies the delay between each forced enumeration of the used YoctoHubs.
    ''' <para>
    '''   By default, the library performs a full enumeration every 10 seconds.
    '''   To reduce network traffic, you can increase this delay.
    '''   It's particularly useful when a YoctoHub is connected to the GSM network
    '''   where traffic is billed. This parameter doesn't impact modules connected by USB,
    '''   nor the working of module arrival/removal callbacks.
    '''   Note: you must call this function after <c>yInitAPI</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="deviceListValidity">
    '''   nubmer of seconds between each enumeration.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Sub SetDeviceListValidity(deviceListValidity As Integer)
      _yapiSetNetDevListValidity(deviceListValidity)
    End Sub

    '''*
    ''' <summary>
    '''   Returns the delay between each forced enumeration of the used YoctoHubs.
    ''' <para>
    '''   Note: you must call this function after <c>yInitAPI</c>.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the number of seconds between each enumeration.
    ''' </returns>
    '''/
    Public Overridable Function GetDeviceListValidity() As Integer
      Dim res As Integer = 0
      res = _yapiGetNetDevListValidity()
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Adds a UDEV rule which authorizes all users to access Yoctopuce modules
    '''   connected to the USB ports.
    ''' <para>
    '''   This function works only under Linux. The process that
    '''   calls this method must have root privileges because this method changes the Linux configuration.
    ''' </para>
    ''' </summary>
    ''' <param name="force">
    '''   if true, overwrites any existing rule.
    ''' </param>
    ''' <returns>
    '''   an empty string if the rule has been added.
    ''' </returns>
    ''' <para>
    '''   On failure, returns a string that starts with "error:".
    ''' </para>
    '''/
    Public Overridable Function AddUdevRule(force As Boolean) As String
      Dim msg As String
      Dim res As Integer = 0
      Dim c_force As Integer = 0
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      If (force) Then
        c_force = 1
      Else
        c_force = 0
      End If
      res = _yapiAddUdevRulesForYocto(c_force, errmsg)
      If (res < 0) Then
        msg = "error: " + errmsg.ToString()
      Else
        msg = ""
      End If
      Return msg
    End Function

    '''*
    ''' <summary>
    '''   Download the TLS/SSL certificate from the hub.
    ''' <para>
    '''   This function allows to download a TLS/SSL certificate to add it
    '''   to the list of trusted certificates using the AddTrustedCertificates method.
    ''' </para>
    ''' </summary>
    ''' <param name="url">
    '''   the root URL of the VirtualHub V2 or HTTP server.
    ''' </param>
    ''' <param name="mstimeout">
    '''   the number of milliseconds available to download the certificate.
    ''' </param>
    ''' <returns>
    '''   a string containing the certificate. In case of error, returns a string starting with "error:".
    ''' </returns>
    '''/
    Public Overridable Function DownloadHostCertificate(url As String, mstimeout As Long) As String
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim smallbuff As StringBuilder = New StringBuilder(4096)
      Dim bigbuff As StringBuilder
      Dim buffsize As Integer = 0
      Dim fullsize As Integer
      Dim res As Integer = 0
      Dim certifcate As String
      fullsize = 0
      res = _yapiGetRemoteCertificate(New StringBuilder(url), CType(mstimeout, yu64), smallbuff, 4096, fullsize, errmsg)
      If (res < 0) Then
        If (res = YAPI.BUFFER_TOO_SMALL) Then
          fullsize = fullsize * 2
          buffsize = fullsize
          bigbuff = New StringBuilder(buffsize)
          res = _yapiGetRemoteCertificate(New StringBuilder(url), CType(mstimeout, yu64), bigbuff, buffsize, fullsize, errmsg)
          If (res < 0) Then
            certifcate = "error:" + errmsg.ToString()
          Else
            certifcate = bigbuff.ToString()
          End If
          bigbuff = Nothing
        Else
          certifcate = "error:" + errmsg.ToString()
        End If
        Return certifcate
      Else
        certifcate = smallbuff.ToString()
      End If
      Return certifcate
    End Function

    '''*
    ''' <summary>
    '''   Adds a TLS/SSL certificate to the list of trusted certificates.
    ''' <para>
    '''   By default, the library
    '''   library will reject TLS/SSL connections to servers whose certificate is not known. This function
    '''   function allows to add a list of known certificates. It is also possible to disable the verification
    '''   using the SetNetworkSecurityOptions method.
    ''' </para>
    ''' </summary>
    ''' <param name="certificate">
    '''   a string containing one or more certificates.
    ''' </param>
    ''' <returns>
    '''   an empty string if the certificate has been added correctly.
    '''   In case of error, returns a string starting with "error:".
    ''' </returns>
    '''/
    Public Overridable Function AddTrustedCertificates(certificate As String) As String
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim size As Integer = 0
      Dim res As Integer = 0
      REM // null char must be inclued
      size = (certificate).Length + 1
      res = _yapiAddSSLCertificateCli(New StringBuilder(certificate), size, errmsg)
      If (res < 0) Then
        Return errmsg.ToString()
      Else
        Return ""
      End If
    End Function

    '''*
    ''' <summary>
    '''   Set the path of Certificate Authority file on local filesystem.
    ''' <para>
    '''   This method takes as a parameter the path of a file containing all certificates in PEM format.
    '''   For technical reasons, only one file can be specified. So if you need to connect to several Hubs
    '''   instances with self-signed certificates, you'll need to use
    '''   a single file containing all the certificates end-to-end. Passing a empty string will restore the
    '''   default settings. This option is only supported by PHP library.
    ''' </para>
    ''' </summary>
    ''' <param name="certificatePath">
    '''   the path of the file containing all certificates in PEM format.
    ''' </param>
    ''' <returns>
    '''   an empty string if the certificate has been added correctly.
    '''   In case of error, returns a string starting with "error:".
    ''' </returns>
    '''/
    Public Overridable Function SetTrustedCertificatesList(certificatePath As String) As String
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As Integer = 0
      res = _yapiSetTrustedCertificatesList(New StringBuilder(certificatePath), errmsg)
      If (res < 0) Then
        Return errmsg.ToString()
      Else
        Return ""
      End If
    End Function

    '''*
    ''' <summary>
    '''   Enables or disables certain TLS/SSL certificate checks.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="opts">
    '''   The options are <c>YAPI.NO_TRUSTED_CA_CHECK</c>,
    '''   <c>YAPI.NO_EXPIRATION_CHECK</c>, <c>YAPI.NO_HOSTNAME_CHECK</c>.
    ''' </param>
    ''' <returns>
    '''   an empty string if the options are taken into account.
    '''   On error, returns a string beginning with "error:".
    ''' </returns>
    '''/
    Public Overridable Function SetNetworkSecurityOptions(opts As Integer) As String
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As Integer = 0
      res = _yapiSetNetworkSecurityOptions(opts, errmsg)
      If (res < 0) Then
        Return errmsg.ToString()
      Else
        Return ""
      End If
    End Function

    '''*
    ''' <summary>
    '''   Modifies the network connection delay for <c>yRegisterHub()</c> and <c>yUpdateDeviceList()</c>.
    ''' <para>
    '''   This delay impacts only the YoctoHubs and VirtualHub
    '''   which are accessible through the network. By default, this delay is of 20000 milliseconds,
    '''   but depending on your network you may want to change this delay,
    '''   gor example if your network infrastructure is based on a GSM connection.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="networkMsTimeout">
    '''   the network connection delay in milliseconds.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Sub SetNetworkTimeout(networkMsTimeout As Integer)
      _yapiSetNetworkTimeout(networkMsTimeout)
    End Sub

    '''*
    ''' <summary>
    '''   Returns the network connection delay for <c>yRegisterHub()</c> and <c>yUpdateDeviceList()</c>.
    ''' <para>
    '''   This delay impacts only the YoctoHubs and VirtualHub
    '''   which are accessible through the network. By default, this delay is of 20000 milliseconds,
    '''   but depending on your network you may want to change this delay,
    '''   for example if your network infrastructure is based on a GSM connection.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the network connection delay in milliseconds.
    ''' </returns>
    '''/
    Public Overridable Function GetNetworkTimeout() As Integer
      Dim res As Integer = 0
      res = _yapiGetNetworkTimeout()
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Change the validity period of the data loaded by the library.
    ''' <para>
    '''   By default, when accessing a module, all the attributes of the
    '''   module functions are automatically kept in cache for the standard
    '''   duration (5 ms). This method can be used to change this standard duration,
    '''   for example in order to reduce network or USB traffic. This parameter
    '''   does not affect value change callbacks
    '''   Note: This function must be called after <c>yInitAPI</c>.
    ''' </para>
    ''' </summary>
    ''' <param name="cacheValidityMs">
    '''   an integer corresponding to the validity attributed to the
    '''   loaded function parameters, in milliseconds.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Sub SetCacheValidity(cacheValidityMs As Long)
      Me._defaultCacheValidity = cacheValidityMs
    End Sub

    '''*
    ''' <summary>
    '''   Returns the validity period of the data loaded by the library.
    ''' <para>
    '''   This method returns the cache validity of all attributes
    '''   module functions.
    '''   Note: This function must be called after <c>yInitAPI </c>.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the validity attributed to the
    '''   loaded function parameters, in milliseconds
    ''' </returns>
    '''/
    Public Overridable Function GetCacheValidity() As Long
      Return Me._defaultCacheValidity
    End Function

    Public Overridable Function nextHubInUseInternal(hubref As Integer) As YHub
      Dim nextref As Integer = 0
      nextref = _yapiGetNextHubRef(hubref)
      If (nextref < 0) Then
        Return Nothing
      End If
      Return Me.getYHubObj(nextref)
    End Function

    Public Overridable Function getYHubObj(hubref As Integer) As YHub
      Dim obj As YHub
      obj = Me._findYHubFromCache(hubref)
      If ((obj Is Nothing)) Then
        obj = New YHub(Me, hubref)
        Me._addYHubToCache(hubref, obj)
      End If
      Return obj
    End Function



    REM --- (end of generated code: YAPIContext public methods declaration)

    Private Sub _addYHubToCache(hubref As Integer, obj As YHub)
      _yhub_cache(hubref) = obj
    End Sub

    Public Property _yhub_cache As Dictionary(Of Integer, YHub) = New Dictionary(Of Integer, YHub)
    Private Function _findYHubFromCache(hubref As Integer) As YHub
      If (_yhub_cache.ContainsKey(hubref)) Then
        Return _yhub_cache(hubref)
      End If
      Return Nothing
    End Function
    Private Sub _ClearCache()
      _yhub_cache.Clear()
    End Sub
  End Class

  REM --- (generated code: YAPIContext functions)


  REM --- (end of generated code: YAPIContext functions)


  Public Class YAPI
    Public Shared DefaultEncoding As Encoding = System.Text.Encoding.GetEncoding(1252)

    REM Switch to turn off exceptions and use return codes instead, for source-code compatibility
    REM with languages without exception support like C
    Public Shared ExceptionsDisabled As Boolean = False
    Private Shared apiInitialized As Boolean = False

    REM Default cache validity (in [ms]) before reloading data from device. This saves a lots of trafic.
    REM Note that a value undger 2 ms makes little sense since a USB bus itself has a 2ms roundtrip period
    Public Shared DefaultCacheValidity As Integer = 5

    Public Const INVALID_STRING As String = "!INVALID!"
    Public Const INVALID_DOUBLE As Double = -1.7976931348623157E+308
    Public Const MIN_DOUBLE As Double = Double.MinValue
    Public Const MAX_DOUBLE As Double = Double.MaxValue
    Public Const INVALID_INT As Integer = -2147483648
    Public Const INVALID_UINT As Integer = -1
    Public Const INVALID_LONG As Long = -9223372036854775807
    Public Const HARDWAREID_INVALID As String = INVALID_STRING
    Public Const FUNCTIONID_INVALID As String = INVALID_STRING
    Public Const FRIENDLYNAME_INVALID As String = INVALID_STRING

    REM yInitAPI argument
    Public Const DETECT_NONE As Integer = 0
    Public Const DETECT_USB As Integer = 1
    Public Const DETECT_NET As Integer = 2
    Public Const RESEND_MISSING_PKT As Integer = 4
    Public Const DETECT_ALL As Integer = DETECT_USB Or DETECT_NET
    Public Const HASH_BUF_SIZE As Integer = 28
    REM --- (generated code: YFunction return codes)
    REM Yoctopuce error codes, also used by default as function return value
    Public Const SUCCESS As Integer = 0         REM everything worked all right
    Public Const NOT_INITIALIZED As Integer = -1 REM call yInitAPI() first !
    Public Const INVALID_ARGUMENT As Integer = -2 REM one of the arguments passed to the function is invalid
    Public Const NOT_SUPPORTED As Integer = -3  REM the operation attempted is (currently) not supported
    Public Const DEVICE_NOT_FOUND As Integer = -4 REM the requested device is not reachable
    Public Const VERSION_MISMATCH As Integer = -5 REM the device firmware is incompatible with this API version
    Public Const DEVICE_BUSY As Integer = -6    REM the device is busy with another task and cannot answer
    Public Const TIMEOUT As Integer = -7        REM the device took too long to provide an answer
    Public Const IO_ERROR As Integer = -8       REM there was an I/O problem while talking to the device
    Public Const NO_MORE_DATA As Integer = -9   REM there is no more data to read from
    Public Const EXHAUSTED As Integer = -10     REM you have run out of a limited resource, check the documentation
    Public Const DOUBLE_ACCES As Integer = -11  REM you have two process that try to access to the same device
    Public Const UNAUTHORIZED As Integer = -12  REM unauthorized access to password-protected device
    Public Const RTC_NOT_READY As Integer = -13 REM real-time clock has not been initialized (or time was lost)
    Public Const FILE_NOT_FOUND As Integer = -14 REM the file is not found
    Public Const SSL_ERROR As Integer = -15     REM Error reported by mbedSSL
    Public Const RFID_SOFT_ERROR As Integer = -16 REM Recoverable error with RFID tag (eg. tag out of reach), check YRfidStatus for details
    Public Const RFID_HARD_ERROR As Integer = -17 REM Serious RFID error (eg. write-protected, out-of-boundary), check YRfidStatus for details
    Public Const BUFFER_TOO_SMALL As Integer = -18 REM The buffer provided is too small
    Public Const DNS_ERROR As Integer = -19     REM Error during name resolutions (invalid hostname or dns communication error)
    Public Const SSL_UNK_CERT As Integer = -20  REM The certificate is not correctly signed by the trusted CA

  REM TLS / SSL definitions
    Public Const NO_TRUSTED_CA_CHECK As Integer = 1 REM Disables certificate checking
    Public Const NO_EXPIRATION_CHECK As Integer = 2 REM Disables certificate expiration date checking
    Public Const NO_HOSTNAME_CHECK As Integer = 4 REM Disable hostname checking
    Public Const LEGACY As Integer = 8          REM Allow non secure connection (similar to v1.10)

    REM --- (end of generated code: YFunction return codes)
    Public Shared _yapiContext As YAPIContext = New YAPIContext()


    Friend Shared Function ParseHTTP(data As String, start As Integer, [stop] As Integer, ByRef headerlen As Integer, ByRef errmsg As String) As Integer
      Const httpheader As String = "HTTP/1.1 "
      Const okHeader As String = "OK" & vbCr & vbLf
      Dim p1 As Integer = 0
      Dim p2 As Integer = 0
      Const CR As String = vbCr & vbLf
      Dim httpcode As Integer

      If ([stop] - start) > okHeader.Length AndAlso data.Substring(start, okHeader.Length) = okHeader Then
        httpcode = 200
        errmsg = ""
      Else
        If ([stop] - start) < httpheader.Length OrElse data.Substring(start, httpheader.Length) <> httpheader Then
          errmsg = Convert.ToString("data should start with ") & httpheader
          headerlen = 0
          Return -1
        End If

        p1 = data.IndexOf(" ", start + httpheader.Length - 1)
        p2 = data.IndexOf(" ", p1 + 1)
        If p1 < 0 OrElse p2 < 0 Then
          errmsg = "Invalid HTTP header (invalid first line)"
          headerlen = 0
          Return -1
        End If

        httpcode = Convert.ToInt32(data.Substring(p1, p2 - p1 + 1))
        If httpcode <> 200 Then
          errmsg = String.Format("Unexpected HTTP return code:{0}", httpcode)
        Else
          errmsg = ""
        End If
      End If
      p1 = data.IndexOf(CR & CR, start)
      'json data is a structure
      If p1 < 0 Then
        errmsg = "Invalid HTTP header (missing header end)"
        headerlen = 0
        Return -1
      End If
      headerlen = p1 + 4
      Return httpcode
    End Function

    Friend Shared Function _atof(ByVal val As String) As Double
      Dim res As Double
      Double.TryParse(val, NumberStyles.Number, CultureInfo.InvariantCulture, res)
      Return res
    End Function

    Friend Shared Function _escapeAttr(ByVal changeval As String) As String
      Dim i, c_ord As Integer
      Dim uchangeval, h As String
      uchangeval = ""
      Dim c As Char
      For i = 0 To changeval.Length - 1
        c = changeval.Chars(i)
        If (c <= " ") Or ((c > Chr(122)) And (c <> "~")) Or (c = Chr(34)) Or (c = "%") Or (c = "&") Or
           (c = "+") Or (c = "<") Or (c = "=") Or (c = ">") Or (c = "\") Or (c = "^") Or (c = "`") Then
          c_ord = Asc(c)
          If (c_ord = &HC2 Or c_ord = &HC3) And (i + 1 < changeval.Length) Then
            If ((Asc(changeval(i + 1)) And &HEC0) = &H80) Then
              REM UTF8-encoded ISO-8859-1 character: translate to plain ISO-8859-1
              c_ord = (c_ord And 1) * &H40
              i += 1
              c_ord += Asc(changeval(i))
            End If
          End If
          h = Hex(c_ord)
          If (h.Length < 2) Then h = "0" + h
          uchangeval = uchangeval + "%" + h
        Else
          uchangeval = uchangeval + c
        End If
      Next
      _escapeAttr = uchangeval
    End Function


    REM calibration handlers
    Private Shared _CalibHandlers As New Dictionary(Of String, yCalibrationHandler)

    Public Shared Sub handlersCleanUp()
      _CalibHandlers.Clear()
    End Sub


    Public Shared Function _getCalibrationHandler(ByVal calType As Integer) As yCalibrationHandler

      Dim key As String

      key = calType.ToString()
      If (_CalibHandlers.ContainsKey(key)) Then
        _getCalibrationHandler = _CalibHandlers(key)
        Exit Function
      End If

      _getCalibrationHandler = Nothing
    End Function


    Public Shared Function _checkFirmware(ByVal serial As String, ByVal rev As String, ByVal path As String) As String
      Return ""
    End Function


    Public Shared Function _flattenJsonStruct(current_settings As Byte()) As Byte()
      Return Nothing
    End Function


    Private Shared ReadOnly decexp() As Double = {
  0.000001, 0.00001, 0.0001, 0.001, 0.01, 0.1, 1.0,
  10.0, 100.0, 1000.0, 10000.0, 100000.0, 1000000.0, 10000000.0, 100000000.0, 1000000000.0}

    REM Convert Yoctopuce 16-bit decimal floats to standard double-precision floats
    REM
    Public Shared Function _decimalToDouble(ByVal val As Long) As Double
      Dim negate As Boolean = False
      Dim res As Double
      Dim mantis As Long = val And 2047
      If mantis = 0 Then
        Return 0.0
      End If
      If val > 32767 Then
        negate = True
        val = 65536 - val
      ElseIf val < 0 Then
        negate = True
        val = -val
      End If
      Dim exp As Long = val >> 11
      res = (mantis) * decexp(CInt(exp))
      If negate Then
        Return -res
      Else
        Return res
      End If
    End Function

    REM Convert standard double-precision floats to Yoctopuce 16-bit decimal floats
    REM
    Public Shared Function _doubleToDecimal(ByVal val As Double) As Integer
      Dim negate As Integer = 0
      Dim comp As Double
      Dim mant As Double
      Dim decpw As Integer
      Dim res As Integer

      If val = 0.0 Then
        Return 0
      End If
      If val < 0 Then
        negate = 1
        val = -val
      End If
      comp = val / 1999.0
      decpw = 0
      While comp > decexp(decpw) And decpw < 15
        decpw = decpw + 1
      End While
      mant = val / decexp(decpw)
      If (decpw = 15 And mant > 2047.0) Then
        res = (15 << 11) + 2047 REM overflow
      Else
        res = (decpw << 11) + Convert.ToInt32(mant)
      End If
      If negate <> 0 Then
        Return -res
      Else
        Return res
      End If
    End Function


    Public Shared Function _decodeWords(ByVal sdat As String) As List(Of Integer)
      Dim udat As New List(Of Integer)()
      For p As Integer = 0 To sdat.Length - 1 Step 0
        Dim val As UInteger
        Dim c As UInteger
        c = CUInt(Asc(sdat.Substring(p, 1)))
        p += 1
        If c = 42 Then REM 42 == '*'
          val = 0
        ElseIf c = 88 Then REM 88 == 'X'
          val = &HFFFF
        ElseIf c = 89 Then REM 89 == 'Y'
          val = &H7FFF
        ElseIf (c >= 97) Then REM 97 ='a'
          Dim srcpos As Integer = CInt((udat.Count - 1 - (c - 97)))
          If (srcpos < 0) Then
            val = 0
          Else
            val = CUInt(udat.ElementAt(srcpos))
          End If
        Else
          If (p + 2 > sdat.Length) Then
            Return udat
          End If
          val = CUInt(c - 48) REM 48='0'
          c = CUInt(Asc(sdat.Substring(p, 1)))
          p += 1
          val += CUInt(c - 48) << 5
          c = CUInt(Asc(sdat.Substring(p, 1)))
          p += 1
          If (c = 122) Then REM 122 ='z'
            c = 92 REM 92 ='\'
          End If
          val += CUInt(c - 48) << 10
        End If
        udat.Add(CInt(val))
      Next p
      Return udat
    End Function


    Public Shared Function _decodeFloats(ByVal sdat As String) As List(Of Integer)
      Dim idat As New List(Of Integer)()
      For p As Integer = 0 To sdat.Length - 1 Step 0
        Dim val As Integer = 0
        Dim sign As Integer = 1
        Dim dec As Integer = 0
        Dim decInc As Integer = 0
        Dim c As Integer
        c = Asc(sdat.Substring(p, 1))
        p += 1
        While c <> 45 And (c < 48 Or c > 57) REM 45='-', 48='0', 57='9'
          If p > sdat.Length Then
            Return idat
          End If
          c = Asc(sdat.Substring(p, 1))
          p += 1
        End While
        If c = 45 Then REM 45='-'
          If p > sdat.Length Then
            Return idat
          End If
          sign = -sign
          c = Asc(sdat.Substring(p, 1))
          p += 1
        End If
        While (c >= 48 And c <= 57) Or c = 46 REM 48='0', 57='9', 46='.'
          If c = 46 Then REM 46='.'
            decInc = 1
          Else
            val = val * 10 + c - 48 REM 48='0'
            dec += decInc
          End If
          If p < sdat.Length Then
            c = Asc(sdat.Substring(p, 1))
            p += 1
          Else
            c = 0
          End If
        End While
        If dec < 3 Then
          If dec = 0 Then
            val *= 1000
          ElseIf dec = 1 Then
            val *= 100
          Else
            val *= 10
          End If
        End If
        idat.Add(sign * val)
      Next p
      Return idat
    End Function

    Public Shared Function _floatToStr(ByVal value As Double) As String
      Dim res As String = ""
      Dim rounded As Long
      Dim decim As Long

      rounded = CLng(Math.Round(value * 1000))
      If (rounded < 0) Then
        res = "-"
        rounded = -rounded
      End If
      res += Convert.ToString(rounded \ 1000)
      decim = rounded Mod 1000
      If decim > 0 Then
        res += "."
        If (decim < 100) Then res += "0"
        If (decim < 10) Then res += "0"
        If (decim Mod 10) = 0 Then decim \= 10
        If (decim Mod 10) = 0 Then decim \= 10
        res += Convert.ToString(decim)
      End If
      Return res
    End Function

    Public Shared Function _atoi(ByVal str As String) As Integer
      Dim p As Integer = 0
      Dim start As Integer
      While (p < str.Length)
        If Char.IsWhiteSpace(str(p)) Then
          p = p + 1
        Else
          Exit While
        End If
      End While
      start = p
      If p < str.Length Then
        If (str(p) = "-" Or str(p) = "+") Then
          p = p + 1
        End If
      End If
      While p < str.Length
        If Char.IsDigit(str(p)) Then
          p = p + 1
        Else
          Exit While
        End If
      End While
      If (start < p) Then
        Return Integer.Parse(str.Substring(start, p - start))
      End If
      Return 0
    End Function

    Public Shared Function _bytesToHexStr(ByVal bytes As Byte(), ByVal ofs As Integer, ByVal len As Integer) As String
      Dim hexArray As String = "0123456789ABCDEF"
      Dim hexChars(len * 2 - 1) As Char
      Dim j As Integer = 0
      While (j < len)
        Dim v As Integer
        v = bytes(j + ofs) And &HFF
        hexChars(j * 2) = hexArray(v >> 4)
        hexChars(j * 2 + 1) = hexArray(v And &HF)
        j += 1
      End While
      Return New String(hexChars)
    End Function

    Public Shared Function _hexStrToBin(ByVal hex_str As String) As Byte()
      Dim len As Integer = hex_str.Length \ 2
      Dim res(len - 1) As Byte
      Dim i As Integer = 0
      While (i < len)
        Dim val As Integer = 0
        For n As Integer = 0 To 1
          Dim c As Integer = Asc(hex_str(i * 2 + n))
          val <<= 4
          REM  57='9', 70='F', 102='f'
          If (c <= 57) Then
            val += c - 48
          ElseIf (c <= 70) Then
            val += c - 65 + 10
          Else
            val += c - 97 + 10
          End If
        Next
        res(i) = CType(val And &HFF, Byte)
        i += 1
      End While
      Return res
    End Function


    Public Shared Function _bytesMerge(ByVal array_a As Byte(), ByVal array_b As Byte()) As Byte()
      Dim res As Byte()
      ReDim res(array_a.Length + array_b.Length)
      System.Buffer.BlockCopy(array_a, 0, res, 0, array_a.Length)
      System.Buffer.BlockCopy(array_b, 0, res, array_a.Length, array_b.Length)
      Return res
    End Function


    Public Shared Function _boolToStr(b As Boolean) As String
      If b Then Return "1" Else Return "0"
    End Function

    Public Shared Function _intToHex(h As Integer, width As Integer) As String
      Dim res As String
      res = h.ToString("X")
      While (Len(res) < width)
        res = "0" + res
      End While
      Return res
    End Function


    Public Shared Sub RegisterCalibrationHandler(ByVal calibType As Integer, ByVal callback As yCalibrationHandler)

      Dim key As String
      key = calibType.ToString()
      _CalibHandlers.Add(key, callback)
    End Sub

    Private Shared Function yLinearCalibrationHandler(ByVal rawValue As Double, ByVal calibType As Integer, ByVal parameters As List(Of Integer), ByVal rawValues As List(Of Double), ByVal refValues As List(Of Double)) As Double

      Dim npt, i As Integer
      Dim x, adj As Double
      Dim x2, adj2 As Double

      x = rawValues(0)
      adj = refValues(0) - x
      i = 0

      If (calibType < YOCTO_CALIB_TYPE_OFS) Then
        npt = calibType Mod 10
        If (npt > rawValues.Count) Then npt = rawValues.Count
        If (npt > refValues.Count) Then npt = refValues.Count
      Else
        npt = refValues.Count
      End If

      While ((rawValue > rawValues(i)) And (i + 1 < npt))
        i = i + 1
        x2 = x
        adj2 = adj
        x = rawValues(i)
        adj = refValues(i) - x
        If ((rawValue < x) And (x > x2)) Then
          adj = adj2 + (adj - adj2) * (rawValue - x2) / (x - x2)
        End If
      End While
      yLinearCalibrationHandler = rawValue + adj
    End Function


    REM --- (generated code: YAPIContext yapiwrapper)
    '''*
    ''' <summary>
    '''   Modifies the delay between each forced enumeration of the used YoctoHubs.
    ''' <para>
    '''   By default, the library performs a full enumeration every 10 seconds.
    '''   To reduce network traffic, you can increase this delay.
    '''   It's particularly useful when a YoctoHub is connected to the GSM network
    '''   where traffic is billed. This parameter doesn't impact modules connected by USB,
    '''   nor the working of module arrival/removal callbacks.
    '''   Note: you must call this function after <c>yInitAPI</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="deviceListValidity">
    '''   nubmer of seconds between each enumeration.
    ''' @noreturn
    ''' </param>
    '''/
    Public Shared Sub SetDeviceListValidity(deviceListValidity As Integer)
        YAPI.InitAPI(0, Nothing)
        _yapiContext.SetDeviceListValidity(deviceListValidity)
    End Sub
    '''*
    ''' <summary>
    '''   Returns the delay between each forced enumeration of the used YoctoHubs.
    ''' <para>
    '''   Note: you must call this function after <c>yInitAPI</c>.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the number of seconds between each enumeration.
    ''' </returns>
    '''/
    Public Shared Function GetDeviceListValidity() As Integer
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.GetDeviceListValidity()
    End Function
    '''*
    ''' <summary>
    '''   Adds a UDEV rule which authorizes all users to access Yoctopuce modules
    '''   connected to the USB ports.
    ''' <para>
    '''   This function works only under Linux. The process that
    '''   calls this method must have root privileges because this method changes the Linux configuration.
    ''' </para>
    ''' </summary>
    ''' <param name="force">
    '''   if true, overwrites any existing rule.
    ''' </param>
    ''' <returns>
    '''   an empty string if the rule has been added.
    ''' </returns>
    ''' <para>
    '''   On failure, returns a string that starts with "error:".
    ''' </para>
    '''/
    Public Shared Function AddUdevRule(force As Boolean) As String
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.AddUdevRule(force)
    End Function
    '''*
    ''' <summary>
    '''   Download the TLS/SSL certificate from the hub.
    ''' <para>
    '''   This function allows to download a TLS/SSL certificate to add it
    '''   to the list of trusted certificates using the AddTrustedCertificates method.
    ''' </para>
    ''' </summary>
    ''' <param name="url">
    '''   the root URL of the VirtualHub V2 or HTTP server.
    ''' </param>
    ''' <param name="mstimeout">
    '''   the number of milliseconds available to download the certificate.
    ''' </param>
    ''' <returns>
    '''   a string containing the certificate. In case of error, returns a string starting with "error:".
    ''' </returns>
    '''/
    Public Shared Function DownloadHostCertificate(url As String, mstimeout As Long) As String
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.DownloadHostCertificate(url, mstimeout)
    End Function
    '''*
    ''' <summary>
    '''   Adds a TLS/SSL certificate to the list of trusted certificates.
    ''' <para>
    '''   By default, the library
    '''   library will reject TLS/SSL connections to servers whose certificate is not known. This function
    '''   function allows to add a list of known certificates. It is also possible to disable the verification
    '''   using the SetNetworkSecurityOptions method.
    ''' </para>
    ''' </summary>
    ''' <param name="certificate">
    '''   a string containing one or more certificates.
    ''' </param>
    ''' <returns>
    '''   an empty string if the certificate has been added correctly.
    '''   In case of error, returns a string starting with "error:".
    ''' </returns>
    '''/
    Public Shared Function AddTrustedCertificates(certificate As String) As String
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.AddTrustedCertificates(certificate)
    End Function
    '''*
    ''' <summary>
    '''   Set the path of Certificate Authority file on local filesystem.
    ''' <para>
    '''   This method takes as a parameter the path of a file containing all certificates in PEM format.
    '''   For technical reasons, only one file can be specified. So if you need to connect to several Hubs
    '''   instances with self-signed certificates, you'll need to use
    '''   a single file containing all the certificates end-to-end. Passing a empty string will restore the
    '''   default settings. This option is only supported by PHP library.
    ''' </para>
    ''' </summary>
    ''' <param name="certificatePath">
    '''   the path of the file containing all certificates in PEM format.
    ''' </param>
    ''' <returns>
    '''   an empty string if the certificate has been added correctly.
    '''   In case of error, returns a string starting with "error:".
    ''' </returns>
    '''/
    Public Shared Function SetTrustedCertificatesList(certificatePath As String) As String
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.SetTrustedCertificatesList(certificatePath)
    End Function
    '''*
    ''' <summary>
    '''   Enables or disables certain TLS/SSL certificate checks.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="opts">
    '''   The options are <c>YAPI.NO_TRUSTED_CA_CHECK</c>,
    '''   <c>YAPI.NO_EXPIRATION_CHECK</c>, <c>YAPI.NO_HOSTNAME_CHECK</c>.
    ''' </param>
    ''' <returns>
    '''   an empty string if the options are taken into account.
    '''   On error, returns a string beginning with "error:".
    ''' </returns>
    '''/
    Public Shared Function SetNetworkSecurityOptions(opts As Integer) As String
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.SetNetworkSecurityOptions(opts)
    End Function
    '''*
    ''' <summary>
    '''   Modifies the network connection delay for <c>yRegisterHub()</c> and <c>yUpdateDeviceList()</c>.
    ''' <para>
    '''   This delay impacts only the YoctoHubs and VirtualHub
    '''   which are accessible through the network. By default, this delay is of 20000 milliseconds,
    '''   but depending on your network you may want to change this delay,
    '''   gor example if your network infrastructure is based on a GSM connection.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="networkMsTimeout">
    '''   the network connection delay in milliseconds.
    ''' @noreturn
    ''' </param>
    '''/
    Public Shared Sub SetNetworkTimeout(networkMsTimeout As Integer)
        YAPI.InitAPI(0, Nothing)
        _yapiContext.SetNetworkTimeout(networkMsTimeout)
    End Sub
    '''*
    ''' <summary>
    '''   Returns the network connection delay for <c>yRegisterHub()</c> and <c>yUpdateDeviceList()</c>.
    ''' <para>
    '''   This delay impacts only the YoctoHubs and VirtualHub
    '''   which are accessible through the network. By default, this delay is of 20000 milliseconds,
    '''   but depending on your network you may want to change this delay,
    '''   for example if your network infrastructure is based on a GSM connection.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the network connection delay in milliseconds.
    ''' </returns>
    '''/
    Public Shared Function GetNetworkTimeout() As Integer
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.GetNetworkTimeout()
    End Function
    '''*
    ''' <summary>
    '''   Change the validity period of the data loaded by the library.
    ''' <para>
    '''   By default, when accessing a module, all the attributes of the
    '''   module functions are automatically kept in cache for the standard
    '''   duration (5 ms). This method can be used to change this standard duration,
    '''   for example in order to reduce network or USB traffic. This parameter
    '''   does not affect value change callbacks
    '''   Note: This function must be called after <c>yInitAPI</c>.
    ''' </para>
    ''' </summary>
    ''' <param name="cacheValidityMs">
    '''   an integer corresponding to the validity attributed to the
    '''   loaded function parameters, in milliseconds.
    ''' @noreturn
    ''' </param>
    '''/
    Public Shared Sub SetCacheValidity(cacheValidityMs As Long)
        YAPI.InitAPI(0, Nothing)
        _yapiContext.SetCacheValidity(cacheValidityMs)
    End Sub
    '''*
    ''' <summary>
    '''   Returns the validity period of the data loaded by the library.
    ''' <para>
    '''   This method returns the cache validity of all attributes
    '''   module functions.
    '''   Note: This function must be called after <c>yInitAPI </c>.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the validity attributed to the
    '''   loaded function parameters, in milliseconds
    ''' </returns>
    '''/
    Public Shared Function GetCacheValidity() As Long
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.GetCacheValidity()
    End Function
    Public Shared Function nextHubInUseInternal(hubref As Integer) As YHub
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.nextHubInUseInternal(hubref)
    End Function
    Public Shared Function getYHubObj(hubref As Integer) As YHub
        YAPI.InitAPI(0, Nothing)
        return _yapiContext.getYHubObj(hubref)
    End Function
   REM --- (end of generated code: YAPIContext yapiwrapper)


    '''*
    ''' <summary>
    '''   Returns the version identifier for the Yoctopuce library in use.
    ''' <para>
    '''   The version is a string in the form <c>"Major.Minor.Build"</c>,
    '''   for instance <c>"1.01.5535"</c>. For languages using an external
    '''   DLL (for instance C#, VisualBasic or Delphi), the character string
    '''   includes as well the DLL version, for instance
    '''   <c>"1.01.5535 (1.01.5439)"</c>.
    ''' </para>
    ''' <para>
    '''   If you want to verify in your code that the library version is
    '''   compatible with the version that you have used during development,
    '''   verify that the major number is strictly equal and that the minor
    '''   number is greater or equal. The build number is not relevant
    '''   with respect to the library compatibility.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a character string describing the library version.
    ''' </returns>
    '''/
    Public Shared Function GetAPIVersion() As String
      Dim version As String = ""
      Dim apidate As String = ""
      yapiGetAPIVersion(version, apidate)
      Return YOCTO_API_VERSION_STR + "." + YOCTO_API_BUILD_NO + " (" + version + ")"
    End Function


    '''*
    ''' <summary>
    '''   Initializes the Yoctopuce programming library explicitly.
    ''' <para>
    '''   It is not strictly needed to call <c>yInitAPI()</c>, as the library is
    '''   automatically  initialized when calling <c>yRegisterHub()</c> for the
    '''   first time.
    ''' </para>
    ''' <para>
    '''   When <c>YAPI.DETECT_NONE</c> is used as detection <c>mode</c>,
    '''   you must explicitly use <c>yRegisterHub()</c> to point the API to the
    '''   VirtualHub on which your devices are connected before trying to access them.
    ''' </para>
    ''' </summary>
    ''' <param name="mode">
    '''   an integer corresponding to the type of automatic
    '''   device detection to use. Possible values are
    '''   <c>YAPI.DETECT_NONE</c>, <c>YAPI.DETECT_USB</c>, <c>YAPI.DETECT_NET</c>,
    '''   and <c>YAPI.DETECT_ALL</c>.
    ''' </param>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure returns a negative error code.
    ''' </para>
    '''/
    Public Shared Function InitAPI(ByVal mode As Integer, ByRef errmsg As String) As Integer
      Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE
      Dim version As String = ""
      Dim apidate As String = ""
      Dim i As Integer
      Dim dll_ver As yu16

      If (apiInitialized) Then
        InitAPI = YAPI_SUCCESS
        Exit Function
      End If
      Try
        dll_ver = yapiGetAPIVersion(version, apidate)
      Catch ex As System.DllNotFoundException
        errmsg = "Unable to load yapi.dll (" + ex.Message + ")"
        InitAPI = YAPI_FILE_NOT_FOUND
        Exit Function
      Catch ex As System.BadImageFormatException
        If IntPtr.Size = 4 Then
          errmsg = "Invalid yapi.dll (Using 64 bits yapi.dll with 32 bit application)"
        Else
          errmsg = "Invalid yapi.dll (Using 32 bits yapi.dll with 64 bit application)"
        End If
        InitAPI = YAPI_VERSION_MISMATCH
        Exit Function
      End Try

      If (YOCTO_API_VERSION_BCD <> dll_ver) Then
        errmsg = "yapi.dll does does not match the version of the Libary"
        errmsg += "(Libary=" + YOCTO_API_VERSION_STR + "." + YOCTO_API_BUILD_NO + " yapi.dll=" + version.ToString + ")"
        InitAPI = YAPI_VERSION_MISMATCH
        Exit Function
      End If

      vbmodule_initialization()

      buffer.Length = 0
      res = _yapiInitAPI(mode, buffer)
      errmsg = buffer.ToString()
      If (YISERR(res)) Then
        InitAPI = res
        Exit Function
      End If

      _yapiRegisterDeviceArrivalCallback(Marshal.GetFunctionPointerForDelegate(native_yDeviceArrivalDelegate))
      _yapiRegisterDeviceRemovalCallback(Marshal.GetFunctionPointerForDelegate(native_yDeviceRemovalDelegate))
      _yapiRegisterDeviceChangeCallback(Marshal.GetFunctionPointerForDelegate(native_yDeviceChangeDelegate))
      _yapiRegisterDeviceConfigChangeCallback(Marshal.GetFunctionPointerForDelegate(native_yDeviceConfigChangeDelegate))
      _yapiRegisterBeaconCallback(Marshal.GetFunctionPointerForDelegate(native_yBeaconChangeDelegate))
      _yapiRegisterFunctionUpdateCallback(Marshal.GetFunctionPointerForDelegate(native_yFunctionUpdateDelegate))
      _yapiRegisterTimedReportCallback(Marshal.GetFunctionPointerForDelegate(native_yTimedReportDelegate))
      _yapiRegisterLogFunction(Marshal.GetFunctionPointerForDelegate(native_yLogFunctionDelegate))
      _yapiRegisterDeviceLogCallback(Marshal.GetFunctionPointerForDelegate(native_yDeviceLogDelegate))
      _yapiRegisterHubDiscoveryCallback(Marshal.GetFunctionPointerForDelegate(native_HubDiscoveryDelegate))
      For i = 1 To 20
        RegisterCalibrationHandler(i, AddressOf yLinearCalibrationHandler)
      Next i
      RegisterCalibrationHandler(YOCTO_CALIB_TYPE_OFS, AddressOf yLinearCalibrationHandler)

      apiInitialized = True
      InitAPI = res
    End Function

    '''*
    ''' <summary>
    '''   Waits for all pending communications with Yoctopuce devices to be
    '''   completed then frees dynamically allocated resources used by
    '''   the Yoctopuce library.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   From an operating system standpoint, it is generally not required to call
    '''   this function since the OS will automatically free allocated resources
    '''   once your program is completed. However, there are two situations when
    '''   you may really want to use that function:
    ''' </para>
    ''' <para>
    '''   - Free all dynamically allocated memory blocks in order to
    '''   track a memory leak.
    ''' </para>
    ''' <para>
    '''   - Send commands to devices right before the end
    '''   of the program. Since commands are sent in an asynchronous way
    '''   the program could exit before all commands are effectively sent.
    ''' </para>
    ''' <para>
    '''   You should not call any other library function after calling
    '''   <c>yFreeAPI()</c>, or your program will crash.
    ''' </para>
    ''' </summary>
    '''/
    Public Shared Sub FreeAPI()
      If Not (apiInitialized) Then Exit Sub
      _yapiFreeAPI()
      apiInitialized = False
      vbmodule_cleanup()
    End Sub

    '''*
    ''' <summary>
    '''   Set up the Yoctopuce library to use modules connected on a given machine.
    ''' <para>
    '''   Idealy this
    '''   call will be made once at the begining of your application.  The
    '''   parameter will determine how the API will work. Use the following values:
    ''' </para>
    ''' <para>
    '''   <b>usb</b>: When the <c>usb</c> keyword is used, the API will work with
    '''   devices connected directly to the USB bus. Some programming languages such a JavaScript,
    '''   PHP, and Java don't provide direct access to USB hardware, so <c>usb</c> will
    '''   not work with these. In this case, use a VirtualHub or a networked YoctoHub (see below).
    ''' </para>
    ''' <para>
    '''   <b><i>x.x.x.x</i></b> or <b><i>hostname</i></b>: The API will use the devices connected to the
    '''   host with the given IP address or hostname. That host can be a regular computer
    '''   running a <i>native VirtualHub</i>, a <i>VirtualHub for web</i> hosted on a server,
    '''   or a networked YoctoHub such as YoctoHub-Ethernet or
    '''   YoctoHub-Wireless. If you want to use the VirtualHub running on you local
    '''   computer, use the IP address 127.0.0.1. If the given IP is unresponsive, <c>yRegisterHub</c>
    '''   will not return until a time-out defined by <c>ySetNetworkTimeout</c> has elapsed.
    '''   However, it is possible to preventively test a connection  with <c>yTestHub</c>.
    '''   If you cannot afford a network time-out, you can use the non-blocking <c>yPregisterHub</c>
    '''   function that will establish the connection as soon as it is available.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   <b>callback</b>: that keyword make the API run in "<i>HTTP Callback</i>" mode.
    '''   This a special mode allowing to take control of Yoctopuce devices
    '''   through a NAT filter when using a VirtualHub or a networked YoctoHub. You only
    '''   need to configure your hub to call your server script on a regular basis.
    '''   This mode is currently available for PHP and Node.JS only.
    ''' </para>
    ''' <para>
    '''   Be aware that only one application can use direct USB access at a
    '''   given time on a machine. Multiple access would cause conflicts
    '''   while trying to access the USB modules. In particular, this means
    '''   that you must stop the VirtualHub software before starting
    '''   an application that uses direct USB access. The workaround
    '''   for this limitation is to set up the library to use the VirtualHub
    '''   rather than direct USB access.
    ''' </para>
    ''' <para>
    '''   If access control has been activated on the hub, virtual or not, you want to
    '''   reach, the URL parameter should look like:
    ''' </para>
    ''' <para>
    '''   <c>http://username:password@address:port</c>
    ''' </para>
    ''' <para>
    '''   You can call <i>RegisterHub</i> several times to connect to several machines. On
    '''   the other hand, it is useless and even counterproductive to call <i>RegisterHub</i>
    '''   with to same address multiple times during the life of the application.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="url">
    '''   a string containing either <c>"usb"</c>,<c>"callback"</c> or the
    '''   root URL of the hub to monitor
    ''' </param>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure returns a negative error code.
    ''' </para>
    '''/
    Public Shared Function RegisterHub(ByVal url As String, ByRef errmsg As String) As Integer
      Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE

      res = InitAPI(0, errmsg)
      If YISERR(res) Then
        Return res
      End If
      buffer.Length = 0
      res = _yapiRegisterHub(New StringBuilder(url), buffer)
      If (YISERR(res)) Then
        errmsg = buffer.ToString()
      End If
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Fault-tolerant alternative to <c>yRegisterHub()</c>.
    ''' <para>
    '''   This function has the same
    '''   purpose and same arguments as <c>yRegisterHub()</c>, but does not trigger
    '''   an error when the selected hub is not available at the time of the function call.
    '''   If the connexion cannot be established immediately, a background task will automatically
    '''   perform periodic retries. This makes it possible to register a network hub independently of the current
    '''   connectivity, and to try to contact it only when a device is actively needed.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="url">
    '''   a string containing either <c>"usb"</c>,<c>"callback"</c> or the
    '''   root URL of the hub to monitor
    ''' </param>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure returns a negative error code.
    ''' </para>
    '''/
    Public Shared Function PreregisterHub(ByVal url As String, ByRef errmsg As String) As Integer
      Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE

      res = InitAPI(0, errmsg)
      If YISERR(res) Then
        Return res
      End If
      buffer.Length = 0
      res = _yapiPreregisterHub(New StringBuilder(url), buffer)
      If (YISERR(res)) Then
        errmsg = buffer.ToString()
      End If
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Set up the Yoctopuce library to no more use modules connected on a previously
    '''   registered machine with RegisterHub.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="url">
    '''   a string containing either <c>"usb"</c> or the
    '''   root URL of the hub to monitor
    ''' </param>
    '''/
    Public Shared Sub UnregisterHub(ByVal url As String)
      If Not (apiInitialized) Then Exit Sub
      _yapiUnregisterHub(New StringBuilder(url))
    End Sub


    '''*
    ''' <summary>
    '''   Test if the hub is reachable.
    ''' <para>
    '''   This method do not register the hub, it only test if the
    '''   hub is usable. The url parameter follow the same convention as the <c>yRegisterHub</c>
    '''   method. This method is useful to verify the authentication parameters for a hub. It
    '''   is possible to force this method to return after mstimeout milliseconds.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="url">
    '''   a string containing either <c>"usb"</c>,<c>"callback"</c> or the
    '''   root URL of the hub to monitor
    ''' </param>
    ''' <param name="mstimeout">
    '''   the number of millisecond available to test the connection.
    ''' </param>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure returns a negative error code.
    ''' </para>
    '''/
    Public Shared Function TestHub(ByVal url As String, ByVal mstimeout As Integer, ByRef errmsg As String) As Integer
      Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE

      buffer.Length = 0
      res = _yapiTestHub(New StringBuilder(url), mstimeout, buffer)
      If (YISERR(res)) Then
        errmsg = buffer.ToString()
      End If
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Triggers a (re)detection of connected Yoctopuce modules.
    ''' <para>
    '''   The library searches the machines or USB ports previously registered using
    '''   <c>yRegisterHub()</c>, and invokes any user-defined callback function
    '''   in case a change in the list of connected devices is detected.
    ''' </para>
    ''' <para>
    '''   This function can be called as frequently as desired to refresh the device list
    '''   and to make the application aware of hot-plug events. However, since device
    '''   detection is quite a heavy process, UpdateDeviceList shouldn't be called more
    '''   than once every two seconds.
    ''' </para>
    ''' </summary>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure returns a negative error code.
    ''' </para>
    '''/
    Public Shared Function UpdateDeviceList(ByRef errmsg As String) As YRETCODE
      Dim errbuffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE
      Dim p As PlugEvent

      If Not (apiInitialized) Then
        res = InitAPI(0, errmsg)
        If (YISERR(res)) Then Return res
      End If
      errbuffer.Length = 0
      res = _yapiUpdateDeviceList(0, errbuffer)
      If (YISERR(res)) Then
        errmsg = errbuffer.ToString()
      Else
        res = _yapiHandleEvents(errbuffer)
        If (YISERR(res)) Then
          errmsg = errbuffer.ToString()
          Return res
        End If
      End If
      While (_PlugEvents.Count > 0)
        yapiLockDeviceCallBack(errmsg)
        p = _PlugEvents(0)
        _PlugEvents.RemoveAt(0)
        yapiUnlockDeviceCallBack(errmsg)
        p.invoke()
      End While
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Maintains the device-to-library communication channel.
    ''' <para>
    '''   If your program includes significant loops, you may want to include
    '''   a call to this function to make sure that the library takes care of
    '''   the information pushed by the modules on the communication channels.
    '''   This is not strictly necessary, but it may improve the reactivity
    '''   of the library for the following commands.
    ''' </para>
    ''' <para>
    '''   This function may signal an error in case there is a communication problem
    '''   while contacting a module.
    ''' </para>
    ''' </summary>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure returns a negative error code.
    ''' </para>
    '''/
    Public Shared Function HandleEvents(ByRef errmsg As String) As YRETCODE

      Dim errBuffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE
      Dim ev As DataEvent

      errBuffer.Length = 0
      res = _yapiHandleEvents(errBuffer)

      If (YISERR(res)) Then
        errmsg = errBuffer.ToString()
        HandleEvents = res
        Exit Function
      End If

      While (_DataEvents.Count > 0)
        yapiLockFunctionCallBack(errmsg)
        If (_DataEvents.Count = 0) Then      REM not sure this if is really ussefull
          yapiUnlockFunctionCallBack(errmsg)

        Else
          ev = _DataEvents(0)
          _DataEvents.RemoveAt(0)
          yapiUnlockFunctionCallBack(errmsg)
          ev.invoke()
        End If

      End While
      HandleEvents = YAPI_SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Pauses the execution flow for a specified duration.
    ''' <para>
    '''   This function implements a passive waiting loop, meaning that it does not
    '''   consume CPU cycles significantly. The processor is left available for
    '''   other threads and processes. During the pause, the library nevertheless
    '''   reads from time to time information from the Yoctopuce modules by
    '''   calling <c>yHandleEvents()</c>, in order to stay up-to-date.
    ''' </para>
    ''' <para>
    '''   This function may signal an error in case there is a communication problem
    '''   while contacting a module.
    ''' </para>
    ''' </summary>
    ''' <param name="ms_duration">
    '''   an integer corresponding to the duration of the pause,
    '''   in milliseconds.
    ''' </param>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure returns a negative error code.
    ''' </para>
    '''/
    Public Shared Function Sleep(ByVal ms_duration As Integer, ByRef errmsg As String) As Integer

      Dim errBuffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim timeout As yu64
      Dim res As Integer


      timeout = CULng(GetTickCount() + ms_duration)
      res = YAPI_SUCCESS
      errBuffer.Length = 0

      Do
        res = HandleEvents(errmsg)
        If (YISERR(res)) Then
          Sleep = res
          Exit Function
        End If
        If (GetTickCount() < timeout) Then
          res = _yapiSleep(2, errBuffer)
          If (YISERR(res)) Then
            Sleep = res
            errmsg = errBuffer.ToString()
            Exit Function
          End If
        End If

      Loop Until GetTickCount() >= timeout
      errmsg = errBuffer.ToString()
      Sleep = res
    End Function

    '''*
    ''' <summary>
    '''   Force a hub discovery, if a callback as been registered with <c>yRegisterHubDiscoveryCallback</c> it
    '''   will be called for each net work hub that will respond to the discovery.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    '''   On failure returns a negative error code.
    ''' </returns>
    '''/
    Public Shared Function TriggerHubDiscovery(ByRef errmsg As String) As Integer
      Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE

      res = InitAPI(0, errmsg)
      If YISERR(res) Then
        Return res
      End If
      buffer.Length = 0
      res = _yapiTriggerHubDiscovery(buffer)
      If (YISERR(res)) Then
        errmsg = buffer.ToString()
      End If
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of a monotone millisecond-based time counter.
    ''' <para>
    '''   This counter can be used to compute delays in relation with
    '''   Yoctopuce devices, which also uses the millisecond as timebase.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a long integer corresponding to the millisecond counter.
    ''' </returns>
    '''/
    Public Shared Function GetTickCount() As Long
      GetTickCount = CLng(_yapiGetTickCount())
    End Function

    '''*
    ''' <summary>
    '''   Checks if a given string is valid as logical name for a module or a function.
    ''' <para>
    '''   A valid logical name has a maximum of 19 characters, all among
    '''   <c>A...Z</c>, <c>a...z</c>, <c>0...9</c>, <c>_</c>, and <c>-</c>.
    '''   If you try to configure a logical name with an incorrect string,
    '''   the invalid characters are ignored.
    ''' </para>
    ''' </summary>
    ''' <param name="name">
    '''   a string containing the name to check.
    ''' </param>
    ''' <returns>
    '''   <c>true</c> if the name is valid, <c>false</c> otherwise.
    ''' </returns>
    '''/
    Public Shared Function CheckLogicalName(ByVal name As String) As Boolean
      If (_yapiCheckLogicalName(New StringBuilder(name)) = 0) Then
        CheckLogicalName = False
      Else
        CheckLogicalName = True
      End If
    End Function


    '''*
    ''' <summary>
    '''   Registers a log callback function.
    ''' <para>
    '''   This callback will be called each time
    '''   the API have something to say. Quite useful to debug the API.
    ''' </para>
    ''' </summary>
    ''' <param name="logfun">
    '''   a procedure taking a string parameter, or <c>Nothing</c>
    '''   to unregister a previously registered  callback.
    ''' </param>
    '''/
    Public Shared Sub RegisterLogFunction(ByVal logfun As yLogFunc)
      ylog = logfun
    End Sub

    '''*
    ''' <summary>
    '''   Disables the use of exceptions to report runtime errors.
    ''' <para>
    '''   When exceptions are disabled, every function returns a specific
    '''   error value which depends on its type and which is documented in
    '''   this reference manual.
    ''' </para>
    ''' </summary>
    '''/
    Public Shared Sub DisableExceptions()
      YAPI.ExceptionsDisabled = True
    End Sub

    '''*
    ''' <summary>
    '''   Re-enables the use of exceptions for runtime error handling.
    ''' <para>
    '''   Be aware than when exceptions are enabled, every function that fails
    '''   triggers an exception. If the exception is not caught by the user code,
    '''   it either fires the debugger or aborts (i.e. crash) the program.
    ''' </para>
    ''' </summary>
    '''/
    Public Shared Sub EnableExceptions()
      YAPI.ExceptionsDisabled = False
    End Sub


    '''*
    ''' <summary>
    '''   Register a callback function, to be called each time
    '''   a device is plugged.
    ''' <para>
    '''   This callback will be invoked while <c>yUpdateDeviceList</c>
    '''   is running. You will have to call this function on a regular basis.
    ''' </para>
    ''' </summary>
    ''' <param name="arrivalCallback">
    '''   a procedure taking a <c>YModule</c> parameter, or <c>Nothing</c>
    '''   to unregister a previously registered  callback.
    ''' </param>
    '''/
    Public Shared Sub RegisterDeviceArrivalCallback(ByVal arrivalCallback As yDeviceUpdateFunc)
      yArrival = arrivalCallback
      If Not (arrivalCallback Is Nothing) Then
        Dim m As YModule = YModule.FirstModule()
        Dim errmsg As String = ""
        While m IsNot Nothing
          If (m.isOnline()) Then
            yapiLockDeviceCallBack(errmsg)
            native_yDeviceArrivalCallback(m.get_functionDescriptor())
            yapiUnlockDeviceCallBack(errmsg)
          End If
          m = m.nextModule()
        End While
      End If
    End Sub

    '''*
    ''' <summary>
    '''   Register a callback function, to be called each time
    '''   a device is unplugged.
    ''' <para>
    '''   This callback will be invoked while <c>yUpdateDeviceList</c>
    '''   is running. You will have to call this function on a regular basis.
    ''' </para>
    ''' </summary>
    ''' <param name="removalCallback">
    '''   a procedure taking a <c>YModule</c> parameter, or <c>Nothing</c>
    '''   to unregister a previously registered  callback.
    ''' </param>
    '''/
    Public Shared Sub RegisterDeviceRemovalCallback(ByVal removalCallback As yDeviceUpdateFunc)
      yRemoval = removalCallback
      If yRemoval IsNot Nothing Then
        _yapiRegisterDeviceRemovalCallback(Marshal.GetFunctionPointerForDelegate(native_yDeviceRemovalDelegate))
      Else
        _yapiRegisterDeviceRemovalCallback(Nothing)
      End If
    End Sub

    '''*
    ''' <summary>
    '''   Register a callback function, to be called each time an Network Hub send
    '''   an SSDP message.
    ''' <para>
    '''   The callback has two string parameter, the first one
    '''   contain the serial number of the hub and the second contain the URL of the
    '''   network hub (this URL can be passed to RegisterHub). This callback will be invoked
    '''   while yUpdateDeviceList is running. You will have to call this function on a regular basis.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="hubDiscoveryCallback">
    '''   a procedure taking two string parameter, the serial
    '''   number and the hub URL. Use <c>Nothing</c> to unregister a previously registered  callback.
    ''' </param>
    '''/
    Public Shared Sub RegisterHubDiscoveryCallback(ByVal hubDiscoveryCallback As YHubDiscoveryCallback)
      Dim errmsg As String = ""
      _HubDiscoveryCallback = hubDiscoveryCallback
      YAPI.TriggerHubDiscovery(errmsg)
    End Sub

    Public Shared Sub RegisterDeviceChangeCallback(ByVal callback As yDeviceUpdateFunc)
      yChange = callback
      If yChange IsNot Nothing Then
        _yapiRegisterDeviceChangeCallback(Marshal.GetFunctionPointerForDelegate(native_yDeviceChangeDelegate))
      Else
        _yapiRegisterDeviceChangeCallback(Nothing)
      End If
    End Sub
  End Class


  <StructLayout(LayoutKind.Sequential, Pack:=1, CharSet:=CharSet.Ansi)>
  Public Structure yDeviceSt
    Dim vendorid As yu16
    Dim deviceid As yu16
    Dim devrelease As yu16
    Dim nbinbterfaces As yu16
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=YOCTO_MANUFACTURER_LEN)> Public manufacturer As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=YOCTO_PRODUCTNAME_LEN)> Public productname As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=YOCTO_SERIAL_LEN)> Public serial As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=YOCTO_LOGICAL_LEN)> Public logicalname As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=YOCTO_FIRMWARE_LEN)> Public firmware As String
    Dim beacon As yu8
    Dim pad As yu8
  End Structure

  Public Const YIOHDL_SIZE As Integer = 8

  <StructLayout(LayoutKind.Sequential, Pack:=1, CharSet:=CharSet.Ansi)>
  Public Structure YIOHDL
    <MarshalAs(UnmanagedType.U1, SizeConst:=YIOHDL_SIZE)> Public raw As yu8
  End Structure

  Enum yDEVICE_PROP
    PROP_VENDORID
    PROP_DEVICEID
    PROP_DEVRELEASE
    PROP_FIRMWARELEVEL
    PROP_MANUFACTURER
    PROP_PRODUCTNAME
    PROP_SERIAL
    PROP_LOGICALNAME
    PROP_URL
  End Enum


  Enum yFACE_STATUS
    YFACE_EMPTY
    YFACE_RUNNING
    YFACE_ERROR
  End Enum

  Dim ySerialList As String()
  Dim yHandleArray As YHANDLE

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Delegate Function yFlashCallback(ByVal stepnumber As yu32, ByVal totalStep As yu32, ByVal context As IntPtr) As Integer

  <StructLayout(LayoutKind.Sequential, Pack:=1, CharSet:=CharSet.Ansi)>
  Public Structure yFlashArg
    Dim OSDeviceName As StringBuilder      REM device windows name on os (used to acces device)
    Dim serial2assign As StringBuilder      REM serial number of the device
    Dim firmwarePtr As IntPtr        REM pointer to the content of the Hex file
    Dim firmwareLen As yu32            REM len of the Hexfile
    Dim progress As yFlashCallback
    Dim context As IntPtr
  End Structure


  REM --- (generated code: YFunction globals)

  REM Yoctopuce error codes, also used by default as function return value
  Public Const YAPI_SUCCESS As Integer = 0         REM everything worked all right
  Public Const YAPI_NOT_INITIALIZED As Integer = -1 REM call yInitAPI() first !
  Public Const YAPI_INVALID_ARGUMENT As Integer = -2 REM one of the arguments passed to the function is invalid
  Public Const YAPI_NOT_SUPPORTED As Integer = -3  REM the operation attempted is (currently) not supported
  Public Const YAPI_DEVICE_NOT_FOUND As Integer = -4 REM the requested device is not reachable
  Public Const YAPI_VERSION_MISMATCH As Integer = -5 REM the device firmware is incompatible with this API version
  Public Const YAPI_DEVICE_BUSY As Integer = -6    REM the device is busy with another task and cannot answer
  Public Const YAPI_TIMEOUT As Integer = -7        REM the device took too long to provide an answer
  Public Const YAPI_IO_ERROR As Integer = -8       REM there was an I/O problem while talking to the device
  Public Const YAPI_NO_MORE_DATA As Integer = -9   REM there is no more data to read from
  Public Const YAPI_EXHAUSTED As Integer = -10     REM you have run out of a limited resource, check the documentation
  Public Const YAPI_DOUBLE_ACCES As Integer = -11  REM you have two process that try to access to the same device
  Public Const YAPI_UNAUTHORIZED As Integer = -12  REM unauthorized access to password-protected device
  Public Const YAPI_RTC_NOT_READY As Integer = -13 REM real-time clock has not been initialized (or time was lost)
  Public Const YAPI_FILE_NOT_FOUND As Integer = -14 REM the file is not found
  Public Const YAPI_SSL_ERROR As Integer = -15     REM Error reported by mbedSSL
  Public Const YAPI_RFID_SOFT_ERROR As Integer = -16 REM Recoverable error with RFID tag (eg. tag out of reach), check YRfidStatus for details
  Public Const YAPI_RFID_HARD_ERROR As Integer = -17 REM Serious RFID error (eg. write-protected, out-of-boundary), check YRfidStatus for details
  Public Const YAPI_BUFFER_TOO_SMALL As Integer = -18 REM The buffer provided is too small
  Public Const YAPI_DNS_ERROR As Integer = -19     REM Error during name resolutions (invalid hostname or dns communication error)
  Public Const YAPI_SSL_UNK_CERT As Integer = -20  REM The certificate is not correctly signed by the trusted CA

  REM TLS / SSL definitions
  Public Const YAPI_NO_TRUSTED_CA_CHECK As Integer = 1 REM Disables certificate checking
  Public Const YAPI_NO_EXPIRATION_CHECK As Integer = 2 REM Disables certificate expiration date checking
  Public Const YAPI_NO_HOSTNAME_CHECK As Integer = 4 REM Disable hostname checking
  Public Const YAPI_LEGACY As Integer = 8          REM Allow non secure connection (similar to v1.10)

  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YFunctionValueCallback(ByVal func As YFunction, ByVal value As String)
  Public Delegate Sub YFunctionTimedReportCallback(ByVal func As YFunction, ByVal measure As YMeasure)
  REM --- (end of generated code: YFunction globals)

  REM --- (generated code: YModule globals)

  Public Const Y_PRODUCTNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SERIALNUMBER_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PRODUCTID_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PRODUCTRELEASE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_FIRMWARERELEASE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PERSISTENTSETTINGS_LOADED As Integer = 0
  Public Const Y_PERSISTENTSETTINGS_SAVED As Integer = 1
  Public Const Y_PERSISTENTSETTINGS_MODIFIED As Integer = 2
  Public Const Y_PERSISTENTSETTINGS_INVALID As Integer = -1
  Public Const Y_LUMINOSITY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_BEACON_OFF As Integer = 0
  Public Const Y_BEACON_ON As Integer = 1
  Public Const Y_BEACON_INVALID As Integer = -1
  Public Const Y_UPTIME_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_USBCURRENT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_REBOOTCOUNTDOWN_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_USERVAR_INVALID As Integer = YAPI.INVALID_INT
  Public Delegate Sub YModuleLogCallback(ByVal modul As YModule, ByVal logline As String)
  Public Delegate Sub YModuleConfigChangeCallback(ByVal modul As YModule)
  Public Delegate Sub YModuleBeaconCallback(ByVal modul As YModule, ByVal beacon As Integer)
  Public Delegate Sub YModuleValueCallback(ByVal func As YModule, ByVal value As String)
  Public Delegate Sub YModuleTimedReportCallback(ByVal func As YModule, ByVal measure As YMeasure)
  REM --- (end of generated code: YModule globals)

  REM --- (generated code: YSensor globals)

  Public Const Y_UNIT_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CURRENTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_LOWESTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_HIGHESTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_CURRENTRAWVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_LOGFREQUENCY_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_REPORTFREQUENCY_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVMODE_IMMEDIATE As Integer = 0
  Public Const Y_ADVMODE_PERIOD_AVG As Integer = 1
  Public Const Y_ADVMODE_PERIOD_MIN As Integer = 2
  Public Const Y_ADVMODE_PERIOD_MAX As Integer = 3
  Public Const Y_ADVMODE_INVALID As Integer = -1
  Public Const Y_CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_RESOLUTION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_SENSORSTATE_INVALID As Integer = YAPI.INVALID_INT
  Public Delegate Sub YSensorValueCallback(ByVal func As YSensor, ByVal value As String)
  Public Delegate Sub YSensorTimedReportCallback(ByVal func As YSensor, ByVal measure As YMeasure)
  REM --- (end of generated code: YSensor globals)

  REM --- (generated code: YFirmwareUpdate globals)

  REM --- (end of generated code: YFirmwareUpdate globals)

  REM --- (generated code: YDataStream globals)

  REM --- (end of generated code: YDataStream globals)

  REM --- (generated code: YMeasure globals)

  REM --- (end of generated code: YMeasure globals)

  REM --- (generated code: YDataSet globals)

  REM --- (end of generated code: YDataSet globals)

  REM --- (generated code: YConsolidatedDataSet globals)

  REM --- (end of generated code: YConsolidatedDataSet globals)

  Public Class YAPI_Exception
    Inherits ApplicationException
    Public errorType As YRETCODE

    Public Sub New(ByVal errType As YRETCODE, ByVal errMsg As String)
      MyBase.New(errMsg)
      errorType = errorType
    End Sub
  End Class

  Dim YDevice_devCache As List(Of YDevice)


  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub HTTPRequestCallback(ByVal device As YDevice, ByRef context As blockingCallbackCtx, ByVal returnval As YRETCODE, ByVal result As String, ByVal errmsg As String)

  REM - Types used for public yocto_api callbacks
  Public Delegate Sub yLogFunc(ByVal log As String)
  Public Delegate Sub yDeviceUpdateFunc(ByVal modul As YModule)
  Public Delegate Sub yFunctionUpdateFunc(ByVal modul As YModule, ByVal functionId As String, ByVal functionName As String, ByVal functionValue As String)
  Public Delegate Function yCalibrationHandler(ByVal rawValue As Double, ByVal calibType As Integer, ByVal parameters As List(Of Integer), ByVal rawValues As List(Of Double), ByVal refValues As List(Of Double)) As Double
  Public Delegate Sub YHubDiscoveryCallback(ByVal serial As String, ByVal url As String)


  REM - Types used for internal yapi callbacks
  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub _yapiLogFunc(ByVal log As IntPtr, ByVal loglen As yu32)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub _yapiDeviceUpdateFunc(ByVal dev As YDEV_DESCR)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub _yapiBeaconUpdateFunc(ByVal dev As YDEV_DESCR, ByVal beacon As Integer)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub _yapiFunctionUpdateFunc(ByVal dev As YFUN_DESCR, ByVal value As IntPtr)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub _yapiTimedReportFunc(ByVal dev As YFUN_DESCR, ByVal timestamp As Double, ByVal data As IntPtr, ByVal len As System.UInt32, ByVal duration As Double)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub _yapiHubDiscoveryCallback(ByVal serial As IntPtr, ByVal url As IntPtr)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub _yapiDeviceLogCallback(ByVal dev As YFUN_DESCR, ByVal value As IntPtr)


  REM - Variables used to store public yocto_api callbacks
  Private ylog As yLogFunc = Nothing
  Private yArrival As yDeviceUpdateFunc = Nothing
  Private yRemoval As yDeviceUpdateFunc = Nothing
  Private yChange As yDeviceUpdateFunc = Nothing
  Private _HubDiscoveryCallback As YHubDiscoveryCallback = Nothing

  Public Function YISERR(ByVal retcode As YRETCODE) As Boolean
    If retcode < 0 Then YISERR = True Else YISERR = False
  End Function

  Public Class blockingCallbackCtx
    Public res As YRETCODE
    Public response As String
    Public errmsg As String
  End Class

  Public Sub YblockingCallback(ByVal device As YDevice, ByRef context As blockingCallbackCtx, ByVal returnval As YRETCODE, ByVal result As String, ByVal errmsg As String)
    context.res = returnval
    context.response = result
    context.errmsg = errmsg
  End Sub


  REM --- (generated code: YFirmwareUpdate class start)

  '''*
  ''' <summary>
  '''   The <c>YFirmwareUpdate</c> class let you control the firmware update of a Yoctopuce
  '''   module.
  ''' <para>
  '''   This class should not be instantiate directly, but instances should be retrieved
  '''   using the <c>YModule</c> method <c>module.updateFirmware</c>.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YFirmwareUpdate
    REM --- (end of generated code: YFirmwareUpdate class start)
    REM --- (generated code: YFirmwareUpdate definitions)
    REM --- (end of generated code: YFirmwareUpdate definitions)
    Public Const DATA_INVALID As Double = YAPI.INVALID_DOUBLE

    REM --- (generated code: YFirmwareUpdate attributes declaration)
    Protected _serial As String
    Protected _settings As Byte()
    Protected _firmwarepath As String
    Protected _progress_msg As String
    Protected _progress_c As Integer
    Protected _progress As Integer
    Protected _restore_step As Integer
    Protected _force As Boolean
    REM --- (end of generated code: YFirmwareUpdate attributes declaration)

    Public Sub New(serial As String, path As String, settings As Byte(), force As Boolean)
      _serial = serial
      _firmwarepath = path
      _settings = settings
      _force = force
      REM --- (generated code: YFirmwareUpdate attributes initialization)
      _settings = New Byte(){}
      _progress_c = 0
      _progress = 0
      _restore_step = 0
      REM --- (end of generated code: YFirmwareUpdate attributes initialization)
    End Sub

    Public Sub New(serial As String, path As String, settings As Byte())
      _serial = serial
      _firmwarepath = path
      _settings = settings
      _force = False
      REM --- (generated code: YFirmwareUpdate attributes initialization)
      _settings = New Byte(){}
      _progress_c = 0
      _progress = 0
      _restore_step = 0
      REM --- (end of generated code: YFirmwareUpdate attributes initialization)
    End Sub


    REM --- (generated code: YFirmwareUpdate private methods declaration)

    REM --- (end of generated code: YFirmwareUpdate private methods declaration)

    REM --- (generated code: YFirmwareUpdate public methods declaration)
    Public Overridable Function _processMore(newupdate As Integer) As Integer
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim m As YModule
      Dim res As Integer = 0
      Dim serial As String
      Dim firmwarepath As String
      Dim settings As String
      Dim prod_prefix As String
      Dim force As Integer = 0
      If ((Me._progress_c < 100) AndAlso (Me._progress_c <> YAPI.VERSION_MISMATCH)) Then
        serial = Me._serial
        firmwarepath = Me._firmwarepath
        settings = YAPI.DefaultEncoding.GetString(Me._settings)
        If (Me._force) Then
          force = 1
        Else
          force = 0
        End If
        res = _yapiUpdateFirmwareEx(New StringBuilder(serial), New StringBuilder(firmwarepath), New StringBuilder(settings), force, newupdate, errmsg)
        If ((res = YAPI.VERSION_MISMATCH) AndAlso ((Me._settings).Length <> 0)) Then
          Me._progress_c = res
          Me._progress_msg = errmsg.ToString()
          Return Me._progress
        End If
        If (res < 0) Then
          Me._progress = res
          Me._progress_msg = errmsg.ToString()
          Return res
        End If
        Me._progress_c = res
        Me._progress = (Me._progress_c * 9 \ 10)
        Me._progress_msg = errmsg.ToString()
      Else
        If (((Me._settings).Length <> 0) AndAlso ( Me._progress_c <> 101)) Then
          Me._progress_msg = "restoring settings"
          m = YModule.FindModule(Me._serial + ".module")
          If (Not (m.isOnline())) Then
            Return Me._progress
          End If
          If (Me._progress < 95) Then
            prod_prefix = (m.get_productName()).Substring(0, 8)
            If (prod_prefix = "YoctoHub") Then
              YAPI.Sleep(1000, Nothing)
              Me._progress = Me._progress + 1
              Return Me._progress
            Else
              Me._progress = 95
            End If
          End If
          If (Me._progress < 100) Then
            m.set_allSettingsAndFiles(Me._settings)
            m.saveToFlash()
            ReDim Me._settings(0-1)
            If (Me._progress_c = YAPI.VERSION_MISMATCH) Then
              Me._progress = YAPI.IO_ERROR
              Me._progress_msg = "Unable to update firmware"
            Else
              Me._progress = 100
              Me._progress_msg = "success"
            End If
          End If
        Else
          Me._progress = 100
          Me._progress_msg = "success"
        End If
      End If
      Return Me._progress
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of all the modules in "firmware update" mode.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an array of strings containing the serial numbers of devices in "firmware update" mode.
    ''' </returns>
    '''/
    Public Shared Function GetAllBootLoaders() As List(Of String)
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim smallbuff As StringBuilder = New StringBuilder(1024)
      Dim bigbuff As StringBuilder
      Dim buffsize As Integer = 0
      Dim fullsize As Integer
      Dim yapi_res As Integer = 0
      Dim bootloader_list As String
      Dim bootladers As List(Of String) = New List(Of String)()
      fullsize = 0
      yapi_res = _yapiGetBootloaders(smallbuff, 1024, fullsize, errmsg)
      If (yapi_res < 0) Then
        Return bootladers
      End If
      If (fullsize <= 1024) Then
        bootloader_list = smallbuff.ToString()
      Else
        buffsize = fullsize
        bigbuff = New StringBuilder(buffsize)
        yapi_res = _yapiGetBootloaders(bigbuff, buffsize, fullsize, errmsg)
        If (yapi_res < 0) Then
          bigbuff = Nothing
          Return bootladers
        Else
          bootloader_list = bigbuff.ToString()
        End If
        bigbuff = Nothing
      End If
      If (Not (bootloader_list = "")) Then
        bootladers = New List(Of String)(bootloader_list.Split(New Char() {","c}))
      End If
      Return bootladers
    End Function

    '''*
    ''' <summary>
    '''   Test if the byn file is valid for this module.
    ''' <para>
    '''   It is possible to pass a directory instead of a file.
    '''   In that case, this method returns the path of the most recent appropriate byn file. This method will
    '''   ignore any firmware older than minrelease.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="serial">
    '''   the serial number of the module to update
    ''' </param>
    ''' <param name="path">
    '''   the path of a byn file or a directory that contains byn files
    ''' </param>
    ''' <param name="minrelease">
    '''   a positive integer
    ''' </param>
    ''' <returns>
    '''   : the path of the byn file to use, or an empty string if no byn files matches the requirement
    ''' </returns>
    ''' <para>
    '''   On failure, returns a string that starts with "error:".
    ''' </para>
    '''/
    Public Shared Function CheckFirmware(serial As String, path As String, minrelease As Integer) As String
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim smallbuff As StringBuilder = New StringBuilder(1024)
      Dim bigbuff As StringBuilder
      Dim buffsize As Integer = 0
      Dim fullsize As Integer
      Dim res As Integer = 0
      Dim firmware_path As String
      Dim release As String
      fullsize = 0
      release = (minrelease).ToString()
      res = _yapiCheckFirmware(New StringBuilder(serial), New StringBuilder(release), New StringBuilder(path), smallbuff, 1024, fullsize, errmsg)
      If (res < 0) Then
        firmware_path = "error:" + errmsg.ToString()
        Return "error:" + errmsg.ToString()
      End If
      If (fullsize <= 1024) Then
        firmware_path = smallbuff.ToString()
      Else
        buffsize = fullsize
        bigbuff = New StringBuilder(buffsize)
        res = _yapiCheckFirmware(New StringBuilder(serial), New StringBuilder(release), New StringBuilder(path), bigbuff, buffsize, fullsize, errmsg)
        If (res < 0) Then
          firmware_path = "error:" + errmsg.ToString()
        Else
          firmware_path = bigbuff.ToString()
        End If
        bigbuff = Nothing
      End If
      Return firmware_path
    End Function

    '''*
    ''' <summary>
    '''   Returns the progress of the firmware update, on a scale from 0 to 100.
    ''' <para>
    '''   When the object is
    '''   instantiated, the progress is zero. The value is updated during the firmware update process until
    '''   the value of 100 is reached. The 100 value means that the firmware update was completed
    '''   successfully. If an error occurs during the firmware update, a negative value is returned, and the
    '''   error message can be retrieved with <c>get_progressMessage</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer in the range 0 to 100 (percentage of completion)
    '''   or a negative error code in case of failure.
    ''' </returns>
    '''/
    Public Overridable Function get_progress() As Integer
      If (Me._progress >= 0) Then
        Me._processMore(0)
      End If
      Return Me._progress
    End Function

    '''*
    ''' <summary>
    '''   Returns the last progress message of the firmware update process.
    ''' <para>
    '''   If an error occurs during the
    '''   firmware update process, the error message is returned
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string  with the latest progress message, or the error message.
    ''' </returns>
    '''/
    Public Overridable Function get_progressMessage() As String
      Return Me._progress_msg
    End Function

    '''*
    ''' <summary>
    '''   Starts the firmware update process.
    ''' <para>
    '''   This method starts the firmware update process in background. This method
    '''   returns immediately. You can monitor the progress of the firmware update with the <c>get_progress()</c>
    '''   and <c>get_progressMessage()</c> methods.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer in the range 0 to 100 (percentage of completion),
    '''   or a negative error code in case of failure.
    ''' </returns>
    ''' <para>
    '''   On failure returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function startUpdate() As Integer
      Dim err As String
      Dim leng As Integer = 0
      err = YAPI.DefaultEncoding.GetString(Me._settings)
      leng = (err).Length
      If ((leng >= 6) AndAlso ("error:" = (err).Substring(0, 6))) Then
        Me._progress = -1
        Me._progress_msg = (err).Substring(6, leng - 6)
      Else
        Me._progress = 0
        Me._progress_c = 0
        Me._processMore(1)
      End If
      Return Me._progress
    End Function



    REM --- (end of generated code: YFirmwareUpdate public methods declaration)
  End Class

  REM --- (generated code: YFirmwareUpdate functions)


  REM --- (end of generated code: YFirmwareUpdate functions)


  REM --- (generated code: YDataStream class start)

  '''*
  ''' <c>DataStream</c> objects represent bare recorded measure sequences,
  ''' exactly as found within the data logger present on Yoctopuce
  ''' sensors.
  ''' <para>
  '''   In most cases, it is not necessary to use <c>DataStream</c> objects
  '''   directly, as the <c>DataSet</c> objects (returned by the
  '''   <c>get_recordedData()</c> method from sensors and the
  '''   <c>get_dataSets()</c> method from the data logger) provide
  '''   a more convenient interface.
  ''' </para>
  '''/
  Public Class YDataStream
    REM --- (end of generated code: YDataStream class start)
    REM --- (generated code: YDataStream definitions)
    REM --- (end of generated code: YDataStream definitions)
    Public Const DATA_INVALID As Double = YAPI.INVALID_DOUBLE

    REM --- (generated code: YDataStream attributes declaration)
    Protected _parent As YFunction
    Protected _runNo As Integer
    Protected _utcStamp As Long
    Protected _nCols As Integer
    Protected _nRows As Integer
    Protected _startTime As Double
    Protected _duration As Double
    Protected _dataSamplesInterval As Double
    Protected _firstMeasureDuration As Double
    Protected _columnNames As List(Of String)
    Protected _functionId As String
    Protected _isClosed As Boolean
    Protected _isAvg As Boolean
    Protected _minVal As Double
    Protected _avgVal As Double
    Protected _maxVal As Double
    Protected _caltyp As Integer
    Protected _calpar As List(Of Integer)
    Protected _calraw As List(Of Double)
    Protected _calref As List(Of Double)
    Protected _values As List(Of List(Of Double))
    Protected _isLoaded As Boolean
    REM --- (end of generated code: YDataStream attributes declaration)
    Protected _calhdl As yCalibrationHandler

    Public Sub New(parent As YFunction)
      _parent = parent
      REM --- (generated code: YDataStream attributes initialization)
      _runNo = 0
      _utcStamp = 0
      _nCols = 0
      _nRows = 0
      _startTime = 0
      _duration = 0
      _dataSamplesInterval = 0
      _firstMeasureDuration = 0
      _columnNames = New List(Of String)()
      _minVal = 0
      _avgVal = 0
      _maxVal = 0
      _caltyp = 0
      _calpar = New List(Of Integer)()
      _calraw = New List(Of Double)()
      _calref = New List(Of Double)()
      _values = New List(Of List(Of Double))()
      REM --- (end of generated code: YDataStream attributes initialization)
    End Sub

    Public Sub New(parent As YFunction, dataset As YDataSet, encoded As List(Of Integer))
      _parent = parent
      REM --- (generated code: YDataStream attributes initialization)
      _runNo = 0
      _utcStamp = 0
      _nCols = 0
      _nRows = 0
      _startTime = 0
      _duration = 0
      _dataSamplesInterval = 0
      _firstMeasureDuration = 0
      _columnNames = New List(Of String)()
      _minVal = 0
      _avgVal = 0
      _maxVal = 0
      _caltyp = 0
      _calpar = New List(Of Integer)()
      _calraw = New List(Of Double)()
      _calref = New List(Of Double)()
      _values = New List(Of List(Of Double))()
      REM --- (end of generated code: YDataStream attributes initialization)
      _initFromDataSet(dataset, encoded)
    End Sub

    REM --- (generated code: YDataStream private methods declaration)

    REM --- (end of generated code: YDataStream private methods declaration)

    REM --- (generated code: YDataStream public methods declaration)
    Public Overridable Function _initFromDataSet(dataset As YDataSet, encoded As List(Of Integer)) As Integer
      Dim val As Integer = 0
      Dim i As Integer = 0
      Dim maxpos As Integer = 0
      Dim ms_offset As Integer = 0
      Dim samplesPerHour As Integer = 0
      Dim fRaw As Double = 0
      Dim fRef As Double = 0
      Dim iCalib As List(Of Integer) = New List(Of Integer)()
      REM // decode sequence header to extract data
      Me._runNo = encoded(0) + (((encoded(1)) << 16))
      Me._utcStamp = encoded(2) + (((encoded(3)) << 16))
      val = encoded(4)
      Me._isAvg = (((val) And (&H100)) = 0)
      samplesPerHour = ((val) And (&Hff))
      If (((val) And (&H100)) <> 0) Then
        samplesPerHour = samplesPerHour * 3600
      Else
        If (((val) And (&H200)) <> 0) Then
          samplesPerHour = samplesPerHour * 60
        End If
      End If
      Me._dataSamplesInterval = 3600.0 / samplesPerHour
      ms_offset = encoded(6)
      If (ms_offset < 1000) Then
        REM // new encoding . add the ms to the UTC timestamp
        Me._startTime = Me._utcStamp + (ms_offset / 1000.0)
      Else
        REM // legacy encoding subtract the measure interval form the UTC timestamp
        Me._startTime = Me._utcStamp - Me._dataSamplesInterval
      End If
      Me._firstMeasureDuration = encoded(5)
      If (Not (Me._isAvg)) Then
        Me._firstMeasureDuration = Me._firstMeasureDuration / 1000.0
      End If
      val = encoded(7)
      Me._isClosed = (val <> &Hffff)
      If (val = &Hffff) Then
        val = 0
      End If
      Me._nRows = val
      If (Me._nRows > 0) Then
        If (Me._firstMeasureDuration > 0) Then
          Me._duration = Me._firstMeasureDuration + (Me._nRows - 1) * Me._dataSamplesInterval
        Else
          Me._duration = Me._nRows * Me._dataSamplesInterval
        End If
      Else
        Me._duration = 0
      End If
      REM // precompute decoding parameters
      iCalib = dataset._get_calibration()
      Me._caltyp = iCalib(0)
      If (Me._caltyp <> 0) Then
        Me._calhdl = YAPI._getCalibrationHandler(Me._caltyp)
        maxpos = iCalib.Count
        Me._calpar.Clear()
        Me._calraw.Clear()
        Me._calref.Clear()
        i = 1
        While (i < maxpos)
          Me._calpar.Add(iCalib(i))
          i = i + 1
        End While
        i = 1
        While (i + 1 < maxpos)
          fRaw = iCalib(i)
          fRaw = fRaw / 1000.0
          fRef = iCalib(i + 1)
          fRef = fRef / 1000.0
          Me._calraw.Add(fRaw)
          Me._calref.Add(fRef)
          i = i + 2
        End While
      End If
      REM // preload column names for backward-compatibility
      Me._functionId = dataset.get_functionId()
      If (Me._isAvg) Then
        Me._columnNames.Clear()
        Me._columnNames.Add("" + Me._functionId + "_min")
        Me._columnNames.Add("" + Me._functionId + "_avg")
        Me._columnNames.Add("" + Me._functionId + "_max")
        Me._nCols = 3
      Else
        Me._columnNames.Clear()
        Me._columnNames.Add(Me._functionId)
        Me._nCols = 1
      End If
      REM // decode min/avg/max values for the sequence
      If (Me._nRows > 0) Then
        Me._avgVal = Me._decodeAvg(encoded(8) + ((((encoded(9)) Xor (&H8000)) << 16)), 1)
        Me._minVal = Me._decodeVal(encoded(10) + (((encoded(11)) << 16)))
        Me._maxVal = Me._decodeVal(encoded(12) + (((encoded(13)) << 16)))
      End If
      Return 0
    End Function

    Public Overridable Function _parseStream(sdata As Byte()) As Integer
      Dim idx As Integer = 0
      Dim udat As List(Of Integer) = New List(Of Integer)()
      Dim dat As List(Of Double) = New List(Of Double)()
      If (Me._isLoaded AndAlso Not (Me._isClosed)) Then
        Return YAPI.SUCCESS
      End If
      If ((sdata).Length = 0) Then
        Me._nRows = 0
        Return YAPI.SUCCESS
      End If

      udat = YAPI._decodeWords(Me._parent._json_get_string(sdata))
      Me._values.Clear()
      idx = 0
      If (Me._isAvg) Then
        While (idx + 3 < udat.Count)
          dat.Clear()
          If ((udat(idx) = 65535) AndAlso (udat(idx + 1) = 65535)) Then
            dat.Add(Double.NaN)
            dat.Add(Double.NaN)
            dat.Add(Double.NaN)
          Else
            dat.Add(Me._decodeVal(udat(idx + 2) + (((udat(idx + 3)) << 16))))
            dat.Add(Me._decodeAvg(udat(idx) + ((((udat(idx + 1)) Xor (&H8000)) << 16)), 1))
            dat.Add(Me._decodeVal(udat(idx + 4) + (((udat(idx + 5)) << 16))))
          End If
          idx = idx + 6
          Me._values.Add(New List(Of Double)(dat))
        End While
      Else
        While (idx + 1 < udat.Count)
          dat.Clear()
          If ((udat(idx) = 65535) AndAlso (udat(idx + 1) = 65535)) Then
            dat.Add(Double.NaN)
          Else
            dat.Add(Me._decodeAvg(udat(idx) + ((((udat(idx + 1)) Xor (&H8000)) << 16)), 1))
          End If
          Me._values.Add(New List(Of Double)(dat))
          idx = idx + 2
        End While
      End If

      Me._nRows = Me._values.Count
      Me._isLoaded = True
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function _wasLoaded() As Boolean
      Return Me._isLoaded
    End Function

    Public Overridable Function _get_url() As String
      Dim url As String
      url = "logger.json?id=" + Me._functionId + "&run=" + Convert.ToString(Me._runNo) + "&utc=" + Convert.ToString(Me._utcStamp)
      Return url
    End Function

    Public Overridable Function _get_baseurl() As String
      Dim url As String
      url = "logger.json?id=" + Me._functionId + "&run=" + Convert.ToString(Me._runNo) + "&utc="
      Return url
    End Function

    Public Overridable Function _get_urlsuffix() As String
      Dim url As String
      url = "" + Convert.ToString(Me._utcStamp)
      Return url
    End Function

    Public Overridable Function loadStream() As Integer
      Return Me._parseStream(Me._parent._download(Me._get_url()))
    End Function

    Public Overridable Function _decodeVal(w As Integer) As Double
      Dim val As Double = 0
      val = w
      val = val / 1000.0
      If (Me._caltyp <> 0) Then
        If (Not (Me._calhdl Is Nothing)) Then
          val = Me._calhdl(val, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
        End If
      End If
      Return val
    End Function

    Public Overridable Function _decodeAvg(dw As Integer, count As Integer) As Double
      Dim val As Double = 0
      val = dw
      val = val / 1000.0
      If (Me._caltyp <> 0) Then
        If (Not (Me._calhdl Is Nothing)) Then
          val = Me._calhdl(val, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
        End If
      End If
      Return val
    End Function

    Public Overridable Function isClosed() As Boolean
      Return Me._isClosed
    End Function

    '''*
    ''' <summary>
    '''   Returns the run index of the data stream.
    ''' <para>
    '''   A run can be made of
    '''   multiple datastreams, for different time intervals.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the run index.
    ''' </returns>
    '''/
    Public Overridable Function get_runIndex() As Integer
      Return Me._runNo
    End Function

    '''*
    ''' <summary>
    '''   Returns the relative start time of the data stream, measured in seconds.
    ''' <para>
    '''   For recent firmwares, the value is relative to the present time,
    '''   which means the value is always negative.
    '''   If the device uses a firmware older than version 13000, value is
    '''   relative to the start of the time the device was powered on, and
    '''   is always positive.
    '''   If you need an absolute UTC timestamp, use <c>get_realStartTimeUTC()</c>.
    ''' </para>
    ''' <para>
    '''   <b>DEPRECATED</b>: This method has been replaced by <c>get_realStartTimeUTC()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of seconds
    '''   between the start of the run and the beginning of this data
    '''   stream.
    ''' </returns>
    '''/
    Public Overridable Function get_startTime() As Integer
      Return CInt(Me._utcStamp - CInt((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds))
    End Function

    '''*
    ''' <summary>
    '''   Returns the start time of the data stream, relative to the Jan 1, 1970.
    ''' <para>
    '''   If the UTC time was not set in the datalogger at the time of the recording
    '''   of this data stream, this method returns 0.
    ''' </para>
    ''' <para>
    '''   <b>DEPRECATED</b>: This method has been replaced by <c>get_realStartTimeUTC()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of seconds
    '''   between the Jan 1, 1970 and the beginning of this data
    '''   stream (i.e. Unix time representation of the absolute time).
    ''' </returns>
    '''/
    Public Overridable Function get_startTimeUTC() As Long
      Return CType(Math.Round(Me._startTime), Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns the start time of the data stream, relative to the Jan 1, 1970.
    ''' <para>
    '''   If the UTC time was not set in the datalogger at the time of the recording
    '''   of this data stream, this method returns 0.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number  corresponding to the number of seconds
    '''   between the Jan 1, 1970 and the beginning of this data
    '''   stream (i.e. Unix time representation of the absolute time).
    ''' </returns>
    '''/
    Public Overridable Function get_realStartTimeUTC() As Double
      Return Me._startTime
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of milliseconds between two consecutive
    '''   rows of this data stream.
    ''' <para>
    '''   By default, the data logger records one row
    '''   per second, but the recording frequency can be changed for
    '''   each device function
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to a number of milliseconds.
    ''' </returns>
    '''/
    Public Overridable Function get_dataSamplesIntervalMs() As Integer
      Return CType(Math.Round(Me._dataSamplesInterval*1000), Integer)
    End Function

    Public Overridable Function get_dataSamplesInterval() As Double
      Return Me._dataSamplesInterval
    End Function

    Public Overridable Function get_firstDataSamplesInterval() As Double
      Return Me._firstMeasureDuration
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of data rows present in this stream.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   If the device uses a firmware older than version 13000,
    '''   this method fetches the whole data stream from the device
    '''   if not yet done, which can cause a little delay.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of rows.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns zero.
    ''' </para>
    '''/
    Public Overridable Function get_rowCount() As Integer
      If ((Me._nRows <> 0) AndAlso Me._isClosed) Then
        Return Me._nRows
      End If
      Me.loadStream()
      Return Me._nRows
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of data columns present in this stream.
    ''' <para>
    '''   The meaning of the values present in each column can be obtained
    '''   using the method <c>get_columnNames()</c>.
    ''' </para>
    ''' <para>
    '''   If the device uses a firmware older than version 13000,
    '''   this method fetches the whole data stream from the device
    '''   if not yet done, which can cause a little delay.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of columns.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns zero.
    ''' </para>
    '''/
    Public Overridable Function get_columnCount() As Integer
      If (Me._nCols <> 0) Then
        Return Me._nCols
      End If
      Me.loadStream()
      Return Me._nCols
    End Function

    '''*
    ''' <summary>
    '''   Returns the title (or meaning) of each data column present in this stream.
    ''' <para>
    '''   In most case, the title of the data column is the hardware identifier
    '''   of the sensor that produced the data. For streams recorded at a lower
    '''   recording rate, the dataLogger stores the min, average and max value
    '''   during each measure interval into three columns with suffixes _min,
    '''   _avg and _max respectively.
    ''' </para>
    ''' <para>
    '''   If the device uses a firmware older than version 13000,
    '''   this method fetches the whole data stream from the device
    '''   if not yet done, which can cause a little delay.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list containing as many strings as there are columns in the
    '''   data stream.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_columnNames() As List(Of String)
      If (Me._columnNames.Count <> 0) Then
        Return Me._columnNames
      End If
      Me.loadStream()
      Return Me._columnNames
    End Function

    '''*
    ''' <summary>
    '''   Returns the smallest measure observed within this stream.
    ''' <para>
    '''   If the device uses a firmware older than version 13000,
    '''   this method will always return YDataStream.DATA_INVALID.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the smallest value,
    '''   or YDataStream.DATA_INVALID if the stream is not yet complete (still recording).
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns YDataStream.DATA_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_minValue() As Double
      Return Me._minVal
    End Function

    '''*
    ''' <summary>
    '''   Returns the average of all measures observed within this stream.
    ''' <para>
    '''   If the device uses a firmware older than version 13000,
    '''   this method will always return YDataStream.DATA_INVALID.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the average value,
    '''   or YDataStream.DATA_INVALID if the stream is not yet complete (still recording).
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns YDataStream.DATA_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_averageValue() As Double
      Return Me._avgVal
    End Function

    '''*
    ''' <summary>
    '''   Returns the largest measure observed within this stream.
    ''' <para>
    '''   If the device uses a firmware older than version 13000,
    '''   this method will always return YDataStream.DATA_INVALID.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the largest value,
    '''   or YDataStream.DATA_INVALID if the stream is not yet complete (still recording).
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns YDataStream.DATA_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_maxValue() As Double
      Return Me._maxVal
    End Function

    Public Overridable Function get_realDuration() As Double
      If (Me._isClosed) Then
        Return Me._duration
      End If
      Return CType(CInt((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds) - Me._utcStamp, Double)
    End Function

    '''*
    ''' <summary>
    '''   Returns the whole data set contained in the stream, as a bidimensional
    '''   table of numbers.
    ''' <para>
    '''   The meaning of the values present in each column can be obtained
    '''   using the method <c>get_columnNames()</c>.
    ''' </para>
    ''' <para>
    '''   This method fetches the whole data stream from the device,
    '''   if not yet done.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list containing as many elements as there are rows in the
    '''   data stream. Each row itself is a list of floating-point
    '''   numbers.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_dataRows() As List(Of List(Of Double))
      If ((Me._values.Count = 0) OrElse Not (Me._isClosed)) Then
        Me.loadStream()
      End If
      Return Me._values
    End Function

    '''*
    ''' <summary>
    '''   Returns a single measure from the data stream, specified by its
    '''   row and column index.
    ''' <para>
    '''   The meaning of the values present in each column can be obtained
    '''   using the method get_columnNames().
    ''' </para>
    ''' <para>
    '''   This method fetches the whole data stream from the device,
    '''   if not yet done.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="row">
    '''   row index
    ''' </param>
    ''' <param name="col">
    '''   column index
    ''' </param>
    ''' <returns>
    '''   a floating-point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns YDataStream.DATA_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_data(row As Integer, col As Integer) As Double
      If ((Me._values.Count = 0) OrElse Not (Me._isClosed)) Then
        Me.loadStream()
      End If
      If (row >= Me._values.Count) Then
        Return DATA_INVALID
      End If
      If (col >= Me._values(row).Count) Then
        Return DATA_INVALID
      End If
      Return Me._values(row)(col)
    End Function



    REM --- (end of generated code: YDataStream public methods declaration)
  End Class

  REM --- (generated code: YDataStream functions)


  REM --- (end of generated code: YDataStream functions)


  REM --- (generated code: YMeasure class start)

  '''*
  ''' <c>YMeasure</c> objects are used within the API to represent
  ''' a value measured at a specified time. These objects are
  ''' used in particular in conjunction with the <c>YDataSet</c> class,
  ''' but also for sensors periodic timed reports
  ''' (see <c>sensor.registerTimedReportCallback</c>).
  '''/
  Public Class YMeasure
    REM --- (end of generated code: YMeasure class start)

    REM --- (generated code: YMeasure definitions)
    REM --- (end of generated code: YMeasure definitions)

    REM --- (generated code: YMeasure attributes declaration)
    Protected _start As Double
    Protected _end As Double
    Protected _minVal As Double
    Protected _avgVal As Double
    Protected _maxVal As Double
    REM --- (end of generated code: YMeasure attributes declaration)
    Protected _start_datetime As DateTime
    Protected _end_datetime As DateTime

    Sub New(start As Double, endt As Double, minVal As Double, avgVal As Double, maxVal As Double)
      Dim epoch As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
      REM --- (generated code: YMeasure attributes initialization)
      _start = 0
      _end = 0
      _minVal = 0
      _avgVal = 0
      _maxVal = 0
      REM --- (end of generated code: YMeasure attributes initialization)
      _start = start
      _end = endt
      _minVal = minVal
      _avgVal = avgVal
      _maxVal = maxVal
      _start_datetime = epoch.AddSeconds(_start)
      _end_datetime = epoch.AddSeconds(_end)
    End Sub

    REM --- (generated code: YMeasure private methods declaration)

    REM --- (end of generated code: YMeasure private methods declaration)

    REM --- (generated code: YMeasure public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the start time of the measure, relative to the Jan 1, 1970 UTC
    '''   (Unix timestamp).
    ''' <para>
    '''   When the recording rate is higher then 1 sample
    '''   per second, the timestamp may have a fractional part.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the number of seconds
    '''   between the Jan 1, 1970 UTC and the beginning of this measure.
    ''' </returns>
    '''/
    Public Overridable Function get_startTimeUTC() As Double
      Return Me._start
    End Function

    '''*
    ''' <summary>
    '''   Returns the end time of the measure, relative to the Jan 1, 1970 UTC
    '''   (Unix timestamp).
    ''' <para>
    '''   When the recording rate is higher than 1 sample
    '''   per second, the timestamp may have a fractional part.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the number of seconds
    '''   between the Jan 1, 1970 UTC and the end of this measure.
    ''' </returns>
    '''/
    Public Overridable Function get_endTimeUTC() As Double
      Return Me._end
    End Function

    '''*
    ''' <summary>
    '''   Returns the smallest value observed during the time interval
    '''   covered by this measure.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the smallest value observed.
    ''' </returns>
    '''/
    Public Overridable Function get_minValue() As Double
      Return Me._minVal
    End Function

    '''*
    ''' <summary>
    '''   Returns the average value observed during the time interval
    '''   covered by this measure.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the average value observed.
    ''' </returns>
    '''/
    Public Overridable Function get_averageValue() As Double
      Return Me._avgVal
    End Function

    '''*
    ''' <summary>
    '''   Returns the largest value observed during the time interval
    '''   covered by this measure.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the largest value observed.
    ''' </returns>
    '''/
    Public Overridable Function get_maxValue() As Double
      Return Me._maxVal
    End Function



    REM --- (end of generated code: YMeasure public methods declaration)

    Function get_startTimeUTC_asDateTime() As DateTime
      Return _start_datetime
    End Function

    Function get_endTimeUTC_asDateTime() As DateTime
      Return _end_datetime
    End Function
  End Class

  REM --- (generated code: YMeasure functions)


  REM --- (end of generated code: YMeasure functions)


  REM --- (generated code: YDataSet class start)

  '''*
  ''' <c>YDataSet</c> objects make it possible to retrieve a set of recorded measures
  ''' for a given sensor and a specified time interval. They can be used
  ''' to load data points with a progress report. When the <c>YDataSet</c> object is
  ''' instantiated by the <c>sensor.get_recordedData()</c>  function, no data is
  ''' yet loaded from the module. It is only when the <c>loadMore()</c>
  ''' method is called over and over than data will be effectively loaded
  ''' from the dataLogger.
  ''' <para>
  '''   A preview of available measures is available using the function
  '''   <c>get_preview()</c> as soon as <c>loadMore()</c> has been called
  '''   once. Measures themselves are available using function <c>get_measures()</c>
  '''   when loaded by subsequent calls to <c>loadMore()</c>.
  ''' </para>
  ''' <para>
  '''   This class can only be used on devices that use a relatively recent firmware,
  '''   as <c>YDataSet</c> objects are not supported by firmwares older than version 13000.
  ''' </para>
  '''/
  Public Class YDataSet
    REM --- (end of generated code: YDataSet class start)
    REM --- (generated code: YDataSet definitions)
    REM --- (end of generated code: YDataSet definitions)

    REM --- (generated code: YDataSet attributes declaration)
    Protected _parent As YFunction
    Protected _hardwareId As String
    Protected _functionId As String
    Protected _unit As String
    Protected _bulkLoad As Integer
    Protected _startTimeMs As Double
    Protected _endTimeMs As Double
    Protected _progress As Integer
    Protected _calib As List(Of Integer)
    Protected _streams As List(Of YDataStream)
    Protected _summary As YMeasure
    Protected _preview As List(Of YMeasure)
    Protected _measures As List(Of YMeasure)
    Protected _summaryMinVal As Double
    Protected _summaryMaxVal As Double
    Protected _summaryTotalAvg As Double
    Protected _summaryTotalTime As Double
    REM --- (end of generated code: YDataSet attributes declaration)


    Sub New(parent As YFunction, functionId As String, unit As String, startTime As Double, endTime As Double)
      REM --- (generated code: YDataSet attributes initialization)
      _bulkLoad = 0
      _startTimeMs = 0
      _endTimeMs = 0
      _progress = 0
      _calib = New List(Of Integer)()
      _streams = New List(Of YDataStream)()
      _preview = New List(Of YMeasure)()
      _measures = New List(Of YMeasure)()
      _summaryMinVal = 0
      _summaryMaxVal = 0
      _summaryTotalAvg = 0
      _summaryTotalTime = 0
      REM --- (end of generated code: YDataSet attributes initialization)
      _parent = parent
      _functionId = functionId
      _unit = unit
      _startTimeMs = startTime * 1000
      _endTimeMs = endTime * 1000
      _summary = New YMeasure(0, 0, 0, 0, 0)
      _progress = -1
    End Sub


    Sub New(parent As YFunction)
      REM --- (generated code: YDataSet attributes initialization)
      _bulkLoad = 0
      _startTimeMs = 0
      _endTimeMs = 0
      _progress = 0
      _calib = New List(Of Integer)()
      _streams = New List(Of YDataStream)()
      _preview = New List(Of YMeasure)()
      _measures = New List(Of YMeasure)()
      _summaryMinVal = 0
      _summaryMaxVal = 0
      _summaryTotalAvg = 0
      _summaryTotalTime = 0
      REM --- (end of generated code: YDataSet attributes initialization)
      _parent = parent
      _startTimeMs = 0
      _endTimeMs = 0
      _summary = New YMeasure(0, 0, 0, 0, 0)
    End Sub


    Public Function _parse(data As String) As Integer
      Dim p As New YJSONObject(data)
      Dim arr As YJSONArray

      If Not YAPI.ExceptionsDisabled Then
        p.parse()
      Else
        Try
          p.parse()
        Catch
          Return YAPI.IO_ERROR
        End Try
      End If

      Dim stream As YDataStream
      Dim streamStartTime As Double
      Dim streamEndTime As Double

      Me._functionId = p.getString("id")
      Me._unit = p.getString("unit")
      If p.has("bulk") Then
        Me._bulkLoad = YAPI._atoi(p.getString("bulk"))
      End If
      If p.has("calib") Then
        Me._calib = YAPI._decodeFloats(p.getString("calib"))
        Me._calib(0) = Me._calib(0) \ 1000
      Else
        Me._calib = YAPI._decodeWords(p.getString("cal"))
      End If
      arr = p.getYJSONArray("streams")
      Me._streams = New List(Of YDataStream)()
      Me._preview = New List(Of YMeasure)()
      Me._measures = New List(Of YMeasure)()
      For i As Integer = 0 To arr.Length - 1
        stream = _parent._findDataStream(Me, arr.getString(i))
        streamStartTime = Math.Round(stream.get_realStartTimeUTC() * 1000)
        streamEndTime = streamStartTime + Math.Round(stream.get_realDuration() * 1000)
        If _startTimeMs > 0 AndAlso streamEndTime <= _startTimeMs Then
          REM this stream is too early, drop it
        ElseIf _endTimeMs > 0 AndAlso streamStartTime >= Me._endTimeMs Then
          REM this stream is too late, drop it
        Else
          _streams.Add(stream)
        End If
      Next i
      _progress = 0
      Return get_progress()
    End Function

    REM --- (generated code: YDataSet private methods declaration)

    REM --- (end of generated code: YDataSet private methods declaration)

    REM --- (generated code: YDataSet public methods declaration)
    Public Overridable Function _get_calibration() As List(Of Integer)
      Return Me._calib
    End Function

    Public Overridable Function loadSummary(data As Byte()) As Integer
      Dim dataRows As List(Of List(Of Double)) = New List(Of List(Of Double))()
      Dim tim As Double = 0
      Dim mitv As Double = 0
      Dim itv As Double = 0
      Dim fitv As Double = 0
      Dim end_ As Double = 0
      Dim nCols As Integer = 0
      Dim minCol As Integer = 0
      Dim avgCol As Integer = 0
      Dim maxCol As Integer = 0
      Dim res As Integer = 0
      Dim m_pos As Integer = 0
      Dim previewTotalTime As Double = 0
      Dim previewTotalAvg As Double = 0
      Dim previewMinVal As Double = 0
      Dim previewMaxVal As Double = 0
      Dim previewAvgVal As Double = 0
      Dim previewStartMs As Double = 0
      Dim previewStopMs As Double = 0
      Dim previewDuration As Double = 0
      Dim streamStartTimeMs As Double = 0
      Dim streamDuration As Double = 0
      Dim streamEndTimeMs As Double = 0
      Dim minVal As Double = 0
      Dim avgVal As Double = 0
      Dim maxVal As Double = 0
      Dim summaryStartMs As Double = 0
      Dim summaryStopMs As Double = 0
      Dim summaryTotalTime As Double = 0
      Dim summaryTotalAvg As Double = 0
      Dim summaryMinVal As Double = 0
      Dim summaryMaxVal As Double = 0
      Dim url As String
      Dim strdata As String
      Dim measure_data As List(Of Double) = New List(Of Double)()

      If (Me._progress < 0) Then
        strdata = YAPI.DefaultEncoding.GetString(data)
        If (strdata = "{}") Then
          Me._parent._throw(YAPI.VERSION_MISMATCH, "device firmware is too old")
          Return YAPI.VERSION_MISMATCH
        End If
        res = Me._parse(strdata)
        If (res < 0) Then
          Return res
        End If
      End If
      summaryTotalTime = 0
      summaryTotalAvg = 0
      summaryMinVal = YAPI.MAX_DOUBLE
      summaryMaxVal = YAPI.MIN_DOUBLE
      summaryStartMs = YAPI.MAX_DOUBLE
      summaryStopMs = YAPI.MIN_DOUBLE

      REM // Parse complete streams
      Dim ii_0 As Integer
      For ii_0 = 0 To Me._streams.Count - 1
        streamStartTimeMs = Math.Round(Me._streams(ii_0).get_realStartTimeUTC() * 1000)
        streamDuration = Me._streams(ii_0).get_realDuration()
        streamEndTimeMs = streamStartTimeMs + Math.Round(streamDuration * 1000)
        If ((streamStartTimeMs >= Me._startTimeMs) AndAlso ((Me._endTimeMs = 0) OrElse (streamEndTimeMs <= Me._endTimeMs))) Then
          REM // stream that are completely inside the dataset
          previewMinVal = Me._streams(ii_0).get_minValue()
          previewAvgVal = Me._streams(ii_0).get_averageValue()
          previewMaxVal = Me._streams(ii_0).get_maxValue()
          previewStartMs = streamStartTimeMs
          previewStopMs = streamEndTimeMs
          previewDuration = streamDuration
        Else
          REM // stream that are partially in the dataset
          REM // we need to parse data to filter value outside the dataset
          If (Not (Me._streams(ii_0)._wasLoaded())) Then
            url = Me._streams(ii_0)._get_url()
            data = Me._parent._download(url)
            Me._streams(ii_0)._parseStream(data)
          End If
          dataRows = Me._streams(ii_0).get_dataRows()
          If (dataRows.Count = 0) Then
            Return Me.get_progress()
          End If
          tim = streamStartTimeMs
          fitv = Math.Round(Me._streams(ii_0).get_firstDataSamplesInterval() * 1000)
          itv = Math.Round(Me._streams(ii_0).get_dataSamplesInterval() * 1000)
          nCols = dataRows(0).Count
          minCol = 0
          If (nCols > 2) Then
            avgCol = 1
          Else
            avgCol = 0
          End If
          If (nCols > 2) Then
            maxCol = 2
          Else
            maxCol = 0
          End If
          previewTotalTime = 0
          previewTotalAvg = 0
          previewStartMs = streamEndTimeMs
          previewStopMs = streamStartTimeMs
          previewMinVal = YAPI.MAX_DOUBLE
          previewMaxVal = YAPI.MIN_DOUBLE
          m_pos = 0
          While (m_pos < dataRows.Count)
            measure_data = dataRows(m_pos)
            If (m_pos = 0) Then
              mitv = fitv
            Else
              mitv = itv
            End If
            end_ = tim + mitv
            If ((end_ > Me._startTimeMs) AndAlso ((Me._endTimeMs = 0) OrElse (tim < Me._endTimeMs))) Then
              minVal = measure_data(minCol)
              avgVal = measure_data(avgCol)
              maxVal = measure_data(maxCol)
              If (previewStartMs > tim) Then
                previewStartMs = tim
              End If
              If (previewStopMs < end_) Then
                previewStopMs = end_
              End If
              If (previewMinVal > minVal) Then
                previewMinVal = minVal
              End If
              If (previewMaxVal < maxVal) Then
                previewMaxVal = maxVal
              End If
              If (Not (Double.IsNaN(avgVal))) Then
                previewTotalAvg = previewTotalAvg + (avgVal * mitv)
                previewTotalTime = previewTotalTime + mitv
              End If
            End If
            tim = end_
            m_pos = m_pos + 1
          End While
          If (previewTotalTime > 0) Then
            previewAvgVal = previewTotalAvg / previewTotalTime
            previewDuration = (previewStopMs - previewStartMs) / 1000.0
          Else
            previewAvgVal = 0.0
            previewDuration = 0.0
          End If
        End If
        Me._preview.Add(New YMeasure(previewStartMs / 1000.0, previewStopMs / 1000.0, previewMinVal, previewAvgVal, previewMaxVal))
        If (summaryMinVal > previewMinVal) Then
          summaryMinVal = previewMinVal
        End If
        If (summaryMaxVal < previewMaxVal) Then
          summaryMaxVal = previewMaxVal
        End If
        If (summaryStartMs > previewStartMs) Then
          summaryStartMs = previewStartMs
        End If
        If (summaryStopMs < previewStopMs) Then
          summaryStopMs = previewStopMs
        End If
        summaryTotalAvg = summaryTotalAvg + (previewAvgVal * previewDuration)
        summaryTotalTime = summaryTotalTime + previewDuration
      Next ii_0
      If ((Me._startTimeMs = 0) OrElse (Me._startTimeMs > summaryStartMs)) Then
        Me._startTimeMs = summaryStartMs
      End If
      If ((Me._endTimeMs = 0) OrElse (Me._endTimeMs < summaryStopMs)) Then
        Me._endTimeMs = summaryStopMs
      End If
      If (summaryTotalTime > 0) Then
        Me._summary = New YMeasure(summaryStartMs / 1000.0, summaryStopMs / 1000.0, summaryMinVal, summaryTotalAvg / summaryTotalTime, summaryMaxVal)
      Else
        Me._summary = New YMeasure(0.0, 0.0, YAPI.INVALID_DOUBLE, YAPI.INVALID_DOUBLE, YAPI.INVALID_DOUBLE)
      End If
      Return Me.get_progress()
    End Function

    Public Overridable Function processMore(progress As Integer, data As Byte()) As Integer
      Dim stream As YDataStream
      Dim dataRows As List(Of List(Of Double)) = New List(Of List(Of Double))()
      Dim tim As Double = 0
      Dim itv As Double = 0
      Dim fitv As Double = 0
      Dim avgv As Double = 0
      Dim end_ As Double = 0
      Dim nCols As Integer = 0
      Dim minCol As Integer = 0
      Dim avgCol As Integer = 0
      Dim maxCol As Integer = 0
      Dim firstMeasure As Boolean
      Dim baseurl As String
      Dim url As String
      Dim suffix As String
      Dim suffixes As List(Of String) = New List(Of String)()
      Dim idx As Integer = 0
      Dim bulkFile As Byte() = New Byte(){}
      Dim urlIdx As Integer = 0
      Dim streamBin As List(Of Byte()) = New List(Of Byte())()

      If (progress <> Me._progress) Then
        Return Me._progress
      End If
      If (Me._progress < 0) Then
        Return Me.loadSummary(data)
      End If
      stream = Me._streams(Me._progress)
      If (Not (stream._wasLoaded())) Then
        stream._parseStream(data)
      End If
      dataRows = stream.get_dataRows()
      Me._progress = Me._progress + 1
      If (dataRows.Count = 0) Then
        Return Me.get_progress()
      End If
      tim = Math.Round(stream.get_realStartTimeUTC() * 1000)
      fitv = Math.Round(stream.get_firstDataSamplesInterval() * 1000)
      itv = Math.Round(stream.get_dataSamplesInterval() * 1000)
      If (fitv = 0) Then
        fitv = itv
      End If
      If (tim < itv) Then
        tim = itv
      End If
      nCols = dataRows(0).Count
      minCol = 0
      If (nCols > 2) Then
        avgCol = 1
      Else
        avgCol = 0
      End If
      If (nCols > 2) Then
        maxCol = 2
      Else
        maxCol = 0
      End If

      firstMeasure = True
      Dim ii_0 As Integer
      For ii_0 = 0 To dataRows.Count - 1
        If (firstMeasure) Then
          end_ = tim + fitv
          firstMeasure = False
        Else
          end_ = tim + itv
        End If
        avgv = dataRows(ii_0)(avgCol)
        If ((end_ > Me._startTimeMs) AndAlso ((Me._endTimeMs = 0) OrElse (tim < Me._endTimeMs)) AndAlso Not (Double.IsNaN(avgv))) Then
          Me._measures.Add(New YMeasure(tim / 1000, end_ / 1000, dataRows(ii_0)(minCol), avgv, dataRows(ii_0)(maxCol)))
        End If
        tim = end_
      Next ii_0

      REM // Perform bulk preload to speed-up network transfer
      If ((Me._bulkLoad > 0) AndAlso (Me._progress < Me._streams.Count)) Then
        stream = Me._streams(Me._progress)
        If (stream._wasLoaded()) Then
          Return Me.get_progress()
        End If
        baseurl = stream._get_baseurl()
        url = stream._get_url()
        suffix = stream._get_urlsuffix()
        suffixes.Add(suffix)
        idx = Me._progress + 1
        While ((idx < Me._streams.Count) AndAlso (suffixes.Count < Me._bulkLoad))
          stream = Me._streams(idx)
          If (Not (stream._wasLoaded()) AndAlso (stream._get_baseurl() = baseurl)) Then
            suffix = stream._get_urlsuffix()
            suffixes.Add(suffix)
            url = url + "," + suffix
          End If
          idx = idx + 1
        End While
        bulkFile = Me._parent._download(url)
        streamBin = Me._parent._json_get_array(bulkFile)
        urlIdx = 0
        idx = Me._progress
        While ((idx < Me._streams.Count) AndAlso (urlIdx < suffixes.Count) AndAlso (urlIdx < streamBin.Count))
          stream = Me._streams(idx)
          If ((stream._get_baseurl() = baseurl) AndAlso (stream._get_urlsuffix() = suffixes(urlIdx))) Then
            stream._parseStream(streamBin(urlIdx))
            urlIdx = urlIdx + 1
          End If
          idx = idx + 1
        End While
      End If
      Return Me.get_progress()
    End Function

    Public Overridable Function get_privateDataStreams() As List(Of YDataStream)
      Return Me._streams
    End Function

    '''*
    ''' <summary>
    '''   Returns the unique hardware identifier of the function who performed the measures,
    '''   in the form <c>SERIAL.FUNCTIONID</c>.
    ''' <para>
    '''   The unique hardware identifier is composed of the
    '''   device serial number and of the hardware identifier of the function
    '''   (for example <c>THRMCPL1-123456.temperature1</c>)
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string that uniquely identifies the function (ex: <c>THRMCPL1-123456.temperature1</c>)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns  <c>YDataSet.HARDWAREID_INVALID</c>.
    ''' </para>
    '''/
    Public Overridable Function get_hardwareId() As String
      Dim mo As YModule
      If (Not (Me._hardwareId = "")) Then
        Return Me._hardwareId
      End If
      mo = Me._parent.get_module()
      Me._hardwareId = "" + mo.get_serialNumber() + "." + Me.get_functionId()
      Return Me._hardwareId
    End Function

    '''*
    ''' <summary>
    '''   Returns the hardware identifier of the function that performed the measure,
    '''   without reference to the module.
    ''' <para>
    '''   For example <c>temperature1</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string that identifies the function (ex: <c>temperature1</c>)
    ''' </returns>
    '''/
    Public Overridable Function get_functionId() As String
      Return Me._functionId
    End Function

    '''*
    ''' <summary>
    '''   Returns the measuring unit for the measured value.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string that represents a physical unit.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns  <c>YDataSet.UNIT_INVALID</c>.
    ''' </para>
    '''/
    Public Overridable Function get_unit() As String
      Return Me._unit
    End Function

    '''*
    ''' <summary>
    '''   Returns the start time of the dataset, relative to the Jan 1, 1970.
    ''' <para>
    '''   When the <c>YDataSet</c> object is created, the start time is the value passed
    '''   in parameter to the <c>get_dataSet()</c> function. After the
    '''   very first call to <c>loadMore()</c>, the start time is updated
    '''   to reflect the timestamp of the first measure actually found in the
    '''   dataLogger within the specified range.
    ''' </para>
    ''' <para>
    '''   <b>DEPRECATED</b>: This method has been replaced by <c>get_summary()</c>
    '''   which contain more precise informations.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of seconds
    '''   between the Jan 1, 1970 and the beginning of this data
    '''   set (i.e. Unix time representation of the absolute time).
    ''' </returns>
    '''/
    Public Overridable Function get_startTimeUTC() As Long
      Return Me.imm_get_startTimeUTC()
    End Function

    Public Overridable Function imm_get_startTimeUTC() As Long
      Return CType((Me._startTimeMs / 1000.0), Long)
    End Function

    '''*
    ''' <summary>
    '''   Returns the end time of the dataset, relative to the Jan 1, 1970.
    ''' <para>
    '''   When the <c>YDataSet</c> object is created, the end time is the value passed
    '''   in parameter to the <c>get_dataSet()</c> function. After the
    '''   very first call to <c>loadMore()</c>, the end time is updated
    '''   to reflect the timestamp of the last measure actually found in the
    '''   dataLogger within the specified range.
    ''' </para>
    ''' <para>
    '''   <b>DEPRECATED</b>: This method has been replaced by <c>get_summary()</c>
    '''   which contain more precise informations.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of seconds
    '''   between the Jan 1, 1970 and the end of this data
    '''   set (i.e. Unix time representation of the absolute time).
    ''' </returns>
    '''/
    Public Overridable Function get_endTimeUTC() As Long
      Return Me.imm_get_endTimeUTC()
    End Function

    Public Overridable Function imm_get_endTimeUTC() As Long
      Return CType(Math.Round(Me._endTimeMs / 1000.0), Long)
    End Function

    '''*
    ''' <summary>
    '''   Returns the progress of the downloads of the measures from the data logger,
    '''   on a scale from 0 to 100.
    ''' <para>
    '''   When the object is instantiated by <c>get_dataSet</c>,
    '''   the progress is zero. Each time <c>loadMore()</c> is invoked, the progress
    '''   is updated, to reach the value 100 only once all measures have been loaded.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer in the range 0 to 100 (percentage of completion).
    ''' </returns>
    '''/
    Public Overridable Function get_progress() As Integer
      If (Me._progress < 0) Then
        Return 0
      End If
      REM // index not yet loaded
      If (Me._progress >= Me._streams.Count) Then
        Return 100
      End If
      Return (1 + (1 + Me._progress) * 98 \ (1 + Me._streams.Count))
    End Function

    '''*
    ''' <summary>
    '''   Loads the next block of measures from the dataLogger, and updates
    '''   the progress indicator.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer in the range 0 to 100 (percentage of completion),
    '''   or a negative error code in case of failure.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadMore() As Integer
      Dim url As String
      Dim stream As YDataStream
      If (Me._progress < 0) Then
        url = "logger.json?id=" + Me._functionId
        If (Me._startTimeMs <> 0) Then
          url = "" + url + "&from=" + Convert.ToString(Me.imm_get_startTimeUTC())
        End If
        If (Me._endTimeMs <> 0) Then
          url = "" + url + "&to=" + Convert.ToString(Me.imm_get_endTimeUTC() + 1)
        End If
      Else
        If (Me._progress >= Me._streams.Count) Then
          Return 100
        Else
          stream = Me._streams(Me._progress)
          If (stream._wasLoaded()) Then
            REM // Do not reload stream if it was already loaded
            Return Me.processMore(Me._progress, YAPI.DefaultEncoding.GetBytes(""))
          End If
          url = stream._get_url()
        End If
      End If
      Try
          Return Me.processMore(Me._progress, Me._parent._download(url))
      Catch
          Return Me.processMore(Me._progress, Me._parent._download(url))
      End Try
    End Function

    '''*
    ''' <summary>
    '''   Returns an <c>YMeasure</c> object which summarizes the whole
    '''   <c>YDataSet</c>.
    ''' <para>
    '''   In includes the following information:
    '''   - the start of a time interval
    '''   - the end of a time interval
    '''   - the minimal value observed during the time interval
    '''   - the average value observed during the time interval
    '''   - the maximal value observed during the time interval
    ''' </para>
    ''' <para>
    '''   This summary is available as soon as <c>loadMore()</c> has
    '''   been called for the first time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an <c>YMeasure</c> object
    ''' </returns>
    '''/
    Public Overridable Function get_summary() As YMeasure
      Return Me._summary
    End Function

    '''*
    ''' <summary>
    '''   Returns a condensed version of the measures that can
    '''   retrieved in this <c>YDataSet</c>, as a list of <c>YMeasure</c>
    '''   objects.
    ''' <para>
    '''   Each item includes:
    '''   - the start of a time interval
    '''   - the end of a time interval
    '''   - the minimal value observed during the time interval
    '''   - the average value observed during the time interval
    '''   - the maximal value observed during the time interval
    ''' </para>
    ''' <para>
    '''   This preview is available as soon as <c>loadMore()</c> has
    '''   been called for the first time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a table of records, where each record depicts the
    '''   measured values during a time interval
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_preview() As List(Of YMeasure)
      Return Me._preview
    End Function

    '''*
    ''' <summary>
    '''   Returns the detailed set of measures for the time interval corresponding
    '''   to a given condensed measures previously returned by <c>get_preview()</c>.
    ''' <para>
    '''   The result is provided as a list of <c>YMeasure</c> objects.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="measure">
    '''   condensed measure from the list previously returned by
    '''   <c>get_preview()</c>.
    ''' </param>
    ''' <returns>
    '''   a table of records, where each record depicts the
    '''   measured values during a time interval
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_measuresAt(measure As YMeasure) As List(Of YMeasure)
      Dim startUtcMs As Double = 0
      Dim stream As YDataStream
      Dim dataRows As List(Of List(Of Double)) = New List(Of List(Of Double))()
      Dim measures As List(Of YMeasure) = New List(Of YMeasure)()
      Dim tim As Double = 0
      Dim itv As Double = 0
      Dim end_ As Double = 0
      Dim nCols As Integer = 0
      Dim minCol As Integer = 0
      Dim avgCol As Integer = 0
      Dim maxCol As Integer = 0

      startUtcMs = measure.get_startTimeUTC() * 1000
      stream = Nothing
      Dim ii_0 As Integer
      For ii_0 = 0 To Me._streams.Count - 1
        If (Math.Round(Me._streams(ii_0).get_realStartTimeUTC() *1000) = startUtcMs) Then
          stream = Me._streams(ii_0)
        End If
      Next ii_0
      If ((stream Is Nothing)) Then
        Return measures
      End If
      dataRows = stream.get_dataRows()
      If (dataRows.Count = 0) Then
        Return measures
      End If
      tim = Math.Round(stream.get_realStartTimeUTC() * 1000)
      itv = Math.Round(stream.get_dataSamplesInterval() * 1000)
      If (tim < itv) Then
        tim = itv
      End If
      nCols = dataRows(0).Count
      minCol = 0
      If (nCols > 2) Then
        avgCol = 1
      Else
        avgCol = 0
      End If
      If (nCols > 2) Then
        maxCol = 2
      Else
        maxCol = 0
      End If

      Dim ii_1 As Integer
      For ii_1 = 0 To dataRows.Count - 1
        end_ = tim + itv
        If ((end_ > Me._startTimeMs) AndAlso ((Me._endTimeMs = 0) OrElse (tim < Me._endTimeMs))) Then
          measures.Add(New YMeasure(tim / 1000.0, end_ / 1000.0, dataRows(ii_1)(minCol), dataRows(ii_1)(avgCol), dataRows(ii_1)(maxCol)))
        End If
        tim = end_
      Next ii_1

      Return measures
    End Function

    '''*
    ''' <summary>
    '''   Returns all measured values currently available for this DataSet,
    '''   as a list of <c>YMeasure</c> objects.
    ''' <para>
    '''   Each item includes:
    '''   - the start of the measure time interval
    '''   - the end of the measure time interval
    '''   - the minimal value observed during the time interval
    '''   - the average value observed during the time interval
    '''   - the maximal value observed during the time interval
    ''' </para>
    ''' <para>
    '''   Before calling this method, you should call <c>loadMore()</c>
    '''   to load data from the device. You may have to call loadMore()
    '''   several time until all rows are loaded, but you can start
    '''   looking at available data rows before the load is complete.
    ''' </para>
    ''' <para>
    '''   The oldest measures are always loaded first, and the most
    '''   recent measures will be loaded last. As a result, timestamps
    '''   are normally sorted in ascending order within the measure table,
    '''   unless there was an unexpected adjustment of the datalogger UTC
    '''   clock.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a table of records, where each record depicts the
    '''   measured value for a given time interval
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function get_measures() As List(Of YMeasure)
      Return Me._measures
    End Function



    REM --- (end of generated code: YDataSet public methods declaration)
  End Class

  REM --- (generated code: YDataSet functions)


  REM --- (end of generated code: YDataSet functions)

  REM --- (generated code: YConsolidatedDataSet class start)

  '''*
  ''' <c>YConsolidatedDataSet</c> objects make it possible to retrieve a set of
  ''' recorded measures from multiple sensors, for a specified time interval.
  ''' They can be used to load data points progressively, and to receive
  ''' data records by timestamp, one by one..
  '''/
  Public Class YConsolidatedDataSet
    REM --- (end of generated code: YConsolidatedDataSet class start)
    REM --- (generated code: YConsolidatedDataSet definitions)
    REM --- (end of generated code: YConsolidatedDataSet definitions)

    REM --- (generated code: YConsolidatedDataSet attributes declaration)
    Protected _start As Double
    Protected _end As Double
    Protected _nsensors As Integer
    Protected _sensors As List(Of YSensor)
    Protected _datasets As List(Of YDataSet)
    Protected _progresss As List(Of Integer)
    Protected _nextidx As List(Of Integer)
    Protected _nexttim As List(Of Double)
    REM --- (end of generated code: YConsolidatedDataSet attributes declaration)


    Sub New(startTime As Double, endTime As Double, sensorList As List(Of YSensor))
      REM --- (generated code: YConsolidatedDataSet attributes initialization)
      _start = 0
      _end = 0
      _nsensors = 0
      _sensors = New List(Of YSensor)()
      _datasets = New List(Of YDataSet)()
      _progresss = New List(Of Integer)()
      _nextidx = New List(Of Integer)()
      _nexttim = New List(Of Double)()
      REM --- (end of generated code: YConsolidatedDataSet attributes initialization)
      imm_init(startTime, endTime, sensorList)
    End Sub

    REM --- (generated code: YConsolidatedDataSet private methods declaration)

    REM --- (end of generated code: YConsolidatedDataSet private methods declaration)

    REM --- (generated code: YConsolidatedDataSet public methods declaration)
    Public Overridable Function imm_init(startt As Double, endt As Double, sensorList As List(Of YSensor)) As Integer
      Me._start = startt
      Me._end = endt
      Me._sensors = sensorList
      Me._nsensors = -1
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns an object holding historical data for multiple
    '''   sensors, for a specified time interval.
    ''' <para>
    '''   The measures will be retrieved from the data logger, which must have been turned
    '''   on at the desired time. The resulting object makes it possible to load progressively
    '''   a large set of measures from multiple sensors, consolidating data on the fly
    '''   to align records based on measurement timestamps.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="sensorNames">
    '''   array of logical names or hardware identifiers of the sensors
    '''   for which data must be loaded from their data logger.
    ''' </param>
    ''' <param name="startTime">
    '''   the start of the desired measure time interval,
    '''   as a Unix timestamp, i.e. the number of seconds since
    '''   January 1, 1970 UTC. The special value 0 can be used
    '''   to include any measure, without initial limit.
    ''' </param>
    ''' <param name="endTime">
    '''   the end of the desired measure time interval,
    '''   as a Unix timestamp, i.e. the number of seconds since
    '''   January 1, 1970 UTC. The special value 0 can be used
    '''   to include any measure, without ending limit.
    ''' </param>
    ''' <returns>
    '''   an instance of <c>YConsolidatedDataSet</c>, providing access to
    '''   consolidated historical data. Records can be loaded progressively
    '''   using the <c>YConsolidatedDataSet.nextRecord()</c> method.
    ''' </returns>
    '''/
    Public Shared Function Init(sensorNames As List(Of String), startTime As Double, endTime As Double) As YConsolidatedDataSet
      Dim nSensors As Integer = 0
      Dim sensorList As List(Of YSensor) = New List(Of YSensor)()
      Dim idx As Integer = 0
      Dim sensorName As String
      Dim s As YSensor
      Dim obj As YConsolidatedDataSet
      nSensors = sensorNames.Count
      sensorList.Clear()
      idx = 0
      While (idx < nSensors)
        sensorName = sensorNames(idx)
        s = YSensor.FindSensor(sensorName)
        sensorList.Add(s)
        idx = idx + 1
      End While

      obj = New YConsolidatedDataSet(startTime, endTime, sensorList)
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Extracts the next data record from the data logger of all sensors linked to this
    '''   object.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="datarec">
    '''   array of floating point numbers, that will be filled by the
    '''   function with the timestamp of the measure in first position,
    '''   followed by the measured value in next positions.
    ''' </param>
    ''' <returns>
    '''   an integer in the range 0 to 100 (percentage of completion),
    '''   or a negative error code in case of failure.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function nextRecord(datarec As List(Of Double)) As Integer
      Dim s As Integer = 0
      Dim idx As Integer = 0
      Dim sensor As YSensor
      Dim newdataset As YDataSet
      Dim globprogress As Integer = 0
      Dim currprogress As Integer = 0
      Dim currnexttim As Double = 0
      Dim newvalue As Double = 0
      Dim measures As List(Of YMeasure) = New List(Of YMeasure)()
      Dim nexttime As Double = 0
      REM //
      REM // Ensure the dataset have been retrieved
      REM //
      If (Me._nsensors = -1) Then
        Me._nsensors = Me._sensors.Count
        Me._datasets.Clear()
        Me._progresss.Clear()
        Me._nextidx.Clear()
        Me._nexttim.Clear()
        s = 0
        While (s < Me._nsensors)
          sensor = Me._sensors(s)
          newdataset = sensor.get_recordedData(Me._start, Me._end)
          Me._datasets.Add(newdataset)
          Me._progresss.Add(0)
          Me._nextidx.Add(0)
          Me._nexttim.Add(0.0)
          s = s + 1
        End While
      End If
      datarec.Clear()
      REM //
      REM // Find next timestamp to process
      REM //
      nexttime = 0
      s = 0
      While (s < Me._nsensors)
        currnexttim = Me._nexttim(s)
        If (currnexttim = 0) Then
          idx = Me._nextidx(s)
          measures = Me._datasets(s).get_measures()
          currprogress = Me._progresss(s)
          While ((idx >= measures.Count) AndAlso (currprogress < 100))
            currprogress = Me._datasets(s).loadMore()
            If (currprogress < 0) Then
              currprogress = 100
            End If
            Me._progresss(s) = currprogress
            measures = Me._datasets(s).get_measures()
          End While
          If (idx < measures.Count) Then
            currnexttim = measures(idx).get_endTimeUTC()
            Me._nexttim(s) = currnexttim
          End If
        End If
        If (currnexttim > 0) Then
          If ((nexttime = 0) OrElse (nexttime > currnexttim)) Then
            nexttime = currnexttim
          End If
        End If
        s = s + 1
      End While
      If (nexttime = 0) Then
        Return 100
      End If
      REM //
      REM // Extract data for this timestamp
      REM //
      datarec.Clear()
      datarec.Add(nexttime)
      globprogress = 0
      s = 0
      While (s < Me._nsensors)
        If (Me._nexttim(s) = nexttime) Then
          idx = Me._nextidx(s)
          measures = Me._datasets(s).get_measures()
          newvalue = measures(idx).get_averageValue()
          datarec.Add(newvalue)
          Me._nexttim(s) = 0.0
          Me._nextidx(s) = idx + 1
        Else
          datarec.Add(Double.NaN)
        End If
        currprogress = Me._progresss(s)
        globprogress = globprogress + currprogress
        s = s + 1
      End While
      If (globprogress > 0) Then
        globprogress = (globprogress \ Me._nsensors)
        If (globprogress > 99) Then
          globprogress = 99
        End If
      End If

      Return globprogress
    End Function



    REM --- (end of generated code: YConsolidatedDataSet public methods declaration)
  End Class

  REM --- (generated code: YConsolidatedDataSet functions)


  REM --- (end of generated code: YConsolidatedDataSet functions)


  Public Class YDevice
    Private ReadOnly _devdescr As YDEV_DESCR
    Private _cacheStamp As yu64
    Private _cacheJson As YJSONObject
    Private ReadOnly _lock As New [Object]()
    Private ReadOnly _functions As New List(Of yu32)()

    Private _rootdevice As String
    Private _subpath As String

    Private _subpathinit As Boolean

    Private Sub New(devdesc As YDEV_DESCR)
      _devdescr = devdesc
      _cacheStamp = 0
      _cacheJson = Nothing
    End Sub


    Friend Sub dispose()
      clearCache(True)
    End Sub


    Friend Sub clearCache(clearSubpath As Boolean)
      SyncLock _lock
        _cacheStamp = 0
        If clearSubpath Then
          _cacheJson = Nothing
          _subpathinit = False
        End If
      End SyncLock
    End Sub


    Friend Shared Sub PlugDevice(devdescr As YDEV_DESCR)
      For idx As Integer = 0 To YDevice_devCache.Count - 1
        Dim dev As YDevice = YDevice_devCache(idx)
        If dev._devdescr = devdescr Then
          dev.clearCache(True)
          Exit For
        End If
      Next
    End Sub

    Friend Shared Function getDevice(devdescr As YDEV_DESCR) As YDevice
      Dim idx As Integer
      Dim dev As YDevice = Nothing
      For idx = 0 To YDevice_devCache.Count - 1
        If YDevice_devCache(idx)._devdescr = devdescr Then
          Return YDevice_devCache(idx)
        End If
      Next
      dev = New YDevice(devdescr)
      YDevice_devCache.Add(dev)
      Return dev
    End Function

    Private Function HTTPRequestSync(request_org As Byte(), ByRef reply As Byte(), ByRef errmsg As String) As YRETCODE
      Dim iohdl As YIOHDL
      Dim requestbuf As IntPtr = IntPtr.Zero
      Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim preply As IntPtr = IntPtr.Zero
      Dim replysize As Integer = 0
      Dim fullrequest As Byte() = Nothing
      Dim res As YRETCODE
      Dim enter As Boolean
      Do
        enter = Monitor.TryEnter(_lock)
        If Not enter Then
          Thread.Sleep(50)
        End If
      Loop Until enter
      Try
        res = HTTPRequestPrepare(request_org, fullrequest, errmsg)
        If YISERR(res) Then
          Return res
        End If

        iohdl.raw = 0  REM dummy, useless init to avoid compiler warning

        requestbuf = Marshal.AllocHGlobal(fullrequest.Length)
        Marshal.Copy(fullrequest, 0, requestbuf, fullrequest.Length)

        res = _yapiHTTPRequestSyncStartEx(iohdl, New StringBuilder(_rootdevice), requestbuf, fullrequest.Length, preply, replysize,
                                                            buffer)
        Marshal.FreeHGlobal(requestbuf)
        If res < 0 Then
          errmsg = buffer.ToString()
          Return res
        End If
        reply = New Byte(replysize - 1) {}
        If reply.Length > 0 AndAlso preply <> IntPtr.Zero Then
          Marshal.Copy(preply, reply, 0, replysize)
        End If
        res = _yapiHTTPRequestSyncDone(iohdl, buffer)
      Finally
        Monitor.Exit(_lock)
      End Try
      errmsg = buffer.ToString()
      Return res
    End Function

    Private Function HTTPRequestAsync(request As Byte(), ByRef errmsg As String) As YRETCODE
      Dim fullrequest As Byte() = Nothing
      Dim requestbuf As IntPtr = IntPtr.Zero
      Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE
      SyncLock _lock
        res = HTTPRequestPrepare(request, fullrequest, errmsg)
        If YISERR(res) Then
          Return res
        End If

        requestbuf = Marshal.AllocHGlobal(fullrequest.Length)
        Marshal.Copy(fullrequest, 0, requestbuf, fullrequest.Length)
        res = _yapiHTTPRequestAsyncEx(New StringBuilder(_rootdevice), requestbuf, fullrequest.Length, Nothing, Nothing, buffer)
      End SyncLock
      Marshal.FreeHGlobal(requestbuf)
      errmsg = buffer.ToString()
      Return res
    End Function

    Private Function HTTPRequestPrepare(request As Byte(), ByRef fullrequest As Byte(), ByRef errmsg As String) As YRETCODE
      Dim res As YRETCODE = Nothing
      Dim errbuf As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim b As StringBuilder = Nothing
      Dim neededsize As Integer = 0
      Dim p As Integer = 0
      Dim root As New StringBuilder(YOCTO_SERIAL_LEN)
      Dim tmp As Integer = 0

      ' no need to lock since it's already done by the called.
      If Not _subpathinit Then
        res = _yapiGetDevicePath(_devdescr, root, Nothing, 0, neededsize, errbuf)

        If YISERR(res) Then
          errmsg = errbuf.ToString()
          Return res
        End If

        b = New StringBuilder(neededsize)
        res = _yapiGetDevicePath(_devdescr, root, b, neededsize, tmp, errbuf)
        If YISERR(res) Then
          errmsg = errbuf.ToString()
          Return res
        End If

        _rootdevice = root.ToString()
        _subpath = b.ToString()
        _subpathinit = True
      End If
      ' search for the first '/'
      p = 0
      While p < request.Length AndAlso request(p) <> 47
        p += 1
      End While
      fullrequest = New Byte(request.Length - 1 + (_subpath.Length - 1)) {}
      Buffer.BlockCopy(request, 0, fullrequest, 0, p)
      Buffer.BlockCopy(System.Text.Encoding.ASCII.GetBytes(_subpath), 0, fullrequest, p, _subpath.Length)
      Buffer.BlockCopy(request, p + 1, fullrequest, p + _subpath.Length, request.Length - p - 1)

      Return YAPI.SUCCESS
    End Function


    Friend Function requestAPI(ByRef apires As YJSONObject, ByRef errmsg As String) As YRETCODE
      Dim buffer As String = ""
      Dim res As Integer = 0
      Dim http_headerlen As Integer
      Dim fw_release As String

      apires = Nothing
      SyncLock _lock
        ' Check if we have a valid cache value
        If _cacheStamp > YAPI.GetTickCount() Then
          apires = _cacheJson
          Return YAPI.SUCCESS
        End If
        Dim request As String
        If _cacheJson Is Nothing Then
          request = "GET /api.json " & vbCr & vbLf & vbCr & vbLf
        Else
          fw_release = YAPI._escapeAttr(_cacheJson.getYJSONObject("module").getString("firmwareRelease"))
          request = "GET /api.json?fw=" + fw_release + " " & vbCr & vbLf & vbCr & vbLf
        End If
        res = HTTPRequest(request, buffer, errmsg)
        If YISERR(res) Then
          ' make sure a device scan does not solve the issue
          res = yapiUpdateDeviceList(1, errmsg)
          If YISERR(res) Then
            Return res
          End If
          res = HTTPRequest(request, buffer, errmsg)
          If YISERR(res) Then
            Return res
          End If
        End If
        Dim httpcode As Integer = YAPI.ParseHTTP(buffer, 0, buffer.Length, http_headerlen, errmsg)
        If httpcode <> 200 Then
          Return YAPI.IO_ERROR
        End If
        Try
          apires = New YJSONObject(buffer, http_headerlen, buffer.Length)
          apires.parseWithRef(_cacheJson)
        Catch E As Exception
          errmsg = "unexpected JSON structure: " + E.Message
          Return YAPI.IO_ERROR
        End Try


        ' store result in cache
        _cacheJson = apires
        _cacheStamp = CULng(YAPI.GetTickCount() + YAPI.DefaultCacheValidity)
      End SyncLock
      Return YAPI.SUCCESS
    End Function


    Friend Function getFunctions(ByRef functions As List(Of yu32), ByRef errmsg As String) As YRETCODE
      Dim res As Integer = 0
      Dim neededsize As Integer = 0
      Dim i As Integer = 0
      Dim count As Integer = 0
      Dim p As IntPtr = Nothing
      Dim ids As ys32() = Nothing
      SyncLock _lock
        If _functions.Count = 0 Then
          res = yapiGetFunctionsByDevice(_devdescr, 0, IntPtr.Zero, 64, neededsize, errmsg)
          If YISERR(res) Then
            Return res
          End If

          p = Marshal.AllocHGlobal(neededsize)

          res = yapiGetFunctionsByDevice(_devdescr, 0, p, 64, neededsize, errmsg)
          If YISERR(res) Then
            Marshal.FreeHGlobal(p)
            Return res
          End If

          count = Convert.ToInt32(neededsize / Marshal.SizeOf(i))
          '  i is an 32 bits integer
          Array.Resize(ids, count + 1)
          Marshal.Copy(p, ids, 0, count)
          For i = 0 To count - 1
            _functions.Add(Convert.ToUInt32(ids(i)))
          Next

          Marshal.FreeHGlobal(p)
        End If
        functions = _functions
      End SyncLock
      Return YAPI.SUCCESS
    End Function

    '
    '         * Thread safe hepers
    '


    Friend Function HTTPRequest(request As Byte(), ByRef buffer As Byte(), ByRef errmsg As String) As YRETCODE
      Return HTTPRequestSync(request, buffer, errmsg)
    End Function


    Friend Function HTTPRequest(request As String, ByRef buffer As String, ByRef errmsg As String) As YRETCODE
      Dim binreply As Byte() = New Byte(-1) {}
      Dim res As YRETCODE = HTTPRequestSync(YAPI.DefaultEncoding.GetBytes(request), binreply, errmsg)
      buffer = YAPI.DefaultEncoding.GetString(binreply)
      Return res
    End Function

    Friend Function HTTPRequest(request As String, ByRef buffer As Byte(), ByRef errmsg As String) As YRETCODE
      Return HTTPRequestSync(YAPI.DefaultEncoding.GetBytes(request), buffer, errmsg)
    End Function


    Friend Function HTTPRequestAsync(request As String, ByRef errmsg As String) As YRETCODE
      Return Me.HTTPRequestAsync(YAPI.DefaultEncoding.GetBytes(request), errmsg)
    End Function
  End Class


  REM --- (generated code: YFunction class start)

  '''*
  ''' <summary>
  '''   This is the parent class for all public objects representing device functions documented in
  '''   the high-level programming API.
  ''' <para>
  '''   This abstract class does all the real job, but without
  '''   knowledge of the specific function attributes.
  ''' </para>
  ''' <para>
  '''   Instantiating a child class of YFunction does not cause any communication.
  '''   The instance simply keeps track of its function identifier, and will dynamically bind
  '''   to a matching device at the time it is really being used to read or set an attribute.
  '''   In order to allow true hot-plug replacement of one device by another, the binding stay
  '''   dynamic through the life of the object.
  ''' </para>
  ''' <para>
  '''   The YFunction class implements a generic high-level cache for the attribute values of
  '''   the specified function, pre-parsed from the REST API string.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YFunction
    REM --- (end of generated code: YFunction class start)
    REM --- (generated code: YFunction definitions)
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of generated code: YFunction definitions)
    Public Const FUNCTIONDESCRIPTOR_INVALID As YFUN_DESCR = -1

    Public Shared _cache As Dictionary(Of String, YFunction) = New Dictionary(Of String, YFunction)
    Public Shared _ValueCallbackList As List(Of YFunction) = New List(Of YFunction)
    Public Shared _TimedReportCallbackList As List(Of YSensor) = New List(Of YSensor)

    REM --- (generated code: YFunction attributes declaration)
    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _valueCallbackFunction As YFunctionValueCallback
    Protected _cacheExpiration As Long
    Protected _serial As String
    Protected _funId As String
    Protected _hwId As String
    REM --- (end of generated code: YFunction attributes declaration)
    Protected _className As String
    Protected _func As String
    Protected _lastErrorType As YRETCODE
    Protected _lastErrorMsg As String
    Protected _fundescr As YFUN_DESCR
    Protected _userData As Object
    Protected _dataStreams As Hashtable = New Hashtable()

    Public Sub New(ByRef func As String)
      REM --- (generated code: YFunction attributes initialization)
      _logicalName = LOGICALNAME_INVALID
      _advertisedValue = ADVERTISEDVALUE_INVALID
      _valueCallbackFunction = Nothing
      _cacheExpiration = 0
      REM --- (end of generated code: YFunction attributes initialization)
      _className = "Function"
      _func = func
      _lastErrorType = YAPI_SUCCESS
      _lastErrorMsg = ""
      _cacheExpiration = 0
      _fundescr = FUNCTIONDESCRIPTOR_INVALID
      _userData = Nothing
    End Sub


    Protected Shared Function _FindFromCache(ByVal classname As String, ByVal func As String) As YFunction
      Dim key As String
      key = classname + "_" + func
      If (_cache.ContainsKey(key)) Then
        Return _cache.Item(key)
      End If
      Return Nothing
    End Function

    Protected Shared Sub _AddToCache(ByVal classname As String, ByVal func As String, obj As YFunction)
      _cache.Add(classname + "_" + func, obj)
    End Sub

    Public Shared Sub _ClearCache(ByVal classname As String, ByVal func As String)
      _cache.Clear()
    End Sub

    Public Sub _throw(ByVal errType As YRETCODE, ByVal errMsg As String)
      _lastErrorType = errType
      _lastErrorMsg = errMsg
      If Not (YAPI.ExceptionsDisabled) Then
        Throw New YAPI_Exception(errType, "YoctoApi error : " + errMsg)
      End If
    End Sub

    REM  Method used to resolve our name to our unique function descriptor (may trigger a hub scan)
    Protected Function _getDescriptor(ByRef fundescr As YFUN_DESCR, ByRef errMsg As String) As YRETCODE
      Dim res As Integer
      Dim tmp_fundescr As YFUN_DESCR

      tmp_fundescr = yapiGetFunction(_className, _func, errMsg)
      If (YISERR(tmp_fundescr)) Then
        res = yapiUpdateDeviceList(1, errMsg)
        If (YISERR(res)) Then Return res

        tmp_fundescr = yapiGetFunction(_className, _func, errMsg)
        If (YISERR(tmp_fundescr)) Then Return tmp_fundescr

      End If

      fundescr = tmp_fundescr
      _fundescr = tmp_fundescr
      _getDescriptor = YAPI_SUCCESS
    End Function

    REM Return a pointer to our device caching object (may trigger a hub scan)
    Protected Function _getDevice(ByRef dev As YDevice, ByRef errMsg As String) As YRETCODE
      Dim fundescr As YFUN_DESCR
      Dim devdescr As YDEV_DESCR
      Dim res As YRETCODE

      REM Resolve function name
      res = _getDescriptor(fundescr, errMsg)
      If (YISERR(res)) Then
        _getDevice = res
        Exit Function
      End If

      REM Get device descriptor
      devdescr = yapiGetDeviceByFunction(fundescr, errMsg)
      If (YISERR(devdescr)) Then
        _getDevice = res
        Exit Function
      End If

      REM Get device object
      dev = YDevice.getDevice(devdescr)

      _getDevice = YAPI_SUCCESS
    End Function

    REM Return the next known function of current class listed in the yellow pages
    Protected Function _nextFunction(ByRef hwid As String) As YRETCODE

      Dim fundescr As YFUN_DESCR
      Dim devdescr As YDEV_DESCR
      Dim serial As String = ""
      Dim funcId As String = ""
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res, count As Integer
      Dim neededsize, maxsize As Integer
      Dim p As IntPtr

      Const n_element As Integer = 1
      Dim pdata(n_element) As Integer

      res = _getDescriptor(fundescr, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        _nextFunction = res
        Exit Function
      End If

      maxsize = n_element * Marshal.SizeOf(pdata(0))
      p = Marshal.AllocHGlobal(maxsize)
      res = yapiGetFunctionsByClass(_className, fundescr, p, maxsize, neededsize, errmsg)
      Marshal.Copy(p, pdata, 0, n_element)
      Marshal.FreeHGlobal(p)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        _nextFunction = res
        Exit Function
      End If

      count = CInt(neededsize / Marshal.SizeOf(pdata(0)))
      If count = 0 Then
        hwid = ""
        _nextFunction = YAPI_SUCCESS
        Exit Function
      End If

      res = yapiGetFunctionInfo(pdata(0), devdescr, serial, funcId, funcName, funcVal, errmsg)

      If (YISERR(res)) Then
        _throw(res, errmsg)
        _nextFunction = YAPI_SUCCESS
        Exit Function
      End If

      hwid = serial + "." + funcId
      _nextFunction = YAPI_SUCCESS
    End Function

    Protected Function _escapeAttr(ByVal changeval As String) As String
      _escapeAttr = YAPI._escapeAttr(changeval)

    End Function


    Private Function _buildSetRequest(ByVal changeattr As String, ByVal changeval As String, ByRef request As String, ByRef errmsg As String) As YRETCODE
      Dim res As Integer
      Dim fundesc As YFUN_DESCR
      Dim funcid As New StringBuilder(YOCTO_FUNCTION_LEN)
      Dim errbuff As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim devdesc As YDEV_DESCR

      funcid.Length = 0
      errbuff.Length = 0

      REM Resolve the function name
      res = _getDescriptor(fundesc, errmsg)

      If (YISERR(res)) Then
        _buildSetRequest = res
        Exit Function
      End If

      res = _yapiGetFunctionInfoEx(fundesc, devdesc, Nothing, funcid, Nothing, Nothing, Nothing, errbuff)
      If YISERR(res) Then
        errmsg = errbuff.ToString()
        _throw(res, errmsg)
        _buildSetRequest = res
        Exit Function
      End If

      request = "GET /api/" + funcid.ToString() + "/"
      If (changeattr <> "") Then
        request = request + changeattr + "?" + changeattr + "=" + _escapeAttr(changeval)
      End If
      request = request + "&. " + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      _buildSetRequest = YAPI_SUCCESS
    End Function


    REM Set an attribute in the function, and parse the resulting new function state
    Protected Function _setAttr(ByVal attrname As String, ByVal newvalue As String) As YRETCODE

      Dim errmsg As String = ""
      Dim request As String = ""
      Dim res As Integer
      Dim dev As YDevice = Nothing

      REM  Execute http request
      res = _buildSetRequest(attrname, newvalue, request, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        _setAttr = res
        Exit Function
      End If

      REM Get device Object
      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        _setAttr = res
        Exit Function
      End If

      res = dev.HTTPRequestAsync(request, errmsg)
      If (YISERR(res)) Then
        REM make sure a device scan does not solve the issue
        res = yapiUpdateDeviceList(1, errmsg)
        If (YISERR(res)) Then
          _throw(res, errmsg)
          _setAttr = res
          Exit Function
        End If
        res = dev.HTTPRequestAsync(request, errmsg)
        If (YISERR(res)) Then
          _throw(res, errmsg)
          _setAttr = res
          Exit Function
        End If
      End If
      dev.clearCache(False)
      If (_cacheExpiration <> 0) Then
        _cacheExpiration = YAPI.GetTickCount()
      End If
      _setAttr = YAPI_SUCCESS
    End Function


    REM Method used to send http request to the device (not the function)
    Protected Function _request(ByVal request As String) As Byte()
      Return _request(YAPI.DefaultEncoding.GetBytes(request))
    End Function

    REM Method used to send http request to the device (not the function)
    Protected Function _request(ByVal request As Byte()) As Byte()
      Dim errbuffer As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      errbuffer.Length = 0
      Dim dev As YDevice = Nothing
      Dim errmsg As String = ""
      Dim outbuf As Byte()
      Dim check As Byte()
      Dim res As Integer

      REM Resolve our reference to our device, load REST API
      ReDim outbuf(-1)
      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        Return outbuf
      End If
      res = dev.HTTPRequest(request, outbuf, errmsg)
      If (YISERR(res)) Then
        REM Check if an update of the device list does notb solve the issue
        res = _yapiUpdateDeviceList(1, errbuffer)
        If (YISERR(res)) Then
          _throw(res, errbuffer.ToString)
          ReDim outbuf(-1)
          Return outbuf
        End If
        res = dev.HTTPRequest(request, outbuf, errmsg)
        If (YISERR(res)) Then
          _throw(res, errmsg)
          ReDim outbuf(-1)
          Return outbuf
        End If
      End If
      If outbuf.Length >= 4 Then
        ReDim check(3)
        Buffer.BlockCopy(outbuf, 0, check, 0, 4)
        If (YAPI.DefaultEncoding.GetString(check) = "OK" + Chr(13) + Chr(10)) Then
          Return outbuf
        End If
        If outbuf.Length >= 17 Then
          ReDim check(16)
          Buffer.BlockCopy(outbuf, 0, check, 0, 17)
          If (YAPI.DefaultEncoding.GetString(check) = "HTTP/1.1 200 OK" + Chr(13) + Chr(10)) Then
            Return outbuf
          End If
        End If
      End If
      _throw(YAPI.IO_ERROR, "http request failed")
      ReDim outbuf(-1)
      Return outbuf
    End Function

    REM Method used to send http request to the device (not the function)
    Public Function _download(ByVal url As String) As Byte()
      Dim request As String
      Dim outbuf As Byte()

      request = "GET /" + url + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      outbuf = _request(request)
      Return _strip_http_header(outbuf)
    End Function

    Private Function _strip_http_header(outbuf As Byte()) As Byte()
      Dim found As Integer
      Dim body As Integer
      Dim res As Byte()

      If (outbuf.Length = 0) Then
        Return outbuf
      End If
      found = 0
      Do While found < outbuf.Length - 4 And
                 (outbuf(found) <> 13 Or outbuf(found + 1) <> 10 Or
                  outbuf(found + 2) <> 13 Or outbuf(found + 3) <> 10)
        found += 1
      Loop
      If found > outbuf.Length - 4 Then
        _throw(YAPI.IO_ERROR, "http request failed")
        ReDim outbuf(-1)
        Return outbuf
      End If
      If found = outbuf.Length - 4 Then
        ReDim outbuf(0)
        Return outbuf
      End If
      body = found + 4
      ReDim res(outbuf.Length - body - 1)
      Buffer.BlockCopy(outbuf, body, res, 0, outbuf.Length - body)
      Return res
    End Function

    REM Method used to upload a file to the device
    Public Function _upload(ByVal path As String, ByVal content As String) As Integer
      Return _upload(path, YAPI.DefaultEncoding.GetBytes(content))
    End Function

    REM Method used to upload a file to the device
    Public Function _upload(ByVal path As String, ByVal content As List(Of Byte)) As Integer
      Return _upload(path, content.ToArray())
    End Function

    REM Method used to upload a file to the device
    Public Function _upload(ByVal path As String, ByVal content As Byte()) As Integer
      Dim outbuf As Byte()
      outbuf = _uploadEx(path, content)
      If outbuf.Length = 0 Then
        Return YAPI.IO_ERROR
      End If
      Return YAPI.SUCCESS
    End Function

    REM Method used to upload a file to the device
    Public Function _uploadEx(ByVal path As String, ByVal content As Byte()) As Byte()
      Dim bodystr, boundary As String
      Dim body, bb, header, footer, fullrequest, outbuf As Byte()

      bodystr = "Content-Disposition: form-data; name=""" + path + """; filename=""api""" + Chr(13) + Chr(10) +
                "Content-Type: application/octet-stream" + Chr(13) + Chr(10) +
                "Content-Transfer-Encoding: binary" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      ReDim body(bodystr.Length + content.Length - 1)
      Buffer.BlockCopy(YAPI.DefaultEncoding.GetBytes(bodystr), 0, body, 0, bodystr.Length)
      Buffer.BlockCopy(content, 0, body, bodystr.Length, content.Length)

      Dim random As Random = New Random()
      Dim pos, i As Integer
      Do
        boundary = "Zz" + (CInt(random.Next(100000, 999999))).ToString() + "zZ"
        bb = YAPI.DefaultEncoding.GetBytes(boundary)
        pos = 0
        i = 0
        Do While pos <= body.Length - bb.Length And i < bb.Length
          If body(pos) = 90 Then REM 'Z'
            i = 1
            Do While i < bb.Length And body(pos + i) = bb(i)
              i += 1
            Loop
            If i < bb.Length Then
              pos += i
            End If
          Else
            pos += 1
          End If
        Loop
      Loop While pos <= body.Length - bb.Length

      header = YAPI.DefaultEncoding.GetBytes("POST /upload.html HTTP/1.1" + Chr(13) + Chr(10) +
                                             "Content-Type: multipart/form-data, boundary=" + boundary + Chr(13) + Chr(10) +
                                             Chr(13) + Chr(10) + "--" + boundary + Chr(13) + Chr(10))
      footer = YAPI.DefaultEncoding.GetBytes(Chr(13) + Chr(10) + "--" + boundary + "--" + Chr(13) + Chr(10))
      ReDim fullrequest(header.Length + body.Length + footer.Length - 1)
      Buffer.BlockCopy(header, 0, fullrequest, 0, header.Length)
      Buffer.BlockCopy(body, 0, fullrequest, header.Length, body.Length)
      Buffer.BlockCopy(footer, 0, fullrequest, header.Length + body.Length, footer.Length)

      outbuf = _request(fullrequest)
      Return _strip_http_header(outbuf)
    End Function

    Public Function _findDataStream(dataset As YDataSet, def As String) As YDataStream
      Dim key As String = dataset.get_functionId() + ":" + def
      Dim newDataStream As YDataStream
      Dim words As List(Of Integer)
      If (_dataStreams.ContainsKey(key)) Then
        Return CType(_dataStreams(key), YDataStream)
      End If
      words = YAPI._decodeWords(def)
      If (words.Count < 14) Then
        _throw(YAPI_VERSION_MISMATCH, "device firmware is too old")
        Return Nothing
      End If
      newDataStream = New YDataStream(Me, dataset, words)
      Return newDataStream
    End Function

    Public Sub _clearDataStreamCache()
      _dataStreams.Clear()
    End Sub

    Protected Function _parse(ByRef j As YJSONObject) As Integer
      _parseAttr(j)
      _parserHelper()
      Return 0
    End Function

    Protected Shared Sub _UpdateValueCallbackList(func As YFunction, add As Boolean)
      If (add) Then
        func.isOnline()
        If Not (_ValueCallbackList.Contains(func)) Then
          _ValueCallbackList.Add(func)
        End If
      Else
        _ValueCallbackList.Remove(func)
      End If
    End Sub

    Protected Shared Sub _UpdateTimedReportCallbackList(func As YSensor, add As Boolean)
      If (add) Then
        func.isOnline()
        If Not (_TimedReportCallbackList.Contains(func)) Then
          _TimedReportCallbackList.Add(func)
        End If
      Else
        _TimedReportCallbackList.Remove(func)
      End If
    End Sub

    REM --- (generated code: YFunction private methods declaration)

    Protected Overridable Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("logicalName") Then
        _logicalName = json_val.getString("logicalName")
      End If
      If json_val.has("advertisedValue") Then
        _advertisedValue = json_val.getString("advertisedValue")
      End If
      Return 0
    End Function

    REM --- (end of generated code: YFunction private methods declaration)

    REM --- (generated code: YFunction public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the logical name of the function.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YFunction.LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LOGICALNAME_INVALID
        End If
      End If
      res = Me._logicalName
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the logical name of the function.
    ''' <para>
    '''   You can use <c>yCheckLogicalName()</c>
    '''   prior to this call to make sure that your parameter is valid.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the logical name of the function
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_logicalName(ByVal newval As String) As Integer
      Dim rest_val As String
      If Not YAPI.CheckLogicalName(newval) Then
        _throw(YAPI.INVALID_ARGUMENT, "Invalid name :" + newval)
        Return YAPI.INVALID_ARGUMENT
      End If
      rest_val = newval
      Return _setAttr("logicalName", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns a short string representing the current state of the function.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to a short string representing the current state of the function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YFunction.ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ADVERTISEDVALUE_INVALID
        End If
      End If
      res = Me._advertisedValue
      Return res
    End Function


    Public Function set_advertisedValue(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("advertisedValue", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a function for a given identifier.
    ''' <para>
    '''   The identifier can be specified using several formats:
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   - FunctionLogicalName
    ''' </para>
    ''' <para>
    '''   - ModuleSerialNumber.FunctionIdentifier
    ''' </para>
    ''' <para>
    '''   - ModuleSerialNumber.FunctionLogicalName
    ''' </para>
    ''' <para>
    '''   - ModuleLogicalName.FunctionIdentifier
    ''' </para>
    ''' <para>
    '''   - ModuleLogicalName.FunctionLogicalName
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   This function does not require that the function is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YFunction.isOnline()</c> to test if the function is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a function by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' <para>
    '''   If a call to this object's is_online() method returns FALSE although
    '''   you are certain that the matching device is plugged, make sure that you did
    '''   call registerHub() at application initialization time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the function, for instance
    '''   <c>MyDevice.</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YFunction</c> object allowing you to drive the function.
    ''' </returns>
    '''/
    Public Shared Function FindFunction(func As String) As YFunction
      Dim obj As YFunction
      obj = CType(YFunction._FindFromCache("Function", func), YFunction)
      If ((obj Is Nothing)) Then
        obj = New YFunction(func)
        YFunction._AddToCache("Function", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Function registerValueCallback(callback As YFunctionValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackFunction = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) AndAlso Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overridable Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackFunction Is Nothing)) Then
        Me._valueCallbackFunction(Me, value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Disables the propagation of every new advertised value to the parent hub.
    ''' <para>
    '''   You can use this function to save bandwidth and CPU on computers with limited
    '''   resources, or to prevent unwanted invocations of the HTTP callback.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function muteValueCallbacks() As Integer
      Return Me.set_advertisedValue("SILENT")
    End Function

    '''*
    ''' <summary>
    '''   Re-enables the propagation of every new advertised value to the parent hub.
    ''' <para>
    '''   This function reverts the effect of a previous call to <c>muteValueCallbacks()</c>.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function unmuteValueCallbacks() As Integer
      Return Me.set_advertisedValue("")
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of a single function attribute, as a text string, as quickly as
    '''   possible but without using the cached value.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="attrName">
    '''   the name of the requested attribute
    ''' </param>
    ''' <returns>
    '''   a string with the value of the the attribute
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function loadAttribute(attrName As String) As String
      Dim url As String
      Dim attrVal As Byte() = New Byte(){}
      url = "api/" + Me.get_functionId() + "/" + attrName
      attrVal = Me._download(url)
      Return YAPI.DefaultEncoding.GetString(attrVal)
    End Function

    '''*
    ''' <summary>
    '''   Indicates whether changes to the function are prohibited or allowed.
    ''' <para>
    '''   Returns <c>true</c> if the function is blocked by an admin password
    '''   or if the function is not available.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>true</c> if the function is write-protected or not online.
    ''' </returns>
    '''/
    Public Overridable Function isReadOnly() As Boolean
      Dim serial As String
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As Integer = 0
      Try
          serial = Me.get_serialNumber()
      Catch
          Return True
      End Try
      res = _yapiIsModuleWritable(New StringBuilder(serial), errmsg)
      If (res > 0) Then
        Return False
      End If
      Return True
    End Function

    '''*
    ''' <summary>
    '''   Returns the serial number of the module, as set by the factory.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the serial number of the module, as set by the factory.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns YFunction.SERIALNUMBER_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_serialNumber() As String
      Dim m As YModule
      m = Me.get_module()
      Return m.get_serialNumber()
    End Function

    Public Overridable Function _parserHelper() As Integer
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   c
    ''' <para>
    '''   omment from .yc definition
    ''' </para>
    ''' </summary>
    '''/
    Public Function nextFunction() As YFunction
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YFunction.FindFunction(hwid)
    End Function

    '''*
    ''' <summary>
    '''   c
    ''' <para>
    '''   omment from .yc definition
    ''' </para>
    ''' </summary>
    '''/
    Public Shared Function FirstFunction() As YFunction
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Function", 0, p, size, neededsize, errmsg)
      Marshal.Copy(p, v_fundescr, 0, 1)
      Marshal.FreeHGlobal(p)

      If (YISERR(err) Or (neededsize = 0)) Then
        Return Nothing
      End If
      serial = ""
      funcId = ""
      funcName = ""
      funcVal = ""
      errmsg = ""
      If (YISERR(yapiGetFunctionInfo(v_fundescr(0), dev, serial, funcId, funcName, funcVal, errmsg))) Then
        Return Nothing
      End If
      Return YFunction.FindFunction(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YFunction public methods declaration)

    '''*
    ''' <summary>
    '''   Returns a short text that describes unambiguously the instance of the function in the form <c>TYPE(NAME)=SERIAL&#46;FUNCTIONID</c>.
    ''' <para>
    '''   More precisely,
    '''   <c>TYPE</c>       is the type of the function,
    '''   <c>NAME</c>       it the name used for the first access to the function,
    '''   <c>SERIAL</c>     is the serial number of the module if the module is connected or <c>"unresolved"</c>, and
    '''   <c>FUNCTIONID</c> is  the hardware identifier of the function if the module is connected.
    '''   For example, this method returns <c>Relay(MyCustomName.relay1)=RELAYLO1-123456.relay1</c> if the
    '''   module is already connected or <c>Relay(BadCustomeName.relay1)=unresolved</c> if the module has
    '''   not yet been connected. This method does not trigger any USB or TCP transaction and can therefore be used in
    '''   a debugger.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string that describes the function
    '''   (ex: <c>Relay(MyCustomName.relay1)=RELAYLO1-123456.relay1</c>)
    ''' </returns>
    '''/
    Public Function describe() As String
      Dim fundescr As YFUN_DESCR
      Dim devdescr As YDEV_DESCR
      Dim errmsg As String = ""
      Dim serial As String = ""
      Dim funcId As String = ""
      Dim funcName As String = ""
      Dim funcValue As String = ""
      fundescr = yapiGetFunction(_className, _func, errmsg)
      If (Not (YISERR(fundescr))) Then
        If (Not (YISERR(yapiGetFunctionInfo(fundescr, devdescr, serial, funcId, funcName, funcValue, errmsg)))) Then
          describe = _className + "(" + _func + ")=" + serial + "." + funcId
          Exit Function
        End If
      End If
      describe = _className + "(" + _func + ")=unresolved"
    End Function

    '''*
    ''' <summary>
    '''   Returns the unique hardware identifier of the function in the form <c>SERIAL.FUNCTIONID</c>.
    ''' <para>
    '''   The unique hardware identifier is composed of the device serial
    '''   number and of the hardware identifier of the function (for example <c>RELAYLO1-123456.relay1</c>).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string that uniquely identifies the function (ex: <c>RELAYLO1-123456.relay1</c>)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns  <c>YFunction.HARDWAREID_INVALID</c>.
    ''' </para>
    '''/

    Public Overridable Function get_hardwareId() As String
      Dim retcode As YRETCODE
      Dim fundesc As YFUN_DESCR = 0
      Dim devdesc As YDEV_DESCR = 0
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim snum As String = ""
      Dim funcid As String = ""

      REM Resolve the function name
      retcode = _getDescriptor(fundesc, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_hardwareId = YAPI.HARDWAREID_INVALID
        Exit Function
      End If

      retcode = yapiGetFunctionInfo(fundesc, devdesc, snum, funcid, funcName, funcVal, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_hardwareId = YAPI.HARDWAREID_INVALID
        Exit Function
      End If

      get_hardwareId = snum + "." + funcid
    End Function

    '''*
    ''' <summary>
    '''   Returns the hardware identifier of the function, without reference to the module.
    ''' <para>
    '''   For example
    '''   <c>relay1</c>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string that identifies the function (ex: <c>relay1</c>)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns  <c>YFunction.FUNCTIONID_INVALID</c>.
    ''' </para>
    '''/

    Public Function get_functionId() As String
      Dim retcode As YRETCODE
      Dim fundesc As YFUN_DESCR = 0
      Dim devdesc As YDEV_DESCR = 0
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim snum As String = ""
      Dim funcid As String = ""

      REM Resolve the function name
      retcode = _getDescriptor(fundesc, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_functionId = YAPI.FUNCTIONID_INVALID
        Exit Function
      End If

      retcode = yapiGetFunctionInfo(fundesc, devdesc, snum, funcid, funcName, funcVal, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_functionId = YAPI.FUNCTIONID_INVALID
        Exit Function
      End If

      get_functionId = funcid
    End Function

    Public Function get_friendlyName() As String
      Dim retcode As YRETCODE
      Dim fundesc As YFUN_DESCR = 0
      Dim devdesc As YDEV_DESCR = 0
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim snum As String = ""
      Dim funcid As String = ""
      Dim friendfunc As String
      Dim friendmod As String
      REM Resolve the function name
      retcode = _getDescriptor(fundesc, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_friendlyName = YAPI.FRIENDLYNAME_INVALID
        Exit Function
      End If

      retcode = yapiGetFunctionInfo(fundesc, devdesc, snum, funcid, funcName, funcVal, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_friendlyName = YAPI.FRIENDLYNAME_INVALID
        Exit Function
      End If
      Dim moddescr As YFUN_DESCR
      moddescr = yapiGetFunction("Module", snum, errmsg)
      If (YISERR(moddescr)) Then
        _throw(moddescr, errmsg)
        get_friendlyName = YAPI.FRIENDLYNAME_INVALID
        Exit Function
      End If

      friendmod = snum
      friendfunc = funcid
      If funcName <> "" Then
        friendfunc = funcName
      End If
      retcode = yapiGetFunctionInfo(moddescr, devdesc, snum, funcid, funcName, funcVal, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_friendlyName = YAPI.FRIENDLYNAME_INVALID
        Exit Function
      End If
      If funcName <> "" Then
        friendmod = funcName
      End If
      get_friendlyName = friendmod + "." + friendfunc
    End Function


    Public Overrides Function ToString() As String
      Return describe()
    End Function

    '''*
    ''' <summary>
    '''   Returns the numerical error code of the latest error with the function.
    ''' <para>
    '''   This method is mostly useful when using the Yoctopuce library with
    '''   exceptions disabled.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a number corresponding to the code of the latest error that occurred while
    '''   using the function object
    ''' </returns>
    '''/
    Public Function get_errorType() As YRETCODE
      Return _lastErrorType
    End Function

    Public Function errorType() As YRETCODE
      Return _lastErrorType
    End Function

    Public Function errType() As YRETCODE
      Return _lastErrorType
    End Function

    '''*
    ''' <summary>
    '''   Returns the error message of the latest error with the function.
    ''' <para>
    '''   This method is mostly useful when using the Yoctopuce library with
    '''   exceptions disabled.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest error message that occured while
    '''   using the function object
    ''' </returns>
    '''/
    Public Function get_errorMessage() As String
      Return _lastErrorMsg
    End Function

    Public Function errorMessage() As String
      Return _lastErrorMsg
    End Function

    Public Function errMessage() As String
      Return _lastErrorMsg
    End Function

    '''*
    ''' <summary>
    '''   Checks if the function is currently reachable, without raising any error.
    ''' <para>
    '''   If there is a cached value for the function in cache, that has not yet
    '''   expired, the device is considered reachable.
    '''   No exception is raised if there is an error while trying to contact the
    '''   device hosting the function.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>true</c> if the function can be reached, and <c>false</c> otherwise
    ''' </returns>
    '''/
    Public Function isOnline() As Boolean

      Dim dev As YDevice = Nothing
      Dim errmsg As String = ""
      Dim apires As YJSONObject = Nothing

      REM  A valid value in cache means that the device is online
      If (_cacheExpiration > YAPI.GetTickCount()) Then
        isOnline = True
        Exit Function
      End If

      REM Check that the function is available, without throwing exceptions
      If (YISERR(_getDevice(dev, errmsg))) Then
        isOnline = False
        Exit Function
      End If

      REM Try to execute a function request to be positively sure that the device is ready
      If (YISERR(dev.requestAPI(apires, errmsg))) Then
        isOnline = False
        Exit Function
      End If

      REM Preload the function data, since we have it in device cache
      load(YAPI.DefaultCacheValidity)

      isOnline = True
    End Function


    Protected Function _json_get_key(ByVal data As Byte(), ByVal key As String) As String
      Dim obj As New YJSONObject(YAPI.DefaultEncoding.GetString(data))
      obj.parse()
      If obj.has(key) Then
        Dim val As String = obj.getString(key)
        If val Is Nothing Then
          val = obj.ToString()
        End If
        Return val
      End If
      Throw New YAPI_Exception(YAPI.INVALID_ARGUMENT, (Convert.ToString("No key ") & key) + "in JSON struct")
    End Function

    Public Function _json_get_array(data As Byte()) As List(Of Byte())
      Dim array As New YJSONArray(YAPI.DefaultEncoding.GetString(data))
      array.parse()
      Dim list As New List(Of Byte())()
      Dim len As Integer = array.Length
      For i As Integer = 0 To len - 1
        Dim o As YJSONContent = array.[get](i)
        list.Add(o.toJSON())
      Next
      Return list
    End Function

    Public Function _json_get_string(data As Byte()) As String
      Dim s As String = YAPI.DefaultEncoding.GetString(data)
      Dim jstring As New YJSONString(s, 0, s.Length)
      jstring.parse()
      Return jstring.getString()
    End Function


    Private Function get_json_path_struct(jsonObject As YJSONObject, paths As String(), ofs As Integer) As Byte()

      Dim key As String = paths(ofs)
      If Not jsonObject.has(key) Then
        Return New Byte() {}
      End If

      Dim obj As YJSONContent = jsonObject.[get](key)
      If obj IsNot Nothing Then
        If paths.Length = ofs + 1 Then
          Return obj.toJSON()
        End If

        If TypeOf obj Is YJSONArray Then
          Return get_json_path_array(jsonObject.getYJSONArray(key), paths, ofs + 1)
        ElseIf TypeOf obj Is YJSONObject Then
          Return get_json_path_struct(jsonObject.getYJSONObject(key), paths, ofs + 1)
        End If
      End If
      Return New Byte() {}
    End Function

    Private Function get_json_path_array(jsonArray As YJSONArray, paths As String(), ofs As Integer) As Byte()
      Dim key As Integer = Convert.ToInt32(paths(ofs))
      If jsonArray.Length <= key Then
        Return New Byte() {}
      End If

      Dim obj As YJSONContent = jsonArray.[get](key)
      If obj IsNot Nothing Then
        If paths.Length = ofs + 1 Then
          Return obj.toJSON()
        End If

        If TypeOf obj Is YJSONArray Then
          Return get_json_path_array(jsonArray.getYJSONArray(key), paths, ofs + 1)
        ElseIf TypeOf obj Is YJSONObject Then
          Return get_json_path_struct(jsonArray.getYJSONObject(key), paths, ofs + 1)
        End If
      End If
      Return New Byte() {}
    End Function


    Public Function _get_json_path(ByVal json As Byte(), ByVal path As String) As Byte()

      Dim jsonObject As YJSONObject = Nothing
      jsonObject = New YJSONObject(YAPI.DefaultEncoding.GetString(json))
      jsonObject.parse()
      Dim split As String() = path.Split(New Char() {"\"c, "|"c})
      Return get_json_path_struct(jsonObject, split, 0)
    End Function


    Public Function _decode_json_int(ByVal json As Byte()) As Integer
      Dim s As String = YAPI.DefaultEncoding.GetString(json)
      Dim obj As YJSONNumber = New YJSONNumber(s, 0, s.Length)
      obj.parse()
      Return obj.getInt()
    End Function

    Public Function _decode_json_string(ByVal json As Byte()) As String
      Dim len As Integer
      Dim buffer As StringBuilder
      Dim decoded_len As Integer
      len = json.Length
      buffer = New StringBuilder(len)
      If len > 0 Then
        decoded_len = _yapiJsonDecodeString(New StringBuilder(YAPI.DefaultEncoding.GetString(json)), buffer)
      End If
      Return buffer.ToString()
    End Function

    '''*
    ''' <summary>
    '''   Preloads the function cache with a specified validity duration.
    ''' <para>
    '''   By default, whenever accessing a device, all function attributes
    '''   are kept in cache for the standard duration (5 ms). This method can be
    '''   used to temporarily mark the cache as valid for a longer period, in order
    '''   to reduce network traffic for instance.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="msValidity">
    '''   an integer corresponding to the validity attributed to the
    '''   loaded function parameters, in milliseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function load(ByVal msValidity As Long) As YRETCODE
      Dim dev As YDevice = Nothing
      Dim errmsg As String = ""
      Dim apires As YJSONObject = Nothing
      Dim fundescr As YFUN_DESCR = Nothing
      Dim res As Integer = 0
      Dim errbuf As String = ""
      Dim funcId As String = ""
      Dim devdesc As YDEV_DESCR = Nothing
      Dim serial As String = ""
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim node As YJSONObject

      REM Resolve our reference to our device, load REST API
      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        load = res
        Exit Function
      End If

      res = dev.requestAPI(apires, errmsg)
      If YISERR(res) Then
        _throw(res, errmsg)
        load = res
        Exit Function
      End If

      REM Get our function Id
      fundescr = yapiGetFunction(_className, _func, errmsg)
      If YISERR(fundescr) Then
        _throw(res, errmsg)
        load = res
        Exit Function
      End If

      devdesc = 0
      res = yapiGetFunctionInfo(fundescr, devdesc, serial, funcId, funcName, funcVal, errbuf)
      If YISERR(res) Then
        _throw(res, errmsg)
        load = res
        Exit Function
      End If
      _cacheExpiration = YAPI.GetTickCount() + msValidity
      _serial = serial
      _funId = funcId
      _hwId = _serial + "." + _funId

      Try
        node = apires.getYJSONObject(funcId)
      Catch generatedExceptionName As Exception
        _throw(YAPI.IO_ERROR, Convert.ToString("unexpected JSON structure: missing function ") & funcId)
        load = YAPI.IO_ERROR
        Exit Function
      End Try

      _parse(node)
      load = YAPI_SUCCESS
    End Function


    '''*
    ''' <summary>
    '''   Invalidates the cache.
    ''' <para>
    '''   Invalidates the cache of the function attributes. Forces the
    '''   next call to get_xxx() or loadxxx() to use values that come from the device.
    ''' </para>
    ''' <para>
    ''' @noreturn
    ''' </para>
    ''' </summary>
    '''/
    Public Sub clearCache()

      Dim dev As YDevice = Nothing
      Dim errmsg As String = ""
      Dim res As Integer
      REM Resolve our reference to our device, load REST API
      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        Exit Sub
      End If
      dev.clearCache(False)
      If _cacheExpiration > 0 Then
        _cacheExpiration = YAPI.GetTickCount()
      End If
    End Sub

    '''*
    ''' <summary>
    '''   Gets the <c>YModule</c> object for the device on which the function is located.
    ''' <para>
    '''   If the function cannot be located on any module, the returned instance of
    '''   <c>YModule</c> is not shown as on-line.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an instance of <c>YModule</c>
    ''' </returns>
    '''/
    Public Function get_module() As YModule
      Dim fundescr As YFUN_DESCR
      Dim devdescr As YDEV_DESCR
      Dim errmsg As String = ""
      Dim serial As String = ""
      Dim funcId As String = ""
      Dim funcName As String = ""
      Dim funcValue As String = ""
      If (Not (_serial = "")) Then
        Return YModule.FindModule(_serial + ".module")
      End If
      fundescr = yapiGetFunction(_className, _func, errmsg)
      If (Not (YISERR(fundescr))) Then
        If (Not (YISERR(yapiGetFunctionInfo(fundescr, devdescr, serial, funcId, funcName, funcValue, errmsg)))) Then
          Return yFindModule(serial + ".module")

        End If
      End If

      REM return a true YModule object even if it is not a module valid for communicating
      Return yFindModule("module_of_" + _className + "_" + _func)
    End Function

    '''*
    ''' <summary>
    '''   Returns a unique identifier of type <c>YFUN_DESCR</c> corresponding to the function.
    ''' <para>
    '''   This identifier can be used to test if two instances of <c>YFunction</c> reference the same
    '''   physical function on the same physical device.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an identifier of type <c>YFUN_DESCR</c>.
    ''' </returns>
    ''' <para>
    '''   If the function has never been contacted, the returned value is <c>Y$CLASSNAME$.FUNCTIONDESCRIPTOR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_functionDescriptor() As YFUN_DESCR
      Return _fundescr
    End Function

    '''*
    ''' <summary>
    '''   Returns the value of the userData attribute, as previously stored using method
    '''   <c>set_userData</c>.
    ''' <para>
    '''   This attribute is never touched directly by the API, and is at disposal of the caller to
    '''   store a context.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the object stored previously by the caller.
    ''' </returns>
    '''/
    Public Function get_userData() As Object
      Return _userData
    End Function

    '''*
    ''' <summary>
    '''   Stores a user context provided as argument in the userData attribute of the function.
    ''' <para>
    '''   This attribute is never touched by the API, and is at disposal of the caller to store a context.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="data">
    '''   any kind of object to be stored
    ''' @noreturn
    ''' </param>
    '''/
    Public Sub set_userData(ByVal data As Object)
      _userData = data
    End Sub
  End Class

  REM --- (generated code: YFunction functions)

  '''*
  ''' <summary>
  '''   Retrieves a function for a given identifier.
  ''' <para>
  '''   The identifier can be specified using several formats:
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   - FunctionLogicalName
  ''' </para>
  ''' <para>
  '''   - ModuleSerialNumber.FunctionIdentifier
  ''' </para>
  ''' <para>
  '''   - ModuleSerialNumber.FunctionLogicalName
  ''' </para>
  ''' <para>
  '''   - ModuleLogicalName.FunctionIdentifier
  ''' </para>
  ''' <para>
  '''   - ModuleLogicalName.FunctionLogicalName
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   This function does not require that the function is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YFunction.isOnline()</c> to test if the function is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a function by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' <para>
  '''   If a call to this object's is_online() method returns FALSE although
  '''   you are certain that the matching device is plugged, make sure that you did
  '''   call registerHub() at application initialization time.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the function, for instance
  '''   <c>MyDevice.</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YFunction</c> object allowing you to drive the function.
  ''' </returns>
  '''/
  Public Function yFindFunction(ByVal func As String) As YFunction
    Return YFunction.FindFunction(func)
  End Function

  '''*
  ''' <summary>
  '''   A
  ''' <para>
  '''   lias for Y{$classname}.First{$classname}()
  ''' </para>
  ''' </summary>
  '''/
  Public Function yFirstFunction() As YFunction
    Return YFunction.FirstFunction()
  End Function


  REM --- (end of generated code: YFunction functions)


  REM --- (generated code: YModule class start)

  '''*
  ''' <summary>
  '''   The <c>YModule</c> class can be used with all Yoctopuce USB devices.
  ''' <para>
  '''   It can be used to control the module global parameters, and
  '''   to enumerate the functions provided by each module.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YModule
    Inherits YFunction
    REM --- (end of generated code: YModule class start)

    REM --- (generated code: YModule definitions)
    Public Const PRODUCTNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const SERIALNUMBER_INVALID As String = YAPI.INVALID_STRING
    Public Const PRODUCTID_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PRODUCTRELEASE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const FIRMWARERELEASE_INVALID As String = YAPI.INVALID_STRING
    Public Const PERSISTENTSETTINGS_LOADED As Integer = 0
    Public Const PERSISTENTSETTINGS_SAVED As Integer = 1
    Public Const PERSISTENTSETTINGS_MODIFIED As Integer = 2
    Public Const PERSISTENTSETTINGS_INVALID As Integer = -1
    Public Const LUMINOSITY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const BEACON_OFF As Integer = 0
    Public Const BEACON_ON As Integer = 1
    Public Const BEACON_INVALID As Integer = -1
    Public Const UPTIME_INVALID As Long = YAPI.INVALID_LONG
    Public Const USBCURRENT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const REBOOTCOUNTDOWN_INVALID As Integer = YAPI.INVALID_INT
    Public Const USERVAR_INVALID As Integer = YAPI.INVALID_INT
    REM --- (end of generated code: YModule definitions)

    REM --- (generated code: YModule attributes declaration)
    Protected _productName As String
    Protected _serialNumber As String
    Protected _productId As Integer
    Protected _productRelease As Integer
    Protected _firmwareRelease As String
    Protected _persistentSettings As Integer
    Protected _luminosity As Integer
    Protected _beacon As Integer
    Protected _upTime As Long
    Protected _usbCurrent As Integer
    Protected _rebootCountdown As Integer
    Protected _userVar As Integer
    Protected _valueCallbackModule As YModuleValueCallback
    Protected _logCallback As YModuleLogCallback
    Protected _confChangeCallback As YModuleConfigChangeCallback
    Protected _beaconCallback As YModuleBeaconCallback
    REM --- (end of generated code: YModule attributes declaration)
    Public Shared _moduleCallbackList As Dictionary(Of YModule, Integer) = New Dictionary(Of YModule, Integer)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _className = "Module"
      REM --- (generated code: YModule attributes initialization)
      _productName = PRODUCTNAME_INVALID
      _serialNumber = SERIALNUMBER_INVALID
      _productId = PRODUCTID_INVALID
      _productRelease = PRODUCTRELEASE_INVALID
      _firmwareRelease = FIRMWARERELEASE_INVALID
      _persistentSettings = PERSISTENTSETTINGS_INVALID
      _luminosity = LUMINOSITY_INVALID
      _beacon = BEACON_INVALID
      _upTime = UPTIME_INVALID
      _usbCurrent = USBCURRENT_INVALID
      _rebootCountdown = REBOOTCOUNTDOWN_INVALID
      _userVar = USERVAR_INVALID
      _valueCallbackModule = Nothing
      _logCallback = Nothing
      _confChangeCallback = Nothing
      REM --- (end of generated code: YModule attributes initialization)
    End Sub


    Shared Sub _updateModuleCallbackList(ByVal modul As YModule, ByVal ad As Boolean)

      If (ad) Then
        modul.isOnline()
        If (Not _moduleCallbackList.ContainsKey(modul)) Then
          _moduleCallbackList(modul) = 1
        Else
          _moduleCallbackList(modul) += 1
        End If
      Else
        If (_moduleCallbackList.ContainsKey(modul) And _moduleCallbackList(modul) > 1) Then
          _moduleCallbackList(modul) -= 1
        End If
      End If
    End Sub


    REM --- (generated code: YModule private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("productName") Then
        _productName = json_val.getString("productName")
      End If
      If json_val.has("serialNumber") Then
        _serialNumber = json_val.getString("serialNumber")
      End If
      If json_val.has("productId") Then
        _productId = CInt(json_val.getLong("productId"))
      End If
      If json_val.has("productRelease") Then
        _productRelease = CInt(json_val.getLong("productRelease"))
      End If
      If json_val.has("firmwareRelease") Then
        _firmwareRelease = json_val.getString("firmwareRelease")
      End If
      If json_val.has("persistentSettings") Then
        _persistentSettings = CInt(json_val.getLong("persistentSettings"))
      End If
      If json_val.has("luminosity") Then
        _luminosity = CInt(json_val.getLong("luminosity"))
      End If
      If json_val.has("beacon") Then
        If (json_val.getInt("beacon") > 0) Then _beacon = 1 Else _beacon = 0
      End If
      If json_val.has("upTime") Then
        _upTime = json_val.getLong("upTime")
      End If
      If json_val.has("usbCurrent") Then
        _usbCurrent = CInt(json_val.getLong("usbCurrent"))
      End If
      If json_val.has("rebootCountdown") Then
        _rebootCountdown = CInt(json_val.getLong("rebootCountdown"))
      End If
      If json_val.has("userVar") Then
        _userVar = CInt(json_val.getLong("userVar"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YModule private methods declaration)


    Public Overloads Function get_friendlyName() As String
      Dim retcode As YRETCODE
      Dim fundesc As YFUN_DESCR = 0
      Dim devdesc As YDEV_DESCR = 0
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim snum As String = ""
      Dim funcid As String = ""
      REM Resolve the function name
      retcode = _getDescriptor(fundesc, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_friendlyName = YAPI.FRIENDLYNAME_INVALID
        Exit Function
      End If

      retcode = yapiGetFunctionInfo(fundesc, devdesc, snum, funcid, funcName, funcVal, errmsg)
      If (YISERR(retcode)) Then
        _throw(retcode, errmsg)
        get_friendlyName = YAPI.FRIENDLYNAME_INVALID
        Exit Function
      End If
      If funcName <> "" Then
        get_friendlyName = funcName
      Else
        get_friendlyName = snum
      End If
    End Function


    Public Function get_logicalName_internal() As String
      Return _logicalName
    End Function


    Public Sub setImmutableAttributes(ByRef infos As yDeviceSt)
      _serialNumber = infos.serial
      _productName = infos.productname
      _productId = infos.deviceid
    End Sub

    REM Return the properties of the nth function of our device
    Private Function _getFunction(ByVal idx As Integer, ByRef serial As String, ByRef funcId As String, ByRef baseType As String,
                              ByRef funcName As String, ByRef funcVal As String, ByRef errmsg As String) As YRETCODE


      Dim functions As List(Of yu32) = Nothing
      Dim dev As YDevice = Nothing
      Dim res As Integer
      Dim fundescr As YFUN_DESCR
      Dim devdescr As YDEV_DESCR

      REM retrieve device object
      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        _getFunction = res
        Exit Function
      End If


      REM get reference to all functions from the device
      res = dev.getFunctions(functions, errmsg)
      If (YISERR(res)) Then
        _getFunction = res
        Exit Function
      End If

      REM get latest function info from yellow pages
      fundescr = CInt(functions(idx))

      res = yapiGetFunctionInfoEx(fundescr, devdescr, serial, funcId, baseType, funcName, funcVal, errmsg)
      If (YISERR(res)) Then
        _getFunction = res
        Exit Function
      End If

      _getFunction = YAPI_SUCCESS
    End Function


    '''*
    ''' <summary>
    '''   Returns the number of functions (beside the "module" interface) available on the module.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the number of functions on the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/

    Public Function functionCount() As Integer
      Dim functions As List(Of yu32) = Nothing
      Dim dev As YDevice = Nothing
      Dim errmsg As String = ""
      Dim res As Integer

      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        functionCount = res
        Exit Function
      End If

      res = dev.getFunctions(functions, errmsg)
      If (YISERR(res)) Then
        functions = Nothing
        _throw(res, errmsg)
        functionCount = res
        Exit Function
      End If

      functionCount = functions.Count
    End Function


    '''*
    ''' <summary>
    '''   Retrieves the hardware identifier of the <i>n</i>th function on the module.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="functionIndex">
    '''   the index of the function for which the information is desired, starting at 0 for the first function.
    ''' </param>
    ''' <returns>
    '''   a string corresponding to the unambiguous hardware identifier of the requested module function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/

    Public Function functionId(ByVal functionIndex As Integer) As String
      Dim serial As String = ""
      Dim funcId As String = ""
      Dim baseType As String = ""
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res As Integer
      res = _getFunction(functionIndex, serial, funcId, baseType, funcName, funcVal, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        functionId = YAPI.INVALID_STRING
        Exit Function
      End If
      functionId = funcId
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the base type of the <i>n</i>th function on the module.
    ''' <para>
    '''   For instance, the base type of all measuring functions is "Sensor".
    ''' </para>
    ''' </summary>
    ''' <param name="functionIndex">
    '''   the index of the function for which the information is desired, starting at 0 for the first function.
    ''' </param>
    ''' <returns>
    '''   a string corresponding to the base type of the function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/

    Public Function functionBaseType(ByVal functionIndex As Integer) As String
      Dim serial As String = ""
      Dim funcId As String = ""
      Dim funcName As String = ""
      Dim baseType As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res As Integer

      res = _getFunction(functionIndex, serial, funcId, baseType, funcName, funcVal, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        Return YAPI.INVALID_STRING
      End If
      Return baseType
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the type of the <i>n</i>th function on the module.
    ''' <para>
    '''   Yoctopuce functions type names match their class names without the <i>Y</i> prefix, for instance
    '''   <i>Relay</i>, <i>Temperature</i> etc..
    ''' </para>
    ''' </summary>
    ''' <param name="functionIndex">
    '''   the index of the function for which the information is desired, starting at 0 for the first function.
    ''' </param>
    ''' <returns>
    '''   a string corresponding to the type of the function.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/

    Public Function functionType(ByVal functionIndex As Integer) As String
      Dim serial As String = ""
      Dim funcId As String = ""
      Dim baseType As String = ""
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res As Integer
      Dim first As Char
      Dim i As Integer

      res = _getFunction(functionIndex, serial, funcId, baseType, funcName, funcVal, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        Return YAPI.INVALID_STRING
      End If
      first = funcId(0)
      i = funcId.Length
      While (i > 0)
        If (Char.IsLetter(funcId(i - 1))) Then
          Exit While
        End If
        i -= 1
      End While
      Return Char.ToUpper(first) + funcId.Substring(1, i - 1)
    End Function

    '''*
    ''' <summary>
    '''   Retrieves the logical name of the <i>n</i>th function on the module.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="functionIndex">
    '''   the index of the function for which the information is desired, starting at 0 for the first function.
    ''' </param>
    ''' <returns>
    '''   a string corresponding to the logical name of the requested module function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/

    Public Function functionName(ByVal functionIndex As Integer) As String
      Dim serial As String = ""
      Dim funcId As String = ""
      Dim baseType As String = ""
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res As Integer

      res = _getFunction(functionIndex, serial, funcId, baseType, funcName, funcVal, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        functionName = YAPI.INVALID_STRING
        Exit Function
      End If

      functionName = funcName
    End Function


    '''*
    ''' <summary>
    '''   Retrieves the advertised value of the <i>n</i>th function on the module.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="functionIndex">
    '''   the index of the function for which the information is desired, starting at 0 for the first function.
    ''' </param>
    ''' <returns>
    '''   a short string (up to 6 characters) corresponding to the advertised value of the requested module function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/

    Public Function functionValue(ByVal functionIndex As Integer) As String
      Dim serial As String = ""
      Dim funcId As String = ""
      Dim baseType As String = ""
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res As Integer

      res = _getFunction(functionIndex, serial, funcId, baseType, funcName, funcVal, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        functionValue = YAPI.INVALID_STRING
        Exit Function
      End If
      functionValue = funcVal
    End Function


    REM --- (generated code: YModule public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the commercial name of the module, as set by the factory.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the commercial name of the module, as set by the factory
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.PRODUCTNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_productName() As String
      Dim res As String
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PRODUCTNAME_INVALID
        End If
      End If
      res = Me._productName
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the serial number of the module, as set by the factory.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the serial number of the module, as set by the factory
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns YModule.SERIALNUMBER_INVALID.
    ''' </para>
    '''/
    Public Overrides Function get_serialNumber() As String
      Dim res As String
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SERIALNUMBER_INVALID
        End If
      End If
      res = Me._serialNumber
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the USB device identifier of the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the USB device identifier of the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.PRODUCTID_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_productId() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PRODUCTID_INVALID
        End If
      End If
      res = Me._productId
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the release number of the module hardware, preprogrammed at the factory.
    ''' <para>
    '''   The original hardware release returns value 1, revision B returns value 2, etc.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the release number of the module hardware, preprogrammed at the factory
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.PRODUCTRELEASE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_productRelease() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PRODUCTRELEASE_INVALID
        End If
      End If
      res = Me._productRelease
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the version of the firmware embedded in the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the version of the firmware embedded in the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.FIRMWARERELEASE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_firmwareRelease() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return FIRMWARERELEASE_INVALID
        End If
      End If
      res = Me._firmwareRelease
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current state of persistent module settings.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YModule.PERSISTENTSETTINGS_LOADED</c>, <c>YModule.PERSISTENTSETTINGS_SAVED</c> and
    '''   <c>YModule.PERSISTENTSETTINGS_MODIFIED</c> corresponding to the current state of persistent module settings
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.PERSISTENTSETTINGS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_persistentSettings() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PERSISTENTSETTINGS_INVALID
        End If
      End If
      res = Me._persistentSettings
      Return res
    End Function


    Public Function set_persistentSettings(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("persistentSettings", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the luminosity of the  module informative LEDs (from 0 to 100).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the luminosity of the  module informative LEDs (from 0 to 100)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.LUMINOSITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_luminosity() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LUMINOSITY_INVALID
        End If
      End If
      res = Me._luminosity
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the luminosity of the module informative leds.
    ''' <para>
    '''   The parameter is a
    '''   value between 0 and 100.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the luminosity of the module informative leds
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_luminosity(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("luminosity", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the state of the localization beacon.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YModule.BEACON_OFF</c> or <c>YModule.BEACON_ON</c>, according to the state of the localization beacon
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.BEACON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_beacon() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BEACON_INVALID
        End If
      End If
      res = Me._beacon
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Turns on or off the module localization beacon.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YModule.BEACON_OFF</c> or <c>YModule.BEACON_ON</c>
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_beacon(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("beacon", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of milliseconds spent since the module was powered on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of milliseconds spent since the module was powered on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.UPTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_upTime() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return UPTIME_INVALID
        End If
      End If
      res = Me._upTime
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current consumed by the module on the USB bus, in milli-amps.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current consumed by the module on the USB bus, in milli-amps
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.USBCURRENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_usbCurrent() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return USBCURRENT_INVALID
        End If
      End If
      res = Me._usbCurrent
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the remaining number of seconds before the module restarts, or zero when no
    '''   reboot has been scheduled.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the remaining number of seconds before the module restarts, or zero when no
    '''   reboot has been scheduled
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.REBOOTCOUNTDOWN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rebootCountdown() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return REBOOTCOUNTDOWN_INVALID
        End If
      End If
      res = Me._rebootCountdown
      Return res
    End Function


    Public Function set_rebootCountdown(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("rebootCountdown", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the value previously stored in this attribute.
    ''' <para>
    '''   On startup and after a device reboot, the value is always reset to zero.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the value previously stored in this attribute
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YModule.USERVAR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_userVar() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return USERVAR_INVALID
        End If
      End If
      res = Me._userVar
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Stores a 32 bit value in the device RAM.
    ''' <para>
    '''   This attribute is at programmer disposal,
    '''   should he need to store a state variable.
    '''   On startup and after a device reboot, the value is always reset to zero.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_userVar(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("userVar", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Allows you to find a module from its serial number or from its logical name.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   This function does not require that the module is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YModule.isOnline()</c> to test if the module is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a module by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   If a call to this object's is_online() method returns FALSE although
    '''   you are certain that the device is plugged, make sure that you did
    '''   call registerHub() at application initialization time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string containing either the serial number or
    '''   the logical name of the desired module
    ''' </param>
    ''' <returns>
    '''   a <c>YModule</c> object allowing you to drive the module
    '''   or get additional information on the module.
    ''' </returns>
    '''/
    Public Shared Function FindModule(func As String) As YModule
      Dim obj As YModule
      Dim cleanHwId As String
      Dim modpos As Integer = 0
      cleanHwId = func
      modpos = func.IndexOf(".module")
      If (modpos <> ((func).Length - 7)) Then
        cleanHwId = func + ".module"
      End If
      obj = CType(YFunction._FindFromCache("Module", cleanHwId), YModule)
      If ((obj Is Nothing)) Then
        obj = New YModule(cleanHwId)
        YFunction._AddToCache("Module", cleanHwId, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerValueCallback(callback As YModuleValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackModule = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) AndAlso Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overrides Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackModule Is Nothing)) Then
        Me._valueCallbackModule(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function get_productNameAndRevision() As String
      Dim prodname As String
      Dim prodrel As Integer = 0
      Dim fullname As String

      prodname = Me.get_productName()
      prodrel = Me.get_productRelease()
      If (prodrel > 1) Then
        fullname = "" + prodname + " rev. " + Chr(64 + prodrel)
      Else
        fullname = prodname
      End If
      Return fullname
    End Function

    '''*
    ''' <summary>
    '''   Saves current settings in the nonvolatile memory of the module.
    ''' <para>
    '''   Warning: the number of allowed save operations during a module life is
    '''   limited (about 100000 cycles). Do not call this function within a loop.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function saveToFlash() As Integer
      Return Me.set_persistentSettings(PERSISTENTSETTINGS_SAVED)
    End Function

    '''*
    ''' <summary>
    '''   Reloads the settings stored in the nonvolatile memory, as
    '''   when the module is powered on.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function revertFromFlash() As Integer
      Return Me.set_persistentSettings(PERSISTENTSETTINGS_LOADED)
    End Function

    '''*
    ''' <summary>
    '''   Schedules a simple module reboot after the given number of seconds.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="secBeforeReboot">
    '''   number of seconds before rebooting
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function reboot(secBeforeReboot As Integer) As Integer
      Return Me.set_rebootCountdown(secBeforeReboot)
    End Function

    '''*
    ''' <summary>
    '''   Schedules a module reboot into special firmware update mode.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="secBeforeReboot">
    '''   number of seconds before rebooting
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function triggerFirmwareUpdate(secBeforeReboot As Integer) As Integer
      Return Me.set_rebootCountdown(-secBeforeReboot)
    End Function

    Public Overridable Sub _startStopDevLog(serial As String, start As Boolean)
      Dim i_start As Integer = 0
      If (start) Then
        i_start = 1
      Else
        i_start = 0
      End If

      _yapiStartStopDeviceLogCallback(New StringBuilder(serial), i_start)
    End Sub

    '''*
    ''' <summary>
    '''   Registers a device log callback function.
    ''' <para>
    '''   This callback will be called each time
    '''   that a module sends a new log message. Mostly useful to debug a Yoctopuce module.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer.
    '''   The callback function should take two
    '''   arguments: the module object that emitted the log message,
    '''   and the character string containing the log.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </param>
    '''/
    Public Overridable Function registerLogCallback(callback As YModuleLogCallback) As Integer
      Dim serial As String

      serial = Me.get_serialNumber()
      If (serial = YAPI.INVALID_STRING) Then
        Return YAPI.DEVICE_NOT_FOUND
      End If
      Me._logCallback = callback
      Me._startStopDevLog(serial, Not (callback Is Nothing))
      Return 0
    End Function

    Public Overridable Function get_logCallback() As YModuleLogCallback
      Return Me._logCallback
    End Function

    '''*
    ''' <summary>
    '''   Register a callback function, to be called when a persistent settings in
    '''   a device configuration has been changed (e.g.
    ''' <para>
    '''   change of unit, etc).
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   a procedure taking a YModule parameter, or <c>Nothing</c>
    '''   to unregister a previously registered  callback.
    ''' </param>
    '''/
    Public Overridable Function registerConfigChangeCallback(callback As YModuleConfigChangeCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YModule._updateModuleCallbackList(Me, True)
      Else
        YModule._updateModuleCallbackList(Me, False)
      End If
      Me._confChangeCallback = callback
      Return 0
    End Function

    Public Overridable Function _invokeConfigChangeCallback() As Integer
      If (Not (Me._confChangeCallback Is Nothing)) Then
        Me._confChangeCallback(Me)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Register a callback function, to be called when the localization beacon of the module
    '''   has been changed.
    ''' <para>
    '''   The callback function should take two arguments: the YModule object of
    '''   which the beacon has changed, and an integer describing the new beacon state.
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   The callback function to call, or <c>Nothing</c> to unregister a
    '''   previously registered callback.
    ''' </param>
    '''/
    Public Overridable Function registerBeaconCallback(callback As YModuleBeaconCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YModule._updateModuleCallbackList(Me, True)
      Else
        YModule._updateModuleCallbackList(Me, False)
      End If
      Me._beaconCallback = callback
      Return 0
    End Function

    Public Overridable Function _invokeBeaconCallback(beaconState As Integer) As Integer
      If (Not (Me._beaconCallback Is Nothing)) Then
        Me._beaconCallback(Me, beaconState)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Triggers a configuration change callback, to check if they are supported or not.
    ''' <para>
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function triggerConfigChangeCallback() As Integer
      Me._setAttr("persistentSettings", "2")
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Tests whether the byn file is valid for this module.
    ''' <para>
    '''   This method is useful to test if the module needs to be updated.
    '''   It is possible to pass a directory as argument instead of a file. In this case, this method returns
    '''   the path of the most recent
    '''   appropriate <c>.byn</c> file. If the parameter <c>onlynew</c> is true, the function discards
    '''   firmwares that are older or
    '''   equal to the installed firmware.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="path">
    '''   the path of a byn file or a directory that contains byn files
    ''' </param>
    ''' <param name="onlynew">
    '''   returns only files that are strictly newer
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   the path of the byn file to use or a empty string if no byn files matches the requirement
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a string that start with "error:".
    ''' </para>
    '''/
    Public Overridable Function checkFirmware(path As String, onlynew As Boolean) As String
      Dim serial As String
      Dim release As Integer = 0
      Dim tmp_res As String
      If (onlynew) Then
        release = YAPI._atoi(Me.get_firmwareRelease())
      Else
        release = 0
      End If
      REM //may throw an exception
      serial = Me.get_serialNumber()
      tmp_res = YFirmwareUpdate.CheckFirmware(serial, path, release)
      If (tmp_res.IndexOf("error:") = 0) Then
        Me._throw(YAPI.INVALID_ARGUMENT, tmp_res)
      End If
      Return tmp_res
    End Function

    '''*
    ''' <summary>
    '''   Prepares a firmware update of the module.
    ''' <para>
    '''   This method returns a <c>YFirmwareUpdate</c> object which
    '''   handles the firmware update process.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="path">
    '''   the path of the <c>.byn</c> file to use.
    ''' </param>
    ''' <param name="force">
    '''   true to force the firmware update even if some prerequisites appear not to be met
    ''' </param>
    ''' <returns>
    '''   a <c>YFirmwareUpdate</c> object or Nothing on error.
    ''' </returns>
    '''/
    Public Overridable Function updateFirmwareEx(path As String, force As Boolean) As YFirmwareUpdate
      Dim serial As String
      Dim settings As Byte() = New Byte(){}

      serial = Me.get_serialNumber()
      settings = Me.get_allSettings()
      If ((settings).Length = 0) Then
        Me._throw(YAPI.IO_ERROR, "Unable to get device settings")
        settings = YAPI.DefaultEncoding.GetBytes("error:Unable to get device settings")
      End If

      Return New YFirmwareUpdate(serial, path, settings, force)
    End Function

    '''*
    ''' <summary>
    '''   Prepares a firmware update of the module.
    ''' <para>
    '''   This method returns a <c>YFirmwareUpdate</c> object which
    '''   handles the firmware update process.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="path">
    '''   the path of the <c>.byn</c> file to use.
    ''' </param>
    ''' <returns>
    '''   a <c>YFirmwareUpdate</c> object or Nothing on error.
    ''' </returns>
    '''/
    Public Overridable Function updateFirmware(path As String) As YFirmwareUpdate
      Return Me.updateFirmwareEx(path, False)
    End Function

    '''*
    ''' <summary>
    '''   Returns all the settings and uploaded files of the module.
    ''' <para>
    '''   Useful to backup all the
    '''   logical names, calibrations parameters, and uploaded files of a device.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a binary buffer with all the settings.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an binary object of size 0.
    ''' </para>
    '''/
    Public Overridable Function get_allSettings() As Byte()
      Dim settings As Byte() = New Byte(){}
      Dim json As Byte() = New Byte(){}
      Dim res As Byte() = New Byte(){}
      Dim sep As String
      Dim name As String
      Dim item As String
      Dim t_type As String
      Dim pageid As String
      Dim url As String
      Dim file_data As String
      Dim file_data_bin As Byte() = New Byte(){}
      Dim temp_data_bin As Byte() = New Byte(){}
      Dim ext_settings As String
      Dim filelist As List(Of Byte()) = New List(Of Byte())()
      Dim templist As List(Of String) = New List(Of String)()

      settings = Me._download("api.json")
      If ((settings).Length = 0) Then
        Return settings
      End If
      ext_settings = ", ""extras"":["
      templist = Me.get_functionIds("Temperature")
      sep = ""
      Dim ii_0 As Integer
      For ii_0 = 0 To templist.Count - 1
        If (YAPI._atoi(Me.get_firmwareRelease()) > 9000) Then
          url = "api/" + templist(ii_0) + "/sensorType"
          t_type = YAPI.DefaultEncoding.GetString(Me._download(url))
          If (t_type = "RES_NTC" OrElse t_type = "RES_LINEAR") Then
            pageid = (templist(ii_0)).Substring(11, (templist(ii_0)).Length - 11)
            If (pageid = "") Then
              pageid = "1"
            End If
            temp_data_bin = Me._download("extra.json?page=" + pageid)
            If ((temp_data_bin).Length > 0) Then
              item = "" + sep + "{""fid"":""" + templist(ii_0) + """, ""json"":" + YAPI.DefaultEncoding.GetString(temp_data_bin) + "}" + vbLf + ""
              ext_settings = ext_settings + item
              sep = ","
            End If
          End If
        End If
      Next ii_0
      ext_settings = ext_settings + "]," + vbLf + """files"":["
      If (Me.hasFunction("files")) Then
        json = Me._download("files.json?a=dir&f=")
        If ((json).Length = 0) Then
          Return json
        End If
        filelist = Me._json_get_array(json)
        sep = ""
        Dim ii_1 As Integer
        For ii_1 = 0 To filelist.Count - 1
          name = Me._json_get_key(filelist(ii_1), "name")
          If (((name).Length > 0) AndAlso Not (name = "startupConf.json")) Then
            file_data_bin = Me._download(Me._escapeAttr(name))
            file_data = YAPI._bytesToHexStr(file_data_bin, 0, file_data_bin.Length)
            item = "" + sep + "{""name"":""" + name + """, ""data"":""" + file_data + """}" + vbLf + ""
            ext_settings = ext_settings + item
            sep = ","
          End If
        Next ii_1
      End If
      res = YAPI.DefaultEncoding.GetBytes("{ ""api"":" + YAPI.DefaultEncoding.GetString(settings) + ext_settings + "]}")
      Return res
    End Function

    Public Overridable Function loadThermistorExtra(funcId As String, jsonExtra As String) As Integer
      Dim values As List(Of Byte()) = New List(Of Byte())()
      Dim url As String
      Dim curr As String
      Dim binCurr As Byte() = New Byte(){}
      Dim currTemp As String
      Dim binCurrTemp As Byte() = New Byte(){}
      Dim ofs As Integer = 0
      Dim size As Integer = 0
      url = "api/" + funcId + ".json?command=Z"

      Me._download(url)
      REM // add records in growing resistance value
      values = Me._json_get_array(YAPI.DefaultEncoding.GetBytes(jsonExtra))
      ofs = 0
      size = values.Count
      While (ofs + 1 < size)
        binCurr = values(ofs)
        binCurrTemp = values(ofs + 1)
        curr = Me._json_get_string(binCurr)
        currTemp = Me._json_get_string(binCurrTemp)
        url = "api/" + funcId + ".json?command=m" + curr + ":" + currTemp
        Me._download(url)
        ofs = ofs + 2
      End While
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_extraSettings(jsonExtra As String) As Integer
      Dim extras As List(Of Byte()) = New List(Of Byte())()
      Dim tmp As Byte() = New Byte(){}
      Dim functionId As String
      Dim data As Byte() = New Byte(){}
      extras = Me._json_get_array(YAPI.DefaultEncoding.GetBytes(jsonExtra))
      Dim ii_0 As Integer
      For ii_0 = 0 To extras.Count - 1
        tmp = Me._get_json_path(extras(ii_0), "fid")
        functionId = Me._json_get_string(tmp)
        data = Me._get_json_path(extras(ii_0), "json")
        If (Me.hasFunction(functionId)) Then
          Me.loadThermistorExtra(functionId, YAPI.DefaultEncoding.GetString(data))
        End If
      Next ii_0
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Restores all the settings and uploaded files to the module.
    ''' <para>
    '''   This method is useful to restore all the logical names and calibrations parameters,
    '''   uploaded files etc. of a device from a backup.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modifications must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="settings">
    '''   a binary buffer with all the settings.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_allSettingsAndFiles(settings As Byte()) As Integer
      Dim down As Byte() = New Byte(){}
      Dim json_bin As Byte() = New Byte(){}
      Dim json_api As Byte() = New Byte(){}
      Dim json_files As Byte() = New Byte(){}
      Dim json_extra As Byte() = New Byte(){}
      Dim fuperror As Integer = 0
      Dim globalres As Integer = 0
      fuperror = 0
      json_api = Me._get_json_path(settings, "api")
      If ((json_api).Length = 0) Then
        Return Me.set_allSettings(settings)
      End If
      json_extra = Me._get_json_path(settings, "extras")
      If ((json_extra).Length > 0) Then
        Me.set_extraSettings(YAPI.DefaultEncoding.GetString(json_extra))
      End If
      Me.set_allSettings(json_api)
      If (Me.hasFunction("files")) Then
        Dim files As List(Of Byte()) = New List(Of Byte())()
        Dim res As String
        Dim tmp As Byte() = New Byte(){}
        Dim name As String
        Dim data As String
        down = Me._download("files.json?a=format")
        down = Me._get_json_path(down, "res")
        res = Me._json_get_string(down)
        If Not(res = "ok") Then
          me._throw(YAPI.IO_ERROR, "format failed")
          return YAPI.IO_ERROR
        end if
        json_files = Me._get_json_path(settings, "files")
        files = Me._json_get_array(json_files)
        Dim ii_0 As Integer
        For ii_0 = 0 To files.Count - 1
          tmp = Me._get_json_path(files(ii_0), "name")
          name = Me._json_get_string(tmp)
          tmp = Me._get_json_path(files(ii_0), "data")
          data = Me._json_get_string(tmp)
          If (name = "") Then
            fuperror = fuperror + 1
          Else
            Me._upload(name, YAPI._hexStrToBin(data))
          End If
        Next ii_0
      End If
      REM // Apply settings a second time for file-dependent settings and dynamic sensor nodes
      globalres = Me.set_allSettings(json_api)
      If Not(fuperror = 0) Then
        me._throw(YAPI.IO_ERROR, "Error during file upload")
        return YAPI.IO_ERROR
      end if
      Return globalres
    End Function

    '''*
    ''' <summary>
    '''   Tests if the device includes a specific function.
    ''' <para>
    '''   This method takes a function identifier
    '''   and returns a boolean.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="funcId">
    '''   the requested function identifier
    ''' </param>
    ''' <returns>
    '''   true if the device has the function identifier
    ''' </returns>
    '''/
    Public Overridable Function hasFunction(funcId As String) As Boolean
      Dim count As Integer = 0
      Dim i As Integer = 0
      Dim fid As String

      count = Me.functionCount()
      i = 0
      While (i < count)
        fid = Me.functionId(i)
        If (fid = funcId) Then
          Return True
        End If
        i = i + 1
      End While
      Return False
    End Function

    '''*
    ''' <summary>
    '''   Retrieve all hardware identifier that match the type passed in argument.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="funType">
    '''   The type of function (Relay, LightSensor, Voltage,...)
    ''' </param>
    ''' <returns>
    '''   an array of strings.
    ''' </returns>
    '''/
    Public Overridable Function get_functionIds(funType As String) As List(Of String)
      Dim count As Integer = 0
      Dim i As Integer = 0
      Dim ftype As String
      Dim res As List(Of String) = New List(Of String)()

      count = Me.functionCount()
      i = 0

      While (i < count)
        ftype = Me.functionType(i)
        If (ftype = funType) Then
          res.Add(Me.functionId(i))
        Else
          ftype = Me.functionBaseType(i)
          If (ftype = funType) Then
            res.Add(Me.functionId(i))
          End If
        End If
        i = i + 1
      End While

      Return res
    End Function

    Public Overridable Function _flattenJsonStruct(jsoncomplex As Byte()) As Byte()
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim smallbuff As StringBuilder = New StringBuilder(1024)
      Dim bigbuff As StringBuilder
      Dim buffsize As Integer = 0
      Dim fullsize As Integer
      Dim res As Integer = 0
      Dim jsonflat As String
      Dim jsoncomplexstr As String
      fullsize = 0
      jsoncomplexstr = YAPI.DefaultEncoding.GetString(jsoncomplex)
      res = _yapiGetAllJsonKeys(New StringBuilder(jsoncomplexstr), smallbuff, 1024, fullsize, errmsg)
      If (res < 0) Then
        Me._throw(YAPI.INVALID_ARGUMENT, errmsg.ToString())
        jsonflat = "error:" + errmsg.ToString()
        Return YAPI.DefaultEncoding.GetBytes(jsonflat)
      End If
      If (fullsize <= 1024) Then
        jsonflat = smallbuff.ToString()
      Else
        fullsize = fullsize * 2
        buffsize = fullsize
        bigbuff = New StringBuilder(buffsize)
        res = _yapiGetAllJsonKeys(New StringBuilder(jsoncomplexstr), bigbuff, buffsize, fullsize, errmsg)
        If (res < 0) Then
          Me._throw(YAPI.INVALID_ARGUMENT, errmsg.ToString())
          jsonflat = "error:" + errmsg.ToString()
        Else
          jsonflat = bigbuff.ToString()
        End If
        bigbuff = Nothing
      End If
      Return YAPI.DefaultEncoding.GetBytes(jsonflat)
    End Function

    Public Overridable Function calibVersion(cparams As String) As Integer
      If (cparams = "0,") Then
        Return 3
      End If
      If (cparams.IndexOf(",") >= 0) Then
        If (cparams.IndexOf(" ") > 0) Then
          Return 3
        Else
          Return 1
        End If
      End If
      If (cparams = "" OrElse cparams = "0") Then
        Return 1
      End If
      If (((cparams).Length < 2) OrElse (cparams.IndexOf(".") >= 0)) Then
        Return 0
      Else
        Return 2
      End If
    End Function

    Public Overridable Function calibScale(unit_name As String, sensorType As String) As Integer
      If (unit_name = "g" OrElse unit_name = "gauss" OrElse unit_name = "W") Then
        Return 1000
      End If
      If (unit_name = "C") Then
        If (sensorType = "") Then
          Return 16
        End If
        If (YAPI._atoi(sensorType) < 8) Then
          Return 16
        Else
          Return 100
        End If
      End If
      If (unit_name = "m" OrElse unit_name = "deg") Then
        Return 10
      End If
      Return 1
    End Function

    Public Overridable Function calibOffset(unit_name As String) As Integer
      If (unit_name = "% RH" OrElse unit_name = "mbar" OrElse unit_name = "lx") Then
        Return 0
      End If
      Return 32767
    End Function

    Public Overridable Function calibConvert(param As String, currentFuncValue As String, unit_name As String, sensorType As String) As String
      Dim paramVer As Integer = 0
      Dim funVer As Integer = 0
      Dim funScale As Integer = 0
      Dim funOffset As Integer = 0
      Dim paramScale As Integer = 0
      Dim paramOffset As Integer = 0
      Dim words As List(Of Integer) = New List(Of Integer)()
      Dim words_str As List(Of String) = New List(Of String)()
      Dim calibData As List(Of Double) = New List(Of Double)()
      Dim iCalib As List(Of Integer) = New List(Of Integer)()
      Dim calibType As Integer = 0
      Dim i As Integer = 0
      Dim maxSize As Integer = 0
      Dim ratio As Double = 0
      Dim nPoints As Integer = 0
      Dim wordVal As Double = 0
      REM // Initial guess for parameter encoding
      paramVer = Me.calibVersion(param)
      funVer = Me.calibVersion(currentFuncValue)
      funScale = Me.calibScale(unit_name, sensorType)
      funOffset = Me.calibOffset(unit_name)
      paramScale = funScale
      paramOffset = funOffset
      If (funVer < 3) Then
        REM // Read the effective device scale if available
        If (funVer = 2) Then
          words = YAPI._decodeWords(currentFuncValue)
          If ((words(0) = 1366) AndAlso (words(1) = 12500)) Then
            REM // Yocto-3D RefFrame used a special encoding
            funScale = 1
            funOffset = 0
          Else
            funScale = words(1)
            funOffset = words(0)
          End If
        Else
          If (funVer = 1) Then
            If (currentFuncValue = "" OrElse (YAPI._atoi(currentFuncValue) > 10)) Then
              funScale = 0
            End If
          End If
        End If
      End If
      calibData.Clear()
      calibType = 0
      If (paramVer < 3) Then
        REM // Handle old 16 bit parameters formats
        If (paramVer = 2) Then
          words = YAPI._decodeWords(param)
          If ((words(0) = 1366) AndAlso (words(1) = 12500)) Then
            REM // Yocto-3D RefFrame used a special encoding
            paramScale = 1
            paramOffset = 0
          Else
            paramScale = words(1)
            paramOffset = words(0)
          End If
          If ((words.Count >= 3) AndAlso (words(2) > 0)) Then
            maxSize = 3 + 2 * ((words(2)) Mod (10))
            If (maxSize > words.Count) Then
              maxSize = words.Count
            End If
            i = 3
            While (i < maxSize)
              calibData.Add(CType(words(i), Double))
              i = i + 1
            End While
          End If
        Else
          If (paramVer = 1) Then
            words_str = New List(Of String)(param.Split(New Char() {","c}))
            Dim ii_0 As Integer
            For ii_0 = 0 To words_str.Count - 1
              words.Add(YAPI._atoi(words_str(ii_0)))
            Next ii_0
            If (param = "" OrElse (words(0) > 10)) Then
              paramScale = 0
            End If
            If ((words.Count > 0) AndAlso (words(0) > 0)) Then
              maxSize = 1 + 2 * ((words(0)) Mod (10))
              If (maxSize > words.Count) Then
                maxSize = words.Count
              End If
              i = 1
              While (i < maxSize)
                calibData.Add(CType(words(i), Double))
                i = i + 1
              End While
            End If
          Else
            If (paramVer = 0) Then
              ratio = YAPI._atof(param)
              If (ratio > 0) Then
                calibData.Add(0.0)
                calibData.Add(0.0)
                calibData.Add(Math.Round(65535 / ratio))
                calibData.Add(65535.0)
              End If
            End If
          End If
        End If
        i = 0
        While (i < calibData.Count)
          If (paramScale > 0) Then
            REM // scalar decoding
            calibData(i) = (calibData(i) - paramOffset) / paramScale
          Else
            REM // floating-point decoding
            calibData(i) = YAPI._decimalToDouble(CType(Math.Round(calibData(i)), Integer))
          End If
          i = i + 1
        End While
      Else
        REM // Handle latest 32bit parameter format
        iCalib = YAPI._decodeFloats(param)
        calibType = CType(Math.Round(iCalib(0) / 1000.0), Integer)
        If (calibType >= 30) Then
          calibType = calibType - 30
        End If
        i = 1
        While (i < iCalib.Count)
          calibData.Add(iCalib(i) / 1000.0)
          i = i + 1
        End While
      End If
      If (funVer >= 3) Then
        REM // Encode parameters in new format
        If (calibData.Count = 0) Then
          param = "0,"
        Else
          param = (30 + calibType).ToString()
          i = 0
          While (i < calibData.Count)
            If (((i) And (1)) > 0) Then
              param = param + ":"
            Else
              param = param + " "
            End If
            param = param + (CType(Math.Round(calibData(i) * 1000.0 / 1000.0), Integer)).ToString()
            i = i + 1
          End While
          param = param + ","
        End If
      Else
        If (funVer >= 1) Then
          REM // Encode parameters for older devices
          nPoints = (calibData.Count \ 2)
          param = (nPoints).ToString()
          i = 0
          While (i < 2 * nPoints)
            If (funScale = 0) Then
              wordVal = YAPI._doubleToDecimal(CType(Math.Round(calibData(i)), Integer))
            Else
              wordVal = calibData(i) * funScale + funOffset
            End If
            param = param + "," + (Math.Round(wordVal)).ToString()
            i = i + 1
          End While
        Else
          REM // Initial V0 encoding used for old Yocto-Light
          If (calibData.Count = 4) Then
            param = (Math.Round(1000 * (calibData(3) - calibData(1)) / calibData(2) - calibData(0))).ToString()
          End If
        End If
      End If
      Return param
    End Function

    Public Overridable Function _tryExec(url As String) As Integer
      Dim res As Integer = 0
      Dim done As Integer = 0
      res = YAPI.SUCCESS
      done = 1
      Try
          Me._download(url)
      Catch
          done = 0
      End Try
      If (done = 0) Then
        REM // retry silently after a short wait
        Try
            YAPI.Sleep(500, Nothing)
            Me._download(url)
        Catch
            REM // second failure, return error code
            res = Me.get_errorType()
        End Try
      End If
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Restores all the settings of the device.
    ''' <para>
    '''   Useful to restore all the logical names and calibrations parameters
    '''   of a module from a backup.Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modifications must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="settings">
    '''   a binary buffer with all the settings.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_allSettings(settings As Byte()) As Integer
      Dim restoreLast As List(Of String) = New List(Of String)()
      Dim old_json_flat As Byte() = New Byte(){}
      Dim old_dslist As List(Of Byte()) = New List(Of Byte())()
      Dim old_jpath As List(Of String) = New List(Of String)()
      Dim old_jpath_len As List(Of Integer) = New List(Of Integer)()
      Dim old_val_arr As List(Of String) = New List(Of String)()
      Dim actualSettings As Byte() = New Byte(){}
      Dim new_dslist As List(Of Byte()) = New List(Of Byte())()
      Dim new_jpath As List(Of String) = New List(Of String)()
      Dim new_jpath_len As List(Of Integer) = New List(Of Integer)()
      Dim new_val_arr As List(Of String) = New List(Of String)()
      Dim cpos As Integer = 0
      Dim eqpos As Integer = 0
      Dim leng As Integer = 0
      Dim i As Integer = 0
      Dim j As Integer = 0
      Dim subres As Integer = 0
      Dim res As Integer = 0
      Dim njpath As String
      Dim jpath As String
      Dim fun As String
      Dim attr As String
      Dim value As String
      Dim old_serial As String
      Dim new_serial As String
      Dim url As String
      Dim tmp As String
      Dim binTmp As Byte() = New Byte(){}
      Dim sensorType As String
      Dim unit_name As String
      Dim newval As String
      Dim oldval As String
      Dim old_calib As String
      Dim each_str As String
      Dim do_update As Boolean
      Dim found As Boolean
      res = YAPI.SUCCESS
      binTmp = Me._get_json_path(settings, "api")
      If ((binTmp).Length > 0) Then
        settings = binTmp
      End If
      old_serial = ""
      oldval = ""
      newval = ""
      old_json_flat = Me._flattenJsonStruct(settings)
      old_dslist = Me._json_get_array(old_json_flat)



      Dim ii_0 As Integer
      For ii_0 = 0 To old_dslist.Count - 1
        each_str = Me._json_get_string(old_dslist(ii_0))
        REM // split json path and attr
        leng = (each_str).Length
        eqpos = each_str.IndexOf("=")
        If ((eqpos < 0) OrElse (leng = 0)) Then
          Me._throw(YAPI.INVALID_ARGUMENT, "Invalid settings")
          Return YAPI.INVALID_ARGUMENT
        End If
        jpath = (each_str).Substring(0, eqpos)
        eqpos = eqpos + 1
        value = (each_str).Substring(eqpos, leng - eqpos)
        old_jpath.Add(jpath)
        old_jpath_len.Add((jpath).Length)
        old_val_arr.Add(value)
        If (jpath = "module/serialNumber") Then
          old_serial = value
        End If
      Next ii_0




      Try
          actualSettings = Me._download("api.json")
      Catch
          REM // retry silently after a short wait
          YAPI.Sleep(500, Nothing)
          actualSettings = Me._download("api.json")
      End Try
      new_serial = Me.get_serialNumber()
      If (old_serial = new_serial OrElse old_serial = "") Then
        old_serial = "_NO_SERIAL_FILTER_"
      End If
      actualSettings = Me._flattenJsonStruct(actualSettings)
      new_dslist = Me._json_get_array(actualSettings)



      Dim ii_1 As Integer
      For ii_1 = 0 To new_dslist.Count - 1
        REM // remove quotes
        each_str = Me._json_get_string(new_dslist(ii_1))
        REM // split json path and attr
        leng = (each_str).Length
        eqpos = each_str.IndexOf("=")
        If ((eqpos < 0) OrElse (leng = 0)) Then
          Me._throw(YAPI.INVALID_ARGUMENT, "Invalid settings")
          Return YAPI.INVALID_ARGUMENT
        End If
        jpath = (each_str).Substring(0, eqpos)
        eqpos = eqpos + 1
        value = (each_str).Substring(eqpos, leng - eqpos)
        new_jpath.Add(jpath)
        new_jpath_len.Add((jpath).Length)
        new_val_arr.Add(value)
      Next ii_1




      i = 0
      While (i < new_jpath.Count)
        njpath = new_jpath(i)
        leng = (njpath).Length
        cpos = njpath.IndexOf("/")
        If ((cpos < 0) OrElse (leng = 0)) Then
          Continue While
        End If
        fun = (njpath).Substring(0, cpos)
        cpos = cpos + 1
        attr = (njpath).Substring(cpos, leng - cpos)
        do_update = True
        If (fun = "services") Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "firmwareRelease")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "usbCurrent")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "upTime")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "persistentSettings")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "adminPassword")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "userPassword")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "rebootCountdown")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "advertisedValue")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "poeCurrent")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "readiness")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "ipAddress")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "subnetMask")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "router")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "linkQuality")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "ssid")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "channel")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "security")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "message")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "signalValue")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "currentValue")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "currentRawValue")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "currentRunIndex")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "pulseTimer")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "lastTimePressed")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "lastTimeReleased")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "filesCount")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "freeSpace")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "timeUTC")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "rtcTime")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "unixTime")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "dateTime")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "rawValue")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "lastMsg")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "delayedPulseTimer")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "rxCount")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "txCount")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "msgCount")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "rxMsgCount")) Then
          do_update = False
        End If
        If (do_update AndAlso (attr = "txMsgCount")) Then
          do_update = False
        End If
        If (do_update) Then
          do_update = False
          j = 0
          found = False
          newval = new_val_arr(i)
          While ((j < old_jpath.Count) AndAlso Not (found))
            If ((new_jpath_len(i) = old_jpath_len(j)) AndAlso (new_jpath(i) = old_jpath(j))) Then
              found = True
              oldval = old_val_arr(j)
              If (Not (newval = oldval) AndAlso Not (oldval = old_serial)) Then
                do_update = True
              End If
            End If
            j = j + 1
          End While
        End If
        If (do_update) Then
          If (attr = "calibrationParam") Then
            old_calib = ""
            unit_name = ""
            sensorType = ""
            j = 0
            found = False
            While ((j < old_jpath.Count) AndAlso Not (found))
              If ((new_jpath_len(i) = old_jpath_len(j)) AndAlso (new_jpath(i) = old_jpath(j))) Then
                found = True
                old_calib = old_val_arr(j)
              End If
              j = j + 1
            End While
            tmp = fun + "/unit"
            j = 0
            found = False
            While ((j < new_jpath.Count) AndAlso Not (found))
              If (tmp = new_jpath(j)) Then
                found = True
                unit_name = new_val_arr(j)
              End If
              j = j + 1
            End While
            tmp = fun + "/sensorType"
            j = 0
            found = False
            While ((j < new_jpath.Count) AndAlso Not (found))
              If (tmp = new_jpath(j)) Then
                found = True
                sensorType = new_val_arr(j)
              End If
              j = j + 1
            End While
            newval = Me.calibConvert(old_calib, new_val_arr(i), unit_name, sensorType)
            url = "api/" + fun + ".json?" + attr + "=" + Me._escapeAttr(newval)
            subres = Me._tryExec(url)
            If ((res = YAPI.SUCCESS) AndAlso (subres <> YAPI.SUCCESS)) Then
              res = subres
            End If
          Else
            url = "api/" + fun + ".json?" + attr + "=" + Me._escapeAttr(oldval)
            If (attr = "resolution") Then
              restoreLast.Add(url)
            Else
              subres = Me._tryExec(url)
              If ((res = YAPI.SUCCESS) AndAlso (subres <> YAPI.SUCCESS)) Then
                res = subres
              End If
            End If
          End If
        End If
        i = i + 1
      End While

      Dim ii_2 As Integer
      For ii_2 = 0 To restoreLast.Count - 1
        subres = Me._tryExec(restoreLast(ii_2))
        If ((res = YAPI.SUCCESS) AndAlso (subres <> YAPI.SUCCESS)) Then
          res = subres
        End If
      Next ii_2
      Me.clearCache()
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Adds a file to the uploaded data at the next HTTP callback.
    ''' <para>
    '''   This function only affects the next HTTP callback and only works in
    '''   HTTP callback mode.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="filename">
    '''   the name of the file to upload at the next HTTP callback
    ''' </param>
    ''' <returns>
    '''   nothing.
    ''' </returns>
    '''/
    Public Overridable Function addFileToHTTPCallback(filename As String) As Integer
      Dim content As Byte() = New Byte(){}

      content = Me._download("@YCB+" + filename)
      If ((content).Length = 0) Then
        Return YAPI.NOT_SUPPORTED
      End If
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns the unique hardware identifier of the module.
    ''' <para>
    '''   The unique hardware identifier is made of the device serial
    '''   number followed by string ".module".
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string that uniquely identifies the module
    ''' </returns>
    '''/
    Public Overrides Function get_hardwareId() As String
      Dim serial As String

      serial = Me.get_serialNumber()
      Return serial + ".module"
    End Function

    '''*
    ''' <summary>
    '''   Downloads the specified built-in file and returns a binary buffer with its content.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="pathname">
    '''   name of the new file to load
    ''' </param>
    ''' <returns>
    '''   a binary buffer with the file content
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns  <c>YAPI.INVALID_STRING</c>.
    ''' </para>
    '''/
    Public Overridable Function download(pathname As String) As Byte()
      Return Me._download(pathname)
    End Function

    '''*
    ''' <summary>
    '''   Returns the icon of the module.
    ''' <para>
    '''   The icon is a PNG image and does not
    '''   exceed 1536 bytes.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a binary buffer with module icon, in png format.
    '''   On failure, throws an exception or returns  <c>YAPI.INVALID_STRING</c>.
    ''' </returns>
    '''/
    Public Overridable Function get_icon2d() As Byte()
      Return Me._download("icon2d.png")
    End Function

    '''*
    ''' <summary>
    '''   Returns a string with last logs of the module.
    ''' <para>
    '''   This method return only
    '''   logs that are still in the module.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with last logs of the module.
    '''   On failure, throws an exception or returns  <c>YAPI.INVALID_STRING</c>.
    ''' </returns>
    '''/
    Public Overridable Function get_lastLogs() As String
      Dim content As Byte() = New Byte(){}

      content = Me._download("logs.txt")
      Return YAPI.DefaultEncoding.GetString(content)
    End Function

    '''*
    ''' <summary>
    '''   Adds a text message to the device logs.
    ''' <para>
    '''   This function is useful in
    '''   particular to trace the execution of HTTP callbacks. If a newline
    '''   is desired after the message, it must be included in the string.
    ''' </para>
    ''' </summary>
    ''' <param name="text">
    '''   the string to append to the logs.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function log(text As String) As Integer
      Return Me._upload("logs.txt", YAPI.DefaultEncoding.GetBytes(text))
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of all the modules that are plugged into the current module.
    ''' <para>
    '''   This method only makes sense when called for a YoctoHub/VirtualHub.
    '''   Otherwise, an empty array will be returned.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an array of strings containing the sub modules.
    ''' </returns>
    '''/
    Public Overridable Function get_subDevices() As List(Of String)
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim smallbuff As StringBuilder = New StringBuilder(1024)
      Dim bigbuff As StringBuilder
      Dim buffsize As Integer = 0
      Dim fullsize As Integer
      Dim yapi_res As Integer = 0
      Dim subdevice_list As String
      Dim subdevices As List(Of String) = New List(Of String)()
      Dim serial As String

      serial = Me.get_serialNumber()
      fullsize = 0
      yapi_res = _yapiGetSubdevices(New StringBuilder(serial), smallbuff, 1024, fullsize, errmsg)
      If (yapi_res < 0) Then
        Return subdevices
      End If
      If (fullsize <= 1024) Then
        subdevice_list = smallbuff.ToString()
      Else
        buffsize = fullsize
        bigbuff = New StringBuilder(buffsize)
        yapi_res = _yapiGetSubdevices(New StringBuilder(serial), bigbuff, buffsize, fullsize, errmsg)
        If (yapi_res < 0) Then
          bigbuff = Nothing
          Return subdevices
        Else
          subdevice_list = bigbuff.ToString()
        End If
        bigbuff = Nothing
      End If
      If (Not (subdevice_list = "")) Then
        subdevices = New List(Of String)(subdevice_list.Split(New Char() {","c}))
      End If
      Return subdevices
    End Function

    '''*
    ''' <summary>
    '''   Returns the serial number of the YoctoHub on which this module is connected.
    ''' <para>
    '''   If the module is connected by USB, or if the module is the root YoctoHub, an
    '''   empty string is returned.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with the serial number of the YoctoHub or an empty string
    ''' </returns>
    '''/
    Public Overridable Function get_parentHub() As String
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim hubserial As StringBuilder = New StringBuilder(YOCTO_SERIAL_LEN)
      Dim pathsize As Integer
      Dim yapi_res As Integer = 0
      Dim serial As String

      serial = Me.get_serialNumber()
      REM // retrieve device object
      pathsize = 0
      yapi_res = _yapiGetDevicePathEx(New StringBuilder(serial), hubserial, Nothing, 0, pathsize, errmsg)
      If (yapi_res < 0) Then
        Return ""
      End If
      Return hubserial.ToString()
    End Function

    '''*
    ''' <summary>
    '''   Returns the URL used to access the module.
    ''' <para>
    '''   If the module is connected by USB, the
    '''   string 'usb' is returned.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with the URL of the module.
    ''' </returns>
    '''/
    Public Overridable Function get_url() As String
      Dim errmsg As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim path As StringBuilder = New StringBuilder(1024)
      Dim pathsize As Integer
      Dim yapi_res As Integer = 0
      Dim serial As String

      serial = Me.get_serialNumber()
      REM // retrieve device object
      pathsize = 0
      yapi_res = _yapiGetDevicePathEx(New StringBuilder(serial), Nothing, path, 1024, pathsize, errmsg)
      If (yapi_res < 0) Then
        Return ""
      End If
      Return path.ToString()
    End Function


    '''*
    ''' <summary>
    '''   Continues the module enumeration started using <c>yFirstModule()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned modules order.
    '''   If you want to find a specific module, use <c>Module.findModule()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YModule</c> object, corresponding to
    '''   the next module found, or a <c>Nothing</c> pointer
    '''   if there are no more modules to enumerate.
    ''' </returns>
    '''/
    Public Function nextModule() As YModule
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YModule.FindModule(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of modules currently accessible.
    ''' <para>
    '''   Use the method <c>YModule.nextModule()</c> to iterate on the
    '''   next modules.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YModule</c> object, corresponding to
    '''   the first module currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstModule() As YModule
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Module", 0, p, size, neededsize, errmsg)
      Marshal.Copy(p, v_fundescr, 0, 1)
      Marshal.FreeHGlobal(p)

      If (YISERR(err) Or (neededsize = 0)) Then
        Return Nothing
      End If
      serial = ""
      funcId = ""
      funcName = ""
      funcVal = ""
      errmsg = ""
      If (YISERR(yapiGetFunctionInfo(v_fundescr(0), dev, serial, funcId, funcName, funcVal, errmsg))) Then
        Return Nothing
      End If
      Return YModule.FindModule(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YModule public methods declaration)
  End Class


  REM --- (generated code: YModule functions)

  '''*
  ''' <summary>
  '''   Allows you to find a module from its serial number or from its logical name.
  ''' <para>
  ''' </para>
  ''' <para>
  '''   This function does not require that the module is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YModule.isOnline()</c> to test if the module is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a module by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   If a call to this object's is_online() method returns FALSE although
  '''   you are certain that the device is plugged, make sure that you did
  '''   call registerHub() at application initialization time.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string containing either the serial number or
  '''   the logical name of the desired module
  ''' </param>
  ''' <returns>
  '''   a <c>YModule</c> object allowing you to drive the module
  '''   or get additional information on the module.
  ''' </returns>
  '''/
  Public Function yFindModule(ByVal func As String) As YModule
    Return YModule.FindModule(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of modules currently accessible.
  ''' <para>
  '''   Use the method <c>YModule.nextModule()</c> to iterate on the
  '''   next modules.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YModule</c> object, corresponding to
  '''   the first module currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstModule() As YModule
    Return YModule.FirstModule()
  End Function


  REM --- (end of generated code: YModule functions)


  REM --- (generated code: YSensor class start)

  '''*
  ''' <summary>
  '''   The <c>YSensor</c> class is the parent class for all Yoctopuce sensor types.
  ''' <para>
  '''   It can be
  '''   used to read the current value and unit of any sensor, read the min/max
  '''   value, configure autonomous recording frequency and access recorded data.
  '''   It also provides a function to register a callback invoked each time the
  '''   observed value changes, or at a predefined interval. Using this class rather
  '''   than a specific subclass makes it possible to create generic applications
  '''   that work with any Yoctopuce sensor, even those that do not yet exist.
  '''   Note: The <c>YAnButton</c> class is the only analog input which does not inherit
  '''   from <c>YSensor</c>.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSensor
    Inherits YFunction
    REM --- (end of generated code: YSensor class start)

    REM --- (generated code: YSensor definitions)
    Public Const UNIT_INVALID As String = YAPI.INVALID_STRING
    Public Const CURRENTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const LOWESTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const HIGHESTVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const CURRENTRAWVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const LOGFREQUENCY_INVALID As String = YAPI.INVALID_STRING
    Public Const REPORTFREQUENCY_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVMODE_IMMEDIATE As Integer = 0
    Public Const ADVMODE_PERIOD_AVG As Integer = 1
    Public Const ADVMODE_PERIOD_MIN As Integer = 2
    Public Const ADVMODE_PERIOD_MAX As Integer = 3
    Public Const ADVMODE_INVALID As Integer = -1
    Public Const CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING
    Public Const RESOLUTION_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const SENSORSTATE_INVALID As Integer = YAPI.INVALID_INT
    REM --- (end of generated code: YSensor definitions)

    REM --- (generated code: YSensor attributes declaration)
    Protected _unit As String
    Protected _currentValue As Double
    Protected _lowestValue As Double
    Protected _highestValue As Double
    Protected _currentRawValue As Double
    Protected _logFrequency As String
    Protected _reportFrequency As String
    Protected _advMode As Integer
    Protected _calibrationParam As String
    Protected _resolution As Double
    Protected _sensorState As Integer
    Protected _valueCallbackSensor As YSensorValueCallback
    Protected _timedReportCallbackSensor As YSensorTimedReportCallback
    Protected _prevTimedReport As Double
    Protected _iresol As Double
    Protected _offset As Double
    Protected _scale As Double
    Protected _decexp As Double
    Protected _caltyp As Integer
    Protected _calpar As List(Of Integer)
    Protected _calraw As List(Of Double)
    Protected _calref As List(Of Double)
    Protected _calhdl As yCalibrationHandler
    REM --- (end of generated code: YSensor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _className = "Sensor"
      REM --- (generated code: YSensor attributes initialization)
      _unit = UNIT_INVALID
      _currentValue = CURRENTVALUE_INVALID
      _lowestValue = LOWESTVALUE_INVALID
      _highestValue = HIGHESTVALUE_INVALID
      _currentRawValue = CURRENTRAWVALUE_INVALID
      _logFrequency = LOGFREQUENCY_INVALID
      _reportFrequency = REPORTFREQUENCY_INVALID
      _advMode = ADVMODE_INVALID
      _calibrationParam = CALIBRATIONPARAM_INVALID
      _resolution = RESOLUTION_INVALID
      _sensorState = SENSORSTATE_INVALID
      _valueCallbackSensor = Nothing
      _timedReportCallbackSensor = Nothing
      _prevTimedReport = 0
      _iresol = 0
      _offset = 0
      _scale = 0
      _decexp = 0
      _caltyp = 0
      _calpar = New List(Of Integer)()
      _calraw = New List(Of Double)()
      _calref = New List(Of Double)()
      _calhdl = Nothing
      REM --- (end of generated code: YSensor attributes initialization)
    End Sub

    REM --- (generated code: YSensor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("unit") Then
        _unit = json_val.getString("unit")
      End If
      If json_val.has("currentValue") Then
        _currentValue = Math.Round(json_val.getDouble("currentValue") / 65.536) / 1000.0
      End If
      If json_val.has("lowestValue") Then
        _lowestValue = Math.Round(json_val.getDouble("lowestValue") / 65.536) / 1000.0
      End If
      If json_val.has("highestValue") Then
        _highestValue = Math.Round(json_val.getDouble("highestValue") / 65.536) / 1000.0
      End If
      If json_val.has("currentRawValue") Then
        _currentRawValue = Math.Round(json_val.getDouble("currentRawValue") / 65.536) / 1000.0
      End If
      If json_val.has("logFrequency") Then
        _logFrequency = json_val.getString("logFrequency")
      End If
      If json_val.has("reportFrequency") Then
        _reportFrequency = json_val.getString("reportFrequency")
      End If
      If json_val.has("advMode") Then
        _advMode = CInt(json_val.getLong("advMode"))
      End If
      If json_val.has("calibrationParam") Then
        _calibrationParam = json_val.getString("calibrationParam")
      End If
      If json_val.has("resolution") Then
        _resolution = Math.Round(json_val.getDouble("resolution") / 65.536) / 1000.0
      End If
      If json_val.has("sensorState") Then
        _sensorState = CInt(json_val.getLong("sensorState"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YSensor private methods declaration)

    REM --- (generated code: YSensor public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the measuring unit for the measure.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the measuring unit for the measure
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.UNIT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_unit() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return UNIT_INVALID
        End If
      End If
      res = Me._unit
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of the measure, in the specified unit, as a floating point number.
    ''' <para>
    '''   Note that a get_currentValue() call will *not* start a measure in the device, it
    '''   will just return the last measure that occurred in the device. Indeed, internally, each Yoctopuce
    '''   devices is continuously making measurements at a hardware specific frequency.
    ''' </para>
    ''' <para>
    '''   If continuously calling  get_currentValue() leads you to performances issues, then
    '''   you might consider to switch to callback programming model. Check the "advanced
    '''   programming" chapter in in your device user manual for more information.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current value of the measure, in the specified unit,
    '''   as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.CURRENTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CURRENTVALUE_INVALID
        End If
      End If
      res = Me._applyCalibration(Me._currentRawValue)
      If (res = CURRENTVALUE_INVALID) Then
        res = Me._currentValue
      End If
      res = res * Me._iresol
      res = Math.Round(res) / Me._iresol
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the recorded minimal value observed.
    ''' <para>
    '''   Can be used to reset the value returned
    '''   by get_lowestValue().
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the recorded minimal value observed
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_lowestValue(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("lowestValue", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the minimal value observed for the measure since the device was started.
    ''' <para>
    '''   Can be reset to an arbitrary value thanks to set_lowestValue().
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the minimal value observed for the measure since the device was started
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.LOWESTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lowestValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LOWESTVALUE_INVALID
        End If
      End If
      res = Me._lowestValue * Me._iresol
      res = Math.Round(res) / Me._iresol
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the recorded maximal value observed.
    ''' <para>
    '''   Can be used to reset the value returned
    '''   by get_lowestValue().
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the recorded maximal value observed
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_highestValue(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("highestValue", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the maximal value observed for the measure since the device was started.
    ''' <para>
    '''   Can be reset to an arbitrary value thanks to set_highestValue().
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the maximal value observed for the measure since the device was started
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.HIGHESTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_highestValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return HIGHESTVALUE_INVALID
        End If
      End If
      res = Me._highestValue * Me._iresol
      res = Math.Round(res) / Me._iresol
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the uncalibrated, unrounded raw value returned by the
    '''   sensor, in the specified unit, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the uncalibrated, unrounded raw value returned by the
    '''   sensor, in the specified unit, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.CURRENTRAWVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentRawValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CURRENTRAWVALUE_INVALID
        End If
      End If
      res = Me._currentRawValue
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the datalogger recording frequency for this function, or "OFF"
    '''   when measures are not stored in the data logger flash memory.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the datalogger recording frequency for this function, or "OFF"
    '''   when measures are not stored in the data logger flash memory
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.LOGFREQUENCY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logFrequency() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LOGFREQUENCY_INVALID
        End If
      End If
      res = Me._logFrequency
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the datalogger recording frequency for this function.
    ''' <para>
    '''   The frequency can be specified as samples per second,
    '''   as sample per minute (for instance "15/m") or in samples per
    '''   hour (eg. "4/h"). To disable recording for this function, use
    '''   the value "OFF". Note that setting the  datalogger recording frequency
    '''   to a greater value than the sensor native sampling frequency is useless,
    '''   and even counterproductive: those two frequencies are not related.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the datalogger recording frequency for this function
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_logFrequency(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("logFrequency", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the timed value notification frequency, or "OFF" if timed
    '''   value notifications are disabled for this function.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the timed value notification frequency, or "OFF" if timed
    '''   value notifications are disabled for this function
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.REPORTFREQUENCY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_reportFrequency() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return REPORTFREQUENCY_INVALID
        End If
      End If
      res = Me._reportFrequency
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the timed value notification frequency for this function.
    ''' <para>
    '''   The frequency can be specified as samples per second,
    '''   as sample per minute (for instance "15/m") or in samples per
    '''   hour (e.g. "4/h"). To disable timed value notifications for this
    '''   function, use the value "OFF". Note that setting the  timed value
    '''   notification frequency to a greater value than the sensor native
    '''   sampling frequency is unless, and even counterproductive: those two
    '''   frequencies are not related.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the timed value notification frequency for this function
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_reportFrequency(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("reportFrequency", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the measuring mode used for the advertised value pushed to the parent hub.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YSensor.ADVMODE_IMMEDIATE</c>, <c>YSensor.ADVMODE_PERIOD_AVG</c>,
    '''   <c>YSensor.ADVMODE_PERIOD_MIN</c> and <c>YSensor.ADVMODE_PERIOD_MAX</c> corresponding to the
    '''   measuring mode used for the advertised value pushed to the parent hub
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.ADVMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ADVMODE_INVALID
        End If
      End If
      res = Me._advMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the measuring mode used for the advertised value pushed to the parent hub.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YSensor.ADVMODE_IMMEDIATE</c>, <c>YSensor.ADVMODE_PERIOD_AVG</c>,
    '''   <c>YSensor.ADVMODE_PERIOD_MIN</c> and <c>YSensor.ADVMODE_PERIOD_MAX</c> corresponding to the
    '''   measuring mode used for the advertised value pushed to the parent hub
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_advMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("advMode", rest_val)
    End Function
    Public Function get_calibrationParam() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CALIBRATIONPARAM_INVALID
        End If
      End If
      res = Me._calibrationParam
      Return res
    End Function


    Public Function set_calibrationParam(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("calibrationParam", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the resolution of the measured physical values.
    ''' <para>
    '''   The resolution corresponds to the numerical precision
    '''   when displaying value. It does not change the precision of the measure itself.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the resolution of the measured physical values
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_resolution(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("resolution", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the resolution of the measured values.
    ''' <para>
    '''   The resolution corresponds to the numerical precision
    '''   of the measures, which is not always the same as the actual precision of the sensor.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the resolution of the measured values
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.RESOLUTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_resolution() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RESOLUTION_INVALID
        End If
      End If
      res = Me._resolution
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensor state code, which is zero when there is an up-to-date measure
    '''   available or a positive code if the sensor is not able to provide a measure right now.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the sensor state code, which is zero when there is an up-to-date measure
    '''   available or a positive code if the sensor is not able to provide a measure right now
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSensor.SENSORSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_sensorState() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SENSORSTATE_INVALID
        End If
      End If
      res = Me._sensorState
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a sensor for a given identifier.
    ''' <para>
    '''   The identifier can be specified using several formats:
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   - FunctionLogicalName
    ''' </para>
    ''' <para>
    '''   - ModuleSerialNumber.FunctionIdentifier
    ''' </para>
    ''' <para>
    '''   - ModuleSerialNumber.FunctionLogicalName
    ''' </para>
    ''' <para>
    '''   - ModuleLogicalName.FunctionIdentifier
    ''' </para>
    ''' <para>
    '''   - ModuleLogicalName.FunctionLogicalName
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   This function does not require that the sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSensor.isOnline()</c> to test if the sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a sensor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' <para>
    '''   If a call to this object's is_online() method returns FALSE although
    '''   you are certain that the matching device is plugged, make sure that you did
    '''   call registerHub() at application initialization time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the sensor, for instance
    '''   <c>MyDevice.</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YSensor</c> object allowing you to drive the sensor.
    ''' </returns>
    '''/
    Public Shared Function FindSensor(func As String) As YSensor
      Dim obj As YSensor
      obj = CType(YFunction._FindFromCache("Sensor", func), YSensor)
      If ((obj Is Nothing)) Then
        obj = New YSensor(func)
        YFunction._AddToCache("Sensor", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerValueCallback(callback As YSensorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSensor = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) AndAlso Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overrides Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackSensor Is Nothing)) Then
        Me._valueCallbackSensor(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overrides Function _parserHelper() As Integer
      Dim position As Integer = 0
      Dim maxpos As Integer = 0
      Dim iCalib As List(Of Integer) = New List(Of Integer)()
      Dim iRaw As Integer = 0
      Dim iRef As Integer = 0
      Dim fRaw As Double = 0
      Dim fRef As Double = 0
      Me._caltyp = -1
      Me._scale = -1
      Me._calpar.Clear()
      Me._calraw.Clear()
      Me._calref.Clear()
      REM // Store inverted resolution, to provide better rounding
      If (Me._resolution > 0) Then
        Me._iresol = Math.Round(1.0 / Me._resolution)
      Else
        Me._iresol = 10000
        Me._resolution = 0.0001
      End If
      REM // Old format: supported when there is no calibration
      If (Me._calibrationParam = "" OrElse Me._calibrationParam = "0") Then
        Me._caltyp = 0
        Return 0
      End If
      If (Me._calibrationParam.IndexOf(",") >= 0) Then
        REM // Plain text format
        iCalib = YAPI._decodeFloats(Me._calibrationParam)
        Me._caltyp = (iCalib(0) \ 1000)
        If (Me._caltyp > 0) Then
          If (Me._caltyp < YOCTO_CALIB_TYPE_OFS) Then
            REM // Unknown calibration type: calibrated value will be provided by the device
            Me._caltyp = -1
            Return 0
          End If
          Me._calhdl = YAPI._getCalibrationHandler(Me._caltyp)
          If (Not (Not (Me._calhdl Is Nothing))) Then
            REM // Unknown calibration type: calibrated value will be provided by the device
            Me._caltyp = -1
            Return 0
          End If
        End If
        REM // New 32 bits text format
        Me._offset = 0
        Me._scale = 1000
        maxpos = iCalib.Count
        Me._calpar.Clear()
        position = 1
        While (position < maxpos)
          Me._calpar.Add(iCalib(position))
          position = position + 1
        End While
        Me._calraw.Clear()
        Me._calref.Clear()
        position = 1
        While (position + 1 < maxpos)
          fRaw = iCalib(position)
          fRaw = fRaw / 1000.0
          fRef = iCalib(position + 1)
          fRef = fRef / 1000.0
          Me._calraw.Add(fRaw)
          Me._calref.Add(fRef)
          position = position + 2
        End While
      Else
        REM // Recorder-encoded format, including encoding
        iCalib = YAPI._decodeWords(Me._calibrationParam)
        REM // In case of unknown format, calibrated value will be provided by the device
        If (iCalib.Count < 2) Then
          Me._caltyp = -1
          Return 0
        End If
        REM // Save variable format (scale for scalar, or decimal exponent)
        Me._offset = 0
        Me._scale = 1
        Me._decexp = 1.0
        position = iCalib(0)
        While (position > 0)
          Me._decexp = Me._decexp * 10
          position = position - 1
        End While
        REM // Shortcut when there is no calibration parameter
        If (iCalib.Count = 2) Then
          Me._caltyp = 0
          Return 0
        End If
        Me._caltyp = iCalib(2)
        Me._calhdl = YAPI._getCalibrationHandler(Me._caltyp)
        REM // parse calibration points
        If (Me._caltyp <= 10) Then
          maxpos = Me._caltyp
        Else
          If (Me._caltyp <= 20) Then
            maxpos = Me._caltyp - 10
          Else
            maxpos = 5
          End If
        End If
        maxpos = 3 + 2 * maxpos
        If (maxpos > iCalib.Count) Then
          maxpos = iCalib.Count
        End If
        Me._calpar.Clear()
        Me._calraw.Clear()
        Me._calref.Clear()
        position = 3
        While (position + 1 < maxpos)
          iRaw = iCalib(position)
          iRef = iCalib(position + 1)
          Me._calpar.Add(iRaw)
          Me._calpar.Add(iRef)
          Me._calraw.Add(YAPI._decimalToDouble(iRaw))
          Me._calref.Add(YAPI._decimalToDouble(iRef))
          position = position + 2
        End While
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Checks if the sensor is currently able to provide an up-to-date measure.
    ''' <para>
    '''   Returns false if the device is unreachable, or if the sensor does not have
    '''   a current measure to transmit. No exception is raised if there is an error
    '''   while trying to contact the device hosting $THEFUNCTION$.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>true</c> if the sensor can provide an up-to-date measure, and <c>false</c> otherwise
    ''' </returns>
    '''/
    Public Overridable Function isSensorReady() As Boolean
      If (Not (Me.isOnline())) Then
        Return False
      End If
      If (Not (Me._sensorState = 0)) Then
        Return False
      End If
      Return True
    End Function

    '''*
    ''' <summary>
    '''   Returns the <c>YDatalogger</c> object of the device hosting the sensor.
    ''' <para>
    '''   This method returns an object
    '''   that can control global parameters of the data logger. The returned object
    '''   should not be freed.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an <c>YDatalogger</c> object, or Nothing on error.
    ''' </returns>
    '''/
    Public Overridable Function get_dataLogger() As YDataLogger
      Dim logger As YDataLogger
      Dim modu As YModule
      Dim serial As String
      Dim hwid As String

      modu = Me.get_module()
      serial = modu.get_serialNumber()
      If (serial = YAPI.INVALID_STRING) Then
        Return Nothing
      End If
      hwid = serial + ".dataLogger"

      logger = YDataLogger.FindDataLogger(hwid)
      Return logger
    End Function

    '''*
    ''' <summary>
    '''   Starts the data logger on the device.
    ''' <para>
    '''   Note that the data logger
    '''   will only save the measures on this sensor if the logFrequency
    '''   is not set to "OFF".
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    '''/
    Public Overridable Function startDataLogger() As Integer
      Dim res As Byte() = New Byte(){}

      res = Me._download("api/dataLogger/recording?recording=1")
      If Not((res).Length > 0) Then
        me._throw(YAPI.IO_ERROR, "unable to start datalogger")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Stops the datalogger on the device.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    '''/
    Public Overridable Function stopDataLogger() As Integer
      Dim res As Byte() = New Byte(){}

      res = Me._download("api/dataLogger/recording?recording=0")
      If Not((res).Length > 0) Then
        me._throw(YAPI.IO_ERROR, "unable to stop datalogger")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a <c>YDataSet</c> object holding historical data for this
    '''   sensor, for a specified time interval.
    ''' <para>
    '''   The measures will be
    '''   retrieved from the data logger, which must have been turned
    '''   on at the desired time. See the documentation of the <c>YDataSet</c>
    '''   class for information on how to get an overview of the
    '''   recorded data, and how to load progressively a large set
    '''   of measures from the data logger.
    ''' </para>
    ''' <para>
    '''   This function only works if the device uses a recent firmware,
    '''   as <c>YDataSet</c> objects are not supported by firmwares older than
    '''   version 13000.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="startTime">
    '''   the start of the desired measure time interval,
    '''   as a Unix timestamp, i.e. the number of seconds since
    '''   January 1, 1970 UTC. The special value 0 can be used
    '''   to include any measure, without initial limit.
    ''' </param>
    ''' <param name="endTime">
    '''   the end of the desired measure time interval,
    '''   as a Unix timestamp, i.e. the number of seconds since
    '''   January 1, 1970 UTC. The special value 0 can be used
    '''   to include any measure, without ending limit.
    ''' </param>
    ''' <returns>
    '''   an instance of <c>YDataSet</c>, providing access to historical
    '''   data. Past measures can be loaded progressively
    '''   using methods from the <c>YDataSet</c> object.
    ''' </returns>
    '''/
    Public Overridable Function get_recordedData(startTime As Double, endTime As Double) As YDataSet
      Dim funcid As String
      Dim funit As String

      funcid = Me.get_functionId()
      funit = Me.get_unit()

      Return New YDataSet(Me, funcid, funit, startTime, endTime)
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every periodic timed notification.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and an <c>YMeasure</c> object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Function registerTimedReportCallback(callback As YSensorTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackSensor = callback
      Return 0
    End Function

    Public Overridable Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackSensor Is Nothing)) Then
        Me._timedReportCallbackSensor(Me, value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Configures error correction data points, in particular to compensate for
    '''   a possible perturbation of the measure caused by an enclosure.
    ''' <para>
    '''   It is possible
    '''   to configure up to five correction points. Correction points must be provided
    '''   in ascending order, and be in the range of the sensor. The device will automatically
    '''   perform a linear interpolation of the error correction between specified
    '''   points. Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    '''   For more information on advanced capabilities to refine the calibration of
    '''   sensors, please contact support@yoctopuce.com.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="rawValues">
    '''   array of floating point numbers, corresponding to the raw
    '''   values returned by the sensor for the correction points.
    ''' </param>
    ''' <param name="refValues">
    '''   array of floating point numbers, corresponding to the corrected
    '''   values for the correction points.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function calibrateFromPoints(rawValues As List(Of Double), refValues As List(Of Double)) As Integer
      Dim rest_val As String
      Dim res As Integer = 0

      rest_val = Me._encodeCalibrationPoints(rawValues, refValues)
      res = Me._setAttr("calibrationParam", rest_val)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves error correction data points previously entered using the method
    '''   <c>calibrateFromPoints</c>.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="rawValues">
    '''   array of floating point numbers, that will be filled by the
    '''   function with the raw sensor values for the correction points.
    ''' </param>
    ''' <param name="refValues">
    '''   array of floating point numbers, that will be filled by the
    '''   function with the desired values for the correction points.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadCalibrationPoints(rawValues As List(Of Double), refValues As List(Of Double)) As Integer
      rawValues.Clear()
      refValues.Clear()
      REM // Load function parameters if not yet loaded
      If ((Me._scale = 0) OrElse (Me._cacheExpiration <= YAPI.GetTickCount())) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
      End If
      If (Me._caltyp < 0) Then
        Me._throw(YAPI.NOT_SUPPORTED, "Calibration parameters format mismatch. Please upgrade your library or firmware.")
        Return YAPI.NOT_SUPPORTED
      End If
      rawValues.Clear()
      refValues.Clear()
      Dim ii_0 As Integer
      For ii_0 = 0 To Me._calraw.Count - 1
        rawValues.Add(Me._calraw(ii_0))
      Next ii_0
      Dim ii_1 As Integer
      For ii_1 = 0 To Me._calref.Count - 1
        refValues.Add(Me._calref(ii_1))
      Next ii_1
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function _encodeCalibrationPoints(rawValues As List(Of Double), refValues As List(Of Double)) As String
      Dim res As String
      Dim npt As Integer = 0
      Dim idx As Integer = 0
      npt = rawValues.Count
      If (npt <> refValues.Count) Then
        Me._throw(YAPI.INVALID_ARGUMENT, "Invalid calibration parameters (size mismatch)")
        Return YAPI.INVALID_STRING
      End If
      REM // Shortcut when building empty calibration parameters
      If (npt = 0) Then
        Return "0"
      End If
      REM // Load function parameters if not yet loaded
      If (Me._scale = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return YAPI.INVALID_STRING
        End If
      End If
      REM // Detect old firmware
      If ((Me._caltyp < 0) OrElse (Me._scale < 0)) Then
        Me._throw(YAPI.NOT_SUPPORTED, "Calibration parameters format mismatch. Please upgrade your library or firmware.")
        Return "0"
      End If
      REM // 32-bit fixed-point encoding
      res = "" + Convert.ToString(YOCTO_CALIB_TYPE_OFS)
      idx = 0
      While (idx < npt)
        res = "" + res + "," + YAPI._floatToStr(rawValues(idx)) + "," + YAPI._floatToStr(refValues(idx))
        idx = idx + 1
      End While
      Return res
    End Function

    Public Overridable Function _applyCalibration(rawValue As Double) As Double
      If (rawValue = CURRENTVALUE_INVALID) Then
        Return CURRENTVALUE_INVALID
      End If
      If (Me._caltyp = 0) Then
        Return rawValue
      End If
      If (Me._caltyp < 0) Then
        Return CURRENTVALUE_INVALID
      End If
      If (Not (Not (Me._calhdl Is Nothing))) Then
        Return CURRENTVALUE_INVALID
      End If
      Return Me._calhdl(rawValue, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
    End Function

    Public Overridable Function _decodeTimedReport(timestamp As Double, duration As Double, report As List(Of Integer)) As YMeasure
      Dim i As Integer = 0
      Dim byteVal As Integer = 0
      Dim poww As Double = 0
      Dim minRaw As Double = 0
      Dim avgRaw As Double = 0
      Dim maxRaw As Double = 0
      Dim sublen As Integer = 0
      Dim difRaw As Double = 0
      Dim startTime As Double = 0
      Dim endTime As Double = 0
      Dim minVal As Double = 0
      Dim avgVal As Double = 0
      Dim maxVal As Double = 0
      If (duration > 0) Then
        startTime = timestamp - duration
      Else
        startTime = Me._prevTimedReport
      End If
      endTime = timestamp
      Me._prevTimedReport = endTime
      If (startTime = 0) Then
        startTime = endTime
      End If
      REM // 32 bits timed report format
      If (report.Count <= 5) Then
        REM // sub-second report, 1-4 bytes
        poww = 1
        avgRaw = 0
        byteVal = 0
        i = 1
        While (i < report.Count)
          byteVal = report(i)
          avgRaw = avgRaw + poww * byteVal
          poww = poww * &H100
          i = i + 1
        End While
        If (((byteVal) And (&H80)) <> 0) Then
          avgRaw = avgRaw - poww
        End If
        avgVal = avgRaw / 1000.0
        If (Me._caltyp <> 0) Then
          If (Not (Me._calhdl Is Nothing)) Then
            avgVal = Me._calhdl(avgVal, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
          End If
        End If
        minVal = avgVal
        maxVal = avgVal
      Else
        REM // averaged report: avg,avg-min,max-avg
        sublen = 1 + ((report(1)) And (3))
        poww = 1
        avgRaw = 0
        byteVal = 0
        i = 2
        While ((sublen > 0) AndAlso (i < report.Count))
          byteVal = report(i)
          avgRaw = avgRaw + poww * byteVal
          poww = poww * &H100
          i = i + 1
          sublen = sublen - 1
        End While
        If (((byteVal) And (&H80)) <> 0) Then
          avgRaw = avgRaw - poww
        End If
        sublen = 1 + ((((report(1)) >> 2)) And (3))
        poww = 1
        difRaw = 0
        While ((sublen > 0) AndAlso (i < report.Count))
          byteVal = report(i)
          difRaw = difRaw + poww * byteVal
          poww = poww * &H100
          i = i + 1
          sublen = sublen - 1
        End While
        minRaw = avgRaw - difRaw
        sublen = 1 + ((((report(1)) >> 4)) And (3))
        poww = 1
        difRaw = 0
        While ((sublen > 0) AndAlso (i < report.Count))
          byteVal = report(i)
          difRaw = difRaw + poww * byteVal
          poww = poww * &H100
          i = i + 1
          sublen = sublen - 1
        End While
        maxRaw = avgRaw + difRaw
        avgVal = avgRaw / 1000.0
        minVal = minRaw / 1000.0
        maxVal = maxRaw / 1000.0
        If (Me._caltyp <> 0) Then
          If (Not (Me._calhdl Is Nothing)) Then
            avgVal = Me._calhdl(avgVal, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
            minVal = Me._calhdl(minVal, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
            maxVal = Me._calhdl(maxVal, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
          End If
        End If
      End If
      Return New YMeasure(startTime, endTime, minVal, avgVal, maxVal)
    End Function

    Public Overridable Function _decodeVal(w As Integer) As Double
      Dim val As Double = 0
      val = w
      If (Me._caltyp <> 0) Then
        If (Not (Me._calhdl Is Nothing)) Then
          val = Me._calhdl(val, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
        End If
      End If
      Return val
    End Function

    Public Overridable Function _decodeAvg(dw As Integer) As Double
      Dim val As Double = 0
      val = dw
      If (Me._caltyp <> 0) Then
        If (Not (Me._calhdl Is Nothing)) Then
          val = Me._calhdl(val, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
        End If
      End If
      Return val
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of sensors started using <c>yFirstSensor()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned sensors order.
    '''   If you want to find a specific a sensor, use <c>Sensor.findSensor()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSensor</c> object, corresponding to
    '''   a sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextSensor() As YSensor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSensor.FindSensor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YSensor.nextSensor()</c> to iterate on
    '''   next sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSensor</c> object, corresponding to
    '''   the first sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstSensor() As YSensor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Sensor", 0, p, size, neededsize, errmsg)
      Marshal.Copy(p, v_fundescr, 0, 1)
      Marshal.FreeHGlobal(p)

      If (YISERR(err) Or (neededsize = 0)) Then
        Return Nothing
      End If
      serial = ""
      funcId = ""
      funcName = ""
      funcVal = ""
      errmsg = ""
      If (YISERR(yapiGetFunctionInfo(v_fundescr(0), dev, serial, funcId, funcName, funcVal, errmsg))) Then
        Return Nothing
      End If
      Return YSensor.FindSensor(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YSensor public methods declaration)
  End Class

  REM --- (generated code: YSensor functions)

  '''*
  ''' <summary>
  '''   Retrieves a sensor for a given identifier.
  ''' <para>
  '''   The identifier can be specified using several formats:
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   - FunctionLogicalName
  ''' </para>
  ''' <para>
  '''   - ModuleSerialNumber.FunctionIdentifier
  ''' </para>
  ''' <para>
  '''   - ModuleSerialNumber.FunctionLogicalName
  ''' </para>
  ''' <para>
  '''   - ModuleLogicalName.FunctionIdentifier
  ''' </para>
  ''' <para>
  '''   - ModuleLogicalName.FunctionLogicalName
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   This function does not require that the sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSensor.isOnline()</c> to test if the sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a sensor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' <para>
  '''   If a call to this object's is_online() method returns FALSE although
  '''   you are certain that the matching device is plugged, make sure that you did
  '''   call registerHub() at application initialization time.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the sensor, for instance
  '''   <c>MyDevice.</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YSensor</c> object allowing you to drive the sensor.
  ''' </returns>
  '''/
  Public Function yFindSensor(ByVal func As String) As YSensor
    Return YSensor.FindSensor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YSensor.nextSensor()</c> to iterate on
  '''   next sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YSensor</c> object, corresponding to
  '''   the first sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSensor() As YSensor
    Return YSensor.FirstSensor()
  End Function


  REM --- (end of generated code: YSensor functions)

  REM --- (generated code: YDataLogger globals)

  Public Const Y_CURRENTRUNINDEX_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_TIMEUTC_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_RECORDING_OFF As Integer = 0
  Public Const Y_RECORDING_ON As Integer = 1
  Public Const Y_RECORDING_PENDING As Integer = 2
  Public Const Y_RECORDING_INVALID As Integer = -1
  REM Y_AUTOSTART is defined in yocto_api.vb
  Public Const Y_BEACONDRIVEN_OFF As Integer = 0
  Public Const Y_BEACONDRIVEN_ON As Integer = 1
  Public Const Y_BEACONDRIVEN_INVALID As Integer = -1
  Public Const Y_USAGE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CLEARHISTORY_FALSE As Integer = 0
  Public Const Y_CLEARHISTORY_TRUE As Integer = 1
  Public Const Y_CLEARHISTORY_INVALID As Integer = -1
  Public Delegate Sub YDataLoggerValueCallback(ByVal func As YDataLogger, ByVal value As String)
  Public Delegate Sub YDataLoggerTimedReportCallback(ByVal func As YDataLogger, ByVal measure As YMeasure)
  REM --- (end of generated code: YDataLogger globals)


  REM --- (generated code: YDataLogger class start)

  '''*
  ''' <summary>
  '''   A non-volatile memory for storing ongoing measured data is available on most Yoctopuce
  '''   sensors.
  ''' <para>
  '''   Recording can happen automatically, without requiring a permanent
  '''   connection to a computer.
  '''   The <c>YDataLogger</c> class controls the global parameters of the internal data
  '''   logger. Recording control (start/stop) as well as data retrieval is done at
  '''   sensor objects level.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YDataLogger
    Inherits YFunction
    REM --- (end of generated code: YDataLogger class start)

    REM --- (generated code: YDataLogger definitions)
    Public Const CURRENTRUNINDEX_INVALID As Integer = YAPI.INVALID_UINT
    Public Const TIMEUTC_INVALID As Long = YAPI.INVALID_LONG
    Public Const RECORDING_OFF As Integer = 0
    Public Const RECORDING_ON As Integer = 1
    Public Const RECORDING_PENDING As Integer = 2
    Public Const RECORDING_INVALID As Integer = -1
    Public Const AUTOSTART_OFF As Integer = 0
    Public Const AUTOSTART_ON As Integer = 1
    Public Const AUTOSTART_INVALID As Integer = -1
    Public Const BEACONDRIVEN_OFF As Integer = 0
    Public Const BEACONDRIVEN_ON As Integer = 1
    Public Const BEACONDRIVEN_INVALID As Integer = -1
    Public Const USAGE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CLEARHISTORY_FALSE As Integer = 0
    Public Const CLEARHISTORY_TRUE As Integer = 1
    Public Const CLEARHISTORY_INVALID As Integer = -1
    REM --- (end of generated code: YDataLogger definitions)

    REM --- (generated code: YDataLogger attributes declaration)
    Protected _currentRunIndex As Integer
    Protected _timeUTC As Long
    Protected _recording As Integer
    Protected _autoStart As Integer
    Protected _beaconDriven As Integer
    Protected _usage As Integer
    Protected _clearHistory As Integer
    Protected _valueCallbackDataLogger As YDataLoggerValueCallback
    REM --- (end of generated code: YDataLogger attributes declaration)
    Protected _dataLoggerURL As String

    Public Sub New(ByVal func As String)
      MyBase.new(func)
      _className = "DataLogger"
      REM --- (generated code: YDataLogger attributes initialization)
      _currentRunIndex = CURRENTRUNINDEX_INVALID
      _timeUTC = TIMEUTC_INVALID
      _recording = RECORDING_INVALID
      _autoStart = AUTOSTART_INVALID
      _beaconDriven = BEACONDRIVEN_INVALID
      _usage = USAGE_INVALID
      _clearHistory = CLEARHISTORY_INVALID
      _valueCallbackDataLogger = Nothing
      REM --- (end of generated code: YDataLogger attributes initialization)
    End Sub

    REM --- (generated code: YDataLogger private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("currentRunIndex") Then
        _currentRunIndex = CInt(json_val.getLong("currentRunIndex"))
      End If
      If json_val.has("timeUTC") Then
        _timeUTC = json_val.getLong("timeUTC")
      End If
      If json_val.has("recording") Then
        _recording = CInt(json_val.getLong("recording"))
      End If
      If json_val.has("autoStart") Then
        If (json_val.getInt("autoStart") > 0) Then _autoStart = 1 Else _autoStart = 0
      End If
      If json_val.has("beaconDriven") Then
        If (json_val.getInt("beaconDriven") > 0) Then _beaconDriven = 1 Else _beaconDriven = 0
      End If
      If json_val.has("usage") Then
        _usage = CInt(json_val.getLong("usage"))
      End If
      If json_val.has("clearHistory") Then
        If (json_val.getInt("clearHistory") > 0) Then _clearHistory = 1 Else _clearHistory = 0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YDataLogger private methods declaration)

    REM --- (generated code: YDataLogger public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current run number, corresponding to the number of times the module was
    '''   powered on with the dataLogger enabled at some point.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current run number, corresponding to the number of times the module was
    '''   powered on with the dataLogger enabled at some point
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDataLogger.CURRENTRUNINDEX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentRunIndex() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CURRENTRUNINDEX_INVALID
        End If
      End If
      res = Me._currentRunIndex
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the Unix timestamp for current UTC time, if known.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the Unix timestamp for current UTC time, if known
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDataLogger.TIMEUTC_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_timeUTC() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TIMEUTC_INVALID
        End If
      End If
      res = Me._timeUTC
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current UTC time reference used for recorded data.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current UTC time reference used for recorded data
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_timeUTC(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("timeUTC", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current activation state of the data logger.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YDataLogger.RECORDING_OFF</c>, <c>YDataLogger.RECORDING_ON</c> and
    '''   <c>YDataLogger.RECORDING_PENDING</c> corresponding to the current activation state of the data logger
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDataLogger.RECORDING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_recording() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RECORDING_INVALID
        End If
      End If
      res = Me._recording
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the activation state of the data logger to start/stop recording data.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YDataLogger.RECORDING_OFF</c>, <c>YDataLogger.RECORDING_ON</c> and
    '''   <c>YDataLogger.RECORDING_PENDING</c> corresponding to the activation state of the data logger to
    '''   start/stop recording data
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_recording(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("recording", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the default activation state of the data logger on power up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YDataLogger.AUTOSTART_OFF</c> or <c>YDataLogger.AUTOSTART_ON</c>, according to the
    '''   default activation state of the data logger on power up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDataLogger.AUTOSTART_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_autoStart() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return AUTOSTART_INVALID
        End If
      End If
      res = Me._autoStart
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the default activation state of the data logger on power up.
    ''' <para>
    '''   Do not forget to call the <c>saveToFlash()</c> method of the module to save the
    '''   configuration change.  Note: if the device doesn't have any time source at his disposal when
    '''   starting up, it will wait for ~8 seconds before automatically starting to record  with
    '''   an arbitrary timestamp
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YDataLogger.AUTOSTART_OFF</c> or <c>YDataLogger.AUTOSTART_ON</c>, according to the
    '''   default activation state of the data logger on power up
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_autoStart(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("autoStart", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns true if the data logger is synchronised with the localization beacon.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YDataLogger.BEACONDRIVEN_OFF</c> or <c>YDataLogger.BEACONDRIVEN_ON</c>, according to true
    '''   if the data logger is synchronised with the localization beacon
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDataLogger.BEACONDRIVEN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_beaconDriven() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BEACONDRIVEN_INVALID
        End If
      End If
      res = Me._beaconDriven
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the type of synchronisation of the data logger.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YDataLogger.BEACONDRIVEN_OFF</c> or <c>YDataLogger.BEACONDRIVEN_ON</c>, according to the
    '''   type of synchronisation of the data logger
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_beaconDriven(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("beaconDriven", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the percentage of datalogger memory in use.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the percentage of datalogger memory in use
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YDataLogger.USAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_usage() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return USAGE_INVALID
        End If
      End If
      res = Me._usage
      Return res
    End Function

    Public Function get_clearHistory() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CLEARHISTORY_INVALID
        End If
      End If
      res = Me._clearHistory
      Return res
    End Function


    Public Function set_clearHistory(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("clearHistory", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a data logger for a given identifier.
    ''' <para>
    '''   The identifier can be specified using several formats:
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   - FunctionLogicalName
    ''' </para>
    ''' <para>
    '''   - ModuleSerialNumber.FunctionIdentifier
    ''' </para>
    ''' <para>
    '''   - ModuleSerialNumber.FunctionLogicalName
    ''' </para>
    ''' <para>
    '''   - ModuleLogicalName.FunctionIdentifier
    ''' </para>
    ''' <para>
    '''   - ModuleLogicalName.FunctionLogicalName
    ''' </para>
    ''' <para>
    ''' </para>
    ''' <para>
    '''   This function does not require that the data logger is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YDataLogger.isOnline()</c> to test if the data logger is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a data logger by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' <para>
    '''   If a call to this object's is_online() method returns FALSE although
    '''   you are certain that the matching device is plugged, make sure that you did
    '''   call registerHub() at application initialization time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the data logger, for instance
    '''   <c>LIGHTMK4.dataLogger</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YDataLogger</c> object allowing you to drive the data logger.
    ''' </returns>
    '''/
    Public Shared Function FindDataLogger(func As String) As YDataLogger
      Dim obj As YDataLogger
      obj = CType(YFunction._FindFromCache("DataLogger", func), YDataLogger)
      If ((obj Is Nothing)) Then
        obj = New YDataLogger(func)
        YFunction._AddToCache("DataLogger", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerValueCallback(callback As YDataLoggerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackDataLogger = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) AndAlso Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overrides Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackDataLogger Is Nothing)) Then
        Me._valueCallbackDataLogger(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Clears the data logger memory and discards all recorded data streams.
    ''' <para>
    '''   This method also resets the current run index to zero.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function forgetAllDataStreams() As Integer
      Return Me.set_clearHistory(CLEARHISTORY_TRUE)
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of <c>YDataSet</c> objects that can be used to retrieve
    '''   all measures stored by the data logger.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   This function only works if the device uses a recent firmware,
    '''   as <c>YDataSet</c> objects are not supported by firmwares older than
    '''   version 13000.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of <c>YDataSet</c> object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list.
    ''' </para>
    '''/
    Public Overridable Function get_dataSets() As List(Of YDataSet)
      Return Me.parse_dataSets(Me._download("logger.json"))
    End Function

    Public Overridable Function parse_dataSets(jsonbuff As Byte()) As List(Of YDataSet)
      Dim dslist As List(Of Byte()) = New List(Of Byte())()
      Dim dataset As YDataSet
      Dim res As List(Of YDataSet) = New List(Of YDataSet)()


      dslist = Me._json_get_array(jsonbuff)
      res.Clear()
      Dim ii_0 As Integer
      For ii_0 = 0 To dslist.Count - 1
        dataset = New YDataSet(Me)
        dataset._parse(YAPI.DefaultEncoding.GetString(dslist(ii_0)))
        res.Add(dataset)
      Next ii_0
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of data loggers started using <c>yFirstDataLogger()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned data loggers order.
    '''   If you want to find a specific a data logger, use <c>DataLogger.findDataLogger()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDataLogger</c> object, corresponding to
    '''   a data logger currently online, or a <c>Nothing</c> pointer
    '''   if there are no more data loggers to enumerate.
    ''' </returns>
    '''/
    Public Function nextDataLogger() As YDataLogger
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YDataLogger.FindDataLogger(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of data loggers currently accessible.
    ''' <para>
    '''   Use the method <c>YDataLogger.nextDataLogger()</c> to iterate on
    '''   next data loggers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YDataLogger</c> object, corresponding to
    '''   the first data logger currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstDataLogger() As YDataLogger
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("DataLogger", 0, p, size, neededsize, errmsg)
      Marshal.Copy(p, v_fundescr, 0, 1)
      Marshal.FreeHGlobal(p)

      If (YISERR(err) Or (neededsize = 0)) Then
        Return Nothing
      End If
      serial = ""
      funcId = ""
      funcName = ""
      funcVal = ""
      errmsg = ""
      If (YISERR(yapiGetFunctionInfo(v_fundescr(0), dev, serial, funcId, funcName, funcVal, errmsg))) Then
        Return Nothing
      End If
      Return YDataLogger.FindDataLogger(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YDataLogger public methods declaration)


    Public Function getData(runIdx As Long, timeIdx As Long, ByRef jsondata As YJSONContent) As Integer
      Dim dev As YDevice = Nothing
      Dim errmsg As String = ""
      Dim query As String = Nothing
      Dim buffer As String = ""
      Dim res As Integer = 0
      Dim http_headerlen As Integer

      If _dataLoggerURL = "" Then
        _dataLoggerURL = "/logger.json"
      End If
      REM Resolve our reference to our device, load REST API
      res = _getDevice(dev, errmsg)
      If YISERR(res) Then
        _throw(res, errmsg)
        jsondata = Nothing
        Return res
      End If

      If timeIdx > 0 Then
        query = "GET " + _dataLoggerURL + "?run=" + LTrim(Str(runIdx)) + "&time=" + LTrim(Str(timeIdx)) + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      Else
        query = "GET " + _dataLoggerURL + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      End If

      res = dev.HTTPRequest(query, buffer, errmsg)
      REM make sure a device scan does not solve the issue
      If YISERR(res) Then
        res = yapiUpdateDeviceList(1, errmsg)
        If YISERR(res) Then
          _throw(res, errmsg)
          jsondata = Nothing
          Return res
        End If

        res = dev.HTTPRequest("GET " + _dataLoggerURL + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10), buffer, errmsg)
        If YISERR(res) Then
          _throw(res, errmsg)
          jsondata = Nothing
          Return res
        End If
      End If

      Dim httpcode As Integer = YAPI.ParseHTTP(buffer, 0, buffer.Length, http_headerlen, errmsg)
      If httpcode = 404 AndAlso _dataLoggerURL <> "/dataLogger.json" Then
        REM retry using backward-compatible datalogger URL
        _dataLoggerURL = "/dataLogger.json"
        Return Me.getData(runIdx, timeIdx, jsondata)
      End If

      If httpcode <> 200 Then
        jsondata = Nothing
        Return YAPI.IO_ERROR
      End If
      Try
        jsondata = YJSONContent.ParseJson(buffer, http_headerlen, buffer.Length)
      Catch E As Exception
        errmsg = "unexpected JSON structure: " + E.Message
        _throw(YAPI_IO_ERROR, errmsg)
        jsondata = Nothing
        Return YAPI_IO_ERROR
      End Try
      Return YAPI_SUCCESS
    End Function


  End Class

  REM --- (generated code: YDataLogger functions)

  '''*
  ''' <summary>
  '''   Retrieves a data logger for a given identifier.
  ''' <para>
  '''   The identifier can be specified using several formats:
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   - FunctionLogicalName
  ''' </para>
  ''' <para>
  '''   - ModuleSerialNumber.FunctionIdentifier
  ''' </para>
  ''' <para>
  '''   - ModuleSerialNumber.FunctionLogicalName
  ''' </para>
  ''' <para>
  '''   - ModuleLogicalName.FunctionIdentifier
  ''' </para>
  ''' <para>
  '''   - ModuleLogicalName.FunctionLogicalName
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   This function does not require that the data logger is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YDataLogger.isOnline()</c> to test if the data logger is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a data logger by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' <para>
  '''   If a call to this object's is_online() method returns FALSE although
  '''   you are certain that the matching device is plugged, make sure that you did
  '''   call registerHub() at application initialization time.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the data logger, for instance
  '''   <c>LIGHTMK4.dataLogger</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YDataLogger</c> object allowing you to drive the data logger.
  ''' </returns>
  '''/
  Public Function yFindDataLogger(ByVal func As String) As YDataLogger
    Return YDataLogger.FindDataLogger(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of data loggers currently accessible.
  ''' <para>
  '''   Use the method <c>YDataLogger.nextDataLogger()</c> to iterate on
  '''   next data loggers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YDataLogger</c> object, corresponding to
  '''   the first data logger currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstDataLogger() As YDataLogger
    Return YDataLogger.FirstDataLogger()
  End Function


  REM --- (end of generated code: YDataLogger functions)


  '''*
  ''' <summary>
  '''   Disables the use of exceptions to report runtime errors.
  ''' <para>
  '''   When exceptions are disabled, every function returns a specific
  '''   error value which depends on its type and which is documented in
  '''   this reference manual.
  ''' </para>
  ''' </summary>
  '''/
  Public Sub yDisableExceptions()
    YAPI.DisableExceptions()
  End Sub

  '''*
  ''' <summary>
  '''   Re-enables the use of exceptions for runtime error handling.
  ''' <para>
  '''   Be aware than when exceptions are enabled, every function that fails
  '''   triggers an exception. If the exception is not caught by the user code,
  '''   it either fires the debugger or aborts (i.e. crash) the program.
  ''' </para>
  ''' </summary>
  '''/
  Public Sub yEnableExceptions()
    YAPI.EnableExceptions()
  End Sub

  REM - Internal callback registered into YAPI using a protected delegate
  Private Sub native_yLogFunction(ByVal log As IntPtr, ByVal loglen As yu32)
    If Not ylog Is Nothing Then ylog(Marshal.PtrToStringAnsi(log))
  End Sub

  '''*
  ''' <summary>
  '''   Registers a log callback function.
  ''' <para>
  '''   This callback will be called each time
  '''   the API have something to say. Quite useful to debug the API.
  ''' </para>
  ''' </summary>
  ''' <param name="logfun">
  '''   a procedure taking a string parameter, or <c>Nothing</c>
  '''   to unregister a previously registered  callback.
  ''' </param>
  '''/
  Public Sub yRegisterLogFunction(ByVal logfun As yLogFunc)
    YAPI.RegisterLogFunction(logfun)
  End Sub


  Private Class PlugEvent
    Enum EVTYPE
      ARRIVAL
      REMOVAL
      CHANGE
      HUB_DISCOVERY
    End Enum

    Private _ev As EVTYPE
    Private _module As YModule
    Private _serial As String
    Private _url As String

    Public Sub New(ByVal ev As EVTYPE, ByVal modul As YModule)
      _ev = ev
      _module = modul
    End Sub

    Public Sub New(ByVal serial As String, ByVal url As String)
      _ev = EVTYPE.HUB_DISCOVERY
      _serial = serial
      _url = url
    End Sub

    Public Sub invoke()
      Select Case _ev
        Case EVTYPE.ARRIVAL
          If Not (yArrival Is Nothing) Then
            yArrival(_module)
          End If
        Case EVTYPE.REMOVAL
          If Not (yRemoval Is Nothing) Then
            yRemoval(_module)
          End If
        Case EVTYPE.CHANGE
          If Not (yChange Is Nothing) Then
            yChange(_module)
          End If
        Case EVTYPE.HUB_DISCOVERY
          If Not (_HubDiscoveryCallback Is Nothing) Then
            _HubDiscoveryCallback(_serial, _url)
          End If
      End Select
    End Sub
  End Class

  Private Class DataEvent
    Private _fun As YFunction
    Private _sensor As YSensor
    Private _mod As YModule
    Private _value As String
    Private _report As List(Of Integer)
    Private _timestamp As Double
    Private _duration As Double
    Private _beacon As Integer

    Public Sub New(ByVal fun As YFunction, ByVal value As String)
      _fun = fun
      _mod = Nothing
      _sensor = Nothing
      _value = value
      _report = Nothing
      _timestamp = 0
      _duration = 0
      _beacon = -1
    End Sub

    Public Sub New(ByVal fun As YSensor, ByVal timestamp As Double, ByVal duration As Double, ByVal report As List(Of Integer))
      _sensor = fun
      _fun = Nothing
      _mod = Nothing
      _value = Nothing
      _timestamp = timestamp
      _duration = duration
      _report = report
      _beacon = -1
    End Sub

    Public Sub New(ByVal modul As YModule)
      _fun = Nothing
      _mod = modul
      _sensor = Nothing
      _value = Nothing
      _report = Nothing
      _timestamp = 0
      _duration = 0
      _beacon = -1
    End Sub

    Public Sub New(ByVal modul As YModule, ByVal beacon As Integer)
      _fun = Nothing
      _mod = modul
      _sensor = Nothing
      _value = Nothing
      _report = Nothing
      _timestamp = 0
      _duration = 0
      _beacon = beacon
    End Sub

    Public Sub invoke()
      If (_sensor IsNot Nothing) Then
        Dim measure As YMeasure = _sensor._decodeTimedReport(_timestamp, _duration, _report)
        _sensor._invokeTimedReportCallback(measure)
      ElseIf (_fun IsNot Nothing) Then
        If _value Is Nothing Then
          _fun.isOnline()
        Else
          REM new value
          _fun._invokeValueCallback(_value)
        End If
      ElseIf (_mod IsNot Nothing) Then
        If (_beacon < 0) Then
          _mod._invokeConfigChangeCallback()
        Else
          _mod._invokeBeaconCallback(_beacon)
        End If
      End If
    End Sub
  End Class


  Private Function emptyDeviceSt() As yDeviceSt
    Dim infos As yDeviceSt
    infos.vendorid = 0
    infos.deviceid = 0
    infos.devrelease = 0
    infos.nbinbterfaces = 0
    infos.manufacturer = ""
    infos.productname = ""
    infos.serial = ""
    infos.logicalname = ""
    infos.firmware = ""
    infos.beacon = 0
    emptyDeviceSt = infos
  End Function

  Dim _PlugEvents As List(Of PlugEvent)
  Dim _DataEvents As List(Of DataEvent)

  Private Sub native_yDeviceArrivalCallback(ByVal d As YDEV_DESCR)
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim ev As PlugEvent
    Dim d_ev As DataEvent
    Dim modul As YModule
    Dim errmsg As String = ""
    For i As Integer = 0 To YFunction._ValueCallbackList.Count - 1
      If YFunction._ValueCallbackList(i).get_functionDescriptor() = YFunction.FUNCTIONDESCRIPTOR_INVALID Then
        d_ev = New DataEvent(YFunction._ValueCallbackList(i), Nothing)
        _DataEvents.Add(d_ev)
      End If
    Next
    YDevice.PlugDevice(d)
    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then
      Exit Sub
    End If
    modul = YModule.FindModule(infos.serial + ".module")
    modul.setImmutableAttributes(infos)
    ev = New PlugEvent(PlugEvent.EVTYPE.ARRIVAL, modul)
    If Not (yArrival Is Nothing) Then _PlugEvents.Add(ev)
  End Sub

  '''*
  ''' <summary>
  '''   Register a callback function, to be called each time
  '''   a device is plugged.
  ''' <para>
  '''   This callback will be invoked while <c>yUpdateDeviceList</c>
  '''   is running. You will have to call this function on a regular basis.
  ''' </para>
  ''' </summary>
  ''' <param name="arrivalCallback">
  '''   a procedure taking a <c>YModule</c> parameter, or <c>Nothing</c>
  '''   to unregister a previously registered  callback.
  ''' </param>
  '''/
  Public Sub yRegisterDeviceArrivalCallback(ByVal arrivalCallback As yDeviceUpdateFunc)
    YAPI.RegisterDeviceArrivalCallback(arrivalCallback)
  End Sub

  Private Sub native_yDeviceRemovalCallback(ByVal d As YDEV_DESCR)
    Dim ev As PlugEvent
    Dim modul As YModule
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim errmsg As String = ""
    If (yRemoval Is Nothing) Then Exit Sub
    infos.deviceid = 0
    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then Exit Sub
    modul = YModule.FindModule(infos.serial + ".module")
    ev = New PlugEvent(PlugEvent.EVTYPE.REMOVAL, modul)
    _PlugEvents.Add(ev)
  End Sub

  '''*
  ''' <summary>
  '''   Register a callback function, to be called each time
  '''   a device is unplugged.
  ''' <para>
  '''   This callback will be invoked while <c>yUpdateDeviceList</c>
  '''   is running. You will have to call this function on a regular basis.
  ''' </para>
  ''' </summary>
  ''' <param name="removalCallback">
  '''   a procedure taking a <c>YModule</c> parameter, or <c>Nothing</c>
  '''   to unregister a previously registered  callback.
  ''' </param>
  '''/
  Public Sub yRegisterDeviceRemovalCallback(ByVal removalCallback As yDeviceUpdateFunc)
    YAPI.RegisterDeviceRemovalCallback(removalCallback)
  End Sub

  Public Sub native_HubDiscoveryCallback(ByVal serial_ptr As IntPtr, ByVal url_ptr As IntPtr)
    Dim ev As PlugEvent
    Dim serial As String = Marshal.PtrToStringAnsi(serial_ptr)
    Dim url As String = Marshal.PtrToStringAnsi(url_ptr)
    ev = New PlugEvent(serial, url)
    _PlugEvents.Add(ev)
  End Sub


  '''*
  ''' <summary>
  '''   Register a callback function, to be called each time an Network Hub send
  '''   an SSDP message.
  ''' <para>
  '''   The callback has two string parameter, the first one
  '''   contain the serial number of the hub and the second contain the URL of the
  '''   network hub (this URL can be passed to RegisterHub). This callback will be invoked
  '''   while yUpdateDeviceList is running. You will have to call this function on a regular basis.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="hubDiscoveryCallback">
  '''   a procedure taking two string parameter, the serial
  '''   number and the hub URL. Use <c>Nothing</c> to unregister a previously registered  callback.
  ''' </param>
  '''/
  Public Sub yRegisterHubDiscoveryCallback(ByVal hubDiscoveryCallback As YHubDiscoveryCallback)
    YAPI.RegisterHubDiscoveryCallback(hubDiscoveryCallback)
  End Sub


  Public Sub native_yDeviceChangeCallback(ByVal d As YDEV_DESCR)
    Dim ev As PlugEvent
    Dim modul As YModule
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim errmsg As String = ""

    If (yChange Is Nothing) Then Exit Sub
    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then Exit Sub
    modul = YModule.FindModule(infos.serial + ".module")
    ev = New PlugEvent(PlugEvent.EVTYPE.CHANGE, modul)
    _PlugEvents.Add(ev)
  End Sub

  Public Sub yRegisterDeviceChangeCallback(ByVal callback As yDeviceUpdateFunc)
    YAPI.RegisterDeviceChangeCallback(callback)
  End Sub

  Public Sub native_yDeviceConfigChangeCallback(ByVal d As YDEV_DESCR)
    Dim ev As DataEvent
    Dim modul As YModule
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim errmsg As String = ""

    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then Exit Sub
    modul = YModule.FindModule(infos.serial + ".module")
    If (YModule._moduleCallbackList.ContainsKey(modul) AndAlso YModule._moduleCallbackList(modul) > 0) Then
      ev = New DataEvent(modul)
      _DataEvents.Add(ev)
    End If
  End Sub


  Public Sub native_yBeaconChangeCallback(ByVal d As YDEV_DESCR, ByVal beacon As Integer)

    Dim ev As DataEvent
    Dim modul As YModule
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim errmsg As String = ""

    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then Exit Sub
    modul = YModule.FindModule(infos.serial + ".module")
    If (YModule._moduleCallbackList.ContainsKey(modul) AndAlso YModule._moduleCallbackList(modul) > 0) Then
      ev = New DataEvent(modul, beacon)
      _DataEvents.Add(ev)
    End If
  End Sub


  Private Sub queuesCleanUp()
    _PlugEvents.Clear()
    _PlugEvents = Nothing
    _DataEvents.Clear()
    _DataEvents = Nothing
  End Sub

  Private Sub native_yFunctionUpdateCallback(ByVal fundescr As YFUN_DESCR, ByVal data As IntPtr)
    Dim ev As DataEvent
    If (Not (data = IntPtr.Zero)) Then
      For i As Integer = 0 To YFunction._ValueCallbackList.Count - 1
        If (YFunction._ValueCallbackList(i).get_functionDescriptor() = fundescr) Then
          ev = New DataEvent(YFunction._ValueCallbackList(i), Marshal.PtrToStringAnsi(data))
          _DataEvents.Add(ev)
          Return
        End If
      Next i
    End If
  End Sub


  Private Sub native_yTimedReportCallback(ByVal fundescr As YFUN_DESCR, timestamp As Double, rawdata As IntPtr, len As System.UInt32, duration As Double)
    Dim ev As DataEvent
    Dim data As Byte()
    Dim report As List(Of Integer)
    Dim intlen As Integer = CInt(len)
    Dim p As Integer
    For i As Integer = 0 To YFunction._TimedReportCallbackList.Count - 1
      If (YFunction._TimedReportCallbackList(i).get_functionDescriptor() = fundescr) Then
        ReDim data(intlen)
        Marshal.Copy(rawdata, data, 0, intlen)
        If (data(0) <= 2) Then
          report = New List(Of Integer)(intlen)
          p = 0
          While (p < intlen)
            report.Add(data(p))
            p = p + 1
          End While
          ev = New DataEvent(YFunction._TimedReportCallbackList(i), timestamp, duration, report)
          _DataEvents.Add(ev)
        End If
        Return
      End If
    Next i
  End Sub


  Private Function yapiLockDeviceCallBack(ByRef errmsg As String) As Integer
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiLockDeviceCallBack = _yapiLockDeviceCallBack(buffer)
    errmsg = buffer.ToString()
  End Function

  Private Function yapiUnlockDeviceCallBack(ByRef errmsg As String) As Integer
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiUnlockDeviceCallBack = _yapiUnlockDeviceCallBack(buffer)
    errmsg = buffer.ToString()
  End Function

  Private Function yapiLockFunctionCallBack(ByRef errmsg As String) As Integer
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiLockFunctionCallBack = _yapiLockFunctionCallBack(buffer)
    errmsg = buffer.ToString()
  End Function

  Private Function yapiUnlockFunctionCallBack(ByRef errmsg As String) As Integer
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiUnlockFunctionCallBack = _yapiUnlockFunctionCallBack(buffer)
    errmsg = buffer.ToString()
  End Function


  Private Sub native_yDeviceLogCallback(ByVal devdescr As YFUN_DESCR, ByVal data As IntPtr)
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim modul As YModule
    Dim errmsg As String = ""
    Dim callback As YModuleLogCallback

    If (yapiGetDeviceInfo(devdescr, infos, errmsg) <> YAPI_SUCCESS) Then
      Exit Sub
    End If
    modul = YModule.FindModule(infos.serial + ".module")
    callback = modul.get_logCallback()
    If Not (callback Is Nothing) Then
      callback(modul, Marshal.PtrToStringAnsi(data))
    End If
  End Sub


  REM - Delegate object for our internal callback, protected from GC
  Public native_yLogFunctionDelegate As _yapiLogFunc = AddressOf native_yLogFunction
  Dim native_yLogFunctionAnchor As GCHandle = GCHandle.Alloc(native_yLogFunctionDelegate)

  Public native_yFunctionUpdateDelegate As _yapiFunctionUpdateFunc = AddressOf native_yFunctionUpdateCallback
  Dim native_yFunctionUpdateAnchor As GCHandle = GCHandle.Alloc(native_yFunctionUpdateDelegate)

  Public native_yTimedReportDelegate As _yapiTimedReportFunc = AddressOf native_yTimedReportCallback
  Dim native_yTimedReportAnchor As GCHandle = GCHandle.Alloc(native_yTimedReportDelegate)

  Public native_yDeviceArrivalDelegate As _yapiDeviceUpdateFunc = AddressOf native_yDeviceArrivalCallback
  Dim native_yDeviceArrivalAnchor As GCHandle = GCHandle.Alloc(native_yDeviceArrivalDelegate)

  Public native_yDeviceRemovalDelegate As _yapiDeviceUpdateFunc = AddressOf native_yDeviceRemovalCallback
  Dim native_yDeviceRemovalAnchor As GCHandle = GCHandle.Alloc(native_yDeviceRemovalDelegate)

  Public native_yDeviceChangeDelegate As _yapiDeviceUpdateFunc = AddressOf native_yDeviceChangeCallback
  Dim native_yDeviceChangeAnchor As GCHandle = GCHandle.Alloc(native_yDeviceChangeDelegate)

  Public native_yDeviceConfigChangeDelegate As _yapiDeviceUpdateFunc = AddressOf native_yDeviceConfigChangeCallback
  Dim native_yDeviceConfigChangeAnchor As GCHandle = GCHandle.Alloc(native_yDeviceConfigChangeDelegate)

  Public native_yBeaconChangeDelegate As _yapiBeaconUpdateFunc = AddressOf native_yBeaconChangeCallback
  Dim native_yBeaconChangeAnchor As GCHandle = GCHandle.Alloc(native_yBeaconChangeDelegate)

  Public native_HubDiscoveryDelegate As _yapiHubDiscoveryCallback = AddressOf native_HubDiscoveryCallback
  Dim native_HubDiscoveryAnchor As GCHandle = GCHandle.Alloc(native_HubDiscoveryDelegate)

  Public native_yDeviceLogDelegate As _yapiDeviceLogCallback = AddressOf native_yDeviceLogCallback
  Dim native_yDeviceLogAnchor As GCHandle = GCHandle.Alloc(native_yDeviceLogDelegate)

  '''*
  ''' <summary>
  '''   Returns the version identifier for the Yoctopuce library in use.
  ''' <para>
  '''   The version is a string in the form <c>"Major.Minor.Build"</c>,
  '''   for instance <c>"1.01.5535"</c>. For languages using an external
  '''   DLL (for instance C#, VisualBasic or Delphi), the character string
  '''   includes as well the DLL version, for instance
  '''   <c>"1.01.5535 (1.01.5439)"</c>.
  ''' </para>
  ''' <para>
  '''   If you want to verify in your code that the library version is
  '''   compatible with the version that you have used during development,
  '''   verify that the major number is strictly equal and that the minor
  '''   number is greater or equal. The build number is not relevant
  '''   with respect to the library compatibility.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a character string describing the library version.
  ''' </returns>
  '''/
  Public Function yGetAPIVersion() As String
    Return YAPI.GetAPIVersion()
  End Function


  '''*
  '''/
  Public Const Y_DETECT_NONE As Integer = YAPI.DETECT_NONE
  Public Const Y_DETECT_USB As Integer = YAPI.DETECT_USB
  Public Const Y_DETECT_NET As Integer = YAPI.DETECT_NET
  Public Const Y_DETECT_ALL As Integer = YAPI.DETECT_ALL
  Public Const Y_RESEND_MISSING_PKT As Integer = YAPI.RESEND_MISSING_PKT
  Public Function yInitAPI(ByVal mode As Integer, ByRef errmsg As String) As Integer
    Return YAPI.InitAPI(mode, errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Waits for all pending communications with Yoctopuce devices to be
  '''   completed then frees dynamically allocated resources used by
  '''   the Yoctopuce library.
  ''' <para>
  ''' </para>
  ''' <para>
  '''   From an operating system standpoint, it is generally not required to call
  '''   this function since the OS will automatically free allocated resources
  '''   once your program is completed. However, there are two situations when
  '''   you may really want to use that function:
  ''' </para>
  ''' <para>
  '''   - Free all dynamically allocated memory blocks in order to
  '''   track a memory leak.
  ''' </para>
  ''' <para>
  '''   - Send commands to devices right before the end
  '''   of the program. Since commands are sent in an asynchronous way
  '''   the program could exit before all commands are effectively sent.
  ''' </para>
  ''' <para>
  '''   You should not call any other library function after calling
  '''   <c>yFreeAPI()</c>, or your program will crash.
  ''' </para>
  ''' </summary>
  '''/
  Public Sub yFreeAPI()
    YAPI.FreeAPI()
  End Sub

  '''*
  ''' <summary>
  '''   Set up the Yoctopuce library to use modules connected on a given machine.
  ''' <para>
  '''   Idealy this
  '''   call will be made once at the begining of your application.  The
  '''   parameter will determine how the API will work. Use the following values:
  ''' </para>
  ''' <para>
  '''   <b>usb</b>: When the <c>usb</c> keyword is used, the API will work with
  '''   devices connected directly to the USB bus. Some programming languages such a JavaScript,
  '''   PHP, and Java don't provide direct access to USB hardware, so <c>usb</c> will
  '''   not work with these. In this case, use a VirtualHub or a networked YoctoHub (see below).
  ''' </para>
  ''' <para>
  '''   <b><i>x.x.x.x</i></b> or <b><i>hostname</i></b>: The API will use the devices connected to the
  '''   host with the given IP address or hostname. That host can be a regular computer
  '''   running a <i>native VirtualHub</i>, a <i>VirtualHub for web</i> hosted on a server,
  '''   or a networked YoctoHub such as YoctoHub-Ethernet or
  '''   YoctoHub-Wireless. If you want to use the VirtualHub running on you local
  '''   computer, use the IP address 127.0.0.1. If the given IP is unresponsive, <c>yRegisterHub</c>
  '''   will not return until a time-out defined by <c>ySetNetworkTimeout</c> has elapsed.
  '''   However, it is possible to preventively test a connection  with <c>yTestHub</c>.
  '''   If you cannot afford a network time-out, you can use the non-blocking <c>yPregisterHub</c>
  '''   function that will establish the connection as soon as it is available.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' <para>
  '''   <b>callback</b>: that keyword make the API run in "<i>HTTP Callback</i>" mode.
  '''   This a special mode allowing to take control of Yoctopuce devices
  '''   through a NAT filter when using a VirtualHub or a networked YoctoHub. You only
  '''   need to configure your hub to call your server script on a regular basis.
  '''   This mode is currently available for PHP and Node.JS only.
  ''' </para>
  ''' <para>
  '''   Be aware that only one application can use direct USB access at a
  '''   given time on a machine. Multiple access would cause conflicts
  '''   while trying to access the USB modules. In particular, this means
  '''   that you must stop the VirtualHub software before starting
  '''   an application that uses direct USB access. The workaround
  '''   for this limitation is to set up the library to use the VirtualHub
  '''   rather than direct USB access.
  ''' </para>
  ''' <para>
  '''   If access control has been activated on the hub, virtual or not, you want to
  '''   reach, the URL parameter should look like:
  ''' </para>
  ''' <para>
  '''   <c>http://username:password@address:port</c>
  ''' </para>
  ''' <para>
  '''   You can call <i>RegisterHub</i> several times to connect to several machines. On
  '''   the other hand, it is useless and even counterproductive to call <i>RegisterHub</i>
  '''   with to same address multiple times during the life of the application.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="url">
  '''   a string containing either <c>"usb"</c>,<c>"callback"</c> or the
  '''   root URL of the hub to monitor
  ''' </param>
  ''' <param name="errmsg">
  '''   a string passed by reference to receive any error message.
  ''' </param>
  ''' <returns>
  '''   <c>YAPI.SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure returns a negative error code.
  ''' </para>
  '''/
  Public Function yRegisterHub(ByVal url As String, ByRef errmsg As String) As Integer
    Return YAPI.RegisterHub(url, errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Fault-tolerant alternative to <c>yRegisterHub()</c>.
  ''' <para>
  '''   This function has the same
  '''   purpose and same arguments as <c>yRegisterHub()</c>, but does not trigger
  '''   an error when the selected hub is not available at the time of the function call.
  '''   If the connexion cannot be established immediately, a background task will automatically
  '''   perform periodic retries. This makes it possible to register a network hub independently of the current
  '''   connectivity, and to try to contact it only when a device is actively needed.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="url">
  '''   a string containing either <c>"usb"</c>,<c>"callback"</c> or the
  '''   root URL of the hub to monitor
  ''' </param>
  ''' <param name="errmsg">
  '''   a string passed by reference to receive any error message.
  ''' </param>
  ''' <returns>
  '''   <c>YAPI.SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure returns a negative error code.
  ''' </para>
  '''/
  Public Function yPreregisterHub(ByVal url As String, ByRef errmsg As String) As Integer
    Return YAPI.PreregisterHub(url, errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Set up the Yoctopuce library to no more use modules connected on a previously
  '''   registered machine with RegisterHub.
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="url">
  '''   a string containing either <c>"usb"</c> or the
  '''   root URL of the hub to monitor
  ''' </param>
  '''/
  Public Sub yUnregisterHub(ByVal url As String)
    YAPI.UnregisterHub(url)
  End Sub

  '''*
  ''' <summary>
  '''   Test if the hub is reachable.
  ''' <para>
  '''   This method do not register the hub, it only test if the
  '''   hub is usable. The url parameter follow the same convention as the <c>yRegisterHub</c>
  '''   method. This method is useful to verify the authentication parameters for a hub. It
  '''   is possible to force this method to return after mstimeout milliseconds.
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="url">
  '''   a string containing either <c>"usb"</c>,<c>"callback"</c> or the
  '''   root URL of the hub to monitor
  ''' </param>
  ''' <param name="mstimeout">
  '''   the number of millisecond available to test the connection.
  ''' </param>
  ''' <param name="errmsg">
  '''   a string passed by reference to receive any error message.
  ''' </param>
  ''' <returns>
  '''   <c>YAPI.SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure returns a negative error code.
  ''' </para>
  '''/
  Public Function yTestHub(ByVal url As String, ByVal mstimeout As Integer, ByRef errmsg As String) As Integer
    Return YAPI.TestHub(url, mstimeout, errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Triggers a (re)detection of connected Yoctopuce modules.
  ''' <para>
  '''   The library searches the machines or USB ports previously registered using
  '''   <c>yRegisterHub()</c>, and invokes any user-defined callback function
  '''   in case a change in the list of connected devices is detected.
  ''' </para>
  ''' <para>
  '''   This function can be called as frequently as desired to refresh the device list
  '''   and to make the application aware of hot-plug events. However, since device
  '''   detection is quite a heavy process, UpdateDeviceList shouldn't be called more
  '''   than once every two seconds.
  ''' </para>
  ''' </summary>
  ''' <param name="errmsg">
  '''   a string passed by reference to receive any error message.
  ''' </param>
  ''' <returns>
  '''   <c>YAPI.SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure returns a negative error code.
  ''' </para>
  '''/
  Public Function yUpdateDeviceList(ByRef errmsg As String) As YRETCODE
    Return YAPI.UpdateDeviceList(errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Maintains the device-to-library communication channel.
  ''' <para>
  '''   If your program includes significant loops, you may want to include
  '''   a call to this function to make sure that the library takes care of
  '''   the information pushed by the modules on the communication channels.
  '''   This is not strictly necessary, but it may improve the reactivity
  '''   of the library for the following commands.
  ''' </para>
  ''' <para>
  '''   This function may signal an error in case there is a communication problem
  '''   while contacting a module.
  ''' </para>
  ''' </summary>
  ''' <param name="errmsg">
  '''   a string passed by reference to receive any error message.
  ''' </param>
  ''' <returns>
  '''   <c>YAPI.SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure returns a negative error code.
  ''' </para>
  '''/
  Public Function yHandleEvents(ByRef errmsg As String) As YRETCODE
    Return YAPI.HandleEvents(errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Pauses the execution flow for a specified duration.
  ''' <para>
  '''   This function implements a passive waiting loop, meaning that it does not
  '''   consume CPU cycles significantly. The processor is left available for
  '''   other threads and processes. During the pause, the library nevertheless
  '''   reads from time to time information from the Yoctopuce modules by
  '''   calling <c>yHandleEvents()</c>, in order to stay up-to-date.
  ''' </para>
  ''' <para>
  '''   This function may signal an error in case there is a communication problem
  '''   while contacting a module.
  ''' </para>
  ''' </summary>
  ''' <param name="ms_duration">
  '''   an integer corresponding to the duration of the pause,
  '''   in milliseconds.
  ''' </param>
  ''' <param name="errmsg">
  '''   a string passed by reference to receive any error message.
  ''' </param>
  ''' <returns>
  '''   <c>YAPI.SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure returns a negative error code.
  ''' </para>
  '''/
  Public Function ySleep(ByVal ms_duration As Integer, ByRef errmsg As String) As Integer
    Return YAPI.Sleep(ms_duration, errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Force a hub discovery, if a callback as been registered with <c>yRegisterHubDiscoveryCallback</c> it
  '''   will be called for each net work hub that will respond to the discovery.
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="errmsg">
  '''   a string passed by reference to receive any error message.
  ''' </param>
  ''' <returns>
  '''   <c>YAPI.SUCCESS</c> when the call succeeds.
  '''   On failure returns a negative error code.
  ''' </returns>
  '''/
  Public Function yTriggerHubDiscovery(ByRef errmsg As String) As Integer
    Return YAPI.TriggerHubDiscovery(errmsg)
  End Function


  '''*
  ''' <summary>
  '''   Returns the current value of a monotone millisecond-based time counter.
  ''' <para>
  '''   This counter can be used to compute delays in relation with
  '''   Yoctopuce devices, which also uses the millisecond as timebase.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a long integer corresponding to the millisecond counter.
  ''' </returns>
  '''/
  Public Function yGetTickCount() As Long
    Return YAPI.GetTickCount()
  End Function

  '''*
  ''' <summary>
  '''   Checks if a given string is valid as logical name for a module or a function.
  ''' <para>
  '''   A valid logical name has a maximum of 19 characters, all among
  '''   <c>A...Z</c>, <c>a...z</c>, <c>0...9</c>, <c>_</c>, and <c>-</c>.
  '''   If you try to configure a logical name with an incorrect string,
  '''   the invalid characters are ignored.
  ''' </para>
  ''' </summary>
  ''' <param name="name">
  '''   a string containing the name to check.
  ''' </param>
  ''' <returns>
  '''   <c>true</c> if the name is valid, <c>false</c> otherwise.
  ''' </returns>
  '''/
  Public Function yCheckLogicalName(ByVal name As String) As Boolean
    Return YAPI.CheckLogicalName(name)
  End Function


  Public Function yapiGetFunctionInfo(ByVal fundesc As YFUN_DESCR,
                                        ByRef devdesc As YDEV_DESCR,
                                        ByRef serial As String,
                                        ByRef funcId As String,
                                        ByRef funcName As String,
                                        ByRef funcVal As String,
                                        ByRef errmsg As String) As Integer

    Dim serialBuffer As New StringBuilder(YOCTO_SERIAL_LEN)
    Dim funcIdBuffer As New StringBuilder(YOCTO_FUNCTION_LEN)
    Dim funcNameBuffer As New StringBuilder(YOCTO_LOGICAL_LEN)
    Dim funcValBuffer As New StringBuilder(YOCTO_PUBVAL_LEN)
    Dim errBuffer As New StringBuilder(YOCTO_ERRMSG_LEN)

    serialBuffer.Length = 0
    funcIdBuffer.Length = 0
    funcNameBuffer.Length = 0
    funcValBuffer.Length = 0
    errBuffer.Length = 0

    yapiGetFunctionInfo = _yapiGetFunctionInfoEx(fundesc, devdesc, serialBuffer, funcIdBuffer, Nothing, funcNameBuffer, funcValBuffer, errBuffer)
    serial = serialBuffer.ToString()
    funcId = funcIdBuffer.ToString()
    funcName = funcNameBuffer.ToString()
    funcVal = funcValBuffer.ToString()
    errmsg = funcValBuffer.ToString()
  End Function

  Public Function yapiGetFunctionInfoEx(ByVal fundesc As YFUN_DESCR,
                                        ByRef devdesc As YDEV_DESCR,
                                        ByRef serial As String,
                                        ByRef funcId As String,
                                        ByRef baseType As String,
                                        ByRef funcName As String,
                                        ByRef funcVal As String,
                                        ByRef errmsg As String) As Integer

    Dim serialBuffer As New StringBuilder(YOCTO_SERIAL_LEN)
    Dim funcIdBuffer As New StringBuilder(YOCTO_FUNCTION_LEN)
    Dim baseTypeBuffer As New StringBuilder(YOCTO_FUNCTION_LEN)
    Dim funcNameBuffer As New StringBuilder(YOCTO_LOGICAL_LEN)
    Dim funcValBuffer As New StringBuilder(YOCTO_PUBVAL_LEN)
    Dim errBuffer As New StringBuilder(YOCTO_ERRMSG_LEN)

    serialBuffer.Length = 0
    funcIdBuffer.Length = 0
    funcNameBuffer.Length = 0
    funcValBuffer.Length = 0
    errBuffer.Length = 0

    yapiGetFunctionInfoEx = _yapiGetFunctionInfoEx(fundesc, devdesc, serialBuffer, funcIdBuffer, baseTypeBuffer, funcNameBuffer, funcValBuffer, errBuffer)
    serial = serialBuffer.ToString()
    funcId = funcIdBuffer.ToString()
    baseType = baseTypeBuffer.ToString()
    funcName = funcNameBuffer.ToString()
    funcVal = funcValBuffer.ToString()
    errmsg = funcValBuffer.ToString()
  End Function


  Public Function yapiGetDeviceByFunction(ByVal fundesc As YFUN_DESCR, ByRef errmsg As String) As Integer
    Dim errBuffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    Dim devdesc As YDEV_DESCR
    Dim res As Integer
    errBuffer.Length = 0
    res = _yapiGetFunctionInfoEx(fundesc, devdesc, Nothing, Nothing, Nothing, Nothing, Nothing, errBuffer)
    errmsg = errBuffer.ToString
    If (res < 0) Then
      yapiGetDeviceByFunction = res
    Else
      yapiGetDeviceByFunction = devdesc
    End If
  End Function

  Public Function yapiGetAPIVersion(ByRef version As String, ByRef subversion As String) As yu16
    Dim pversion As IntPtr
    Dim psubversion As IntPtr
    Dim res As yu16
    res = _yapiGetAPIVersion(pversion, psubversion)
    version = Marshal.PtrToStringAnsi(pversion)
    subversion = Marshal.PtrToStringAnsi(psubversion)
    yapiGetAPIVersion = res
  End Function

  Public Function yapiUpdateDeviceList(ByVal force As Integer, ByRef errmsg As String) As YRETCODE
    Dim buffer As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    Dim res As YRETCODE = _yapiUpdateDeviceList(force, buffer)
    If (YISERR(res)) Then
      errmsg = buffer.ToString()
    End If
    Return res
  End Function


  Public Function yapiGetDevice(ByRef device_str As String, ByVal errmsg As String) As YDEV_DESCR
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiGetDevice = _yapiGetDevice(New StringBuilder(device_str), buffer)
    errmsg = buffer.ToString()
  End Function

  Public Function yapiGetDeviceInfo(ByVal d As YDEV_DESCR, ByRef infos As yDeviceSt, ByRef errmsg As String) As Integer
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiGetDeviceInfo = _yapiGetDeviceInfo(d, infos, buffer)
    errmsg = buffer.ToString()
  End Function

  Public Function yapiGetFunction(ByVal class_str As String, ByVal function_str As String, ByRef errmsg As String) As YFUN_DESCR
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiGetFunction = _yapiGetFunction(New StringBuilder(class_str), New StringBuilder(function_str), buffer)
    errmsg = buffer.ToString()
  End Function

  Public Function yapiGetFunctionsByClass(ByVal class_str As String, ByVal precFuncDesc As YFUN_DESCR, ByVal dbuffer As IntPtr, ByVal maxsize As Integer, ByRef neededsize As Integer, ByRef errmsg As String) As Integer
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiGetFunctionsByClass = _yapiGetFunctionsByClass(New StringBuilder(class_str), precFuncDesc, dbuffer, maxsize, neededsize, buffer)
    errmsg = buffer.ToString()
  End Function

  Function yapiGetFunctionsByDevice(ByVal devdesc As YDEV_DESCR, ByVal precFuncDesc As YFUN_DESCR, ByVal dbuffer As IntPtr, ByVal maxsize As Integer, ByRef neededsize As Integer, ByRef errmsg As String) As Integer
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiGetFunctionsByDevice = _yapiGetFunctionsByDevice(devdesc, precFuncDesc, dbuffer, maxsize, neededsize, buffer)
    errmsg = buffer.ToString()
  End Function

  <DllImport("myDll.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub DllCallTest(ByRef data As yDeviceSt)
  End Sub

  REM --- (generated code: YFunction dlldef)
  <DllImport("yapi.dll", EntryPoint:="yapiInitAPI", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiInitAPI(ByVal mode As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiFreeAPI", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiFreeAPI()
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiSetTraceFile", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiSetTraceFile(ByVal tracefile As StringBuilder)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterLogFunction", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterLogFunction(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterDeviceArrivalCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterDeviceArrivalCallback(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterDeviceRemovalCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterDeviceRemovalCallback(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterDeviceChangeCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterDeviceChangeCallback(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterDeviceConfigChangeCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterDeviceConfigChangeCallback(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterFunctionUpdateCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterFunctionUpdateCallback(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterTimedReportCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterTimedReportCallback(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiLockDeviceCallBack", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiLockDeviceCallBack(ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiUnlockDeviceCallBack", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiUnlockDeviceCallBack(ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiLockFunctionCallBack", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiLockFunctionCallBack(ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiUnlockFunctionCallBack", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiUnlockFunctionCallBack(ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterHub", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiRegisterHub(ByVal rootUrl As StringBuilder, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiPreregisterHub", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiPreregisterHub(ByVal rootUrl As StringBuilder, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiUnregisterHub", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiUnregisterHub(ByVal rootUrl As StringBuilder)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiUpdateDeviceList", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiUpdateDeviceList(ByVal force As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHandleEvents", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHandleEvents(ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetTickCount", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetTickCount() As yu64
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiCheckLogicalName", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiCheckLogicalName(ByVal name As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetAPIVersion", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetAPIVersion(ByRef version As IntPtr, ByRef dat_ As IntPtr) As yu16
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetDevice", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetDevice(ByVal device_str As StringBuilder, ByVal errmsg As StringBuilder) As YDEV_DESCR
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetDeviceInfo", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetDeviceInfo(ByVal d As YDEV_DESCR, ByRef infos As yDeviceSt, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetFunction", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetFunction(ByVal class_str As StringBuilder, ByVal function_str As StringBuilder, ByVal errmsg As StringBuilder) As YFUN_DESCR
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetFunctionsByClass", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetFunctionsByClass(ByVal class_str As StringBuilder, ByVal precFuncDesc As YFUN_DESCR, ByVal buffer As IntPtr, ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetFunctionsByDevice", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetFunctionsByDevice(ByVal device As YDEV_DESCR, ByVal precFuncDesc As YFUN_DESCR, ByVal buffer As IntPtr, ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetFunctionInfoEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetFunctionInfoEx(ByVal fundesc As YFUN_DESCR, ByRef devdesc As YDEV_DESCR, ByVal serial As StringBuilder, ByVal funcId As StringBuilder, ByVal baseType As StringBuilder, ByVal funcName As StringBuilder, ByVal funcVal As StringBuilder, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncStart", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestSyncStart(ByRef iohdl As YIOHDL, ByVal device As StringBuilder, ByVal request As StringBuilder, ByRef reply As IntPtr, ByRef replysize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncStartEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestSyncStartEx(ByRef iohdl As YIOHDL, ByVal device As StringBuilder, ByVal request As IntPtr, ByVal requestlen As Integer, ByRef reply As IntPtr, ByRef replysize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncDone", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestSyncDone(ByRef iohdl As YIOHDL, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestAsync", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestAsync(ByVal device As StringBuilder, ByVal request As IntPtr, ByVal callback As IntPtr, ByRef context As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestAsyncEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestAsyncEx(ByVal device As StringBuilder, ByVal request As IntPtr, ByVal requestlen As Integer, ByVal callback As IntPtr, ByRef context As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequest", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequest(ByVal device As StringBuilder, ByVal url As StringBuilder, ByVal buffer As StringBuilder, ByVal buffsize As Integer, ByRef fullsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetDevicePath", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetDevicePath(ByVal devdesc As Integer, ByVal rootdevice As StringBuilder, ByVal path As StringBuilder, ByVal pathsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiSleep", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiSleep(ByVal duration_ms As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterHubDiscoveryCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterHubDiscoveryCallback(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiTriggerHubDiscovery", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiTriggerHubDiscovery(ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterDeviceLogCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterDeviceLogCallback(ByVal fct As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiGetAllJsonKeys", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetAllJsonKeys(ByVal jsonbuffer As StringBuilder, ByVal out_buffer As StringBuilder, ByVal out_buffersize As Integer, ByRef fullsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiCheckFirmware", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiCheckFirmware(ByVal serial As StringBuilder, ByVal rev As StringBuilder, ByVal path As StringBuilder, ByVal buffer As StringBuilder, ByVal buffersize As Integer, ByRef fullsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetBootloaders", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetBootloaders(ByVal buffer As StringBuilder, ByVal buffersize As Integer, ByRef totalSize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiUpdateFirmwareEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiUpdateFirmwareEx(ByVal serial As StringBuilder, ByVal firmwarePath As StringBuilder, ByVal settings As StringBuilder, ByVal force As Integer, ByVal startUpdate As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncStartOutOfBand", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestSyncStartOutOfBand(ByRef iohdl As YIOHDL, ByVal channel As Integer, ByVal device As StringBuilder, ByVal request As StringBuilder, ByVal requestsize As Integer, ByRef reply As IntPtr, ByRef replysize As Integer, ByVal progress_cb As IntPtr, ByRef progress_ctx As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestAsyncOutOfBand", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestAsyncOutOfBand(ByVal channel As Integer, ByVal device As StringBuilder, ByVal request As StringBuilder, ByVal requestsize As Integer, ByVal callback As IntPtr, ByRef context As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiTestHub", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiTestHub(ByVal url As StringBuilder, ByVal mstimeout As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiJsonGetPath", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiJsonGetPath(ByVal path As StringBuilder, ByVal json_data As StringBuilder, ByVal json_len As Integer, ByRef result As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiJsonDecodeString", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiJsonDecodeString(ByVal json_data As StringBuilder, ByVal output As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetSubdevices", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetSubdevices(ByVal serial As StringBuilder, ByVal buffer As StringBuilder, ByVal buffersize As Integer, ByRef totalSize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiFreeMem", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiFreeMem(ByRef buffer As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiGetDevicePathEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetDevicePathEx(ByVal serial As StringBuilder, ByVal rootdevice As StringBuilder, ByVal path As StringBuilder, ByVal pathsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiSetNetDevListValidity", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiSetNetDevListValidity(ByVal sValidity As Integer)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiGetNetDevListValidity", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetNetDevListValidity() As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiRegisterBeaconCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiRegisterBeaconCallback(ByVal beaconCallback As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiStartStopDeviceLogCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiStartStopDeviceLogCallback(ByVal serial As StringBuilder, ByVal start As Integer)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiIsModuleWritable", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiIsModuleWritable(ByVal serial As StringBuilder, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetDLLPath", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetDLLPath(ByVal path As StringBuilder, ByVal pathsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiSetNetworkTimeout", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiSetNetworkTimeout(ByVal sValidity As Integer)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiGetNetworkTimeout", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetNetworkTimeout() As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiAddUdevRulesForYocto", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiAddUdevRulesForYocto(ByVal force As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiSetSSLCertificateSrv", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiSetSSLCertificateSrv(ByVal certfile As StringBuilder, ByVal keyfile As StringBuilder, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiAddSSLCertificateCli", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiAddSSLCertificateCli(ByVal cert As StringBuilder, ByVal cert_len As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiSetNetworkSecurityOptions", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiSetNetworkSecurityOptions(ByVal options As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetRemoteCertificate", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetRemoteCertificate(ByVal rooturl As StringBuilder, ByVal timeout As yu64, ByVal buffer As StringBuilder, ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetNextHubRef", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetNextHubRef(ByVal hubref As Integer) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetHubStrAttr", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetHubStrAttr(ByVal hubref As Integer, ByVal attrname As StringBuilder, ByVal attrval As StringBuilder, ByVal maxsize As Integer, ByRef neededsize As Integer) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetHubIntAttr", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetHubIntAttr(ByVal hubref As Integer, ByVal attrname As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiSetHubIntAttr", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiSetHubIntAttr(ByVal hubref As Integer, ByVal attrname As StringBuilder, ByVal value As Integer) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiSetTrustedCertificatesList", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiSetTrustedCertificatesList(ByVal certificatePath As StringBuilder, ByVal errmsg As StringBuilder) As Integer
  End Function
    REM --- (end of generated code: YFunction dlldef)


  Private Sub vbmodule_initialization()
    YDevice_devCache = New List(Of YDevice)
    REM --- (generated code: YFunction initialization)
    REM --- (end of generated code: YFunction initialization)
    REM --- (generated code: YModule initialization)
    REM --- (end of generated code: YModule initialization)
    REM --- (generated code: YDataStream initialization)
    REM --- (end of generated code: YDataStream initialization)
    REM --- (generated code: YMeasure initialization)
    REM --- (end of generated code: YMeasure initialization)
    REM --- (generated code: YDataSet initialization)
    REM --- (end of generated code: YDataSet initialization)
    _PlugEvents = New List(Of PlugEvent)
    _DataEvents = New List(Of DataEvent)
  End Sub

  Private Sub vbmodule_cleanup()
    YDevice_devCache.Clear()
    YDevice_devCache = Nothing
    REM --- (generated code: YFunction cleanup)
    REM --- (end of generated code: YFunction cleanup)
    REM --- (generated code: YModule cleanup)
    REM --- (end of generated code: YModule cleanup)
    REM --- (generated code: YDataStream cleanup)
    REM --- (end of generated code: YDataStream cleanup)
    REM --- (generated code: YMeasure cleanup)
    REM --- (end of generated code: YMeasure cleanup)
    REM --- (generated code: YDataSet cleanup)
    REM --- (end of generated code: YDataSet cleanup)
    _PlugEvents.Clear()
    _PlugEvents = Nothing
    _DataEvents.Clear()
    _DataEvents = Nothing
    YAPI.handlersCleanUp()
  End Sub
End Module
'/********************************************************************
'*
'* $Id: yocto_api.vb 31238 2018-07-17 11:08:47Z mvuilleu $
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
Imports System.Text
Imports System.Math
Imports System.Threading

Module yocto_api
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

    Public MustOverride Function toJSON() As String
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

      If _data(cur_pos) <> "["c Then
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

    Public Overrides Function toJSON() As String
      Dim res As New StringBuilder()
      res.Append("["c)
      Dim sep As String = ""
      For Each yjsonContent As YJSONContent In _arrayValue
        Dim subres As String = yjsonContent.toJSON()
        res.Append(sep)
        res.Append(subres)
        sep = ","
      Next
      res.Append("]"c)
      Return res.ToString()
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

      If _data(cur_pos) <> """"c Then
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

    Public Overrides Function toJSON() As String
      Dim res As New StringBuilder(_stringValue.Length * 2)
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
      Return res.ToString()
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
        End If
        cur_pos += 1
      End While
      Throw New System.Exception(FormatError("unexpected end of data", cur_pos))
    End Function

    Public Overrides Function toJSON() As String
      If _isFloat Then
        Return _doubleValue.ToString()
      Else
        Return _intValue.ToString()
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

      If _data.Length <= cur_pos OrElse _data(cur_pos) <> "{"c Then
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
      Dim ystr As YJSONString = DirectCast(parsed(key), YJSONString)
      Return ystr.getString()
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

    Public Overrides Function toJSON() As String
      Dim res As New StringBuilder()
      res.Append("{"c)
      Dim sep As String = ""
      For Each key As String In parsed.Keys.ToArray()
        Dim subContent As YJSONContent = parsed(key)
        Dim subres As String = subContent.toJSON()
        res.Append(sep)
        res.Append(""""c)
        res.Append(key)
        res.Append(""":")
        res.Append(subres)
        sep = ","
      Next
      res.Append("}"c)
      Return res.ToString()
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

















  Public Const YOCTO_API_VERSION_STR As String = "1.10"
  Public Const YOCTO_API_VERSION_BCD As Integer = &H110
  Public Const YOCTO_API_BUILD_NO As String = "31315"

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

    REM --- (end of generated code: YFunction return codes)



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
    Public Shared Function _decimalToDouble(ByVal val As Integer) As Double
      Dim negate As Boolean = False
      Dim res As Double
      Dim mantis As Integer = val And 2047
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
      Dim exp As Integer = val >> 11
      res = (mantis) * decexp(exp)
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
      Dim rounded As Integer
      Dim decim As Integer

      rounded = CInt(Math.Round(value * 1000))
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
    '''   When <c>Y_DETECT_NONE</c> is used as detection <c>mode</c>,
    '''   you must explicitly use <c>yRegisterHub()</c> to point the API to the
    '''   VirtualHub on which your devices are connected before trying to access them.
    ''' </para>
    ''' </summary>
    ''' <param name="mode">
    '''   an integer corresponding to the type of automatic
    '''   device detection to use. Possible values are
    '''   <c>Y_DETECT_NONE</c>, <c>Y_DETECT_USB</c>, <c>Y_DETECT_NET</c>,
    '''   and <c>Y_DETECT_ALL</c>.
    ''' </param>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Shared Function InitAPI(ByVal mode As Integer, ByRef errmsg As String) As Integer
      Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE
      Dim version As String = ""
      Dim apidate As String = ""
      Dim i As Integer
      Dim dll_ver as yu16

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
    '''   Frees dynamically allocated memory blocks used by the Yoctopuce library.
    ''' <para>
    '''   It is generally not required to call this function, unless you
    '''   want to free all dynamically allocated memory blocks in order to
    '''   track a memory leak for instance.
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
    '''   Setup the Yoctopuce library to use modules connected on a given machine.
    ''' <para>
    '''   The
    '''   parameter will determine how the API will work. Use the following values:
    ''' </para>
    ''' <para>
    '''   <b>usb</b>: When the <c>usb</c> keyword is used, the API will work with
    '''   devices connected directly to the USB bus. Some programming languages such a Javascript,
    '''   PHP, and Java don't provide direct access to USB hardware, so <c>usb</c> will
    '''   not work with these. In this case, use a VirtualHub or a networked YoctoHub (see below).
    ''' </para>
    ''' <para>
    '''   <b><i>x.x.x.x</i></b> or <b><i>hostname</i></b>: The API will use the devices connected to the
    '''   host with the given IP address or hostname. That host can be a regular computer
    '''   running a VirtualHub, or a networked YoctoHub such as YoctoHub-Ethernet or
    '''   YoctoHub-Wireless. If you want to use the VirtualHub running on you local
    '''   computer, use the IP address 127.0.0.1.
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
    '''   for this limitation is to setup the library to use the VirtualHub
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
    '''   You can call <i>RegisterHub</i> several times to connect to several machines.
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
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
    '''   Fault-tolerant alternative to <c>RegisterHub()</c>.
    ''' <para>
    '''   This function has the same
    '''   purpose and same arguments as <c>RegisterHub()</c>, but does not trigger
    '''   an error when the selected hub is not available at the time of the function call.
    '''   This makes it possible to register a network hub independently of the current
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
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
    '''   Setup the Yoctopuce library to no more use modules connected on a previously
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
    '''   hub is usable. The url parameter follow the same convention as the <c>RegisterHub</c>
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
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
    '''   and to make the application aware of hot-plug events.
    ''' </para>
    ''' </summary>
    ''' <param name="errmsg">
    '''   a string passed by reference to receive any error message.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
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
        Return res
      End If

      res = _yapiHandleEvents(errbuffer)
      If (YISERR(res)) Then
        errmsg = errbuffer.ToString()
        Return res
      End If
      While (_PlugEvents.Count > 0)
        yapiLockDeviceCallBack(errmsg)
        p = _PlugEvents(0)
        _PlugEvents.RemoveAt(0)
        yapiUnlockDeviceCallBack(errmsg)
        p.invoke()
      End While
      Return YAPI_SUCCESS
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
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
    '''   <c>A..Z</c>, <c>a..z</c>, <c>0..9</c>, <c>_</c>, and <c>-</c>.
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
    '''   it  either fires the debugger or aborts (i.e. crash) the program.
    '''   On failure, throws an exception or returns a negative error code.
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
  Public Delegate Sub _yapiFunctionUpdateFunc(ByVal dev As YFUN_DESCR, ByVal value As IntPtr)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)>
  Public Delegate Sub _yapiTimedReportFunc(ByVal dev As YFUN_DESCR, ByVal timestamp As Double, ByVal data As IntPtr, ByVal len As System.UInt32)

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
  '''   The YFirmwareUpdate class let you control the firmware update of a Yoctopuce
  '''   module.
  ''' <para>
  '''   This class should not be instantiate directly, instead the method
  '''   <c>updateFirmware</c> should be called to get an instance of YFirmwareUpdate.
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
        If (((Me._settings).Length <> 0)) Then
          Me._progress_msg = "restoring settings"
          m = YModule.FindModule(Me._serial + ".module")
          If (Not (m.isOnline())) Then
            Return Me._progress
          End If
          If (Me._progress < 95) Then
            prod_prefix = (m.get_productName()).Substring( 0, 8)
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
              Me._progress =  100
              Me._progress_msg = "success"
            End If
          End If
        Else
          Me._progress =  100
          Me._progress_msg = "success"
        End If
      End If
      Return Me._progress
    End Function

    '''*
    ''' <summary>
    '''   Returns a list of all the modules in "firmware update" mode.
    ''' <para>
    '''   Only devices
    '''   connected over USB are listed. For devices connected to a YoctoHub, you
    '''   must connect yourself to the YoctoHub web interface.
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
      If (( leng >= 6) AndAlso ("error:" = (err).Substring(0, 6))) Then
        Me._progress = -1
        Me._progress_msg = (err).Substring( 6, leng - 6)
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
  ''' <summary>
  '''   YDataStream objects represent bare recorded measure sequences,
  '''   exactly as found within the data logger present on Yoctopuce
  '''   sensors.
  ''' <para>
  ''' </para>
  ''' <para>
  '''   In most cases, it is not necessary to use YDataStream objects
  '''   directly, as the YDataSet objects (returned by the
  '''   <c>get_recordedData()</c> method from sensors and the
  '''   <c>get_dataSets()</c> method from the data logger) provide
  '''   a more convenient interface.
  ''' </para>
  ''' </summary>
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
    Protected _duration As Integer
    Protected _columnNames As List(Of String)
    Protected _functionId As String
    Protected _isClosed As Boolean
    Protected _isAvg As Boolean
    Protected _isScal As Boolean
    Protected _isScal32 As Boolean
    Protected _decimals As Integer
    Protected _offset As Double
    Protected _scale As Double
    Protected _samplesPerHour As Integer
    Protected _minVal As Double
    Protected _avgVal As Double
    Protected _maxVal As Double
    Protected _decexp As Double
    Protected _caltyp As Integer
    Protected _calpar As List(Of Integer)
    Protected _calraw As List(Of Double)
    Protected _calref As List(Of Double)
    Protected _values As List(Of List(Of Double))
    REM --- (end of generated code: YDataStream attributes declaration)
    Protected _calhdl As yCalibrationHandler

    Public Sub New(parent As YFunction)
      _parent = parent
      REM --- (generated code: YDataStream attributes initialization)
      _runNo = 0
      _utcStamp = 0
      _nCols = 0
      _nRows = 0
      _duration = 0
      _columnNames = New List(Of String)()
      _decimals = 0
      _offset = 0
      _scale = 0
      _samplesPerHour = 0
      _minVal = 0
      _avgVal = 0
      _maxVal = 0
      _decexp = 0
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
      _duration = 0
      _columnNames = New List(Of String)()
      _decimals = 0
      _offset = 0
      _scale = 0
      _samplesPerHour = 0
      _minVal = 0
      _avgVal = 0
      _maxVal = 0
      _decexp = 0
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
      Dim iRaw As Integer = 0
      Dim iRef As Integer = 0
      Dim fRaw As Double = 0
      Dim fRef As Double = 0
      Dim duration_float As Double = 0
      Dim iCalib As List(Of Integer) = New List(Of Integer)()
      REM // decode sequence header to extract data
      Me._runNo = encoded(0) + (((encoded(1)) << (16)))
      Me._utcStamp = encoded(2) + (((encoded(3)) << (16)))
      val = encoded(4)
      Me._isAvg = (((val) And (&H100)) = 0)
      Me._samplesPerHour = ((val) And (&Hff))
      If (((val) And (&H100)) <> 0) Then
        Me._samplesPerHour = Me._samplesPerHour * 3600
      Else
        If (((val) And (&H200)) <> 0) Then
          Me._samplesPerHour = Me._samplesPerHour * 60
        End If
      End If
      val = encoded(5)
      If (val > 32767) Then
        val = val - 65536
      End If
      Me._decimals = val
      Me._offset = val
      Me._scale = encoded(6)
      Me._isScal = (Me._scale <> 0)
      Me._isScal32 = (encoded.Count >= 14)
      val = encoded(7)
      Me._isClosed = (val <> &Hffff)
      If (val = &Hffff) Then
        val = 0
      End If
      Me._nRows = val
      duration_float = Me._nRows * 3600 / Me._samplesPerHour
      Me._duration = CType(Math.Round(duration_float), Integer)
      REM // precompute decoding parameters
      Me._decexp = 1.0
      If (Me._scale = 0) Then
        i = 0
        While (i < Me._decimals)
          Me._decexp = Me._decexp * 10.0
          i = i + 1
        End While
      End If
      iCalib = dataset._get_calibration()
      Me._caltyp = iCalib(0)
      If (Me._caltyp <> 0) Then
        Me._calhdl = YAPI._getCalibrationHandler(Me._caltyp)
        maxpos = iCalib.Count
        Me._calpar.Clear()
        Me._calraw.Clear()
        Me._calref.Clear()
        If (Me._isScal32) Then
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
        Else
          i = 1
          While (i + 1 < maxpos)
            iRaw = iCalib(i)
            iRef = iCalib(i + 1)
            Me._calpar.Add(iRaw)
            Me._calpar.Add(iRef)
            If (Me._isScal) Then
              fRaw = iRaw
              fRaw = (fRaw - Me._offset) / Me._scale
              fRef = iRef
              fRef = (fRef - Me._offset) / Me._scale
              Me._calraw.Add(fRaw)
              Me._calref.Add(fRef)
            Else
              Me._calraw.Add(YAPI._decimalToDouble(iRaw))
              Me._calref.Add(YAPI._decimalToDouble(iRef))
            End If
            i = i + 2
          End While
        End If
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
        If (Me._isScal32) Then
          Me._avgVal = Me._decodeAvg(encoded(8) + (((((encoded(9)) Xor (&H8000))) << (16))), 1)
          Me._minVal = Me._decodeVal(encoded(10) + (((encoded(11)) << (16))))
          Me._maxVal = Me._decodeVal(encoded(12) + (((encoded(13)) << (16))))
        Else
          Me._minVal = Me._decodeVal(encoded(8))
          Me._maxVal = Me._decodeVal(encoded(9))
          Me._avgVal = Me._decodeAvg(encoded(10) + (((encoded(11)) << (16))), Me._nRows)
        End If
      End If
      Return 0
    End Function

    Public Overridable Function _parseStream(sdata As Byte()) As Integer
      Dim idx As Integer = 0
      Dim udat As List(Of Integer) = New List(Of Integer)()
      Dim dat As List(Of Double) = New List(Of Double)()
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
          If (Me._isScal32) Then
            dat.Add(Me._decodeVal(udat(idx + 2) + (((udat(idx + 3)) << (16)))))
            dat.Add(Me._decodeAvg(udat(idx) + (((((udat(idx + 1)) Xor (&H8000))) << (16))), 1))
            dat.Add(Me._decodeVal(udat(idx + 4) + (((udat(idx + 5)) << (16)))))
            idx = idx + 6
          Else
            dat.Add(Me._decodeVal(udat(idx)))
            dat.Add(Me._decodeAvg(udat(idx + 2) + (((udat(idx + 3)) << (16))), 1))
            dat.Add(Me._decodeVal(udat(idx + 1)))
            idx = idx + 4
          End If
          Me._values.Add(New List(Of Double)(dat))
        End While
      Else
        If (Me._isScal AndAlso Not (Me._isScal32)) Then
          While (idx < udat.Count)
            dat.Clear()
            dat.Add(Me._decodeVal(udat(idx)))
            Me._values.Add(New List(Of Double)(dat))
            idx = idx + 1
          End While
        Else
          While (idx + 1 < udat.Count)
            dat.Clear()
            dat.Add(Me._decodeAvg(udat(idx) + (((((udat(idx + 1)) Xor (&H8000))) << (16))), 1))
            Me._values.Add(New List(Of Double)(dat))
            idx = idx + 2
          End While
        End If
      End If

      Me._nRows = Me._values.Count
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function _get_url() As String
      Dim url As String
      url = "logger.json?id=" +
      Me._functionId + "&run=" + Convert.ToString(Me._runNo) + "&utc=" + Convert.ToString(Me._utcStamp)
      Return url
    End Function

    Public Overridable Function loadStream() As Integer
      Return Me._parseStream(Me._parent._download(Me._get_url()))
    End Function

    Public Overridable Function _decodeVal(w As Integer) As Double
      Dim val As Double = 0
      val = w
      If (Me._isScal32) Then
        val = val / 1000.0
      Else
        If (Me._isScal) Then
          val = (val - Me._offset) / Me._scale
        Else
          val = YAPI._decimalToDouble(w)
        End If
      End If
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
      If (Me._isScal32) Then
        val = val / 1000.0
      Else
        If (Me._isScal) Then
          val = (val / (100 * count) - Me._offset) / Me._scale
        Else
          val = val / (count * Me._decexp)
        End If
      End If
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
    '''   If you need an absolute UTC timestamp, use <c>get_startTimeUTC()</c>.
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
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to the number of seconds
    '''   between the Jan 1, 1970 and the beginning of this data
    '''   stream (i.e. Unix time representation of the absolute time).
    ''' </returns>
    '''/
    Public Overridable Function get_startTimeUTC() As Long
      Return Me._utcStamp
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
      Return (3600000 \ Me._samplesPerHour)
    End Function

    Public Overridable Function get_dataSamplesInterval() As Double
      Return 3600.0 / Me._samplesPerHour
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
    '''   this method will always return Y_DATA_INVALID.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the smallest value,
    '''   or Y_DATA_INVALID if the stream is not yet complete (still recording).
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_DATA_INVALID.
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
    '''   this method will always return Y_DATA_INVALID.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the average value,
    '''   or Y_DATA_INVALID if the stream is not yet complete (still recording).
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_DATA_INVALID.
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
    '''   this method will always return Y_DATA_INVALID.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating-point number corresponding to the largest value,
    '''   or Y_DATA_INVALID if the stream is not yet complete (still recording).
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_DATA_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_maxValue() As Double
      Return Me._maxVal
    End Function

    '''*
    ''' <summary>
    '''   Returns the approximate duration of this stream, in seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the number of seconds covered by this stream.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns Y_DURATION_INVALID.
    ''' </para>
    '''/
    Public Overridable Function get_duration() As Integer
      If (Me._isClosed) Then
        Return Me._duration
      End If
      Return CInt(CInt((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds) - Me._utcStamp)
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
    '''   On failure, throws an exception or returns Y_DATA_INVALID.
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
  ''' <summary>
  '''   YMeasure objects are used within the API to represent
  '''   a value measured at a specified time.
  ''' <para>
  '''   These objects are
  '''   used in particular in conjunction with the YDataSet class.
  ''' </para>
  ''' </summary>
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
    '''   an floating point number corresponding to the number of seconds
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
    '''   an floating point number corresponding to the number of seconds
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
  ''' <summary>
  '''   YDataSet objects make it possible to retrieve a set of recorded measures
  '''   for a given sensor and a specified time interval.
  ''' <para>
  '''   They can be used
  '''   to load data points with a progress report. When the YDataSet object is
  '''   instantiated by the <c>get_recordedData()</c>  function, no data is
  '''   yet loaded from the module. It is only when the <c>loadMore()</c>
  '''   method is called over and over than data will be effectively loaded
  '''   from the dataLogger.
  ''' </para>
  ''' <para>
  '''   A preview of available measures is available using the function
  '''   <c>get_preview()</c> as soon as <c>loadMore()</c> has been called
  '''   once. Measures themselves are available using function <c>get_measures()</c>
  '''   when loaded by subsequent calls to <c>loadMore()</c>.
  ''' </para>
  ''' <para>
  '''   This class can only be used on devices that use a recent firmware,
  '''   as YDataSet objects are not supported by firmwares older than version 13000.
  ''' </para>
  ''' </summary>
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
    Protected _startTime As Long
    Protected _endTime As Long
    Protected _progress As Integer
    Protected _calib As List(Of Integer)
    Protected _streams As List(Of YDataStream)
    Protected _summary As YMeasure
    Protected _preview As List(Of YMeasure)
    Protected _measures As List(Of YMeasure)
    REM --- (end of generated code: YDataSet attributes declaration)



    Sub New(parent As YFunction, functionId As String, unit As String, startTime As Long, endTime As Long)
      REM --- (generated code: YDataSet attributes initialization)
      _startTime = 0
      _endTime = 0
      _progress = 0
      _calib = New List(Of Integer)()
      _streams = New List(Of YDataStream)()
      _preview = New List(Of YMeasure)()
      _measures = New List(Of YMeasure)()
      REM --- (end of generated code: YDataSet attributes initialization)
      _parent = parent
      _functionId = functionId
      _unit = unit
      _startTime = startTime
      _endTime = endTime
      _summary = New YMeasure(0, 0, 0, 0, 0)
      _progress = -1
    End Sub


    Sub New(parent As YFunction)
      REM --- (generated code: YDataSet attributes initialization)
      _startTime = 0
      _endTime = 0
      _progress = 0
      _calib = New List(Of Integer)()
      _streams = New List(Of YDataStream)()
      _preview = New List(Of YMeasure)()
      _measures = New List(Of YMeasure)()
      REM --- (end of generated code: YDataSet attributes initialization)
      _parent = parent
      _startTime = 0
      _endTime = 0
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
      Dim streamStartTime As Long
      Dim streamEndTime As Long
      Dim startTime As Long = &H7fffffff
      Dim endTime As Long = 0
      Dim summaryMinVal As Double = [Double].MaxValue
      Dim summaryMaxVal As Double = -[Double].MaxValue
      Dim summaryTotalTime As Double = 0
      Dim summaryTotalAvg As Double = 0

      Me._functionId = p.getString("id")
      Me._unit = p.getString("unit")
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
        streamStartTime = stream.get_startTimeUTC() - CLng(stream.get_dataSamplesIntervalMs() / 1000)
        streamEndTime = stream.get_startTimeUTC() + stream.get_duration()
        If _startTime > 0 AndAlso streamEndTime <= _startTime Then
          REM this stream is too early, drop it
        ElseIf _endTime > 0 AndAlso stream.get_startTimeUTC() > Me._endTime Then
          REM this stream is too late, drop it
        Else
          _streams.Add(stream)
          If startTime > streamStartTime Then
            startTime = streamStartTime
          End If
          If endTime < streamEndTime Then
            endTime = streamEndTime
          End If
          If stream.isClosed() AndAlso stream.get_startTimeUTC() >= Me._startTime AndAlso (Me._endTime = 0 OrElse streamEndTime <= Me._endTime) Then
            If summaryMinVal > stream.get_minValue() Then
              summaryMinVal = stream.get_minValue()
            End If
            If summaryMaxVal < stream.get_maxValue() Then
              summaryMaxVal = stream.get_maxValue()
            End If
            summaryTotalAvg += stream.get_averageValue() * stream.get_duration()
            summaryTotalTime += stream.get_duration()
            Dim rec As New YMeasure(stream.get_startTimeUTC(),
                                    streamEndTime,
                                    stream.get_minValue(),
                                    stream.get_averageValue(),
                                    stream.get_maxValue())
            _preview.Add(rec)
          End If
        End If
      Next i
      If (_streams.Count > 0) And (summaryTotalTime > 0) Then
        REM update time boundaries with actual data
        If (_startTime < startTime) Then
          _startTime = startTime
        End If
        If (_endTime = 0 Or _endTime > endTime) Then
          _endTime = endTime
        End If
        _summary = New YMeasure(_startTime,
                                _endTime,
                                summaryMinVal,
                                summaryTotalAvg / summaryTotalTime,
                                summaryMaxVal)
      End If
      _progress = 0
      Return get_progress()
    End Function

    REM --- (generated code: YDataSet private methods declaration)

    REM --- (end of generated code: YDataSet private methods declaration)

    REM --- (generated code: YDataSet public methods declaration)
    Public Overridable Function _get_calibration() As List(Of Integer)
      Return Me._calib
    End Function

    Public Overridable Function processMore(progress As Integer, data As Byte()) As Integer
      Dim i_i As Integer
      Dim stream As YDataStream
      Dim dataRows As List(Of List(Of Double)) = New List(Of List(Of Double))()
      Dim strdata As String
      Dim tim As Double = 0
      Dim itv As Double = 0
      Dim nCols As Integer = 0
      Dim minCol As Integer = 0
      Dim avgCol As Integer = 0
      Dim maxCol As Integer = 0

      If (progress <> Me._progress) Then
        Return Me._progress
      End If
      If (Me._progress < 0) Then
        strdata = YAPI.DefaultEncoding.GetString(data)
        If (strdata = "{}") Then
          Me._parent._throw(YAPI.VERSION_MISMATCH, "device firmware is too old")
          Return YAPI.VERSION_MISMATCH
        End If
        Return Me._parse(strdata)
      End If
      stream = Me._streams(Me._progress)
      stream._parseStream(data)
      dataRows = stream.get_dataRows()
      Me._progress = Me._progress + 1
      If (dataRows.Count = 0) Then
        Return Me.get_progress()
      End If
      tim = CType(stream.get_startTimeUTC(), Double)
      itv = stream.get_dataSamplesInterval()
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

      For i_i = 0 To dataRows.Count - 1
        If ((tim >= Me._startTime) AndAlso ((Me._endTime = 0) OrElse (tim <= Me._endTime))) Then
          Me._measures.Add(New YMeasure(tim - itv, tim, dataRows(i_i)(minCol), dataRows(i_i)(avgCol), dataRows(i_i)(maxCol)))
        End If
        tim = tim + itv
        tim = Math.Round(tim * 1000) / 1000.0
      Next i_i

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
    '''   On failure, throws an exception or returns  <c>Y_HARDWAREID_INVALID</c>.
    ''' </para>
    '''/
    Public Overridable Function get_hardwareId() As String
      Dim mo As YModule
      If (Not (Me._hardwareId = "")) Then
        Return Me._hardwareId
      End If
      mo = Me._parent.get_module()
      Me._hardwareId = "" +  mo.get_serialNumber() + "." + Me.get_functionId()
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
    '''   On failure, throws an exception or returns  <c>Y_UNIT_INVALID</c>.
    ''' </para>
    '''/
    Public Overridable Function get_unit() As String
      Return Me._unit
    End Function

    '''*
    ''' <summary>
    '''   Returns the start time of the dataset, relative to the Jan 1, 1970.
    ''' <para>
    '''   When the YDataSet is created, the start time is the value passed
    '''   in parameter to the <c>get_dataSet()</c> function. After the
    '''   very first call to <c>loadMore()</c>, the start time is updated
    '''   to reflect the timestamp of the first measure actually found in the
    '''   dataLogger within the specified range.
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
      Return Me._startTime
    End Function

    '''*
    ''' <summary>
    '''   Returns the end time of the dataset, relative to the Jan 1, 1970.
    ''' <para>
    '''   When the YDataSet is created, the end time is the value passed
    '''   in parameter to the <c>get_dataSet()</c> function. After the
    '''   very first call to <c>loadMore()</c>, the end time is updated
    '''   to reflect the timestamp of the last measure actually found in the
    '''   dataLogger within the specified range.
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
      Return Me._endTime
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
    '''   Loads the the next block of measures from the dataLogger, and updates
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
        If (Me._startTime <> 0) Then
          url = "" + url + "&from=" + Convert.ToString(Me._startTime)
        End If
        If (Me._endTime <> 0) Then
          url = "" + url + "&to=" + Convert.ToString(Me._endTime)
        End If
      Else
        If (Me._progress >= Me._streams.Count) Then
          Return 100
        Else
          stream = Me._streams(Me._progress)
          url = stream._get_url()
        End If
      End If
      Return Me.processMore(Me._progress, Me._parent._download(url))
    End Function

    '''*
    ''' <summary>
    '''   Returns an YMeasure object which summarizes the whole
    '''   DataSet.
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
    '''   an YMeasure object
    ''' </returns>
    '''/
    Public Overridable Function get_summary() As YMeasure
      Return Me._summary
    End Function

    '''*
    ''' <summary>
    '''   Returns a condensed version of the measures that can
    '''   retrieved in this YDataSet, as a list of YMeasure
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
    '''   The result is provided as a list of YMeasure objects.
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
      Dim i_i As Integer
      Dim startUtc As Long = 0
      Dim stream As YDataStream
      Dim dataRows As List(Of List(Of Double)) = New List(Of List(Of Double))()
      Dim measures As List(Of YMeasure) = New List(Of YMeasure)()
      Dim tim As Double = 0
      Dim itv As Double = 0
      Dim nCols As Integer = 0
      Dim minCol As Integer = 0
      Dim avgCol As Integer = 0
      Dim maxCol As Integer = 0

      startUtc = CType(Math.Round(measure.get_startTimeUTC()), Long)
      stream = Nothing
      For i_i = 0 To Me._streams.Count - 1
        If (Me._streams(i_i).get_startTimeUTC() = startUtc) Then
          stream = Me._streams(i_i)
        End If
      Next i_i
      If ((stream Is Nothing)) Then
        Return measures
      End If
      dataRows = stream.get_dataRows()
      If (dataRows.Count = 0) Then
        Return measures
      End If
      tim = CType(stream.get_startTimeUTC(), Double)
      itv = stream.get_dataSamplesInterval()
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

      For i_i = 0 To dataRows.Count - 1
        If ((tim >= Me._startTime) AndAlso ((Me._endTime = 0) OrElse (tim <= Me._endTime))) Then
          measures.Add(New YMeasure(tim - itv, tim, dataRows(i_i)(minCol), dataRows(i_i)(avgCol), dataRows(i_i)(maxCol)))
        End If
        tim = tim + itv
      Next i_i

      Return measures
    End Function

    '''*
    ''' <summary>
    '''   Returns all measured values currently available for this DataSet,
    '''   as a list of YMeasure objects.
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
          request = "GET /api.json?fw=" + _cacheJson.getYJSONObject("module").getString("firmwareRelease") + " " & vbCr & vbLf & vbCr & vbLf
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
    Public Shared _TimedReportCallbackList As List(Of YFunction) = New List(Of YFunction)

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
      Dim i, c_ord As Integer
      Dim uchangeval, h As String
      uchangeval = ""
      Dim c As Char
      For i = 0 To changeval.Length - 1
        c = changeval.Chars(i)
        If (c < " ") Or ((c > Chr(122)) And (c <> "~")) Or (c = Chr(34)) Or (c = "%") Or (c = "&") Or
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
      Dim outbuf, res As Byte()
      Dim found, body As Integer

      request = "GET /" + url + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      outbuf = _request(request)
      If (outbuf.Length = 0) Then
        Return outbuf
      End If
      found = 0
      Do While found < outbuf.Length - 4 And
          (outbuf(found) <> 13 Or outbuf(found + 1) <> 10 Or
           outbuf(found + 2) <> 13 Or outbuf(found + 3) <> 10)
        found += 1
      Loop
      If found >= outbuf.Length - 4 Then
        _throw(YAPI.IO_ERROR, "http request failed")
        ReDim outbuf(-1)
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
      If outbuf.Length = 0 Then
        _throw(YAPI.IO_ERROR, "http request failed")
        Return YAPI.IO_ERROR
      End If
      Return YAPI.SUCCESS
    End Function

    Public Function _findDataStream(dataset As YDataSet, def As String) As YDataStream
      Dim key As String = dataset.get_functionId() + ":" + def
      Dim newDataStream As YDataStream
      If (_dataStreams.ContainsKey(key)) Then
        Return CType(_dataStreams(key), YDataStream)
      End If
      newDataStream = New YDataStream(Me, dataset, YAPI._decodeWords(def))
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

    Protected Shared Sub _UpdateTimedReportCallbackList(func As YFunction, add As Boolean)
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
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   a string that uniquely characterizes the function
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
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
      Dim attrVal As Byte()
      url = "api/" +  Me.get_functionId() + "/" + attrName
      attrVal = Me._download(url)
      Return YAPI.DefaultEncoding.GetString(attrVal)
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
    '''   On failure, throws an exception or returns  <c>Y_HARDWAREID_INVALID</c>.
    ''' </para>
    '''/

    Public Function get_hardwareId() As String
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
    '''   On failure, throws an exception or returns  <c>Y_FUNCTIONID_INVALID</c>.
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

    Protected Function _json_get_array(data As Byte()) As List(Of String)
      Dim array As New YJSONArray(YAPI.DefaultEncoding.GetString(data))
      array.parse()
      Dim list As New List(Of String)()
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


    Private Function get_json_path_struct(jsonObject As YJSONObject, paths As String(), ofs As Integer) As String

      Dim key As String = paths(ofs)
      If Not jsonObject.has(key) Then
        Return ""
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
      Return ""
    End Function

    Private Function get_json_path_array(jsonArray As YJSONArray, paths As String(), ofs As Integer) As String
      Dim key As Integer = Convert.ToInt32(paths(ofs))
      If jsonArray.Length <= key Then
        Return ""
      End If

      Dim obj As YJSONContent = jsonArray.[get](key)
      If obj IsNot Nothing Then
        If paths.Length = ofs + 1 Then
          Return obj.ToString()
        End If

        If TypeOf obj Is YJSONArray Then
          Return get_json_path_array(jsonArray.getYJSONArray(key), paths, ofs + 1)
        ElseIf TypeOf obj Is YJSONObject Then
          Return get_json_path_struct(jsonArray.getYJSONObject(key), paths, ofs + 1)
        End If
      End If
      Return ""
    End Function


    Public Function _get_json_path(ByVal json As String, ByVal path As String) As String

      Dim jsonObject As YJSONObject = Nothing
      jsonObject = New YJSONObject(json)
      jsonObject.parse()
      Dim split As String() = path.Split(New Char() {"\"C, "|"C})
      Return get_json_path_struct(jsonObject, split, 0)
    End Function

    Public Function _decode_json_string(ByVal json As String) As String
      Dim len As Integer
      Dim buffer As StringBuilder
      Dim decoded_len As Integer
      len = json.Length
      buffer = New StringBuilder(len)
      decoded_len = _yapiJsonDecodeString(New StringBuilder(json), buffer)
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function load(ByVal msValidity As Integer) As YRETCODE
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
    '''   If the function has never been contacted, the returned value is <c>Y_FUNCTIONDESCRIPTOR_INVALID</c>.
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
  '''   a string that uniquely characterizes the function
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
  '''   This interface is identical for all Yoctopuce USB modules.
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
    REM --- (end of generated code: YModule attributes declaration)

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
    ''' </para>
    ''' </summary>
    ''' <param name="functionIndex">
    '''   the index of the function for which the information is desired, starting at 0 for the first function.
    ''' </param>
    ''' <returns>
    '''   a string corresponding to the type of the function
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
      i = 1
      While (i < funcId.Length)
        If (Not Char.IsLetter(funcId(i))) Then
          Exit While
        End If
        i += 1
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
    '''   the callback function to call, or a Nothing pointer. The callback function should take two
    '''   arguments: the module object that emitted the log message, and the character string containing the log.
    ''' @noreturn
    ''' </param>
    '''/
    Public Function registerLogCallback(ByVal callback As YModuleLogCallback) As Integer
      _logCallback = callback
      If _logCallback Is Nothing Then
        _yapiStartStopDeviceLogCallback(New StringBuilder(_serialNumber), 0)
      Else
        _yapiStartStopDeviceLogCallback(New StringBuilder(_serialNumber), 1)
      End If
      Return YAPI_SUCCESS
    End Function

    Public Function get_logCallback() As YModuleLogCallback
      Return _logCallback
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
    '''   On failure, throws an exception or returns <c>Y_PRODUCTNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_productName() As String
      Dim res As String
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_SERIALNUMBER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_serialNumber() As String
      Dim res As String
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_PRODUCTID_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_productId() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PRODUCTID_INVALID
        End If
      End If
      res = Me._productId
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the hardware release version of the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the hardware release version of the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PRODUCTRELEASE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_productRelease() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_FIRMWARERELEASE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_firmwareRelease() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   a value among <c>Y_PERSISTENTSETTINGS_LOADED</c>, <c>Y_PERSISTENTSETTINGS_SAVED</c> and
    '''   <c>Y_PERSISTENTSETTINGS_MODIFIED</c> corresponding to the current state of persistent module settings
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PERSISTENTSETTINGS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_persistentSettings() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   Returns the luminosity of the  module informative leds (from 0 to 100).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the luminosity of the  module informative leds (from 0 to 100)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LUMINOSITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_luminosity() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   either <c>Y_BEACON_OFF</c> or <c>Y_BEACON_ON</c>, according to the state of the localization beacon
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BEACON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_beacon() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   either <c>Y_BEACON_OFF</c> or <c>Y_BEACON_ON</c>
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   On failure, throws an exception or returns <c>Y_UPTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_upTime() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_USBCURRENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_usbCurrent() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_REBOOTCOUNTDOWN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rebootCountdown() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_USERVAR_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_userVar() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
      obj = CType(YFunction._FindFromCache("Module", func), YModule)
      If ((obj Is Nothing)) Then
        obj = New YModule(func)
        YFunction._AddToCache("Module", func, obj)
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

    '''*
    ''' <summary>
    '''   Saves current settings in the nonvolatile memory of the module.
    ''' <para>
    '''   Warning: the number of allowed save operations during a module life is
    '''   limited (about 100000 cycles). Do not call this function within a loop.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function triggerFirmwareUpdate(secBeforeReboot As Integer) As Integer
      Return Me.set_rebootCountdown(-secBeforeReboot)
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
    '''   Triggers a configuration change callback, to check if they are supported or not.
    ''' <para>
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function triggerConfigChangeCallback() As Integer
      Me._setAttr("persistentSettings","2")
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
      tmp_res = YFirmwareUpdate.CheckFirmware(serial,path, release)
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
      Dim settings As Byte()

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
      Dim i_i As Integer
      Dim settings As Byte()
      Dim json As Byte()
      Dim res As Byte()
      Dim sep As String
      Dim name As String
      Dim item As String
      Dim t_type As String
      Dim id As String
      Dim url As String
      Dim file_data As String
      Dim file_data_bin As Byte()
      Dim temp_data_bin As Byte()
      Dim ext_settings As String
      Dim filelist As List(Of String) = New List(Of String)()
      Dim templist As List(Of String) = New List(Of String)()

      settings = Me._download("api.json")
      If ((settings).Length = 0) Then
        Return settings
      End If
      ext_settings = ", ""extras"":["
      templist = Me.get_functionIds("Temperature")
      sep = ""
      For i_i = 0 To  templist.Count - 1
        If (YAPI._atoi(Me.get_firmwareRelease()) > 9000) Then
          url = "api/" +  templist(i_i) + "/sensorType"
          t_type = YAPI.DefaultEncoding.GetString(Me._download(url))
          If (t_type = "RES_NTC") Then
            id = ( templist(i_i)).Substring( 11, ( templist(i_i)).Length - 11)
            temp_data_bin = Me._download("extra.json?page=" + id)
            If ((temp_data_bin).Length = 0) Then
              Return temp_data_bin
            End If
            item = "" +  sep + "{""fid"":""" +   templist(i_i) + """, ""json"":" + YAPI.DefaultEncoding.GetString(temp_data_bin) + "}" + vbLf + ""
            ext_settings = ext_settings + item
            sep = ","
          End If
        End If
      Next i_i
      ext_settings = ext_settings + "]," + vbLf + """files"":["
      If (Me.hasFunction("files")) Then
        json = Me._download("files.json?a=dir&f=")
        If ((json).Length = 0) Then
          Return json
        End If
        filelist = Me._json_get_array(json)
        sep = ""
        For i_i = 0 To  filelist.Count - 1
          name = Me._json_get_key(YAPI.DefaultEncoding.GetBytes( filelist(i_i)), "name")
          If (((name).Length > 0) AndAlso Not (name = "startupConf.json")) Then
            file_data_bin = Me._download(Me._escapeAttr(name))
            file_data = YAPI._bytesToHexStr(file_data_bin, 0, file_data_bin.Length)
            item = "" +  sep + "{""name"":""" +  name + """, ""data"":""" + file_data + """}" + vbLf + ""
            ext_settings = ext_settings + item
            sep = ","
          End If
        Next i_i
      End If
      res = YAPI.DefaultEncoding.GetBytes("{ ""api"":" + YAPI.DefaultEncoding.GetString(settings) + ext_settings + "]}")
      Return res
    End Function

    Public Overridable Function loadThermistorExtra(funcId As String, jsonExtra As String) As Integer
      Dim values As List(Of String) = New List(Of String)()
      Dim url As String
      Dim curr As String
      Dim currTemp As String
      Dim ofs As Integer = 0
      Dim size As Integer = 0
      url = "api/" + funcId + ".json?command=Z"

      Me._download(url)
      REM // add records in growing resistance value
      values = Me._json_get_array(YAPI.DefaultEncoding.GetBytes(jsonExtra))
      ofs = 0
      size = values.Count
      While (ofs + 1 < size)
        curr = values(ofs)
        currTemp = values(ofs + 1)
        url = "api/" +   funcId + "/.json?command=m" +  curr + ":" + currTemp
        Me._download(url)
        ofs = ofs + 2
      End While
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function set_extraSettings(jsonExtra As String) As Integer
      Dim i_i As Integer
      Dim extras As List(Of String) = New List(Of String)()
      Dim functionId As String
      Dim data As String
      extras = Me._json_get_array(YAPI.DefaultEncoding.GetBytes(jsonExtra))
      For i_i = 0 To  extras.Count - 1
        functionId = Me._get_json_path( extras(i_i), "fid")
        functionId = Me._decode_json_string(functionId)
        data = Me._get_json_path( extras(i_i), "json")
        If (Me.hasFunction(functionId)) Then
          Me.loadThermistorExtra(functionId, data)
        End If
      Next i_i
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_allSettingsAndFiles(settings As Byte()) As Integer
      Dim i_i As Integer
      Dim down As Byte()
      Dim json As String
      Dim json_api As String
      Dim json_files As String
      Dim json_extra As String
      json = YAPI.DefaultEncoding.GetString(settings)
      json_api = Me._get_json_path(json, "api")
      If (json_api = "") Then
        Return Me.set_allSettings(settings)
      End If
      json_extra = Me._get_json_path(json, "extras")
      If (Not (json_extra = "")) Then
        Me.set_extraSettings(json_extra)
      End If
      Me.set_allSettings(YAPI.DefaultEncoding.GetBytes(json_api))
      If (Me.hasFunction("files")) Then
        Dim files As List(Of String) = New List(Of String)()
        Dim res As String
        Dim name As String
        Dim data As String
        down = Me._download("files.json?a=format")
        res = Me._get_json_path(YAPI.DefaultEncoding.GetString(down), "res")
        res = Me._decode_json_string(res)
        If Not(res = "ok") Then
          me._throw( YAPI.IO_ERROR,  "format failed")
          return YAPI.IO_ERROR
        end if
        json_files = Me._get_json_path(json, "files")
        files = Me._json_get_array(YAPI.DefaultEncoding.GetBytes(json_files))
        For i_i = 0 To  files.Count - 1
          name = Me._get_json_path( files(i_i), "name")
          name = Me._decode_json_string(name)
          data = Me._get_json_path( files(i_i), "data")
          data = Me._decode_json_string(data)
          Me._upload(name, YAPI._hexStrToBin(data))
        Next i_i
      End If
      Return YAPI.SUCCESS
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

      count  = Me.functionCount()
      i = 0
      While (i < count)
        fid  = Me.functionId(i)
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
      Dim i_i As Integer
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
            For i_i = 0 To words_str.Count - 1
              words.Add(YAPI._atoi(words_str(i_i)))
            Next i_i
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
              ratio = Double.Parse(param)
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
            calibData( i) = (calibData(i) - paramOffset) / paramScale
          Else
            REM // floating-point decoding
            calibData( i) = YAPI._decimalToDouble(CType(Math.Round(calibData(i)), Integer))
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
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_allSettings(settings As Byte()) As Integer
      Dim i_i As Integer
      Dim restoreLast As List(Of String) = New List(Of String)()
      Dim old_json_flat As Byte()
      Dim old_dslist As List(Of String) = New List(Of String)()
      Dim old_jpath As List(Of String) = New List(Of String)()
      Dim old_jpath_len As List(Of Integer) = New List(Of Integer)()
      Dim old_val_arr As List(Of String) = New List(Of String)()
      Dim actualSettings As Byte()
      Dim new_dslist As List(Of String) = New List(Of String)()
      Dim new_jpath As List(Of String) = New List(Of String)()
      Dim new_jpath_len As List(Of Integer) = New List(Of Integer)()
      Dim new_val_arr As List(Of String) = New List(Of String)()
      Dim cpos As Integer = 0
      Dim eqpos As Integer = 0
      Dim leng As Integer = 0
      Dim i As Integer = 0
      Dim j As Integer = 0
      Dim njpath As String
      Dim jpath As String
      Dim fun As String
      Dim attr As String
      Dim value As String
      Dim url As String
      Dim tmp As String
      Dim new_calib As String
      Dim sensorType As String
      Dim unit_name As String
      Dim newval As String
      Dim oldval As String
      Dim old_calib As String
      Dim each_str As String
      Dim do_update As Boolean
      Dim found As Boolean
      tmp = YAPI.DefaultEncoding.GetString(settings)
      tmp = Me._get_json_path(tmp, "api")
      If (Not (tmp = "")) Then
        settings = YAPI.DefaultEncoding.GetBytes(tmp)
      End If
      oldval = ""
      newval = ""
      old_json_flat = Me._flattenJsonStruct(settings)
      old_dslist = Me._json_get_array(old_json_flat)



      For i_i = 0 To old_dslist.Count - 1
        each_str = Me._json_get_string(YAPI.DefaultEncoding.GetBytes(old_dslist(i_i)))
        REM // split json path and attr
        leng = (each_str).Length
        eqpos = each_str.IndexOf("=")
        If ((eqpos < 0) OrElse (leng = 0)) Then
          Me._throw(YAPI.INVALID_ARGUMENT, "Invalid settings")
          Return YAPI.INVALID_ARGUMENT
        End If
        jpath = (each_str).Substring( 0, eqpos)
        eqpos = eqpos + 1
        value = (each_str).Substring( eqpos, leng - eqpos)
        old_jpath.Add(jpath)
        old_jpath_len.Add((jpath).Length)
        old_val_arr.Add(value)
      Next i_i




      actualSettings = Me._download("api.json")
      actualSettings = Me._flattenJsonStruct(actualSettings)
      new_dslist = Me._json_get_array(actualSettings)



      For i_i = 0 To new_dslist.Count - 1
        REM // remove quotes
        each_str = Me._json_get_string(YAPI.DefaultEncoding.GetBytes(new_dslist(i_i)))
        REM // split json path and attr
        leng = (each_str).Length
        eqpos = each_str.IndexOf("=")
        If ((eqpos < 0) OrElse (leng = 0)) Then
          Me._throw(YAPI.INVALID_ARGUMENT, "Invalid settings")
          Return YAPI.INVALID_ARGUMENT
        End If
        jpath = (each_str).Substring( 0, eqpos)
        eqpos = eqpos + 1
        value = (each_str).Substring( eqpos, leng - eqpos)
        new_jpath.Add(jpath)
        new_jpath_len.Add((jpath).Length)
        new_val_arr.Add(value)
      Next i_i




      i = 0
      While (i < new_jpath.Count)
        njpath = new_jpath(i)
        leng = (njpath).Length
        cpos = njpath.IndexOf("/")
        If ((cpos < 0) OrElse (leng = 0)) Then
          Continue While
        End If
        fun = (njpath).Substring( 0, cpos)
        cpos = cpos + 1
        attr = (njpath).Substring( cpos, leng - cpos)
        do_update = True
        If (fun = "services") Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "firmwareRelease")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "usbCurrent")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "upTime")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "persistentSettings")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "adminPassword")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "userPassword")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "rebootCountdown")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "advertisedValue")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "poeCurrent")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "readiness")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "ipAddress")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "subnetMask")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "router")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "linkQuality")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "ssid")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "channel")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "security")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "message")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "currentValue")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "currentRawValue")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "currentRunIndex")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "pulseTimer")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "lastTimePressed")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "lastTimeReleased")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "filesCount")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "freeSpace")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "timeUTC")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "rtcTime")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "unixTime")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "dateTime")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "rawValue")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "lastMsg")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "delayedPulseTimer")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "rxCount")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "txCount")) Then
          do_update = False
        End If
        If ((do_update) AndAlso (attr = "msgCount")) Then
          do_update = False
        End If
        If (do_update) Then
          do_update = False
          newval = new_val_arr(i)
          j = 0
          found = False
          While ((j < old_jpath.Count) AndAlso Not (found))
            If ((new_jpath_len(i) = old_jpath_len(j)) AndAlso (new_jpath(i) = old_jpath(j))) Then
              found = True
              oldval = old_val_arr(j)
              If (Not (newval = oldval)) Then
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
            new_calib = newval
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
            Me._download(url)
          Else
            url = "api/" + fun + ".json?" + attr + "=" + Me._escapeAttr(oldval)
            If (attr = "resolution") Then
              restoreLast.Add(url)
            Else
              Me._download(url)
            End If
          End If
        End If
        i = i + 1
      End While

      For i_i = 0 To restoreLast.Count - 1
        Me._download(restoreLast(i_i))
      Next i_i
      Me.clearCache()
      Return YAPI.SUCCESS
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
    '''   On failure, throws an exception or returns  <c>YAPI_INVALID_STRING</c>.
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
    '''   exceeds 1536 bytes.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a binary buffer with module icon, in png format.
    '''   On failure, throws an exception or returns  <c>YAPI_INVALID_STRING</c>.
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
    '''   On failure, throws an exception or returns  <c>YAPI_INVALID_STRING</c>.
    ''' </returns>
    '''/
    Public Overridable Function get_lastLogs() As String
      Dim content As Byte()

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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
  '''   The YSensor class is the parent class for all Yoctopuce sensors.
  ''' <para>
  '''   It can be
  '''   used to read the current value and unit of any sensor, read the min/max
  '''   value, configure autonomous recording frequency and access recorded data.
  '''   It also provide a function to register a callback invoked each time the
  '''   observed value changes, or at a predefined interval. Using this class rather
  '''   than a specific subclass makes it possible to create generic applications
  '''   that work with any Yoctopuce sensor, even those that do not yet exist.
  '''   Note: The YAnButton class is the only analog input which does not inherit
  '''   from YSensor.
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
    Protected _isScal As Boolean
    Protected _isScal32 As Boolean
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
        _currentValue = Math.Round(json_val.getDouble("currentValue") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("lowestValue") Then
        _lowestValue = Math.Round(json_val.getDouble("lowestValue") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("highestValue") Then
        _highestValue = Math.Round(json_val.getDouble("highestValue") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("currentRawValue") Then
        _currentRawValue = Math.Round(json_val.getDouble("currentRawValue") * 1000.0 / 65536.0) / 1000.0
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
        _resolution = Math.Round(json_val.getDouble("resolution") * 1000.0 / 65536.0) / 1000.0
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
    '''   On failure, throws an exception or returns <c>Y_UNIT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_unit() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current value of the measure, in the specified unit,
    '''   as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   On failure, throws an exception or returns <c>Y_LOWESTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lowestValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   On failure, throws an exception or returns <c>Y_HIGHESTVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_highestValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return HIGHESTVALUE_INVALID
        End If
      End If
      res = Me._highestValue * Me._iresol
      res = Math.Round(res) / Me._iresol
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the uncalibrated, unrounded raw value returned by the sensor, in the specified unit, as a floating point number.
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
    '''   On failure, throws an exception or returns <c>Y_CURRENTRAWVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentRawValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_LOGFREQUENCY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logFrequency() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   the value "OFF".
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   On failure, throws an exception or returns <c>Y_REPORTFREQUENCY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_reportFrequency() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   hour (eg. "4/h"). To disable timed value notifications for this
    '''   function, use the value "OFF".
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   a value among <c>Y_ADVMODE_IMMEDIATE</c>, <c>Y_ADVMODE_PERIOD_AVG</c>, <c>Y_ADVMODE_PERIOD_MIN</c>
    '''   and <c>Y_ADVMODE_PERIOD_MAX</c> corresponding to the measuring mode used for the advertised value
    '''   pushed to the parent hub
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_ADVMODE_IMMEDIATE</c>, <c>Y_ADVMODE_PERIOD_AVG</c>, <c>Y_ADVMODE_PERIOD_MIN</c>
    '''   and <c>Y_ADVMODE_PERIOD_MAX</c> corresponding to the measuring mode used for the advertised value
    '''   pushed to the parent hub
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the resolution of the measured values
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RESOLUTION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_resolution() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RESOLUTION_INVALID
        End If
      End If
      res = Me._resolution
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensor health state code, which is zero when there is an up-to-date measure
    '''   available or a positive code if the sensor is not able to provide a measure right now.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the sensor health state code, which is zero when there is an up-to-date measure
    '''   available or a positive code if the sensor is not able to provide a measure right now
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SENSORSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_sensorState() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   a string that uniquely characterizes the sensor
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
      Me._isScal32 = False
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
        REM // New 32bit text format
        Me._isScal = True
        Me._isScal32 = True
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
        Me._isScal = (iCalib(1) > 0)
        If (Me._isScal) Then
          Me._offset = iCalib(0)
          If (Me._offset > 32767) Then
            Me._offset = Me._offset - 65536
          End If
          Me._scale = iCalib(1)
          Me._decexp = 0
        Else
          Me._offset = 0
          Me._scale = 1
          Me._decexp = 1.0
          position = iCalib(0)
          While (position > 0)
            Me._decexp = Me._decexp * 10
            position = position - 1
          End While
        End If
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
          If (Me._isScal) Then
            fRaw = iRaw
            fRaw = (fRaw - Me._offset) / Me._scale
            fRef = iRef
            fRef = (fRef - Me._offset) / Me._scale
            Me._calraw.Add(fRaw)
            Me._calref.Add(fRef)
          Else
            Me._calraw.Add(YAPI._decimalToDouble(iRaw))
            Me._calref.Add(YAPI._decimalToDouble(iRef))
          End If
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
    '''   Returns the YDatalogger object of the device hosting the sensor.
    ''' <para>
    '''   This method returns an object of
    '''   class YDatalogger that can control global parameters of the data logger. The returned object
    '''   should not be freed.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an YDataLogger object or Nothing on error.
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
      logger  = YDataLogger.FindDataLogger(hwid)
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    '''/
    Public Overridable Function startDataLogger() As Integer
      Dim res As Byte()

      res = Me._download("api/dataLogger/recording?recording=1")
      If Not((res).Length>0) Then
        me._throw( YAPI.IO_ERROR,  "unable to start datalogger")
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    '''/
    Public Overridable Function stopDataLogger() As Integer
      Dim res As Byte()

      res = Me._download("api/dataLogger/recording?recording=0")
      If Not((res).Length>0) Then
        me._throw( YAPI.IO_ERROR,  "unable to stop datalogger")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a DataSet object holding historical data for this
    '''   sensor, for a specified time interval.
    ''' <para>
    '''   The measures will be
    '''   retrieved from the data logger, which must have been turned
    '''   on at the desired time. See the documentation of the DataSet
    '''   class for information on how to get an overview of the
    '''   recorded data, and how to load progressively a large set
    '''   of measures from the data logger.
    ''' </para>
    ''' <para>
    '''   This function only works if the device uses a recent firmware,
    '''   as DataSet objects are not supported by firmwares older than
    '''   version 13000.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="startTime">
    '''   the start of the desired measure time interval,
    '''   as a Unix timestamp, i.e. the number of seconds since
    '''   January 1, 1970 UTC. The special value 0 can be used
    '''   to include any meaasure, without initial limit.
    ''' </param>
    ''' <param name="endTime">
    '''   the end of the desired measure time interval,
    '''   as a Unix timestamp, i.e. the number of seconds since
    '''   January 1, 1970 UTC. The special value 0 can be used
    '''   to include any meaasure, without ending limit.
    ''' </param>
    ''' <returns>
    '''   an instance of YDataSet, providing access to historical
    '''   data. Past measures can be loaded progressively
    '''   using methods from the YDataSet object.
    ''' </returns>
    '''/
    Public Overridable Function get_recordedData(startTime As Long, endTime As Long) As YDataSet
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
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function loadCalibrationPoints(rawValues As List(Of Double), refValues As List(Of Double)) As Integer
      Dim i_i As Integer
      rawValues.Clear()
      refValues.Clear()
      REM // Load function parameters if not yet loaded
      If (Me._scale = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return YAPI.DEVICE_NOT_FOUND
        End If
      End If
      If (Me._caltyp < 0) Then
        Me._throw(YAPI.NOT_SUPPORTED, "Calibration parameters format mismatch. Please upgrade your library or firmware.")
        Return YAPI.NOT_SUPPORTED
      End If
      rawValues.Clear()
      refValues.Clear()
      For i_i = 0 To Me._calraw.Count - 1
        rawValues.Add(Me._calraw(i_i))
      Next i_i
      For i_i = 0 To Me._calref.Count - 1
        refValues.Add(Me._calref(i_i))
      Next i_i
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function _encodeCalibrationPoints(rawValues As List(Of Double), refValues As List(Of Double)) As String
      Dim res As String
      Dim npt As Integer = 0
      Dim idx As Integer = 0
      Dim iRaw As Integer = 0
      Dim iRef As Integer = 0
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
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return YAPI.INVALID_STRING
        End If
      End If
      REM // Detect old firmware
      If ((Me._caltyp < 0) OrElse (Me._scale < 0)) Then
        Me._throw(YAPI.NOT_SUPPORTED, "Calibration parameters format mismatch. Please upgrade your library or firmware.")
        Return "0"
      End If
      If (Me._isScal32) Then
        REM // 32-bit fixed-point encoding
        res = "" + Convert.ToString(YOCTO_CALIB_TYPE_OFS)
        idx = 0
        While (idx < npt)
          res = "" +  res + "," + YAPI._floatToStr( rawValues(idx)) + "," + YAPI._floatToStr(refValues(idx))
          idx = idx + 1
        End While
      Else
        If (Me._isScal) Then
          REM // 16-bit fixed-point encoding
          res = "" + Convert.ToString(npt)
          idx = 0
          While (idx < npt)
            iRaw = CType(Math.Round(rawValues(idx) * Me._scale + Me._offset), Integer)
            iRef = CType(Math.Round(refValues(idx) * Me._scale + Me._offset), Integer)
            res = "" +  res + "," + Convert.ToString( iRaw) + "," + Convert.ToString(iRef)
            idx = idx + 1
          End While
        Else
          REM // 16-bit floating-point decimal encoding
          res = "" + Convert.ToString(10 + npt)
          idx = 0
          While (idx < npt)
            iRaw = CType(YAPI._doubleToDecimal(rawValues(idx)), Integer)
            iRef = CType(YAPI._doubleToDecimal(refValues(idx)), Integer)
            res = "" +  res + "," + Convert.ToString( iRaw) + "," + Convert.ToString(iRef)
            idx = idx + 1
          End While
        End If
      End If
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

    Public Overridable Function _decodeTimedReport(timestamp As Double, report As List(Of Integer)) As YMeasure
      Dim i As Integer = 0
      Dim byteVal As Integer = 0
      Dim poww As Integer = 0
      Dim minRaw As Integer = 0
      Dim avgRaw As Integer = 0
      Dim maxRaw As Integer = 0
      Dim sublen As Integer = 0
      Dim difRaw As Integer = 0
      Dim startTime As Double = 0
      Dim endTime As Double = 0
      Dim minVal As Double = 0
      Dim avgVal As Double = 0
      Dim maxVal As Double = 0
      startTime = Me._prevTimedReport
      endTime = timestamp
      Me._prevTimedReport = endTime
      If (startTime = 0) Then
        startTime = endTime
      End If
      If (report(0) = 2) Then
        REM // 32bit timed report format
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
          sublen = 1 + ((((report(1)) >> (2))) And (3))
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
          sublen = 1 + ((((report(1)) >> (4))) And (3))
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
      Else
        REM // 16bit timed report format
        If (report(0) = 0) Then
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
          If (Me._isScal) Then
            avgVal = Me._decodeVal(avgRaw)
          Else
            If (((byteVal) And (&H80)) <> 0) Then
              avgRaw = avgRaw - poww
            End If
            avgVal = Me._decodeAvg(avgRaw)
          End If
          minVal = avgVal
          maxVal = avgVal
        Else
          REM // averaged report 2+4+2 bytes
          minRaw = report(1) + &H100 * report(2)
          maxRaw = report(3) + &H100 * report(4)
          avgRaw = report(5) + &H100 * report(6) + &H10000 * report(7)
          byteVal = report(8)
          If (((byteVal) And (&H80)) = 0) Then
            avgRaw = avgRaw + &H1000000 * byteVal
          Else
            avgRaw = avgRaw - &H1000000 * (&H100 - byteVal)
          End If
          minVal = Me._decodeVal(minRaw)
          avgVal = Me._decodeAvg(avgRaw)
          maxVal = Me._decodeVal(maxRaw)
        End If
      End If
      Return New YMeasure(startTime, endTime, minVal, avgVal, maxVal)
    End Function

    Public Overridable Function _decodeVal(w As Integer) As Double
      Dim val As Double = 0
      val = w
      If (Me._isScal) Then
        val = (val - Me._offset) / Me._scale
      Else
        val = YAPI._decimalToDouble(w)
      End If
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
      If (Me._isScal) Then
        val = (val / 100 - Me._offset) / Me._scale
      Else
        val = val / Me._decexp
      End If
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
  '''   a string that uniquely characterizes the sensor
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
  Public Const Y_CLEARHISTORY_FALSE As Integer = 0
  Public Const Y_CLEARHISTORY_TRUE As Integer = 1
  Public Const Y_CLEARHISTORY_INVALID As Integer = -1
  Public Delegate Sub YDataLoggerValueCallback(ByVal func As YDataLogger, ByVal value As String)
  Public Delegate Sub YDataLoggerTimedReportCallback(ByVal func As YDataLogger, ByVal measure As YMeasure)
  REM --- (end of generated code: YDataLogger globals)


  Public Class YOldDataStream
    Inherits YDataStream
    Protected _dataLogger As YDataLogger
    Protected _timeStamp As Long
    Protected _interval As Long

    Public Sub New(parent As YDataLogger, run As Integer, ByVal stamp As Integer, ByVal utc As Long, ByVal itv As Integer)
      MyBase.new(parent)
      _dataLogger = parent
      _runNo = run
      _timeStamp = stamp
      _utcStamp = CUInt(utc)
      _interval = itv
      _samplesPerHour = CInt((3600 / _interval))
      _isClosed = True
      _minVal = DATA_INVALID
      _avgVal = DATA_INVALID
      _maxVal = DATA_INVALID
    End Sub

    Protected Overridable Overloads Sub Dispose()
      _columnNames = Nothing
      _values = Nothing
    End Sub


    '''*
    ''' <summary>
    '''   Returns the relative start time of the data stream, measured in seconds.
    ''' <para>
    '''   For recent firmwares, the value is relative to the present time,
    '''   which means the value is always negative.
    '''   If the device uses a firmware older than version 13000, value is
    '''   relative to the start of the time the device was powered on, and
    '''   is always positive.
    '''   If you need an absolute UTC timestamp, use <c>get_startTimeUTC()</c>.
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
    Public Overloads Function get_startTime() As Integer
      get_startTime = CInt(_timeStamp)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of seconds elapsed between  two consecutive
    '''   rows of this data stream.
    ''' <para>
    '''   By default, the data logger records one row
    '''   per second, but there might be alternative streams at lower resolution
    '''   created by summarizing the original stream for archiving purposes.
    ''' </para>
    ''' <para>
    '''   This method does not cause any access to the device, as the value
    '''   is preloaded in the object at instantiation time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an unsigned number corresponding to a number of seconds.
    ''' </returns>
    '''/
    Public Overloads Function get_dataSamplesInterval() As Integer
      get_dataSamplesInterval = CInt(_interval)
    End Function

    Private Overloads Function loadStream() As Integer
      Dim raw_json As YJSONContent = Nothing
      Dim json As YJSONObject
      Dim res As Integer
      Dim coldiv As List(Of Integer) = New List(Of Integer)
      Dim coltype As List(Of Integer) = New List(Of Integer)
      Dim udat As List(Of Integer) = New List(Of Integer)
      Dim [date] As New List(Of Double)()
      Dim colscl As List(Of Double) = New List(Of Double)
      Dim colofs As List(Of Integer) = New List(Of Integer)
      Dim caltyp As List(Of Integer) = New List(Of Integer)
      Dim calhdl As List(Of yCalibrationHandler) = New List(Of yCalibrationHandler)
      Dim calpar As List(Of List(Of Integer)) = New List(Of List(Of Integer))
      Dim calraw As List(Of List(Of Double)) = New List(Of List(Of Double))
      Dim calref As List(Of List(Of Double)) = New List(Of List(Of Double))

      Dim x As Integer = 0
      Dim j As Integer = 0

      res = _dataLogger.getData(_runNo, _timeStamp, raw_json)
      If (res <> YAPI.SUCCESS) Then
        Return res
      End If

      _nRows = 0
      _nCols = 0
      _columnNames.Clear()
      _values = New List(Of List(Of Double))()
      json = DirectCast(raw_json, YJSONObject)


      If json.has("time") Then
        _timeStamp = json.getInt("time")
      End If
      If json.has("UTC") Then
        _utcStamp = json.getLong("UTC")
      End If
      If json.has("interval") Then
        _interval = json.getInt("interval")
      End If
      If json.has("nRows") Then
        _nRows = json.getInt("nRows")
      End If
      If json.has("keys") Then
        Dim jsonkeys As YJSONArray = json.getYJSONArray("keys")
        _nCols = jsonkeys.Length
        For j = 0 To _nCols - 1
          _columnNames.Add(jsonkeys.getString(j))
        Next
      End If
      If json.has("div") Then
        Dim arr As YJSONArray = json.getYJSONArray("div")
        _nCols = arr.Length
        For j = 0 To _nCols - 1
          coldiv.Add(arr.getInt(j))
        Next
      End If
      If json.has("type") Then
        Dim arr As YJSONArray = json.getYJSONArray("type")
        _nCols = arr.Length
        For j = 0 To _nCols - 1
          coltype.Add(arr.getInt(j))
        Next
      End If
      If json.has("scal") Then
        Dim arr As YJSONArray = json.getYJSONArray("type")
        _nCols = arr.Length
        For j = 0 To _nCols - 1
          colscl.Add(arr.getInt(j) / 65536.0)
          If coltype(j) <> 0 Then
            colofs.Add(-32767)
          Else
            colofs.Add(0)
          End If

        Next
      End If
      If json.has("cal") Then
        REM old calibration is not supported
      End If
      If json.has("data") Then
        If colscl.Count <= 0 Then
          For j = 0 To _nCols - 1
            colscl.Add(1.0 / coldiv(j))
            If coltype(j) <> 0 Then
              colofs.Add(-32767)
            Else
              colofs.Add(0)
            End If
          Next
        End If
        udat.Clear()
        Dim data As String = Nothing
        Try
          data = json.getString("data")
          udat = YAPI._decodeWords(data)
        Catch generatedExceptionName As Exception
        End Try
        If data Is Nothing Then
          Dim jsonData As YJSONArray = json.getYJSONArray("data")
          For j = 0 To jsonData.Length - 1
            Dim tmp As Integer = CInt(jsonData.getInt(j))
            udat.Add(tmp)
          Next
        End If
        _values = New List(Of List(Of Double))()
        Dim dat As New List(Of Double)()
        For Each uval As Integer In udat
          Dim value As Double
          If coltype(x) < 2 Then
            value = (uval + colofs(x)) * colscl(x)
          Else
            value = YAPI._decimalToDouble(uval - 32767)
          End If
          If caltyp(x) > 0 AndAlso calhdl(x) IsNot Nothing Then
              Dim handler As yCalibrationHandler = calhdl(x)
            If caltyp(x) <= 10 Then
              value = handler((uval + colofs(x)) / coldiv(x), caltyp(x), calpar(x), calraw(x), calref(x))
            ElseIf caltyp(x) > 20 Then
              value = handler(value, caltyp(x), calpar(x), calraw(x), calref(x))
            End If
          End If
          dat.Add(value)
          x += 1
          If x = _nCols Then
            _values.Add(dat)
            dat.Clear()
            x = 0
          End If
        Next
      End If
      Return YAPI.SUCCESS
    End Function



  End Class

 REM --- (generated code: YDataLogger class start)

  '''*
  ''' <summary>
  '''   Yoctopuce sensors include a non-volatile memory capable of storing ongoing measured
  '''   data automatically, without requiring a permanent connection to a computer.
  ''' <para>
  '''   The DataLogger function controls the global parameters of the internal data
  '''   logger.
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
    '''   On failure, throws an exception or returns <c>Y_CURRENTRUNINDEX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentRunIndex() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   On failure, throws an exception or returns <c>Y_TIMEUTC_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_timeUTC() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   a value among <c>Y_RECORDING_OFF</c>, <c>Y_RECORDING_ON</c> and <c>Y_RECORDING_PENDING</c>
    '''   corresponding to the current activation state of the data logger
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RECORDING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_recording() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   a value among <c>Y_RECORDING_OFF</c>, <c>Y_RECORDING_ON</c> and <c>Y_RECORDING_PENDING</c>
    '''   corresponding to the activation state of the data logger to start/stop recording data
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   either <c>Y_AUTOSTART_OFF</c> or <c>Y_AUTOSTART_ON</c>, according to the default activation state
    '''   of the data logger on power up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_AUTOSTART_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_autoStart() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_AUTOSTART_OFF</c> or <c>Y_AUTOSTART_ON</c>, according to the default activation state
    '''   of the data logger on power up
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   either <c>Y_BEACONDRIVEN_OFF</c> or <c>Y_BEACONDRIVEN_ON</c>, according to true if the data logger
    '''   is synchronised with the localization beacon
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BEACONDRIVEN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_beaconDriven() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   either <c>Y_BEACONDRIVEN_OFF</c> or <c>Y_BEACONDRIVEN_ON</c>, according to the type of
    '''   synchronisation of the data logger
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    Public Function get_clearHistory() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
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
    '''   a string that uniquely characterizes the data logger
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   Returns a list of YDataSet objects that can be used to retrieve
    '''   all measures stored by the data logger.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   This function only works if the device uses a recent firmware,
    '''   as YDataSet objects are not supported by firmwares older than
    '''   version 13000.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a list of YDataSet object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty list.
    ''' </para>
    '''/
    Public Overridable Function get_dataSets() As List(Of YDataSet)
      Return Me.parse_dataSets(Me._download("logger.json"))
    End Function

    Public Overridable Function parse_dataSets(json As Byte()) As List(Of YDataSet)
      Dim i_i As Integer
      Dim dslist As List(Of String) = New List(Of String)()
      Dim dataset As YDataSet
      Dim res As List(Of YDataSet) = New List(Of YDataSet)()

      dslist = Me._json_get_array(json)
      res.Clear()
      For i_i = 0 To dslist.Count - 1
        dataset = New YDataSet(Me)
        dataset._parse(dslist(i_i))
        res.Add(dataset)
      Next i_i
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of data loggers started using <c>yFirstDataLogger()</c>.
    ''' <para>
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


    '''*
    ''' <summary>
    '''   Builds a list of all data streams hold by the data logger (legacy method).
    ''' <para>
    '''   The caller must pass by reference an empty array to hold YDataStream
    '''   objects, and the function fills it with objects describing available
    '''   data sequences.
    ''' </para>
    ''' <para>
    '''   This is the old way to retrieve data from the DataLogger.
    '''   For new applications, you should rather use <c>get_dataSets()</c>
    '''   method, or call directly <c>get_recordedData()</c> on the
    '''   sensor object.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="v">
    '''   an array of YDataStream objects to be filled in
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/

    Public Function get_dataStreams(ByVal v As List(Of YDataStream)) As Integer
      Dim raw_json As YJSONContent = Nothing
      Dim root As YJSONArray
      Dim i As Integer
      Dim res As Integer

      v.Clear()
      res = getData(0, 0, raw_json)
      If res <> YAPI_SUCCESS Then
        Return res
      End If
      root = DirectCast(raw_json, YJSONArray)


      If root.Length = 0 Then
        Return YAPI.SUCCESS
      End If
      If root.[get](0).getJSONType() = YJSONContent.YJSONType.ARRAY Then
        REM old datalogger format: [runIdx, timerel, utc, interval]
        For i = 0 To root.Length - 1
          Dim el As YJSONArray = root.getYJSONArray(i)
          v.Add(New YOldDataStream(Me, el.getInt(0), el.getInt(1), CUInt(el.getLong(2)), el.getInt(1)))
        Next
      Else
        REM new datalogger format: {"id":"...","unit":"...","streams":["...",...]}
        Dim json_buffer As String = root.toJSON()
        Dim sets As List(Of YDataSet) = parse_dataSets(YAPI.DefaultEncoding.GetBytes(json_buffer))
        For sj As Integer = 0 To sets.Count - 1
          Dim ds As List(Of YDataStream) = sets(sj).get_privateDataStreams()
          For si As Integer = 0 To ds.Count - 1
            v.Add(ds(si))
          Next
        Next
      End If
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
  '''   a string that uniquely characterizes the data logger
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
  '''   it  either fires the debugger or aborts (i.e. crash) the program.
  '''   On failure, throws an exception or returns a negative error code.
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
    Private _mod As YModule
    Private _value As String
    Private _report As List(Of Integer)
    Private _timestamp As Double

    Public Sub New(ByVal fun As YFunction, ByVal value As String)
      _fun = fun
      _mod = Nothing
      _value = value
      _report = Nothing
      _timestamp = 0
    End Sub

    Public Sub New(ByVal fun As YFunction, ByVal timestamp As Double, ByVal report As List(Of Integer))
      _fun = fun
      _mod = Nothing
      _value = Nothing
      _timestamp = timestamp
      _report = report
    End Sub

    Public Sub New(ByVal modul As YModule)
      _fun = Nothing
      _mod = modul
      _value = Nothing
      _report = Nothing
      _timestamp = 0
    End Sub

    Public Sub invoke()
      If (_fun IsNot Nothing) Then
        If (_value Is Nothing) Then
          Dim sensor As YSensor = CType(_fun, YSensor)
          Dim measure As YMeasure = sensor._decodeTimedReport(_timestamp, _report)
          sensor._invokeTimedReportCallback(measure)
        Else
          REM new value
          _fun._invokeValueCallback(_value)
        End If
      ElseIf (_mod IsNot Nothing) Then
        _mod._invokeConfigChangeCallback()
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
    Dim modul As YModule
    Dim errmsg As String = ""

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
    ev = New DataEvent(modul)
    _DataEvents.Add(ev)
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


  Private Sub native_yTimedReportCallback(ByVal fundescr As YFUN_DESCR, timestamp As Double, rawdata As IntPtr, len As System.UInt32)
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
          ev = New DataEvent(YFunction._TimedReportCallbackList(i), timestamp, report)
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
  Public Const Y_DETECT_USB As Integer = YAPI.DETECT_USB
  Public Const Y_DETECT_NET As Integer = YAPI.DETECT_NET
  Public Const Y_DETECT_ALL As Integer = YAPI.DETECT_ALL
  Public Const Y_RESEND_MISSING_PKT As Integer = YAPI.RESEND_MISSING_PKT
  Public Function yInitAPI(ByVal mode As Integer, ByRef errmsg As String) As Integer
    Return YAPI.InitAPI(mode, errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Frees dynamically allocated memory blocks used by the Yoctopuce library.
  ''' <para>
  '''   It is generally not required to call this function, unless you
  '''   want to free all dynamically allocated memory blocks in order to
  '''   track a memory leak for instance.
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
  '''   Setup the Yoctopuce library to use modules connected on a given machine.
  ''' <para>
  '''   The
  '''   parameter will determine how the API will work. Use the following values:
  ''' </para>
  ''' <para>
  '''   <b>usb</b>: When the <c>usb</c> keyword is used, the API will work with
  '''   devices connected directly to the USB bus. Some programming languages such a Javascript,
  '''   PHP, and Java don't provide direct access to USB hardware, so <c>usb</c> will
  '''   not work with these. In this case, use a VirtualHub or a networked YoctoHub (see below).
  ''' </para>
  ''' <para>
  '''   <b><i>x.x.x.x</i></b> or <b><i>hostname</i></b>: The API will use the devices connected to the
  '''   host with the given IP address or hostname. That host can be a regular computer
  '''   running a VirtualHub, or a networked YoctoHub such as YoctoHub-Ethernet or
  '''   YoctoHub-Wireless. If you want to use the VirtualHub running on you local
  '''   computer, use the IP address 127.0.0.1.
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
  '''   for this limitation is to setup the library to use the VirtualHub
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
  '''   You can call <i>RegisterHub</i> several times to connect to several machines.
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
  '''   <c>YAPI_SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure, throws an exception or returns a negative error code.
  ''' </para>
  '''/
  Public Function yRegisterHub(ByVal url As String, ByRef errmsg As String) As Integer
    Return YAPI.RegisterHub(url, errmsg)
  End Function

  '''*
  ''' <summary>
  '''   Fault-tolerant alternative to <c>RegisterHub()</c>.
  ''' <para>
  '''   This function has the same
  '''   purpose and same arguments as <c>RegisterHub()</c>, but does not trigger
  '''   an error when the selected hub is not available at the time of the function call.
  '''   This makes it possible to register a network hub independently of the current
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
  '''   <c>YAPI_SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure, throws an exception or returns a negative error code.
  ''' </para>
  '''/
  Public Function yPreregisterHub(ByVal url As String, ByRef errmsg As String) As Integer
    Return YAPI.PreregisterHub(url, errmsg)
  End Function
  '''*
  ''' <summary>
  '''   Setup the Yoctopuce library to no more use modules connected on a previously
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
  '''   hub is usable. The url parameter follow the same convention as the <c>RegisterHub</c>
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
  '''   <c>YAPI_SUCCESS</c> when the call succeeds.
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
  '''   and to make the application aware of hot-plug events.
  ''' </para>
  ''' </summary>
  ''' <param name="errmsg">
  '''   a string passed by reference to receive any error message.
  ''' </param>
  ''' <returns>
  '''   <c>YAPI_SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure, throws an exception or returns a negative error code.
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
  '''   <c>YAPI_SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure, throws an exception or returns a negative error code.
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
  '''   <c>YAPI_SUCCESS</c> when the call succeeds.
  ''' </returns>
  ''' <para>
  '''   On failure, throws an exception or returns a negative error code.
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
  '''   <c>YAPI_SUCCESS</c> when the call succeeds.
  '''   On failure, throws an exception or returns a negative error code.
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
  '''   <c>A..Z</c>, <c>a..z</c>, <c>0..9</c>, <c>_</c>, and <c>-</c>.
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

  Public Function yapiGetAllDevices(ByVal dbuffer As IntPtr, ByVal maxsize As Integer, ByVal neededsize As Integer, ByRef errmsg As String) As Integer
    Dim buffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    buffer.Length = 0
    yapiGetAllDevices = _yapiGetAllDevices(dbuffer, maxsize, neededsize, buffer)
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

  <DllImport("yapi.dll", EntryPoint:="yapiInitAPI", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiInitAPI(ByVal mode As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiFreeAPI", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiFreeAPI()
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
  Private Function _yapiGetAPIVersion(ByRef version As IntPtr, ByRef subversion As IntPtr) As yu16
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetDevice", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetDevice(ByVal device_str As StringBuilder, ByVal errmsg As StringBuilder) As YDEV_DESCR
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetAllDevices", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetAllDevices(ByVal buffer As IntPtr, ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetDeviceInfo", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetDeviceInfo(ByVal d As YDEV_DESCR, ByRef infos As yDeviceSt, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetFunction", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetFunction(ByVal class_str As StringBuilder, ByVal function_str As StringBuilder, ByVal errmsg As StringBuilder) As YFUN_DESCR
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetFunctionsByClass", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetFunctionsByClass(ByVal class_str As StringBuilder, ByVal precFuncDesc As YFUN_DESCR, ByVal buffer As IntPtr,
                                       ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetFunctionsByDevice", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetFunctionsByDevice(ByVal device As YDEV_DESCR, ByVal precFuncDesc As YFUN_DESCR, ByVal buffer As IntPtr, ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetFunctionInfoEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetFunctionInfoEx(ByVal fundesc As YFUN_DESCR,
                                         ByRef devdesc As YDEV_DESCR,
                                         ByVal serial As StringBuilder,
                                         ByVal funcId As StringBuilder,
                                         ByVal baseType As StringBuilder,
                                         ByVal funcName As StringBuilder,
                                         ByVal funcVal As StringBuilder,
                                         ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetErrorString", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetErrorString(ByVal errorcode As Integer, ByVal buffer As StringBuilder, ByVal maxsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncStart", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestSyncStart(ByRef iohdl As YIOHDL, ByVal device As StringBuilder, ByVal url As StringBuilder, ByRef reply As IntPtr, ByRef replysize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncStartEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestSyncStartEx(ByRef iohdl As YIOHDL, ByVal device As StringBuilder, ByVal url As IntPtr, ByVal urllen As Integer, ByRef reply As IntPtr, ByRef replysize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function


  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncDone", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestSyncDone(ByRef iohdl As YIOHDL, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestAsync", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestAsync(ByVal device As StringBuilder, ByVal url As IntPtr, ByVal callback As IntPtr, ByVal context As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestAsyncEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequestAsyncEx(ByVal device As StringBuilder, ByVal url As IntPtr, ByVal urllen As Integer, ByVal callback As IntPtr, ByVal context As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequest", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiHTTPRequest(ByVal device As StringBuilder, ByVal url As StringBuilder, ByVal buffer As StringBuilder, ByVal buffsize As Integer, ByRef fullsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetBootloadersDevs", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetBootloadersDevs(ByVal serials As StringBuilder, ByVal maxNbSerial As yu32, ByRef totalBootladers As yu32, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiGetDevicePath", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Function _yapiGetDevicePath(ByVal devdesc As Integer,
                                           ByVal rootdevice As StringBuilder,
                                           ByVal path As StringBuilder,
                                           ByVal pathsize As Integer,
                                           ByRef neededsize As Integer,
                                           ByVal errmsg As StringBuilder) As Integer
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

  <DllImport("yapi.dll", EntryPoint:="yapiStartStopDeviceLogCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)>
  Private Sub _yapiStartStopDeviceLogCallback(ByVal serial As StringBuilder, ByVal startStop As Integer)
  End Sub
  REM --- (generated code: YFunction dlldef)
  <DllImport("yapi.dll", EntryPoint:="yapiGetAllJsonKeys", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetAllJsonKeys(ByVal jsonbuffer As StringBuilder, ByVal out_buffer As StringBuilder, ByVal out_buffersize As Integer, ByRef fullsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiCheckFirmware", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiCheckFirmware(ByVal serial As StringBuilder, ByVal rev As StringBuilder, ByVal path As StringBuilder, ByVal buffer As StringBuilder, ByVal buffersize As Integer, ByRef fullsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetBootloaders", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetBootloaders(ByVal buffer As StringBuilder, ByVal buffersize As Integer, ByRef totalSize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiUpdateFirmwareEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiUpdateFirmwareEx(ByVal serial As StringBuilder, ByVal firmwarePath As StringBuilder, ByVal settings As StringBuilder, ByVal force As Integer, ByVal startUpdate As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncStartOutOfBand", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHTTPRequestSyncStartOutOfBand(ByRef iohdl As YIOHDL, ByVal channel As Integer, ByVal device As StringBuilder, ByVal request As StringBuilder, ByVal requestsize As Integer, ByRef reply As IntPtr, ByRef replysize As Integer, ByRef progress_cb As IntPtr, ByRef progress_ctx As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestAsyncOutOfBand", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHTTPRequestAsyncOutOfBand(ByVal channel As Integer, ByVal device As StringBuilder, ByVal request As StringBuilder, ByVal requestsize As Integer, ByRef callback As IntPtr, ByRef context As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiTestHub", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiTestHub(ByVal url As StringBuilder, ByVal mstimeout As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiJsonGetPath", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiJsonGetPath(ByVal path As StringBuilder, ByVal json_data As StringBuilder, ByVal json_len As Integer, ByRef result As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiJsonDecodeString", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiJsonDecodeString(ByVal json_data As StringBuilder, ByVal output As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiGetSubdevices", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetSubdevices(ByVal serial As StringBuilder, ByVal buffer As StringBuilder, ByVal buffersize As Integer, ByRef totalSize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function
  <DllImport("yapi.dll", EntryPoint:="yapiFreeMem", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiFreeMem(ByRef buffer As IntPtr)
  End Sub
  <DllImport("yapi.dll", EntryPoint:="yapiGetDevicePathEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetDevicePathEx(ByVal serial As StringBuilder, ByVal rootdevice As StringBuilder, ByVal path As StringBuilder, ByVal pathsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
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
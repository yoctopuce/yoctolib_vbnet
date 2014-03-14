'/********************************************************************
'*
'* $Id: yocto_api.vb 15376 2014-03-10 16:22:13Z seb $
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

Module yocto_api

  Public Enum TJSONRECORDTYPE
    JSON_STRING
    JSON_INTEGER
    JSON_BOOLEAN
    JSON_STRUCT
    JSON_ARRAY
  End Enum

  Public Structure TJSONRECORD
    Dim name As String
    Dim recordtype As TJSONRECORDTYPE
    Dim svalue As String
    Dim ivalue As Long
    Dim bvalue As Boolean
    Dim membercount As Integer
    Dim memberAllocated As Integer
    Dim members() As TJSONRECORD
    Dim itemcount As Integer
    Dim itemAllocated As Integer
    Dim items() As TJSONRECORD
  End Structure

  Public Class TJsonParser
    Private Enum Tjstate
      JSTART
      JWAITFORNAME
      JWAITFORENDOFNAME
      JWAITFORCOLON
      JWAITFORDATA
      JWAITFORNEXTSTRUCTMEMBER
      JWAITFORNEXTARRAYITEM
      JSCOMPLETED
      JWAITFORSTRINGVALUE
      JWAITFORINTVALUE
      JWAITFORBOOLVALUE
    End Enum

    Private Const JSONGRANULARITY As Integer = 10
    Public httpcode As Integer
    Private data As TJSONRECORD

    Public Sub New(ByVal jsonData As String)
      Me.New(jsonData, True)
    End Sub

    Public Sub New(ByVal jsonData As String, ByVal withHTTPHeader As Boolean)
      Const httpheader As String = "HTTP/1.1 "
      Const okHeader As String = "OK" + Chr(13) + Chr(10)
      Dim errmsg As String
      Dim p1, p2 As Integer
      Const CR As String = Chr(13) + Chr(10)

      If withHTTPHeader Then
        If Mid(jsonData, 1, Len(okHeader)) = okHeader Then
          httpcode = 200
        Else
          If Mid(jsonData, 1, Len(httpheader)) <> httpheader Then
            errmsg = "data should start with " + httpheader
            Throw New System.Exception(errmsg)
          End If
          p1 = InStr(Len(httpheader) + 1, jsonData, " ")
          p2 = InStr(jsonData, CR)
          httpcode = CInt(Val(Mid(jsonData, Len(httpheader), p1 - Len(httpheader))))
          If (httpcode <> 200) Then Exit Sub
        End If
        p1 = InStr(jsonData, CR + CR + "{") REM json data is a structure
        If p1 <= 0 Then p1 = InStr(jsonData, CR + CR + "[") REM json data is an array

        If p1 <= 0 Then
          errmsg = "data  does not contain JSON data "
          Throw New System.Exception(errmsg)
        End If

        jsonData = Mid(jsonData, p1 + 4, Len(jsonData) - p1 - 3)
      Else
        Dim start_struct As Integer = InStr(jsonData, "{") REM json data is a structure
        Dim start_array As Integer = InStr(jsonData, "[") REM json data is an array
        If ((start_struct < 0) And (start_array < 0)) Then
          errmsg = "data  does not contain JSON data "
          Throw New System.Exception(errmsg)
        End If

      End If
      data = CType(Parse(jsonData), TJSONRECORD)
    End Sub

    Public Function convertToString(ByVal p As Nullable(Of TJSONRECORD), ByVal showNamePrefix As Boolean) As String
      Dim buffer As String

      If (p Is Nothing) Then p = data

      If (p.Value.name <> "" And showNamePrefix) Then
        buffer = """" + p.Value.name + """:"
      Else
        buffer = ""
      End If

      Select Case p.Value.recordtype
        Case TJSONRECORDTYPE.JSON_STRING
          buffer = buffer + """" + p.Value.svalue + """"
        Case TJSONRECORDTYPE.JSON_INTEGER
          buffer = buffer + CStr(p.Value.ivalue)
        Case TJSONRECORDTYPE.JSON_BOOLEAN
          If p.Value.bvalue Then
            buffer = buffer + "TRUE"
          Else
            buffer = buffer + "FALSE"
          End If
        Case TJSONRECORDTYPE.JSON_STRUCT
          buffer = buffer + "{"
          For i As Integer = 0 To p.Value.membercount - 1
            If (i > 0) Then buffer = buffer + ","
            buffer = buffer + Me.convertToString(p.Value.members(i), True)
          Next i
          buffer = buffer + "}"
        Case TJSONRECORDTYPE.JSON_ARRAY
          buffer = buffer + "["
          For i As Integer = 0 To p.Value.itemcount - 1
            If (i > 0) Then buffer = buffer + ","
            buffer = buffer + Me.convertToString(p.Value.items(i), False)
          Next i
          buffer = buffer + "]"
      End Select

      Return buffer
    End Function



    Public Sub Dispose()
      freestructure(data)
    End Sub

    Public Function GetRootNode() As TJSONRECORD
      GetRootNode = data
    End Function

    Private Function Parse(ByVal st As String) As Nullable(Of TJSONRECORD)
      Dim i As Integer
      i = 1
      st = """root"" : " + st + " "
      Parse = ParseEx(Tjstate.JWAITFORNAME, "", st, i)
    End Function

    Private Sub ParseError(ByRef st As String, ByVal i As Integer, ByVal errmsg As String)
      Dim ststart, stend As Integer
      ststart = i - 10
      stend = i + 10
      If (ststart < 1) Then ststart = 1
      If (stend > Len(st)) Then stend = Len(st)
      errmsg = errmsg + " near " + Mid(st, ststart, i - ststart) + "*" + Mid(st, i, stend - i)
      Throw New System.Exception(errmsg)
    End Sub

    Private Function createStructRecord(ByVal name As String) As TJSONRECORD
      Dim res As TJSONRECORD
      res.recordtype = TJSONRECORDTYPE.JSON_STRUCT
      res.name = name
      res.svalue = ""
      res.ivalue = 0
      res.bvalue = False
      res.membercount = 0
      res.memberAllocated = JSONGRANULARITY
      ReDim Preserve res.members(res.memberAllocated - 1)
      res.itemcount = 0
      res.itemAllocated = 0
      res.items = Nothing
      createStructRecord = res
    End Function

    Private Function createArrayRecord(ByVal name As String) As TJSONRECORD
      Dim res As TJSONRECORD
      res.recordtype = TJSONRECORDTYPE.JSON_ARRAY
      res.name = name
      res.svalue = ""
      res.ivalue = 0
      res.bvalue = False
      res.itemcount = 0
      res.itemAllocated = JSONGRANULARITY
      ReDim Preserve res.items(res.itemAllocated - 1)
      res.membercount = 0
      res.memberAllocated = 0
      res.members = Nothing
      createArrayRecord = res
    End Function

    Private Function createStrRecord(ByVal name As String, ByVal value As String) As TJSONRECORD
      Dim res As TJSONRECORD
      res.recordtype = TJSONRECORDTYPE.JSON_STRING
      res.name = name
      res.svalue = value
      res.ivalue = 0
      res.bvalue = False
      res.itemcount = 0
      res.itemAllocated = 0
      res.items = Nothing
      res.membercount = 0
      res.memberAllocated = 0
      res.members = Nothing
      createStrRecord = res
    End Function

    Private Function createIntRecord(ByVal name As String, ByVal value As Long) As TJSONRECORD
      Dim res As TJSONRECORD
      res.recordtype = TJSONRECORDTYPE.JSON_INTEGER
      res.name = name
      res.svalue = ""
      res.ivalue = value
      res.bvalue = False
      res.itemcount = 0
      res.itemAllocated = 0
      res.items = Nothing
      res.membercount = 0
      res.memberAllocated = 0
      res.members = Nothing
      createIntRecord = res
    End Function

    Private Function createBoolRecord(ByVal name As String, ByVal value As Boolean) As TJSONRECORD
      Dim res As TJSONRECORD
      res.recordtype = TJSONRECORDTYPE.JSON_BOOLEAN
      res.name = name
      res.svalue = ""
      res.ivalue = 0
      res.bvalue = value
      res.itemcount = 0
      res.itemAllocated = 0
      res.items = Nothing
      res.membercount = 0
      res.memberAllocated = 0
      res.members = Nothing
      createBoolRecord = res
    End Function

    Private Sub add2StructRecord(ByRef container As TJSONRECORD, ByRef element As TJSONRECORD)
      If container.recordtype <> TJSONRECORDTYPE.JSON_STRUCT Then Throw New System.Exception("container is not a struct type")
      If (container.membercount >= container.memberAllocated) Then
        ReDim Preserve container.members(0 To container.memberAllocated + JSONGRANULARITY - 1)
        container.memberAllocated = container.memberAllocated + JSONGRANULARITY
      End If
      container.members(container.membercount) = element
      container.membercount = container.membercount + 1
    End Sub

    Private Sub add2ArrayRecord(ByRef container As TJSONRECORD, ByRef element As TJSONRECORD)
      If container.recordtype <> TJSONRECORDTYPE.JSON_ARRAY Then Throw New System.Exception("container is not an array type")
      If (container.itemcount >= container.itemAllocated) Then
        ReDim Preserve container.items(0 To container.itemAllocated + JSONGRANULARITY - 1)
        container.itemAllocated = container.itemAllocated + JSONGRANULARITY
      End If
      container.items(container.itemcount) = element
      container.itemcount = container.itemcount + 1
    End Sub

    Private Function Skipgarbage(ByRef st As String, ByRef i As Integer) As Char
      Dim sti As Char = CChar(Mid(st, i, 1))
      While (i <= Len(st) And (sti = Chr(32) Or sti = Chr(13) Or sti = Chr(10)))
        i = i + 1
        If (i <= Len(st)) Then sti = CChar(Mid(st, i, 1))
      End While
      Skipgarbage = sti
    End Function


    Private Function ParseEx(ByVal initialstate As Tjstate, ByVal defaultname As String, ByRef st As String, ByRef i As Integer) As Nullable(Of TJSONRECORD)
      Dim res, value As TJSONRECORD
      Dim state As Tjstate
      Dim svalue As String = ""
      Dim ivalue, isign As Long
      Dim sti As Char

      Dim name As String

      name = defaultname
      state = initialstate
      isign = 1
      res = Nothing
      ivalue = 0

      While i < Len(st)
        sti = CChar(Mid(st, i, 1))
        Select Case state
          Case Tjstate.JWAITFORNAME
            If sti = """" Then
              state = Tjstate.JWAITFORENDOFNAME
            Else
              If Asc(sti) <> 32 And Asc(sti) <> 13 And Asc(sti) <> 10 Then ParseError(st, i, "invalid char: was expecting """)
            End If

          Case Tjstate.JWAITFORENDOFNAME
            If sti = """" Then
              state = Tjstate.JWAITFORCOLON
            Else
              If Asc(sti) >= 32 Then name = name + sti Else ParseError(st, i, "invalid char: was expecting an identifier compliant char")
            End If

          Case Tjstate.JWAITFORCOLON
            If sti = ":" Then
              state = Tjstate.JWAITFORDATA
            Else
              If Asc(sti) <> 32 And Asc(sti) <> 13 And Asc(sti) <> 10 Then ParseError(st, i, "invalid char: was expecting """)
            End If
          Case Tjstate.JWAITFORDATA
            If sti = "{" Then
              res = createStructRecord(name)
              state = Tjstate.JWAITFORNEXTSTRUCTMEMBER
            ElseIf sti = "[" Then
              res = createArrayRecord(name)
              state = Tjstate.JWAITFORNEXTARRAYITEM
            ElseIf sti = """" Then
              svalue = ""
              state = Tjstate.JWAITFORSTRINGVALUE
            ElseIf sti >= "0" And sti <= "9" Then
              state = Tjstate.JWAITFORINTVALUE
              ivalue = Asc(sti) - 48
              isign = 1
            ElseIf sti = "-" Then
              state = Tjstate.JWAITFORINTVALUE
              ivalue = 0
              isign = -1
            ElseIf UCase(sti) = "T" Or UCase(sti) = "F" Then
              svalue = UCase(sti)
              state = Tjstate.JWAITFORBOOLVALUE
            ElseIf Asc(sti) <> 32 And Asc(sti) <> 13 And Asc(sti) <> 10 Then
              ParseError(st, i, "invalid char: was expecting  "",0..9,t or f")
            End If
          Case Tjstate.JWAITFORSTRINGVALUE
            If sti = """" Then
              state = Tjstate.JSCOMPLETED
              res = createStrRecord(name, svalue)
            ElseIf Asc(sti) < 32 Then
              ParseError(st, i, "invalid char: was expecting string value")
            Else
              svalue = svalue + sti
            End If
          Case Tjstate.JWAITFORINTVALUE
            If sti >= "0" And sti <= "9" Then
              ivalue = (ivalue * 10) + Asc(sti) - 48
            Else
              res = createIntRecord(name, isign * ivalue)
              state = Tjstate.JSCOMPLETED
              i = i - 1
            End If
          Case Tjstate.JWAITFORBOOLVALUE
            If UCase(sti) < "A" Or UCase(sti) > "Z" Then
              If svalue <> "TRUE" And svalue <> "FALSE" Then ParseError(st, i, "unexpected value, was expecting ""true"" or ""false""")
              If svalue = "TRUE" Then res = createBoolRecord(name, True) Else res = createBoolRecord(name, False)
              state = Tjstate.JSCOMPLETED
              i = i - 1
            Else
              svalue = svalue + UCase(sti)
            End If
          Case Tjstate.JWAITFORNEXTSTRUCTMEMBER
            sti = Skipgarbage(st, i)
            If (i <= Len(st)) Then
              If sti = "}" Then
                ParseEx = res
                i = i + 1
                Exit Function
              Else
                value = CType(ParseEx(Tjstate.JWAITFORNAME, "", st, i), TJSONRECORD)
                add2StructRecord(res, value)
                sti = Skipgarbage(st, i)
                If i < Len(st) Then
                  If sti = "}" And i < Len(st) Then
                    i = i - 1
                  ElseIf Asc(sti) <> 32 And Asc(sti) <> 13 And Asc(sti) <> 10 And sti <> "," Then
                    ParseError(st, i, "invalid char: vas expecting , or }")
                  End If
                End If
              End If

            End If
          Case Tjstate.JWAITFORNEXTARRAYITEM
            sti = Skipgarbage(st, i)
            If i < Len(st) Then
              If sti = "]" Then
                ParseEx = res
                i = i + 1
                Exit Function
              Else
                value = CType(ParseEx(Tjstate.JWAITFORDATA, Str(res.itemcount), st, i), TJSONRECORD)
                add2ArrayRecord(res, value)
                sti = Skipgarbage(st, i)
                If i < Len(st) Then
                  If sti = "]" And i < Len(st) Then
                    i = i - 1
                  ElseIf Asc(sti) <> 32 And Asc(sti) <> 13 And Asc(sti) <> 10 And sti <> "," Then
                    ParseError(st, i, "invalid char: vas expecting , or ]")
                  End If
                End If
              End If
            End If
          Case Tjstate.JSCOMPLETED
            ParseEx = res
            Exit Function
        End Select
        i = i + 1
      End While
      ParseError(st, i, "unexpected end of data")
      ParseEx = Nothing
    End Function

    Private Sub DumpStructureRec(ByRef p As TJSONRECORD, ByRef deep As Integer)
      Dim line, indent As String
      Dim i As Integer
      line = ""
      indent = ""
      For i = 0 To deep * 2
        indent = indent + " "
      Next i
      line = indent + p.name + ":"
      Select Case p.recordtype
        Case TJSONRECORDTYPE.JSON_STRING
          line = line + " str=" + p.svalue
          Console.WriteLine(line)
        Case TJSONRECORDTYPE.JSON_INTEGER
          line = line + " int =" + Str(p.ivalue)
          Console.WriteLine(line)
        Case TJSONRECORDTYPE.JSON_BOOLEAN
          If p.bvalue Then line = line + " bool = TRUE" Else line = line + " bool = FALSE"
          Console.WriteLine(line)
        Case TJSONRECORDTYPE.JSON_STRUCT
          Console.WriteLine(line + " struct")
          For i = 0 To p.membercount - 1
            DumpStructureRec(p.members(i), deep + 1)
          Next i
        Case TJSONRECORDTYPE.JSON_ARRAY
          Console.WriteLine(line + " array")
          For i = 0 To p.itemcount - 1
            DumpStructureRec(p.items(i), deep + 1)
          Next i
      End Select
    End Sub


    Private Sub freestructure(ByRef p As TJSONRECORD)
      Select Case p.recordtype
        Case TJSONRECORDTYPE.JSON_STRUCT
          For i As Integer = p.membercount - 1 To 0 Step -1
            freestructure(p.members(i))
          Next i
          ReDim p.members(0)

        Case TJSONRECORDTYPE.JSON_ARRAY
          For i As Integer = p.itemcount - 1 To 0 Step -1
            freestructure(p.items(i))
          Next i
          ReDim p.items(0)
      End Select
    End Sub


    Public Sub DumpStructure()
      DumpStructureRec(data, 0)
    End Sub




    Public Function GetChildNode(ByVal parent As Nullable(Of TJSONRECORD), ByVal nodename As String) As Nullable(Of TJSONRECORD)
      Dim i, index As Integer
      Dim p As Nullable(Of TJSONRECORD) = parent

      If p Is Nothing Then p = data

      If p.Value.recordtype = TJSONRECORDTYPE.JSON_STRUCT Then
        For i = 0 To p.Value.membercount - 1
          If p.Value.members(i).name = nodename Then
            GetChildNode = p.Value.members(i)
            Exit Function
          End If

        Next
      ElseIf p.Value.recordtype = TJSONRECORDTYPE.JSON_ARRAY Then
        index = CInt(Val(nodename))
        If (index >= p.Value.itemcount) Then Throw New System.Exception("index out of bounds " + nodename + ">=" + Str(p.Value.itemcount))
        GetChildNode = p.Value.items(index)
        Exit Function
      End If

      GetChildNode = Nothing
    End Function

    Public Function GetAllChilds(ByVal parent As Nullable(Of TJSONRECORD)) As List(Of String)
      Dim res As List(Of String) = New List(Of String)()
      Dim p As Nullable(Of TJSONRECORD) = parent

      If p Is Nothing Then p = data

      If (p.Value.recordtype = TJSONRECORDTYPE.JSON_STRUCT) Then
        For i As Integer = 0 To p.Value.membercount - 1
          res.Add(Me.convertToString(p.Value.members(i), False))
        Next i
      ElseIf (p.Value.recordtype = TJSONRECORDTYPE.JSON_ARRAY) Then
        For i As Integer = 0 To p.Value.itemcount - 1
          res.Add(Me.convertToString(p.Value.items(i), False))
        Next i
      End If
      Return res
    End Function
  End Class

  Public Const YOCTO_API_VERSION_STR As String = "1.10"
  Public Const YOCTO_API_VERSION_BCD As Integer = &H110
  Public Const YOCTO_API_BUILD_NO As String = "15466"

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
    Public Const DefaultCacheValidity As Integer = 5

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
    Public Const DETECT_ALL As Integer = DETECT_USB Or DETECT_NET

    REM --- (generated code: YFunction return codes)
    REM Yoctopuce error codes, also used by default as function return value
    Public Const SUCCESS As Integer = 0         REM everything worked allright
    Public Const NOT_INITIALIZED As Integer = -1 REM call yInitAPI() first !
    Public Const INVALID_ARGUMENT As Integer = -2 REM one of the arguments passed to the function is invalid
    Public Const NOT_SUPPORTED As Integer = -3  REM the operation attempted is (currently) not supported
    Public Const DEVICE_NOT_FOUND As Integer = -4 REM the requested device is not reachable
    Public Const VERSION_MISMATCH As Integer = -5 REM the device firmware is incompatible with this API version
    Public Const DEVICE_BUSY As Integer = -6    REM the device is busy with another task and cannot answer
    Public Const TIMEOUT As Integer = -7        REM the device took too long to provide an answer
    Public Const IO_ERROR As Integer = -8       REM there was an I/O problem while talking to the device
    Public Const NO_MORE_DATA As Integer = -9   REM there is no more data to read from
    Public Const EXHAUSTED As Integer = -10     REM you have run out of a limited ressource, check the documentation
    Public Const DOUBLE_ACCES As Integer = -11  REM you have two process that try to acces to the same device
    Public Const UNAUTHORIZED As Integer = -12  REM unauthorized access to password-protected device
    Public Const RTC_NOT_READY As Integer = -13 REM real-time clock has not been initialized (or time was lost)

    REM --- (end of generated code: YFunction return codes)

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




    Private Shared ReadOnly decexp() As Double = {
  0.000001, 0.00001, 0.0001, 0.001, 0.01, 0.1, 1.0,
  10.0, 100.0, 1000.0, 10000.0, 100000.0, 1000000.0, 10000000.0, 100000000.0, 1000000000.0}

    REM Convert Yoctopuce 16-bit decimal floats to standard double-precision floats
    REM
    Public Shared Function _decimalToDouble(ByVal val As Integer) As Double
      Dim negate As Boolean = False
      Dim res As Double
      If val = 0 Then
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
      res = (val And 2047) * decexp(exp)
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


    Public Shared Sub RegisterCalibrationHandler(ByVal calibType As Integer, ByVal callback As yCalibrationHandler)

      Dim key As String
      key = calibType.ToString()
      _CalibHandlers.Add(key, callback)
    End Sub

    Private Shared Function yLinearCalibrationHandler(ByVal rawValue As Double, ByVal calibType As Integer, ByVal parameters As List(Of Integer), ByVal rawValues As List(Of Double), ByVal refValues As List(Of Double)) As Double

      Dim npt, i As Integer
      Dim x, adj As Double
      Dim x2, adj2 As Double


      npt = calibType Mod 10
      x = rawValues(0)
      adj = refValues(0) - x
      i = 0

      If (npt > rawValues.Count) Then npt = rawValues.Count
      If (npt > refValues.Count) Then npt = refValues.Count + 1
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
    Public Shared Function yGetAPIVersion() As String
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

      If (apiInitialized) Then
        InitAPI = YAPI_SUCCESS
        Exit Function
      End If

      If (YOCTO_API_VERSION_BCD <> yapiGetAPIVersion(version, apidate)) Then
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
      _yapiRegisterFunctionUpdateCallback(Marshal.GetFunctionPointerForDelegate(native_yFunctionUpdateDelegate))
      _yapiRegisterTimedReportCallback(Marshal.GetFunctionPointerForDelegate(native_yTimedReportDelegate))
      _yapiRegisterLogFunction(Marshal.GetFunctionPointerForDelegate(native_yLogFunctionDelegate))
      _yapiRegisterDeviceLogCallback(Marshal.GetFunctionPointerForDelegate(native_yDeviceLogDelegate))
      _yapiRegisterHubDiscoveryCallback(Marshal.GetFunctionPointerForDelegate(native_HubDiscoveryDelegate))
      For i = 1 To 20
        RegisterCalibrationHandler(i, AddressOf yLinearCalibrationHandler)
      Next i

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
    '''   <c>http://username:password@adresse:port</c>
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
    '''   Fault-tolerant alternative to RegisterHub().
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
          res = _yapiSleep(1, errBuffer)
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
    '''   Force a hub discovery, if a callback as been registered with <c>yRegisterDeviceRemovalCallback</c> it
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
    '''   a procedure taking a string parameter, or <c>null</c>
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
    '''   a procedure taking a <c>YModule</c> parameter, or <c>null</c>
    '''   to unregister a previously registered  callback.
    ''' </param>
    '''/
    Public Shared Sub RegisterDeviceArrivalCallback(ByVal arrivalCallback As yDeviceUpdateFunc)
      yArrival = arrivalCallback
      If (arrivalCallback <> Nothing) Then
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
    '''   a procedure taking a <c>YModule</c> parameter, or <c>null</c>
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
    '''   a procedure taking two string parameter, or null
    '''   to unregister a previously registered  callback.
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





  <StructLayout(LayoutKind.Sequential, pack:=1, CharSet:=CharSet.Ansi)> _
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
  <StructLayout(LayoutKind.Sequential, pack:=1, CharSet:=CharSet.Ansi)> _
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

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Delegate Function yFlashCallback(ByVal stepnumber As yu32, ByVal totalStep As yu32, ByVal context As IntPtr) As Integer

  <StructLayout(LayoutKind.Sequential, pack:=1, CharSet:=CharSet.Ansi)> _
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
  Public Const YAPI_SUCCESS As Integer = 0         REM everything worked allright
  Public Const YAPI_NOT_INITIALIZED As Integer = -1 REM call yInitAPI() first !
  Public Const YAPI_INVALID_ARGUMENT As Integer = -2 REM one of the arguments passed to the function is invalid
  Public Const YAPI_NOT_SUPPORTED As Integer = -3  REM the operation attempted is (currently) not supported
  Public Const YAPI_DEVICE_NOT_FOUND As Integer = -4 REM the requested device is not reachable
  Public Const YAPI_VERSION_MISMATCH As Integer = -5 REM the device firmware is incompatible with this API version
  Public Const YAPI_DEVICE_BUSY As Integer = -6    REM the device is busy with another task and cannot answer
  Public Const YAPI_TIMEOUT As Integer = -7        REM the device took too long to provide an answer
  Public Const YAPI_IO_ERROR As Integer = -8       REM there was an I/O problem while talking to the device
  Public Const YAPI_NO_MORE_DATA As Integer = -9   REM there is no more data to read from
  Public Const YAPI_EXHAUSTED As Integer = -10     REM you have run out of a limited ressource, check the documentation
  Public Const YAPI_DOUBLE_ACCES As Integer = -11  REM you have two process that try to acces to the same device
  Public Const YAPI_UNAUTHORIZED As Integer = -12  REM unauthorized access to password-protected device
  Public Const YAPI_RTC_NOT_READY As Integer = -13 REM real-time clock has not been initialized (or time was lost)

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
  Public Const Y_USBBANDWIDTH_SIMPLE As Integer = 0
  Public Const Y_USBBANDWIDTH_DOUBLE As Integer = 1
  Public Const Y_USBBANDWIDTH_INVALID As Integer = -1

  Public Delegate Sub YModuleLogCallback(ByVal modul As YModule, ByVal logline As String)
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
  Public Const Y_CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_RESOLUTION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YSensorValueCallback(ByVal func As YSensor, ByVal value As String)
  Public Delegate Sub YSensorTimedReportCallback(ByVal func As YSensor, ByVal measure As YMeasure)
  REM --- (end of generated code: YSensor globals)

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
    End Sub ' New
  End Class

  Dim YDevice_devCache As List(Of YDevice)



  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub HTTPRequestCallback(ByVal device As YDevice, ByRef context As blockingCallbackCtx, ByVal returnval As YRETCODE, ByVal result As String, ByVal errmsg As String)

  REM - Types used for public yocto_api callbacks
  Public Delegate Sub yLogFunc(ByVal log As String)
  Public Delegate Sub yDeviceUpdateFunc(ByVal modul As YModule)
  Public Delegate Sub yFunctionUpdateFunc(ByVal modul As YModule, ByVal functionId As String, ByVal functionName As String, ByVal functionValue As String)
  Public Delegate Function yCalibrationHandler(ByVal rawValue As Double, ByVal calibType As Integer, ByVal parameters As List(Of Integer), ByVal rawValues As List(Of Double), ByVal refValues As List(Of Double)) As Double
  Public Delegate Sub YHubDiscoveryCallback(ByVal serial As String, ByVal url As String)


  REM - Types used for internal yapi callbacks
  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub _yapiLogFunc(ByVal log As IntPtr, ByVal loglen As yu32)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub _yapiDeviceUpdateFunc(ByVal dev As YDEV_DESCR)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub _yapiFunctionUpdateFunc(ByVal dev As YFUN_DESCR, ByVal value As IntPtr)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub _yapiTimedReportFunc(ByVal dev As YFUN_DESCR, ByVal timestamp As Double, ByVal data As IntPtr, ByVal len As System.UInt32)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub _yapiHubDiscoveryCallback(ByVal serial As IntPtr, ByVal url As IntPtr)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
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
      REM decode sequence header to extract data
      _runNo = encoded.ElementAt(0) + (encoded.ElementAt(1) << 16)
      _utcStamp = CUInt(encoded.ElementAt(2) + CUInt(encoded.ElementAt(3) << 16))
      _isAvg = (encoded.ElementAt(4) And &H100) = 0
      _samplesPerHour = encoded.ElementAt(4) And &HFF
      If ((encoded.ElementAt(4) And &H100) <> 0) Then
        _samplesPerHour *= 3600
      ElseIf ((encoded.ElementAt(4) And &H200) <> 0) Then
        _samplesPerHour *= 60
      End If
      _decimals = encoded.ElementAt(5)
      _offset = encoded.ElementAt(5)
      _scale = encoded.ElementAt(6)
      _isScal = (_scale <> 0)
      _isClosed = (encoded.ElementAt(7) <> &HFFFF)
      If (_isClosed) Then
        _nRows = encoded.ElementAt(7)
      Else
        _nRows = 0
      End If
      _duration = CInt((_nRows * 3600 + (_samplesPerHour / 2)) / _samplesPerHour)

      REM precompute decoding parameters
      _decexp = 1.0
      For i As Integer = 0 To _decimals - 1 Step 1
        _decexp *= 10.0
      Next i
      _columnNames = New List(Of String)()
      _calpar = New List(Of Integer)()
      _calraw = New List(Of Double)()
      _calref = New List(Of Double)()
      _values = New List(Of List(Of Double))()
      Dim calib As List(Of Integer) = dataset.get_calibration()
      _caltyp = calib.ElementAt(0)
      If (_caltyp > 0) Then
        _calhdl = YAPI._getCalibrationHandler(_caltyp)
        For i As Integer = 1 To calib.Count - 1 Step 1
          Dim ival As Integer = calib.ElementAt(i)
          Dim fval As Double
          If (_caltyp <= 10) Then
            fval = (ival - _offset) / _scale
          Else
            fval = YAPI._decimalToDouble(ival)
          End If
          _calpar.Add(ival)
          If ((i And 1) <> 0) Then
            _calraw.Add(fval)
          Else
            _calref.Add(fval)
          End If
        Next i
      End If

      REM preload column names for backward-compatibility
      _functionId = dataset.get_functionId()
      If (_isAvg) Then
        _columnNames.Add(_functionId + "_min")
        _columnNames.Add(_functionId + "_avg")
        _columnNames.Add(_functionId + "_max")
        _nCols = 3
      Else
        _columnNames.Add(_functionId)
        _nCols = 1
      End If

      REM decode min/avg/max values for the sequence
      If (_nRows > 0) Then
        _minVal = _decodeVal(encoded.ElementAt(8))
        _maxVal = _decodeVal(encoded.ElementAt(9))
        _avgVal = _decodeAvg(encoded.ElementAt(10) + (encoded.ElementAt(11) << 16), _nRows)
      End If
    End Sub

    REM --- (generated code: YDataStream private methods declaration)

    REM --- (end of generated code: YDataStream private methods declaration)

    REM --- (generated code: YDataStream public methods declaration)
    Public Overridable Function _initFromDataSet(dataset As YDataSet, encoded As List(Of Integer)) As Integer
      Dim val As Integer = 0
      Dim i As Integer = 0
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
      iCalib = dataset.get_calibration()
      Me._caltyp = iCalib(0)
      If (Me._caltyp <> 0) Then
        Me._calhdl = YAPI._getCalibrationHandler(Me._caltyp)
        Me._calpar.Clear()
        Me._calraw.Clear()
        Me._calref.Clear()
        i = 1
        While (i + 1 < iCalib.Count)
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
        Me._minVal = Me._decodeVal(encoded(8))
        Me._maxVal = Me._decodeVal(encoded(9))
        Me._avgVal = Me._decodeAvg(encoded(10) + (((encoded(11)) << (16))), Me._nRows)
      End If
      Return 0
    End Function

    Public Overridable Function parse(sdata As Byte()) As Integer
      Dim idx As Integer = 0
      Dim udat As List(Of Integer) = New List(Of Integer)()
      Dim dat As List(Of Double) = New List(Of Double)()
      REM // may throw an exception
      udat = YAPI._decodeWords(Me._parent._json_get_string(sdata))
      Me._values.Clear()
      idx = 0
      If (Me._isAvg) Then
        While (idx + 3 < udat.Count)
          dat.Clear()
          dat.Add(Me._decodeVal(udat(idx)))
          dat.Add(Me._decodeAvg(udat(idx + 2) + (((udat(idx + 3)) << (16))), 1))
          dat.Add(Me._decodeVal(udat(idx + 1)))
          Me._values.Add(New List(Of Double)(dat))
          idx = idx + 4
        End While
      Else
        If (Me._isScal) Then
          While (idx < udat.Count)
            dat.Clear()
            dat.Add(Me._decodeVal(udat(idx)))
            Me._values.Add(New List(Of Double)(dat))
            idx = idx + 1
          End While
        Else
          While (idx + 1 < udat.Count)
            dat.Clear()
            dat.Add(Me._decodeAvg(udat(idx) + (((udat(idx + 1)) << (16))), 1))
            Me._values.Add(New List(Of Double)(dat))
            idx = idx + 2
          End While
        End If
      End If
      
      Me._nRows = Me._values.Count
      Return YAPI.SUCCESS
    End Function

    Public Overridable Function get_url() As String
      Dim url As String
      url = "logger.json?id=" +
      Me._functionId + "&run=" + Convert.ToString(Me._runNo) + "&utc=" + Convert.ToString(Me._utcStamp)
      Return url
    End Function

    Public Overridable Function loadStream() As Integer
      REM // may throw an exception
      Return Me.parse(Me._parent._download(Me.get_url()))
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
        val = Me._calhdl(val, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
      End If
      Return val
    End Function

    Public Overridable Function _decodeAvg(dw As Integer, count As Integer) As Double
      Dim val As Double = 0
      val = dw
      If (Me._isScal) Then
        val = (val / (100 * count) - Me._offset) / Me._scale
      Else
        val = val / (count * Me._decexp)
      End If
      If (Me._caltyp <> 0) Then
        val = Me._calhdl(val, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
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
      If ((Me._nRows <> 0) And Me._isClosed) Then
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
      If ((Me._values.Count = 0) Or Not (Me._isClosed)) Then
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
      If ((Me._values.Count = 0) Or Not (Me._isClosed)) Then
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

  REM --- (generated code: DataStream functions)


  REM --- (end of generated code: DataStream functions)



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
    '''   When the recording rate is higher then 1 sample
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

  REM --- (generated code: Measure functions)


  REM --- (end of generated code: Measure functions)


  REM --- (generated code: YDataSet class start)

  '''*
  ''' <summary>
  '''   YDataSet objects make it possible to retrieve a set of recorded measures
  '''   for a given sensor and a specified time interval.
  ''' <para>
  '''   They can be used
  '''   to load data points with a progress report. When the YDataSet object is
  '''   instanciated by the <c>get_recordedData()</c>  function, no data is
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


    Sub New(parent As YFunction, data As String)
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
      _parse(data)
    End Sub


    Protected Function _parse(data As String) As Integer
      Dim p As TJsonParser
      Dim node As TJSONRECORD
      Dim arr As TJSONRECORD
      Dim stream As YDataStream
      Dim summaryMinVal As Double = Double.MaxValue
      Dim summaryMaxVal As Double = -Double.MaxValue
      Dim summaryTotalTime As Double = 0
      Dim summaryTotalAvg As Double = 0
      Dim startTime As UInt32
      Dim endtime As UInt32

      If Not (YAPI.ExceptionsDisabled) Then
        p = New TJsonParser(data, False)
      Else
        Try
          p = New TJsonParser(data, False)
        Catch E As Exception
          Return YAPI_NOT_SUPPORTED
          Exit Function
        End Try
      End If


      node = CType(p.GetChildNode(Nothing, "id"), TJSONRECORD)
      _functionId = node.svalue
      node = CType(p.GetChildNode(Nothing, "unit"), TJSONRECORD)
      _unit = node.svalue
      node = CType(p.GetChildNode(Nothing, "cal"), TJSONRECORD)
      _calib = YAPI._decodeWords(node.svalue)
      arr = CType(p.GetChildNode(Nothing, "streams"), TJSONRECORD)
      _streams = New List(Of YDataStream)()
      _preview = New List(Of YMeasure)()
      _measures = New List(Of YMeasure)()

      For i As Integer = 0 To arr.itemcount - 1 Step 1
        stream = _parent._findDataStream(Me, arr.items.ElementAt(i).svalue)
        If (_startTime > 0 And stream.get_startTimeUTC() + stream.get_duration() <= _startTime) Then
          REM this stream is too early, drop it
        ElseIf (_endTime > 0 And stream.get_startTimeUTC() > _endTime) Then
          REM this stream is too late, drop it
        Else
          _streams.Add(stream)
          If (stream.isClosed() And stream.get_startTimeUTC() >= _startTime And (_endTime = 0 Or stream.get_startTimeUTC() + stream.get_duration() <= _endTime)) Then
            If (summaryMinVal > stream.get_minValue()) Then
              summaryMinVal = stream.get_minValue()
            End If
            If (summaryMaxVal < stream.get_maxValue()) Then
              summaryMaxVal = stream.get_maxValue()
            End If
            summaryTotalAvg += stream.get_averageValue() * stream.get_duration()
            summaryTotalTime += stream.get_duration()
            Dim rec As New YMeasure(stream.get_startTimeUTC(),
                                    stream.get_startTimeUTC() + stream.get_duration(),
                                    stream.get_minValue(),
                                    stream.get_averageValue(),
                                    stream.get_maxValue())
            _preview.Add(rec)
          End If
        End If
      Next i
      If (_streams.Count > 0) And (summaryTotalTime > 0) Then
        REM update time boundaries with actual data
        stream = _streams.ElementAt(_streams.Count - 1)
        endtime = CUInt(stream.get_startTimeUTC() + stream.get_duration())
        startTime = CUInt((_streams.ElementAt(0).get_startTimeUTC() - stream.get_dataSamplesIntervalMs() / 1000))
        If (_startTime < startTime) Then
          _startTime = startTime
        End If
        If (_endTime = 0 Or _endTime > endtime) Then
          _endTime = endtime
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
    Public Overridable Function get_calibration() As List(Of Integer)
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
      REM // may throw an exception
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
      stream.parse(data)
      dataRows = stream.get_dataRows()
      Me._progress = Me._progress + 1
      If (dataRows.Count = 0) Then
        Return Me.get_progress()
      End If
      tim = CType(stream.get_startTimeUTC(), Double)
      itv = stream.get_dataSamplesInterval()
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
        If ((tim >= Me._startTime) And ((Me._endTime = 0) Or (tim <= Me._endTime))) Then
          Me._measures.Add(New YMeasure(tim - itv, tim, dataRows(i_i)(minCol), dataRows(i_i)(avgCol), dataRows(i_i)(maxCol)))
          tim = tim + itv
        End If
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
    '''   When the object is instanciated by <c>get_dataSet</c>,
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
      Else
        If (Me._progress >= Me._streams.Count) Then
          Return 100
        Else
          stream = Me._streams(Me._progress)
          url = stream.get_url()
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

  REM --- (generated code: DataSet functions)


  REM --- (end of generated code: DataSet functions)



  Public Class YDevice
    Private _devdescr As YDEV_DESCR
    Private _cacheStamp As yu64
    Private _cacheJson As TJsonParser
    Private _functions As New List(Of yu32)
    Private _http_result As String
    Private _rootdevice As String
    Private _subpath As String
    Private _subpathinit As Boolean

    Public Sub New(ByVal devdesc As YDEV_DESCR)
      _devdescr = devdesc
      _cacheStamp = 0
      _cacheJson = Nothing
    End Sub

    Public Sub dispose()

      If _cacheJson IsNot Nothing Then _cacheJson.Dispose()
      _cacheJson = Nothing

    End Sub

    Public Shared Function getDevice(ByVal devdescr As YDEV_DESCR) As YDevice
      Dim idx As Integer
      Dim dev As YDevice
      For idx = 0 To YDevice_devCache.Count - 1
        If YDevice_devCache(idx)._devdescr = devdescr Then
          getDevice = YDevice_devCache(idx)
          Exit Function
        End If
      Next
      dev = New YDevice(devdescr)
      YDevice_devCache.Add(dev)
      getDevice = dev
    End Function

    Public Shared Function HTTPRequestSync(ByVal device As String, ByVal request As String, ByRef reply As String, ByRef errmsg As String) As YRETCODE
      Dim binreply As Byte()
      Dim res As YRETCODE

      ReDim binreply(-1)
      res = HTTPRequestSync(device, YAPI.DefaultEncoding.GetBytes(request), binreply, errmsg)
      reply = YAPI.DefaultEncoding.GetString(binreply)
      Return res
    End Function

    Public Shared Function HTTPRequestSync(ByVal device As String, ByVal request As Byte(), ByRef reply As Byte(), ByRef errmsg As String) As YRETCODE
      Dim iohdl As YIOHDL
      Dim requestbuf As IntPtr = IntPtr.Zero
      Dim buffer As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim preply As IntPtr = IntPtr.Zero
      Dim replysize As Integer = 0
      Dim res As YRETCODE

      iohdl.raw = 0  REM dummy, useless init to avoid compiler warning

      requestbuf = Marshal.AllocHGlobal(request.Length)
      Marshal.Copy(request, 0, requestbuf, request.Length)
      res = _yapiHTTPRequestSyncStartEx(iohdl, New StringBuilder(device), requestbuf, request.Length, preply, replysize, buffer)
      If (res < 0) Then
        errmsg = buffer.ToString()
        Return res
      End If

      ReDim reply(replysize - 1)
      Marshal.Copy(preply, reply, 0, replysize)
      res = _yapiHTTPRequestSyncDone(iohdl, buffer)
      errmsg = buffer.ToString()
      Return res
    End Function

    Public Function HTTPRequestAsync(ByVal request As String, ByRef errmsg As String) As YRETCODE
      Return HTTPRequestAsync(YAPI.DefaultEncoding.GetBytes(request), errmsg)
    End Function

    Public Function HTTPRequestAsync(ByVal request As Byte(), ByRef errmsg As String) As YRETCODE
      Dim fullrequest As Byte()
      Dim requestbuf As IntPtr = IntPtr.Zero
      Dim buffer As StringBuilder = New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim res As YRETCODE

      ReDim fullrequest(-1)
      res = HTTPRequestPrepare(request, fullrequest, errmsg)
      requestbuf = Marshal.AllocHGlobal(fullrequest.Length)
      Marshal.Copy(fullrequest, 0, requestbuf, fullrequest.Length)
      res = _yapiHTTPRequestAsyncEx(New StringBuilder(_rootdevice), requestbuf, fullrequest.Length, IntPtr.Zero, IntPtr.Zero, buffer)
      Marshal.FreeHGlobal(requestbuf)
      errmsg = buffer.ToString()
      Return res
    End Function

    Public Function HTTPRequestPrepare(ByVal request As Byte(), ByRef fullrequest As Byte(), ByRef errmsg As String) As YRETCODE
      Dim res As YRETCODE
      Dim errbuf As New StringBuilder(YOCTO_ERRMSG_LEN)
      Dim b As StringBuilder
      Dim neededsize As Integer
      Dim p As Integer
      Dim root As New StringBuilder(YOCTO_SERIAL_LEN)
      Dim tmp As Integer

      _cacheStamp = CULng(YAPI.GetTickCount())  REM invalidate cache

      If (Not (_subpathinit)) Then
        res = _yapiGetDevicePath(_devdescr, root, Nothing, 0, neededsize, errbuf)
        If (YISERR(res)) Then
          errmsg = errbuf.ToString()
          Return res
        End If

        b = New StringBuilder(neededsize)
        res = _yapiGetDevicePath(_devdescr, root, b, neededsize, tmp, errbuf)
        If (YISERR(res)) Then
          errmsg = errbuf.ToString()
          Return res
        End If

        _rootdevice = root.ToString()
        _subpath = b.ToString()
        _subpathinit = True
      End If

      REM Search for the first '/'
      p = 0
      While p < request.Length And request(p) <> 47
        p += 1
      End While
      ReDim fullrequest(request.Length - 1 + _subpath.Length - 1)
      Buffer.BlockCopy(request, 0, fullrequest, 0, p)
      Buffer.BlockCopy(System.Text.Encoding.ASCII.GetBytes(_subpath), 0, fullrequest, p, _subpath.Length)
      Buffer.BlockCopy(request, p + 1, fullrequest, p + _subpath.Length, request.Length - p - 1)

      Return YAPI_SUCCESS
    End Function

    Public Function HTTPRequest(ByVal request As String, ByRef buffer As String, ByRef errmsg As String) As YRETCODE
      Dim binreply As Byte()
      Dim res As YRETCODE

      ReDim binreply(-1)
      res = HTTPRequest(YAPI.DefaultEncoding.GetBytes(request), binreply, errmsg)
      buffer = YAPI.DefaultEncoding.GetString(binreply)

      Return res
    End Function

    Public Function HTTPRequest(ByVal request As String, ByRef buffer As Byte(), ByRef errmsg As String) As YRETCODE
      Return HTTPRequest(YAPI.DefaultEncoding.GetBytes(request), buffer, errmsg)
    End Function

    Public Function HTTPRequest(ByVal request As Byte(), ByRef buffer As Byte(), ByRef errmsg As String) As YRETCODE
      Dim fullrequest As Byte() = Nothing

      Dim res As Integer = HTTPRequestPrepare(request, fullrequest, errmsg)
      If (YISERR(res)) Then
        Return res
      End If

      HTTPRequest = HTTPRequestSync(_rootdevice, fullrequest, buffer, errmsg)
    End Function


    Function requestAPI(ByRef apires As TJsonParser, ByRef errmsg As String) As YRETCODE
      Dim buffer As String = ""
      Dim res As Integer
      apires = Nothing

      REM Check if we have a valid cache value
      If (_cacheStamp > YAPI.GetTickCount()) Then
        apires = _cacheJson
        requestAPI = YAPI_SUCCESS
        Exit Function
      End If

      res = HTTPRequest("GET /api.json " + Chr(13) + Chr(10) + Chr(13) + Chr(10), buffer, errmsg)

      If (YISERR(res)) Then

        REM make sure a device scan does not solve the issue
        res = yapiUpdateDeviceList(1, errmsg)
        If (YISERR(res)) Then
          requestAPI = res
          Exit Function
        End If
        res = HTTPRequest("GET /api.json " + Chr(13) + Chr(10) + Chr(13) + Chr(10), buffer, errmsg)
        If (YISERR(res)) Then
          requestAPI = res
          Exit Function
        End If
      End If

      Try
        apires = New TJsonParser(buffer)
      Catch E As Exception
        errmsg = "unexpected JSON structure: " + E.Message
        requestAPI = YAPI_IO_ERROR
        Exit Function
      End Try

      REM store result in cache
      _cacheJson = apires
      _cacheStamp = CULng(YAPI.GetTickCount() + YAPI.DefaultCacheValidity)

      requestAPI = YAPI_SUCCESS
    End Function

    Public Function getFunctions(ByRef functions As List(Of yu32), ByRef errmsg As String) As YRETCODE
      Dim res, neededsize, i, count As Integer
      Dim p As IntPtr
      Dim ids() As ys32
      If (_functions.Count = 0) Then
        res = yapiGetFunctionsByDevice(_devdescr, 0, Nothing, 64, neededsize, errmsg)
        If (YISERR(res)) Then
          getFunctions = res
          Exit Function
        End If

        p = Marshal.AllocHGlobal(neededsize)

        res = yapiGetFunctionsByDevice(_devdescr, 0, p, 64, neededsize, errmsg)
        If (YISERR(res)) Then
          Marshal.FreeHGlobal(p)
          getFunctions = res
          Exit Function
        End If

        count = CInt(neededsize / Marshal.SizeOf(i))  REM  i is an 32 bits integer 
        ReDim Preserve ids(count)
        Marshal.Copy(p, ids, 0, count)
        For i = 0 To count - 1
          _functions.Add(CUInt(ids(i)))
        Next i

        Marshal.FreeHGlobal(p)
      End If
      functions = _functions
      getFunctions = YAPI_SUCCESS
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
    Protected _cacheExpiration As ULong
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

    Private Function _buildSetRequest(ByVal changeattr As String, ByVal changeval As String, ByRef request As String, ByRef errmsg As String) As YRETCODE

      Dim res, i As Integer
      Dim fundesc As YFUN_DESCR
      Dim funcid As New StringBuilder(YOCTO_FUNCTION_LEN)
      Dim errbuff As New StringBuilder(YOCTO_ERRMSG_LEN)

      Dim uchangeval, h As String
      Dim c As Char
      Dim devdesc As YDEV_DESCR

      funcid.Length = 0
      errbuff.Length = 0


      REM Resolve the function name
      res = _getDescriptor(fundesc, errmsg)

      If (YISERR(res)) Then
        _buildSetRequest = res
        Exit Function
      End If

      res = _yapiGetFunctionInfo(fundesc, devdesc, Nothing, funcid, Nothing, Nothing, errbuff)
      If YISERR(res) Then
        errmsg = errbuff.ToString()
        _throw(res, errmsg)
        _buildSetRequest = res
        Exit Function
      End If


      request = "GET /api/" + funcid.ToString() + "/"
      uchangeval = ""

      If (changeattr <> "") Then
        request = request + changeattr + "?" + changeattr + "="
        For i = 0 To changeval.Length - 1
          c = changeval.Chars(i)
          If (c < " ") Or ((c > Chr(122)) And (c <> "~")) Or (c = Chr(34)) Or (c = "%") Or (c = "&") Or
             (c = "+") Or (c = "<") Or (c = "=") Or (c = ">") Or (c = "\") Or (c = "^") Or (c = "`") Then
            h = Hex(Asc(c))
            If (h.Length < 2) Then h = "0" + h
            uchangeval = uchangeval + "%" + h
          Else
            uchangeval = uchangeval + c
          End If
        Next
      End If

      request = request + uchangeval + " " + Chr(13) + Chr(10) + Chr(13) + Chr(10)
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
      _cacheExpiration = 0
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

      bodystr = "Content-Disposition: form-data; name=""" + path + """; filename=""api""" + Chr(13) + Chr(10) + _
                "Content-Type: application/octet-stream" + Chr(13) + Chr(10) + _
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

      header = YAPI.DefaultEncoding.GetBytes("POST /upload.html HTTP/1.1" + Chr(13) + Chr(10) + _
                                             "Content-Type: multipart/form-data, boundary=" + boundary + Chr(13) + Chr(10) + _
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

    Protected Function _parse(ByRef j As TJSONRECORD) As Integer
      Dim member As TJSONRECORD
      Dim i As Integer
      If (j.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then
        Return -1
      End If
      For i = 0 To j.membercount - 1
        member = j.members(i)
        _parseAttr(member)
      Next i
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

    Protected Overridable Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "logicalName") Then
        _logicalName = member.svalue
        Return 1
      End If
      If (member.name = "advertisedValue") Then
        _advertisedValue = member.svalue
        Return 1
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LOGICALNAME_INVALID
        End If
      End If
      Return Me._logicalName
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
      rest_val = newval
      Return _setAttr("logicalName", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current value of the function (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the function (no more than 6 characters)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ADVERTISEDVALUE_INVALID
        End If
      End If
      Return Me._advertisedValue
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
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
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
      If (Not (callback Is Nothing) And Me.isOnline()) Then
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

    Public Overridable Function _parserHelper() As Integer
      REM // By default, nothing to do
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
    '''   number and of the hardware identifier of the function. (for example <c>RELAYLO1-123456.relay1</c>)
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
    '''   a number corresponding to the code of the latest error that occured while
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
      Dim apires As TJsonParser = Nothing

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
      Dim node As Nullable(Of TJSONRECORD)

      Dim st As String = YAPI.DefaultEncoding.GetString(data)
      Dim p As TJsonParser

      If Not (YAPI.ExceptionsDisabled) Then
        p = New TJsonParser(st)
      Else
        Try
          p = New TJsonParser(st)
        Catch E As Exception
          Return ""
          Exit Function
        End Try
      End If

      node = p.GetChildNode(Nothing, key)

      Return node.Value.svalue
    End Function

    Protected Function _json_get_array(ByVal data As Byte()) As List(Of String)
      Dim st As String = YAPI.DefaultEncoding.GetString(data)
      Dim p As TJsonParser

      If Not (YAPI.ExceptionsDisabled) Then
        p = New TJsonParser(st)
      Else
        Try
          p = New TJsonParser(st)
        Catch E As Exception
          Return Nothing
          Exit Function
        End Try
      End If


      Return p.GetAllChilds(Nothing)
    End Function

    Public Function _json_get_string(ByVal data As Byte()) As String
      Dim json_str As String = YAPI.DefaultEncoding.GetString(data)
      Dim p As TJsonParser = New TJsonParser("[" + json_str + "]", False)
      Dim node As TJSONRECORD = p.GetRootNode()
      Return node.items.ElementAt(0).svalue
    End Function




    '''*
    ''' <summary>
    '''   Preloads the function cache with a specified validity duration.
    ''' <para>
    '''   By default, whenever accessing a device, all function attributes
    '''   are kept in cache for the standard duration (5 ms). This method can be
    '''   used to temporarily mark the cache as valid for a longer period, in order
    '''   to reduce network trafic for instance.
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
      Dim apires As TJsonParser = Nothing
      Dim fundescr As YFUN_DESCR
      Dim res As Integer
      Dim errbuf As String = ""
      Dim funcId As String = ""
      Dim devdesc As YDEV_DESCR
      Dim serial As String = ""
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim node As Nullable(Of TJSONRECORD)

      REM Resolve our reference to our device, load REST API
      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        load = res
        Exit Function
      End If

      res = dev.requestAPI(apires, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        load = res
        Exit Function
      End If

      REM Get our function Id
      fundescr = yapiGetFunction(_className, _func, errmsg)
      If (YISERR(fundescr)) Then
        _throw(res, errmsg)
        load = res
        Exit Function
      End If

      devdesc = 0
      res = yapiGetFunctionInfo(fundescr, devdesc, serial, funcId, funcName, funcVal, errbuf)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        load = res
        Exit Function
      End If
      _cacheExpiration = CULng(YAPI.GetTickCount() + msValidity)
      _serial = serial
      _funId = funcId
      _hwId = _serial + "." + _funId

      node = apires.GetChildNode(Nothing, funcId)
      If Not (node.HasValue) Then
        _throw(YAPI_IO_ERROR, "unexpected JSON structure: missing function " + funcId)
        load = res
        Exit Function
      End If

      _parse(CType(node, TJSONRECORD))
      load = YAPI_SUCCESS
    End Function

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

  REM --- (generated code: Function functions)

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


  REM --- (end of generated code: Function functions)



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
    Public Const USBBANDWIDTH_SIMPLE As Integer = 0
    Public Const USBBANDWIDTH_DOUBLE As Integer = 1
    Public Const USBBANDWIDTH_INVALID As Integer = -1

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
    Protected _usbBandwidth As Integer
    Protected _valueCallbackModule As YModuleValueCallback
    Protected _logCallback As YModuleLogCallback
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
      _usbBandwidth = USBBANDWIDTH_INVALID
      _valueCallbackModule = Nothing
      _logCallback = Nothing
      REM --- (end of generated code: YModule attributes initialization)
    End Sub

    REM --- (generated code: YModule private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "productName") Then
        _productName = member.svalue
        Return 1
      End If
      If (member.name = "serialNumber") Then
        _serialNumber = member.svalue
        Return 1
      End If
      If (member.name = "productId") Then
        _productId = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "productRelease") Then
        _productRelease = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "firmwareRelease") Then
        _firmwareRelease = member.svalue
        Return 1
      End If
      If (member.name = "persistentSettings") Then
        _persistentSettings = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "luminosity") Then
        _luminosity = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "beacon") Then
        If (member.ivalue > 0) Then _beacon = 1 Else _beacon = 0
        Return 1
      End If
      If (member.name = "upTime") Then
        _upTime = member.ivalue
        Return 1
      End If
      If (member.name = "usbCurrent") Then
        _usbCurrent = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "rebootCountdown") Then
        _rebootCountdown = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "usbBandwidth") Then
        _usbBandwidth = CInt(member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
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
    Private Function _getFunction(ByVal idx As Integer, ByRef serial As String, ByRef funcId As String, _
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

      res = yapiGetFunctionInfo(fundescr, devdescr, serial, funcId, funcName, funcVal, errmsg)
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
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res As Integer
      res = _getFunction(functionIndex, serial, funcId, funcName, funcVal, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        functionId = YAPI.INVALID_STRING
        Exit Function
      End If
      functionId = funcId
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
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res As Integer

      res = _getFunction(functionIndex, serial, funcId, funcName, funcVal, errmsg)
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
      Dim funcName As String = ""
      Dim funcVal As String = ""
      Dim errmsg As String = ""
      Dim res As Integer

      res = _getFunction(functionIndex, serial, funcId, funcName, funcVal, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        functionValue = YAPI.INVALID_STRING
        Exit Function
      End If
      functionValue = funcVal
    End Function


    '''*
    ''' <summary>
    '''   todo
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Function registerLogCallback(ByVal callback As YModuleLogCallback) As Integer
      _logCallback = callback
      If _logCallback Is Nothing Then
        _yapiStartStopDeviceLogCallback(New StringBuilder(_serial), 0)
      Else
        _yapiStartStopDeviceLogCallback(New StringBuilder(_serial), 1)
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
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PRODUCTNAME_INVALID
        End If
      End If
      Return Me._productName
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
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SERIALNUMBER_INVALID
        End If
      End If
      Return Me._serialNumber
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
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PRODUCTID_INVALID
        End If
      End If
      Return Me._productId
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
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PRODUCTRELEASE_INVALID
        End If
      End If
      Return Me._productRelease
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return FIRMWARERELEASE_INVALID
        End If
      End If
      Return Me._firmwareRelease
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PERSISTENTSETTINGS_INVALID
        End If
      End If
      Return Me._persistentSettings
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LUMINOSITY_INVALID
        End If
      End If
      Return Me._luminosity
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BEACON_INVALID
        End If
      End If
      Return Me._beacon
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return UPTIME_INVALID
        End If
      End If
      Return Me._upTime
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return USBCURRENT_INVALID
        End If
      End If
      Return Me._usbCurrent
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return REBOOTCOUNTDOWN_INVALID
        End If
      End If
      Return Me._rebootCountdown
    End Function


    Public Function set_rebootCountdown(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("rebootCountdown", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of USB interfaces used by the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_USBBANDWIDTH_SIMPLE</c> or <c>Y_USBBANDWIDTH_DOUBLE</c>, according to the number of USB
    '''   interfaces used by the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_USBBANDWIDTH_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_usbBandwidth() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return USBBANDWIDTH_INVALID
        End If
      End If
      Return Me._usbBandwidth
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of USB interfaces used by the module.
    ''' <para>
    '''   You must reboot the module
    '''   after changing this setting.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_USBBANDWIDTH_SIMPLE</c> or <c>Y_USBBANDWIDTH_DOUBLE</c>, according to the number of USB
    '''   interfaces used by the module
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
    Public Function set_usbBandwidth(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("usbBandwidth", rest_val)
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
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
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
      If (Not (callback Is Nothing) And Me.isOnline()) Then
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
    '''   On failure, throws an exception or returns an empty content.
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
    ''' </returns>
    '''/
    Public Overridable Function get_icon2d() As Byte()
      REM // may throw an exception
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
    ''' </returns>
    '''/
    Public Overridable Function get_lastLogs() As String
      Dim content As Byte()
      REM // may throw an exception
      content = Me._download("logs.txt")
      Return YAPI.DefaultEncoding.GetString(content)
    End Function


    '''*
    ''' <summary>
    '''   Continues the module enumeration started using <c>yFirstModule()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YModule</c> object, corresponding to
    '''   the next module found, or a <c>null</c> pointer
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
    '''   the first module currently online, or a <c>null</c> pointer
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


  REM --- (generated code: Module functions)

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
  '''   the first module currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstModule() As YModule
    Return YModule.FirstModule()
  End Function


  REM --- (end of generated code: Module functions)




  REM --- (generated code: YSensor class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to read an instant
  '''   measure of the sensor, as well as the minimal and maximal values observed.
  ''' <para>
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
    Public Const CALIBRATIONPARAM_INVALID As String = YAPI.INVALID_STRING
    Public Const RESOLUTION_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of generated code: YSensor definitions)

    REM --- (generated code: YSensor attributes declaration)
    Protected _unit As String
    Protected _currentValue As Double
    Protected _lowestValue As Double
    Protected _highestValue As Double
    Protected _currentRawValue As Double
    Protected _logFrequency As String
    Protected _reportFrequency As String
    Protected _calibrationParam As String
    Protected _resolution As Double
    Protected _valueCallbackSensor As YSensorValueCallback
    Protected _timedReportCallbackSensor As YSensorTimedReportCallback
    Protected _prevTimedReport As Double
    Protected _iresol As Double
    Protected _offset As Double
    Protected _scale As Double
    Protected _decexp As Double
    Protected _isScal As Boolean
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
      _calibrationParam = CALIBRATIONPARAM_INVALID
      _resolution = RESOLUTION_INVALID
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

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "unit") Then
        _unit = member.svalue
        Return 1
      End If
      If (member.name = "currentValue") Then
        _currentValue = member.ivalue / 65536.0
        Return 1
      End If
      If (member.name = "lowestValue") Then
        _lowestValue = member.ivalue / 65536.0
        Return 1
      End If
      If (member.name = "highestValue") Then
        _highestValue = member.ivalue / 65536.0
        Return 1
      End If
      If (member.name = "currentRawValue") Then
        _currentRawValue = member.ivalue / 65536.0
        Return 1
      End If
      If (member.name = "logFrequency") Then
        _logFrequency = member.svalue
        Return 1
      End If
      If (member.name = "reportFrequency") Then
        _reportFrequency = member.svalue
        Return 1
      End If
      If (member.name = "calibrationParam") Then
        _calibrationParam = member.svalue
        Return 1
      End If
      If (member.name = "resolution") Then
        If (member.ivalue > 100) Then _resolution = 1.0 / Math.Round(65536.0 / member.ivalue) Else _resolution = 0.001 / Math.Round(67.0 / member.ivalue)
        Return 1
      End If
      Return MyBase._parseAttr(member)
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return UNIT_INVALID
        End If
      End If
      Return Me._unit
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of the measure.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current value of the measure
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
      Return Math.Round(res) / Me._iresol
    End Function


    '''*
    ''' <summary>
    '''   Changes the recorded minimal value observed.
    ''' <para>
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
      Return Math.Round(res) / Me._iresol
    End Function


    '''*
    ''' <summary>
    '''   Changes the recorded maximal value observed.
    ''' <para>
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
      Return Math.Round(res) / Me._iresol
    End Function

    '''*
    ''' <summary>
    '''   Returns the uncalibrated, unrounded raw value returned by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the uncalibrated, unrounded raw value returned by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENTRAWVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentRawValue() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CURRENTRAWVALUE_INVALID
        End If
      End If
      Return Me._currentRawValue
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LOGFREQUENCY_INVALID
        End If
      End If
      Return Me._logFrequency
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return REPORTFREQUENCY_INVALID
        End If
      End If
      Return Me._reportFrequency
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
    Public Function get_calibrationParam() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALIBRATIONPARAM_INVALID
        End If
      End If
      Return Me._calibrationParam
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RESOLUTION_INVALID
        End If
      End If
      Return Me._resolution
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
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
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
      If (Not (callback Is Nothing) And Me.isOnline()) Then
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
      REM // Store inverted resolution, to provide better rounding
      If (Me._resolution > 0) Then
        Me._iresol = Math.Round(1.0 / Me._resolution)
      Else
        Return 0
      End If
      
      Me._scale = -1
      Me._calpar.Clear()
      Me._calraw.Clear()
      Me._calref.Clear()
      
      REM // Old format: supported when there is no calibration
      If (Me._calibrationParam = "" Or Me._calibrationParam = "0") Then
        Me._caltyp = 0
        Return 0
      End If
      REM // Old format: calibrated value will be provided by the device
      If (Me._calibrationParam.IndexOf(",") >= 0) Then
        Me._caltyp = -1
        Return 0
      End If
      REM // New format, decode parameters
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
      position = 3
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
      
      
      
      Return 0
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
      REM // may throw an exception
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
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overridable Function registerTimedReportCallback(callback As YSensorTimedReportCallback) As Integer
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(Me, True)
      Else
        YFunction._UpdateTimedReportCallbackList(Me, False)
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
      REM // may throw an exception
      rest_val = Me._encodeCalibrationPoints(rawValues, refValues)
      Return Me._setAttr("calibrationParam", rest_val)
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
        Me._throw(YAPI.NOT_SUPPORTED, "Device does not support new calibration parameters. Please upgrade your firmware")
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
      If ((Me._caltyp < 0) Or (Me._scale < 0)) Then
        Me._throw(YAPI.NOT_SUPPORTED, "Device does not support new calibration parameters. Please upgrade your firmware")
        Return "0"
      End If
      If (Me._isScal) Then
        REM
        res = "" + Convert.ToString(npt)
        idx = 0
        While (idx < npt)
          iRaw = CType(Math.Round(rawValues(idx) * Me._scale - Me._offset), Integer)
          iRef = CType(Math.Round(refValues(idx) * Me._scale - Me._offset), Integer)
          res = "" +  res + "," + Convert.ToString( iRaw) + "," + Convert.ToString(iRef)
          idx = idx + 1
        End While
      Else
        REM
        res = "" + Convert.ToString(10 + npt)
        idx = 0
        While (idx < npt)
          iRaw = CType(YAPI._doubleToDecimal(rawValues(idx)), Integer)
          iRef = CType(YAPI._doubleToDecimal(refValues(idx)), Integer)
          res = "" +  res + "," + Convert.ToString( iRaw) + "," + Convert.ToString(iRef)
          idx = idx + 1
        End While
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
      If (report(0) > 0) Then
        REM
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
      Else
        REM
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
        val = Me._calhdl(val, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
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
        val = Me._calhdl(val, Me._caltyp, Me._calpar, Me._calraw, Me._calref)
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
    '''   a sensor currently online, or a <c>null</c> pointer
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
    '''   the first sensor currently online, or a <c>null</c> pointer
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

  REM --- (generated code: Sensor functions)

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
  '''   the first sensor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSensor() As YSensor
    Return YSensor.FirstSensor()
  End Function


  REM --- (end of generated code: Sensor functions)





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
    If ylog <> Nothing Then ylog(Marshal.PtrToStringAnsi(log))
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
  '''   a procedure taking a string parameter, or <c>null</c>
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
          If yArrival <> Nothing Then
            yArrival(_module)
          End If
        Case EVTYPE.REMOVAL
          If yRemoval <> Nothing Then
            yRemoval(_module)
          End If
        Case EVTYPE.CHANGE
          If (yChange <> Nothing) Then
            yChange(_module)
          End If
        Case EVTYPE.HUB_DISCOVERY
          If (_HubDiscoveryCallback <> Nothing) Then
            _HubDiscoveryCallback(_serial, _url)
          End If
      End Select
    End Sub
  End Class

  Private Class DataEvent
    Private _fun As YFunction
    Private _value As String
    Private _report As List(Of Integer)
    Private _timestamp As Double

    Public Sub New(ByVal fun As YFunction, ByVal value As String)
      _fun = fun
      _value = value
      _report = Nothing
      _timestamp = 0
    End Sub

    Public Sub New(ByVal fun As YFunction, ByVal timestamp As Double, ByVal report As List(Of Integer))
      _fun = fun
      _value = Nothing
      _timestamp = timestamp
      _report = report
    End Sub

    Public Sub invoke()
      If (_value Is Nothing) Then
        Dim sensor As YSensor = CType(_fun, YSensor)
        Dim measure As YMeasure = sensor._decodeTimedReport(_timestamp, _report)
        sensor._invokeTimedReportCallback(measure)
      Else
        REM new value
        _fun._invokeValueCallback(_value)
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

    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then
      Exit Sub
    End If
    modul = YModule.FindModule(infos.serial + ".module")
    modul.setImmutableAttributes(infos)
    ev = New PlugEvent(PlugEvent.EVTYPE.ARRIVAL, modul)
    If (yArrival <> Nothing) Then _PlugEvents.Add(ev)
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
  '''   a procedure taking a <c>YModule</c> parameter, or <c>null</c>
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
    If (yRemoval = Nothing) Then Exit Sub
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
  '''   a procedure taking a <c>YModule</c> parameter, or <c>null</c>
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
  '''   a procedure taking two string parameter, or null
  '''   to unregister a previously registered  callback.
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

    If (yChange = Nothing) Then Exit Sub
    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then Exit Sub
    modul = YModule.FindModule(infos.serial + ".module")
    ev = New PlugEvent(PlugEvent.EVTYPE.CHANGE, modul)
    _PlugEvents.Add(ev)
  End Sub

  Public Sub yRegisterDeviceChangeCallback(ByVal callback As yDeviceUpdateFunc)
    YAPI.RegisterDeviceChangeCallback(callback)
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
        report = New List(Of Integer)(intlen)
        p = 0
        While (p < intlen)
          report.Add(data(p))
          p = p + 1
        End While
        ev = New DataEvent(YFunction._TimedReportCallbackList(i), timestamp, report)
        _DataEvents.Add(ev)
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
    Return YAPI.yGetAPIVersion()
  End Function


  '''*
  '''/
  Public Const Y_DETECT_USB As Integer = YAPI.DETECT_USB
  Public Const Y_DETECT_NET As Integer = YAPI.DETECT_NET
  Public Const Y_DETECT_ALL As Integer = YAPI.DETECT_ALL
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
  '''   <c>http://username:password@adresse:port</c>
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
  '''   Fault-tolerant alternative to RegisterHub().
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
  '''   Force a hub discovery, if a callback as been registered with <c>yRegisterDeviceRemovalCallback</c> it
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


  Public Function yapiGetFunctionInfo(ByVal fundesc As YFUN_DESCR, _
                                        ByRef devdesc As YDEV_DESCR, _
                                        ByRef serial As String, _
                                        ByRef funcId As String, _
                                        ByRef funcName As String, _
                                        ByRef funcVal As String, _
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

    yapiGetFunctionInfo = _yapiGetFunctionInfo(fundesc, devdesc, serialBuffer, funcIdBuffer, funcNameBuffer, funcValBuffer, errBuffer)
    serial = serialBuffer.ToString()
    funcId = funcIdBuffer.ToString()
    funcName = funcNameBuffer.ToString()
    funcVal = funcValBuffer.ToString()
    errmsg = funcValBuffer.ToString()
  End Function


  Public Function yapiGetDeviceByFunction(ByVal fundesc As YFUN_DESCR, ByRef errmsg As String) As Integer
    Dim errBuffer As New StringBuilder(YOCTO_ERRMSG_LEN)
    Dim devdesc As YDEV_DESCR
    Dim res As Integer
    errBuffer.Length = 0
    res = _yapiGetFunctionInfo(fundesc, devdesc, Nothing, Nothing, Nothing, Nothing, errBuffer)
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

  <DllImport("myDll.dll", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub DllCallTest(ByRef data As yDeviceSt)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiInitAPI", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiInitAPI(ByVal mode As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiFreeAPI", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiFreeAPI()
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiRegisterLogFunction", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiRegisterLogFunction(ByVal fct As IntPtr)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiRegisterDeviceArrivalCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiRegisterDeviceArrivalCallback(ByVal fct As IntPtr)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiRegisterDeviceRemovalCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiRegisterDeviceRemovalCallback(ByVal fct As IntPtr)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiRegisterDeviceChangeCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiRegisterDeviceChangeCallback(ByVal fct As IntPtr)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiRegisterFunctionUpdateCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiRegisterFunctionUpdateCallback(ByVal fct As IntPtr)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiRegisterTimedReportCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiRegisterTimedReportCallback(ByVal fct As IntPtr)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiLockDeviceCallBack", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiLockDeviceCallBack(ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiUnlockDeviceCallBack", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiUnlockDeviceCallBack(ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiLockFunctionCallBack", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiLockFunctionCallBack(ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiUnlockFunctionCallBack", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiUnlockFunctionCallBack(ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiRegisterHub", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiRegisterHub(ByVal rootUrl As StringBuilder, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiPreregisterHub", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiPreregisterHub(ByVal rootUrl As StringBuilder, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiUnregisterHub", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiUnregisterHub(ByVal rootUrl As StringBuilder)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiUpdateDeviceList", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiUpdateDeviceList(ByVal force As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiHandleEvents", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHandleEvents(ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetTickCount", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetTickCount() As yu64
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiCheckLogicalName", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiCheckLogicalName(ByVal name As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetAPIVersion", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetAPIVersion(ByRef version As IntPtr, ByRef subversion As IntPtr) As yu16
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetDevice", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetDevice(ByVal device_str As StringBuilder, ByVal errmsg As StringBuilder) As YDEV_DESCR
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetAllDevices", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetAllDevices(ByVal buffer As IntPtr, ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetDeviceInfo", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetDeviceInfo(ByVal d As YDEV_DESCR, ByRef infos As yDeviceSt, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetFunction", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetFunction(ByVal class_str As StringBuilder, ByVal function_str As StringBuilder, ByVal errmsg As StringBuilder) As YFUN_DESCR
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetFunctionsByClass", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetFunctionsByClass(ByVal class_str As StringBuilder, ByVal precFuncDesc As YFUN_DESCR, ByVal buffer As IntPtr, _
                                       ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetFunctionsByDevice", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetFunctionsByDevice(ByVal device As YDEV_DESCR, ByVal precFuncDesc As YFUN_DESCR, ByVal buffer As IntPtr, ByVal maxsize As Integer, ByRef neededsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetFunctionInfo", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetFunctionInfo(ByVal fundesc As YFUN_DESCR, _
                                         ByRef devdesc As YDEV_DESCR, _
                                         ByVal serial As StringBuilder, _
                                         ByVal funcId As StringBuilder, _
                                         ByVal funcName As StringBuilder, _
                                         ByVal funcVal As StringBuilder, _
                                         ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetErrorString", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetErrorString(ByVal errorcode As Integer, ByVal buffer As StringBuilder, ByVal maxsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncStart", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHTTPRequestSyncStart(ByRef iohdl As YIOHDL, ByVal device As StringBuilder, ByVal url As StringBuilder, ByRef reply As IntPtr, ByRef replysize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncStartEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHTTPRequestSyncStartEx(ByRef iohdl As YIOHDL, ByVal device As StringBuilder, ByVal url As IntPtr, ByVal urllen As Integer, ByRef reply As IntPtr, ByRef replysize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function


  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestSyncDone", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHTTPRequestSyncDone(ByRef iohdl As YIOHDL, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestAsync", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHTTPRequestAsync(ByVal device As StringBuilder, ByVal url As IntPtr, ByVal callback As IntPtr, ByVal context As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", EntryPoint:="yapiHTTPRequestAsyncEx", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHTTPRequestAsyncEx(ByVal device As StringBuilder, ByVal url As IntPtr, ByVal urllen As Integer, ByVal callback As IntPtr, ByVal context As IntPtr, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiHTTPRequest", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiHTTPRequest(ByVal device As StringBuilder, ByVal url As StringBuilder, ByVal buffer As StringBuilder, ByVal buffsize As Integer, ByRef fullsize As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetBootloadersDevs", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetBootloadersDevs(ByVal serials As StringBuilder, ByVal maxNbSerial As yu32, ByRef totalBootladers As yu32, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiFlashDevice", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiFlashDevice(ByRef args As yFlashArg, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiVerifyDevice", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiVerifyDevice(ByRef args As yFlashArg, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiGetDevicePath", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiGetDevicePath(ByVal devdesc As Integer, _
                                           ByVal rootdevice As StringBuilder, _
                                           ByVal path As StringBuilder, _
                                           ByVal pathsize As Integer, _
                                           ByRef neededsize As Integer, _
                                           ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiSleep", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiSleep(ByVal duration_ms As Integer, ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiRegisterHubDiscoveryCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiRegisterHubDiscoveryCallback(ByVal fct As IntPtr)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiTriggerHubDiscovery", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Function _yapiTriggerHubDiscovery(ByVal errmsg As StringBuilder) As Integer
  End Function

  <DllImport("yapi.dll", entrypoint:="yapiRegisterDeviceLogCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiRegisterDeviceLogCallback(ByVal fct As IntPtr)
  End Sub

  <DllImport("yapi.dll", entrypoint:="yapiStartStopDeviceLogCallback", CharSet:=CharSet.Ansi, CallingConvention:=CallingConvention.Cdecl)> _
  Private Sub _yapiStartStopDeviceLogCallback(ByVal serial As StringBuilder, ByVal startStop As Integer)
  End Sub


  Private Sub vbmodule_initialization()
    YDevice_devCache = New List(Of YDevice)
    REM --- (generated code: Function initialization)
    REM --- (end of generated code: Function initialization)
    REM --- (generated code: Module initialization)
    REM --- (end of generated code: Module initialization)
    REM --- (generated code: DataStream initialization)
    REM --- (end of generated code: DataStream initialization)
    REM --- (generated code: Measure initialization)
    REM --- (end of generated code: Measure initialization)
    REM --- (generated code: DataSet initialization)
    REM --- (end of generated code: DataSet initialization)
    _PlugEvents = New List(Of PlugEvent)
    _DataEvents = New List(Of DataEvent)
  End Sub

  Private Sub vbmodule_cleanup()
    YDevice_devCache.Clear()
    YDevice_devCache = Nothing
    REM --- (generated code: Function cleanup)
    REM --- (end of generated code: Function cleanup)
    REM --- (generated code: Module cleanup)
    REM --- (end of generated code: Module cleanup)
    REM --- (generated code: DataStream cleanup)
    REM --- (end of generated code: DataStream cleanup)
    REM --- (generated code: Measure cleanup)
    REM --- (end of generated code: Measure cleanup)
    REM --- (generated code: DataSet cleanup)
    REM --- (end of generated code: DataSet cleanup)
    _PlugEvents.Clear()
    _PlugEvents = Nothing
    _DataEvents.Clear()
    _DataEvents = Nothing
    YAPI.handlersCleanUp()
  End Sub


End Module
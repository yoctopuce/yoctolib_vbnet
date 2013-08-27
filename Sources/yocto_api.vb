'/********************************************************************
'*
'* $Id: yocto_api.vb 12326 2013-08-13 15:52:20Z mvuilleu $
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
    Dim ivalue As Integer
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
        Buffer = ""
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
          For i = 0 To p.Value.membercount - 1
            If (i > 0) Then buffer = buffer + ","
            buffer = buffer + Me.convertToString(p.Value.members(i), True)
          Next i
          buffer = buffer + "}"
        Case TJSONRECORDTYPE.JSON_ARRAY
          buffer = buffer + "["
          For i = 0 To p.Value.itemcount - 1
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

    Private Function createIntRecord(ByVal name As String, ByVal value As Integer) As TJSONRECORD
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
      Dim ivalue, isign As Integer
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
          For i = p.membercount - 1 To 0 Step -1
            freestructure(p.members(i))
          Next i
          ReDim p.members(0)

        Case TJSONRECORDTYPE.JSON_ARRAY
          For i = p.itemcount - 1 To 0 Step -1
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

    Public Function GetAllChilds(ByVal parent As Nullable(Of TJSONRECORD)) As String()
      Dim res As String()
      Dim p As Nullable(Of TJSONRECORD) = parent

      If p Is Nothing Then p = Data

      If (p.Value.recordtype = TJSONRECORDTYPE.JSON_STRUCT) Then
        ReDim res(p.Value.membercount - 1)
        For i = 0 To p.Value.membercount - 1
          res(i) = Me.convertToString(p.Value.members(i), False)
        Next i
      ElseIf (p.Value.recordtype = TJSONRECORDTYPE.JSON_ARRAY) Then
        ReDim res(p.Value.itemcount - 1)
        For i = 0 To p.Value.itemcount - 1
          res(i) = Me.convertToString(p.Value.items(i), False)
        Next i
      Else
        ReDim res(-1)
      End If

      Return res
    End Function
  End Class

  Public Const YOCTO_API_VERSION_STR = "1.01"
  Public Const YOCTO_API_VERSION_BCD = &H101
  Public Const YOCTO_API_BUILD_NO = "12553"

  Public Const YOCTO_DEFAULT_PORT = 4444
  Public Const YOCTO_VENDORID = &H24E0
  Public Const YOCTO_DEVID_FACTORYBOOT = 1
  Public Const YOCTO_DEVID_BOOTLOADER = 2

  Public Const YOCTO_ERRMSG_LEN = 256
  Public Const YOCTO_MANUFACTURER_LEN = 20
  Public Const YOCTO_SERIAL_LEN = 20
  Public Const YOCTO_BASE_SERIAL_LEN = 8
  Public Const YOCTO_PRODUCTNAME_LEN = 28
  Public Const YOCTO_FIRMWARE_LEN = 22
  Public Const YOCTO_LOGICAL_LEN = 20
  Public Const YOCTO_FUNCTION_LEN = 20
  Public Const YOCTO_PUBVAL_SIZE = 6 ' Size of the data (can be non null terminated)
  Public Const YOCTO_PUBVAL_LEN = 16 ' Temporary storage, > YOCTO_PUBVAL_SIZE
  Public Const YOCTO_PASS_LEN = 20
  Public Const YOCTO_REALM_LEN = 20
  Public Const INVALID_YHANDLE = 0
  Public Const yUnknowSize = 1024

  REM Global definitions for YRelay,YDatalogger an YWatchdog class
  Public Const Y_STATE_A = 0
  Public Const Y_STATE_B = 1
  Public Const Y_STATE_INVALID = -1
  Public Const Y_OUTPUT_OFF = 0
  Public Const Y_OUTPUT_ON = 1
  Public Const Y_OUTPUT_INVALID = -1
  Public Const Y_AUTOSTART_OFF = 0
  Public Const Y_AUTOSTART_ON = 1
  Public Const Y_AUTOSTART_INVALID = -1
  Public Const Y_RUNNING_OFF = 0
  Public Const Y_RUNNING_ON = 1
  Public Const Y_RUNNING_INVALID = -1



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
    Public Const INVALID_LONG As Long = -9223372036854775807
    Public Const INVALID_UNSIGNED As Integer = -1
    Public Const HARDWAREID_INVALID As String = INVALID_STRING
    Public Const FUNCTIONID_INVALID As String = INVALID_STRING
    Public Const FRIENDLYNAME_INVALID As String = INVALID_STRING

    REM yInitAPI argument
    Public Const DETECT_NONE As Integer = 0
    Public Const DETECT_USB As Integer = 1
    Public Const DETECT_NET As Integer = 2
    Public Const DETECT_ALL As Integer = DETECT_USB Or DETECT_NET

    REM --- (generated code: return codes)
    REM Yoctopuce error codes, also used by default as function return value
    Public Const SUCCESS = 0                    REM everything worked allright
    Public Const NOT_INITIALIZED = -1           REM call yInitAPI() first !
    Public Const INVALID_ARGUMENT = -2          REM one of the arguments passed to the function is invalid
    Public Const NOT_SUPPORTED = -3             REM the operation attempted is (currently) not supported
    Public Const DEVICE_NOT_FOUND = -4          REM the requested device is not reachable
    Public Const VERSION_MISMATCH = -5          REM the device firmware is incompatible with this API version
    Public Const DEVICE_BUSY = -6               REM the device is busy with another task and cannot answer
    Public Const TIMEOUT = -7                   REM the device took too long to provide an answer
    Public Const IO_ERROR = -8                  REM there was an I/O problem while talking to the device
    Public Const NO_MORE_DATA = -9              REM there is no more data to read from
    Public Const EXHAUSTED = -10                REM you have run out of a limited ressource, check the documentation
    Public Const DOUBLE_ACCES = -11             REM you have two process that try to acces to the same device
    Public Const UNAUTHORIZED = -12             REM unauthorized access to password-protected device
    Public Const RTC_NOT_READY = -13            REM real-time clock has not been initialized (or time was lost)

  REM --- (end of generated code: return codes)

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



    Public Shared Function _encodeCalibrationPoints(ByVal rawValues() As Double, ByVal refValues() As Double, ByVal resolution As Double, ByVal calibrationOffset As Long, ByVal actualCparams As String) As String
      Dim i As Integer
      Dim npt As Integer
      If (rawValues.Length < refValues.Length) Then
        npt = rawValues.Length
      Else
        npt = refValues.Length
      End If
      Dim rawVal As Integer
      Dim refVal As Integer
      Dim calibtype As Integer
      Dim minRaw As Integer = 0
      Dim res As String
      If npt = 0 Then
        Return ""
      End If
      If actualCparams = "" Then
        calibtype = 10 + npt
      Else
        Dim pos As Integer = actualCparams.IndexOf(","c)
        calibtype = Convert.ToInt32(actualCparams.Substring(0, pos))
        If (calibtype <= 10) Then
          calibtype = npt
        Else
          calibtype = 10 + npt
        End If
      End If

      res = calibtype.ToString()
      If (calibtype <= 10) Then
        For i = 0 To npt - 1 Step 1
          rawVal = CInt(Math.Round(rawValues(i) / resolution - calibrationOffset))
          If (rawVal >= minRaw And rawVal < 65536) Then
            refVal = CInt(Math.Round(refValues(i) / resolution - calibrationOffset))
            If (refVal >= 0 And refVal < 65536) Then
              res += "," + rawVal.ToString() + "," + refVal.ToString()
              minRaw = rawVal + 1
            End If
          End If
        Next
      Else
        REM 16-bit floating-point decimal encoding
        For i = 0 To npt - 1 Step 1
          rawVal = CInt(_doubleToDecimal(rawValues(i)))
          refVal = CInt(_doubleToDecimal(refValues(i)))
          res += "," + rawVal.ToString() + "," + refVal.ToString()
        Next
      End If
      Return res
    End Function



    Public Shared Function _decodeCalibrationPoints(ByVal calibParams As String, ByRef intPt() As Integer, ByRef rawPt() As Double, ByRef calPt() As Double, ByRef resolution As Double, ByVal calibrationOffset As Long) As Integer
      Dim valuesStr() As String
      valuesStr = calibParams.Split(","c)
      If valuesStr.Length <= 1 Then
        Return 0
      End If
      Dim calibType As Integer = Convert.ToInt32(valuesStr(0))
      Dim nval As Integer = 99
      If (calibType < 20) Then
        nval = 2 * (calibType Mod 10)
      End If
      Array.Resize(intPt, nval)
      Array.Resize(rawPt, CInt(nval / 2))
      Array.Resize(calPt, CInt(nval / 2))
      Dim i As Integer = 1
      While (i < nval And i < valuesStr.Length)
        Dim rawval As Integer = Convert.ToInt32(valuesStr(i))
        Dim calval As Integer = Convert.ToInt32(valuesStr(i + 1))
        Dim rawval_d As Double
        Dim calval_d As Double
        If (calibType <= 10) Then
          rawval_d = (rawval + calibrationOffset) * resolution
          calval_d = (calval + calibrationOffset) * resolution
        Else
          rawval_d = _decimalToDouble(rawval)
          calval_d = _decimalToDouble(calval)
        End If
        intPt(i - 1) = rawval
        intPt(i) = calval
        rawPt((i - 1) >> 1) = rawval_d
        calPt((i - 1) >> 1) = calval_d
        i = i + 2
      End While
      Return calibType
    End Function


    Public Shared Function _applyCalibration(ByVal rawValue As Double, ByVal calibParams As String, ByVal calibOffset As Long, ByVal resolution As Double) As Double
      If ((rawValue = YAPI.INVALID_DOUBLE) Or (resolution = YAPI.INVALID_DOUBLE)) Then
        Return YAPI.INVALID_DOUBLE
      End If
      If (calibParams.IndexOf(","c) <= 0) Then
        Return YAPI.INVALID_DOUBLE
      End If
      Dim cur_calpar() As Integer = Nothing
      Dim cur_calraw() As Double = Nothing
      Dim cur_calref() As Double = Nothing
      Dim calibType As Integer = YAPI._decodeCalibrationPoints(calibParams, cur_calpar, cur_calraw, cur_calref, resolution, calibOffset)
      If (calibType = 0) Then
        Return rawValue
      End If
      Dim calhdl As yCalibrationHandler
      calhdl = YAPI._getCalibrationHandler(calibType)
      If (calhdl = Nothing) Then
        Return YAPI.INVALID_DOUBLE
      End If
      Return calhdl(rawValue, calibType, cur_calpar, cur_calraw, cur_calref)
    End Function

    Public Shared Sub RegisterCalibrationHandler(ByVal calibType As Integer, ByVal callback As yCalibrationHandler)

      Dim key As String
      key = calibType.ToString()
      _CalibHandlers.Add(key, callback)
    End Sub

    Private Shared Function yLinearCalibrationHandler(ByVal rawValue As Double, ByVal calibType As Integer, ByVal parameters() As Integer, ByVal rawValues() As Double, ByVal refValues() As Double) As Double

      Dim npt, i As Integer
      Dim x, adj As Double
      Dim x2, adj2 As Double


      npt = calibType Mod 10
      x = rawValues(0)
      adj = refValues(0) - x
      i = 0

      If (npt > rawValues.Length) Then npt = rawValues.Length
      If (npt > refValues.Length) Then npt = refValues.Length + 1
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

      _yapiRegisterLogFunction(Marshal.GetFunctionPointerForDelegate(native_yLogFunctionDelegate))


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
    '''   When using Yoctopuce modules through the VirtualHub gateway,
    '''   you should provide as parameter the address of the machine on which the
    '''   VirtualHub software is running (typically <c>"http://127.0.0.1:4444"</c>,
    '''   which represents the local machine).
    '''   When you use a language which has direct access to the USB hardware,
    '''   you can use the pseudo-URL <c>"usb"</c> instead.
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
    '''   If acces control has been activated on the VirtualHub you want to
    '''   reach, the URL parameter should look like:
    ''' </para>
    ''' <para>
    '''   <c>http://username:password@adresse:port</c>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="url">
    '''   a string containing either <c>"usb"</c> or the
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
    ''' 
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
      Dim p As yapiEvent


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
        Select Case p.eventtype
          Case yapiEventType.YAPI_DEV_ARRIVAL
            If yArrival <> Nothing Then yArrival(p.modul)

          Case yapiEventType.YAPI_DEV_REMOVAL
            If yRemoval <> Nothing Then yRemoval(p.modul)

          Case yapiEventType.YAPI_DEV_CHANGE
            If (yChange <> Nothing) Then yChange(p.modul)
        End Select

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
      Dim p As yapiEvent
      REM Dim serial, funcId, funcName, discard As String
      REM Dim devdescr As YDEV_DESCR
      REM Dim modul As YModule
      Dim i As Integer

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
          p = _DataEvents(0)
          _DataEvents.RemoveAt(0)
          yapiUnlockFunctionCallBack(errmsg)
          If (p.eventtype = yapiEventType.YAPI_FUN_VALUE) Then
            For i = 0 To YFunction._FunctionCallbacks.Count - 1
              If (YFunction._FunctionCallbacks(i).get_functionDescriptor() = p.fun_descr) Then
                YFunction._FunctionCallbacks(i).advertiseValue(p.value)
              End If
            Next i
          End If
        End If

      End While
      HandleEvents = YAPI_SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Pauses the execution flow for a specified duration.
    ''' <para>
    '''   This function implements a passive waiting loop, meaning that it does not
    '''   consume CPU cycles significatively. The processor is left available for
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
    '''   Returns the current value of a monotone millisecond-based time counter.
    ''' <para>
    '''   This counter can be used to compute delays in relation with
    '''   Yoctopuce devices, which also uses the milisecond as timebase.
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
    '''   the API have something to say. Quite usefull to debug the API.
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
    '''   a device is pluged.
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
        Dim m As YModule = yFirstModule()
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
    '''   a device is unpluged.
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

    Public Shared Sub yRegisterDeviceChangeCallback(ByVal callback As yDeviceUpdateFunc)
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
  Public Const YIOHDL_SIZE = 8
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




  REM --- (generated code: YModule definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YModule, ByVal value As String)


 REM Yoctopuce error codes, also used by default as function return value
  Public Const YAPI_SUCCESS = 0                    REM everything worked allright
  Public Const YAPI_NOT_INITIALIZED = -1           REM call yInitAPI() first !
  Public Const YAPI_INVALID_ARGUMENT = -2          REM one of the arguments passed to the function is invalid
  Public Const YAPI_NOT_SUPPORTED = -3             REM the operation attempted is (currently) not supported
  Public Const YAPI_DEVICE_NOT_FOUND = -4          REM the requested device is not reachable
  Public Const YAPI_VERSION_MISMATCH = -5          REM the device firmware is incompatible with this API version
  Public Const YAPI_DEVICE_BUSY = -6               REM the device is busy with another task and cannot answer
  Public Const YAPI_TIMEOUT = -7                   REM the device took too long to provide an answer
  Public Const YAPI_IO_ERROR = -8                  REM there was an I/O problem while talking to the device
  Public Const YAPI_NO_MORE_DATA = -9              REM there is no more data to read from
  Public Const YAPI_EXHAUSTED = -10                REM you have run out of a limited ressource, check the documentation
  Public Const YAPI_DOUBLE_ACCES = -11             REM you have two process that try to acces to the same device
  Public Const YAPI_UNAUTHORIZED = -12             REM unauthorized access to password-protected device
  Public Const YAPI_RTC_NOT_READY = -13            REM real-time clock has not been initialized (or time was lost)

  Public Const Y_PRODUCTNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SERIALNUMBER_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PRODUCTID_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_PRODUCTRELEASE_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_FIRMWARERELEASE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PERSISTENTSETTINGS_LOADED = 0
  Public Const Y_PERSISTENTSETTINGS_SAVED = 1
  Public Const Y_PERSISTENTSETTINGS_MODIFIED = 2
  Public Const Y_PERSISTENTSETTINGS_INVALID = -1

  Public Const Y_LUMINOSITY_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_BEACON_OFF = 0
  Public Const Y_BEACON_ON = 1
  Public Const Y_BEACON_INVALID = -1

  Public Const Y_UPTIME_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_USBCURRENT_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_REBOOTCOUNTDOWN_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_USBBANDWIDTH_SIMPLE = 0
  Public Const Y_USBBANDWIDTH_DOUBLE = 1
  Public Const Y_USBBANDWIDTH_INVALID = -1



  REM --- (end of generated code: YModule definitions)



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
  Public Delegate Function yCalibrationHandler(ByVal rawValue As Double, ByVal calibType As Integer, ByVal parameters() As Integer, ByVal rawValues() As Double, ByVal refValues() As Double) As Double


  REM - Types used for internal yapi callbacks
  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub _yapiLogFunc(ByVal log As IntPtr, ByVal loglen As yu32)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub _yapiDeviceUpdateFunc(ByVal dev As YDEV_DESCR)

  <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
  Public Delegate Sub _yapiFunctionUpdateFunc(ByVal dev As YFUN_DESCR, ByVal value As IntPtr)

  REM - Variables used to store public yocto_api callbacks
  Dim ylog As yLogFunc = Nothing
  Dim yArrival As yDeviceUpdateFunc = Nothing
  Dim yRemoval As yDeviceUpdateFunc = Nothing
  Dim yChange As yDeviceUpdateFunc = Nothing

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
      Dim replysize = 0
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


  '''*
  ''' <summary>YFunction Class (virtual class, used internally)
  ''' <para>
  ''' This is the parent class for all public objects representing device functions documented in
  ''' the high-level programming API. This abstract class does all the real job, but without
  ''' knowledge of the specific function attributes.
  ''' </para>
  ''' <para>
  ''' Instantiating a child class of YFunction does not cause any communication.
  ''' The instance simply keeps track of its function identifier, and will dynamically bind
  ''' to a matching device at the time it is really beeing used to read or set an attribute.
  ''' In order to allow true hot-plug replacement of one device by another, the binding stay
  ''' dynamic through the life of the object.
  ''' </para>
  ''' <para>
  ''' The YFunction class implements a generic high-level cache for the attribute values of
  ''' the specified function, pre-parsed from the REST API string.
  ''' </para>
  ''' </summary>
  '''/



  Public MustInherit Class YFunction

    Public Shared _FunctionCache As List(Of YFunction) = New List(Of YFunction)
    Public Shared _FunctionCallbacks As List(Of YFunction) = New List(Of YFunction)
    Public Delegate Sub GenericUpdateCallback(ByVal func As YFunction, ByVal value As String)

    Public Const FUNCTIONDESCRIPTOR_INVALID As YFUN_DESCR = -1

    Protected _className As String
    Protected _func As String
    Protected _lastErrorType As YRETCODE
    Protected _lastErrorMsg As String

    Protected _fundescr As YFUN_DESCR
    Protected _userData As Object
    REM protected Action<T,T>   _callback;
    Protected _genCallback As GenericUpdateCallback



    Protected _cacheExpiration As System.UInt64

    Public Sub New(ByRef classname As String, ByRef func As String)
      _className = classname
      _func = func
      _lastErrorType = YAPI_SUCCESS
      _lastErrorMsg = ""
      _cacheExpiration = 0
      _fundescr = FUNCTIONDESCRIPTOR_INVALID
      _userData = Nothing
      _genCallback = Nothing
      _FunctionCache.Add(Me)

    End Sub

    Protected Sub _throw(ByVal errType As YRETCODE, ByVal errMsg As String)
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

      Const n_element = 1
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

    Protected MustOverride Function _parse(ByRef parser As TJSONRECORD) As Integer

    '''*
    ''' <summary>
    '''   Returns a short text that describes the function in the form <c>TYPE(NAME)=SERIAL&#46;FUNCTIONID</c>.
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
    '''   Returns the unique hardware identifier of the function in the form <c>SERIAL&#46;FUNCTIONID</c>.
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
    '''   Returns the numerical error code of the latest error with this function.
    ''' <para>
    '''   This method is mostly useful when using the Yoctopuce library with
    '''   exceptions disabled.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a number corresponding to the code of the latest error that occured while
    '''   using this function object
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
    '''   Returns the error message of the latest error with this function.
    ''' <para>
    '''   This method is mostly useful when using the Yoctopuce library with
    '''   exceptions disabled.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest error message that occured while
    '''   using this function object
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
    '''   device hosting the requested function.
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

      isOnline = True
    End Function

    Protected Function _json_get_key(ByVal data As Byte(), ByVal key As String) As String
      Dim node As Nullable(Of TJSONRECORD)

      Dim st As String = YAPI.DefaultEncoding.GetString(data)
      Dim p As TJsonParser = New TJsonParser(st, False)
      node = p.GetChildNode(Nothing, key)

      Return node.Value.svalue
    End Function

    Protected Function _json_get_array(ByVal data As Byte()) As String()
      Dim st As String = YAPI.DefaultEncoding.GetString(data)
      Dim p As TJsonParser = New TJsonParser(st, False)

      Return p.GetAllChilds(Nothing)
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

      node = apires.GetChildNode(Nothing, funcId)
      If Not (node.HasValue) Then
        _throw(YAPI_IO_ERROR, "unexpected JSON structure: missing function " + funcId)
        load = res
        Exit Function
      End If

      _parse(CType(node, TJSONRECORD))
      _cacheExpiration = CULng(YAPI.GetTickCount() + msValidity)
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

    Protected Sub registerFuncCallback(ByVal func As YFunction)
      isOnline()
      If Not (_FunctionCallbacks.Contains(Me)) Then
        _FunctionCallbacks.Add(Me)
      End If
    End Sub

    Protected Sub unregisterFuncCallback(ByVal func As YFunction)
      _FunctionCallbacks.Remove(Me)
    End Sub

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
    Public Sub registerValueCallback(ByVal callback As GenericUpdateCallback)
      If (callback IsNot Nothing) Then
        registerFuncCallback(Me)
      Else
        unregisterFuncCallback(Me)
      End If
      _genCallback = callback
    End Sub



    Public Overridable Sub advertiseValue(ByVal value As String)
      If (_genCallback IsNot Nothing) Then _genCallback(Me, value)
    End Sub


  End Class

  REM --- (generated code: YModule implementation)

  Private _ModuleCache As New Hashtable()
  Private _callback As UpdateCallback

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
    Public Const PRODUCTNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const SERIALNUMBER_INVALID As String = YAPI.INVALID_STRING
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const PRODUCTID_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const PRODUCTRELEASE_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const FIRMWARERELEASE_INVALID As String = YAPI.INVALID_STRING
    Public Const PERSISTENTSETTINGS_LOADED = 0
    Public Const PERSISTENTSETTINGS_SAVED = 1
    Public Const PERSISTENTSETTINGS_MODIFIED = 2
    Public Const PERSISTENTSETTINGS_INVALID = -1

    Public Const LUMINOSITY_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const BEACON_OFF = 0
    Public Const BEACON_ON = 1
    Public Const BEACON_INVALID = -1

    Public Const UPTIME_INVALID As Long = YAPI.INVALID_LONG
    Public Const USBCURRENT_INVALID As Long = YAPI.INVALID_LONG
    Public Const REBOOTCOUNTDOWN_INVALID As Integer = YAPI.INVALID_INT
    Public Const USBBANDWIDTH_SIMPLE = 0
    Public Const USBBANDWIDTH_DOUBLE = 1
    Public Const USBBANDWIDTH_INVALID = -1


    Protected _productName As String
    Protected _serialNumber As String
    Protected _logicalName As String
    Protected _productId As Long
    Protected _productRelease As Long
    Protected _firmwareRelease As String
    Protected _persistentSettings As Long
    Protected _luminosity As Long
    Protected _beacon As Long
    Protected _upTime As Long
    Protected _usbCurrent As Long
    Protected _rebootCountdown As Long
    Protected _usbBandwidth As Long

    Public Sub New(ByVal func As String)
      MyBase.new("Module", func)
      _productName = Y_PRODUCTNAME_INVALID
      _serialNumber = Y_SERIALNUMBER_INVALID
      _logicalName = Y_LOGICALNAME_INVALID
      _productId = Y_PRODUCTID_INVALID
      _productRelease = Y_PRODUCTRELEASE_INVALID
      _firmwareRelease = Y_FIRMWARERELEASE_INVALID
      _persistentSettings = Y_PERSISTENTSETTINGS_INVALID
      _luminosity = Y_LUMINOSITY_INVALID
      _beacon = Y_BEACON_INVALID
      _upTime = Y_UPTIME_INVALID
      _usbCurrent = Y_USBCURRENT_INVALID
      _rebootCountdown = Y_REBOOTCOUNTDOWN_INVALID
      _usbBandwidth = Y_USBBANDWIDTH_INVALID
    End Sub

    Protected Overrides Function _parse(ByRef j As TJSONRECORD) As Integer
      Dim member As TJSONRECORD
      Dim i As Integer
      If (j.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then
        Return -1
      End If
      For i = 0 To j.membercount - 1
        member = j.members(i)
        If (member.name = "productName") Then
          _productName = member.svalue
        ElseIf (member.name = "serialNumber") Then
          _serialNumber = member.svalue
        ElseIf (member.name = "logicalName") Then
          _logicalName = member.svalue
        ElseIf (member.name = "productId") Then
          _productId = CLng(member.ivalue)
        ElseIf (member.name = "productRelease") Then
          _productRelease = CLng(member.ivalue)
        ElseIf (member.name = "firmwareRelease") Then
          _firmwareRelease = member.svalue
        ElseIf (member.name = "persistentSettings") Then
          _persistentSettings = CLng(member.ivalue)
        ElseIf (member.name = "luminosity") Then
          _luminosity = CLng(member.ivalue)
        ElseIf (member.name = "beacon") Then
          If (member.ivalue > 0) Then _beacon = 1 Else _beacon = 0
        ElseIf (member.name = "upTime") Then
          _upTime = CLng(member.ivalue)
        ElseIf (member.name = "usbCurrent") Then
          _usbCurrent = CLng(member.ivalue)
        ElseIf (member.name = "rebootCountdown") Then
          _rebootCountdown = member.ivalue
        ElseIf (member.name = "usbBandwidth") Then
          _usbBandwidth = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

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
      If (_productName = Y_PRODUCTNAME_INVALID) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PRODUCTNAME_INVALID
        End If
      End If
      Return _productName
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
      If (_serialNumber = Y_SERIALNUMBER_INVALID) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_SERIALNUMBER_INVALID
        End If
      End If
      Return _serialNumber
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the module.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the module
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LOGICALNAME_INVALID
        End If
      End If
      Return _logicalName
    End Function

    '''*
    ''' <summary>
    '''   Changes the logical name of the module.
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
    '''   a string corresponding to the logical name of the module
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
      If (_productId = Y_PRODUCTID_INVALID) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PRODUCTID_INVALID
        End If
      End If
      Return CType(_productId,Integer)
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
      If (_productRelease = Y_PRODUCTRELEASE_INVALID) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PRODUCTRELEASE_INVALID
        End If
      End If
      Return CType(_productRelease,Integer)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_FIRMWARERELEASE_INVALID
        End If
      End If
      Return _firmwareRelease
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PERSISTENTSETTINGS_INVALID
        End If
      End If
      Return CType(_persistentSettings,Integer)
    End Function

    Public Function set_persistentSettings(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("persistentSettings", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Saves current settings in the nonvolatile memory of the module.
    ''' <para>
    '''   Warning: the number of allowed save operations during a module life is
    '''   limited (about 100000 cycles). Do not call this function within a loop.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function saveToFlash() As Integer
      Dim rest_val As String
      rest_val = "1"
      Return _setAttr("persistentSettings", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Reloads the settings stored in the nonvolatile memory, as
    '''   when the module is powered on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function revertFromFlash() As Integer
      Dim rest_val As String
      rest_val = "0"
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LUMINOSITY_INVALID
        End If
      End If
      Return CType(_luminosity,Integer)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_BEACON_INVALID
        End If
      End If
      Return CType(_beacon,Integer)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_UPTIME_INVALID
        End If
      End If
      Return _upTime
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
    Public Function get_usbCurrent() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_USBCURRENT_INVALID
        End If
      End If
      Return _usbCurrent
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_REBOOTCOUNTDOWN_INVALID
        End If
      End If
      Return CType(_rebootCountdown,Integer)
    End Function

    Public Function set_rebootCountdown(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("rebootCountdown", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Schedules a simple module reboot after the given number of seconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="secBeforeReboot">
    '''   number of seconds before rebooting
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
    Public Function reboot(ByVal secBeforeReboot As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(secBeforeReboot))
      Return _setAttr("rebootCountdown", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Schedules a module reboot into special firmware update mode.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="secBeforeReboot">
    '''   number of seconds before rebooting
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
    Public Function triggerFirmwareUpdate(ByVal secBeforeReboot As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(-secBeforeReboot))
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_USBBANDWIDTH_INVALID
        End If
      End If
      Return CType(_usbBandwidth,Integer)
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
    public function download(pathname as string) as byte()
        Return Me._download(pathname)
        
     end function

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
    public function get_icon2d() as byte()
        Return Me._download("icon2d.png")
        
     end function

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
    public function get_lastLogs() as string
        dim  content as byte()
        content = Me._download("logs.txt")
        Return YAPI.DefaultEncoding.GetString(content)
        
     end function


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
    Public Function nextModule() as YModule
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindModule(hwid)
    End Function

    '''*
    ''' <summary>
    '''   comment from .
    ''' <para>
    '''   yc definition
    ''' </para>
    ''' </summary>
    '''/
  Public Overloads Sub registerValueCallback(ByVal callback As UpdateCallback)
   If (callback IsNot Nothing) Then
     registerFuncCallback(Me)
   Else
     unregisterFuncCallback(Me)
   End If
   _callback = callback
  End Sub

  Public Sub set_callback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Sub setCallback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Overrides Sub advertiseValue(ByVal value As String)
    If (_callback IsNot Nothing) Then _callback(Me, value)
  End Sub


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
    Public Shared Function FindModule(ByVal func As String) As YModule
      Dim res As YModule
      If (_ModuleCache.ContainsKey(func)) Then
        Return CType(_ModuleCache(func), YModule)
      End If
      res = New YModule(func)
      _ModuleCache.Add(func, res)
      Return res
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

    REM --- (end of generated code: YModule implementation)
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

  End Class

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
  '''   the API have something to say. Quite usefull to debug the API.
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

  Enum yapiEventType
    YAPI_DEV_ARRIVAL
    YAPI_DEV_REMOVAL
    YAPI_DEV_CHANGE
    YAPI_FUN_UPDATE
    YAPI_FUN_VALUE
    YAPI_NOP
  End Enum

  Private Structure yapiEvent
    Dim eventtype As yapiEventType
    Dim modul As YModule
    Dim fun_descr As YFUN_DESCR
    Dim value As String
  End Structure

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

  Private Function emptyApiEvent() As yapiEvent
    Dim ev As yapiEvent
    ev.eventtype = yapiEventType.YAPI_NOP
    ev.modul = Nothing
    ev.fun_descr = 0
    ev.value = ""
    emptyApiEvent = ev
  End Function

  Dim _PlugEvents As List(Of yapiEvent)
  Dim _DataEvents As List(Of yapiEvent)

  Private Sub native_yDeviceArrivalCallback(ByVal d As YDEV_DESCR)
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim ev As yapiEvent = emptyApiEvent()
    Dim errmsg As String = ""

    ev.eventtype = yapiEventType.YAPI_DEV_ARRIVAL
    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then
      Exit Sub
    End If
    ev.modul = yFindModule(infos.serial + ".module")
    ev.modul.setImmutableAttributes(infos)
    If (yArrival <> Nothing) Then _PlugEvents.Add(ev)
  End Sub

  '''*
  ''' <summary>
  '''   Register a callback function, to be called each time
  '''   a device is pluged.
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
    Dim ev As yapiEvent = emptyApiEvent()
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim errmsg As String = ""
    If (yRemoval = Nothing) Then Exit Sub
    ev.fun_descr = 0
    ev.value = ""
    ev.eventtype = yapiEventType.YAPI_DEV_REMOVAL
    infos.deviceid = 0
    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then Exit Sub
    ev.modul = yFindModule(infos.serial + ".module")
    _PlugEvents.Add(ev)
  End Sub

  '''*
  ''' <summary>
  '''   Register a callback function, to be called each time
  '''   a device is unpluged.
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

  Public Sub native_yDeviceChangeCallback(ByVal d As YDEV_DESCR)
    Dim ev As yapiEvent = emptyApiEvent()
    Dim infos As yDeviceSt = emptyDeviceSt()
    Dim errmsg As String = ""

    If (yChange = Nothing) Then Exit Sub
    ev.eventtype = yapiEventType.YAPI_DEV_CHANGE
    If (yapiGetDeviceInfo(d, infos, errmsg) <> YAPI_SUCCESS) Then Exit Sub
    ev.modul = yFindModule(infos.serial + ".module")
    _PlugEvents.Add(ev)
  End Sub

  Public Sub yRegisterDeviceChangeCallback(ByVal callback As yDeviceUpdateFunc)
    YAPI.yRegisterDeviceChangeCallback(callback)
  End Sub

  Private Sub queuesCleanUp()
    _PlugEvents.Clear()
    _PlugEvents = Nothing
    _DataEvents.Clear()
    _DataEvents = Nothing
  End Sub

  Private Sub native_yFunctionUpdateCallback(ByVal f As YFUN_DESCR, ByVal data As IntPtr)
    Dim ev As yapiEvent = emptyApiEvent()
    ev.fun_descr = f

    If (data = IntPtr.Zero) Then
      ev.eventtype = yapiEventType.YAPI_FUN_UPDATE
    Else
      ev.eventtype = yapiEventType.YAPI_FUN_VALUE
      ev.value = Marshal.PtrToStringAnsi(data)
    End If
    _DataEvents.Add(ev)
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




  REM - Delegate object for our internal callback, protected from GC
  Public native_yLogFunctionDelegate As _yapiLogFunc = AddressOf native_yLogFunction
  Dim native_yLogFunctionAnchor As GCHandle = GCHandle.Alloc(native_yLogFunctionDelegate)

  Public native_yFunctionUpdateDelegate As _yapiFunctionUpdateFunc = AddressOf native_yFunctionUpdateCallback
  Dim native_yFunctionUpdateAnchor As GCHandle = GCHandle.Alloc(native_yFunctionUpdateDelegate)

  Public native_yDeviceArrivalDelegate As _yapiDeviceUpdateFunc = AddressOf native_yDeviceArrivalCallback
  Dim native_yDeviceArrivalAnchor As GCHandle = GCHandle.Alloc(native_yDeviceArrivalDelegate)

  Public native_yDeviceRemovalDelegate As _yapiDeviceUpdateFunc = AddressOf native_yDeviceRemovalCallback
  Dim native_yDeviceRemovalAnchor As GCHandle = GCHandle.Alloc(native_yDeviceRemovalDelegate)

  Public native_yDeviceChangeDelegate As _yapiDeviceUpdateFunc = AddressOf native_yDeviceChangeCallback
  Dim native_yDeviceChangeAnchor As GCHandle = GCHandle.Alloc(native_yDeviceChangeDelegate)


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
  '''   When using Yoctopuce modules through the VirtualHub gateway,
  '''   you should provide as parameter the address of the machine on which the
  '''   VirtualHub software is running (typically <c>"http://127.0.0.1:4444"</c>,
  '''   which represents the local machine).
  '''   When you use a language which has direct access to the USB hardware,
  '''   you can use the pseudo-URL <c>"usb"</c> instead.
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
  '''   If acces control has been activated on the VirtualHub you want to
  '''   reach, the URL parameter should look like:
  ''' </para>
  ''' <para>
  '''   <c>http://username:password@adresse:port</c>
  ''' </para>
  ''' <para>
  ''' </para>
  ''' </summary>
  ''' <param name="url">
  '''   a string containing either <c>"usb"</c> or the
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
  '''
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
  '''   consume CPU cycles significatively. The processor is left available for
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
  '''   Returns the current value of a monotone millisecond-based time counter.
  ''' <para>
  '''   This counter can be used to compute delays in relation with
  '''   Yoctopuce devices, which also uses the milisecond as timebase.
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

  REM--- (generated code: Module functions)

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

  Private Sub _ModuleCleanup()
  End Sub


  REM --- (end of generated code: Module functions)


  Private Sub vbmodule_initialization()

    YDevice_devCache = New List(Of YDevice)
    REM --- (generated code: Module initialization)
    _ModuleCache = New Hashtable()
    REM --- (end of generated code: Module initialization)
    _PlugEvents = New List(Of yapiEvent)
    _DataEvents = New List(Of yapiEvent)

  End Sub

  Private Sub vbmodule_cleanup()
    YDevice_devCache.Clear()
    YDevice_devCache = Nothing
    REM --- (generated code: Module cleanup)
    _ModuleCache.Clear()
    _ModuleCache = Nothing
    REM --- (end of generated code: Module cleanup)
    _PlugEvents.Clear()
    _PlugEvents = Nothing
    _DataEvents.Clear()
    _DataEvents = Nothing
    YAPI.handlersCleanUp()
  End Sub


End Module
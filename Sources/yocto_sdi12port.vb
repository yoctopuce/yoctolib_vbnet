' ********************************************************************
'
'  $Id: yocto_sdi12port.vb 52943 2023-01-26 15:46:47Z mvuilleu $
'
'  Implements yFindSdi12Port(), the high-level API for Sdi12Port functions
'
'  - - - - - - - - - License information: - - - - - - - - -
'
'  Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'
'  Yoctopuce Sarl (hereafter Licensor) grants to you a perpetual
'  non-exclusive license to use, modify, copy and integrate this
'  file into your software for the sole purpose of interfacing
'  with Yoctopuce products.
'
'  You may reproduce and distribute copies of this file in
'  source or object form, as long as the sole purpose of this
'  code is to interface with Yoctopuce products. You must retain
'  this notice in the distributed source file.
'
'  You should refer to Yoctopuce General Terms and Conditions
'  for additional information regarding your rights and
'  obligations.
'
'  THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'  WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING
'  WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS
'  FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'  EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'  INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA,
'  COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR
'  SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT
'  LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'  CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'  BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'  WARRANTY, OR OTHERWISE.
'
' *********************************************************************


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_sdi12port

    REM --- (generated code: YSdi12Port return codes)
    REM --- (end of generated code: YSdi12Port return codes)
    REM --- (generated code: YSdi12Port dlldef)
    REM --- (end of generated code: YSdi12Port dlldef)
   REM --- (generated code: YSdi12Port yapiwrapper)
   REM --- (end of generated code: YSdi12Port yapiwrapper)
  REM --- (generated code: YSdi12Port globals)

  Public Const Y_RXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_TXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ERRCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_RXMSGCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_TXMSGCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_LASTMSG_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CURRENTJOB_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_STARTUPJOB_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_JOBMAXTASK_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_JOBMAXSIZE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_PROTOCOL_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_VOLTAGELEVEL_OFF As Integer = 0
  Public Const Y_VOLTAGELEVEL_TTL3V As Integer = 1
  Public Const Y_VOLTAGELEVEL_TTL3VR As Integer = 2
  Public Const Y_VOLTAGELEVEL_TTL5V As Integer = 3
  Public Const Y_VOLTAGELEVEL_TTL5VR As Integer = 4
  Public Const Y_VOLTAGELEVEL_RS232 As Integer = 5
  Public Const Y_VOLTAGELEVEL_RS485 As Integer = 6
  Public Const Y_VOLTAGELEVEL_TTL1V8 As Integer = 7
  Public Const Y_VOLTAGELEVEL_SDI12 As Integer = 8
  Public Const Y_VOLTAGELEVEL_INVALID As Integer = -1
  Public Const Y_SERIALMODE_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YSdi12PortValueCallback(ByVal func As YSdi12Port, ByVal value As String)
  Public Delegate Sub YSdi12PortTimedReportCallback(ByVal func As YSdi12Port, ByVal measure As YMeasure)
  REM --- (end of generated code: YSdi12Port globals)

  REM --- (generated code: YSdi12SnoopingRecord class start)

  Public Class YSdi12SnoopingRecord
    REM --- (end of generated code: YSdi12SnoopingRecord class start)
    REM --- (generated code: YSdi12SnoopingRecord definitions)
    REM --- (end of generated code: YSdi12SnoopingRecord definitions)
    REM --- (generated code: YSdi12SnoopingRecord attributes declaration)
    Protected _tim As Integer
    Protected _pos As Integer
    Protected _dir As Integer
    Protected _msg As String
    REM --- (end of generated code: YSdi12SnoopingRecord attributes declaration)

    REM --- (generated code: YSdi12SnoopingRecord private methods declaration)

    REM --- (end of generated code: YSdi12SnoopingRecord private methods declaration)

    REM --- (generated code: YSdi12SnoopingRecord public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the elapsed time, in ms, since the beginning of the preceding message.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the elapsed time, in ms, since the beginning of the preceding message.
    ''' </returns>
    '''/
    Public Overridable Function get_time() As Integer
      Return Me._tim
    End Function

    '''*
    ''' <summary>
    '''   Returns the absolute position of the message end.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the absolute position of the message end.
    ''' </returns>
    '''/
    Public Overridable Function get_pos() As Integer
      Return Me._pos
    End Function

    '''*
    ''' <summary>
    '''   Returns the message direction (RX=0, TX=1).
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the message direction (RX=0, TX=1).
    ''' </returns>
    '''/
    Public Overridable Function get_direction() As Integer
      Return Me._dir
    End Function

    '''*
    ''' <summary>
    '''   Returns the message content.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the message content.
    ''' </returns>
    '''/
    Public Overridable Function get_message() As String
      Return Me._msg
    End Function



    REM --- (end of generated code: YSdi12SnoopingRecord public methods declaration)



    Public Sub New(ByVal data As String)
      Dim m as string
      Dim json As YJSONObject  = New YJSONObject(data)
      json.parse()
      If json.has("t") Then
        Me._tim = CInt(json.getInt("t"))
      End If
      If json.has("p") Then
        Me._pos = CInt(json.getInt("p"))
      End If
      If json.has("m") Then
        m = json.getString("m")
        If m.Chars(0) = "<" Then
          Me._dir = 1
        Else
          Me._dir = 0
        End If
        Me._msg = m.Substring(1)
      End If
    End Sub

  End Class

  Public Class YSdi12SensorInfo
    REM --- (end of generated code: YSdi12SensorInfo class start)
    REM --- (generated code: YSdi12SensorInfo definitions)
    REM --- (end of generated code: YSdi12SensorInfo definitions)
    REM --- (generated code: YSdi12SensorInfo attributes declaration)
    Protected _sdi12Port As YSdi12Port
    Protected _isValid As Boolean
    Protected _addr As String
    Protected _proto As String
    Protected _mfg As String
    Protected _model As String
    Protected _ver As String
    Protected _sn As String
    Protected _valuesDesc As List(Of List(Of String))
    REM --- (end of generated code: YSdi12SensorInfo attributes declaration)

    REM --- (generated code: YSdi12SensorInfo private methods declaration)

    REM --- (end of generated code: YSdi12SensorInfo private methods declaration)

    Public Overridable Sub _throw(errcode As Integer, msg As String)
      Me._sdi12Port._throw(errcode,msg)
    End Sub

    REM --- (generated code: YSdi12SensorInfo public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the sensor state.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the sensor state.
    ''' </returns>
    '''/
    Public Overridable Function isValid() As Boolean
      Return Me._isValid
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensor address.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the sensor address.
    ''' </returns>
    '''/
    Public Overridable Function get_sensorAddress() As String
      Return Me._addr
    End Function

    '''*
    ''' <summary>
    '''   Returns the compatible SDI-12 version of the sensor.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the compatible SDI-12 version of the sensor.
    ''' </returns>
    '''/
    Public Overridable Function get_sensorProtocol() As String
      Return Me._proto
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensor vendor identification.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the sensor vendor identification.
    ''' </returns>
    '''/
    Public Overridable Function get_sensorVendor() As String
      Return Me._mfg
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensor model number.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the sensor model number.
    ''' </returns>
    '''/
    Public Overridable Function get_sensorModel() As String
      Return Me._model
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensor version.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the sensor version.
    ''' </returns>
    '''/
    Public Overridable Function get_sensorVersion() As String
      Return Me._ver
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensor serial number.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the sensor serial number.
    ''' </returns>
    '''/
    Public Overridable Function get_sensorSerial() As String
      Return Me._sn
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of sensor measurements.
    ''' <para>
    '''   This function only works if the sensor is in version 1.4 SDI-12
    '''   and supports metadata commands.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the number of sensor measurements.
    ''' </returns>
    '''/
    Public Overridable Function get_measureCount() As Integer
      Return Me._valuesDesc.Count
    End Function

    '''*
    ''' <summary>
    '''   Returns the sensor measurement command.
    ''' <para>
    '''   This function only works if the sensor is in version 1.4 SDI-12
    '''   and supports metadata commands.
    ''' </para>
    ''' </summary>
    ''' <param name="measureIndex">
    '''   measurement index
    ''' </param>
    ''' <returns>
    '''   the sensor measurement command.
    '''   On failure, throws an exception or returns an empty string.
    ''' </returns>
    '''/
    Public Overridable Function get_measureCommand(measureIndex As Integer) As String
      If Not(measureIndex < Me._valuesDesc.Count) Then
        me._throw(YAPI.INVALID_ARGUMENT, "Invalid measure index")
        return ""
      end if
      Return Me._valuesDesc(measureIndex)(0)
    End Function

    '''*
    ''' <summary>
    '''   Returns sensor measurement position.
    ''' <para>
    '''   This function only works if the sensor is in version 1.4 SDI-12
    '''   and supports metadata commands.
    ''' </para>
    ''' </summary>
    ''' <param name="measureIndex">
    '''   measurement index
    ''' </param>
    ''' <returns>
    '''   the sensor measurement command.
    '''   On failure, throws an exception or returns 0.
    ''' </returns>
    '''/
    Public Overridable Function get_measurePosition(measureIndex As Integer) As Integer
      If Not(measureIndex < Me._valuesDesc.Count) Then
        me._throw(YAPI.INVALID_ARGUMENT, "Invalid measure index")
        return 0
      end if
      Return YAPI._atoi(Me._valuesDesc(measureIndex)(2))
    End Function

    '''*
    ''' <summary>
    '''   Returns the measured value symbol.
    ''' <para>
    '''   This function only works if the sensor is in version 1.4 SDI-12
    '''   and supports metadata commands.
    ''' </para>
    ''' </summary>
    ''' <param name="measureIndex">
    '''   measurement index
    ''' </param>
    ''' <returns>
    '''   the sensor measurement command.
    '''   On failure, throws an exception or returns an empty string.
    ''' </returns>
    '''/
    Public Overridable Function get_measureSymbol(measureIndex As Integer) As String
      If Not(measureIndex < Me._valuesDesc.Count) Then
        me._throw(YAPI.INVALID_ARGUMENT, "Invalid measure index")
        return ""
      end if
      Return Me._valuesDesc(measureIndex)(3)
    End Function

    '''*
    ''' <summary>
    '''   Returns the unit of the measured value.
    ''' <para>
    '''   This function only works if the sensor is in version 1.4 SDI-12
    '''   and supports metadata commands.
    ''' </para>
    ''' </summary>
    ''' <param name="measureIndex">
    '''   measurement index
    ''' </param>
    ''' <returns>
    '''   the sensor measurement command.
    '''   On failure, throws an exception or returns an empty string.
    ''' </returns>
    '''/
    Public Overridable Function get_measureUnit(measureIndex As Integer) As String
      If Not(measureIndex < Me._valuesDesc.Count) Then
        me._throw(YAPI.INVALID_ARGUMENT, "Invalid measure index")
        return ""
      end if
      Return Me._valuesDesc(measureIndex)(4)
    End Function

    '''*
    ''' <summary>
    '''   Returns the description of the measured value.
    ''' <para>
    '''   This function only works if the sensor is in version 1.4 SDI-12
    '''   and supports metadata commands.
    ''' </para>
    ''' </summary>
    ''' <param name="measureIndex">
    '''   measurement index
    ''' </param>
    ''' <returns>
    '''   the sensor measurement command.
    '''   On failure, throws an exception or returns an empty string.
    ''' </returns>
    '''/
    Public Overridable Function get_measureDescription(measureIndex As Integer) As String
      If Not(measureIndex < Me._valuesDesc.Count) Then
        me._throw(YAPI.INVALID_ARGUMENT, "Invalid measure index")
        return ""
      end if
      Return Me._valuesDesc(measureIndex)(5)
    End Function

    Public Overridable Function get_typeMeasure() As List(Of List(Of String))
      Return Me._valuesDesc
    End Function

    Public Overridable Sub _parseInfoStr(infoStr As String)
      Dim errmsg As String

      If ((infoStr).Length > 1) Then
        If ((infoStr).Substring(0, 2) = "ER") Then
          errmsg = (infoStr).Substring(2, (infoStr).Length-2)
          Me._addr = errmsg
          Me._proto = errmsg
          Me._mfg = errmsg
          Me._model = errmsg
          Me._ver = errmsg
          Me._sn = errmsg
          Me._isValid = False
        Else
          Me._addr = (infoStr).Substring(0, 1)
          Me._proto = (infoStr).Substring(1, 2)
          Me._mfg = (infoStr).Substring(3, 8)
          Me._model = (infoStr).Substring(11, 6)
          Me._ver = (infoStr).Substring(17, 3)
          Me._sn = (infoStr).Substring(20, (infoStr).Length-20)
          Me._isValid = True
        End If
      End If
    End Sub

    Public Overridable Sub _queryValueInfo()
      Dim val As List(Of List(Of String)) = New List(Of List(Of String))()
      Dim data As List(Of String) = New List(Of String)()
      Dim infoNbVal As String
      Dim cmd As String
      Dim infoVal As String
      Dim value As String
      Dim nbVal As Integer = 0
      Dim k As Integer = 0
      Dim i As Integer = 0
      Dim j As Integer = 0
      Dim listVal As List(Of String) = New List(Of String)()
      Dim size As Integer = 0

      k = 0
      size = 4
      While (k < 10)
        infoNbVal = Me._sdi12Port.querySdi12(Me._addr, "IM" + Convert.ToString(k), 5000)
        If ((infoNbVal).Length > 1) Then
          value = (infoNbVal).Substring(4, (infoNbVal).Length-4)
          nbVal = YAPI._atoi(value)
          If (nbVal <> 0) Then
            val.Clear()
            i = 0
            While (i < nbVal)
              cmd = "IM" + Convert.ToString(k) + "_00" + Convert.ToString(i+1)
              infoVal = Me._sdi12Port.querySdi12(Me._addr, cmd, 5000)
              data = New List(Of String)(infoVal.Split(New Char() {";"c}))
              data = New List(Of String)(data(0).Split(New Char() {","c}))
              listVal.Clear()
              listVal.Add("M" + Convert.ToString(k))
              listVal.Add((i+1).ToString())
              j = 0
              While (data.Count < size)
                data.Add("")
              End While
              While (j < data.Count)
                listVal.Add(data(j))
                j = j + 1
              End While
              val.Add(New List(Of String)(listVal))
              i = i + 1
            End While
          End If
        End If
        k = k + 1
      End While
      Me._valuesDesc = val
    End Sub



    REM --- (end of generated code: YSdi12SensorInfo public methods declaration)



        Public Sub New(sdi12Port As YSdi12Port, ByVal data As String)
            Me._sdi12Port = sdi12Port
            _parseInfoStr(data)

        End Sub

    End Class


  REM --- (generated code: YSdi12Port class start)

  '''*
  ''' <summary>
  '''   The <c>YSdi12Port</c> class allows you to fully drive a Yoctopuce SDI12 port.
  ''' <para>
  '''   It can be used to send and receive data, and to configure communication
  '''   parameters (baud rate, bit count, parity, flow control and protocol).
  '''   Note that Yoctopuce SDI12 ports are not exposed as virtual COM ports.
  '''   They are meant to be used in the same way as all Yoctopuce devices.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSdi12Port
    Inherits YFunction
    REM --- (end of generated code: YSdi12Port class start)

    REM --- (generated code: YSdi12Port definitions)
    Public Const RXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const TXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ERRCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const RXMSGCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const TXMSGCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const LASTMSG_INVALID As String = YAPI.INVALID_STRING
    Public Const CURRENTJOB_INVALID As String = YAPI.INVALID_STRING
    Public Const STARTUPJOB_INVALID As String = YAPI.INVALID_STRING
    Public Const JOBMAXTASK_INVALID As Integer = YAPI.INVALID_UINT
    Public Const JOBMAXSIZE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    Public Const PROTOCOL_INVALID As String = YAPI.INVALID_STRING
    Public Const VOLTAGELEVEL_OFF As Integer = 0
    Public Const VOLTAGELEVEL_TTL3V As Integer = 1
    Public Const VOLTAGELEVEL_TTL3VR As Integer = 2
    Public Const VOLTAGELEVEL_TTL5V As Integer = 3
    Public Const VOLTAGELEVEL_TTL5VR As Integer = 4
    Public Const VOLTAGELEVEL_RS232 As Integer = 5
    Public Const VOLTAGELEVEL_RS485 As Integer = 6
    Public Const VOLTAGELEVEL_TTL1V8 As Integer = 7
    Public Const VOLTAGELEVEL_SDI12 As Integer = 8
    Public Const VOLTAGELEVEL_INVALID As Integer = -1
    Public Const SERIALMODE_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of generated code: YSdi12Port definitions)

    REM --- (generated code: YSdi12Port attributes declaration)
    Protected _rxCount As Integer
    Protected _txCount As Integer
    Protected _errCount As Integer
    Protected _rxMsgCount As Integer
    Protected _txMsgCount As Integer
    Protected _lastMsg As String
    Protected _currentJob As String
    Protected _startupJob As String
    Protected _jobMaxTask As Integer
    Protected _jobMaxSize As Integer
    Protected _command As String
    Protected _protocol As String
    Protected _voltageLevel As Integer
    Protected _serialMode As String
    Protected _valueCallbackSdi12Port As YSdi12PortValueCallback
    Protected _rxptr As Integer
    Protected _rxbuff As Byte()
    Protected _rxbuffptr As Integer
    Protected _eventPos As Integer
    REM --- (end of generated code: YSdi12Port attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Sdi12Port"
      REM --- (generated code: YSdi12Port attributes initialization)
      _rxCount = RXCOUNT_INVALID
      _txCount = TXCOUNT_INVALID
      _errCount = ERRCOUNT_INVALID
      _rxMsgCount = RXMSGCOUNT_INVALID
      _txMsgCount = TXMSGCOUNT_INVALID
      _lastMsg = LASTMSG_INVALID
      _currentJob = CURRENTJOB_INVALID
      _startupJob = STARTUPJOB_INVALID
      _jobMaxTask = JOBMAXTASK_INVALID
      _jobMaxSize = JOBMAXSIZE_INVALID
      _command = COMMAND_INVALID
      _protocol = PROTOCOL_INVALID
      _voltageLevel = VOLTAGELEVEL_INVALID
      _serialMode = SERIALMODE_INVALID
      _valueCallbackSdi12Port = Nothing
      _rxptr = 0
      _rxbuff = New Byte(){}
      _rxbuffptr = 0
      _eventPos = 0
      REM --- (end of generated code: YSdi12Port attributes initialization)
    End Sub

    REM --- (generated code: YSdi12Port private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("rxCount") Then
        _rxCount = CInt(json_val.getLong("rxCount"))
      End If
      If json_val.has("txCount") Then
        _txCount = CInt(json_val.getLong("txCount"))
      End If
      If json_val.has("errCount") Then
        _errCount = CInt(json_val.getLong("errCount"))
      End If
      If json_val.has("rxMsgCount") Then
        _rxMsgCount = CInt(json_val.getLong("rxMsgCount"))
      End If
      If json_val.has("txMsgCount") Then
        _txMsgCount = CInt(json_val.getLong("txMsgCount"))
      End If
      If json_val.has("lastMsg") Then
        _lastMsg = json_val.getString("lastMsg")
      End If
      If json_val.has("currentJob") Then
        _currentJob = json_val.getString("currentJob")
      End If
      If json_val.has("startupJob") Then
        _startupJob = json_val.getString("startupJob")
      End If
      If json_val.has("jobMaxTask") Then
        _jobMaxTask = CInt(json_val.getLong("jobMaxTask"))
      End If
      If json_val.has("jobMaxSize") Then
        _jobMaxSize = CInt(json_val.getLong("jobMaxSize"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      If json_val.has("protocol") Then
        _protocol = json_val.getString("protocol")
      End If
      If json_val.has("voltageLevel") Then
        _voltageLevel = CInt(json_val.getLong("voltageLevel"))
      End If
      If json_val.has("serialMode") Then
        _serialMode = json_val.getString("serialMode")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of generated code: YSdi12Port private methods declaration)

    REM --- (generated code: YSdi12Port public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the total number of bytes received since last reset.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the total number of bytes received since last reset
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.RXCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rxCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RXCOUNT_INVALID
        End If
      End If
      res = Me._rxCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the total number of bytes transmitted since last reset.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the total number of bytes transmitted since last reset
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.TXCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_txCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TXCOUNT_INVALID
        End If
      End If
      res = Me._txCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the total number of communication errors detected since last reset.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the total number of communication errors detected since last reset
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.ERRCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_errCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ERRCOUNT_INVALID
        End If
      End If
      res = Me._errCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the total number of messages received since last reset.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the total number of messages received since last reset
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.RXMSGCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rxMsgCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RXMSGCOUNT_INVALID
        End If
      End If
      res = Me._rxMsgCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the total number of messages send since last reset.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the total number of messages send since last reset
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.TXMSGCOUNT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_txMsgCount() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TXMSGCOUNT_INVALID
        End If
      End If
      res = Me._txMsgCount
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the latest message fully received.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest message fully received
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.LASTMSG_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastMsg() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LASTMSG_INVALID
        End If
      End If
      res = Me._lastMsg
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the name of the job file currently in use.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the name of the job file currently in use
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.CURRENTJOB_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentJob() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CURRENTJOB_INVALID
        End If
      End If
      res = Me._currentJob
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Selects a job file to run immediately.
    ''' <para>
    '''   If an empty string is
    '''   given as argument, stops running current job file.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string
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
    Public Function set_currentJob(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("currentJob", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the job file to use when the device is powered on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the job file to use when the device is powered on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.STARTUPJOB_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_startupJob() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return STARTUPJOB_INVALID
        End If
      End If
      res = Me._startupJob
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the job to use when the device is powered on.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the job to use when the device is powered on
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
    Public Function set_startupJob(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("startupJob", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the maximum number of tasks in a job that the device can handle.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximum number of tasks in a job that the device can handle
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.JOBMAXTASK_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_jobMaxTask() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return JOBMAXTASK_INVALID
        End If
      End If
      res = Me._jobMaxTask
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns maximum size allowed for job files.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to maximum size allowed for job files
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.JOBMAXSIZE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_jobMaxSize() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return JOBMAXSIZE_INVALID
        End If
      End If
      res = Me._jobMaxSize
      Return res
    End Function

    Public Function get_command() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      res = Me._command
      Return res
    End Function


    Public Function set_command(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("command", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the type of protocol used over the serial line, as a string.
    ''' <para>
    '''   Possible values are "Line" for ASCII messages separated by CR and/or LF,
    '''   "Frame:[timeout]ms" for binary messages separated by a delay time,
    '''   "Char" for a continuous ASCII stream or
    '''   "Byte" for a continuous binary stream.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the type of protocol used over the serial line, as a string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.PROTOCOL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_protocol() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PROTOCOL_INVALID
        End If
      End If
      res = Me._protocol
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the type of protocol used over the serial line.
    ''' <para>
    '''   Possible values are "Line" for ASCII messages separated by CR and/or LF,
    '''   "Frame:[timeout]ms" for binary messages separated by a delay time,
    '''   "Char" for a continuous ASCII stream or
    '''   "Byte" for a continuous binary stream.
    '''   The suffix "/[wait]ms" can be added to reduce the transmit rate so that there
    '''   is always at lest the specified number of milliseconds between each bytes sent.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the type of protocol used over the serial line
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
    Public Function set_protocol(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("protocol", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the voltage level used on the serial line.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YSdi12Port.VOLTAGELEVEL_OFF</c>, <c>YSdi12Port.VOLTAGELEVEL_TTL3V</c>,
    '''   <c>YSdi12Port.VOLTAGELEVEL_TTL3VR</c>, <c>YSdi12Port.VOLTAGELEVEL_TTL5V</c>,
    '''   <c>YSdi12Port.VOLTAGELEVEL_TTL5VR</c>, <c>YSdi12Port.VOLTAGELEVEL_RS232</c>,
    '''   <c>YSdi12Port.VOLTAGELEVEL_RS485</c>, <c>YSdi12Port.VOLTAGELEVEL_TTL1V8</c> and
    '''   <c>YSdi12Port.VOLTAGELEVEL_SDI12</c> corresponding to the voltage level used on the serial line
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.VOLTAGELEVEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_voltageLevel() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return VOLTAGELEVEL_INVALID
        End If
      End If
      res = Me._voltageLevel
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the voltage type used on the serial line.
    ''' <para>
    '''   Valid
    '''   values  will depend on the Yoctopuce device model featuring
    '''   the serial port feature.  Check your device documentation
    '''   to find out which values are valid for that specific model.
    '''   Trying to set an invalid value will have no effect.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YSdi12Port.VOLTAGELEVEL_OFF</c>, <c>YSdi12Port.VOLTAGELEVEL_TTL3V</c>,
    '''   <c>YSdi12Port.VOLTAGELEVEL_TTL3VR</c>, <c>YSdi12Port.VOLTAGELEVEL_TTL5V</c>,
    '''   <c>YSdi12Port.VOLTAGELEVEL_TTL5VR</c>, <c>YSdi12Port.VOLTAGELEVEL_RS232</c>,
    '''   <c>YSdi12Port.VOLTAGELEVEL_RS485</c>, <c>YSdi12Port.VOLTAGELEVEL_TTL1V8</c> and
    '''   <c>YSdi12Port.VOLTAGELEVEL_SDI12</c> corresponding to the voltage type used on the serial line
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
    Public Function set_voltageLevel(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("voltageLevel", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the serial port communication parameters, as a string such as
    '''   "1200,7E1,Simplex".
    ''' <para>
    '''   The string includes the baud rate, the number of data bits,
    '''   the parity, and the number of stop bits. The suffix "Simplex" denotes
    '''   the fact that transmission in both directions is multiplexed on the
    '''   same transmission line.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the serial port communication parameters, as a string such as
    '''   "1200,7E1,Simplex"
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSdi12Port.SERIALMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_serialMode() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SERIALMODE_INVALID
        End If
      End If
      res = Me._serialMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the serial port communication parameters, with a string such as
    '''   "1200,7E1,Simplex".
    ''' <para>
    '''   The string includes the baud rate, the number of data bits,
    '''   the parity, and the number of stop bits. The suffix "Simplex" denotes
    '''   the fact that transmission in both directions is multiplexed on the
    '''   same transmission line.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the serial port communication parameters, with a string such as
    '''   "1200,7E1,Simplex"
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
    Public Function set_serialMode(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("serialMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves an SDI12 port for a given identifier.
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
    '''   This function does not require that the SDI12 port is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSdi12Port.isOnline()</c> to test if the SDI12 port is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an SDI12 port by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the SDI12 port, for instance
    '''   <c>MyDevice.sdi12Port</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YSdi12Port</c> object allowing you to drive the SDI12 port.
    ''' </returns>
    '''/
    Public Shared Function FindSdi12Port(func As String) As YSdi12Port
      Dim obj As YSdi12Port
      obj = CType(YFunction._FindFromCache("Sdi12Port", func), YSdi12Port)
      If ((obj Is Nothing)) Then
        obj = New YSdi12Port(func)
        YFunction._AddToCache("Sdi12Port", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YSdi12PortValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSdi12Port = callback
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
      If (Not (Me._valueCallbackSdi12Port Is Nothing)) Then
        Me._valueCallbackSdi12Port(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function sendCommand(text As String) As Integer
      Return Me.set_command(text)
    End Function

    '''*
    ''' <summary>
    '''   Reads a single line (or message) from the receive buffer, starting at current stream position.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for a message protocol,
    '''   such as 'Line' mode or frame protocols.
    ''' </para>
    ''' <para>
    '''   If data at current stream position is not available anymore in the receive buffer,
    '''   the function returns the oldest available line and moves the stream position just after.
    '''   If no new full line is received, the function returns an empty line.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string with a single line of text
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function readLine() As String
      Dim url As String
      Dim msgbin As Byte() = New Byte(){}
      Dim msgarr As List(Of Byte()) = New List(Of Byte())()
      Dim msglen As Integer = 0
      Dim res As String

      url = "rxmsg.json?pos=" + Convert.ToString(Me._rxptr) + "&len=1&maxw=1"
      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return ""
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = Me._decode_json_int(msgarr(msglen))
      If (msglen = 0) Then
        Return ""
      End If
      res = Me._json_get_string(msgarr(0))
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Searches for incoming messages in the serial port receive buffer matching a given pattern,
    '''   starting at current position.
    ''' <para>
    '''   This function will only compare and return printable characters
    '''   in the message strings. Binary protocols are handled as hexadecimal strings.
    ''' </para>
    ''' <para>
    '''   The search returns all messages matching the expression provided as argument in the buffer.
    '''   If no matching message is found, the search waits for one up to the specified maximum timeout
    '''   (in milliseconds).
    ''' </para>
    ''' </summary>
    ''' <param name="pattern">
    '''   a limited regular expression describing the expected message format,
    '''   or an empty string if all messages should be returned (no filtering).
    '''   When using binary protocols, the format applies to the hexadecimal
    '''   representation of the message.
    ''' </param>
    ''' <param name="maxWait">
    '''   the maximum number of milliseconds to wait for a message if none is found
    '''   in the receive buffer.
    ''' </param>
    ''' <returns>
    '''   an array of strings containing the messages found, if any.
    '''   Binary messages are converted to hexadecimal representation.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function readMessages(pattern As String, maxWait As Integer) As List(Of String)
      Dim url As String
      Dim msgbin As Byte() = New Byte(){}
      Dim msgarr As List(Of Byte()) = New List(Of Byte())()
      Dim msglen As Integer = 0
      Dim res As List(Of String) = New List(Of String)()
      Dim idx As Integer = 0

      url = "rxmsg.json?pos=" + Convert.ToString(Me._rxptr) + "&maxw=" + Convert.ToString(maxWait) + "&pat=" + pattern
      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return res
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = Me._decode_json_int(msgarr(msglen))
      idx = 0

      While (idx < msglen)
        res.Add(Me._json_get_string(msgarr(idx)))
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Changes the current internal stream position to the specified value.
    ''' <para>
    '''   This function
    '''   does not affect the device, it only changes the value stored in the API object
    '''   for the next read operations.
    ''' </para>
    ''' </summary>
    ''' <param name="absPos">
    '''   the absolute position index for next read operations.
    ''' </param>
    ''' <returns>
    '''   nothing.
    ''' </returns>
    '''/
    Public Overridable Function read_seek(absPos As Integer) As Integer
      Me._rxptr = absPos
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Returns the current absolute stream position pointer of the API object.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the absolute position index for next read operations.
    ''' </returns>
    '''/
    Public Overridable Function read_tell() As Integer
      Return Me._rxptr
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of bytes available to read in the input buffer starting from the
    '''   current absolute stream position pointer of the API object.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the number of bytes available to read
    ''' </returns>
    '''/
    Public Overridable Function read_avail() As Integer
      Dim availPosStr As String
      Dim atPos As Integer = 0
      Dim res As Integer = 0
      Dim databin As Byte() = New Byte(){}

      databin = Me._download("rxcnt.bin?pos=" + Convert.ToString(Me._rxptr))
      availPosStr = YAPI.DefaultEncoding.GetString(databin)
      atPos = availPosStr.IndexOf("@")
      res = YAPI._atoi((availPosStr).Substring(0, atPos))
      Return res
    End Function

    Public Overridable Function end_tell() As Integer
      Dim availPosStr As String
      Dim atPos As Integer = 0
      Dim res As Integer = 0
      Dim databin As Byte() = New Byte(){}

      databin = Me._download("rxcnt.bin?pos=" + Convert.ToString(Me._rxptr))
      availPosStr = YAPI.DefaultEncoding.GetString(databin)
      atPos = availPosStr.IndexOf("@")
      res = YAPI._atoi((availPosStr).Substring(atPos+1, (availPosStr).Length-atPos-1))
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends a text line query to the serial port, and reads the reply, if any.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for 'Line' protocol.
    ''' </para>
    ''' </summary>
    ''' <param name="query">
    '''   the line query to send (without CR/LF)
    ''' </param>
    ''' <param name="maxWait">
    '''   the maximum number of milliseconds to wait for a reply.
    ''' </param>
    ''' <returns>
    '''   the next text line received after sending the text query, as a string.
    '''   Additional lines can be obtained by calling readLine or readMessages.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function queryLine(query As String, maxWait As Integer) As String
      Dim prevpos As Integer = 0
      Dim url As String
      Dim msgbin As Byte() = New Byte(){}
      Dim msgarr As List(Of Byte()) = New List(Of Byte())()
      Dim msglen As Integer = 0
      Dim res As String
      If ((query).Length <= 80) Then
        REM // fast query
        url = "rxmsg.json?len=1&maxw=" + Convert.ToString(maxWait) + "&cmd=!" + Me._escapeAttr(query)
      Else
        REM // long query
        prevpos = Me.end_tell()
        Me._upload("txdata", YAPI.DefaultEncoding.GetBytes(query + "" + vbCr + "" + vbLf + ""))
        url = "rxmsg.json?len=1&maxw=" + Convert.ToString(maxWait) + "&pos=" + Convert.ToString(prevpos)
      End If

      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return ""
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = Me._decode_json_int(msgarr(msglen))
      If (msglen = 0) Then
        Return ""
      End If
      res = Me._json_get_string(msgarr(0))
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends a binary message to the serial port, and reads the reply, if any.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for
    '''   Frame-based protocol.
    ''' </para>
    ''' </summary>
    ''' <param name="hexString">
    '''   the message to send, coded in hexadecimal
    ''' </param>
    ''' <param name="maxWait">
    '''   the maximum number of milliseconds to wait for a reply.
    ''' </param>
    ''' <returns>
    '''   the next frame received after sending the message, as a hex string.
    '''   Additional frames can be obtained by calling readHex or readMessages.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function queryHex(hexString As String, maxWait As Integer) As String
      Dim prevpos As Integer = 0
      Dim url As String
      Dim msgbin As Byte() = New Byte(){}
      Dim msgarr As List(Of Byte()) = New List(Of Byte())()
      Dim msglen As Integer = 0
      Dim res As String
      If ((hexString).Length <= 80) Then
        REM // fast query
        url = "rxmsg.json?len=1&maxw=" + Convert.ToString(maxWait) + "&cmd=$" + hexString
      Else
        REM // long query
        prevpos = Me.end_tell()
        Me._upload("txdata", YAPI._hexStrToBin(hexString))
        url = "rxmsg.json?len=1&maxw=" + Convert.ToString(maxWait) + "&pos=" + Convert.ToString(prevpos)
      End If

      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return ""
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = Me._decode_json_int(msgarr(msglen))
      If (msglen = 0) Then
        Return ""
      End If
      res = Me._json_get_string(msgarr(0))
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Saves the job definition string (JSON data) into a job file.
    ''' <para>
    '''   The job file can be later enabled using <c>selectJob()</c>.
    ''' </para>
    ''' </summary>
    ''' <param name="jobfile">
    '''   name of the job file to save on the device filesystem
    ''' </param>
    ''' <param name="jsonDef">
    '''   a string containing a JSON definition of the job
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function uploadJob(jobfile As String, jsonDef As String) As Integer
      Me._upload(jobfile, YAPI.DefaultEncoding.GetBytes(jsonDef))
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Load and start processing the specified job file.
    ''' <para>
    '''   The file must have
    '''   been previously created using the user interface or uploaded on the
    '''   device filesystem using the <c>uploadJob()</c> function.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="jobfile">
    '''   name of the job file (on the device filesystem)
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function selectJob(jobfile As String) As Integer
      Return Me.set_currentJob(jobfile)
    End Function

    '''*
    ''' <summary>
    '''   Clears the serial port buffer and resets counters to zero.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function reset() As Integer
      Me._eventPos = 0
      Me._rxptr = 0
      Me._rxbuffptr = 0
      ReDim Me._rxbuff(0-1)

      Return Me.sendCommand("Z")
    End Function

    '''*
    ''' <summary>
    '''   Sends a single byte to the serial port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="code">
    '''   the byte to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeByte(code As Integer) As Integer
      Return Me.sendCommand("$" + (code).ToString("X02"))
    End Function

    '''*
    ''' <summary>
    '''   Sends an ASCII string to the serial port, as is.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="text">
    '''   the text string to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeStr(text As String) As Integer
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim idx As Integer = 0
      Dim ch As Integer = 0
      buff = YAPI.DefaultEncoding.GetBytes(text)
      bufflen = (buff).Length
      If (bufflen < 100) Then
        REM // if string is pure text, we can send it as a simple command (faster)
        ch = &H20
        idx = 0
        While ((idx < bufflen) AndAlso (ch <> 0))
          ch = buff(idx)
          If ((ch >= &H20) AndAlso (ch < &H7f)) Then
            idx = idx + 1
          Else
            ch = 0
          End If
        End While
        If (idx >= bufflen) Then
          Return Me.sendCommand("+" + text)
        End If
      End If
      REM // send string using file upload
      Return Me._upload("txdata", buff)
    End Function

    '''*
    ''' <summary>
    '''   Sends a binary buffer to the serial port, as is.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="buff">
    '''   the binary buffer to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeBin(buff As Byte()) As Integer
      Return Me._upload("txdata", buff)
    End Function

    '''*
    ''' <summary>
    '''   Sends a byte sequence (provided as a list of bytes) to the serial port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="byteList">
    '''   a list of byte codes
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeArray(byteList As List(Of Integer)) As Integer
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim idx As Integer = 0
      Dim hexb As Integer = 0
      Dim res As Integer = 0
      bufflen = byteList.Count
      ReDim buff(bufflen-1)
      idx = 0
      While (idx < bufflen)
        hexb = byteList(idx)
        buff(idx) = Convert.ToByte(hexb And &HFF)
        idx = idx + 1
      End While

      res = Me._upload("txdata", buff)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends a byte sequence (provided as a hexadecimal string) to the serial port.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="hexString">
    '''   a string of hexadecimal byte codes
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeHex(hexString As String) As Integer
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim idx As Integer = 0
      Dim hexb As Integer = 0
      Dim res As Integer = 0
      bufflen = (hexString).Length
      If (bufflen < 100) Then
        Return Me.sendCommand("$" + hexString)
      End If
      bufflen = (bufflen >> 1)
      ReDim buff(bufflen-1)
      idx = 0
      While (idx < bufflen)
        hexb = Convert.ToInt32((hexString).Substring(2 * idx, 2), 16)
        buff(idx) = Convert.ToByte(hexb And &HFF)
        idx = idx + 1
      End While

      res = Me._upload("txdata", buff)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends an ASCII string to the serial port, followed by a line break (CR LF).
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="text">
    '''   the text string to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeLine(text As String) As Integer
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim idx As Integer = 0
      Dim ch As Integer = 0
      buff = YAPI.DefaultEncoding.GetBytes("" + text + "" + vbCr + "" + vbLf + "")
      bufflen = (buff).Length-2
      If (bufflen < 100) Then
        REM // if string is pure text, we can send it as a simple command (faster)
        ch = &H20
        idx = 0
        While ((idx < bufflen) AndAlso (ch <> 0))
          ch = buff(idx)
          If ((ch >= &H20) AndAlso (ch < &H7f)) Then
            idx = idx + 1
          Else
            ch = 0
          End If
        End While
        If (idx >= bufflen) Then
          Return Me.sendCommand("!" + text)
        End If
      End If
      REM // send string using file upload
      Return Me._upload("txdata", buff)
    End Function

    '''*
    ''' <summary>
    '''   Reads one byte from the receive buffer, starting at current stream position.
    ''' <para>
    '''   If data at current stream position is not available anymore in the receive buffer,
    '''   or if there is no data available yet, the function returns YAPI.NO_MORE_DATA.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the next byte
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function readByte() As Integer
      Dim currpos As Integer = 0
      Dim reqlen As Integer = 0
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim res As Integer = 0
      REM // first check if we have the requested character in the look-ahead buffer
      bufflen = (Me._rxbuff).Length
      If ((Me._rxptr >= Me._rxbuffptr) AndAlso (Me._rxptr < Me._rxbuffptr+bufflen)) Then
        res = Me._rxbuff(Me._rxptr-Me._rxbuffptr)
        Me._rxptr = Me._rxptr + 1
        Return res
      End If
      REM // try to preload more than one byte to speed-up byte-per-byte access
      currpos = Me._rxptr
      reqlen = 1024
      buff = Me.readBin(reqlen)
      bufflen = (buff).Length
      If ((bufflen > 0) AndAlso (Me._rxptr = currpos+bufflen)) Then
        REM // up to 1024 bytes in buffer, all in direction Rx
        res = buff(0)
        Me._rxptr = currpos+1
        Me._rxbuffptr = currpos
        Me._rxbuff = buff
        Return res
      End If
      REM // mixed bidirectional data, retry with a smaller block
      Me._rxptr = currpos
      reqlen = 16
      buff = Me.readBin(reqlen)
      bufflen = (buff).Length
      If ((bufflen > 0) AndAlso (Me._rxptr = currpos+bufflen)) Then
        REM // up to 16 bytes in buffer, all in direction Rx
        res = buff(0)
        Me._rxptr = currpos+1
        Me._rxbuffptr = currpos
        Me._rxbuff = buff
        Return res
      End If
      REM // still mixed, need to process character by character
      Me._rxptr = currpos

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString(Me._rxptr) + "&len=1")
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) AndAlso (buff(bufflen) <> 64))
        endpos = endpos + mult * (buff(bufflen) - 48)
        mult = mult * 10
        bufflen = bufflen - 1
      End While
      Me._rxptr = endpos
      If (bufflen = 0) Then
        Return YAPI.NO_MORE_DATA
      End If
      res = buff(0)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads data from the receive buffer as a string, starting at current stream position.
    ''' <para>
    '''   If data at current stream position is not available anymore in the receive buffer, the
    '''   function performs a short read.
    ''' </para>
    ''' </summary>
    ''' <param name="nChars">
    '''   the maximum number of characters to read
    ''' </param>
    ''' <returns>
    '''   a string with receive buffer contents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function readStr(nChars As Integer) As String
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim res As String
      If (nChars > 65535) Then
        nChars = 65535
      End If

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString(Me._rxptr) + "&len=" + Convert.ToString(nChars))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) AndAlso (buff(bufflen) <> 64))
        endpos = endpos + mult * (buff(bufflen) - 48)
        mult = mult * 10
        bufflen = bufflen - 1
      End While
      Me._rxptr = endpos
      res = (YAPI.DefaultEncoding.GetString(buff)).Substring(0, bufflen)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads data from the receive buffer as a binary buffer, starting at current stream position.
    ''' <para>
    '''   If data at current stream position is not available anymore in the receive buffer, the
    '''   function performs a short read.
    ''' </para>
    ''' </summary>
    ''' <param name="nChars">
    '''   the maximum number of bytes to read
    ''' </param>
    ''' <returns>
    '''   a binary object with receive buffer contents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function readBin(nChars As Integer) As Byte()
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim idx As Integer = 0
      Dim res As Byte() = New Byte(){}
      If (nChars > 65535) Then
        nChars = 65535
      End If

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString(Me._rxptr) + "&len=" + Convert.ToString(nChars))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) AndAlso (buff(bufflen) <> 64))
        endpos = endpos + mult * (buff(bufflen) - 48)
        mult = mult * 10
        bufflen = bufflen - 1
      End While
      Me._rxptr = endpos
      ReDim res(bufflen-1)
      idx = 0
      While (idx < bufflen)
        res(idx) = Convert.ToByte(buff(idx) And &HFF)
        idx = idx + 1
      End While
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads data from the receive buffer as a list of bytes, starting at current stream position.
    ''' <para>
    '''   If data at current stream position is not available anymore in the receive buffer, the
    '''   function performs a short read.
    ''' </para>
    ''' </summary>
    ''' <param name="nChars">
    '''   the maximum number of bytes to read
    ''' </param>
    ''' <returns>
    '''   a sequence of bytes with receive buffer contents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function readArray(nChars As Integer) As List(Of Integer)
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim idx As Integer = 0
      Dim b As Integer = 0
      Dim res As List(Of Integer) = New List(Of Integer)()
      If (nChars > 65535) Then
        nChars = 65535
      End If

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString(Me._rxptr) + "&len=" + Convert.ToString(nChars))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) AndAlso (buff(bufflen) <> 64))
        endpos = endpos + mult * (buff(bufflen) - 48)
        mult = mult * 10
        bufflen = bufflen - 1
      End While
      Me._rxptr = endpos
      res.Clear()
      idx = 0
      While (idx < bufflen)
        b = buff(idx)
        res.Add(b)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads data from the receive buffer as a hexadecimal string, starting at current stream position.
    ''' <para>
    '''   If data at current stream position is not available anymore in the receive buffer, the
    '''   function performs a short read.
    ''' </para>
    ''' </summary>
    ''' <param name="nBytes">
    '''   the maximum number of bytes to read
    ''' </param>
    ''' <returns>
    '''   a string with receive buffer contents, encoded in hexadecimal
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function readHex(nBytes As Integer) As String
      Dim buff As Byte() = New Byte(){}
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim ofs As Integer = 0
      Dim res As String
      If (nBytes > 65535) Then
        nBytes = 65535
      End If

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString(Me._rxptr) + "&len=" + Convert.ToString(nBytes))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) AndAlso (buff(bufflen) <> 64))
        endpos = endpos + mult * (buff(bufflen) - 48)
        mult = mult * 10
        bufflen = bufflen - 1
      End While
      Me._rxptr = endpos
      res = ""
      ofs = 0
      While (ofs + 3 < bufflen)
        res = "" + res + "" + (buff(ofs)).ToString("X02") + "" + (buff(ofs + 1)).ToString("X02") + "" + (buff(ofs + 2)).ToString("X02") + "" + (buff(ofs + 3)).ToString("X02")
        ofs = ofs + 4
      End While
      While (ofs < bufflen)
        res = "" + res + "" + (buff(ofs)).ToString("X02")
        ofs = ofs + 1
      End While
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends a SDI-12 query to the bus, and reads the sensor immediate reply.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for 'SDI-12' protocol.
    ''' </para>
    ''' </summary>
    ''' <param name="sensorAddr">
    '''   the sensor address, as a string
    ''' </param>
    ''' <param name="cmd">
    '''   the SDI12 query to send (without address and exclamation point)
    ''' </param>
    ''' <param name="maxWait">
    '''   the maximum timeout to wait for a reply from sensor, in millisecond
    ''' </param>
    ''' <returns>
    '''   the reply returned by the sensor, without newline, as a string.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function querySdi12(sensorAddr As String, cmd As String, maxWait As Integer) As String
      Dim fullCmd As String
      Dim cmdChar As String
      Dim pattern As String
      Dim url As String
      Dim msgbin As Byte() = New Byte(){}
      Dim msgarr As List(Of Byte()) = New List(Of Byte())()
      Dim msglen As Integer = 0
      Dim res As String
      cmdChar  = ""

      pattern = sensorAddr
      If ((cmd).Length > 0) Then
        cmdChar = (cmd).Substring(0, 1)
      End If
      If (sensorAddr = "?") Then
        pattern = ".*"
      Else
        If (cmdChar = "M" OrElse cmdChar = "D") Then
          pattern = "" + sensorAddr + ":.*"
        Else
          pattern = "" + sensorAddr + ".*"
        End If
      End If
      pattern = Me._escapeAttr(pattern)
      fullCmd = Me._escapeAttr("+" + sensorAddr + "" + cmd + "!")
      url = "rxmsg.json?len=1&maxw=" + Convert.ToString(maxWait) + "&cmd=" + fullCmd + "&pat=" + pattern

      msgbin = Me._download(url)
      If ((msgbin).Length<2) Then
        Return ""
      End If
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return ""
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = Me._decode_json_int(msgarr(msglen))
      If (msglen = 0) Then
        Return ""
      End If
      res = Me._json_get_string(msgarr(0))
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends a discovery command to the bus, and reads the sensor information reply.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for 'SDI-12' protocol.
    '''   This function work when only one sensor is connected.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   the reply returned by the sensor, as a YSdi12SensorInfo object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function discoverSingleSensor() As YSdi12SensorInfo
      Dim resStr As String

      resStr = Me.querySdi12("?","",5000)
      If (resStr = "") Then
        Return New YSdi12SensorInfo(Me, "ERSensor Not Found")
      End If

      Return Me.getSensorInformation(resStr)
    End Function

    '''*
    ''' <summary>
    '''   Sends a discovery command to the bus, and reads all sensors information reply.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for 'SDI-12' protocol.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   all the information from every connected sensor, as an array of YSdi12SensorInfo object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function discoverAllSensors() As List(Of YSdi12SensorInfo)
      Dim sensors As List(Of YSdi12SensorInfo) = New List(Of YSdi12SensorInfo)()
      Dim idSens As List(Of String) = New List(Of String)()
      Dim res As String
      Dim i As Integer = 0
      Dim lettreMin As String
      Dim lettreMaj As String

      REM // 1. Search for sensors present
      idSens.Clear()
      i = 0
      While (i < 10)
        res = Me.querySdi12((i).ToString(),"!",500)
        If ((res).Length >= 1) Then
          idSens.Add(res)
        End If
        i = i+1
      End While
      lettreMin = "abcdefghijklmnopqrstuvwxyz"
      lettreMaj = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      i = 0
      While (i<26)
        res = Me.querySdi12((lettreMin).Substring(i, 1),"!",500)
        If ((res).Length >= 1) Then
          idSens.Add(res)
        End If
        i = i +1
      End While
      While (i<26)
        res = Me.querySdi12((lettreMaj).Substring(i, 1),"!",500)
        If ((res).Length >= 1) Then
          idSens.Add(res)
        End If
        i = i +1
      End While

      REM // 2. Query existing sensors information
      i = 0
      sensors.Clear()
      While (i < idSens.Count)
        sensors.Add(Me.getSensorInformation(idSens(i)))
        i = i + 1
      End While

      Return sensors
    End Function

    '''*
    ''' <summary>
    '''   Sends a mesurement command to the SDI-12 bus, and reads the sensor immediate reply.
    ''' <para>
    '''   The supported commands are:
    '''   M: Measurement start control
    '''   M1...M9: Additional measurement start command
    '''   D: Measurement reading control
    '''   This function is intended to be used when the serial port is configured for 'SDI-12' protocol.
    ''' </para>
    ''' </summary>
    ''' <param name="sensorAddr">
    '''   the sensor address, as a string
    ''' </param>
    ''' <param name="measCmd">
    '''   the SDI12 query to send (without address and exclamation point)
    ''' </param>
    ''' <param name="maxWait">
    '''   the maximum timeout to wait for a reply from sensor, in millisecond
    ''' </param>
    ''' <returns>
    '''   the reply returned by the sensor, without newline, as a list of float.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function readSensor(sensorAddr As String, measCmd As String, maxWait As Integer) As List(Of Double)
      Dim resStr As String
      Dim res As List(Of Double) = New List(Of Double)()
      Dim tab As List(Of String) = New List(Of String)()
      Dim split As List(Of String) = New List(Of String)()
      Dim i As Integer = 0
      Dim valdouble As Double = 0

      resStr = Me.querySdi12(sensorAddr,measCmd,maxWait)
      tab = New List(Of String)(resStr.Split(New Char() {","c}))
      split = New List(Of String)(tab(0).Split(New Char() {":"c}))
      If (split.Count < 2) Then
        Return res
      End If

      valdouble = YAPI._atof(split(1))
      res.Add(valdouble)
      i = 1
      While (i < tab.Count)
        valdouble = YAPI._atof(tab(i))
        res.Add(valdouble)
        i = i + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Changes the address of the selected sensor, and returns the sensor information with the new address.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for 'SDI-12' protocol.
    ''' </para>
    ''' </summary>
    ''' <param name="oldAddress">
    '''   Actual sensor address, as a string
    ''' </param>
    ''' <param name="newAddress">
    '''   New sensor address, as a string
    ''' </param>
    ''' <returns>
    '''   the sensor address and information , as a YSdi12SensorInfo object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function changeAddress(oldAddress As String, newAddress As String) As YSdi12SensorInfo
      Dim addr As YSdi12SensorInfo

      Me.querySdi12(oldAddress, "A" + newAddress,1000)
      addr = Me.getSensorInformation(newAddress)
      Return addr
    End Function

    '''*
    ''' <summary>
    '''   Sends a information command to the bus, and reads sensors information selected.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for 'SDI-12' protocol.
    ''' </para>
    ''' </summary>
    ''' <param name="sensorAddr">
    '''   Sensor address, as a string
    ''' </param>
    ''' <returns>
    '''   the reply returned by the sensor, as a YSdi12Port object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function getSensorInformation(sensorAddr As String) As YSdi12SensorInfo
      Dim res As String
      Dim sensor As YSdi12SensorInfo

      res = Me.querySdi12(sensorAddr,"I",1000)
      If (res = "") Then
        Return New YSdi12SensorInfo(Me, "ERSensor Not Found")
      End If
      sensor = New YSdi12SensorInfo(Me, res)
      sensor._queryValueInfo()
      Return sensor
    End Function

    '''*
    ''' <summary>
    '''   Sends a information command to the bus, and reads sensors information selected.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for 'SDI-12' protocol.
    ''' </para>
    ''' </summary>
    ''' <param name="sensorAddr">
    '''   Sensor address, as a string
    ''' </param>
    ''' <returns>
    '''   the reply returned by the sensor, as a YSdi12Port object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function readConcurrentMeasurements(sensorAddr As String) As List(Of Double)
      Dim res As List(Of Double) = New List(Of Double)()

      res= Me.readSensor(sensorAddr,"D",1000)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends a information command to the bus, and reads sensors information selected.
    ''' <para>
    '''   This function is intended to be used when the serial port is configured for 'SDI-12' protocol.
    ''' </para>
    ''' </summary>
    ''' <param name="sensorAddr">
    '''   Sensor address, as a string
    ''' </param>
    ''' <returns>
    '''   the reply returned by the sensor, as a YSdi12Port object.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty string.
    ''' </para>
    '''/
    Public Overridable Function requestConcurrentMeasurements(sensorAddr As String) As Integer
      Dim timewait As Integer = 0
      Dim wait As String

      wait = Me.querySdi12(sensorAddr,"C",1000)
      wait = (wait).Substring(1, 3)
      timewait = YAPI._atoi(wait) * 1000
      Return timewait
    End Function

    '''*
    ''' <summary>
    '''   Retrieves messages (both direction) in the SDI12 port buffer, starting at current position.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   If no message is found, the search waits for one up to the specified maximum timeout
    '''   (in milliseconds).
    ''' </para>
    ''' </summary>
    ''' <param name="maxWait">
    '''   the maximum number of milliseconds to wait for a message if none is found
    '''   in the receive buffer.
    ''' </param>
    ''' <param name="maxMsg">
    '''   the maximum number of messages to be returned by the function; up to 254.
    ''' </param>
    ''' <returns>
    '''   an array of <c>YSdi12SnoopingRecord</c> objects containing the messages found, if any.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function snoopMessagesEx(maxWait As Integer, maxMsg As Integer) As List(Of YSdi12SnoopingRecord)
      Dim url As String
      Dim msgbin As Byte() = New Byte(){}
      Dim msgarr As List(Of Byte()) = New List(Of Byte())()
      Dim msglen As Integer = 0
      Dim res As List(Of YSdi12SnoopingRecord) = New List(Of YSdi12SnoopingRecord)()
      Dim idx As Integer = 0

      url = "rxmsg.json?pos=" + Convert.ToString(Me._rxptr) + "&maxw=" + Convert.ToString(maxWait) + "&t=0&len=" + Convert.ToString(maxMsg)
      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return res
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = Me._decode_json_int(msgarr(msglen))
      idx = 0

      While (idx < msglen)
        res.Add(New YSdi12SnoopingRecord(YAPI.DefaultEncoding.GetString(msgarr(idx))))
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves messages (both direction) in the SDI12 port buffer, starting at current position.
    ''' <para>
    ''' </para>
    ''' <para>
    '''   If no message is found, the search waits for one up to the specified maximum timeout
    '''   (in milliseconds).
    ''' </para>
    ''' </summary>
    ''' <param name="maxWait">
    '''   the maximum number of milliseconds to wait for a message if none is found
    '''   in the receive buffer.
    ''' </param>
    ''' <returns>
    '''   an array of <c>YSdi12SnoopingRecord</c> objects containing the messages found, if any.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function snoopMessages(maxWait As Integer) As List(Of YSdi12SnoopingRecord)
      Return Me.snoopMessagesEx(maxWait, 255)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of SDI12 ports started using <c>yFirstSdi12Port()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned SDI12 ports order.
    '''   If you want to find a specific an SDI12 port, use <c>Sdi12Port.findSdi12Port()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSdi12Port</c> object, corresponding to
    '''   an SDI12 port currently online, or a <c>Nothing</c> pointer
    '''   if there are no more SDI12 ports to enumerate.
    ''' </returns>
    '''/
    Public Function nextSdi12Port() As YSdi12Port
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSdi12Port.FindSdi12Port(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of SDI12 ports currently accessible.
    ''' <para>
    '''   Use the method <c>YSdi12Port.nextSdi12Port()</c> to iterate on
    '''   next SDI12 ports.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSdi12Port</c> object, corresponding to
    '''   the first SDI12 port currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstSdi12Port() As YSdi12Port
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Sdi12Port", 0, p, size, neededsize, errmsg)
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
      Return YSdi12Port.FindSdi12Port(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YSdi12Port public methods declaration)

  End Class

  REM --- (generated code: YSdi12Port functions)

  '''*
  ''' <summary>
  '''   Retrieves an SDI12 port for a given identifier.
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
  '''   This function does not require that the SDI12 port is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSdi12Port.isOnline()</c> to test if the SDI12 port is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an SDI12 port by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the SDI12 port, for instance
  '''   <c>MyDevice.sdi12Port</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YSdi12Port</c> object allowing you to drive the SDI12 port.
  ''' </returns>
  '''/
  Public Function yFindSdi12Port(ByVal func As String) As YSdi12Port
    Return YSdi12Port.FindSdi12Port(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of SDI12 ports currently accessible.
  ''' <para>
  '''   Use the method <c>YSdi12Port.nextSdi12Port()</c> to iterate on
  '''   next SDI12 ports.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YSdi12Port</c> object, corresponding to
  '''   the first SDI12 port currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSdi12Port() As YSdi12Port
    Return YSdi12Port.FirstSdi12Port()
  End Function


  REM --- (end of generated code: YSdi12Port functions)

End Module

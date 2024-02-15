'*********************************************************************
'*
'* $Id: yocto_serialport.vb 58921 2024-01-12 09:43:57Z seb $
'*
'* Implements yFindSerialPort(), the high-level API for SerialPort functions
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
'*********************************************************************/


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_serialport

  REM --- (generated code: YSnoopingRecord globals)

  REM --- (end of generated code: YSnoopingRecord globals)


    REM --- (generated code: YSerialPort return codes)
    REM --- (end of generated code: YSerialPort return codes)
    REM --- (generated code: YSerialPort dlldef)
    REM --- (end of generated code: YSerialPort dlldef)
  REM --- (generated code: YSerialPort globals)

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
  Public Delegate Sub YSerialPortValueCallback(ByVal func As YSerialPort, ByVal value As String)
  Public Delegate Sub YSerialPortTimedReportCallback(ByVal func As YSerialPort, ByVal measure As YMeasure)
  Public Delegate Sub YSnoopingCallback(ByVal func As YSerialPort, ByVal rec As YSnoopingRecord)

  Sub yInternalEventCallback(ByVal func As YSerialPort, ByVal value As String)
    func._internalEventHandler(value)
  End Sub
  REM --- (end of generated code: YSerialPort globals)



  REM --- (generated code: YSnoopingRecord class start)

  Public Class YSnoopingRecord
    REM --- (end of generated code: YSnoopingRecord class start)
    REM --- (generated code: YSnoopingRecord definitions)
    REM --- (end of generated code: YSnoopingRecord definitions)
    REM --- (generated code: YSnoopingRecord attributes declaration)
    Protected _tim As Integer
    Protected _pos As Integer
    Protected _dir As Integer
    Protected _msg As String
    REM --- (end of generated code: YSnoopingRecord attributes declaration)

    REM --- (generated code: YSnoopingRecord private methods declaration)

    REM --- (end of generated code: YSnoopingRecord private methods declaration)

    REM --- (generated code: YSnoopingRecord public methods declaration)
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



    REM --- (end of generated code: YSnoopingRecord public methods declaration)



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




  REM --- (generated code: YSerialPort class start)

  '''*
  ''' <summary>
  '''   The <c>YSerialPort</c> class allows you to fully drive a Yoctopuce serial port.
  ''' <para>
  '''   It can be used to send and receive data, and to configure communication
  '''   parameters (baud rate, bit count, parity, flow control and protocol).
  '''   Note that Yoctopuce serial ports are not exposed as virtual COM ports.
  '''   They are meant to be used in the same way as all Yoctopuce devices.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSerialPort
    Inherits YFunction
    REM --- (end of generated code: YSerialPort class start)

    REM --- (generated code: YSerialPort definitions)
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
    REM --- (end of generated code: YSerialPort definitions)

    REM --- (generated code: YSerialPort attributes declaration)
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
    Protected _valueCallbackSerialPort As YSerialPortValueCallback
    Protected _rxptr As Integer
    Protected _rxbuff As Byte()
    Protected _rxbuffptr As Integer
    Protected _eventPos As Integer
    Protected _eventCallback As YSnoopingCallback
    REM --- (end of generated code: YSerialPort attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "SerialPort"
      REM --- (generated code: YSerialPort attributes initialization)
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
      _valueCallbackSerialPort = Nothing
      _rxptr = 0
      _rxbuff = New Byte(){}
      _rxbuffptr = 0
      _eventPos = 0
      REM --- (end of generated code: YSerialPort attributes initialization)
    End Sub

    REM --- (generated code: YSerialPort private methods declaration)

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

    REM --- (end of generated code: YSerialPort private methods declaration)

    REM --- (generated code: YSerialPort public methods declaration)
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
    '''   On failure, throws an exception or returns <c>YSerialPort.RXCOUNT_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>YSerialPort.TXCOUNT_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>YSerialPort.ERRCOUNT_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>YSerialPort.RXMSGCOUNT_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>YSerialPort.TXMSGCOUNT_INVALID</c>.
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
    '''   Returns the latest message fully received (for Line, Frame and Modbus protocols).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest message fully received (for Line, Frame and Modbus protocols)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSerialPort.LASTMSG_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>YSerialPort.CURRENTJOB_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>YSerialPort.STARTUPJOB_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>YSerialPort.JOBMAXTASK_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>YSerialPort.JOBMAXSIZE_INVALID</c>.
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
    '''   "StxEtx" for ASCII messages delimited by STX/ETX codes,
    '''   "Frame:[timeout]ms" for binary messages separated by a delay time,
    '''   "Modbus-ASCII" for MODBUS messages in ASCII mode,
    '''   "Modbus-RTU" for MODBUS messages in RTU mode,
    '''   "Wiegand-ASCII" for Wiegand messages in ASCII mode,
    '''   "Wiegand-26","Wiegand-34", etc for Wiegand messages in byte mode,
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
    '''   On failure, throws an exception or returns <c>YSerialPort.PROTOCOL_INVALID</c>.
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
    '''   "StxEtx" for ASCII messages delimited by STX/ETX codes,
    '''   "Frame:[timeout]ms" for binary messages separated by a delay time,
    '''   "Modbus-ASCII" for MODBUS messages in ASCII mode,
    '''   "Modbus-RTU" for MODBUS messages in RTU mode,
    '''   "Wiegand-ASCII" for Wiegand messages in ASCII mode,
    '''   "Wiegand-26","Wiegand-34", etc for Wiegand messages in byte mode,
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
    '''   a value among <c>YSerialPort.VOLTAGELEVEL_OFF</c>, <c>YSerialPort.VOLTAGELEVEL_TTL3V</c>,
    '''   <c>YSerialPort.VOLTAGELEVEL_TTL3VR</c>, <c>YSerialPort.VOLTAGELEVEL_TTL5V</c>,
    '''   <c>YSerialPort.VOLTAGELEVEL_TTL5VR</c>, <c>YSerialPort.VOLTAGELEVEL_RS232</c>,
    '''   <c>YSerialPort.VOLTAGELEVEL_RS485</c>, <c>YSerialPort.VOLTAGELEVEL_TTL1V8</c> and
    '''   <c>YSerialPort.VOLTAGELEVEL_SDI12</c> corresponding to the voltage level used on the serial line
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSerialPort.VOLTAGELEVEL_INVALID</c>.
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
    '''   a value among <c>YSerialPort.VOLTAGELEVEL_OFF</c>, <c>YSerialPort.VOLTAGELEVEL_TTL3V</c>,
    '''   <c>YSerialPort.VOLTAGELEVEL_TTL3VR</c>, <c>YSerialPort.VOLTAGELEVEL_TTL5V</c>,
    '''   <c>YSerialPort.VOLTAGELEVEL_TTL5VR</c>, <c>YSerialPort.VOLTAGELEVEL_RS232</c>,
    '''   <c>YSerialPort.VOLTAGELEVEL_RS485</c>, <c>YSerialPort.VOLTAGELEVEL_TTL1V8</c> and
    '''   <c>YSerialPort.VOLTAGELEVEL_SDI12</c> corresponding to the voltage type used on the serial line
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
    '''   "9600,8N1".
    ''' <para>
    '''   The string includes the baud rate, the number of data bits,
    '''   the parity, and the number of stop bits. An optional suffix is included
    '''   if flow control is active: "CtsRts" for hardware handshake, "XOnXOff"
    '''   for logical flow control and "Simplex" for acquiring a shared bus using
    '''   the RTS line (as used by some RS485 adapters for instance).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the serial port communication parameters, as a string such as
    '''   "9600,8N1"
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YSerialPort.SERIALMODE_INVALID</c>.
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
    '''   "9600,8N1".
    ''' <para>
    '''   The string includes the baud rate, the number of data bits,
    '''   the parity, and the number of stop bits. An optional suffix can be added
    '''   to enable flow control: "CtsRts" for hardware handshake, "XOnXOff"
    '''   for logical flow control and "Simplex" for acquiring a shared bus using
    '''   the RTS line (as used by some RS485 adapters for instance).
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the serial port communication parameters, with a string such as
    '''   "9600,8N1"
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
    '''   Retrieves a serial port for a given identifier.
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
    '''   This function does not require that the serial port is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSerialPort.isOnline()</c> to test if the serial port is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a serial port by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the serial port, for instance
    '''   <c>RS232MK1.serialPort</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YSerialPort</c> object allowing you to drive the serial port.
    ''' </returns>
    '''/
    Public Shared Function FindSerialPort(func As String) As YSerialPort
      Dim obj As YSerialPort
      obj = CType(YFunction._FindFromCache("SerialPort", func), YSerialPort)
      If ((obj Is Nothing)) Then
        obj = New YSerialPort(func)
        YFunction._AddToCache("SerialPort", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YSerialPortValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSerialPort = callback
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
      If (Not (Me._valueCallbackSerialPort Is Nothing)) Then
        Me._valueCallbackSerialPort(Me, value)
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
      Dim msgarr As List(Of String) = New List(Of String)()
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
      Me._rxptr = YAPI._atoi(msgarr(msglen))
      If (msglen = 0) Then
        Return ""
      End If
      res = Me._json_get_string(YAPI.DefaultEncoding.GetBytes(msgarr(0)))
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
      Dim msgarr As List(Of String) = New List(Of String)()
      Dim msglen As Integer = 0
      Dim res As List(Of String) = New List(Of String)()
      Dim idx As Integer = 0

      url = "rxmsg.json?pos=" + Convert.ToString( Me._rxptr) + "&maxw=" + Convert.ToString( maxWait) + "&pat=" + pattern
      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return res
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = YAPI._atoi(msgarr(msglen))
      idx = 0

      While (idx < msglen)
        res.Add(Me._json_get_string(YAPI.DefaultEncoding.GetBytes(msgarr(idx))))
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
      res = YAPI._atoi((availPosStr).Substring( 0, atPos))
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
      res = YAPI._atoi((availPosStr).Substring( atPos+1, (availPosStr).Length-atPos-1))
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
      Dim msgarr As List(Of String) = New List(Of String)()
      Dim msglen As Integer = 0
      Dim res As String
      If ((query).Length <= 80) Then
        REM // fast query
        url = "rxmsg.json?len=1&maxw=" + Convert.ToString( maxWait) + "&cmd=!" + Me._escapeAttr(query)
      Else
        REM // long query
        prevpos = Me.end_tell()
        Me._upload("txdata", YAPI.DefaultEncoding.GetBytes(query + "" + vbCr + "" + vbLf + ""))
        url = "rxmsg.json?len=1&maxw=" + Convert.ToString( maxWait) + "&pos=" + Convert.ToString(prevpos)
      End If

      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return ""
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = YAPI._atoi(msgarr(msglen))
      If (msglen = 0) Then
        Return ""
      End If
      res = Me._json_get_string(YAPI.DefaultEncoding.GetBytes(msgarr(0)))
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
      Dim msgarr As List(Of String) = New List(Of String)()
      Dim msglen As Integer = 0
      Dim res As String
      If ((hexString).Length <= 80) Then
        REM // fast query
        url = "rxmsg.json?len=1&maxw=" + Convert.ToString( maxWait) + "&cmd=$" + hexString
      Else
        REM // long query
        prevpos = Me.end_tell()
        Me._upload("txdata", YAPI._hexStrToBin(hexString))
        url = "rxmsg.json?len=1&maxw=" + Convert.ToString( maxWait) + "&pos=" + Convert.ToString(prevpos)
      End If

      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return ""
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = YAPI._atoi(msgarr(msglen))
      If (msglen = 0) Then
        Return ""
      End If
      res = Me._json_get_string(YAPI.DefaultEncoding.GetBytes(msgarr(0)))
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
        buff( idx) = Convert.ToByte(hexb And &HFF)
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
      bufflen = ((bufflen) >> (1))
      ReDim buff(bufflen-1)
      idx = 0
      While (idx < bufflen)
        hexb = Convert.ToInt32((hexString).Substring( 2 * idx, 2), 16)
        buff( idx) = Convert.ToByte(hexb And &HFF)
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
      If (Me._rxptr = currpos+bufflen) Then
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
      If (Me._rxptr = currpos+bufflen) Then
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

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString( Me._rxptr) + "&len=" + Convert.ToString(nChars))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) AndAlso (buff(bufflen) <> 64))
        endpos = endpos + mult * (buff(bufflen) - 48)
        mult = mult * 10
        bufflen = bufflen - 1
      End While
      Me._rxptr = endpos
      res = (YAPI.DefaultEncoding.GetString(buff)).Substring( 0, bufflen)
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

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString( Me._rxptr) + "&len=" + Convert.ToString(nChars))
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
        res( idx) = Convert.ToByte(buff(idx) And &HFF)
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

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString( Me._rxptr) + "&len=" + Convert.ToString(nChars))
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

      buff = Me._download("rxdata.bin?pos=" + Convert.ToString( Me._rxptr) + "&len=" + Convert.ToString(nBytes))
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
        res = "" +  res + "" + ( buff(ofs)).ToString("X02") + "" + ( buff(ofs + 1)).ToString("X02") + "" + ( buff(ofs + 2)).ToString("X02") + "" + (buff(ofs + 3)).ToString("X02")
        ofs = ofs + 4
      End While
      While (ofs < bufflen)
        res = "" +  res + "" + (buff(ofs)).ToString("X02")
        ofs = ofs + 1
      End While
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Emits a BREAK condition on the serial interface.
    ''' <para>
    '''   When the specified
    '''   duration is 0, the BREAK signal will be exactly one character wide.
    '''   When the duration is between 1 and 100, the BREAK condition will
    '''   be hold for the specified number of milliseconds.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="duration">
    '''   0 for a standard BREAK, or duration between 1 and 100 ms
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function sendBreak(duration As Integer) As Integer
      Return Me.sendCommand("B" + Convert.ToString(duration))
    End Function

    '''*
    ''' <summary>
    '''   Manually sets the state of the RTS line.
    ''' <para>
    '''   This function has no effect when
    '''   hardware handshake is enabled, as the RTS line is driven automatically.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="val">
    '''   1 to turn RTS on, 0 to turn RTS off
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_RTS(val As Integer) As Integer
      Return Me.sendCommand("R" + Convert.ToString(val))
    End Function

    '''*
    ''' <summary>
    '''   Reads the level of the CTS line.
    ''' <para>
    '''   The CTS line is usually driven by
    '''   the RTS signal of the connected serial device.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   1 if the CTS line is high, 0 if the CTS line is low.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_CTS() As Integer
      Dim buff As Byte() = New Byte(){}
      Dim res As Integer = 0

      buff = Me._download("cts.txt")
      If Not((buff).Length = 1) Then
        me._throw( YAPI.IO_ERROR,  "invalid CTS reply")
        return YAPI.IO_ERROR
      end if
      res = buff(0) - 48
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves messages (both direction) in the serial port buffer, starting at current position.
    ''' <para>
    '''   This function will only compare and return printable characters in the message strings.
    '''   Binary protocols are handled as hexadecimal strings.
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
    '''   an array of <c>YSnoopingRecord</c> objects containing the messages found, if any.
    '''   Binary messages are converted to hexadecimal representation.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function snoopMessages(maxWait As Integer) As List(Of YSnoopingRecord)
      Dim url As String
      Dim msgbin As Byte() = New Byte(){}
      Dim msgarr As List(Of String) = New List(Of String)()
      Dim msglen As Integer = 0
      Dim res As List(Of YSnoopingRecord) = New List(Of YSnoopingRecord)()
      Dim idx As Integer = 0

      url = "rxmsg.json?pos=" + Convert.ToString( Me._rxptr) + "&maxw=" + Convert.ToString(maxWait) + "&t=0"
      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return res
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      Me._rxptr = YAPI._atoi(msgarr(msglen))
      idx = 0

      While (idx < msglen)
        res.Add(New YSnoopingRecord(msgarr(idx)))
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Registers a callback function to be called each time that a message is sent or
    '''   received by the serial port.
    ''' <para>
    '''   The callback is invoked only during the execution of
    '''   <c>ySleep</c> or <c>yHandleEvents</c>. This provides control over the time when
    '''   the callback is triggered. For good responsiveness, remember to call one of these
    '''   two functions periodically. To unregister a callback, pass a Nothing pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a Nothing pointer.
    '''   The callback function should take four arguments:
    '''   the <c>YSerialPort</c> object that emitted the event, and
    '''   the <c>YSnoopingRecord</c> object that describes the message
    '''   sent or received.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </param>
    '''/
    Public Overridable Function registerSnoopingCallback(callback As YSnoopingCallback) As Integer
      If (Not (callback Is Nothing)) Then
        Me.registerValueCallback(AddressOf yInternalEventCallback)
      Else
        Me.registerValueCallback(CType(Nothing, YSerialPortValueCallback))
      End If
      REM // register user callback AFTER the internal pseudo-event,
      REM // to make sure we start with future events only
      Me._eventCallback = callback
      Return 0
    End Function

    Public Overridable Function _internalEventHandler(advstr As String) As Integer
      Dim url As String
      Dim msgbin As Byte() = New Byte(){}
      Dim msgarr As List(Of String) = New List(Of String)()
      Dim msglen As Integer = 0
      Dim idx As Integer = 0
      If (Not (Not (Me._eventCallback Is Nothing))) Then
        REM // first simulated event, use it only to initialize reference values
        Me._eventPos = 0
      End If

      url = "rxmsg.json?pos=" + Convert.ToString(Me._eventPos) + "&maxw=0&t=0"
      msgbin = Me._download(url)
      msgarr = Me._json_get_array(msgbin)
      msglen = msgarr.Count
      If (msglen = 0) Then
        Return YAPI.SUCCESS
      End If
      REM // last element of array is the new position
      msglen = msglen - 1
      If (Not (Not (Me._eventCallback Is Nothing))) Then
        REM // first simulated event, use it only to initialize reference values
        Me._eventPos = YAPI._atoi(msgarr(msglen))
        Return YAPI.SUCCESS
      End If
      Me._eventPos = YAPI._atoi(msgarr(msglen))
      idx = 0
      While (idx < msglen)
        Me._eventCallback(Me, New YSnoopingRecord(msgarr(idx)))
        idx = idx + 1
      End While
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Sends an ASCII string to the serial port, preceeded with an STX code and
    '''   followed by an ETX code.
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
    Public Overridable Function writeStxEtx(text As String) As Integer
      Dim buff As Byte() = New Byte(){}
      buff = YAPI.DefaultEncoding.GetBytes("" + Chr( 2) + "" +  text + "" + Chr(3))
      REM // send string using file upload
      Return Me._upload("txdata", buff)
    End Function

    '''*
    ''' <summary>
    '''   Sends a MODBUS message (provided as a hexadecimal string) to the serial port.
    ''' <para>
    '''   The message must start with the slave address. The MODBUS CRC/LRC is
    '''   automatically added by the function. This function does not wait for a reply.
    ''' </para>
    ''' </summary>
    ''' <param name="hexString">
    '''   a hexadecimal message string, including device address but no CRC/LRC
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeMODBUS(hexString As String) As Integer
      Return Me.sendCommand(":" + hexString)
    End Function

    '''*
    ''' <summary>
    '''   Sends a message to a specified MODBUS slave connected to the serial port, and reads the
    '''   reply, if any.
    ''' <para>
    '''   The message is the PDU, provided as a vector of bytes.
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to query
    ''' </param>
    ''' <param name="pduBytes">
    '''   the message to send (PDU), as a vector of bytes. The first byte of the
    '''   PDU is the MODBUS function code.
    ''' </param>
    ''' <returns>
    '''   the received reply, as a vector of bytes.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array (or a MODBUS error reply).
    ''' </para>
    '''/
    Public Overridable Function queryMODBUS(slaveNo As Integer, pduBytes As List(Of Integer)) As List(Of Integer)
      Dim funCode As Integer = 0
      Dim nib As Integer = 0
      Dim i As Integer = 0
      Dim cmd As String
      Dim prevpos As Integer = 0
      Dim url As String
      Dim pat As String
      Dim msgs As Byte() = New Byte(){}
      Dim reps As List(Of String) = New List(Of String)()
      Dim rep As String
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim replen As Integer = 0
      Dim hexb As Integer = 0
      funCode = pduBytes(0)
      nib = ((funCode) >> (4))
      pat = "" + ( slaveNo).ToString("X02") + "[" + ( nib).ToString("X") + "" + ( (nib+8)).ToString("X") + "]" + (((funCode) And (15))).ToString("X") + ".*"
      cmd = "" + ( slaveNo).ToString("X02") + "" + (funCode).ToString("X02")
      i = 1
      While (i < pduBytes.Count)
        cmd = "" +  cmd + "" + (((pduBytes(i)) And (&Hff))).ToString("X02")
        i = i + 1
      End While
      If ((cmd).Length <= 80) Then
        REM // fast query
        url = "rxmsg.json?cmd=:" +  cmd + "&pat=:" + pat
      Else
        REM // long query
        prevpos = Me.end_tell()
        Me._upload("txdata:", YAPI._hexStrToBin(cmd))
        url = "rxmsg.json?pos=" + Convert.ToString( prevpos) + "&maxw=2000&pat=:" + pat
      End If

      msgs = Me._download(url)
      reps = Me._json_get_array(msgs)
      If Not(reps.Count > 1) Then
        me._throw( YAPI.IO_ERROR,  "no reply from MODBUS slave")
        return res
      end if
      If (reps.Count > 1) Then
        rep = Me._json_get_string(YAPI.DefaultEncoding.GetBytes(reps(0)))
        replen = (((rep).Length - 3) >> (1))
        i = 0
        While (i < replen)
          hexb = Convert.ToInt32((rep).Substring(2 * i + 3, 2), 16)
          res.Add(hexb)
          i = i + 1
        End While
        If (res(0) <> funCode) Then
          i = res(1)
          If Not(i > 1) Then
            me._throw( YAPI.NOT_SUPPORTED,  "MODBUS error: unsupported function code")
            return res
          end if
          If Not(i > 2) Then
            me._throw( YAPI.INVALID_ARGUMENT,  "MODBUS error: illegal data address")
            return res
          end if
          If Not(i > 3) Then
            me._throw( YAPI.INVALID_ARGUMENT,  "MODBUS error: illegal data value")
            return res
          end if
          If Not(i > 4) Then
            me._throw( YAPI.INVALID_ARGUMENT,  "MODBUS error: failed to execute function")
            return res
          end if
        End If
      End If
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads one or more contiguous internal bits (or coil status) from a MODBUS serial device.
    ''' <para>
    '''   This method uses the MODBUS function code 0x01 (Read Coils).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to query
    ''' </param>
    ''' <param name="pduAddr">
    '''   the relative address of the first bit/coil to read (zero-based)
    ''' </param>
    ''' <param name="nBits">
    '''   the number of bits/coils to read
    ''' </param>
    ''' <returns>
    '''   a vector of integers, each corresponding to one bit.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function modbusReadBits(slaveNo As Integer, pduAddr As Integer, nBits As Integer) As List(Of Integer)
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim bitpos As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim mask As Integer = 0

      pdu.Add(&H01)
      pdu.Add(((pduAddr) >> (8)))
      pdu.Add(((pduAddr) And (&Hff)))
      pdu.Add(((nBits) >> (8)))
      pdu.Add(((nBits) And (&Hff)))


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If

      bitpos = 0
      idx = 2
      val = reply(idx)
      mask = 1
      While (bitpos < nBits)
        If (((val) And (mask)) = 0) Then
          res.Add(0)
        Else
          res.Add(1)
        End If
        bitpos = bitpos + 1
        If (mask = &H80) Then
          idx = idx + 1
          val = reply(idx)
          mask = 1
        Else
          mask = ((mask) << (1))
        End If
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads one or more contiguous input bits (or discrete inputs) from a MODBUS serial device.
    ''' <para>
    '''   This method uses the MODBUS function code 0x02 (Read Discrete Inputs).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to query
    ''' </param>
    ''' <param name="pduAddr">
    '''   the relative address of the first bit/input to read (zero-based)
    ''' </param>
    ''' <param name="nBits">
    '''   the number of bits/inputs to read
    ''' </param>
    ''' <returns>
    '''   a vector of integers, each corresponding to one bit.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function modbusReadInputBits(slaveNo As Integer, pduAddr As Integer, nBits As Integer) As List(Of Integer)
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim bitpos As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim mask As Integer = 0

      pdu.Add(&H02)
      pdu.Add(((pduAddr) >> (8)))
      pdu.Add(((pduAddr) And (&Hff)))
      pdu.Add(((nBits) >> (8)))
      pdu.Add(((nBits) And (&Hff)))


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If

      bitpos = 0
      idx = 2
      val = reply(idx)
      mask = 1
      While (bitpos < nBits)
        If (((val) And (mask)) = 0) Then
          res.Add(0)
        Else
          res.Add(1)
        End If
        bitpos = bitpos + 1
        If (mask = &H80) Then
          idx = idx + 1
          val = reply(idx)
          mask = 1
        Else
          mask = ((mask) << (1))
        End If
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads one or more contiguous internal registers (holding registers) from a MODBUS serial device.
    ''' <para>
    '''   This method uses the MODBUS function code 0x03 (Read Holding Registers).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to query
    ''' </param>
    ''' <param name="pduAddr">
    '''   the relative address of the first holding register to read (zero-based)
    ''' </param>
    ''' <param name="nWords">
    '''   the number of holding registers to read
    ''' </param>
    ''' <returns>
    '''   a vector of integers, each corresponding to one 16-bit register value.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function modbusReadRegisters(slaveNo As Integer, pduAddr As Integer, nWords As Integer) As List(Of Integer)
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim regpos As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      If Not(nWords<=256) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "Cannot read more than 256 words")
        return res
      end if

      pdu.Add(&H03)
      pdu.Add(((pduAddr) >> (8)))
      pdu.Add(((pduAddr) And (&Hff)))
      pdu.Add(((nWords) >> (8)))
      pdu.Add(((nWords) And (&Hff)))


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If

      regpos = 0
      idx = 2
      While (regpos < nWords)
        val = ((reply(idx)) << (8))
        idx = idx + 1
        val = val + reply(idx)
        idx = idx + 1
        res.Add(val)
        regpos = regpos + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Reads one or more contiguous input registers (read-only registers) from a MODBUS serial device.
    ''' <para>
    '''   This method uses the MODBUS function code 0x04 (Read Input Registers).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to query
    ''' </param>
    ''' <param name="pduAddr">
    '''   the relative address of the first input register to read (zero-based)
    ''' </param>
    ''' <param name="nWords">
    '''   the number of input registers to read
    ''' </param>
    ''' <returns>
    '''   a vector of integers, each corresponding to one 16-bit input value.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function modbusReadInputRegisters(slaveNo As Integer, pduAddr As Integer, nWords As Integer) As List(Of Integer)
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As List(Of Integer) = New List(Of Integer)()
      Dim regpos As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0

      pdu.Add(&H04)
      pdu.Add(((pduAddr) >> (8)))
      pdu.Add(((pduAddr) And (&Hff)))
      pdu.Add(((nWords) >> (8)))
      pdu.Add(((nWords) And (&Hff)))


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If

      regpos = 0
      idx = 2
      While (regpos < nWords)
        val = ((reply(idx)) << (8))
        idx = idx + 1
        val = val + reply(idx)
        idx = idx + 1
        res.Add(val)
        regpos = regpos + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sets a single internal bit (or coil) on a MODBUS serial device.
    ''' <para>
    '''   This method uses the MODBUS function code 0x05 (Write Single Coil).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to drive
    ''' </param>
    ''' <param name="pduAddr">
    '''   the relative address of the bit/coil to set (zero-based)
    ''' </param>
    ''' <param name="value">
    '''   the value to set (0 for OFF state, non-zero for ON state)
    ''' </param>
    ''' <returns>
    '''   the number of bits/coils affected on the device (1)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns zero.
    ''' </para>
    '''/
    Public Overridable Function modbusWriteBit(slaveNo As Integer, pduAddr As Integer, value As Integer) As Integer
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As Integer = 0
      res = 0
      If (value <> 0) Then
        value = &Hff
      End If

      pdu.Add(&H05)
      pdu.Add(((pduAddr) >> (8)))
      pdu.Add(((pduAddr) And (&Hff)))
      pdu.Add(value)
      pdu.Add(&H00)


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If
      res = 1
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sets several contiguous internal bits (or coils) on a MODBUS serial device.
    ''' <para>
    '''   This method uses the MODBUS function code 0x0f (Write Multiple Coils).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to drive
    ''' </param>
    ''' <param name="pduAddr">
    '''   the relative address of the first bit/coil to set (zero-based)
    ''' </param>
    ''' <param name="bits">
    '''   the vector of bits to be set (one integer per bit)
    ''' </param>
    ''' <returns>
    '''   the number of bits/coils affected on the device
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns zero.
    ''' </para>
    '''/
    Public Overridable Function modbusWriteBits(slaveNo As Integer, pduAddr As Integer, bits As List(Of Integer)) As Integer
      Dim nBits As Integer = 0
      Dim nBytes As Integer = 0
      Dim bitpos As Integer = 0
      Dim val As Integer = 0
      Dim mask As Integer = 0
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As Integer = 0
      res = 0
      nBits = bits.Count
      nBytes = (((nBits + 7)) >> (3))

      pdu.Add(&H0f)
      pdu.Add(((pduAddr) >> (8)))
      pdu.Add(((pduAddr) And (&Hff)))
      pdu.Add(((nBits) >> (8)))
      pdu.Add(((nBits) And (&Hff)))
      pdu.Add(nBytes)
      bitpos = 0
      val = 0
      mask = 1
      While (bitpos < nBits)
        If (bits(bitpos) <> 0) Then
          val = ((val) Or (mask))
        End If
        bitpos = bitpos + 1
        If (mask = &H80) Then
          pdu.Add(val)
          val = 0
          mask = 1
        Else
          mask = ((mask) << (1))
        End If
      End While
      If (mask <> 1) Then
        pdu.Add(val)
      End If


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If
      res = ((reply(3)) << (8))
      res = res + reply(4)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sets a single internal register (or holding register) on a MODBUS serial device.
    ''' <para>
    '''   This method uses the MODBUS function code 0x06 (Write Single Register).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to drive
    ''' </param>
    ''' <param name="pduAddr">
    '''   the relative address of the register to set (zero-based)
    ''' </param>
    ''' <param name="value">
    '''   the 16 bit value to set
    ''' </param>
    ''' <returns>
    '''   the number of registers affected on the device (1)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns zero.
    ''' </para>
    '''/
    Public Overridable Function modbusWriteRegister(slaveNo As Integer, pduAddr As Integer, value As Integer) As Integer
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As Integer = 0
      res = 0

      pdu.Add(&H06)
      pdu.Add(((pduAddr) >> (8)))
      pdu.Add(((pduAddr) And (&Hff)))
      pdu.Add(((value) >> (8)))
      pdu.Add(((value) And (&Hff)))


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If
      res = 1
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sets several contiguous internal registers (or holding registers) on a MODBUS serial device.
    ''' <para>
    '''   This method uses the MODBUS function code 0x10 (Write Multiple Registers).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to drive
    ''' </param>
    ''' <param name="pduAddr">
    '''   the relative address of the first internal register to set (zero-based)
    ''' </param>
    ''' <param name="values">
    '''   the vector of 16 bit values to set
    ''' </param>
    ''' <returns>
    '''   the number of registers affected on the device
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns zero.
    ''' </para>
    '''/
    Public Overridable Function modbusWriteRegisters(slaveNo As Integer, pduAddr As Integer, values As List(Of Integer)) As Integer
      Dim nWords As Integer = 0
      Dim nBytes As Integer = 0
      Dim regpos As Integer = 0
      Dim val As Integer = 0
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As Integer = 0
      res = 0
      nWords = values.Count
      nBytes = 2 * nWords

      pdu.Add(&H10)
      pdu.Add(((pduAddr) >> (8)))
      pdu.Add(((pduAddr) And (&Hff)))
      pdu.Add(((nWords) >> (8)))
      pdu.Add(((nWords) And (&Hff)))
      pdu.Add(nBytes)
      regpos = 0
      While (regpos < nWords)
        val = values(regpos)
        pdu.Add(((val) >> (8)))
        pdu.Add(((val) And (&Hff)))
        regpos = regpos + 1
      End While


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If
      res = ((reply(3)) << (8))
      res = res + reply(4)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sets several contiguous internal registers (holding registers) on a MODBUS serial device,
    '''   then performs a contiguous read of a set of (possibly different) internal registers.
    ''' <para>
    '''   This method uses the MODBUS function code 0x17 (Read/Write Multiple Registers).
    ''' </para>
    ''' </summary>
    ''' <param name="slaveNo">
    '''   the address of the slave MODBUS device to drive
    ''' </param>
    ''' <param name="pduWriteAddr">
    '''   the relative address of the first internal register to set (zero-based)
    ''' </param>
    ''' <param name="values">
    '''   the vector of 16 bit values to set
    ''' </param>
    ''' <param name="pduReadAddr">
    '''   the relative address of the first internal register to read (zero-based)
    ''' </param>
    ''' <param name="nReadWords">
    '''   the number of 16 bit values to read
    ''' </param>
    ''' <returns>
    '''   a vector of integers, each corresponding to one 16-bit register value read.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function modbusWriteAndReadRegisters(slaveNo As Integer, pduWriteAddr As Integer, values As List(Of Integer), pduReadAddr As Integer, nReadWords As Integer) As List(Of Integer)
      Dim nWriteWords As Integer = 0
      Dim nBytes As Integer = 0
      Dim regpos As Integer = 0
      Dim val As Integer = 0
      Dim idx As Integer = 0
      Dim pdu As List(Of Integer) = New List(Of Integer)()
      Dim reply As List(Of Integer) = New List(Of Integer)()
      Dim res As List(Of Integer) = New List(Of Integer)()
      nWriteWords = values.Count
      nBytes = 2 * nWriteWords

      pdu.Add(&H17)
      pdu.Add(((pduReadAddr) >> (8)))
      pdu.Add(((pduReadAddr) And (&Hff)))
      pdu.Add(((nReadWords) >> (8)))
      pdu.Add(((nReadWords) And (&Hff)))
      pdu.Add(((pduWriteAddr) >> (8)))
      pdu.Add(((pduWriteAddr) And (&Hff)))
      pdu.Add(((nWriteWords) >> (8)))
      pdu.Add(((nWriteWords) And (&Hff)))
      pdu.Add(nBytes)
      regpos = 0
      While (regpos < nWriteWords)
        val = values(regpos)
        pdu.Add(((val) >> (8)))
        pdu.Add(((val) And (&Hff)))
        regpos = regpos + 1
      End While


      reply = Me.queryMODBUS(slaveNo, pdu)
      If (reply.Count = 0) Then
        Return res
      End If
      If (reply(0) <> pdu(0)) Then
        Return res
      End If

      regpos = 0
      idx = 2
      While (regpos < nReadWords)
        val = ((reply(idx)) << (8))
        idx = idx + 1
        val = val + reply(idx)
        idx = idx + 1
        res.Add(val)
        regpos = regpos + 1
      End While

      Return res
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of serial ports started using <c>yFirstSerialPort()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned serial ports order.
    '''   If you want to find a specific a serial port, use <c>SerialPort.findSerialPort()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSerialPort</c> object, corresponding to
    '''   a serial port currently online, or a <c>Nothing</c> pointer
    '''   if there are no more serial ports to enumerate.
    ''' </returns>
    '''/
    Public Function nextSerialPort() As YSerialPort
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSerialPort.FindSerialPort(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of serial ports currently accessible.
    ''' <para>
    '''   Use the method <c>YSerialPort.nextSerialPort()</c> to iterate on
    '''   next serial ports.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSerialPort</c> object, corresponding to
    '''   the first serial port currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstSerialPort() As YSerialPort
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("SerialPort", 0, p, size, neededsize, errmsg)
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
      Return YSerialPort.FindSerialPort(serial + "." + funcId)
    End Function

    REM --- (end of generated code: YSerialPort public methods declaration)

  End Class

  REM --- (generated code: YSerialPort functions)

  '''*
  ''' <summary>
  '''   Retrieves a serial port for a given identifier.
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
  '''   This function does not require that the serial port is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSerialPort.isOnline()</c> to test if the serial port is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a serial port by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the serial port, for instance
  '''   <c>RS232MK1.serialPort</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YSerialPort</c> object allowing you to drive the serial port.
  ''' </returns>
  '''/
  Public Function yFindSerialPort(ByVal func As String) As YSerialPort
    Return YSerialPort.FindSerialPort(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of serial ports currently accessible.
  ''' <para>
  '''   Use the method <c>YSerialPort.nextSerialPort()</c> to iterate on
  '''   next serial ports.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YSerialPort</c> object, corresponding to
  '''   the first serial port currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSerialPort() As YSerialPort
    Return YSerialPort.FirstSerialPort()
  End Function


  REM --- (end of generated code: YSerialPort functions)

End Module

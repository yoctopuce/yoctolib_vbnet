' ********************************************************************
'
'  $Id: yocto_i2cport.vb 36207 2019-07-10 20:46:18Z mvuilleu $
'
'  Implements yFindI2cPort(), the high-level API for I2cPort functions
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

Module yocto_i2cport

    REM --- (YI2cPort return codes)
    REM --- (end of YI2cPort return codes)
    REM --- (YI2cPort dlldef)
    REM --- (end of YI2cPort dlldef)
   REM --- (YI2cPort yapiwrapper)
   REM --- (end of YI2cPort yapiwrapper)
  REM --- (YI2cPort globals)

  Public Const Y_RXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_TXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ERRCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_RXMSGCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_TXMSGCOUNT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_LASTMSG_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CURRENTJOB_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_STARTUPJOB_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_VOLTAGELEVEL_OFF As Integer = 0
  Public Const Y_VOLTAGELEVEL_TTL3V As Integer = 1
  Public Const Y_VOLTAGELEVEL_TTL3VR As Integer = 2
  Public Const Y_VOLTAGELEVEL_TTL5V As Integer = 3
  Public Const Y_VOLTAGELEVEL_TTL5VR As Integer = 4
  Public Const Y_VOLTAGELEVEL_RS232 As Integer = 5
  Public Const Y_VOLTAGELEVEL_RS485 As Integer = 6
  Public Const Y_VOLTAGELEVEL_TTL1V8 As Integer = 7
  Public Const Y_VOLTAGELEVEL_INVALID As Integer = -1
  Public Const Y_PROTOCOL_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_I2CMODE_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YI2cPortValueCallback(ByVal func As YI2cPort, ByVal value As String)
  Public Delegate Sub YI2cPortTimedReportCallback(ByVal func As YI2cPort, ByVal measure As YMeasure)
  REM --- (end of YI2cPort globals)

  REM --- (YI2cPort class start)

  '''*
  ''' <summary>
  '''   The I2cPort function interface allows you to fully drive a Yoctopuce
  '''   I2C port, to send and receive data, and to configure communication
  '''   parameters (baud rate, etc).
  ''' <para>
  '''   Note that Yoctopuce I2C ports are not exposed as virtual COM ports.
  '''   They are meant to be used in the same way as all Yoctopuce devices.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YI2cPort
    Inherits YFunction
    REM --- (end of YI2cPort class start)

    REM --- (YI2cPort definitions)
    Public Const RXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const TXCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ERRCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const RXMSGCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const TXMSGCOUNT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const LASTMSG_INVALID As String = YAPI.INVALID_STRING
    Public Const CURRENTJOB_INVALID As String = YAPI.INVALID_STRING
    Public Const STARTUPJOB_INVALID As String = YAPI.INVALID_STRING
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    Public Const VOLTAGELEVEL_OFF As Integer = 0
    Public Const VOLTAGELEVEL_TTL3V As Integer = 1
    Public Const VOLTAGELEVEL_TTL3VR As Integer = 2
    Public Const VOLTAGELEVEL_TTL5V As Integer = 3
    Public Const VOLTAGELEVEL_TTL5VR As Integer = 4
    Public Const VOLTAGELEVEL_RS232 As Integer = 5
    Public Const VOLTAGELEVEL_RS485 As Integer = 6
    Public Const VOLTAGELEVEL_TTL1V8 As Integer = 7
    Public Const VOLTAGELEVEL_INVALID As Integer = -1
    Public Const PROTOCOL_INVALID As String = YAPI.INVALID_STRING
    Public Const I2CMODE_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YI2cPort definitions)

    REM --- (YI2cPort attributes declaration)
    Protected _rxCount As Integer
    Protected _txCount As Integer
    Protected _errCount As Integer
    Protected _rxMsgCount As Integer
    Protected _txMsgCount As Integer
    Protected _lastMsg As String
    Protected _currentJob As String
    Protected _startupJob As String
    Protected _command As String
    Protected _voltageLevel As Integer
    Protected _protocol As String
    Protected _i2cMode As String
    Protected _valueCallbackI2cPort As YI2cPortValueCallback
    Protected _rxptr As Integer
    Protected _rxbuff As Byte()
    Protected _rxbuffptr As Integer
    REM --- (end of YI2cPort attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "I2cPort"
      REM --- (YI2cPort attributes initialization)
      _rxCount = RXCOUNT_INVALID
      _txCount = TXCOUNT_INVALID
      _errCount = ERRCOUNT_INVALID
      _rxMsgCount = RXMSGCOUNT_INVALID
      _txMsgCount = TXMSGCOUNT_INVALID
      _lastMsg = LASTMSG_INVALID
      _currentJob = CURRENTJOB_INVALID
      _startupJob = STARTUPJOB_INVALID
      _command = COMMAND_INVALID
      _voltageLevel = VOLTAGELEVEL_INVALID
      _protocol = PROTOCOL_INVALID
      _i2cMode = I2CMODE_INVALID
      _valueCallbackI2cPort = Nothing
      _rxptr = 0
      _rxbuffptr = 0
      REM --- (end of YI2cPort attributes initialization)
    End Sub

    REM --- (YI2cPort private methods declaration)

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
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      If json_val.has("voltageLevel") Then
        _voltageLevel = CInt(json_val.getLong("voltageLevel"))
      End If
      If json_val.has("protocol") Then
        _protocol = json_val.getString("protocol")
      End If
      If json_val.has("i2cMode") Then
        _i2cMode = json_val.getString("i2cMode")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YI2cPort private methods declaration)

    REM --- (YI2cPort public methods declaration)
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
    '''   On failure, throws an exception or returns <c>Y_RXCOUNT_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>Y_TXCOUNT_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>Y_ERRCOUNT_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>Y_RXMSGCOUNT_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>Y_TXMSGCOUNT_INVALID</c>.
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
    '''   Returns the latest message fully received (for Line and Frame protocols).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the latest message fully received (for Line and Frame protocols)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LASTMSG_INVALID</c>.
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
    '''   On failure, throws an exception or returns <c>Y_CURRENTJOB_INVALID</c>.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   On failure, throws an exception or returns <c>Y_STARTUPJOB_INVALID</c>.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   Returns the voltage level used on the serial line.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_VOLTAGELEVEL_OFF</c>, <c>Y_VOLTAGELEVEL_TTL3V</c>, <c>Y_VOLTAGELEVEL_TTL3VR</c>,
    '''   <c>Y_VOLTAGELEVEL_TTL5V</c>, <c>Y_VOLTAGELEVEL_TTL5VR</c>, <c>Y_VOLTAGELEVEL_RS232</c>,
    '''   <c>Y_VOLTAGELEVEL_RS485</c> and <c>Y_VOLTAGELEVEL_TTL1V8</c> corresponding to the voltage level
    '''   used on the serial line
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VOLTAGELEVEL_INVALID</c>.
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
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_VOLTAGELEVEL_OFF</c>, <c>Y_VOLTAGELEVEL_TTL3V</c>, <c>Y_VOLTAGELEVEL_TTL3VR</c>,
    '''   <c>Y_VOLTAGELEVEL_TTL5V</c>, <c>Y_VOLTAGELEVEL_TTL5VR</c>, <c>Y_VOLTAGELEVEL_RS232</c>,
    '''   <c>Y_VOLTAGELEVEL_RS485</c> and <c>Y_VOLTAGELEVEL_TTL1V8</c> corresponding to the voltage type used
    '''   on the serial line
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
    Public Function set_voltageLevel(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("voltageLevel", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the type of protocol used over the serial line, as a string.
    ''' <para>
    '''   Possible values are
    '''   "Line" for messages separated by LF or
    '''   "Char" for continuous stream of codes.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the type of protocol used over the serial line, as a string
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PROTOCOL_INVALID</c>.
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
    '''   Possible values are
    '''   "Line" for messages separated by LF or
    '''   "Char" for continuous stream of codes.
    '''   The suffix "/[wait]ms" can be added to reduce the transmit rate so that there
    '''   is always at lest the specified number of milliseconds between each message sent.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   Returns the SPI port communication parameters, as a string such as
    '''   "400kbps,2000ms".
    ''' <para>
    '''   The string includes the baud rate and  th  e recovery delay
    '''   after communications errors.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the SPI port communication parameters, as a string such as
    '''   "400kbps,2000ms"
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_I2CMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_i2cMode() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return I2CMODE_INVALID
        End If
      End If
      res = Me._i2cMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the SPI port communication parameters, with a string such as
    '''   "400kbps,2000ms".
    ''' <para>
    '''   The string includes the baud rate and the recovery delay
    '''   after communications errors.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the SPI port communication parameters, with a string such as
    '''   "400kbps,2000ms"
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
    Public Function set_i2cMode(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("i2cMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves an I2C port for a given identifier.
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
    '''   This function does not require that the I2C port is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YI2cPort.isOnline()</c> to test if the I2C port is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an I2C port by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the I2C port
    ''' </param>
    ''' <returns>
    '''   a <c>YI2cPort</c> object allowing you to drive the I2C port.
    ''' </returns>
    '''/
    Public Shared Function FindI2cPort(func As String) As YI2cPort
      Dim obj As YI2cPort
      obj = CType(YFunction._FindFromCache("I2cPort", func), YI2cPort)
      If ((obj Is Nothing)) Then
        obj = New YI2cPort(func)
        YFunction._AddToCache("I2cPort", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YI2cPortValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackI2cPort = callback
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
      If (Not (Me._valueCallbackI2cPort Is Nothing)) Then
        Me._valueCallbackI2cPort(Me, value)
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
      Dim msgbin As Byte()
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
      Dim msgbin As Byte()
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
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim res As Integer = 0

      buff = Me._download("rxcnt.bin?pos=" + Convert.ToString(Me._rxptr))
      bufflen = (buff).Length - 1
      While ((bufflen > 0) AndAlso (buff(bufflen) <> 64))
        bufflen = bufflen - 1
      End While
      res = YAPI._atoi((YAPI.DefaultEncoding.GetString(buff)).Substring( 0, bufflen))
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
      Dim url As String
      Dim msgbin As Byte()
      Dim msgarr As List(Of String) = New List(Of String)()
      Dim msglen As Integer = 0
      Dim res As String

      url = "rxmsg.json?len=1&maxw=" + Convert.ToString( maxWait) + "&cmd=!" + Me._escapeAttr(query)
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function reset() As Integer
      Me._rxptr = 0
      Me._rxbuffptr = 0
      ReDim Me._rxbuff(0-1)

      Return Me.sendCommand("Z")
    End Function

    '''*
    ''' <summary>
    '''   Sends a one-way message (provided as a a binary buffer) to a device on the I2C bus.
    ''' <para>
    '''   This function checks and reports communication errors on the I2C bus.
    ''' </para>
    ''' </summary>
    ''' <param name="slaveAddr">
    '''   the 7-bit address of the slave device (without the direction bit)
    ''' </param>
    ''' <param name="buff">
    '''   the binary buffer to be sent
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function i2cSendBin(slaveAddr As Integer, buff As Byte()) As Integer
      Dim nBytes As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim msg As String
      Dim reply As String
      msg = "@" + (slaveAddr).ToString("x02") + ":"
      nBytes = (buff).Length
      idx = 0
      While (idx < nBytes)
        val = buff(idx)
        msg = "" +  msg + "" + (val).ToString("x02")
        idx = idx + 1
      End While

      reply = Me.queryLine(msg,1000)
      If Not((reply).Length > 0) Then
        me._throw( YAPI.IO_ERROR,  "no response from device")
        return YAPI.IO_ERROR
      end if
      idx = reply.IndexOf("[N]!")
      If Not(idx < 0) Then
        me._throw( YAPI.IO_ERROR,  "No ACK received")
        return YAPI.IO_ERROR
      end if
      idx = reply.IndexOf("!")
      If Not(idx < 0) Then
        me._throw( YAPI.IO_ERROR,  "Protocol error")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Sends a one-way message (provided as a list of integer) to a device on the I2C bus.
    ''' <para>
    '''   This function checks and reports communication errors on the I2C bus.
    ''' </para>
    ''' </summary>
    ''' <param name="slaveAddr">
    '''   the 7-bit address of the slave device (without the direction bit)
    ''' </param>
    ''' <param name="values">
    '''   a list of data bytes to be sent
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function i2cSendArray(slaveAddr As Integer, values As List(Of Integer)) As Integer
      Dim nBytes As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim msg As String
      Dim reply As String
      msg = "@" + (slaveAddr).ToString("x02") + ":"
      nBytes = values.Count
      idx = 0
      While (idx < nBytes)
        val = values(idx)
        msg = "" +  msg + "" + (val).ToString("x02")
        idx = idx + 1
      End While

      reply = Me.queryLine(msg,1000)
      If Not((reply).Length > 0) Then
        me._throw( YAPI.IO_ERROR,  "no response from device")
        return YAPI.IO_ERROR
      end if
      idx = reply.IndexOf("[N]!")
      If Not(idx < 0) Then
        me._throw( YAPI.IO_ERROR,  "No ACK received")
        return YAPI.IO_ERROR
      end if
      idx = reply.IndexOf("!")
      If Not(idx < 0) Then
        me._throw( YAPI.IO_ERROR,  "Protocol error")
        return YAPI.IO_ERROR
      end if
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Sends a one-way message (provided as a a binary buffer) to a device on the I2C bus,
    '''   then read back the specified number of bytes from device.
    ''' <para>
    '''   This function checks and reports communication errors on the I2C bus.
    ''' </para>
    ''' </summary>
    ''' <param name="slaveAddr">
    '''   the 7-bit address of the slave device (without the direction bit)
    ''' </param>
    ''' <param name="buff">
    '''   the binary buffer to be sent
    ''' </param>
    ''' <param name="rcvCount">
    '''   the number of bytes to receive once the data bytes are sent
    ''' </param>
    ''' <returns>
    '''   a list of bytes with the data received from slave device.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty binary buffer.
    ''' </para>
    '''/
    Public Overridable Function i2cSendAndReceiveBin(slaveAddr As Integer, buff As Byte(), rcvCount As Integer) As Byte()
      Dim nBytes As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim msg As String
      Dim reply As String
      Dim rcvbytes As Byte()
      msg = "@" + (slaveAddr).ToString("x02") + ":"
      nBytes = (buff).Length
      idx = 0
      While (idx < nBytes)
        val = buff(idx)
        msg = "" +  msg + "" + (val).ToString("x02")
        idx = idx + 1
      End While
      idx = 0
      While (idx < rcvCount)
        msg = "" + msg + "xx"
        idx = idx + 1
      End While

      reply = Me.queryLine(msg,1000)
      ReDim rcvbytes(0-1)
      If Not((reply).Length > 0) Then
        me._throw( YAPI.IO_ERROR,  "no response from device")
        return rcvbytes
      end if
      idx = reply.IndexOf("[N]!")
      If Not(idx < 0) Then
        me._throw( YAPI.IO_ERROR,  "No ACK received")
        return rcvbytes
      end if
      idx = reply.IndexOf("!")
      If Not(idx < 0) Then
        me._throw( YAPI.IO_ERROR,  "Protocol error")
        return rcvbytes
      end if
      reply = (reply).Substring( (reply).Length-2*rcvCount, 2*rcvCount)
      rcvbytes = YAPI._hexStrToBin(reply)
      Return rcvbytes
    End Function

    '''*
    ''' <summary>
    '''   Sends a one-way message (provided as a list of integer) to a device on the I2C bus,
    '''   then read back the specified number of bytes from device.
    ''' <para>
    '''   This function checks and reports communication errors on the I2C bus.
    ''' </para>
    ''' </summary>
    ''' <param name="slaveAddr">
    '''   the 7-bit address of the slave device (without the direction bit)
    ''' </param>
    ''' <param name="values">
    '''   a list of data bytes to be sent
    ''' </param>
    ''' <param name="rcvCount">
    '''   the number of bytes to receive once the data bytes are sent
    ''' </param>
    ''' <returns>
    '''   a list of bytes with the data received from slave device.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function i2cSendAndReceiveArray(slaveAddr As Integer, values As List(Of Integer), rcvCount As Integer) As List(Of Integer)
      Dim nBytes As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim msg As String
      Dim reply As String
      Dim rcvbytes As Byte()
      Dim res As List(Of Integer) = New List(Of Integer)()
      msg = "@" + (slaveAddr).ToString("x02") + ":"
      nBytes = values.Count
      idx = 0
      While (idx < nBytes)
        val = values(idx)
        msg = "" +  msg + "" + (val).ToString("x02")
        idx = idx + 1
      End While
      idx = 0
      While (idx < rcvCount)
        msg = "" + msg + "xx"
        idx = idx + 1
      End While

      reply = Me.queryLine(msg,1000)
      If Not((reply).Length > 0) Then
        me._throw( YAPI.IO_ERROR,  "no response from device")
        return res
      end if
      idx = reply.IndexOf("[N]!")
      If Not(idx < 0) Then
        me._throw( YAPI.IO_ERROR,  "No ACK received")
        return res
      end if
      idx = reply.IndexOf("!")
      If Not(idx < 0) Then
        me._throw( YAPI.IO_ERROR,  "Protocol error")
        return res
      end if
      reply = (reply).Substring( (reply).Length-2*rcvCount, 2*rcvCount)
      rcvbytes = YAPI._hexStrToBin(reply)
      res.Clear()
      idx = 0
      While (idx < rcvCount)
        val = rcvbytes(idx)
        res.Add(val)
        idx = idx + 1
      End While

      Return res
    End Function

    '''*
    ''' <summary>
    '''   Sends a text-encoded I2C code stream to the I2C bus, as is.
    ''' <para>
    '''   An I2C code stream is a string made of hexadecimal data bytes,
    '''   but that may also include the I2C state transitions code:
    '''   "{S}" to emit a start condition,
    '''   "{R}" for a repeated start condition,
    '''   "{P}" for a stop condition,
    '''   "xx" for receiving a data byte,
    '''   "{A}" to ack a data byte received and
    '''   "{N}" to nack a data byte received.
    '''   If a newline ("\n") is included in the stream, the message
    '''   will be terminated and a newline will also be added to the
    '''   receive stream.
    ''' </para>
    ''' </summary>
    ''' <param name="codes">
    '''   the code stream to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeStr(codes As String) As Integer
      Dim bufflen As Integer = 0
      Dim buff As Byte()
      Dim idx As Integer = 0
      Dim ch As Integer = 0
      buff = YAPI.DefaultEncoding.GetBytes(codes)
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
          Return Me.sendCommand("+" + codes)
        End If
      End If
      REM // send string using file upload
      Return Me._upload("txdata", buff)
    End Function

    '''*
    ''' <summary>
    '''   Sends a text-encoded I2C code stream to the I2C bus, and terminate
    '''   the message en relâchant le bus.
    ''' <para>
    '''   An I2C code stream is a string made of hexadecimal data bytes,
    '''   but that may also include the I2C state transitions code:
    '''   "{S}" to emit a start condition,
    '''   "{R}" for a repeated start condition,
    '''   "{P}" for a stop condition,
    '''   "xx" for receiving a data byte,
    '''   "{A}" to ack a data byte received and
    '''   "{N}" to nack a data byte received.
    '''   At the end of the stream, a stop condition is added if missing
    '''   and a newline is added to the receive buffer as well.
    ''' </para>
    ''' </summary>
    ''' <param name="codes">
    '''   the code stream to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeLine(codes As String) As Integer
      Dim bufflen As Integer = 0
      Dim buff As Byte()
      bufflen = (codes).Length
      If (bufflen < 100) Then
        Return Me.sendCommand("!" + codes)
      End If
      REM // send string using file upload
      buff = YAPI.DefaultEncoding.GetBytes("" + codes + "" + vbLf + "")
      Return Me._upload("txdata", buff)
    End Function

    '''*
    ''' <summary>
    '''   Sends a single byte to the I2C bus.
    ''' <para>
    '''   Depending on the I2C bus state, the byte
    '''   will be interpreted as an address byte or a data byte.
    ''' </para>
    ''' </summary>
    ''' <param name="code">
    '''   the byte to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeByte(code As Integer) As Integer
      Return Me.sendCommand("+" + (code).ToString("X02"))
    End Function

    '''*
    ''' <summary>
    '''   Sends a byte sequence (provided as a hexadecimal string) to the I2C bus.
    ''' <para>
    '''   Depending on the I2C bus state, the first byte will be interpreted as an
    '''   address byte or a data byte.
    ''' </para>
    ''' </summary>
    ''' <param name="hexString">
    '''   a string of hexadecimal byte codes
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeHex(hexString As String) As Integer
      Dim bufflen As Integer = 0
      Dim buff As Byte()
      bufflen = (hexString).Length
      If (bufflen < 100) Then
        Return Me.sendCommand("+" + hexString)
      End If
      buff = YAPI.DefaultEncoding.GetBytes(hexString)

      Return Me._upload("txdata", buff)
    End Function

    '''*
    ''' <summary>
    '''   Sends a binary buffer to the I2C bus, as is.
    ''' <para>
    '''   Depending on the I2C bus state, the first byte will be interpreted
    '''   as an address byte or a data byte.
    ''' </para>
    ''' </summary>
    ''' <param name="buff">
    '''   the binary buffer to send
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeBin(buff As Byte()) As Integer
      Dim nBytes As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim msg As String
      msg = ""
      nBytes = (buff).Length
      idx = 0
      While (idx < nBytes)
        val = buff(idx)
        msg = "" +  msg + "" + (val).ToString("x02")
        idx = idx + 1
      End While

      Return Me.writeHex(msg)
    End Function

    '''*
    ''' <summary>
    '''   Sends a byte sequence (provided as a list of bytes) to the I2C bus.
    ''' <para>
    '''   Depending on the I2C bus state, the first byte will be interpreted as an
    '''   address byte or a data byte.
    ''' </para>
    ''' </summary>
    ''' <param name="byteList">
    '''   a list of byte codes
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeArray(byteList As List(Of Integer)) As Integer
      Dim nBytes As Integer = 0
      Dim idx As Integer = 0
      Dim val As Integer = 0
      Dim msg As String
      msg = ""
      nBytes = byteList.Count
      idx = 0
      While (idx < nBytes)
        val = byteList(idx)
        msg = "" +  msg + "" + (val).ToString("x02")
        idx = idx + 1
      End While

      Return Me.writeHex(msg)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of I2C ports started using <c>yFirstI2cPort()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned I2C ports order.
    '''   If you want to find a specific an I2C port, use <c>I2cPort.findI2cPort()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YI2cPort</c> object, corresponding to
    '''   an I2C port currently online, or a <c>Nothing</c> pointer
    '''   if there are no more I2C ports to enumerate.
    ''' </returns>
    '''/
    Public Function nextI2cPort() As YI2cPort
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YI2cPort.FindI2cPort(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of I2C ports currently accessible.
    ''' <para>
    '''   Use the method <c>YI2cPort.nextI2cPort()</c> to iterate on
    '''   next I2C ports.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YI2cPort</c> object, corresponding to
    '''   the first I2C port currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstI2cPort() As YI2cPort
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("I2cPort", 0, p, size, neededsize, errmsg)
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
      Return YI2cPort.FindI2cPort(serial + "." + funcId)
    End Function

    REM --- (end of YI2cPort public methods declaration)

  End Class

  REM --- (YI2cPort functions)

  '''*
  ''' <summary>
  '''   Retrieves an I2C port for a given identifier.
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
  '''   This function does not require that the I2C port is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YI2cPort.isOnline()</c> to test if the I2C port is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an I2C port by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the I2C port
  ''' </param>
  ''' <returns>
  '''   a <c>YI2cPort</c> object allowing you to drive the I2C port.
  ''' </returns>
  '''/
  Public Function yFindI2cPort(ByVal func As String) As YI2cPort
    Return YI2cPort.FindI2cPort(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of I2C ports currently accessible.
  ''' <para>
  '''   Use the method <c>YI2cPort.nextI2cPort()</c> to iterate on
  '''   next I2C ports.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YI2cPort</c> object, corresponding to
  '''   the first I2C port currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstI2cPort() As YI2cPort
    Return YI2cPort.FirstI2cPort()
  End Function


  REM --- (end of YI2cPort functions)

End Module
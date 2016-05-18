'*********************************************************************
'*
'* $Id: yocto_spiport.vb 24252 2016-04-26 13:39:30Z seb $
'*
'* Implements yFindSpiPort(), the high-level API for SpiPort functions
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

Module yocto_spiport

    REM --- (YSpiPort return codes)
    REM --- (end of YSpiPort return codes)
    REM --- (YSpiPort dlldef)
    REM --- (end of YSpiPort dlldef)
  REM --- (YSpiPort globals)

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
  Public Const Y_VOLTAGELEVEL_INVALID As Integer = -1
  Public Const Y_PROTOCOL_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SPIMODE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SSPOLARITY_ACTIVE_LOW As Integer = 0
  Public Const Y_SSPOLARITY_ACTIVE_HIGH As Integer = 1
  Public Const Y_SSPOLARITY_INVALID As Integer = -1
  Public Const Y_SHITFTSAMPLING_OFF As Integer = 0
  Public Const Y_SHITFTSAMPLING_ON As Integer = 1
  Public Const Y_SHITFTSAMPLING_INVALID As Integer = -1
  Public Delegate Sub YSpiPortValueCallback(ByVal func As YSpiPort, ByVal value As String)
  Public Delegate Sub YSpiPortTimedReportCallback(ByVal func As YSpiPort, ByVal measure As YMeasure)
  REM --- (end of YSpiPort globals)

  REM --- (YSpiPort class start)

  '''*
  ''' <summary>
  '''   The SpiPort function interface allows you to fully drive a Yoctopuce
  '''   SPI port, to send and receive data, and to configure communication
  '''   parameters (baud rate, bit count, parity, flow control and protocol).
  ''' <para>
  '''   Note that Yoctopuce SPI ports are not exposed as virtual COM ports.
  '''   They are meant to be used in the same way as all Yoctopuce devices.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YSpiPort
    Inherits YFunction
    REM --- (end of YSpiPort class start)

    REM --- (YSpiPort definitions)
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
    Public Const VOLTAGELEVEL_INVALID As Integer = -1
    Public Const PROTOCOL_INVALID As String = YAPI.INVALID_STRING
    Public Const SPIMODE_INVALID As String = YAPI.INVALID_STRING
    Public Const SSPOLARITY_ACTIVE_LOW As Integer = 0
    Public Const SSPOLARITY_ACTIVE_HIGH As Integer = 1
    Public Const SSPOLARITY_INVALID As Integer = -1
    Public Const SHITFTSAMPLING_OFF As Integer = 0
    Public Const SHITFTSAMPLING_ON As Integer = 1
    Public Const SHITFTSAMPLING_INVALID As Integer = -1
    REM --- (end of YSpiPort definitions)

    REM --- (YSpiPort attributes declaration)
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
    Protected _spiMode As String
    Protected _ssPolarity As Integer
    Protected _shitftSampling As Integer
    Protected _valueCallbackSpiPort As YSpiPortValueCallback
    Protected _rxptr As Integer
    REM --- (end of YSpiPort attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "SpiPort"
      REM --- (YSpiPort attributes initialization)
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
      _spiMode = SPIMODE_INVALID
      _ssPolarity = SSPOLARITY_INVALID
      _shitftSampling = SHITFTSAMPLING_INVALID
      _valueCallbackSpiPort = Nothing
      _rxptr = 0
      REM --- (end of YSpiPort attributes initialization)
    End Sub

    REM --- (YSpiPort private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "rxCount") Then
        _rxCount = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "txCount") Then
        _txCount = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "errCount") Then
        _errCount = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "rxMsgCount") Then
        _rxMsgCount = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "txMsgCount") Then
        _txMsgCount = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "lastMsg") Then
        _lastMsg = member.svalue
        Return 1
      End If
      If (member.name = "currentJob") Then
        _currentJob = member.svalue
        Return 1
      End If
      If (member.name = "startupJob") Then
        _startupJob = member.svalue
        Return 1
      End If
      If (member.name = "command") Then
        _command = member.svalue
        Return 1
      End If
      If (member.name = "voltageLevel") Then
        _voltageLevel = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "protocol") Then
        _protocol = member.svalue
        Return 1
      End If
      If (member.name = "spiMode") Then
        _spiMode = member.svalue
        Return 1
      End If
      If (member.name = "ssPolarity") Then
        If (member.ivalue > 0) Then _ssPolarity = 1 Else _ssPolarity = 0
        Return 1
      End If
      If (member.name = "shitftSampling") Then
        If (member.ivalue > 0) Then _shitftSampling = 1 Else _shitftSampling = 0
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YSpiPort private methods declaration)

    REM --- (YSpiPort public methods declaration)
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RXCOUNT_INVALID
        End If
      End If
      Return Me._rxCount
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return TXCOUNT_INVALID
        End If
      End If
      Return Me._txCount
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ERRCOUNT_INVALID
        End If
      End If
      Return Me._errCount
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RXMSGCOUNT_INVALID
        End If
      End If
      Return Me._rxMsgCount
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return TXMSGCOUNT_INVALID
        End If
      End If
      Return Me._txMsgCount
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LASTMSG_INVALID
        End If
      End If
      Return Me._lastMsg
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CURRENTJOB_INVALID
        End If
      End If
      Return Me._currentJob
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return STARTUPJOB_INVALID
        End If
      End If
      Return Me._startupJob
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      Return Me._command
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
    '''   <c>Y_VOLTAGELEVEL_TTL5V</c>, <c>Y_VOLTAGELEVEL_TTL5VR</c>, <c>Y_VOLTAGELEVEL_RS232</c> and
    '''   <c>Y_VOLTAGELEVEL_RS485</c> corresponding to the voltage level used on the serial line
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_VOLTAGELEVEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_voltageLevel() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return VOLTAGELEVEL_INVALID
        End If
      End If
      Return Me._voltageLevel
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
    '''   <c>Y_VOLTAGELEVEL_TTL5V</c>, <c>Y_VOLTAGELEVEL_TTL5VR</c>, <c>Y_VOLTAGELEVEL_RS232</c> and
    '''   <c>Y_VOLTAGELEVEL_RS485</c> corresponding to the voltage type used on the serial line
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
    '''   On failure, throws an exception or returns <c>Y_PROTOCOL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_protocol() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PROTOCOL_INVALID
        End If
      End If
      Return Me._protocol
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
    '''   "125000,0,msb".
    ''' <para>
    '''   The string includes the baud rate, the SPI mode (between
    '''   0 and 3) and the bit order.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the SPI port communication parameters, as a string such as
    '''   "125000,0,msb"
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SPIMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_spiMode() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SPIMODE_INVALID
        End If
      End If
      Return Me._spiMode
    End Function


    '''*
    ''' <summary>
    '''   Changes the SPI port communication parameters, with a string such as
    '''   "125000,0,msb".
    ''' <para>
    '''   The string includes the baud rate, the SPI mode (between
    '''   0 and 3) and the bit order.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the SPI port communication parameters, with a string such as
    '''   "125000,0,msb"
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
    Public Function set_spiMode(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("spiMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the SS line polarity.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_SSPOLARITY_ACTIVE_LOW</c> or <c>Y_SSPOLARITY_ACTIVE_HIGH</c>, according to the SS line polarity
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SSPOLARITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_ssPolarity() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SSPOLARITY_INVALID
        End If
      End If
      Return Me._ssPolarity
    End Function


    '''*
    ''' <summary>
    '''   Changes the SS line polarity.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_SSPOLARITY_ACTIVE_LOW</c> or <c>Y_SSPOLARITY_ACTIVE_HIGH</c>, according to the SS line polarity
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
    Public Function set_ssPolarity(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("ssPolarity", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns true when the SDI line phase is shifted with regards to the SDO line.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_SHITFTSAMPLING_OFF</c> or <c>Y_SHITFTSAMPLING_ON</c>, according to true when the SDI
    '''   line phase is shifted with regards to the SDO line
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SHITFTSAMPLING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_shitftSampling() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SHITFTSAMPLING_INVALID
        End If
      End If
      Return Me._shitftSampling
    End Function


    '''*
    ''' <summary>
    '''   Changes the SDI line sampling shift.
    ''' <para>
    '''   When disabled, SDI line is
    '''   sampled in the middle of data output time. When enabled, SDI line is
    '''   samples at the end of data output time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_SHITFTSAMPLING_OFF</c> or <c>Y_SHITFTSAMPLING_ON</c>, according to the SDI line sampling shift
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
    Public Function set_shitftSampling(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("shitftSampling", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a SPI port for a given identifier.
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
    '''   This function does not require that the SPI port is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YSpiPort.isOnline()</c> to test if the SPI port is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a SPI port by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the SPI port
    ''' </param>
    ''' <returns>
    '''   a <c>YSpiPort</c> object allowing you to drive the SPI port.
    ''' </returns>
    '''/
    Public Shared Function FindSpiPort(func As String) As YSpiPort
      Dim obj As YSpiPort
      obj = CType(YFunction._FindFromCache("SpiPort", func), YSpiPort)
      If ((obj Is Nothing)) Then
        obj = New YSpiPort(func)
        YFunction._AddToCache("SpiPort", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YSpiPortValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackSpiPort = callback
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
      If (Not (Me._valueCallbackSpiPort Is Nothing)) Then
        Me._valueCallbackSpiPort(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function sendCommand(text As String) As Integer
      REM // may throw an exception
      Return Me.set_command(text)
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
      REM // may throw an exception
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeByte(code As Integer) As Integer
      REM // may throw an exception
      Return Me.sendCommand("$" + YAPI._intToHex(code,02))
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeStr(text As String) As Integer
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim idx As Integer = 0
      Dim ch As Integer = 0
      buff = YAPI.DefaultEncoding.GetBytes(text)
      bufflen = (buff).Length
      If (bufflen < 100) Then
        REM
        ch = &H20
        idx = 0
        While ((idx < bufflen) And (ch <> 0))
          ch = buff(idx)
          If ((ch >= &H20) And (ch < &H7f)) Then
            idx = idx + 1
          Else
            ch = 0
          End If
        End While
        If (idx >= bufflen) Then
          REM
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeBin(buff As Byte()) As Integer
      REM // may throw an exception
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeArray(byteList As List(Of Integer)) As Integer
      Dim buff As Byte()
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
      REM // may throw an exception
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeHex(hexString As String) As Integer
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim idx As Integer = 0
      Dim hexb As Integer = 0
      Dim res As Integer = 0
      bufflen = (hexString).Length
      If (bufflen < 100) Then
        REM
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
      REM // may throw an exception
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function writeLine(text As String) As Integer
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim idx As Integer = 0
      Dim ch As Integer = 0
      buff = YAPI.DefaultEncoding.GetBytes("" + text + "" + vbCr + "" + vbLf + "")
      bufflen = (buff).Length-2
      If (bufflen < 100) Then
        REM
        ch = &H20
        idx = 0
        While ((idx < bufflen) And (ch <> 0))
          ch = buff(idx)
          If ((ch >= &H20) And (ch < &H7f)) Then
            idx = idx + 1
          Else
            ch = 0
          End If
        End While
        If (idx >= bufflen) Then
          REM
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
    '''   or if there is no data available yet, the function returns YAPI_NO_MORE_DATA.
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
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim res As Integer = 0
      REM // may throw an exception
      buff = Me._download("rxdata.bin?pos=" + Convert.ToString(Me._rxptr) + "&len=1")
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) And (buff(bufflen) <> 64))
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
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim res As String
      If (nChars > 65535) Then
        nChars = 65535
      End If
      REM // may throw an exception
      buff = Me._download("rxdata.bin?pos=" + Convert.ToString( Me._rxptr) + "&len=" + Convert.ToString(nChars))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) And (buff(bufflen) <> 64))
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
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim idx As Integer = 0
      Dim res As Byte()
      If (nChars > 65535) Then
        nChars = 65535
      End If
      REM // may throw an exception
      buff = Me._download("rxdata.bin?pos=" + Convert.ToString( Me._rxptr) + "&len=" + Convert.ToString(nChars))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) And (buff(bufflen) <> 64))
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
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function readArray(nChars As Integer) As List(Of Integer)
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim idx As Integer = 0
      Dim b As Integer = 0
      Dim res As List(Of Integer) = New List(Of Integer)()
      If (nChars > 65535) Then
        nChars = 65535
      End If
      REM // may throw an exception
      buff = Me._download("rxdata.bin?pos=" + Convert.ToString( Me._rxptr) + "&len=" + Convert.ToString(nChars))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) And (buff(bufflen) <> 64))
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
      Dim buff As Byte()
      Dim bufflen As Integer = 0
      Dim mult As Integer = 0
      Dim endpos As Integer = 0
      Dim ofs As Integer = 0
      Dim res As String
      If (nBytes > 65535) Then
        nBytes = 65535
      End If
      REM // may throw an exception
      buff = Me._download("rxdata.bin?pos=" + Convert.ToString( Me._rxptr) + "&len=" + Convert.ToString(nBytes))
      bufflen = (buff).Length - 1
      endpos = 0
      mult = 1
      While ((bufflen > 0) And (buff(bufflen) <> 64))
        endpos = endpos + mult * (buff(bufflen) - 48)
        mult = mult * 10
        bufflen = bufflen - 1
      End While
      Me._rxptr = endpos
      res = ""
      ofs = 0
      While (ofs + 3 < bufflen)
        res = "" +  res + "" + YAPI._intToHex( buff(ofs),02) + "" + YAPI._intToHex( buff(ofs + 1),02) + "" + YAPI._intToHex( buff(ofs + 2),02) + "" + YAPI._intToHex(buff(ofs + 3),02)
        ofs = ofs + 4
      End While
      While (ofs < bufflen)
        res = "" +  res + "" + YAPI._intToHex(buff(ofs),02)
        ofs = ofs + 1
      End While
      Return res
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
      REM // may throw an exception
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
      REM // may throw an exception
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
      REM // may throw an exception
      buff = Me._download("rxcnt.bin?pos=" + Convert.ToString(Me._rxptr))
      bufflen = (buff).Length - 1
      While ((bufflen > 0) And (buff(bufflen) <> 64))
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
    '''   On failure, throws an exception or returns an empty array.
    ''' </para>
    '''/
    Public Overridable Function queryLine(query As String, maxWait As Integer) As String
      Dim url As String
      Dim msgbin As Byte()
      Dim msgarr As List(Of String) = New List(Of String)()
      Dim msglen As Integer = 0
      Dim res As String
      REM // may throw an exception
      url = "rxmsg.json?len=1&maxw=" + Convert.ToString( maxWait) + "&cmd=!" + query
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
      REM // may throw an exception
      Return Me.set_currentJob(jobfile)
    End Function

    '''*
    ''' <summary>
    '''   Manually sets the state of the SS line.
    ''' <para>
    '''   This function has no effect when
    '''   the SS line is handled automatically.
    ''' </para>
    ''' </summary>
    ''' <param name="val">
    '''   1 to turn SS active, 0 to release SS.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_SS(val As Integer) As Integer
      REM // may throw an exception
      Return Me.sendCommand("S" + Convert.ToString(val))
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of SPI ports started using <c>yFirstSpiPort()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSpiPort</c> object, corresponding to
    '''   a SPI port currently online, or a <c>null</c> pointer
    '''   if there are no more SPI ports to enumerate.
    ''' </returns>
    '''/
    Public Function nextSpiPort() As YSpiPort
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YSpiPort.FindSpiPort(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of SPI ports currently accessible.
    ''' <para>
    '''   Use the method <c>YSpiPort.nextSpiPort()</c> to iterate on
    '''   next SPI ports.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YSpiPort</c> object, corresponding to
    '''   the first SPI port currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstSpiPort() As YSpiPort
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("SpiPort", 0, p, size, neededsize, errmsg)
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
      Return YSpiPort.FindSpiPort(serial + "." + funcId)
    End Function

    REM --- (end of YSpiPort public methods declaration)

  End Class

  REM --- (SpiPort functions)

  '''*
  ''' <summary>
  '''   Retrieves a SPI port for a given identifier.
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
  '''   This function does not require that the SPI port is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YSpiPort.isOnline()</c> to test if the SPI port is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a SPI port by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the SPI port
  ''' </param>
  ''' <returns>
  '''   a <c>YSpiPort</c> object allowing you to drive the SPI port.
  ''' </returns>
  '''/
  Public Function yFindSpiPort(ByVal func As String) As YSpiPort
    Return YSpiPort.FindSpiPort(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of SPI ports currently accessible.
  ''' <para>
  '''   Use the method <c>YSpiPort.nextSpiPort()</c> to iterate on
  '''   next SPI ports.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YSpiPort</c> object, corresponding to
  '''   the first SPI port currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstSpiPort() As YSpiPort
    Return YSpiPort.FirstSpiPort()
  End Function


  REM --- (end of SpiPort functions)

End Module

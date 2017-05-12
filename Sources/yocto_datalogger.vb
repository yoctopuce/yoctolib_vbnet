'/********************************************************************
'*
'* $Id: yocto_datalogger.vb 27282 2017-04-25 15:44:42Z seb $
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
Imports System.Runtime.InteropServices

Module yocto_datalogger


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
  
  REM --- (generated code: DataLogger functions)

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


  REM --- (end of generated code: DataLogger functions)


End Module
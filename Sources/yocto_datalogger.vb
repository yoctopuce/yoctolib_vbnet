'/********************************************************************
'*
'* $Id: yocto_datalogger.vb 27104 2017-04-06 22:14:54Z seb $
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
      Dim json As TJsonParser = Nothing
      Dim res, count As Integer
      Dim root, el As TJSONRECORD
      Dim name As String
      Dim coldiv As List(Of Integer) = New List(Of Integer)
      Dim coltype As List(Of Integer) = New List(Of Integer)
      Dim udat As List(Of UInteger) = New List(Of UInteger)
      Dim dat As List(Of Double) = New List(Of Double)
      Dim colscl As List(Of Double) = New List(Of Double)
      Dim colofs As List(Of Integer) = New List(Of Integer)
      Dim caltyp As List(Of Integer) = New List(Of Integer)
      Dim calhdl As List(Of yCalibrationHandler) = New List(Of yCalibrationHandler)
      Dim calpar As List(Of List(Of Integer)) = New List(Of List(Of Integer))
      Dim calraw As List(Of List(Of Double)) = New List(Of List(Of Double))
      Dim calref As List(Of List(Of Double)) = New List(Of List(Of Double))
      Dim x, i, j As Integer
      Dim value As Double
      Dim sdat As String

      res = _dataLogger.getData(_runNo, _timeStamp, json)
      If (res <> YAPI_SUCCESS) Then
        loadStream = res
        Exit Function
      End If

      _nRows = 0
      _nCols = 0
      _columnNames.Clear()

      root = json.GetRootNode()
      For i = 0 To root.membercount - 1

        el = root.members(i)
        name = el.name
        If (name = "time") Then

          _timeStamp = el.ivalue
        ElseIf (name = "UTC") Then
          _utcStamp = CUInt(el.ivalue)
        ElseIf (name = "interval") Then
          _interval = el.ivalue
        ElseIf (name = "nRows") Then
          _nRows = CInt(el.ivalue)
        ElseIf (name = "keys") Then
          _nCols = el.itemcount
          For j = 0 To _nCols - 1
            _columnNames.Add(el.items(j).svalue)
          Next j
        ElseIf (name = "div") Then
          _nCols = el.itemcount
          For j = 0 To _nCols - 1
            coldiv.Add(CInt(el.items(j).ivalue))
          Next j
        ElseIf (name = "type") Then
          _nCols = el.itemcount
          For j = 0 To _nCols - 1
            coltype.Add(CInt(el.items(j).ivalue))
          Next j
        ElseIf (name = "scal") Then
          _nCols = el.itemcount
          For j = 0 To _nCols - 1
            colscl.Add(el.items(j).ivalue / 65536.0)
            If coltype(j) <> 0 Then
              colofs.Add(-32767)
            Else
              colofs.Add(0)
            End If
          Next j
        ElseIf (name = "cal") Then
          _nCols = el.itemcount
          For j = 0 To _nCols - 1
            Dim calibration_Str As String = el.items(j).svalue
            Dim cur_calpar As List(Of Integer) = Nothing
            Dim cur_calraw As List(Of Double) = Nothing
            Dim cur_calref As List(Of Double) = Nothing
            Dim calibType As Integer = 0
            caltyp.Add(calibType)
            calhdl.Add(YAPI._getCalibrationHandler(calibType))
            calpar.Add(cur_calpar)
            calraw.Add(cur_calraw)
            calref.Add(cur_calref)
          Next j
        ElseIf (name = "data") Then
          If (colscl.Count <= 0) Then
            For j = 0 To _nCols - 1
              colscl.Add(1.0 / coldiv(j))
              If (coltype(j) <> 0) Then
                colofs.Add(-32767)
              Else
                colofs.Add(0)
              End If
            Next j
          End If
          count = el.itemcount
          udat.Clear()
          If (el.recordtype = TJSONRECORDTYPE.JSON_STRING) Then
            sdat = el.svalue
            Dim p As Integer = 0
            While (p < sdat.Length)
              Dim val As UInteger
              Dim c As UInteger = CUInt(Asc(sdat.Substring(p, 1)))
              p += 1
              If (c >= 97) Then REM 97 ='a'
                Dim srcpos As Integer = CInt((udat.Count - 1 - (c - 97)))
                If (srcpos < 0) Then
                  _dataLogger.throw_friend(YAPI_IO_ERROR, "Unexpected JSON reply format")
                  Return YAPI_IO_ERROR
                End If
                val = udat.ElementAt(srcpos)
              Else
                If (p + 2 > sdat.Length) Then
                  _dataLogger.throw_friend(YAPI_IO_ERROR, "Unexpected JSON reply format")
                  Return YAPI_IO_ERROR
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
              udat.Add(val)
            End While
          Else
            count = el.itemcount
            For j = 0 To count - 1
              Dim tmp As UInteger = CUInt(el.items(j).ivalue)
              udat.Add(tmp)
            Next
          End If


          _values = New List(Of List(Of Double))()
          dat = New List(Of Double)()
          x = 0
          For Each uval As Integer In udat
            If coltype(x) < 2 Then
              value = (uval + colofs(x)) * colscl(x)
            Else
              value = YAPI._decimalToDouble(uval - 32767)
            End If
            If (caltyp(x) > 0 And calhdl(x) <> Nothing) Then
              Dim handler As yCalibrationHandler = calhdl(x)
              If (caltyp(x) <= 10) Then
                value = handler((uval + coldiv(x)) / coldiv(x), caltyp(x), calpar(x), calraw(x), calref(x))
              ElseIf (caltyp(x) > 20) Then
                value = handler(value, caltyp(x), calpar(x), calraw(x), calref(x))
              End If
            End If
            dat.Add(value)
            x = x + 1
            If (x = _nCols) Then
              _values.Add(dat)
              dat.Clear()
              x = 0
            End If
          Next
        End If
      Next i

      json = Nothing

      loadStream = YAPI_SUCCESS
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

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "currentRunIndex") Then
        _currentRunIndex = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "timeUTC") Then
        _timeUTC = member.ivalue
        Return 1
      End If
      If (member.name = "recording") Then
        _recording = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "autoStart") Then
        If (member.ivalue > 0) Then _autoStart = 1 Else _autoStart = 0
        Return 1
      End If
      If (member.name = "beaconDriven") Then
        If (member.ivalue > 0) Then _beaconDriven = 1 Else _beaconDriven = 0
        Return 1
      End If
      If (member.name = "clearHistory") Then
        If (member.ivalue > 0) Then _clearHistory = 1 Else _clearHistory = 0
        Return 1
      End If
      Return MyBase._parseAttr(member)
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


    Public Function getData(ByVal runIdx As Long, ByVal timeIdx As Long, ByRef jsondata As TJsonParser) As Integer

      Dim dev As YDevice = Nothing
      Dim errmsg As String = ""
      Dim query As String
      Dim buffer As String = ""
      Dim res As Integer

      If (_dataLoggerURL = "") Then
        _dataLoggerURL = "/logger.json"
      End If
      REM Resolve our reference to our device, load REST API
      res = _getDevice(dev, errmsg)
      If (YISERR(res)) Then
        _throw(res, errmsg)
        getData = res
        Exit Function
      End If

      If (timeIdx > 0) Then
        query = "GET " + _dataLoggerURL + "?run=" + LTrim(Str(runIdx)) + "&time=" + LTrim(Str(timeIdx)) + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      Else
        query = "GET " + _dataLoggerURL + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
      End If

      res = dev.HTTPRequest(query, buffer, errmsg)
      REM make sure a device scan does not solve the issue
      If (YISERR(res)) Then
        res = yapiUpdateDeviceList(1, errmsg)
        If (YISERR(res)) Then
          getData = res
          Exit Function
        End If

        res = dev.HTTPRequest("GET " + _dataLoggerURL + " HTTP/1.1" + Chr(13) + Chr(10) + Chr(13) + Chr(10), buffer, errmsg)
        If (YISERR(res)) Then
          _throw(res, errmsg)
          getData = res
          Exit Function
        End If
      End If

      Try
        jsondata = New TJsonParser(buffer)
      Catch e As Exception
        errmsg = "unexpected JSON structure: " + e.Message
        _throw(YAPI_IO_ERROR, errmsg)
        getData = YAPI_IO_ERROR
        Exit Function
      End Try
      If (jsondata.httpcode = 404 And _dataLoggerURL <> "/dataLogger.json") Then
        REM retry using backward-compatible datalogger URL
        _dataLoggerURL = "/dataLogger.json"
        Return getData(runIdx, timeIdx, jsondata)
      End If
      getData = YAPI_SUCCESS
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

      Dim j As TJsonParser = Nothing
      Dim i, res As Integer
      Dim root, el As TJSONRECORD
      Dim ds As List(Of YDataStream)

      v.Clear()
      res = getData(0, 0, j)
      If (res <> YAPI_SUCCESS) Then
        Return res
      End If
      root = j.GetRootNode()
      If (root.itemcount = 0) Then
        Return YAPI_SUCCESS
      End If

      If root.items.ElementAt(0).recordtype = TJSONRECORDTYPE.JSON_ARRAY Then
        For i = 0 To root.itemcount - 1
          el = root.items(i)
          v.Add(New YOldDataStream(Me, CInt(el.items(0).ivalue), CInt(el.items(1).ivalue), el.items(2).ivalue, CInt(el.items(3).ivalue)))
        Next i
      Else
        Dim json_buffer As String = j.convertToString(root, False)
        Dim sets As List(Of YDataSet) = parse_dataSets(YAPI.DefaultEncoding.GetBytes(json_buffer))
        For sj As Integer = 0 To sets.Count - 1 Step 1
          ds = sets.ElementAt(sj).get_privateDataStreams()
          For si As Integer = 0 To ds.Count - 1 Step 1
            v.Add(ds.ElementAt(si))
          Next si
        Next sj
      End If
      j = Nothing
      get_dataStreams = YAPI_SUCCESS
    End Function


    Public Sub throw_friend(ByVal errType As System.Int32, ByVal errMsg As String)
      _throw(errType, errMsg)
    End Sub

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
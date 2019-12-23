' ********************************************************************
'
'  $Id: yocto_proximity.vb 38899 2019-12-20 17:21:03Z mvuilleu $
'
'  Implements yFindProximity(), the high-level API for Proximity functions
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

Module yocto_proximity

    REM --- (YProximity return codes)
    REM --- (end of YProximity return codes)
    REM --- (YProximity dlldef)
    REM --- (end of YProximity dlldef)
   REM --- (YProximity yapiwrapper)
   REM --- (end of YProximity yapiwrapper)
  REM --- (YProximity globals)

  Public Const Y_SIGNALVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_DETECTIONTHRESHOLD_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_DETECTIONHYSTERESIS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_PRESENCEMINTIME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_REMOVALMINTIME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ISPRESENT_FALSE As Integer = 0
  Public Const Y_ISPRESENT_TRUE As Integer = 1
  Public Const Y_ISPRESENT_INVALID As Integer = -1
  Public Const Y_LASTTIMEAPPROACHED_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_LASTTIMEREMOVED_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PULSECOUNTER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PROXIMITYREPORTMODE_NUMERIC As Integer = 0
  Public Const Y_PROXIMITYREPORTMODE_PRESENCE As Integer = 1
  Public Const Y_PROXIMITYREPORTMODE_PULSECOUNT As Integer = 2
  Public Const Y_PROXIMITYREPORTMODE_INVALID As Integer = -1
  Public Delegate Sub YProximityValueCallback(ByVal func As YProximity, ByVal value As String)
  Public Delegate Sub YProximityTimedReportCallback(ByVal func As YProximity, ByVal measure As YMeasure)
  REM --- (end of YProximity globals)

  REM --- (YProximity class start)

  '''*
  ''' <summary>
  '''   The <c>YProximity</c> class allows you to read and configure Yoctopuce proximity sensors.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  '''   This class adds the ability to setup a detection threshold and to count the
  '''   number of detected state changes.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YProximity
    Inherits YSensor
    REM --- (end of YProximity class start)

    REM --- (YProximity definitions)
    Public Const SIGNALVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const DETECTIONTHRESHOLD_INVALID As Integer = YAPI.INVALID_UINT
    Public Const DETECTIONHYSTERESIS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const PRESENCEMINTIME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const REMOVALMINTIME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ISPRESENT_FALSE As Integer = 0
    Public Const ISPRESENT_TRUE As Integer = 1
    Public Const ISPRESENT_INVALID As Integer = -1
    Public Const LASTTIMEAPPROACHED_INVALID As Long = YAPI.INVALID_LONG
    Public Const LASTTIMEREMOVED_INVALID As Long = YAPI.INVALID_LONG
    Public Const PULSECOUNTER_INVALID As Long = YAPI.INVALID_LONG
    Public Const PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
    Public Const PROXIMITYREPORTMODE_NUMERIC As Integer = 0
    Public Const PROXIMITYREPORTMODE_PRESENCE As Integer = 1
    Public Const PROXIMITYREPORTMODE_PULSECOUNT As Integer = 2
    Public Const PROXIMITYREPORTMODE_INVALID As Integer = -1
    REM --- (end of YProximity definitions)

    REM --- (YProximity attributes declaration)
    Protected _signalValue As Double
    Protected _detectionThreshold As Integer
    Protected _detectionHysteresis As Integer
    Protected _presenceMinTime As Integer
    Protected _removalMinTime As Integer
    Protected _isPresent As Integer
    Protected _lastTimeApproached As Long
    Protected _lastTimeRemoved As Long
    Protected _pulseCounter As Long
    Protected _pulseTimer As Long
    Protected _proximityReportMode As Integer
    Protected _valueCallbackProximity As YProximityValueCallback
    Protected _timedReportCallbackProximity As YProximityTimedReportCallback
    REM --- (end of YProximity attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Proximity"
      REM --- (YProximity attributes initialization)
      _signalValue = SIGNALVALUE_INVALID
      _detectionThreshold = DETECTIONTHRESHOLD_INVALID
      _detectionHysteresis = DETECTIONHYSTERESIS_INVALID
      _presenceMinTime = PRESENCEMINTIME_INVALID
      _removalMinTime = REMOVALMINTIME_INVALID
      _isPresent = ISPRESENT_INVALID
      _lastTimeApproached = LASTTIMEAPPROACHED_INVALID
      _lastTimeRemoved = LASTTIMEREMOVED_INVALID
      _pulseCounter = PULSECOUNTER_INVALID
      _pulseTimer = PULSETIMER_INVALID
      _proximityReportMode = PROXIMITYREPORTMODE_INVALID
      _valueCallbackProximity = Nothing
      _timedReportCallbackProximity = Nothing
      REM --- (end of YProximity attributes initialization)
    End Sub

    REM --- (YProximity private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("signalValue") Then
        _signalValue = Math.Round(json_val.getDouble("signalValue") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("detectionThreshold") Then
        _detectionThreshold = CInt(json_val.getLong("detectionThreshold"))
      End If
      If json_val.has("detectionHysteresis") Then
        _detectionHysteresis = CInt(json_val.getLong("detectionHysteresis"))
      End If
      If json_val.has("presenceMinTime") Then
        _presenceMinTime = CInt(json_val.getLong("presenceMinTime"))
      End If
      If json_val.has("removalMinTime") Then
        _removalMinTime = CInt(json_val.getLong("removalMinTime"))
      End If
      If json_val.has("isPresent") Then
        If (json_val.getInt("isPresent") > 0) Then _isPresent = 1 Else _isPresent = 0
      End If
      If json_val.has("lastTimeApproached") Then
        _lastTimeApproached = json_val.getLong("lastTimeApproached")
      End If
      If json_val.has("lastTimeRemoved") Then
        _lastTimeRemoved = json_val.getLong("lastTimeRemoved")
      End If
      If json_val.has("pulseCounter") Then
        _pulseCounter = json_val.getLong("pulseCounter")
      End If
      If json_val.has("pulseTimer") Then
        _pulseTimer = json_val.getLong("pulseTimer")
      End If
      If json_val.has("proximityReportMode") Then
        _proximityReportMode = CInt(json_val.getLong("proximityReportMode"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YProximity private methods declaration)

    REM --- (YProximity public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current value of signal measured by the proximity sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current value of signal measured by the proximity sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SIGNALVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signalValue() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SIGNALVALUE_INVALID
        End If
      End If
      res = Math.Round(Me._signalValue * 1000) / 1000
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the threshold used to determine the logical state of the proximity sensor, when considered
    '''   as a binary input (on/off).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the threshold used to determine the logical state of the proximity
    '''   sensor, when considered
    '''   as a binary input (on/off)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DETECTIONTHRESHOLD_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_detectionThreshold() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DETECTIONTHRESHOLD_INVALID
        End If
      End If
      res = Me._detectionThreshold
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the threshold used to determine the logical state of the proximity sensor, when considered
    '''   as a binary input (on/off).
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the threshold used to determine the logical state of the proximity
    '''   sensor, when considered
    '''   as a binary input (on/off)
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
    Public Function set_detectionThreshold(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("detectionThreshold", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the hysteresis used to determine the logical state of the proximity sensor, when considered
    '''   as a binary input (on/off).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the hysteresis used to determine the logical state of the proximity
    '''   sensor, when considered
    '''   as a binary input (on/off)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DETECTIONHYSTERESIS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_detectionHysteresis() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DETECTIONHYSTERESIS_INVALID
        End If
      End If
      res = Me._detectionHysteresis
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the hysteresis used to determine the logical state of the proximity sensor, when considered
    '''   as a binary input (on/off).
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the hysteresis used to determine the logical state of the proximity
    '''   sensor, when considered
    '''   as a binary input (on/off)
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
    Public Function set_detectionHysteresis(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("detectionHysteresis", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the minimal detection duration before signalling a presence event.
    ''' <para>
    '''   Any shorter detection is
    '''   considered as noise or bounce (false positive) and filtered out.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minimal detection duration before signalling a presence event
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PRESENCEMINTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_presenceMinTime() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PRESENCEMINTIME_INVALID
        End If
      End If
      res = Me._presenceMinTime
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the minimal detection duration before signalling a presence event.
    ''' <para>
    '''   Any shorter detection is
    '''   considered as noise or bounce (false positive) and filtered out.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minimal detection duration before signalling a presence event
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
    Public Function set_presenceMinTime(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("presenceMinTime", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the minimal detection duration before signalling a removal event.
    ''' <para>
    '''   Any shorter detection is
    '''   considered as noise or bounce (false positive) and filtered out.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minimal detection duration before signalling a removal event
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_REMOVALMINTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_removalMinTime() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return REMOVALMINTIME_INVALID
        End If
      End If
      res = Me._removalMinTime
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the minimal detection duration before signalling a removal event.
    ''' <para>
    '''   Any shorter detection is
    '''   considered as noise or bounce (false positive) and filtered out.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minimal detection duration before signalling a removal event
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
    Public Function set_removalMinTime(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("removalMinTime", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns true if the input (considered as binary) is active (detection value is smaller than the specified <c>threshold</c>), and false otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ISPRESENT_FALSE</c> or <c>Y_ISPRESENT_TRUE</c>, according to true if the input
    '''   (considered as binary) is active (detection value is smaller than the specified <c>threshold</c>),
    '''   and false otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ISPRESENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_isPresent() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ISPRESENT_INVALID
        End If
      End If
      res = Me._isPresent
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of elapsed milliseconds between the module power on and the last observed
    '''   detection (the input contact transitioned from absent to present).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of elapsed milliseconds between the module power on and the last observed
    '''   detection (the input contact transitioned from absent to present)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LASTTIMEAPPROACHED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastTimeApproached() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LASTTIMEAPPROACHED_INVALID
        End If
      End If
      res = Me._lastTimeApproached
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of elapsed milliseconds between the module power on and the last observed
    '''   detection (the input contact transitioned from present to absent).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of elapsed milliseconds between the module power on and the last observed
    '''   detection (the input contact transitioned from present to absent)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LASTTIMEREMOVED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastTimeRemoved() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return LASTTIMEREMOVED_INVALID
        End If
      End If
      res = Me._lastTimeRemoved
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the pulse counter value.
    ''' <para>
    '''   The value is a 32 bit integer. In case
    '''   of overflow (>=2^32), the counter will wrap. To reset the counter, just
    '''   call the resetCounter() method.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the pulse counter value
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PULSECOUNTER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pulseCounter() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PULSECOUNTER_INVALID
        End If
      End If
      res = Me._pulseCounter
      Return res
    End Function


    Public Function set_pulseCounter(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pulseCounter", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the timer of the pulse counter (ms).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the timer of the pulse counter (ms)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PULSETIMER_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pulseTimer() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PULSETIMER_INVALID
        End If
      End If
      res = Me._pulseTimer
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the parameter (sensor value, presence or pulse count) returned by the get_currentValue function and callbacks.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_PROXIMITYREPORTMODE_NUMERIC</c>, <c>Y_PROXIMITYREPORTMODE_PRESENCE</c> and
    '''   <c>Y_PROXIMITYREPORTMODE_PULSECOUNT</c> corresponding to the parameter (sensor value, presence or
    '''   pulse count) returned by the get_currentValue function and callbacks
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PROXIMITYREPORTMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_proximityReportMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PROXIMITYREPORTMODE_INVALID
        End If
      End If
      res = Me._proximityReportMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the  parameter  type (sensor value, presence or pulse count) returned by the get_currentValue function and callbacks.
    ''' <para>
    '''   The edge count value is limited to the 6 lowest digits. For values greater than one million, use
    '''   get_pulseCounter().
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_PROXIMITYREPORTMODE_NUMERIC</c>, <c>Y_PROXIMITYREPORTMODE_PRESENCE</c> and
    '''   <c>Y_PROXIMITYREPORTMODE_PULSECOUNT</c> corresponding to the  parameter  type (sensor value,
    '''   presence or pulse count) returned by the get_currentValue function and callbacks
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
    Public Function set_proximityReportMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("proximityReportMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a proximity sensor for a given identifier.
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
    '''   This function does not require that the proximity sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YProximity.isOnline()</c> to test if the proximity sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a proximity sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the proximity sensor, for instance
    '''   <c>YPROXIM1.proximity1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YProximity</c> object allowing you to drive the proximity sensor.
    ''' </returns>
    '''/
    Public Shared Function FindProximity(func As String) As YProximity
      Dim obj As YProximity
      obj = CType(YFunction._FindFromCache("Proximity", func), YProximity)
      If ((obj Is Nothing)) Then
        obj = New YProximity(func)
        YFunction._AddToCache("Proximity", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YProximityValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackProximity = callback
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
      If (Not (Me._valueCallbackProximity Is Nothing)) Then
        Me._valueCallbackProximity(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
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
    Public Overloads Function registerTimedReportCallback(callback As YProximityTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackProximity = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackProximity Is Nothing)) Then
        Me._timedReportCallbackProximity(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Resets the pulse counter value as well as its timer.
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
    Public Overridable Function resetCounter() As Integer
      Return Me.set_pulseCounter(0)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of proximity sensors started using <c>yFirstProximity()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned proximity sensors order.
    '''   If you want to find a specific a proximity sensor, use <c>Proximity.findProximity()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YProximity</c> object, corresponding to
    '''   a proximity sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more proximity sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextProximity() As YProximity
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YProximity.FindProximity(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of proximity sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YProximity.nextProximity()</c> to iterate on
    '''   next proximity sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YProximity</c> object, corresponding to
    '''   the first proximity sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstProximity() As YProximity
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Proximity", 0, p, size, neededsize, errmsg)
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
      Return YProximity.FindProximity(serial + "." + funcId)
    End Function

    REM --- (end of YProximity public methods declaration)

  End Class

  REM --- (YProximity functions)

  '''*
  ''' <summary>
  '''   Retrieves a proximity sensor for a given identifier.
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
  '''   This function does not require that the proximity sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YProximity.isOnline()</c> to test if the proximity sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a proximity sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the proximity sensor, for instance
  '''   <c>YPROXIM1.proximity1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YProximity</c> object allowing you to drive the proximity sensor.
  ''' </returns>
  '''/
  Public Function yFindProximity(ByVal func As String) As YProximity
    Return YProximity.FindProximity(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of proximity sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YProximity.nextProximity()</c> to iterate on
  '''   next proximity sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YProximity</c> object, corresponding to
  '''   the first proximity sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstProximity() As YProximity
    Return YProximity.FirstProximity()
  End Function


  REM --- (end of YProximity functions)

End Module

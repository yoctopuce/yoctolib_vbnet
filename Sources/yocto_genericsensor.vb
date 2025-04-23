' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindGenericSensor(), the high-level API for GenericSensor functions
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

Module yocto_genericsensor

    REM --- (YGenericSensor return codes)
    REM --- (end of YGenericSensor return codes)
    REM --- (YGenericSensor dlldef)
    REM --- (end of YGenericSensor dlldef)
   REM --- (YGenericSensor yapiwrapper)
   REM --- (end of YGenericSensor yapiwrapper)
  REM --- (YGenericSensor globals)

  Public Const Y_SIGNALVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_SIGNALUNIT_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SIGNALRANGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_VALUERANGE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_SIGNALBIAS_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_SIGNALSAMPLING_HIGH_RATE As Integer = 0
  Public Const Y_SIGNALSAMPLING_HIGH_RATE_FILTERED As Integer = 1
  Public Const Y_SIGNALSAMPLING_LOW_NOISE As Integer = 2
  Public Const Y_SIGNALSAMPLING_LOW_NOISE_FILTERED As Integer = 3
  Public Const Y_SIGNALSAMPLING_HIGHEST_RATE As Integer = 4
  Public Const Y_SIGNALSAMPLING_AC As Integer = 5
  Public Const Y_SIGNALSAMPLING_INVALID As Integer = -1
  Public Const Y_ENABLED_FALSE As Integer = 0
  Public Const Y_ENABLED_TRUE As Integer = 1
  Public Const Y_ENABLED_INVALID As Integer = -1
  Public Delegate Sub YGenericSensorValueCallback(ByVal func As YGenericSensor, ByVal value As String)
  Public Delegate Sub YGenericSensorTimedReportCallback(ByVal func As YGenericSensor, ByVal measure As YMeasure)
  REM --- (end of YGenericSensor globals)

  REM --- (YGenericSensor class start)

  '''*
  ''' <summary>
  '''   The <c>YGenericSensor</c> class allows you to read and configure Yoctopuce signal
  '''   transducers.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measures,
  '''   to register callback functions, to access the autonomous datalogger.
  '''   This class adds the ability to configure the automatic conversion between the
  '''   measured signal and the corresponding engineering unit.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YGenericSensor
    Inherits YSensor
    REM --- (end of YGenericSensor class start)

    REM --- (YGenericSensor definitions)
    Public Const SIGNALVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const SIGNALUNIT_INVALID As String = YAPI.INVALID_STRING
    Public Const SIGNALRANGE_INVALID As String = YAPI.INVALID_STRING
    Public Const VALUERANGE_INVALID As String = YAPI.INVALID_STRING
    Public Const SIGNALBIAS_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const SIGNALSAMPLING_HIGH_RATE As Integer = 0
    Public Const SIGNALSAMPLING_HIGH_RATE_FILTERED As Integer = 1
    Public Const SIGNALSAMPLING_LOW_NOISE As Integer = 2
    Public Const SIGNALSAMPLING_LOW_NOISE_FILTERED As Integer = 3
    Public Const SIGNALSAMPLING_HIGHEST_RATE As Integer = 4
    Public Const SIGNALSAMPLING_AC As Integer = 5
    Public Const SIGNALSAMPLING_INVALID As Integer = -1
    Public Const ENABLED_FALSE As Integer = 0
    Public Const ENABLED_TRUE As Integer = 1
    Public Const ENABLED_INVALID As Integer = -1
    REM --- (end of YGenericSensor definitions)

    REM --- (YGenericSensor attributes declaration)
    Protected _signalValue As Double
    Protected _signalUnit As String
    Protected _signalRange As String
    Protected _valueRange As String
    Protected _signalBias As Double
    Protected _signalSampling As Integer
    Protected _enabled As Integer
    Protected _valueCallbackGenericSensor As YGenericSensorValueCallback
    Protected _timedReportCallbackGenericSensor As YGenericSensorTimedReportCallback
    REM --- (end of YGenericSensor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "GenericSensor"
      REM --- (YGenericSensor attributes initialization)
      _signalValue = SIGNALVALUE_INVALID
      _signalUnit = SIGNALUNIT_INVALID
      _signalRange = SIGNALRANGE_INVALID
      _valueRange = VALUERANGE_INVALID
      _signalBias = SIGNALBIAS_INVALID
      _signalSampling = SIGNALSAMPLING_INVALID
      _enabled = ENABLED_INVALID
      _valueCallbackGenericSensor = Nothing
      _timedReportCallbackGenericSensor = Nothing
      REM --- (end of YGenericSensor attributes initialization)
    End Sub

    REM --- (YGenericSensor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("signalValue") Then
        _signalValue = Math.Round(json_val.getDouble("signalValue") / 65.536) / 1000.0
      End If
      If json_val.has("signalUnit") Then
        _signalUnit = json_val.getString("signalUnit")
      End If
      If json_val.has("signalRange") Then
        _signalRange = json_val.getString("signalRange")
      End If
      If json_val.has("valueRange") Then
        _valueRange = json_val.getString("valueRange")
      End If
      If json_val.has("signalBias") Then
        _signalBias = Math.Round(json_val.getDouble("signalBias") / 65.536) / 1000.0
      End If
      If json_val.has("signalSampling") Then
        _signalSampling = CInt(json_val.getLong("signalSampling"))
      End If
      If json_val.has("enabled") Then
        If (json_val.getInt("enabled") > 0) Then _enabled = 1 Else _enabled = 0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YGenericSensor private methods declaration)

    REM --- (YGenericSensor public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the measuring unit for the measured value.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the measuring unit for the measured value
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
    Public Function set_unit(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("unit", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current value of the electrical signal measured by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current value of the electrical signal measured by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YGenericSensor.SIGNALVALUE_INVALID</c>.
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
    '''   Returns the measuring unit of the electrical signal used by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the measuring unit of the electrical signal used by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YGenericSensor.SIGNALUNIT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signalUnit() As String
      Dim res As String
      If (Me._cacheExpiration = 0) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SIGNALUNIT_INVALID
        End If
      End If
      res = Me._signalUnit
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the input signal range used by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the input signal range used by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YGenericSensor.SIGNALRANGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signalRange() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SIGNALRANGE_INVALID
        End If
      End If
      res = Me._signalRange
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the input signal range used by the sensor.
    ''' <para>
    '''   When the input signal gets out of the planned range, the output value
    '''   will be set to an arbitrary large value, whose sign indicates the direction
    '''   of the range overrun.
    ''' </para>
    ''' <para>
    '''   For a 4-20mA sensor, the default input signal range is "4...20".
    '''   For a 0-10V sensor, the default input signal range is "0.1...10".
    '''   For numeric communication interfaces, the default input signal range is
    '''   "-999999.999...999999.999".
    ''' </para>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the input signal range used by the sensor
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
    Public Function set_signalRange(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("signalRange", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the physical value range measured by the sensor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the physical value range measured by the sensor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YGenericSensor.VALUERANGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_valueRange() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return VALUERANGE_INVALID
        End If
      End If
      res = Me._valueRange
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the output value range, corresponding to the physical value measured
    '''   by the sensor.
    ''' <para>
    '''   The default output value range is the same as the input signal
    '''   range (1:1 mapping), but you can change it so that the function automatically
    '''   computes the physical value encoded by the input signal. Be aware that, as a
    '''   side effect, the range modification may automatically modify the display resolution.
    ''' </para>
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the output value range, corresponding to the physical value measured
    '''   by the sensor
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
    Public Function set_valueRange(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("valueRange", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the electric signal bias for zero shift adjustment.
    ''' <para>
    '''   If your electric signal reads positive when it should be zero, set up
    '''   a positive signalBias of the same value to fix the zero shift.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the electric signal bias for zero shift adjustment
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
    Public Function set_signalBias(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("signalBias", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the electric signal bias for zero shift adjustment.
    ''' <para>
    '''   A positive bias means that the signal is over-reporting the measure,
    '''   while a negative bias means that the signal is under-reporting the measure.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the electric signal bias for zero shift adjustment
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YGenericSensor.SIGNALBIAS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signalBias() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SIGNALBIAS_INVALID
        End If
      End If
      res = Me._signalBias
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the electric signal sampling method to use.
    ''' <para>
    '''   The <c>HIGH_RATE</c> method uses the highest sampling frequency, without any filtering.
    '''   The <c>HIGH_RATE_FILTERED</c> method adds a windowed 7-sample median filter.
    '''   The <c>LOW_NOISE</c> method uses a reduced acquisition frequency to reduce noise.
    '''   The <c>LOW_NOISE_FILTERED</c> method combines a reduced frequency with the median filter
    '''   to get measures as stable as possible when working on a noisy signal.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YGenericSensor.SIGNALSAMPLING_HIGH_RATE</c>,
    '''   <c>YGenericSensor.SIGNALSAMPLING_HIGH_RATE_FILTERED</c>, <c>YGenericSensor.SIGNALSAMPLING_LOW_NOISE</c>,
    '''   <c>YGenericSensor.SIGNALSAMPLING_LOW_NOISE_FILTERED</c>, <c>YGenericSensor.SIGNALSAMPLING_HIGHEST_RATE</c>
    '''   and <c>YGenericSensor.SIGNALSAMPLING_AC</c> corresponding to the electric signal sampling method to use
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YGenericSensor.SIGNALSAMPLING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_signalSampling() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SIGNALSAMPLING_INVALID
        End If
      End If
      res = Me._signalSampling
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the electric signal sampling method to use.
    ''' <para>
    '''   The <c>HIGH_RATE</c> method uses the highest sampling frequency, without any filtering.
    '''   The <c>HIGH_RATE_FILTERED</c> method adds a windowed 7-sample median filter.
    '''   The <c>LOW_NOISE</c> method uses a reduced acquisition frequency to reduce noise.
    '''   The <c>LOW_NOISE_FILTERED</c> method combines a reduced frequency with the median filter
    '''   to get measures as stable as possible when working on a noisy signal.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>YGenericSensor.SIGNALSAMPLING_HIGH_RATE</c>,
    '''   <c>YGenericSensor.SIGNALSAMPLING_HIGH_RATE_FILTERED</c>, <c>YGenericSensor.SIGNALSAMPLING_LOW_NOISE</c>,
    '''   <c>YGenericSensor.SIGNALSAMPLING_LOW_NOISE_FILTERED</c>, <c>YGenericSensor.SIGNALSAMPLING_HIGHEST_RATE</c>
    '''   and <c>YGenericSensor.SIGNALSAMPLING_AC</c> corresponding to the electric signal sampling method to use
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
    Public Function set_signalSampling(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("signalSampling", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the activation state of this input.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YGenericSensor.ENABLED_FALSE</c> or <c>YGenericSensor.ENABLED_TRUE</c>, according to the
    '''   activation state of this input
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YGenericSensor.ENABLED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_enabled() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ENABLED_INVALID
        End If
      End If
      res = Me._enabled
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the activation state of this input.
    ''' <para>
    '''   When an input is disabled,
    '''   its value is no more updated. On some devices, disabling an input can
    '''   improve the refresh rate of the other active inputs.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YGenericSensor.ENABLED_FALSE</c> or <c>YGenericSensor.ENABLED_TRUE</c>, according to the
    '''   activation state of this input
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
    Public Function set_enabled(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("enabled", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a generic sensor for a given identifier.
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
    '''   This function does not require that the generic sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YGenericSensor.isOnline()</c> to test if the generic sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a generic sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the generic sensor, for instance
    '''   <c>RX010V01.genericSensor1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YGenericSensor</c> object allowing you to drive the generic sensor.
    ''' </returns>
    '''/
    Public Shared Function FindGenericSensor(func As String) As YGenericSensor
      Dim obj As YGenericSensor
      obj = CType(YFunction._FindFromCache("GenericSensor", func), YGenericSensor)
      If ((obj Is Nothing)) Then
        obj = New YGenericSensor(func)
        YFunction._AddToCache("GenericSensor", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YGenericSensorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackGenericSensor = callback
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
      If (Not (Me._valueCallbackGenericSensor Is Nothing)) Then
        Me._valueCallbackGenericSensor(Me, value)
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
    Public Overloads Function registerTimedReportCallback(callback As YGenericSensorTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackGenericSensor = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackGenericSensor Is Nothing)) Then
        Me._timedReportCallbackGenericSensor(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Adjusts the signal bias so that the current signal value is need
    '''   precisely as zero.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function zeroAdjust() As Integer
      Dim currSignal As Double = 0
      Dim currBias As Double = 0
      currSignal = Me.get_signalValue()
      currBias = Me.get_signalBias()
      Return Me.set_signalBias(currSignal + currBias)
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of generic sensors started using <c>yFirstGenericSensor()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned generic sensors order.
    '''   If you want to find a specific a generic sensor, use <c>GenericSensor.findGenericSensor()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGenericSensor</c> object, corresponding to
    '''   a generic sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more generic sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextGenericSensor() As YGenericSensor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YGenericSensor.FindGenericSensor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of generic sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YGenericSensor.nextGenericSensor()</c> to iterate on
    '''   next generic sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YGenericSensor</c> object, corresponding to
    '''   the first generic sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstGenericSensor() As YGenericSensor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("GenericSensor", 0, p, size, neededsize, errmsg)
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
      Return YGenericSensor.FindGenericSensor(serial + "." + funcId)
    End Function

    REM --- (end of YGenericSensor public methods declaration)

  End Class

  REM --- (YGenericSensor functions)

  '''*
  ''' <summary>
  '''   Retrieves a generic sensor for a given identifier.
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
  '''   This function does not require that the generic sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YGenericSensor.isOnline()</c> to test if the generic sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a generic sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the generic sensor, for instance
  '''   <c>RX010V01.genericSensor1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YGenericSensor</c> object allowing you to drive the generic sensor.
  ''' </returns>
  '''/
  Public Function yFindGenericSensor(ByVal func As String) As YGenericSensor
    Return YGenericSensor.FindGenericSensor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of generic sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YGenericSensor.nextGenericSensor()</c> to iterate on
  '''   next generic sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YGenericSensor</c> object, corresponding to
  '''   the first generic sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstGenericSensor() As YGenericSensor
    Return YGenericSensor.FirstGenericSensor()
  End Function


  REM --- (end of YGenericSensor functions)

End Module

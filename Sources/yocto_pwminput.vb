' ********************************************************************
'
'  $Id: yocto_pwminput.vb 37149 2019-09-12 21:24:53Z mvuilleu $
'
'  Implements yFindPwmInput(), the high-level API for PwmInput functions
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

Module yocto_pwminput

    REM --- (YPwmInput return codes)
    REM --- (end of YPwmInput return codes)
    REM --- (YPwmInput dlldef)
    REM --- (end of YPwmInput dlldef)
   REM --- (YPwmInput yapiwrapper)
   REM --- (end of YPwmInput yapiwrapper)
  REM --- (YPwmInput globals)

  Public Const Y_DUTYCYCLE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_PULSEDURATION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_FREQUENCY_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_PERIOD_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_PULSECOUNTER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PWMREPORTMODE_PWM_DUTYCYCLE As Integer = 0
  Public Const Y_PWMREPORTMODE_PWM_FREQUENCY As Integer = 1
  Public Const Y_PWMREPORTMODE_PWM_PULSEDURATION As Integer = 2
  Public Const Y_PWMREPORTMODE_PWM_EDGECOUNT As Integer = 3
  Public Const Y_PWMREPORTMODE_PWM_PULSECOUNT As Integer = 4
  Public Const Y_PWMREPORTMODE_PWM_CPS As Integer = 5
  Public Const Y_PWMREPORTMODE_PWM_CPM As Integer = 6
  Public Const Y_PWMREPORTMODE_PWM_STATE As Integer = 7
  Public Const Y_PWMREPORTMODE_PWM_FREQ_CPS As Integer = 8
  Public Const Y_PWMREPORTMODE_PWM_FREQ_CPM As Integer = 9
  Public Const Y_PWMREPORTMODE_INVALID As Integer = -1
  Public Const Y_DEBOUNCEPERIOD_INVALID As Integer = YAPI.INVALID_UINT
  Public Delegate Sub YPwmInputValueCallback(ByVal func As YPwmInput, ByVal value As String)
  Public Delegate Sub YPwmInputTimedReportCallback(ByVal func As YPwmInput, ByVal measure As YMeasure)
  REM --- (end of YPwmInput globals)

  REM --- (YPwmInput class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce class YPwmInput allows you to read and configure Yoctopuce PWM
  '''   sensors.
  ''' <para>
  '''   It inherits from YSensor class the core functions to read measurements,
  '''   to register callback functions, to access the autonomous datalogger.
  '''   This class adds the ability to configure the signal parameter used to transmit
  '''   information: the duty cycle, the frequency or the pulse width.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YPwmInput
    Inherits YSensor
    REM --- (end of YPwmInput class start)

    REM --- (YPwmInput definitions)
    Public Const DUTYCYCLE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const PULSEDURATION_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const FREQUENCY_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const PERIOD_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const PULSECOUNTER_INVALID As Long = YAPI.INVALID_LONG
    Public Const PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
    Public Const PWMREPORTMODE_PWM_DUTYCYCLE As Integer = 0
    Public Const PWMREPORTMODE_PWM_FREQUENCY As Integer = 1
    Public Const PWMREPORTMODE_PWM_PULSEDURATION As Integer = 2
    Public Const PWMREPORTMODE_PWM_EDGECOUNT As Integer = 3
    Public Const PWMREPORTMODE_PWM_PULSECOUNT As Integer = 4
    Public Const PWMREPORTMODE_PWM_CPS As Integer = 5
    Public Const PWMREPORTMODE_PWM_CPM As Integer = 6
    Public Const PWMREPORTMODE_PWM_STATE As Integer = 7
    Public Const PWMREPORTMODE_PWM_FREQ_CPS As Integer = 8
    Public Const PWMREPORTMODE_PWM_FREQ_CPM As Integer = 9
    Public Const PWMREPORTMODE_INVALID As Integer = -1
    Public Const DEBOUNCEPERIOD_INVALID As Integer = YAPI.INVALID_UINT
    REM --- (end of YPwmInput definitions)

    REM --- (YPwmInput attributes declaration)
    Protected _dutyCycle As Double
    Protected _pulseDuration As Double
    Protected _frequency As Double
    Protected _period As Double
    Protected _pulseCounter As Long
    Protected _pulseTimer As Long
    Protected _pwmReportMode As Integer
    Protected _debouncePeriod As Integer
    Protected _valueCallbackPwmInput As YPwmInputValueCallback
    Protected _timedReportCallbackPwmInput As YPwmInputTimedReportCallback
    REM --- (end of YPwmInput attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "PwmInput"
      REM --- (YPwmInput attributes initialization)
      _dutyCycle = DUTYCYCLE_INVALID
      _pulseDuration = PULSEDURATION_INVALID
      _frequency = FREQUENCY_INVALID
      _period = PERIOD_INVALID
      _pulseCounter = PULSECOUNTER_INVALID
      _pulseTimer = PULSETIMER_INVALID
      _pwmReportMode = PWMREPORTMODE_INVALID
      _debouncePeriod = DEBOUNCEPERIOD_INVALID
      _valueCallbackPwmInput = Nothing
      _timedReportCallbackPwmInput = Nothing
      REM --- (end of YPwmInput attributes initialization)
    End Sub

    REM --- (YPwmInput private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("dutyCycle") Then
        _dutyCycle = Math.Round(json_val.getDouble("dutyCycle") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("pulseDuration") Then
        _pulseDuration = Math.Round(json_val.getDouble("pulseDuration") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("frequency") Then
        _frequency = Math.Round(json_val.getDouble("frequency") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("period") Then
        _period = Math.Round(json_val.getDouble("period") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("pulseCounter") Then
        _pulseCounter = json_val.getLong("pulseCounter")
      End If
      If json_val.has("pulseTimer") Then
        _pulseTimer = json_val.getLong("pulseTimer")
      End If
      If json_val.has("pwmReportMode") Then
        _pwmReportMode = CInt(json_val.getLong("pwmReportMode"))
      End If
      If json_val.has("debouncePeriod") Then
        _debouncePeriod = CInt(json_val.getLong("debouncePeriod"))
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YPwmInput private methods declaration)

    REM --- (YPwmInput public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the measuring unit for the measured quantity.
    ''' <para>
    '''   That unit
    '''   is just a string which is automatically initialized each time
    '''   the measurement mode is changed. But is can be set to an
    '''   arbitrary value.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the measuring unit for the measured quantity
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
    Public Function set_unit(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("unit", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the PWM duty cycle, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the PWM duty cycle, in per cents
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DUTYCYCLE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_dutyCycle() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DUTYCYCLE_INVALID
        End If
      End If
      res = Me._dutyCycle
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the PWM pulse length in milliseconds, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the PWM pulse length in milliseconds, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PULSEDURATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pulseDuration() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PULSEDURATION_INVALID
        End If
      End If
      res = Me._pulseDuration
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the PWM frequency in Hz.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the PWM frequency in Hz
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_FREQUENCY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_frequency() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return FREQUENCY_INVALID
        End If
      End If
      res = Me._frequency
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the PWM period in milliseconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the PWM period in milliseconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PERIOD_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_period() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PERIOD_INVALID
        End If
      End If
      res = Me._period
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the pulse counter value.
    ''' <para>
    '''   Actually that
    '''   counter is incremented twice per period. That counter is
    '''   limited  to 1 billion.
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
    '''   Returns the timer of the pulses counter (ms).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the timer of the pulses counter (ms)
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
    '''   Returns the parameter (frequency/duty cycle, pulse width, edges count) returned by the get_currentValue function and callbacks.
    ''' <para>
    '''   Attention
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_PWMREPORTMODE_PWM_DUTYCYCLE</c>, <c>Y_PWMREPORTMODE_PWM_FREQUENCY</c>,
    '''   <c>Y_PWMREPORTMODE_PWM_PULSEDURATION</c>, <c>Y_PWMREPORTMODE_PWM_EDGECOUNT</c>,
    '''   <c>Y_PWMREPORTMODE_PWM_PULSECOUNT</c>, <c>Y_PWMREPORTMODE_PWM_CPS</c>,
    '''   <c>Y_PWMREPORTMODE_PWM_CPM</c>, <c>Y_PWMREPORTMODE_PWM_STATE</c>,
    '''   <c>Y_PWMREPORTMODE_PWM_FREQ_CPS</c> and <c>Y_PWMREPORTMODE_PWM_FREQ_CPM</c> corresponding to the
    '''   parameter (frequency/duty cycle, pulse width, edges count) returned by the get_currentValue
    '''   function and callbacks
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PWMREPORTMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pwmReportMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PWMREPORTMODE_INVALID
        End If
      End If
      res = Me._pwmReportMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the  parameter  type (frequency/duty cycle, pulse width, or edge count) returned by the get_currentValue function and callbacks.
    ''' <para>
    '''   The edge count value is limited to the 6 lowest digits. For values greater than one million, use
    '''   get_pulseCounter().
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_PWMREPORTMODE_PWM_DUTYCYCLE</c>, <c>Y_PWMREPORTMODE_PWM_FREQUENCY</c>,
    '''   <c>Y_PWMREPORTMODE_PWM_PULSEDURATION</c>, <c>Y_PWMREPORTMODE_PWM_EDGECOUNT</c>,
    '''   <c>Y_PWMREPORTMODE_PWM_PULSECOUNT</c>, <c>Y_PWMREPORTMODE_PWM_CPS</c>,
    '''   <c>Y_PWMREPORTMODE_PWM_CPM</c>, <c>Y_PWMREPORTMODE_PWM_STATE</c>,
    '''   <c>Y_PWMREPORTMODE_PWM_FREQ_CPS</c> and <c>Y_PWMREPORTMODE_PWM_FREQ_CPM</c> corresponding to the 
    '''   parameter  type (frequency/duty cycle, pulse width, or edge count) returned by the get_currentValue
    '''   function and callbacks
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
    Public Function set_pwmReportMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pwmReportMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the shortest expected pulse duration, in ms.
    ''' <para>
    '''   Any shorter pulse will be automatically ignored (debounce).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the shortest expected pulse duration, in ms
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DEBOUNCEPERIOD_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_debouncePeriod() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DEBOUNCEPERIOD_INVALID
        End If
      End If
      res = Me._debouncePeriod
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the shortest expected pulse duration, in ms.
    ''' <para>
    '''   Any shorter pulse will be automatically ignored (debounce).
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the shortest expected pulse duration, in ms
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
    Public Function set_debouncePeriod(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("debouncePeriod", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a PWM input for a given identifier.
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
    '''   This function does not require that the PWM input is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YPwmInput.isOnline()</c> to test if the PWM input is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a PWM input by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the PWM input
    ''' </param>
    ''' <returns>
    '''   a <c>YPwmInput</c> object allowing you to drive the PWM input.
    ''' </returns>
    '''/
    Public Shared Function FindPwmInput(func As String) As YPwmInput
      Dim obj As YPwmInput
      obj = CType(YFunction._FindFromCache("PwmInput", func), YPwmInput)
      If ((obj Is Nothing)) Then
        obj = New YPwmInput(func)
        YFunction._AddToCache("PwmInput", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YPwmInputValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackPwmInput = callback
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
      If (Not (Me._valueCallbackPwmInput Is Nothing)) Then
        Me._valueCallbackPwmInput(Me, value)
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
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YPwmInputTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackPwmInput = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackPwmInput Is Nothing)) Then
        Me._timedReportCallbackPwmInput(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the pulse counter value as well as its timer.
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
    '''   Continues the enumeration of PWM inputs started using <c>yFirstPwmInput()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned PWM inputs order.
    '''   If you want to find a specific a PWM input, use <c>PwmInput.findPwmInput()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPwmInput</c> object, corresponding to
    '''   a PWM input currently online, or a <c>Nothing</c> pointer
    '''   if there are no more PWM inputs to enumerate.
    ''' </returns>
    '''/
    Public Function nextPwmInput() As YPwmInput
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YPwmInput.FindPwmInput(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of PWM inputs currently accessible.
    ''' <para>
    '''   Use the method <c>YPwmInput.nextPwmInput()</c> to iterate on
    '''   next PWM inputs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPwmInput</c> object, corresponding to
    '''   the first PWM input currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstPwmInput() As YPwmInput
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("PwmInput", 0, p, size, neededsize, errmsg)
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
      Return YPwmInput.FindPwmInput(serial + "." + funcId)
    End Function

    REM --- (end of YPwmInput public methods declaration)

  End Class

  REM --- (YPwmInput functions)

  '''*
  ''' <summary>
  '''   Retrieves a PWM input for a given identifier.
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
  '''   This function does not require that the PWM input is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YPwmInput.isOnline()</c> to test if the PWM input is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a PWM input by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the PWM input
  ''' </param>
  ''' <returns>
  '''   a <c>YPwmInput</c> object allowing you to drive the PWM input.
  ''' </returns>
  '''/
  Public Function yFindPwmInput(ByVal func As String) As YPwmInput
    Return YPwmInput.FindPwmInput(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of PWM inputs currently accessible.
  ''' <para>
  '''   Use the method <c>YPwmInput.nextPwmInput()</c> to iterate on
  '''   next PWM inputs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YPwmInput</c> object, corresponding to
  '''   the first PWM input currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstPwmInput() As YPwmInput
    Return YPwmInput.FirstPwmInput()
  End Function


  REM --- (end of YPwmInput functions)

End Module

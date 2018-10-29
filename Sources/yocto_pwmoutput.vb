' ********************************************************************
'
'  $Id: yocto_pwmoutput.vb 32610 2018-10-10 06:52:20Z seb $
'
'  Implements yFindPwmOutput(), the high-level API for PwmOutput functions
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

Module yocto_pwmoutput

    REM --- (YPwmOutput return codes)
    REM --- (end of YPwmOutput return codes)
    REM --- (YPwmOutput dlldef)
    REM --- (end of YPwmOutput dlldef)
   REM --- (YPwmOutput yapiwrapper)
   REM --- (end of YPwmOutput yapiwrapper)
  REM --- (YPwmOutput globals)

  Public Const Y_ENABLED_FALSE As Integer = 0
  Public Const Y_ENABLED_TRUE As Integer = 1
  Public Const Y_ENABLED_INVALID As Integer = -1
  Public Const Y_FREQUENCY_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_PERIOD_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_DUTYCYCLE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_PULSEDURATION_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_PWMTRANSITION_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ENABLEDATPOWERON_FALSE As Integer = 0
  Public Const Y_ENABLEDATPOWERON_TRUE As Integer = 1
  Public Const Y_ENABLEDATPOWERON_INVALID As Integer = -1
  Public Const Y_DUTYCYCLEATPOWERON_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YPwmOutputValueCallback(ByVal func As YPwmOutput, ByVal value As String)
  Public Delegate Sub YPwmOutputTimedReportCallback(ByVal func As YPwmOutput, ByVal measure As YMeasure)
  REM --- (end of YPwmOutput globals)

  REM --- (YPwmOutput class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce application programming interface allows you to configure, start, and stop the PWM.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YPwmOutput
    Inherits YFunction
    REM --- (end of YPwmOutput class start)

    REM --- (YPwmOutput definitions)
    Public Const ENABLED_FALSE As Integer = 0
    Public Const ENABLED_TRUE As Integer = 1
    Public Const ENABLED_INVALID As Integer = -1
    Public Const FREQUENCY_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const PERIOD_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const DUTYCYCLE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const PULSEDURATION_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const PWMTRANSITION_INVALID As String = YAPI.INVALID_STRING
    Public Const ENABLEDATPOWERON_FALSE As Integer = 0
    Public Const ENABLEDATPOWERON_TRUE As Integer = 1
    Public Const ENABLEDATPOWERON_INVALID As Integer = -1
    Public Const DUTYCYCLEATPOWERON_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YPwmOutput definitions)

    REM --- (YPwmOutput attributes declaration)
    Protected _enabled As Integer
    Protected _frequency As Double
    Protected _period As Double
    Protected _dutyCycle As Double
    Protected _pulseDuration As Double
    Protected _pwmTransition As String
    Protected _enabledAtPowerOn As Integer
    Protected _dutyCycleAtPowerOn As Double
    Protected _valueCallbackPwmOutput As YPwmOutputValueCallback
    REM --- (end of YPwmOutput attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "PwmOutput"
      REM --- (YPwmOutput attributes initialization)
      _enabled = ENABLED_INVALID
      _frequency = FREQUENCY_INVALID
      _period = PERIOD_INVALID
      _dutyCycle = DUTYCYCLE_INVALID
      _pulseDuration = PULSEDURATION_INVALID
      _pwmTransition = PWMTRANSITION_INVALID
      _enabledAtPowerOn = ENABLEDATPOWERON_INVALID
      _dutyCycleAtPowerOn = DUTYCYCLEATPOWERON_INVALID
      _valueCallbackPwmOutput = Nothing
      REM --- (end of YPwmOutput attributes initialization)
    End Sub

    REM --- (YPwmOutput private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("enabled") Then
        If (json_val.getInt("enabled") > 0) Then _enabled = 1 Else _enabled = 0
      End If
      If json_val.has("frequency") Then
        _frequency = Math.Round(json_val.getDouble("frequency") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("period") Then
        _period = Math.Round(json_val.getDouble("period") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("dutyCycle") Then
        _dutyCycle = Math.Round(json_val.getDouble("dutyCycle") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("pulseDuration") Then
        _pulseDuration = Math.Round(json_val.getDouble("pulseDuration") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("pwmTransition") Then
        _pwmTransition = json_val.getString("pwmTransition")
      End If
      If json_val.has("enabledAtPowerOn") Then
        If (json_val.getInt("enabledAtPowerOn") > 0) Then _enabledAtPowerOn = 1 Else _enabledAtPowerOn = 0
      End If
      If json_val.has("dutyCycleAtPowerOn") Then
        _dutyCycleAtPowerOn = Math.Round(json_val.getDouble("dutyCycleAtPowerOn") * 1000.0 / 65536.0) / 1000.0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YPwmOutput private methods declaration)

    REM --- (YPwmOutput public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the state of the PWMs.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>, according to the state of the PWMs
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ENABLED_INVALID</c>.
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
    '''   Stops or starts the PWM.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_ENABLED_FALSE</c> or <c>Y_ENABLED_TRUE</c>
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
    Public Function set_enabled(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("enabled", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the PWM frequency.
    ''' <para>
    '''   The duty cycle is kept unchanged thanks to an
    '''   automatic pulse width change.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the PWM frequency
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
    Public Function set_frequency(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("frequency", rest_val)
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
    '''   Changes the PWM period in milliseconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the PWM period in milliseconds
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
    Public Function set_period(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("period", rest_val)
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
    '''   Changes the PWM duty cycle, in per cents.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the PWM duty cycle, in per cents
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
    Public Function set_dutyCycle(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("dutyCycle", rest_val)
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
    '''   Changes the PWM pulse length, in milliseconds.
    ''' <para>
    '''   A pulse length cannot be longer than period, otherwise it is truncated.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the PWM pulse length, in milliseconds
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
    Public Function set_pulseDuration(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("pulseDuration", rest_val)
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

    Public Function get_pwmTransition() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PWMTRANSITION_INVALID
        End If
      End If
      res = Me._pwmTransition
      Return res
    End Function


    Public Function set_pwmTransition(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("pwmTransition", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the state of the PWM at device power on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ENABLEDATPOWERON_FALSE</c> or <c>Y_ENABLEDATPOWERON_TRUE</c>, according to the state of
    '''   the PWM at device power on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ENABLEDATPOWERON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_enabledAtPowerOn() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ENABLEDATPOWERON_INVALID
        End If
      End If
      res = Me._enabledAtPowerOn
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the state of the PWM at device power on.
    ''' <para>
    '''   Remember to call the matching module <c>saveToFlash()</c>
    '''   method, otherwise this call will have no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_ENABLEDATPOWERON_FALSE</c> or <c>Y_ENABLEDATPOWERON_TRUE</c>, according to the state of
    '''   the PWM at device power on
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
    Public Function set_enabledAtPowerOn(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("enabledAtPowerOn", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the PWM duty cycle at device power on.
    ''' <para>
    '''   Remember to call the matching
    '''   module <c>saveToFlash()</c> method, otherwise this call will have no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the PWM duty cycle at device power on
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
    Public Function set_dutyCycleAtPowerOn(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("dutyCycleAtPowerOn", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the PWMs duty cycle at device power on as a floating point number between 0 and 100.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the PWMs duty cycle at device power on as a floating point
    '''   number between 0 and 100
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DUTYCYCLEATPOWERON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_dutyCycleAtPowerOn() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DUTYCYCLEATPOWERON_INVALID
        End If
      End If
      res = Me._dutyCycleAtPowerOn
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a PWM for a given identifier.
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
    '''   This function does not require that the PWM is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YPwmOutput.isOnline()</c> to test if the PWM is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a PWM by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the PWM
    ''' </param>
    ''' <returns>
    '''   a <c>YPwmOutput</c> object allowing you to drive the PWM.
    ''' </returns>
    '''/
    Public Shared Function FindPwmOutput(func As String) As YPwmOutput
      Dim obj As YPwmOutput
      obj = CType(YFunction._FindFromCache("PwmOutput", func), YPwmOutput)
      If ((obj Is Nothing)) Then
        obj = New YPwmOutput(func)
        YFunction._AddToCache("PwmOutput", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YPwmOutputValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackPwmOutput = callback
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
      If (Not (Me._valueCallbackPwmOutput Is Nothing)) Then
        Me._valueCallbackPwmOutput(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth transistion of the pulse duration toward a given value.
    ''' <para>
    '''   Any period, frequency, duty cycle or pulse width change will cancel any ongoing transition process.
    ''' </para>
    ''' </summary>
    ''' <param name="ms_target">
    '''   new pulse duration at the end of the transition
    '''   (floating-point number, representing the pulse duration in milliseconds)
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the transition, in milliseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function pulseDurationMove(ms_target As Double, ms_duration As Integer) As Integer
      Dim newval As String
      If (ms_target < 0.0) Then
        ms_target = 0.0
      End If
      newval = "" + Convert.ToString( CType(Math.Round(ms_target*65536), Integer)) + "ms:" + Convert.ToString(ms_duration)
      Return Me.set_pwmTransition(newval)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth change of the duty cycle toward a given value.
    ''' <para>
    '''   Any period, frequency, duty cycle or pulse width change will cancel any ongoing transition process.
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   new duty cycle at the end of the transition
    '''   (percentage, floating-point number between 0 and 100)
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the transition, in milliseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function dutyCycleMove(target As Double, ms_duration As Integer) As Integer
      Dim newval As String
      If (target < 0.0) Then
        target = 0.0
      End If
      If (target > 100.0) Then
        target = 100.0
      End If
      newval = "" + Convert.ToString( CType(Math.Round(target*65536), Integer)) + ":" + Convert.ToString(ms_duration)
      Return Me.set_pwmTransition(newval)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth frequency change toward a given value.
    ''' <para>
    '''   Any period, frequency, duty cycle or pulse width change will cancel any ongoing transition process.
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   new freuency at the end of the transition (floating-point number)
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the transition, in milliseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function frequencyMove(target As Double, ms_duration As Integer) As Integer
      Dim newval As String
      If (target < 0.001) Then
        target = 0.001
      End If
      newval = "" + YAPI._floatToStr( target) + "Hz:" + Convert.ToString(ms_duration)
      Return Me.set_pwmTransition(newval)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth transition toward a specified value of the phase shift between this channel
    '''   and the other channel.
    ''' <para>
    '''   The phase shift is executed by slightly changing the frequency
    '''   temporarily during the specified duration. This function only makes sense when both channels
    '''   are running, either at the same frequency, or at a multiple of the channel frequency.
    '''   Any period, frequency, duty cycle or pulse width change will cancel any ongoing transition process.
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   phase shift at the end of the transition, in milliseconds (floating-point number)
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the transition, in milliseconds
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function phaseMove(target As Double, ms_duration As Integer) As Integer
      Dim newval As String
      newval = "" + YAPI._floatToStr( target) + "ps:" + Convert.ToString(ms_duration)
      Return Me.set_pwmTransition(newval)
    End Function

    '''*
    ''' <summary>
    '''   Trigger a given number of pulses of specified duration, at current frequency.
    ''' <para>
    '''   At the end of the pulse train, revert to the original state of the PWM generator.
    ''' </para>
    ''' </summary>
    ''' <param name="ms_target">
    '''   desired pulse duration
    '''   (floating-point number, representing the pulse duration in milliseconds)
    ''' </param>
    ''' <param name="n_pulses">
    '''   desired pulse count
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function triggerPulsesByDuration(ms_target As Double, n_pulses As Integer) As Integer
      Dim newval As String
      If (ms_target < 0.0) Then
        ms_target = 0.0
      End If
      newval = "" + Convert.ToString( CType(Math.Round(ms_target*65536), Integer)) + "ms*" + Convert.ToString(n_pulses)
      Return Me.set_pwmTransition(newval)
    End Function

    '''*
    ''' <summary>
    '''   Trigger a given number of pulses of specified duration, at current frequency.
    ''' <para>
    '''   At the end of the pulse train, revert to the original state of the PWM generator.
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   desired duty cycle for the generated pulses
    '''   (percentage, floating-point number between 0 and 100)
    ''' </param>
    ''' <param name="n_pulses">
    '''   desired pulse count
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function triggerPulsesByDutyCycle(target As Double, n_pulses As Integer) As Integer
      Dim newval As String
      If (target < 0.0) Then
        target = 0.0
      End If
      If (target > 100.0) Then
        target = 100.0
      End If
      newval = "" + Convert.ToString( CType(Math.Round(target*65536), Integer)) + "*" + Convert.ToString(n_pulses)
      Return Me.set_pwmTransition(newval)
    End Function

    '''*
    ''' <summary>
    '''   Trigger a given number of pulses at the specified frequency, using current duty cycle.
    ''' <para>
    '''   At the end of the pulse train, revert to the original state of the PWM generator.
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   desired frequency for the generated pulses (floating-point number)
    ''' </param>
    ''' <param name="n_pulses">
    '''   desired pulse count
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> when the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function triggerPulsesByFrequency(target As Double, n_pulses As Integer) As Integer
      Dim newval As String
      If (target < 0.001) Then
        target = 0.001
      End If
      newval = "" + YAPI._floatToStr( target) + "Hz*" + Convert.ToString(n_pulses)
      Return Me.set_pwmTransition(newval)
    End Function

    Public Overridable Function markForRepeat() As Integer
      Return Me.set_pwmTransition(":")
    End Function

    Public Overridable Function repeatFromMark() As Integer
      Return Me.set_pwmTransition("R")
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of PWMs started using <c>yFirstPwmOutput()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPwmOutput</c> object, corresponding to
    '''   a PWM currently online, or a <c>Nothing</c> pointer
    '''   if there are no more PWMs to enumerate.
    ''' </returns>
    '''/
    Public Function nextPwmOutput() As YPwmOutput
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YPwmOutput.FindPwmOutput(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of PWMs currently accessible.
    ''' <para>
    '''   Use the method <c>YPwmOutput.nextPwmOutput()</c> to iterate on
    '''   next PWMs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YPwmOutput</c> object, corresponding to
    '''   the first PWM currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstPwmOutput() As YPwmOutput
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("PwmOutput", 0, p, size, neededsize, errmsg)
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
      Return YPwmOutput.FindPwmOutput(serial + "." + funcId)
    End Function

    REM --- (end of YPwmOutput public methods declaration)

  End Class

  REM --- (YPwmOutput functions)

  '''*
  ''' <summary>
  '''   Retrieves a PWM for a given identifier.
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
  '''   This function does not require that the PWM is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YPwmOutput.isOnline()</c> to test if the PWM is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a PWM by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the PWM
  ''' </param>
  ''' <returns>
  '''   a <c>YPwmOutput</c> object allowing you to drive the PWM.
  ''' </returns>
  '''/
  Public Function yFindPwmOutput(ByVal func As String) As YPwmOutput
    Return YPwmOutput.FindPwmOutput(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of PWMs currently accessible.
  ''' <para>
  '''   Use the method <c>YPwmOutput.nextPwmOutput()</c> to iterate on
  '''   next PWMs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YPwmOutput</c> object, corresponding to
  '''   the first PWM currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstPwmOutput() As YPwmOutput
    Return YPwmOutput.FirstPwmOutput()
  End Function


  REM --- (end of YPwmOutput functions)

End Module

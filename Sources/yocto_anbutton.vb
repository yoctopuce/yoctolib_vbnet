'*********************************************************************
'*
'* $Id: yocto_anbutton.vb 17674 2014-09-16 16:18:58Z seb $
'*
'* Implements yFindAnButton(), the high-level API for AnButton functions
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

Module yocto_anbutton

    REM --- (YAnButton return codes)
    REM --- (end of YAnButton return codes)
    REM --- (YAnButton dlldef)
    REM --- (end of YAnButton dlldef)
  REM --- (YAnButton globals)

  Public Const Y_CALIBRATEDVALUE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_RAWVALUE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ANALOGCALIBRATION_OFF As Integer = 0
  Public Const Y_ANALOGCALIBRATION_ON As Integer = 1
  Public Const Y_ANALOGCALIBRATION_INVALID As Integer = -1

  Public Const Y_CALIBRATIONMAX_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_CALIBRATIONMIN_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_SENSITIVITY_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ISPRESSED_FALSE As Integer = 0
  Public Const Y_ISPRESSED_TRUE As Integer = 1
  Public Const Y_ISPRESSED_INVALID As Integer = -1

  Public Const Y_LASTTIMEPRESSED_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_LASTTIMERELEASED_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PULSECOUNTER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
  Public Delegate Sub YAnButtonValueCallback(ByVal func As YAnButton, ByVal value As String)
  Public Delegate Sub YAnButtonTimedReportCallback(ByVal func As YAnButton, ByVal measure As YMeasure)
  REM --- (end of YAnButton globals)

  REM --- (YAnButton class start)

  '''*
  ''' <summary>
  '''   Yoctopuce application programming interface allows you to measure the state
  '''   of a simple button as well as to read an analog potentiometer (variable resistance).
  ''' <para>
  '''   This can be use for instance with a continuous rotating knob, a throttle grip
  '''   or a joystick. The module is capable to calibrate itself on min and max values,
  '''   in order to compute a calibrated value that varies proportionally with the
  '''   potentiometer position, regardless of its total resistance.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YAnButton
    Inherits YFunction
    REM --- (end of YAnButton class start)

    REM --- (YAnButton definitions)
    Public Const CALIBRATEDVALUE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const RAWVALUE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ANALOGCALIBRATION_OFF As Integer = 0
    Public Const ANALOGCALIBRATION_ON As Integer = 1
    Public Const ANALOGCALIBRATION_INVALID As Integer = -1

    Public Const CALIBRATIONMAX_INVALID As Integer = YAPI.INVALID_UINT
    Public Const CALIBRATIONMIN_INVALID As Integer = YAPI.INVALID_UINT
    Public Const SENSITIVITY_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ISPRESSED_FALSE As Integer = 0
    Public Const ISPRESSED_TRUE As Integer = 1
    Public Const ISPRESSED_INVALID As Integer = -1

    Public Const LASTTIMEPRESSED_INVALID As Long = YAPI.INVALID_LONG
    Public Const LASTTIMERELEASED_INVALID As Long = YAPI.INVALID_LONG
    Public Const PULSECOUNTER_INVALID As Long = YAPI.INVALID_LONG
    Public Const PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
    REM --- (end of YAnButton definitions)

    REM --- (YAnButton attributes declaration)
    Protected _calibratedValue As Integer
    Protected _rawValue As Integer
    Protected _analogCalibration As Integer
    Protected _calibrationMax As Integer
    Protected _calibrationMin As Integer
    Protected _sensitivity As Integer
    Protected _isPressed As Integer
    Protected _lastTimePressed As Long
    Protected _lastTimeReleased As Long
    Protected _pulseCounter As Long
    Protected _pulseTimer As Long
    Protected _valueCallbackAnButton As YAnButtonValueCallback
    REM --- (end of YAnButton attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "AnButton"
      REM --- (YAnButton attributes initialization)
      _calibratedValue = CALIBRATEDVALUE_INVALID
      _rawValue = RAWVALUE_INVALID
      _analogCalibration = ANALOGCALIBRATION_INVALID
      _calibrationMax = CALIBRATIONMAX_INVALID
      _calibrationMin = CALIBRATIONMIN_INVALID
      _sensitivity = SENSITIVITY_INVALID
      _isPressed = ISPRESSED_INVALID
      _lastTimePressed = LASTTIMEPRESSED_INVALID
      _lastTimeReleased = LASTTIMERELEASED_INVALID
      _pulseCounter = PULSECOUNTER_INVALID
      _pulseTimer = PULSETIMER_INVALID
      _valueCallbackAnButton = Nothing
      REM --- (end of YAnButton attributes initialization)
    End Sub

    REM --- (YAnButton private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "calibratedValue") Then
        _calibratedValue = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "rawValue") Then
        _rawValue = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "analogCalibration") Then
        If (member.ivalue > 0) Then _analogCalibration = 1 Else _analogCalibration = 0
        Return 1
      End If
      If (member.name = "calibrationMax") Then
        _calibrationMax = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "calibrationMin") Then
        _calibrationMin = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "sensitivity") Then
        _sensitivity = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "isPressed") Then
        If (member.ivalue > 0) Then _isPressed = 1 Else _isPressed = 0
        Return 1
      End If
      If (member.name = "lastTimePressed") Then
        _lastTimePressed = member.ivalue
        Return 1
      End If
      If (member.name = "lastTimeReleased") Then
        _lastTimeReleased = member.ivalue
        Return 1
      End If
      If (member.name = "pulseCounter") Then
        _pulseCounter = member.ivalue
        Return 1
      End If
      If (member.name = "pulseTimer") Then
        _pulseTimer = member.ivalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YAnButton private methods declaration)

    REM --- (YAnButton public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current calibrated input value (between 0 and 1000, included).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current calibrated input value (between 0 and 1000, included)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALIBRATEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_calibratedValue() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALIBRATEDVALUE_INVALID
        End If
      End If
      Return Me._calibratedValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the current measured input value as-is (between 0 and 4095, included).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current measured input value as-is (between 0 and 4095, included)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RAWVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rawValue() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RAWVALUE_INVALID
        End If
      End If
      Return Me._rawValue
    End Function

    '''*
    ''' <summary>
    '''   Tells if a calibration process is currently ongoing.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ANALOGCALIBRATION_OFF</c> or <c>Y_ANALOGCALIBRATION_ON</c>
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ANALOGCALIBRATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_analogCalibration() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ANALOGCALIBRATION_INVALID
        End If
      End If
      Return Me._analogCalibration
    End Function


    '''*
    ''' <summary>
    '''   Starts or stops the calibration process.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module at the end of the calibration if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_ANALOGCALIBRATION_OFF</c> or <c>Y_ANALOGCALIBRATION_ON</c>
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
    Public Function set_analogCalibration(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("analogCalibration", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the maximal value measured during the calibration (between 0 and 4095, included).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximal value measured during the calibration (between 0 and 4095, included)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALIBRATIONMAX_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_calibrationMax() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALIBRATIONMAX_INVALID
        End If
      End If
      Return Me._calibrationMax
    End Function


    '''*
    ''' <summary>
    '''   Changes the maximal calibration value for the input (between 0 and 4095, included), without actually
    '''   starting the automated calibration.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the maximal calibration value for the input (between 0 and 4095,
    '''   included), without actually
    '''   starting the automated calibration
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
    Public Function set_calibrationMax(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("calibrationMax", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the minimal value measured during the calibration (between 0 and 4095, included).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minimal value measured during the calibration (between 0 and 4095, included)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CALIBRATIONMIN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_calibrationMin() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return CALIBRATIONMIN_INVALID
        End If
      End If
      Return Me._calibrationMin
    End Function


    '''*
    ''' <summary>
    '''   Changes the minimal calibration value for the input (between 0 and 4095, included), without actually
    '''   starting the automated calibration.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minimal calibration value for the input (between 0 and 4095,
    '''   included), without actually
    '''   starting the automated calibration
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
    Public Function set_calibrationMin(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("calibrationMin", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the sensibility for the input (between 1 and 1000) for triggering user callbacks.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the sensibility for the input (between 1 and 1000) for triggering user callbacks
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SENSITIVITY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_sensitivity() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return SENSITIVITY_INVALID
        End If
      End If
      Return Me._sensitivity
    End Function


    '''*
    ''' <summary>
    '''   Changes the sensibility for the input (between 1 and 1000) for triggering user callbacks.
    ''' <para>
    '''   The sensibility is used to filter variations around a fixed value, but does not preclude the
    '''   transmission of events when the input value evolves constantly in the same direction.
    '''   Special case: when the value 1000 is used, the callback will only be thrown when the logical state
    '''   of the input switches from pressed to released and back.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the sensibility for the input (between 1 and 1000) for triggering user callbacks
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
    Public Function set_sensitivity(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("sensitivity", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns true if the input (considered as binary) is active (closed contact), and false otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_ISPRESSED_FALSE</c> or <c>Y_ISPRESSED_TRUE</c>, according to true if the input
    '''   (considered as binary) is active (closed contact), and false otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ISPRESSED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_isPressed() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ISPRESSED_INVALID
        End If
      End If
      Return Me._isPressed
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of elapsed milliseconds between the module power on and the last time
    '''   the input button was pressed (the input contact transitioned from open to closed).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of elapsed milliseconds between the module power on and the last time
    '''   the input button was pressed (the input contact transitioned from open to closed)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LASTTIMEPRESSED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastTimePressed() As Long
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LASTTIMEPRESSED_INVALID
        End If
      End If
      Return Me._lastTimePressed
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of elapsed milliseconds between the module power on and the last time
    '''   the input button was released (the input contact transitioned from closed to open).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of elapsed milliseconds between the module power on and the last time
    '''   the input button was released (the input contact transitioned from closed to open)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LASTTIMERELEASED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_lastTimeReleased() As Long
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return LASTTIMERELEASED_INVALID
        End If
      End If
      Return Me._lastTimeReleased
    End Function

    '''*
    ''' <summary>
    '''   Returns the pulse counter value
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PULSECOUNTER_INVALID
        End If
      End If
      Return Me._pulseCounter
    End Function


    Public Function set_pulseCounter(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pulseCounter", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the timer of the pulses counter (ms)
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
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return PULSETIMER_INVALID
        End If
      End If
      Return Me._pulseTimer
    End Function

    '''*
    ''' <summary>
    '''   Retrieves an analog input for a given identifier.
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
    '''   This function does not require that the analog input is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YAnButton.isOnline()</c> to test if the analog input is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an analog input by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the analog input
    ''' </param>
    ''' <returns>
    '''   a <c>YAnButton</c> object allowing you to drive the analog input.
    ''' </returns>
    '''/
    Public Shared Function FindAnButton(func As String) As YAnButton
      Dim obj As YAnButton
      obj = CType(YFunction._FindFromCache("AnButton", func), YAnButton)
      If ((obj Is Nothing)) Then
        obj = New YAnButton(func)
        YFunction._AddToCache("AnButton", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YAnButtonValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackAnButton = callback
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
      If (Not (Me._valueCallbackAnButton Is Nothing)) Then
        Me._valueCallbackAnButton(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the pulse counter value as well as his timer
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
    '''   Continues the enumeration of analog inputs started using <c>yFirstAnButton()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAnButton</c> object, corresponding to
    '''   an analog input currently online, or a <c>null</c> pointer
    '''   if there are no more analog inputs to enumerate.
    ''' </returns>
    '''/
    Public Function nextAnButton() As YAnButton
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YAnButton.FindAnButton(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of analog inputs currently accessible.
    ''' <para>
    '''   Use the method <c>YAnButton.nextAnButton()</c> to iterate on
    '''   next analog inputs.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAnButton</c> object, corresponding to
    '''   the first analog input currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstAnButton() As YAnButton
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("AnButton", 0, p, size, neededsize, errmsg)
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
      Return YAnButton.FindAnButton(serial + "." + funcId)
    End Function

    REM --- (end of YAnButton public methods declaration)

  End Class

  REM --- (AnButton functions)

  '''*
  ''' <summary>
  '''   Retrieves an analog input for a given identifier.
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
  '''   This function does not require that the analog input is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YAnButton.isOnline()</c> to test if the analog input is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an analog input by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the analog input
  ''' </param>
  ''' <returns>
  '''   a <c>YAnButton</c> object allowing you to drive the analog input.
  ''' </returns>
  '''/
  Public Function yFindAnButton(ByVal func As String) As YAnButton
    Return YAnButton.FindAnButton(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of analog inputs currently accessible.
  ''' <para>
  '''   Use the method <c>YAnButton.nextAnButton()</c> to iterate on
  '''   next analog inputs.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YAnButton</c> object, corresponding to
  '''   the first analog input currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstAnButton() As YAnButton
    Return YAnButton.FirstAnButton()
  End Function


  REM --- (end of AnButton functions)

End Module

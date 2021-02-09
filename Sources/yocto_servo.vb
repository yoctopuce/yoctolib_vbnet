' ********************************************************************
'
'  $Id: yocto_servo.vb 43580 2021-01-26 17:46:01Z mvuilleu $
'
'  Implements yFindServo(), the high-level API for Servo functions
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

Module yocto_servo

    REM --- (YServo return codes)
    REM --- (end of YServo return codes)
    REM --- (YServo dlldef)
    REM --- (end of YServo dlldef)
   REM --- (YServo yapiwrapper)
   REM --- (end of YServo yapiwrapper)
  REM --- (YServo globals)

Public Class YServoMove
  Public target As Integer = YAPI.INVALID_INT
  Public ms As Integer = YAPI.INVALID_INT
  Public moving As Integer = YAPI.INVALID_UINT
End Class

  Public Const Y_POSITION_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_ENABLED_FALSE As Integer = 0
  Public Const Y_ENABLED_TRUE As Integer = 1
  Public Const Y_ENABLED_INVALID As Integer = -1
  Public Const Y_RANGE_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_NEUTRAL_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_POSITIONATPOWERON_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_ENABLEDATPOWERON_FALSE As Integer = 0
  Public Const Y_ENABLEDATPOWERON_TRUE As Integer = 1
  Public Const Y_ENABLEDATPOWERON_INVALID As Integer = -1
  Public ReadOnly Y_MOVE_INVALID As YServoMove = Nothing
  Public Delegate Sub YServoValueCallback(ByVal func As YServo, ByVal value As String)
  Public Delegate Sub YServoTimedReportCallback(ByVal func As YServo, ByVal measure As YMeasure)
  REM --- (end of YServo globals)

  REM --- (YServo class start)

  '''*
  ''' <summary>
  '''   The <c>YServo</c> class is designed to drive remote-control servo motors
  '''   outputs.
  ''' <para>
  '''   This class allows you not only to move
  '''   a servo to a given position, but also to specify the time interval
  '''   in which the move should be performed. This makes it possible to
  '''   synchronize two servos involved in a same move.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YServo
    Inherits YFunction
    REM --- (end of YServo class start)

    REM --- (YServo definitions)
    Public Const POSITION_INVALID As Integer = YAPI.INVALID_INT
    Public Const ENABLED_FALSE As Integer = 0
    Public Const ENABLED_TRUE As Integer = 1
    Public Const ENABLED_INVALID As Integer = -1
    Public Const RANGE_INVALID As Integer = YAPI.INVALID_UINT
    Public Const NEUTRAL_INVALID As Integer = YAPI.INVALID_UINT
    Public ReadOnly MOVE_INVALID As YServoMove = Nothing
    Public Const POSITIONATPOWERON_INVALID As Integer = YAPI.INVALID_INT
    Public Const ENABLEDATPOWERON_FALSE As Integer = 0
    Public Const ENABLEDATPOWERON_TRUE As Integer = 1
    Public Const ENABLEDATPOWERON_INVALID As Integer = -1
    REM --- (end of YServo definitions)

    REM --- (YServo attributes declaration)
    Protected _position As Integer
    Protected _enabled As Integer
    Protected _range As Integer
    Protected _neutral As Integer
    Protected _move As YServoMove
    Protected _positionAtPowerOn As Integer
    Protected _enabledAtPowerOn As Integer
    Protected _valueCallbackServo As YServoValueCallback
    REM --- (end of YServo attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Servo"
      REM --- (YServo attributes initialization)
      _position = POSITION_INVALID
      _enabled = ENABLED_INVALID
      _range = RANGE_INVALID
      _neutral = NEUTRAL_INVALID
      _move = New YServoMove()
      _positionAtPowerOn = POSITIONATPOWERON_INVALID
      _enabledAtPowerOn = ENABLEDATPOWERON_INVALID
      _valueCallbackServo = Nothing
      REM --- (end of YServo attributes initialization)
    End Sub

    REM --- (YServo private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("position") Then
        _position = CInt(json_val.getLong("position"))
      End If
      If json_val.has("enabled") Then
        If (json_val.getInt("enabled") > 0) Then _enabled = 1 Else _enabled = 0
      End If
      If json_val.has("range") Then
        _range = CInt(json_val.getLong("range"))
      End If
      If json_val.has("neutral") Then
        _neutral = CInt(json_val.getLong("neutral"))
      End If
      If json_val.has("move") Then
        Dim subjson As YJSONObject = json_val.getYJSONObject("move")
        If (subjson.has("moving")) Then
            _move.moving = subjson.getInt("moving")
        End If
        If (subjson.has("target")) Then
            _move.target = subjson.getInt("target")
        End If
        If (subjson.has("ms")) Then
            _move.ms = subjson.getInt("ms")
        End If
      End If
      If json_val.has("positionAtPowerOn") Then
        _positionAtPowerOn = CInt(json_val.getLong("positionAtPowerOn"))
      End If
      If json_val.has("enabledAtPowerOn") Then
        If (json_val.getInt("enabledAtPowerOn") > 0) Then _enabledAtPowerOn = 1 Else _enabledAtPowerOn = 0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YServo private methods declaration)

    REM --- (YServo public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current servo position.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current servo position
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YServo.POSITION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_position() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return POSITION_INVALID
        End If
      End If
      res = Me._position
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes immediately the servo driving position.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to immediately the servo driving position
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
    Public Function set_position(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("position", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the state of the RC servo motors.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YServo.ENABLED_FALSE</c> or <c>YServo.ENABLED_TRUE</c>, according to the state of the RC servo motors
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YServo.ENABLED_INVALID</c>.
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
    '''   Stops or starts the RC servo motor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YServo.ENABLED_FALSE</c> or <c>YServo.ENABLED_TRUE</c>
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
    '''   Returns the current range of use of the servo.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current range of use of the servo
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YServo.RANGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_range() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RANGE_INVALID
        End If
      End If
      res = Me._range
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the range of use of the servo, specified in per cents.
    ''' <para>
    '''   A range of 100% corresponds to a standard control signal, that varies
    '''   from 1 [ms] to 2 [ms], When using a servo that supports a double range,
    '''   from 0.5 [ms] to 2.5 [ms], you can select a range of 200%.
    '''   Be aware that using a range higher than what is supported by the servo
    '''   is likely to damage the servo. Remember to call the matching module
    '''   <c>saveToFlash()</c> method, otherwise this call will have no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the range of use of the servo, specified in per cents
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
    Public Function set_range(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("range", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the duration in microseconds of a neutral pulse for the servo.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the duration in microseconds of a neutral pulse for the servo
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YServo.NEUTRAL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_neutral() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return NEUTRAL_INVALID
        End If
      End If
      res = Me._neutral
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the duration of the pulse corresponding to the neutral position of the servo.
    ''' <para>
    '''   The duration is specified in microseconds, and the standard value is 1500 [us].
    '''   This setting makes it possible to shift the range of use of the servo.
    '''   Be aware that using a range higher than what is supported by the servo is
    '''   likely to damage the servo. Remember to call the matching module
    '''   <c>saveToFlash()</c> method, otherwise this call will have no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the duration of the pulse corresponding to the neutral position of the servo
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
    Public Function set_neutral(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("neutral", rest_val)
    End Function
    Public Function get_move() As YServoMove
      Dim res As YServoMove
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MOVE_INVALID
        End If
      End If
      res = Me._move
      Return res
    End Function


    Public Function set_move(ByVal newval As YServoMove) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target)) + ":" + Ltrim(Str(newval.ms))
      Return _setAttr("move", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Performs a smooth move at constant speed toward a given position.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="target">
    '''   new position at the end of the move
    ''' </param>
    ''' <param name="ms_duration">
    '''   total duration of the move, in milliseconds
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
    Public Function move(ByVal target As Integer, ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(target)) + ":" + Ltrim(Str(ms_duration))
      Return _setAttr("move", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the servo position at device power up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the servo position at device power up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YServo.POSITIONATPOWERON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_positionAtPowerOn() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return POSITIONATPOWERON_INVALID
        End If
      End If
      res = Me._positionAtPowerOn
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Configure the servo position at device power up.
    ''' <para>
    '''   Remember to call the matching
    '''   module <c>saveToFlash()</c> method, otherwise this call will have no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer
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
    Public Function set_positionAtPowerOn(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("positionAtPowerOn", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the servo signal generator state at power up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YServo.ENABLEDATPOWERON_FALSE</c> or <c>YServo.ENABLEDATPOWERON_TRUE</c>, according to
    '''   the servo signal generator state at power up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YServo.ENABLEDATPOWERON_INVALID</c>.
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
    '''   Configure the servo signal generator state at power up.
    ''' <para>
    '''   Remember to call the matching module <c>saveToFlash()</c>
    '''   method, otherwise this call will have no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YServo.ENABLEDATPOWERON_FALSE</c> or <c>YServo.ENABLEDATPOWERON_TRUE</c>
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
    Public Function set_enabledAtPowerOn(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("enabledAtPowerOn", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a RC servo motor for a given identifier.
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
    '''   This function does not require that the RC servo motor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YServo.isOnline()</c> to test if the RC servo motor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a RC servo motor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the RC servo motor, for instance
    '''   <c>SERVORC1.servo1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YServo</c> object allowing you to drive the RC servo motor.
    ''' </returns>
    '''/
    Public Shared Function FindServo(func As String) As YServo
      Dim obj As YServo
      obj = CType(YFunction._FindFromCache("Servo", func), YServo)
      If ((obj Is Nothing)) Then
        obj = New YServo(func)
        YFunction._AddToCache("Servo", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YServoValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackServo = callback
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
      If (Not (Me._valueCallbackServo Is Nothing)) Then
        Me._valueCallbackServo(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of RC servo motors started using <c>yFirstServo()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned RC servo motors order.
    '''   If you want to find a specific a RC servo motor, use <c>Servo.findServo()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YServo</c> object, corresponding to
    '''   a RC servo motor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more RC servo motors to enumerate.
    ''' </returns>
    '''/
    Public Function nextServo() As YServo
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YServo.FindServo(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of RC servo motors currently accessible.
    ''' <para>
    '''   Use the method <c>YServo.nextServo()</c> to iterate on
    '''   next RC servo motors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YServo</c> object, corresponding to
    '''   the first RC servo motor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstServo() As YServo
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Servo", 0, p, size, neededsize, errmsg)
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
      Return YServo.FindServo(serial + "." + funcId)
    End Function

    REM --- (end of YServo public methods declaration)

  End Class

  REM --- (YServo functions)

  '''*
  ''' <summary>
  '''   Retrieves a RC servo motor for a given identifier.
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
  '''   This function does not require that the RC servo motor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YServo.isOnline()</c> to test if the RC servo motor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a RC servo motor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the RC servo motor, for instance
  '''   <c>SERVORC1.servo1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YServo</c> object allowing you to drive the RC servo motor.
  ''' </returns>
  '''/
  Public Function yFindServo(ByVal func As String) As YServo
    Return YServo.FindServo(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of RC servo motors currently accessible.
  ''' <para>
  '''   Use the method <c>YServo.nextServo()</c> to iterate on
  '''   next RC servo motors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YServo</c> object, corresponding to
  '''   the first RC servo motor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstServo() As YServo
    Return YServo.FirstServo()
  End Function


  REM --- (end of YServo functions)

End Module

' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindMotor(), the high-level API for Motor functions
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

Module yocto_motor

    REM --- (YMotor return codes)
    REM --- (end of YMotor return codes)
    REM --- (YMotor dlldef)
    REM --- (end of YMotor dlldef)
   REM --- (YMotor yapiwrapper)
   REM --- (end of YMotor yapiwrapper)
  REM --- (YMotor globals)

  Public Const Y_MOTORSTATUS_IDLE As Integer = 0
  Public Const Y_MOTORSTATUS_BRAKE As Integer = 1
  Public Const Y_MOTORSTATUS_FORWD As Integer = 2
  Public Const Y_MOTORSTATUS_BACKWD As Integer = 3
  Public Const Y_MOTORSTATUS_LOVOLT As Integer = 4
  Public Const Y_MOTORSTATUS_HICURR As Integer = 5
  Public Const Y_MOTORSTATUS_HIHEAT As Integer = 6
  Public Const Y_MOTORSTATUS_FAILSF As Integer = 7
  Public Const Y_MOTORSTATUS_INVALID As Integer = -1
  Public Const Y_DRIVINGFORCE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_BRAKINGFORCE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_CUTOFFVOLTAGE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_OVERCURRENTLIMIT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_FREQUENCY_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_STARTERTIME_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_FAILSAFETIMEOUT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YMotorValueCallback(ByVal func As YMotor, ByVal value As String)
  Public Delegate Sub YMotorTimedReportCallback(ByVal func As YMotor, ByVal measure As YMeasure)
  REM --- (end of YMotor globals)

  REM --- (YMotor class start)

  '''*
  ''' <summary>
  '''   The <c>YMotor</c> class allows you to drive a DC motor.
  ''' <para>
  '''   It can be used to configure the
  '''   power sent to the motor to make it turn both ways, but also to drive accelerations
  '''   and decelerations. The motor will then accelerate automatically: you will not
  '''   have to monitor it. The API also allows to slow down the motor by shortening
  '''   its terminals: the motor will then act as an electromagnetic brake.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YMotor
    Inherits YFunction
    REM --- (end of YMotor class start)

    REM --- (YMotor definitions)
    Public Const MOTORSTATUS_IDLE As Integer = 0
    Public Const MOTORSTATUS_BRAKE As Integer = 1
    Public Const MOTORSTATUS_FORWD As Integer = 2
    Public Const MOTORSTATUS_BACKWD As Integer = 3
    Public Const MOTORSTATUS_LOVOLT As Integer = 4
    Public Const MOTORSTATUS_HICURR As Integer = 5
    Public Const MOTORSTATUS_HIHEAT As Integer = 6
    Public Const MOTORSTATUS_FAILSF As Integer = 7
    Public Const MOTORSTATUS_INVALID As Integer = -1
    Public Const DRIVINGFORCE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const BRAKINGFORCE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const CUTOFFVOLTAGE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const OVERCURRENTLIMIT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const FREQUENCY_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const STARTERTIME_INVALID As Integer = YAPI.INVALID_UINT
    Public Const FAILSAFETIMEOUT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YMotor definitions)

    REM --- (YMotor attributes declaration)
    Protected _motorStatus As Integer
    Protected _drivingForce As Double
    Protected _brakingForce As Double
    Protected _cutOffVoltage As Double
    Protected _overCurrentLimit As Integer
    Protected _frequency As Double
    Protected _starterTime As Integer
    Protected _failSafeTimeout As Integer
    Protected _command As String
    Protected _valueCallbackMotor As YMotorValueCallback
    REM --- (end of YMotor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Motor"
      REM --- (YMotor attributes initialization)
      _motorStatus = MOTORSTATUS_INVALID
      _drivingForce = DRIVINGFORCE_INVALID
      _brakingForce = BRAKINGFORCE_INVALID
      _cutOffVoltage = CUTOFFVOLTAGE_INVALID
      _overCurrentLimit = OVERCURRENTLIMIT_INVALID
      _frequency = FREQUENCY_INVALID
      _starterTime = STARTERTIME_INVALID
      _failSafeTimeout = FAILSAFETIMEOUT_INVALID
      _command = COMMAND_INVALID
      _valueCallbackMotor = Nothing
      REM --- (end of YMotor attributes initialization)
    End Sub

    REM --- (YMotor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("motorStatus") Then
        _motorStatus = CInt(json_val.getLong("motorStatus"))
      End If
      If json_val.has("drivingForce") Then
        _drivingForce = Math.Round(json_val.getDouble("drivingForce") / 65.536) / 1000.0
      End If
      If json_val.has("brakingForce") Then
        _brakingForce = Math.Round(json_val.getDouble("brakingForce") / 65.536) / 1000.0
      End If
      If json_val.has("cutOffVoltage") Then
        _cutOffVoltage = Math.Round(json_val.getDouble("cutOffVoltage") / 65.536) / 1000.0
      End If
      If json_val.has("overCurrentLimit") Then
        _overCurrentLimit = CInt(json_val.getLong("overCurrentLimit"))
      End If
      If json_val.has("frequency") Then
        _frequency = Math.Round(json_val.getDouble("frequency") / 65.536) / 1000.0
      End If
      If json_val.has("starterTime") Then
        _starterTime = CInt(json_val.getLong("starterTime"))
      End If
      If json_val.has("failSafeTimeout") Then
        _failSafeTimeout = CInt(json_val.getLong("failSafeTimeout"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YMotor private methods declaration)

    REM --- (YMotor public methods declaration)
    '''*
    ''' <summary>
    '''   Return the controller state.
    ''' <para>
    '''   Possible states are:
    '''   IDLE   when the motor is stopped/in free wheel, ready to start;
    '''   FORWD  when the controller is driving the motor forward;
    '''   BACKWD when the controller is driving the motor backward;
    '''   BRAKE  when the controller is braking;
    '''   LOVOLT when the controller has detected a low voltage condition;
    '''   HICURR when the controller has detected an over current condition;
    '''   HIHEAT when the controller has detected an overheat condition;
    '''   FAILSF when the controller switched on the failsafe security.
    ''' </para>
    ''' <para>
    '''   When an error condition occurred (LOVOLT, HICURR, HIHEAT, FAILSF), the controller
    '''   status must be explicitly reset using the <c>resetStatus</c> function.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>YMotor.MOTORSTATUS_IDLE</c>, <c>YMotor.MOTORSTATUS_BRAKE</c>,
    '''   <c>YMotor.MOTORSTATUS_FORWD</c>, <c>YMotor.MOTORSTATUS_BACKWD</c>,
    '''   <c>YMotor.MOTORSTATUS_LOVOLT</c>, <c>YMotor.MOTORSTATUS_HICURR</c>,
    '''   <c>YMotor.MOTORSTATUS_HIHEAT</c> and <c>YMotor.MOTORSTATUS_FAILSF</c>
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMotor.MOTORSTATUS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_motorStatus() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MOTORSTATUS_INVALID
        End If
      End If
      res = Me._motorStatus
      Return res
    End Function


    Public Function set_motorStatus(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("motorStatus", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes immediately the power sent to the motor.
    ''' <para>
    '''   The value is a percentage between -100%
    '''   to 100%. If you want go easy on your mechanics and avoid excessive current consumption,
    '''   try to avoid brutal power changes. For example, immediate transition from forward full power
    '''   to reverse full power is a very bad idea. Each time the driving power is modified, the
    '''   braking power is set to zero.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to immediately the power sent to the motor
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
    Public Function set_drivingForce(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("drivingForce", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the power sent to the motor, as a percentage between -100% and +100%.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the power sent to the motor, as a percentage between -100% and +100%
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMotor.DRIVINGFORCE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_drivingForce() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DRIVINGFORCE_INVALID
        End If
      End If
      res = Me._drivingForce
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes immediately the braking force applied to the motor (in percents).
    ''' <para>
    '''   The value 0 corresponds to no braking (free wheel). When the braking force
    '''   is changed, the driving power is set to zero. The value is a percentage.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to immediately the braking force applied to the motor (in percents)
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
    Public Function set_brakingForce(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("brakingForce", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the braking force applied to the motor, as a percentage.
    ''' <para>
    '''   The value 0 corresponds to no braking (free wheel).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the braking force applied to the motor, as a percentage
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMotor.BRAKINGFORCE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_brakingForce() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return BRAKINGFORCE_INVALID
        End If
      End If
      res = Me._brakingForce
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the threshold voltage under which the controller automatically switches to error state
    '''   and prevents further current draw.
    ''' <para>
    '''   This setting prevent damage to a battery that can
    '''   occur when drawing current from an "empty" battery.
    '''   Note that whatever the cutoff threshold, the controller switches to undervoltage
    '''   error state if the power supply goes under 3V, even for a very brief time.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the threshold voltage under which the controller
    '''   automatically switches to error state
    '''   and prevents further current draw
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
    Public Function set_cutOffVoltage(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("cutOffVoltage", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the threshold voltage under which the controller automatically switches to error state
    '''   and prevents further current draw.
    ''' <para>
    '''   This setting prevents damage to a battery that can
    '''   occur when drawing current from an "empty" battery.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the threshold voltage under which the controller
    '''   automatically switches to error state
    '''   and prevents further current draw
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMotor.CUTOFFVOLTAGE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_cutOffVoltage() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CUTOFFVOLTAGE_INVALID
        End If
      End If
      res = Me._cutOffVoltage
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the current threshold (in mA) above which the controller automatically
    '''   switches to error state.
    ''' <para>
    '''   A zero value means that there is no limit.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current threshold (in mA) above which the controller automatically
    '''   switches to error state
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMotor.OVERCURRENTLIMIT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_overCurrentLimit() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return OVERCURRENTLIMIT_INVALID
        End If
      End If
      res = Me._overCurrentLimit
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current threshold (in mA) above which the controller automatically
    '''   switches to error state.
    ''' <para>
    '''   A zero value means that there is no limit. Note that whatever the
    '''   current limit is, the controller switches to OVERCURRENT status if the current
    '''   goes above 32A, even for a very brief time. Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current threshold (in mA) above which the controller automatically
    '''   switches to error state
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
    Public Function set_overCurrentLimit(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("overCurrentLimit", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Changes the PWM frequency used to control the motor.
    ''' <para>
    '''   Low frequency is usually
    '''   more efficient and may help the motor to start, but an audible noise might be
    '''   generated. A higher frequency reduces the noise, but more energy is converted
    '''   into heat. Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the PWM frequency used to control the motor
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
    Public Function set_frequency(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("frequency", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the PWM frequency used to control the motor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the PWM frequency used to control the motor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMotor.FREQUENCY_INVALID</c>.
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
    '''   Returns the duration (in ms) during which the motor is driven at low frequency to help
    '''   it start up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the duration (in ms) during which the motor is driven at low frequency to help
    '''   it start up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMotor.STARTERTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_starterTime() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return STARTERTIME_INVALID
        End If
      End If
      res = Me._starterTime
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the duration (in ms) during which the motor is driven at low frequency to help
    '''   it start up.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the duration (in ms) during which the motor is driven at low frequency to help
    '''   it start up
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
    Public Function set_starterTime(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("starterTime", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the delay in milliseconds allowed for the controller to run autonomously without
    '''   receiving any instruction from the control process.
    ''' <para>
    '''   When this delay has elapsed,
    '''   the controller automatically stops the motor and switches to FAILSAFE error.
    '''   Failsafe security is disabled when the value is zero.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the delay in milliseconds allowed for the controller to run autonomously without
    '''   receiving any instruction from the control process
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YMotor.FAILSAFETIMEOUT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_failSafeTimeout() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return FAILSAFETIMEOUT_INVALID
        End If
      End If
      res = Me._failSafeTimeout
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the delay in milliseconds allowed for the controller to run autonomously without
    '''   receiving any instruction from the control process.
    ''' <para>
    '''   When this delay has elapsed,
    '''   the controller automatically stops the motor and switches to FAILSAFE error.
    '''   Failsafe security is disabled when the value is zero.
    '''   Remember to call the <c>saveToFlash()</c>
    '''   method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the delay in milliseconds allowed for the controller to run autonomously without
    '''   receiving any instruction from the control process
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
    Public Function set_failSafeTimeout(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("failSafeTimeout", rest_val)
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
    '''   Retrieves a motor for a given identifier.
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
    '''   This function does not require that the motor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YMotor.isOnline()</c> to test if the motor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a motor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the motor, for instance
    '''   <c>MOTORCTL.motor</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YMotor</c> object allowing you to drive the motor.
    ''' </returns>
    '''/
    Public Shared Function FindMotor(func As String) As YMotor
      Dim obj As YMotor
      obj = CType(YFunction._FindFromCache("Motor", func), YMotor)
      If ((obj Is Nothing)) Then
        obj = New YMotor(func)
        YFunction._AddToCache("Motor", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YMotorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackMotor = callback
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
      If (Not (Me._valueCallbackMotor Is Nothing)) Then
        Me._valueCallbackMotor(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Rearms the controller failsafe timer.
    ''' <para>
    '''   When the motor is running and the failsafe feature
    '''   is active, this function should be called periodically to prove that the control process
    '''   is running properly. Otherwise, the motor is automatically stopped after the specified
    '''   timeout. Calling a motor <i>set</i> function implicitly rearms the failsafe timer.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function keepALive() As Integer
      Return Me.set_command("K")
    End Function

    '''*
    ''' <summary>
    '''   Reset the controller state to IDLE.
    ''' <para>
    '''   This function must be invoked explicitly
    '''   after any error condition is signaled.
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function resetStatus() As Integer
      Return Me.set_motorStatus(MOTORSTATUS_IDLE)
    End Function

    '''*
    ''' <summary>
    '''   Changes progressively the power sent to the motor for a specific duration.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="targetPower">
    '''   desired motor power, in percents (between -100% and +100%)
    ''' </param>
    ''' <param name="delay">
    '''   duration (in ms) of the transition
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function drivingForceMove(targetPower As Double, delay As Integer) As Integer
      Return Me.set_command("P" + Convert.ToString(CType(Math.Round(targetPower*10), Integer)) + "," + Convert.ToString(delay))
    End Function

    '''*
    ''' <summary>
    '''   Changes progressively the braking force applied to the motor for a specific duration.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="targetPower">
    '''   desired braking force, in percents
    ''' </param>
    ''' <param name="delay">
    '''   duration (in ms) of the transition
    ''' </param>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function brakingForceMove(targetPower As Double, delay As Integer) As Integer
      Return Me.set_command("B" + Convert.ToString(CType(Math.Round(targetPower*10), Integer)) + "," + Convert.ToString(delay))
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of motors started using <c>yFirstMotor()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned motors order.
    '''   If you want to find a specific a motor, use <c>Motor.findMotor()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMotor</c> object, corresponding to
    '''   a motor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more motors to enumerate.
    ''' </returns>
    '''/
    Public Function nextMotor() As YMotor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YMotor.FindMotor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of motors currently accessible.
    ''' <para>
    '''   Use the method <c>YMotor.nextMotor()</c> to iterate on
    '''   next motors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YMotor</c> object, corresponding to
    '''   the first motor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstMotor() As YMotor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Motor", 0, p, size, neededsize, errmsg)
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
      Return YMotor.FindMotor(serial + "." + funcId)
    End Function

    REM --- (end of YMotor public methods declaration)

  End Class

  REM --- (YMotor functions)

  '''*
  ''' <summary>
  '''   Retrieves a motor for a given identifier.
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
  '''   This function does not require that the motor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YMotor.isOnline()</c> to test if the motor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a motor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the motor, for instance
  '''   <c>MOTORCTL.motor</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YMotor</c> object allowing you to drive the motor.
  ''' </returns>
  '''/
  Public Function yFindMotor(ByVal func As String) As YMotor
    Return YMotor.FindMotor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of motors currently accessible.
  ''' <para>
  '''   Use the method <c>YMotor.nextMotor()</c> to iterate on
  '''   next motors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YMotor</c> object, corresponding to
  '''   the first motor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstMotor() As YMotor
    Return YMotor.FirstMotor()
  End Function


  REM --- (end of YMotor functions)

End Module

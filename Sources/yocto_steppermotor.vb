' ********************************************************************
'
'  $Id: yocto_steppermotor.vb 38913 2019-12-20 18:59:49Z mvuilleu $
'
'  Implements yFindStepperMotor(), the high-level API for StepperMotor functions
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

Module yocto_steppermotor

    REM --- (YStepperMotor return codes)
    REM --- (end of YStepperMotor return codes)
    REM --- (YStepperMotor dlldef)
    REM --- (end of YStepperMotor dlldef)
   REM --- (YStepperMotor yapiwrapper)
   REM --- (end of YStepperMotor yapiwrapper)
  REM --- (YStepperMotor globals)

  Public Const Y_MOTORSTATE_ABSENT As Integer = 0
  Public Const Y_MOTORSTATE_ALERT As Integer = 1
  Public Const Y_MOTORSTATE_HI_Z As Integer = 2
  Public Const Y_MOTORSTATE_STOP As Integer = 3
  Public Const Y_MOTORSTATE_RUN As Integer = 4
  Public Const Y_MOTORSTATE_BATCH As Integer = 5
  Public Const Y_MOTORSTATE_INVALID As Integer = -1
  Public Const Y_DIAGS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_STEPPOS_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_SPEED_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_PULLINSPEED_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_MAXACCEL_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_MAXSPEED_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_STEPPING_MICROSTEP16 As Integer = 0
  Public Const Y_STEPPING_MICROSTEP8 As Integer = 1
  Public Const Y_STEPPING_MICROSTEP4 As Integer = 2
  Public Const Y_STEPPING_HALFSTEP As Integer = 3
  Public Const Y_STEPPING_FULLSTEP As Integer = 4
  Public Const Y_STEPPING_INVALID As Integer = -1
  Public Const Y_OVERCURRENT_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_TCURRSTOP_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_TCURRRUN_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_ALERTMODE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_AUXMODE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_AUXSIGNAL_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YStepperMotorValueCallback(ByVal func As YStepperMotor, ByVal value As String)
  Public Delegate Sub YStepperMotorTimedReportCallback(ByVal func As YStepperMotor, ByVal measure As YMeasure)
  REM --- (end of YStepperMotor globals)

  REM --- (YStepperMotor class start)

  '''*
  ''' <summary>
  '''   The <c>YStepperMotor</c> class allows you to drive a stepper motor.
  ''' <para>
  ''' </para>
  ''' </summary>
  '''/
  Public Class YStepperMotor
    Inherits YFunction
    REM --- (end of YStepperMotor class start)

    REM --- (YStepperMotor definitions)
    Public Const MOTORSTATE_ABSENT As Integer = 0
    Public Const MOTORSTATE_ALERT As Integer = 1
    Public Const MOTORSTATE_HI_Z As Integer = 2
    Public Const MOTORSTATE_STOP As Integer = 3
    Public Const MOTORSTATE_RUN As Integer = 4
    Public Const MOTORSTATE_BATCH As Integer = 5
    Public Const MOTORSTATE_INVALID As Integer = -1
    Public Const DIAGS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const STEPPOS_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const SPEED_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const PULLINSPEED_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const MAXACCEL_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const MAXSPEED_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const STEPPING_MICROSTEP16 As Integer = 0
    Public Const STEPPING_MICROSTEP8 As Integer = 1
    Public Const STEPPING_MICROSTEP4 As Integer = 2
    Public Const STEPPING_HALFSTEP As Integer = 3
    Public Const STEPPING_FULLSTEP As Integer = 4
    Public Const STEPPING_INVALID As Integer = -1
    Public Const OVERCURRENT_INVALID As Integer = YAPI.INVALID_UINT
    Public Const TCURRSTOP_INVALID As Integer = YAPI.INVALID_UINT
    Public Const TCURRRUN_INVALID As Integer = YAPI.INVALID_UINT
    Public Const ALERTMODE_INVALID As String = YAPI.INVALID_STRING
    Public Const AUXMODE_INVALID As String = YAPI.INVALID_STRING
    Public Const AUXSIGNAL_INVALID As Integer = YAPI.INVALID_INT
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YStepperMotor definitions)

    REM --- (YStepperMotor attributes declaration)
    Protected _motorState As Integer
    Protected _diags As Integer
    Protected _stepPos As Double
    Protected _speed As Double
    Protected _pullinSpeed As Double
    Protected _maxAccel As Double
    Protected _maxSpeed As Double
    Protected _stepping As Integer
    Protected _overcurrent As Integer
    Protected _tCurrStop As Integer
    Protected _tCurrRun As Integer
    Protected _alertMode As String
    Protected _auxMode As String
    Protected _auxSignal As Integer
    Protected _command As String
    Protected _valueCallbackStepperMotor As YStepperMotorValueCallback
    REM --- (end of YStepperMotor attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "StepperMotor"
      REM --- (YStepperMotor attributes initialization)
      _motorState = MOTORSTATE_INVALID
      _diags = DIAGS_INVALID
      _stepPos = STEPPOS_INVALID
      _speed = SPEED_INVALID
      _pullinSpeed = PULLINSPEED_INVALID
      _maxAccel = MAXACCEL_INVALID
      _maxSpeed = MAXSPEED_INVALID
      _stepping = STEPPING_INVALID
      _overcurrent = OVERCURRENT_INVALID
      _tCurrStop = TCURRSTOP_INVALID
      _tCurrRun = TCURRRUN_INVALID
      _alertMode = ALERTMODE_INVALID
      _auxMode = AUXMODE_INVALID
      _auxSignal = AUXSIGNAL_INVALID
      _command = COMMAND_INVALID
      _valueCallbackStepperMotor = Nothing
      REM --- (end of YStepperMotor attributes initialization)
    End Sub

    REM --- (YStepperMotor private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("motorState") Then
        _motorState = CInt(json_val.getLong("motorState"))
      End If
      If json_val.has("diags") Then
        _diags = CInt(json_val.getLong("diags"))
      End If
      If json_val.has("stepPos") Then
        _stepPos = json_val.getDouble("stepPos") / 16.0
      End If
      If json_val.has("speed") Then
        _speed = Math.Round(json_val.getDouble("speed") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("pullinSpeed") Then
        _pullinSpeed = Math.Round(json_val.getDouble("pullinSpeed") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("maxAccel") Then
        _maxAccel = Math.Round(json_val.getDouble("maxAccel") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("maxSpeed") Then
        _maxSpeed = Math.Round(json_val.getDouble("maxSpeed") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("stepping") Then
        _stepping = CInt(json_val.getLong("stepping"))
      End If
      If json_val.has("overcurrent") Then
        _overcurrent = CInt(json_val.getLong("overcurrent"))
      End If
      If json_val.has("tCurrStop") Then
        _tCurrStop = CInt(json_val.getLong("tCurrStop"))
      End If
      If json_val.has("tCurrRun") Then
        _tCurrRun = CInt(json_val.getLong("tCurrRun"))
      End If
      If json_val.has("alertMode") Then
        _alertMode = json_val.getString("alertMode")
      End If
      If json_val.has("auxMode") Then
        _auxMode = json_val.getString("auxMode")
      End If
      If json_val.has("auxSignal") Then
        _auxSignal = CInt(json_val.getLong("auxSignal"))
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YStepperMotor private methods declaration)

    REM --- (YStepperMotor public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the motor working state.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_MOTORSTATE_ABSENT</c>, <c>Y_MOTORSTATE_ALERT</c>, <c>Y_MOTORSTATE_HI_Z</c>,
    '''   <c>Y_MOTORSTATE_STOP</c>, <c>Y_MOTORSTATE_RUN</c> and <c>Y_MOTORSTATE_BATCH</c> corresponding to
    '''   the motor working state
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MOTORSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_motorState() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MOTORSTATE_INVALID
        End If
      End If
      res = Me._motorState
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the stepper motor controller diagnostics, as a bitmap.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the stepper motor controller diagnostics, as a bitmap
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DIAGS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_diags() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DIAGS_INVALID
        End If
      End If
      res = Me._diags
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current logical motor position, measured in steps.
    ''' <para>
    '''   This command does not cause any motor move, as its purpose is only to setup
    '''   the origin of the position counter. The fractional part of the position,
    '''   that corresponds to the physical position of the rotor, is not changed.
    '''   To trigger a motor move, use methods <c>moveTo()</c> or <c>moveRel()</c>
    '''   instead.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the current logical motor position, measured in steps
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
    Public Function set_stepPos(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 100.0)/100.0))
      Return _setAttr("stepPos", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current logical motor position, measured in steps.
    ''' <para>
    '''   The value may include a fractional part when micro-stepping is in use.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current logical motor position, measured in steps
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_STEPPOS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_stepPos() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return STEPPOS_INVALID
        End If
      End If
      res = Me._stepPos
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns current motor speed, measured in steps per second.
    ''' <para>
    '''   To change speed, use method <c>changeSpeed()</c>.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to current motor speed, measured in steps per second
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SPEED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_speed() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return SPEED_INVALID
        End If
      End If
      res = Me._speed
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the motor speed immediately reachable from stop state, measured in steps per second.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the motor speed immediately reachable from stop state,
    '''   measured in steps per second
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
    Public Function set_pullinSpeed(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("pullinSpeed", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the motor speed immediately reachable from stop state, measured in steps per second.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the motor speed immediately reachable from stop state,
    '''   measured in steps per second
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_PULLINSPEED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_pullinSpeed() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return PULLINSPEED_INVALID
        End If
      End If
      res = Me._pullinSpeed
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the maximal motor acceleration, measured in steps per second^2.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the maximal motor acceleration, measured in steps per second^2
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
    Public Function set_maxAccel(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("maxAccel", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the maximal motor acceleration, measured in steps per second^2.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the maximal motor acceleration, measured in steps per second^2
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MAXACCEL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_maxAccel() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MAXACCEL_INVALID
        End If
      End If
      res = Me._maxAccel
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the maximal motor speed, measured in steps per second.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number corresponding to the maximal motor speed, measured in steps per second
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
    Public Function set_maxSpeed(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("maxSpeed", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the maximal motor speed, measured in steps per second.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the maximal motor speed, measured in steps per second
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MAXSPEED_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_maxSpeed() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MAXSPEED_INVALID
        End If
      End If
      res = Me._maxSpeed
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the stepping mode used to drive the motor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_STEPPING_MICROSTEP16</c>, <c>Y_STEPPING_MICROSTEP8</c>,
    '''   <c>Y_STEPPING_MICROSTEP4</c>, <c>Y_STEPPING_HALFSTEP</c> and <c>Y_STEPPING_FULLSTEP</c>
    '''   corresponding to the stepping mode used to drive the motor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_STEPPING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_stepping() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return STEPPING_INVALID
        End If
      End If
      res = Me._stepping
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the stepping mode used to drive the motor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_STEPPING_MICROSTEP16</c>, <c>Y_STEPPING_MICROSTEP8</c>,
    '''   <c>Y_STEPPING_MICROSTEP4</c>, <c>Y_STEPPING_HALFSTEP</c> and <c>Y_STEPPING_FULLSTEP</c>
    '''   corresponding to the stepping mode used to drive the motor
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
    Public Function set_stepping(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("stepping", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the overcurrent alert and emergency stop threshold, measured in mA.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the overcurrent alert and emergency stop threshold, measured in mA
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OVERCURRENT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_overcurrent() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return OVERCURRENT_INVALID
        End If
      End If
      res = Me._overcurrent
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the overcurrent alert and emergency stop threshold, measured in mA.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the overcurrent alert and emergency stop threshold, measured in mA
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
    Public Function set_overcurrent(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("overcurrent", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the torque regulation current when the motor is stopped, measured in mA.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the torque regulation current when the motor is stopped, measured in mA
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_TCURRSTOP_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_tCurrStop() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TCURRSTOP_INVALID
        End If
      End If
      res = Me._tCurrStop
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the torque regulation current when the motor is stopped, measured in mA.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the torque regulation current when the motor is stopped, measured in mA
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
    Public Function set_tCurrStop(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("tCurrStop", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the torque regulation current when the motor is running, measured in mA.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the torque regulation current when the motor is running, measured in mA
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_TCURRRUN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_tCurrRun() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TCURRRUN_INVALID
        End If
      End If
      res = Me._tCurrRun
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the torque regulation current when the motor is running, measured in mA.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the torque regulation current when the motor is running, measured in mA
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
    Public Function set_tCurrRun(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("tCurrRun", rest_val)
    End Function
    Public Function get_alertMode() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ALERTMODE_INVALID
        End If
      End If
      res = Me._alertMode
      Return res
    End Function


    Public Function set_alertMode(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("alertMode", rest_val)
    End Function
    Public Function get_auxMode() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return AUXMODE_INVALID
        End If
      End If
      res = Me._auxMode
      Return res
    End Function


    Public Function set_auxMode(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("auxMode", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current value of the signal generated on the auxiliary output.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current value of the signal generated on the auxiliary output
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_AUXSIGNAL_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_auxSignal() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return AUXSIGNAL_INVALID
        End If
      End If
      res = Me._auxSignal
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the value of the signal generated on the auxiliary output.
    ''' <para>
    '''   Acceptable values depend on the auxiliary output signal type configured.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the value of the signal generated on the auxiliary output
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
    Public Function set_auxSignal(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("auxSignal", rest_val)
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
    '''   Retrieves a stepper motor for a given identifier.
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
    '''   This function does not require that the stepper motor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YStepperMotor.isOnline()</c> to test if the stepper motor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a stepper motor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the stepper motor, for instance
    '''   <c>MyDevice.stepperMotor1</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YStepperMotor</c> object allowing you to drive the stepper motor.
    ''' </returns>
    '''/
    Public Shared Function FindStepperMotor(func As String) As YStepperMotor
      Dim obj As YStepperMotor
      obj = CType(YFunction._FindFromCache("StepperMotor", func), YStepperMotor)
      If ((obj Is Nothing)) Then
        obj = New YStepperMotor(func)
        YFunction._AddToCache("StepperMotor", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YStepperMotorValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackStepperMotor = callback
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
      If (Not (Me._valueCallbackStepperMotor Is Nothing)) Then
        Me._valueCallbackStepperMotor(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function sendCommand(command As String) As Integer
      Dim id As String
      Dim url As String
      Dim retBin As Byte()
      Dim res As Integer = 0
      id = Me.get_functionId()
      id = (id).Substring( 12, 1)
      url = "cmd.txt?" +  id + "=" + command
      REM //may throw an exception
      retBin = Me._download(url)
      res = retBin(0)
      If (res = 49) Then
        If Not(res = 48) Then
          me._throw( YAPI.DEVICE_BUSY,  "Motor command pipeline is full, try again later")
          return YAPI.DEVICE_BUSY
        end if
      Else
        If Not(res = 48) Then
          me._throw( YAPI.IO_ERROR,  "Motor command failed permanently")
          return YAPI.IO_ERROR
        end if
      End If
      Return YAPI.SUCCESS
    End Function

    '''*
    ''' <summary>
    '''   Reinitialize the controller and clear all alert flags.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function reset() As Integer
      Return Me.set_command("Z")
    End Function

    '''*
    ''' <summary>
    '''   Starts the motor backward at the specified speed, to search for the motor home position.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="speed">
    '''   desired speed, in steps per second.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function findHomePosition(speed As Double) As Integer
      Return Me.sendCommand("H" + Convert.ToString(CType(Math.Round(1000*speed), Integer)))
    End Function

    '''*
    ''' <summary>
    '''   Starts the motor at a given speed.
    ''' <para>
    '''   The time needed to reach the requested speed
    '''   will depend on the acceleration parameters configured for the motor.
    ''' </para>
    ''' </summary>
    ''' <param name="speed">
    '''   desired speed, in steps per second. The minimal non-zero speed
    '''   is 0.001 pulse per second.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function changeSpeed(speed As Double) As Integer
      Return Me.sendCommand("R" + Convert.ToString(CType(Math.Round(1000*speed), Integer)))
    End Function

    '''*
    ''' <summary>
    '''   Starts the motor to reach a given absolute position.
    ''' <para>
    '''   The time needed to reach the requested
    '''   position will depend on the acceleration and max speed parameters configured for
    '''   the motor.
    ''' </para>
    ''' </summary>
    ''' <param name="absPos">
    '''   absolute position, measured in steps from the origin.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function moveTo(absPos As Double) As Integer
      Return Me.sendCommand("M" + Convert.ToString(CType(Math.Round(16*absPos), Integer)))
    End Function

    '''*
    ''' <summary>
    '''   Starts the motor to reach a given relative position.
    ''' <para>
    '''   The time needed to reach the requested
    '''   position will depend on the acceleration and max speed parameters configured for
    '''   the motor.
    ''' </para>
    ''' </summary>
    ''' <param name="relPos">
    '''   relative position, measured in steps from the current position.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function moveRel(relPos As Double) As Integer
      Return Me.sendCommand("m" + Convert.ToString(CType(Math.Round(16*relPos), Integer)))
    End Function

    '''*
    ''' <summary>
    '''   Starts the motor to reach a given relative position, keeping the speed under the
    '''   specified limit.
    ''' <para>
    '''   The time needed to reach the requested position will depend on
    '''   the acceleration parameters configured for the motor.
    ''' </para>
    ''' </summary>
    ''' <param name="relPos">
    '''   relative position, measured in steps from the current position.
    ''' </param>
    ''' <param name="maxSpeed">
    '''   limit speed, in steps per second.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function moveRelSlow(relPos As Double, maxSpeed As Double) As Integer
      Return Me.sendCommand("m" + Convert.ToString(CType(Math.Round(16*relPos), Integer)) + "@" + Convert.ToString(CType(Math.Round(1000*maxSpeed), Integer)))
    End Function

    '''*
    ''' <summary>
    '''   Keep the motor in the same state for the specified amount of time, before processing next command.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="waitMs">
    '''   wait time, specified in milliseconds.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function pause(waitMs As Integer) As Integer
      Return Me.sendCommand("_" + Convert.ToString(waitMs))
    End Function

    '''*
    ''' <summary>
    '''   Stops the motor with an emergency alert, without taking any additional precaution.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function emergencyStop() As Integer
      Return Me.set_command("!")
    End Function

    '''*
    ''' <summary>
    '''   Move one step in the direction opposite the direction set when the most recent alert was raised.
    ''' <para>
    '''   The move occurs even if the system is still in alert mode (end switch depressed). Caution.
    '''   use this function with great care as it may cause mechanical damages !
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function alertStepOut() As Integer
      Return Me.set_command(".")
    End Function

    '''*
    ''' <summary>
    '''   Move one single step in the selected direction without regards to end switches.
    ''' <para>
    '''   The move occurs even if the system is still in alert mode (end switch depressed). Caution.
    '''   use this function with great care as it may cause mechanical damages !
    ''' </para>
    ''' </summary>
    ''' <param name="dir">
    '''   Value +1 or -1, according to the desired direction of the move
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function alertStepDir(dir As Integer) As Integer
      If Not(dir <> 0) Then
        me._throw( YAPI.INVALID_ARGUMENT,  "direction must be +1 or -1")
        return YAPI.INVALID_ARGUMENT
      end if
      If (dir > 0) Then
        Return Me.set_command(".+")
      End If
      Return Me.set_command(".-")
    End Function

    '''*
    ''' <summary>
    '''   Stops the motor smoothly as soon as possible, without waiting for ongoing move completion.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function abortAndBrake() As Integer
      Return Me.set_command("B")
    End Function

    '''*
    ''' <summary>
    '''   Turn the controller into Hi-Z mode immediately, without waiting for ongoing move completion.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function abortAndHiZ() As Integer
      Return Me.set_command("z")
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of stepper motors started using <c>yFirstStepperMotor()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned stepper motors order.
    '''   If you want to find a specific a stepper motor, use <c>StepperMotor.findStepperMotor()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YStepperMotor</c> object, corresponding to
    '''   a stepper motor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more stepper motors to enumerate.
    ''' </returns>
    '''/
    Public Function nextStepperMotor() As YStepperMotor
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YStepperMotor.FindStepperMotor(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of stepper motors currently accessible.
    ''' <para>
    '''   Use the method <c>YStepperMotor.nextStepperMotor()</c> to iterate on
    '''   next stepper motors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YStepperMotor</c> object, corresponding to
    '''   the first stepper motor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstStepperMotor() As YStepperMotor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("StepperMotor", 0, p, size, neededsize, errmsg)
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
      Return YStepperMotor.FindStepperMotor(serial + "." + funcId)
    End Function

    REM --- (end of YStepperMotor public methods declaration)

  End Class

  REM --- (YStepperMotor functions)

  '''*
  ''' <summary>
  '''   Retrieves a stepper motor for a given identifier.
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
  '''   This function does not require that the stepper motor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YStepperMotor.isOnline()</c> to test if the stepper motor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a stepper motor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the stepper motor, for instance
  '''   <c>MyDevice.stepperMotor1</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YStepperMotor</c> object allowing you to drive the stepper motor.
  ''' </returns>
  '''/
  Public Function yFindStepperMotor(ByVal func As String) As YStepperMotor
    Return YStepperMotor.FindStepperMotor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of stepper motors currently accessible.
  ''' <para>
  '''   Use the method <c>YStepperMotor.nextStepperMotor()</c> to iterate on
  '''   next stepper motors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YStepperMotor</c> object, corresponding to
  '''   the first stepper motor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstStepperMotor() As YStepperMotor
    Return YStepperMotor.FirstStepperMotor()
  End Function


  REM --- (end of YStepperMotor functions)

End Module

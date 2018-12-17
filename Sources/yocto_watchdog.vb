' ********************************************************************
'
'  $Id: yocto_watchdog.vb 33719 2018-12-14 14:22:41Z seb $
'
'  Implements yFindWatchdog(), the high-level API for Watchdog functions
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

Module yocto_watchdog

    REM --- (YWatchdog return codes)
    REM --- (end of YWatchdog return codes)
    REM --- (YWatchdog dlldef)
    REM --- (end of YWatchdog dlldef)
   REM --- (YWatchdog yapiwrapper)
   REM --- (end of YWatchdog yapiwrapper)
  REM --- (YWatchdog globals)

Public Class YWatchdogDelayedPulse
  Public target As Integer = YAPI.INVALID_INT
  Public ms As Integer = YAPI.INVALID_INT
  Public moving As Integer = YAPI.INVALID_UINT
End Class

  REM Y_STATE is defined in yocto_api.vb
  Public Const Y_STATEATPOWERON_UNCHANGED As Integer = 0
  Public Const Y_STATEATPOWERON_A As Integer = 1
  Public Const Y_STATEATPOWERON_B As Integer = 2
  Public Const Y_STATEATPOWERON_INVALID As Integer = -1
  Public Const Y_MAXTIMEONSTATEA_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_MAXTIMEONSTATEB_INVALID As Long = YAPI.INVALID_LONG
  REM Y_OUTPUT is defined in yocto_api.vb
  Public Const Y_PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_COUNTDOWN_INVALID As Long = YAPI.INVALID_LONG
  REM Y_AUTOSTART is defined in yocto_api.vb
  REM Y_RUNNING is defined in yocto_api.vb
  Public Const Y_TRIGGERDELAY_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_TRIGGERDURATION_INVALID As Long = YAPI.INVALID_LONG
  Public ReadOnly Y_DELAYEDPULSETIMER_INVALID As YWatchdogDelayedPulse = Nothing
  Public Delegate Sub YWatchdogValueCallback(ByVal func As YWatchdog, ByVal value As String)
  Public Delegate Sub YWatchdogTimedReportCallback(ByVal func As YWatchdog, ByVal measure As YMeasure)
  REM --- (end of YWatchdog globals)

  REM --- (YWatchdog class start)

  '''*
  ''' <summary>
  '''   The watchdog function works like a relay and can cause a brief power cut
  '''   to an appliance after a preset delay to force this appliance to
  '''   reset.
  ''' <para>
  '''   The Watchdog must be called from time to time to reset the
  '''   timer and prevent the appliance reset.
  '''   The watchdog can be driven directly with <i>pulse</i> and <i>delayedpulse</i> methods to switch
  '''   off an appliance for a given duration.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YWatchdog
    Inherits YFunction
    REM --- (end of YWatchdog class start)

    REM --- (YWatchdog definitions)
    Public Const STATE_A As Integer = 0
    Public Const STATE_B As Integer = 1
    Public Const STATE_INVALID As Integer = -1
    Public Const STATEATPOWERON_UNCHANGED As Integer = 0
    Public Const STATEATPOWERON_A As Integer = 1
    Public Const STATEATPOWERON_B As Integer = 2
    Public Const STATEATPOWERON_INVALID As Integer = -1
    Public Const MAXTIMEONSTATEA_INVALID As Long = YAPI.INVALID_LONG
    Public Const MAXTIMEONSTATEB_INVALID As Long = YAPI.INVALID_LONG
    Public Const OUTPUT_OFF As Integer = 0
    Public Const OUTPUT_ON As Integer = 1
    Public Const OUTPUT_INVALID As Integer = -1
    Public Const PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
    Public ReadOnly DELAYEDPULSETIMER_INVALID As YWatchdogDelayedPulse = Nothing
    Public Const COUNTDOWN_INVALID As Long = YAPI.INVALID_LONG
    Public Const AUTOSTART_OFF As Integer = 0
    Public Const AUTOSTART_ON As Integer = 1
    Public Const AUTOSTART_INVALID As Integer = -1
    Public Const RUNNING_OFF As Integer = 0
    Public Const RUNNING_ON As Integer = 1
    Public Const RUNNING_INVALID As Integer = -1
    Public Const TRIGGERDELAY_INVALID As Long = YAPI.INVALID_LONG
    Public Const TRIGGERDURATION_INVALID As Long = YAPI.INVALID_LONG
    REM --- (end of YWatchdog definitions)

    REM --- (YWatchdog attributes declaration)
    Protected _state As Integer
    Protected _stateAtPowerOn As Integer
    Protected _maxTimeOnStateA As Long
    Protected _maxTimeOnStateB As Long
    Protected _output As Integer
    Protected _pulseTimer As Long
    Protected _delayedPulseTimer As YWatchdogDelayedPulse
    Protected _countdown As Long
    Protected _autoStart As Integer
    Protected _running As Integer
    Protected _triggerDelay As Long
    Protected _triggerDuration As Long
    Protected _valueCallbackWatchdog As YWatchdogValueCallback
    REM --- (end of YWatchdog attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Watchdog"
      REM --- (YWatchdog attributes initialization)
      _state = STATE_INVALID
      _stateAtPowerOn = STATEATPOWERON_INVALID
      _maxTimeOnStateA = MAXTIMEONSTATEA_INVALID
      _maxTimeOnStateB = MAXTIMEONSTATEB_INVALID
      _output = OUTPUT_INVALID
      _pulseTimer = PULSETIMER_INVALID
      _delayedPulseTimer = New YWatchdogDelayedPulse()
      _countdown = COUNTDOWN_INVALID
      _autoStart = AUTOSTART_INVALID
      _running = RUNNING_INVALID
      _triggerDelay = TRIGGERDELAY_INVALID
      _triggerDuration = TRIGGERDURATION_INVALID
      _valueCallbackWatchdog = Nothing
      REM --- (end of YWatchdog attributes initialization)
    End Sub

    REM --- (YWatchdog private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("state") Then
        If (json_val.getInt("state") > 0) Then _state = 1 Else _state = 0
      End If
      If json_val.has("stateAtPowerOn") Then
        _stateAtPowerOn = CInt(json_val.getLong("stateAtPowerOn"))
      End If
      If json_val.has("maxTimeOnStateA") Then
        _maxTimeOnStateA = json_val.getLong("maxTimeOnStateA")
      End If
      If json_val.has("maxTimeOnStateB") Then
        _maxTimeOnStateB = json_val.getLong("maxTimeOnStateB")
      End If
      If json_val.has("output") Then
        If (json_val.getInt("output") > 0) Then _output = 1 Else _output = 0
      End If
      If json_val.has("pulseTimer") Then
        _pulseTimer = json_val.getLong("pulseTimer")
      End If
      If json_val.has("delayedPulseTimer") Then
        Dim subjson As YJSONObject = json_val.getYJSONObject("delayedPulseTimer")
        If (subjson.has("moving")) Then
            _delayedPulseTimer.moving = subjson.getInt("moving")
        End If
        If (subjson.has("target")) Then
            _delayedPulseTimer.target = subjson.getInt("target")
        End If
        If (subjson.has("ms")) Then
            _delayedPulseTimer.ms = subjson.getInt("ms")
        End If
      End If
      If json_val.has("countdown") Then
        _countdown = json_val.getLong("countdown")
      End If
      If json_val.has("autoStart") Then
        If (json_val.getInt("autoStart") > 0) Then _autoStart = 1 Else _autoStart = 0
      End If
      If json_val.has("running") Then
        If (json_val.getInt("running") > 0) Then _running = 1 Else _running = 0
      End If
      If json_val.has("triggerDelay") Then
        _triggerDelay = json_val.getLong("triggerDelay")
      End If
      If json_val.has("triggerDuration") Then
        _triggerDuration = json_val.getLong("triggerDuration")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YWatchdog private methods declaration)

    REM --- (YWatchdog public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the state of the watchdog (A for the idle position, B for the active position).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_STATE_A</c> or <c>Y_STATE_B</c>, according to the state of the watchdog (A for the idle
    '''   position, B for the active position)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_STATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_state() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return STATE_INVALID
        End If
      End If
      res = Me._state
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the state of the watchdog (A for the idle position, B for the active position).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_STATE_A</c> or <c>Y_STATE_B</c>, according to the state of the watchdog (A for the idle
    '''   position, B for the active position)
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
    Public Function set_state(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("state", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the state of the watchdog at device startup (A for the idle position, B for the active position, UNCHANGED for no change).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_STATEATPOWERON_UNCHANGED</c>, <c>Y_STATEATPOWERON_A</c> and
    '''   <c>Y_STATEATPOWERON_B</c> corresponding to the state of the watchdog at device startup (A for the
    '''   idle position, B for the active position, UNCHANGED for no change)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_STATEATPOWERON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_stateAtPowerOn() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return STATEATPOWERON_INVALID
        End If
      End If
      res = Me._stateAtPowerOn
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Preset the state of the watchdog at device startup (A for the idle position,
    '''   B for the active position, UNCHANGED for no modification).
    ''' <para>
    '''   Remember to call the matching module <c>saveToFlash()</c>
    '''   method, otherwise this call will have no effect.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_STATEATPOWERON_UNCHANGED</c>, <c>Y_STATEATPOWERON_A</c> and <c>Y_STATEATPOWERON_B</c>
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
    Public Function set_stateAtPowerOn(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("stateAtPowerOn", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retourne the maximum time (ms) allowed for $THEFUNCTIONS$ to stay in state A before automatically switching back in to B state.
    ''' <para>
    '''   Zero means no maximum time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MAXTIMEONSTATEA_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_maxTimeOnStateA() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MAXTIMEONSTATEA_INVALID
        End If
      End If
      res = Me._maxTimeOnStateA
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Sets the maximum time (ms) allowed for $THEFUNCTIONS$ to stay in state A before automatically switching back in to B state.
    ''' <para>
    '''   Use zero for no maximum time.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_maxTimeOnStateA(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("maxTimeOnStateA", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retourne the maximum time (ms) allowed for $THEFUNCTIONS$ to stay in state B before automatically switching back in to A state.
    ''' <para>
    '''   Zero means no maximum time.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MAXTIMEONSTATEB_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_maxTimeOnStateB() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return MAXTIMEONSTATEB_INVALID
        End If
      End If
      res = Me._maxTimeOnStateB
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Sets the maximum time (ms) allowed for $THEFUNCTIONS$ to stay in state B before automatically switching back in to A state.
    ''' <para>
    '''   Use zero for no maximum time.
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
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_maxTimeOnStateB(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("maxTimeOnStateB", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the output state of the watchdog, when used as a simple switch (single throw).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_OUTPUT_OFF</c> or <c>Y_OUTPUT_ON</c>, according to the output state of the watchdog,
    '''   when used as a simple switch (single throw)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_OUTPUT_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_output() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return OUTPUT_INVALID
        End If
      End If
      res = Me._output
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the output state of the watchdog, when used as a simple switch (single throw).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_OUTPUT_OFF</c> or <c>Y_OUTPUT_ON</c>, according to the output state of the watchdog,
    '''   when used as a simple switch (single throw)
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
    Public Function set_output(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("output", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of milliseconds remaining before the watchdog is returned to idle position
    '''   (state A), during a measured pulse generation.
    ''' <para>
    '''   When there is no ongoing pulse, returns zero.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of milliseconds remaining before the watchdog is returned to
    '''   idle position
    '''   (state A), during a measured pulse generation
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


    Public Function set_pulseTimer(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("pulseTimer", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Sets the relay to output B (active) for a specified duration, then brings it
    '''   automatically back to output A (idle state).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="ms_duration">
    '''   pulse duration, in milliseconds
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
    Public Function pulse(ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(ms_duration))
      Return _setAttr("pulseTimer", rest_val)
    End Function
    Public Function get_delayedPulseTimer() As YWatchdogDelayedPulse
      Dim res As YWatchdogDelayedPulse
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return DELAYEDPULSETIMER_INVALID
        End If
      End If
      res = Me._delayedPulseTimer
      Return res
    End Function


    Public Function set_delayedPulseTimer(ByVal newval As YWatchdogDelayedPulse) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target)) + ":" + Ltrim(Str(newval.ms))
      Return _setAttr("delayedPulseTimer", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Schedules a pulse.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="ms_delay">
    '''   waiting time before the pulse, in milliseconds
    ''' </param>
    ''' <param name="ms_duration">
    '''   pulse duration, in milliseconds
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
    Public Function delayedPulse(ByVal ms_delay As Integer, ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(ms_delay)) + ":" + Ltrim(Str(ms_duration))
      Return _setAttr("delayedPulseTimer", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the number of milliseconds remaining before a pulse (delayedPulse() call)
    '''   When there is no scheduled pulse, returns zero.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of milliseconds remaining before a pulse (delayedPulse() call)
    '''   When there is no scheduled pulse, returns zero
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_COUNTDOWN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_countdown() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COUNTDOWN_INVALID
        End If
      End If
      res = Me._countdown
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the watchdog running state at module power on.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_AUTOSTART_OFF</c> or <c>Y_AUTOSTART_ON</c>, according to the watchdog running state at
    '''   module power on
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_AUTOSTART_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_autoStart() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return AUTOSTART_INVALID
        End If
      End If
      res = Me._autoStart
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the watchdog running state at module power on.
    ''' <para>
    '''   Remember to call the
    '''   <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_AUTOSTART_OFF</c> or <c>Y_AUTOSTART_ON</c>, according to the watchdog running state at
    '''   module power on
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
    '''   Returns the watchdog running state.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_RUNNING_OFF</c> or <c>Y_RUNNING_ON</c>, according to the watchdog running state
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RUNNING_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_running() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RUNNING_INVALID
        End If
      End If
      res = Me._running
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the running state of the watchdog.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_RUNNING_OFF</c> or <c>Y_RUNNING_ON</c>, according to the running state of the watchdog
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
    Public Function set_running(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("running", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Resets the watchdog.
    ''' <para>
    '''   When the watchdog is running, this function
    '''   must be called on a regular basis to prevent the watchdog to
    '''   trigger
    ''' </para>
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
    Public Function resetWatchdog() As Integer
      Dim rest_val As String
      rest_val = "1"
      Return _setAttr("running", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns  the waiting duration before a reset is automatically triggered by the watchdog, in milliseconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to  the waiting duration before a reset is automatically triggered by the
    '''   watchdog, in milliseconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_TRIGGERDELAY_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_triggerDelay() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TRIGGERDELAY_INVALID
        End If
      End If
      res = Me._triggerDelay
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the waiting delay before a reset is triggered by the watchdog, in milliseconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the waiting delay before a reset is triggered by the watchdog, in milliseconds
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
    Public Function set_triggerDelay(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("triggerDelay", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the duration of resets caused by the watchdog, in milliseconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the duration of resets caused by the watchdog, in milliseconds
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_TRIGGERDURATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_triggerDuration() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return TRIGGERDURATION_INVALID
        End If
      End If
      res = Me._triggerDuration
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the duration of resets caused by the watchdog, in milliseconds.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the duration of resets caused by the watchdog, in milliseconds
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
    Public Function set_triggerDuration(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("triggerDuration", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a watchdog for a given identifier.
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
    '''   This function does not require that the watchdog is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YWatchdog.isOnline()</c> to test if the watchdog is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a watchdog by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the watchdog
    ''' </param>
    ''' <returns>
    '''   a <c>YWatchdog</c> object allowing you to drive the watchdog.
    ''' </returns>
    '''/
    Public Shared Function FindWatchdog(func As String) As YWatchdog
      Dim obj As YWatchdog
      obj = CType(YFunction._FindFromCache("Watchdog", func), YWatchdog)
      If ((obj Is Nothing)) Then
        obj = New YWatchdog(func)
        YFunction._AddToCache("Watchdog", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YWatchdogValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackWatchdog = callback
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
      If (Not (Me._valueCallbackWatchdog Is Nothing)) Then
        Me._valueCallbackWatchdog(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of watchdog started using <c>yFirstWatchdog()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned watchdog order.
    '''   If you want to find a specific a watchdog, use <c>Watchdog.findWatchdog()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWatchdog</c> object, corresponding to
    '''   a watchdog currently online, or a <c>Nothing</c> pointer
    '''   if there are no more watchdog to enumerate.
    ''' </returns>
    '''/
    Public Function nextWatchdog() As YWatchdog
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YWatchdog.FindWatchdog(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of watchdog currently accessible.
    ''' <para>
    '''   Use the method <c>YWatchdog.nextWatchdog()</c> to iterate on
    '''   next watchdog.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWatchdog</c> object, corresponding to
    '''   the first watchdog currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstWatchdog() As YWatchdog
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Watchdog", 0, p, size, neededsize, errmsg)
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
      Return YWatchdog.FindWatchdog(serial + "." + funcId)
    End Function

    REM --- (end of YWatchdog public methods declaration)

  End Class

  REM --- (YWatchdog functions)

  '''*
  ''' <summary>
  '''   Retrieves a watchdog for a given identifier.
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
  '''   This function does not require that the watchdog is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YWatchdog.isOnline()</c> to test if the watchdog is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a watchdog by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the watchdog
  ''' </param>
  ''' <returns>
  '''   a <c>YWatchdog</c> object allowing you to drive the watchdog.
  ''' </returns>
  '''/
  Public Function yFindWatchdog(ByVal func As String) As YWatchdog
    Return YWatchdog.FindWatchdog(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of watchdog currently accessible.
  ''' <para>
  '''   Use the method <c>YWatchdog.nextWatchdog()</c> to iterate on
  '''   next watchdog.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YWatchdog</c> object, corresponding to
  '''   the first watchdog currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstWatchdog() As YWatchdog
    Return YWatchdog.FirstWatchdog()
  End Function


  REM --- (end of YWatchdog functions)

End Module

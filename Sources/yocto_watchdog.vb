'*********************************************************************
'*
'* $Id: yocto_watchdog.vb 10401 2013-03-17 13:17:31Z martinm $
'*
'* Implements yFindWatchdog(), the high-level API for Watchdog functions
'*
'* - - - - - - - - - License information: - - - - - - - - - 
'*
'* Copyright (C) 2011 and beyond by Yoctopuce Sarl, Switzerland.
'*
'* 1) If you have obtained this file from www.yoctopuce.com,
'*    Yoctopuce Sarl licenses to you (hereafter Licensee) the
'*    right to use, modify, copy, and integrate this source file
'*    into your own solution for the sole purpose of interfacing
'*    a Yoctopuce product with Licensee's solution.
'*
'*    The use of this file and all relationship between Yoctopuce 
'*    and Licensee are governed by Yoctopuce General Terms and 
'*    Conditions.
'*
'*    THE SOFTWARE AND DOCUMENTATION ARE PROVIDED 'AS IS' WITHOUT
'*    WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING 
'*    WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY, FITNESS 
'*    FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO
'*    EVENT SHALL LICENSOR BE LIABLE FOR ANY INCIDENTAL, SPECIAL,
'*    INDIRECT OR CONSEQUENTIAL DAMAGES, LOST PROFITS OR LOST DATA, 
'*    COST OF PROCUREMENT OF SUBSTITUTE GOODS, TECHNOLOGY OR 
'*    SERVICES, ANY CLAIMS BY THIRD PARTIES (INCLUDING BUT NOT 
'*    LIMITED TO ANY DEFENSE THEREOF), ANY CLAIMS FOR INDEMNITY OR
'*    CONTRIBUTION, OR OTHER SIMILAR COSTS, WHETHER ASSERTED ON THE
'*    BASIS OF CONTRACT, TORT (INCLUDING NEGLIGENCE), BREACH OF
'*    WARRANTY, OR OTHERWISE.
'*
'* 2) If your intent is not to interface with Yoctopuce products,
'*    you are not entitled to use, read or create any derived
'*    material from this source file.
'*
'*********************************************************************/


Imports YDEV_DESCR = System.Int32
Imports YFUN_DESCR = System.Int32
Imports System.Runtime.InteropServices
Imports System.Text

Module yocto_watchdog

  REM --- (YWatchdog definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YWatchdog, ByVal value As String)

Public Class YWatchdogDelayedPulse
  Public target As System.Int64 = YAPI.INVALID_LONG
  Public ms As System.Int64 = YAPI.INVALID_LONG
  Public moving As System.Int64 = YAPI.INVALID_LONG
End Class


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  REM Y_STATE is defined in yocto_api.vb
  REM Y_OUTPUT is defined in yocto_api.vb
  Public Const Y_PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_COUNTDOWN_INVALID As Long = YAPI.INVALID_LONG
  REM Y_AUTOSTART is defined in yocto_api.vb
  REM Y_RUNNING is defined in yocto_api.vb
  Public Const Y_TRIGGERDELAY_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_TRIGGERDURATION_INVALID As Long = YAPI.INVALID_LONG

  Public Y_DELAYEDPULSETIMER_INVALID As YWatchdogDelayedPulse

  REM --- (end of YWatchdog definitions)

  REM --- (YWatchdog implementation)

  Private _WatchdogCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   The watchog function works like a relay and can cause a brief power cut
  '''   to an appliance after a preset delay to force this appliance to
  '''   reset.
  ''' <para>
  '''   The Watchdog must be called from time to time to reset the
  '''   timer and prevent the appliance reset.
  '''   The watchdog can be driven direcly with <i>pulse</i> and <i>delayedpulse</i> methods to switch
  '''   off an appliance for a given duration.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YWatchdog
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const STATE_A = 0
    Public Const STATE_B = 1
    Public Const STATE_INVALID = -1

    Public Const OUTPUT_OFF = 0
    Public Const OUTPUT_ON = 1
    Public Const OUTPUT_INVALID = -1

    Public Const PULSETIMER_INVALID As Long = YAPI.INVALID_LONG
    Public Const COUNTDOWN_INVALID As Long = YAPI.INVALID_LONG
    Public Const AUTOSTART_OFF = 0
    Public Const AUTOSTART_ON = 1
    Public Const AUTOSTART_INVALID = -1

    Public Const RUNNING_OFF = 0
    Public Const RUNNING_ON = 1
    Public Const RUNNING_INVALID = -1

    Public Const TRIGGERDELAY_INVALID As Long = YAPI.INVALID_LONG
    Public Const TRIGGERDURATION_INVALID As Long = YAPI.INVALID_LONG

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _state As Long
    Protected _output As Long
    Protected _pulseTimer As Long
    Protected _delayedPulseTimer As YWatchdogDelayedPulse
    Protected _countdown As Long
    Protected _autoStart As Long
    Protected _running As Long
    Protected _triggerDelay As Long
    Protected _triggerDuration As Long

    Public Sub New(ByVal func As String)
      MyBase.new("Watchdog", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _state = Y_STATE_INVALID
      _output = Y_OUTPUT_INVALID
      _pulseTimer = Y_PULSETIMER_INVALID
      _delayedPulseTimer = New YWatchdogDelayedPulse()
      _countdown = Y_COUNTDOWN_INVALID
      _autoStart = Y_AUTOSTART_INVALID
      _running = Y_RUNNING_INVALID
      _triggerDelay = Y_TRIGGERDELAY_INVALID
      _triggerDuration = Y_TRIGGERDURATION_INVALID
    End Sub

    Protected Overrides Function _parse(ByRef j As TJSONRECORD) As Integer
      Dim member As TJSONRECORD
      Dim i As Integer
      If (j.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then
        Return -1
      End If
      For i = 0 To j.membercount - 1
        member = j.members(i)
        If (member.name = "logicalName") Then
          _logicalName = member.svalue
        ElseIf (member.name = "advertisedValue") Then
          _advertisedValue = member.svalue
        ElseIf (member.name = "state") Then
          If (member.ivalue > 0) Then _state = 1 Else _state = 0
        ElseIf (member.name = "output") Then
          If (member.ivalue > 0) Then _output = 1 Else _output = 0
        ElseIf (member.name = "pulseTimer") Then
          _pulseTimer = CLng(member.ivalue)
        ElseIf (member.name = "delayedPulseTimer") Then
          If (member.recordtype <> TJSONRECORDTYPE.JSON_STRUCT) Then 
             _parse = -1
             Exit Function
          End If
          Dim submemb As TJSONRECORD
          Dim l As Integer
          For l=0 To member.membercount-1
             submemb = member.members(l)
             If (submemb.name = "moving") Then
                _delayedPulseTimer.moving = submemb.ivalue
             ElseIf (submemb.name = "target") Then
                _delayedPulseTimer.target = submemb.ivalue
             ElseIf (submemb.name = "ms") Then
                _delayedPulseTimer.ms = submemb.ivalue
             End If
          Next l
        ElseIf (member.name = "countdown") Then
          _countdown = CLng(member.ivalue)
        ElseIf (member.name = "autoStart") Then
          If (member.ivalue > 0) Then _autoStart = 1 Else _autoStart = 0
        ElseIf (member.name = "running") Then
          If (member.ivalue > 0) Then _running = 1 Else _running = 0
        ElseIf (member.name = "triggerDelay") Then
          _triggerDelay = CLng(member.ivalue)
        ElseIf (member.name = "triggerDuration") Then
          _triggerDuration = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the watchdog.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the watchdog
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_LOGICALNAME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_logicalName() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_LOGICALNAME_INVALID
        End If
      End If
      Return _logicalName
    End Function

    '''*
    ''' <summary>
    '''   Changes the logical name of the watchdog.
    ''' <para>
    '''   You can use <c>yCheckLogicalName()</c>
    '''   prior to this call to make sure that your parameter is valid.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the logical name of the watchdog
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
    Public Function set_logicalName(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("logicalName", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the current value of the watchdog (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the watchdog (no more than 6 characters)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ADVERTISEDVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_advertisedValue() As String
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_ADVERTISEDVALUE_INVALID
        End If
      End If
      Return _advertisedValue
    End Function

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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_STATE_INVALID
        End If
      End If
      Return CType(_state,Integer)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_OUTPUT_INVALID
        End If
      End If
      Return CType(_output,Integer)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_PULSETIMER_INVALID
        End If
      End If
      Return _pulseTimer
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
    '''   pulse duration, in millisecondes
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_DELAYEDPULSETIMER_INVALID
        End If
      End If
      Return _delayedPulseTimer
    End Function

    Public Function set_delayedPulseTimer(ByVal newval As YWatchdogDelayedPulse) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval.target))+":"+Ltrim(Str(newval.ms))
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
    '''   waiting time before the pulse, in millisecondes
    ''' </param>
    ''' <param name="ms_duration">
    '''   pulse duration, in millisecondes
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
    Public Function delayedPulse(ByVal ms_delay As Integer,ByVal ms_duration As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(ms_delay))+":"+Ltrim(Str(ms_duration))
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_COUNTDOWN_INVALID
        End If
      End If
      Return _countdown
    End Function

    '''*
    ''' <summary>
    '''   Returns the watchdog runing state at module power up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_AUTOSTART_OFF</c> or <c>Y_AUTOSTART_ON</c>, according to the watchdog runing state at
    '''   module power up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_AUTOSTART_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_autoStart() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_AUTOSTART_INVALID
        End If
      End If
      Return CType(_autoStart,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the watchdog runningsttae at module power up.
    ''' <para>
    '''   Remember to call the
    '''   <c>saveToFlash()</c> method and then to reboot the module to apply this setting.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>Y_AUTOSTART_OFF</c> or <c>Y_AUTOSTART_ON</c>, according to the watchdog runningsttae at
    '''   module power up
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_RUNNING_INVALID
        End If
      End If
      Return CType(_running,Integer)
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
    '''   must be called on a regular basis to prevent the watchog to
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_TRIGGERDELAY_INVALID
        End If
      End If
      Return _triggerDelay
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_TRIGGERDURATION_INVALID
        End If
      End If
      Return _triggerDuration
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
    '''   Continues the enumeration of watchdog started using <c>yFirstWatchdog()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWatchdog</c> object, corresponding to
    '''   a watchdog currently online, or a <c>null</c> pointer
    '''   if there are no more watchdog to enumerate.
    ''' </returns>
    '''/
    Public Function nextWatchdog() as YWatchdog
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindWatchdog(hwid)
    End Function

    '''*
    ''' <summary>
    '''   comment from .
    ''' <para>
    '''   yc definition
    ''' </para>
    ''' </summary>
    '''/
  Public Overloads Sub registerValueCallback(ByVal callback As UpdateCallback)
   If (callback IsNot Nothing) Then
     registerFuncCallback(Me)
   Else
     unregisterFuncCallback(Me)
   End If
   _callback = callback
  End Sub

  Public Sub set_callback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Sub setCallback(ByVal callback As UpdateCallback)
    registerValueCallback(callback)
  End Sub

  Public Overrides Sub advertiseValue(ByVal value As String)
    If (_callback IsNot Nothing) Then _callback(Me, value)
  End Sub


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
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the watchdog
    ''' </param>
    ''' <returns>
    '''   a <c>YWatchdog</c> object allowing you to drive the watchdog.
    ''' </returns>
    '''/
    Public Shared Function FindWatchdog(ByVal func As String) As YWatchdog
      Dim res As YWatchdog
      If (_WatchdogCache.ContainsKey(func)) Then
        Return CType(_WatchdogCache(func), YWatchdog)
      End If
      res = New YWatchdog(func)
      _WatchdogCache.Add(func, res)
      Return res
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
    '''   the first watchdog currently online, or a <c>null</c> pointer
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

    REM --- (end of YWatchdog implementation)

  End Class

  REM --- (Watchdog functions)

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
  '''   the first watchdog currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstWatchdog() As YWatchdog
    Return YWatchdog.FirstWatchdog()
  End Function

  Private Sub _WatchdogCleanup()
  End Sub


  REM --- (end of Watchdog functions)

End Module

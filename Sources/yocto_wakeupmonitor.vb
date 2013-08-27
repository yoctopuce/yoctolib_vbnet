'*********************************************************************
'*
'* $Id: yocto_wakeupmonitor.vb 12324 2013-08-13 15:10:31Z mvuilleu $
'*
'* Implements yFindWakeUpMonitor(), the high-level API for WakeUpMonitor functions
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

Module yocto_wakeupmonitor

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YWakeUpMonitor definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YWakeUpMonitor, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_POWERDURATION_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_SLEEPCOUNTDOWN_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_NEXTWAKEUP_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_WAKEUPREASON_USBPOWER = 0
  Public Const Y_WAKEUPREASON_EXTPOWER = 1
  Public Const Y_WAKEUPREASON_ENDOFSLEEP = 2
  Public Const Y_WAKEUPREASON_EXTSIG1 = 3
  Public Const Y_WAKEUPREASON_EXTSIG2 = 4
  Public Const Y_WAKEUPREASON_EXTSIG3 = 5
  Public Const Y_WAKEUPREASON_EXTSIG4 = 6
  Public Const Y_WAKEUPREASON_SCHEDULE1 = 7
  Public Const Y_WAKEUPREASON_SCHEDULE2 = 8
  Public Const Y_WAKEUPREASON_SCHEDULE3 = 9
  Public Const Y_WAKEUPREASON_SCHEDULE4 = 10
  Public Const Y_WAKEUPREASON_SCHEDULE5 = 11
  Public Const Y_WAKEUPREASON_SCHEDULE6 = 12
  Public Const Y_WAKEUPREASON_INVALID = -1

  Public Const Y_WAKEUPSTATE_SLEEPING = 0
  Public Const Y_WAKEUPSTATE_AWAKE = 1
  Public Const Y_WAKEUPSTATE_INVALID = -1

  Public Const Y_RTCTIME_INVALID As Long = YAPI.INVALID_LONG


  REM --- (end of YWakeUpMonitor definitions)

  REM --- (YWakeUpMonitor implementation)

  Private _WakeUpMonitorCache As New Hashtable()
  Private _callback As UpdateCallback

  Public Class YWakeUpMonitor
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const POWERDURATION_INVALID As Integer = YAPI.INVALID_INT
    Public Const SLEEPCOUNTDOWN_INVALID As Integer = YAPI.INVALID_INT
    Public Const NEXTWAKEUP_INVALID As Long = YAPI.INVALID_LONG
    Public Const WAKEUPREASON_USBPOWER = 0
    Public Const WAKEUPREASON_EXTPOWER = 1
    Public Const WAKEUPREASON_ENDOFSLEEP = 2
    Public Const WAKEUPREASON_EXTSIG1 = 3
    Public Const WAKEUPREASON_EXTSIG2 = 4
    Public Const WAKEUPREASON_EXTSIG3 = 5
    Public Const WAKEUPREASON_EXTSIG4 = 6
    Public Const WAKEUPREASON_SCHEDULE1 = 7
    Public Const WAKEUPREASON_SCHEDULE2 = 8
    Public Const WAKEUPREASON_SCHEDULE3 = 9
    Public Const WAKEUPREASON_SCHEDULE4 = 10
    Public Const WAKEUPREASON_SCHEDULE5 = 11
    Public Const WAKEUPREASON_SCHEDULE6 = 12
    Public Const WAKEUPREASON_INVALID = -1

    Public Const WAKEUPSTATE_SLEEPING = 0
    Public Const WAKEUPSTATE_AWAKE = 1
    Public Const WAKEUPSTATE_INVALID = -1

    Public Const RTCTIME_INVALID As Long = YAPI.INVALID_LONG

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _powerDuration As Long
    Protected _sleepCountdown As Long
    Protected _nextWakeUp As Long
    Protected _wakeUpReason As Long
    Protected _wakeUpState As Long
    Protected _rtcTime As Long
    Protected _endOfTime As Long

    Public Sub New(ByVal func As String)
      MyBase.new("WakeUpMonitor", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _powerDuration = Y_POWERDURATION_INVALID
      _sleepCountdown = Y_SLEEPCOUNTDOWN_INVALID
      _nextWakeUp = Y_NEXTWAKEUP_INVALID
      _wakeUpReason = Y_WAKEUPREASON_INVALID
      _wakeUpState = Y_WAKEUPSTATE_INVALID
      _rtcTime = Y_RTCTIME_INVALID
      _endOfTime = 2145960000
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
        ElseIf (member.name = "powerDuration") Then
          _powerDuration = member.ivalue
        ElseIf (member.name = "sleepCountdown") Then
          _sleepCountdown = member.ivalue
        ElseIf (member.name = "nextWakeUp") Then
          _nextWakeUp = CLng(member.ivalue)
        ElseIf (member.name = "wakeUpReason") Then
          _wakeUpReason = CLng(member.ivalue)
        ElseIf (member.name = "wakeUpState") Then
          _wakeUpState = CLng(member.ivalue)
        ElseIf (member.name = "rtcTime") Then
          _rtcTime = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the monitor.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the monitor
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
    '''   Changes the logical name of the monitor.
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
    '''   a string corresponding to the logical name of the monitor
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
    '''   Returns the current value of the monitor (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the monitor (no more than 6 characters)
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
    '''   Returns the maximal wake up time (seconds) before going to sleep automatically.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the maximal wake up time (seconds) before going to sleep automatically
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_POWERDURATION_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_powerDuration() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_POWERDURATION_INVALID
        End If
      End If
      Return CType(_powerDuration,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the maximal wake up time (seconds) before going to sleep automatically.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the maximal wake up time (seconds) before going to sleep automatically
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
    Public Function set_powerDuration(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("powerDuration", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the delay before next sleep.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the delay before next sleep
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_SLEEPCOUNTDOWN_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_sleepCountdown() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_SLEEPCOUNTDOWN_INVALID
        End If
      End If
      Return CType(_sleepCountdown,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the delay before next sleep.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the delay before next sleep
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
    Public Function set_sleepCountdown(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("sleepCountdown", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the next scheduled wake-up date/time (UNIX format)
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the next scheduled wake-up date/time (UNIX format)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_NEXTWAKEUP_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nextWakeUp() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_NEXTWAKEUP_INVALID
        End If
      End If
      Return _nextWakeUp
    End Function

    '''*
    ''' <summary>
    '''   Changes the days of the week where a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the days of the week where a wake up must take place
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
    Public Function set_nextWakeUp(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("nextWakeUp", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Return the last wake up reason.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_WAKEUPREASON_USBPOWER</c>, <c>Y_WAKEUPREASON_EXTPOWER</c>,
    '''   <c>Y_WAKEUPREASON_ENDOFSLEEP</c>, <c>Y_WAKEUPREASON_EXTSIG1</c>, <c>Y_WAKEUPREASON_EXTSIG2</c>,
    '''   <c>Y_WAKEUPREASON_EXTSIG3</c>, <c>Y_WAKEUPREASON_EXTSIG4</c>, <c>Y_WAKEUPREASON_SCHEDULE1</c>,
    '''   <c>Y_WAKEUPREASON_SCHEDULE2</c>, <c>Y_WAKEUPREASON_SCHEDULE3</c>, <c>Y_WAKEUPREASON_SCHEDULE4</c>,
    '''   <c>Y_WAKEUPREASON_SCHEDULE5</c> and <c>Y_WAKEUPREASON_SCHEDULE6</c>
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_WAKEUPREASON_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_wakeUpReason() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_WAKEUPREASON_INVALID
        End If
      End If
      Return CType(_wakeUpReason,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Returns  the current state of the monitor
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_WAKEUPSTATE_SLEEPING</c> or <c>Y_WAKEUPSTATE_AWAKE</c>, according to  the current state
    '''   of the monitor
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_WAKEUPSTATE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_wakeUpState() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_WAKEUPSTATE_INVALID
        End If
      End If
      Return CType(_wakeUpState,Integer)
    End Function

    Public Function set_wakeUpState(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("wakeUpState", rest_val)
    End Function

    Public Function get_rtcTime() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_RTCTIME_INVALID
        End If
      End If
      Return _rtcTime
    End Function
    '''*
    ''' <summary>
    '''   Forces a wakeup.
    ''' <para>
    ''' </para>
    ''' </summary>
    '''/
    public function wakeUp() as integer
        REM //fixme use real enum value instead of hardcoded int
        Return Me.set_wakeUpState(WAKEUPSTATE_AWAKE)
        
     end function

    '''*
    ''' <summary>
    '''   Go to sleep until the next wakeup condition is met,  the
    '''   RTC time must have been set before calling this function.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="secBeforeSleep">
    '''   number of seconds before going into sleep mode,
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function sleep(secBeforeSleep as integer) as integer
        dim  currTime as integer
        currTime = CInt(Me.get_rtcTime())
        if not(currTime <> 0) then
me._throw( YAPI.RTC_NOT_READY, "RTC time not set")
 return  YAPI.RTC_NOT_READY
end if

        Me.set_nextWakeUp(Me._endOfTime)
        Me.set_sleepCountdown(secBeforeSleep)
        Return YAPI.SUCCESS
     end function

    '''*
    ''' <summary>
    '''   Go to sleep for a specific time or until the next wakeup condition is met, the
    '''   RTC time must have been set before calling this function.
    ''' <para>
    '''   The count down before sleep
    '''   can be canceled with resetSleepCountDown.
    ''' </para>
    ''' </summary>
    ''' <param name="secUntilWakeUp">
    '''   sleep duration, in secondes
    ''' </param>
    ''' <param name="secBeforeSleep">
    '''   number of seconds before going into sleep mode
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function sleepFor(secUntilWakeUp as integer, secBeforeSleep as integer) as integer
        dim  currTime as integer
        currTime = CInt(Me.get_rtcTime())
        if not(currTime <> 0) then
me._throw( YAPI.RTC_NOT_READY, "RTC time not set")
 return  YAPI.RTC_NOT_READY
end if

        Me.set_nextWakeUp(currTime+secUntilWakeUp)
        Me.set_sleepCountdown(secBeforeSleep)
        Return YAPI.SUCCESS
     end function

    '''*
    ''' <summary>
    '''   Go to sleep until a specific date is reached or until the next wakeup condition is met, the
    '''   RTC time must have been set before calling this function.
    ''' <para>
    '''   The count down before sleep
    '''   can be canceled with resetSleepCountDown.
    ''' </para>
    ''' </summary>
    ''' <param name="wakeUpTime">
    '''   wake-up datetime (UNIX format)
    ''' </param>
    ''' <param name="secBeforeSleep">
    '''   number of seconds before going into sleep mode
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function sleepUntil(wakeUpTime as integer, secBeforeSleep as integer) as integer
        dim  currTime as integer
        currTime = CInt(Me.get_rtcTime())
        if not(currTime <> 0) then
me._throw( YAPI.RTC_NOT_READY, "RTC time not set")
 return  YAPI.RTC_NOT_READY
end if

        Me.set_nextWakeUp(wakeUpTime)
        Me.set_sleepCountdown(secBeforeSleep)
        Return YAPI.SUCCESS
     end function

    '''*
    ''' <summary>
    '''   Reset the sleep countdown.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    public function resetSleepCountDown() as integer
        Me.set_sleepCountdown(0)
        Me.set_nextWakeUp(0)
        Return YAPI.SUCCESS
     end function


    '''*
    ''' <summary>
    '''   Continues the enumeration of monitors started using <c>yFirstWakeUpMonitor()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWakeUpMonitor</c> object, corresponding to
    '''   a monitor currently online, or a <c>null</c> pointer
    '''   if there are no more monitors to enumerate.
    ''' </returns>
    '''/
    Public Function nextWakeUpMonitor() as YWakeUpMonitor
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindWakeUpMonitor(hwid)
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
    '''   Retrieves a monitor for a given identifier.
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
    '''   This function does not require that the monitor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YWakeUpMonitor.isOnline()</c> to test if the monitor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a monitor by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the monitor
    ''' </param>
    ''' <returns>
    '''   a <c>YWakeUpMonitor</c> object allowing you to drive the monitor.
    ''' </returns>
    '''/
    Public Shared Function FindWakeUpMonitor(ByVal func As String) As YWakeUpMonitor
      Dim res As YWakeUpMonitor
      If (_WakeUpMonitorCache.ContainsKey(func)) Then
        Return CType(_WakeUpMonitorCache(func), YWakeUpMonitor)
      End If
      res = New YWakeUpMonitor(func)
      _WakeUpMonitorCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of monitors currently accessible.
    ''' <para>
    '''   Use the method <c>YWakeUpMonitor.nextWakeUpMonitor()</c> to iterate on
    '''   next monitors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWakeUpMonitor</c> object, corresponding to
    '''   the first monitor currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstWakeUpMonitor() As YWakeUpMonitor
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("WakeUpMonitor", 0, p, size, neededsize, errmsg)
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
      Return YWakeUpMonitor.FindWakeUpMonitor(serial + "." + funcId)
    End Function

    REM --- (end of YWakeUpMonitor implementation)

  End Class

  REM --- (WakeUpMonitor functions)

  '''*
  ''' <summary>
  '''   Retrieves a monitor for a given identifier.
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
  '''   This function does not require that the monitor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YWakeUpMonitor.isOnline()</c> to test if the monitor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a monitor by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the monitor
  ''' </param>
  ''' <returns>
  '''   a <c>YWakeUpMonitor</c> object allowing you to drive the monitor.
  ''' </returns>
  '''/
  Public Function yFindWakeUpMonitor(ByVal func As String) As YWakeUpMonitor
    Return YWakeUpMonitor.FindWakeUpMonitor(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of monitors currently accessible.
  ''' <para>
  '''   Use the method <c>YWakeUpMonitor.nextWakeUpMonitor()</c> to iterate on
  '''   next monitors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YWakeUpMonitor</c> object, corresponding to
  '''   the first monitor currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstWakeUpMonitor() As YWakeUpMonitor
    Return YWakeUpMonitor.FirstWakeUpMonitor()
  End Function

  Private Sub _WakeUpMonitorCleanup()
  End Sub


  REM --- (end of WakeUpMonitor functions)

End Module

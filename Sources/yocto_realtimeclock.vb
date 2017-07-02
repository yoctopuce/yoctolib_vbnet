'*********************************************************************
'*
'* $Id: yocto_realtimeclock.vb 27699 2017-06-01 12:26:47Z seb $
'*
'* Implements yFindRealTimeClock(), the high-level API for RealTimeClock functions
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

Module yocto_realtimeclock

    REM --- (YRealTimeClock return codes)
    REM --- (end of YRealTimeClock return codes)
    REM --- (YRealTimeClock dlldef)
    REM --- (end of YRealTimeClock dlldef)
  REM --- (YRealTimeClock globals)

  Public Const Y_UNIXTIME_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_DATETIME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_UTCOFFSET_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_TIMESET_FALSE As Integer = 0
  Public Const Y_TIMESET_TRUE As Integer = 1
  Public Const Y_TIMESET_INVALID As Integer = -1
  Public Delegate Sub YRealTimeClockValueCallback(ByVal func As YRealTimeClock, ByVal value As String)
  Public Delegate Sub YRealTimeClockTimedReportCallback(ByVal func As YRealTimeClock, ByVal measure As YMeasure)
  REM --- (end of YRealTimeClock globals)

  REM --- (YRealTimeClock class start)

  '''*
  ''' <summary>
  '''   The RealTimeClock function maintains and provides current date and time, even accross power cut
  '''   lasting several days.
  ''' <para>
  '''   It is the base for automated wake-up functions provided by the WakeUpScheduler.
  '''   The current time may represent a local time as well as an UTC time, but no automatic time change
  '''   will occur to account for daylight saving time.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YRealTimeClock
    Inherits YFunction
    REM --- (end of YRealTimeClock class start)

    REM --- (YRealTimeClock definitions)
    Public Const UNIXTIME_INVALID As Long = YAPI.INVALID_LONG
    Public Const DATETIME_INVALID As String = YAPI.INVALID_STRING
    Public Const UTCOFFSET_INVALID As Integer = YAPI.INVALID_INT
    Public Const TIMESET_FALSE As Integer = 0
    Public Const TIMESET_TRUE As Integer = 1
    Public Const TIMESET_INVALID As Integer = -1
    REM --- (end of YRealTimeClock definitions)

    REM --- (YRealTimeClock attributes declaration)
    Protected _unixTime As Long
    Protected _dateTime As String
    Protected _utcOffset As Integer
    Protected _timeSet As Integer
    Protected _valueCallbackRealTimeClock As YRealTimeClockValueCallback
    REM --- (end of YRealTimeClock attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "RealTimeClock"
      REM --- (YRealTimeClock attributes initialization)
      _unixTime = UNIXTIME_INVALID
      _dateTime = DATETIME_INVALID
      _utcOffset = UTCOFFSET_INVALID
      _timeSet = TIMESET_INVALID
      _valueCallbackRealTimeClock = Nothing
      REM --- (end of YRealTimeClock attributes initialization)
    End Sub

    REM --- (YRealTimeClock private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("unixTime") Then
        _unixTime = json_val.getLong("unixTime")
      End If
      If json_val.has("dateTime") Then
        _dateTime = json_val.getString("dateTime")
      End If
      If json_val.has("utcOffset") Then
        _utcOffset = CInt(json_val.getLong("utcOffset"))
      End If
      If json_val.has("timeSet") Then
        If (json_val.getInt("timeSet") > 0) Then _timeSet = 1 Else _timeSet = 0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YRealTimeClock private methods declaration)

    REM --- (YRealTimeClock public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the current time in Unix format (number of elapsed seconds since Jan 1st, 1970).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the current time in Unix format (number of elapsed seconds since Jan 1st, 1970)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_UNIXTIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_unixTime() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return UNIXTIME_INVALID
        End If
      End If
      res = Me._unixTime
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the current time.
    ''' <para>
    '''   Time is specifid in Unix format (number of elapsed seconds since Jan 1st, 1970).
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the current time
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
    Public Function set_unixTime(ByVal newval As Long) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("unixTime", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current time in the form "YYYY/MM/DD hh:mm:ss".
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current time in the form "YYYY/MM/DD hh:mm:ss"
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_DATETIME_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_dateTime() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return DATETIME_INVALID
        End If
      End If
      res = Me._dateTime
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Returns the number of seconds between current time and UTC time (time zone).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the number of seconds between current time and UTC time (time zone)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_UTCOFFSET_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_utcOffset() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return UTCOFFSET_INVALID
        End If
      End If
      res = Me._utcOffset
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the number of seconds between current time and UTC time (time zone).
    ''' <para>
    '''   The timezone is automatically rounded to the nearest multiple of 15 minutes.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the number of seconds between current time and UTC time (time zone)
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
    Public Function set_utcOffset(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("utcOffset", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns true if the clock has been set, and false otherwise.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>Y_TIMESET_FALSE</c> or <c>Y_TIMESET_TRUE</c>, according to true if the clock has been
    '''   set, and false otherwise
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_TIMESET_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_timeSet() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return TIMESET_INVALID
        End If
      End If
      res = Me._timeSet
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a clock for a given identifier.
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
    '''   This function does not require that the clock is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YRealTimeClock.isOnline()</c> to test if the clock is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a clock by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the clock
    ''' </param>
    ''' <returns>
    '''   a <c>YRealTimeClock</c> object allowing you to drive the clock.
    ''' </returns>
    '''/
    Public Shared Function FindRealTimeClock(func As String) As YRealTimeClock
      Dim obj As YRealTimeClock
      obj = CType(YFunction._FindFromCache("RealTimeClock", func), YRealTimeClock)
      If ((obj Is Nothing)) Then
        obj = New YRealTimeClock(func)
        YFunction._AddToCache("RealTimeClock", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YRealTimeClockValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackRealTimeClock = callback
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
      If (Not (Me._valueCallbackRealTimeClock Is Nothing)) Then
        Me._valueCallbackRealTimeClock(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of clocks started using <c>yFirstRealTimeClock()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRealTimeClock</c> object, corresponding to
    '''   a clock currently online, or a <c>Nothing</c> pointer
    '''   if there are no more clocks to enumerate.
    ''' </returns>
    '''/
    Public Function nextRealTimeClock() As YRealTimeClock
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YRealTimeClock.FindRealTimeClock(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of clocks currently accessible.
    ''' <para>
    '''   Use the method <c>YRealTimeClock.nextRealTimeClock()</c> to iterate on
    '''   next clocks.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRealTimeClock</c> object, corresponding to
    '''   the first clock currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstRealTimeClock() As YRealTimeClock
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("RealTimeClock", 0, p, size, neededsize, errmsg)
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
      Return YRealTimeClock.FindRealTimeClock(serial + "." + funcId)
    End Function

    REM --- (end of YRealTimeClock public methods declaration)

  End Class

  REM --- (RealTimeClock functions)

  '''*
  ''' <summary>
  '''   Retrieves a clock for a given identifier.
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
  '''   This function does not require that the clock is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YRealTimeClock.isOnline()</c> to test if the clock is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a clock by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the clock
  ''' </param>
  ''' <returns>
  '''   a <c>YRealTimeClock</c> object allowing you to drive the clock.
  ''' </returns>
  '''/
  Public Function yFindRealTimeClock(ByVal func As String) As YRealTimeClock
    Return YRealTimeClock.FindRealTimeClock(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of clocks currently accessible.
  ''' <para>
  '''   Use the method <c>YRealTimeClock.nextRealTimeClock()</c> to iterate on
  '''   next clocks.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YRealTimeClock</c> object, corresponding to
  '''   the first clock currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstRealTimeClock() As YRealTimeClock
    Return YRealTimeClock.FirstRealTimeClock()
  End Function


  REM --- (end of RealTimeClock functions)

End Module

'*********************************************************************
'*
'* $Id: yocto_realtimeclock.vb 12324 2013-08-13 15:10:31Z mvuilleu $
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

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YRealTimeClock definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YRealTimeClock, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_UNIXTIME_INVALID As Long = YAPI.INVALID_LONG
  Public Const Y_DATETIME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_UTCOFFSET_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_TIMESET_FALSE = 0
  Public Const Y_TIMESET_TRUE = 1
  Public Const Y_TIMESET_INVALID = -1



  REM --- (end of YRealTimeClock definitions)

  REM --- (YRealTimeClock implementation)

  Private _RealTimeClockCache As New Hashtable()
  Private _callback As UpdateCallback

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
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const UNIXTIME_INVALID As Long = YAPI.INVALID_LONG
    Public Const DATETIME_INVALID As String = YAPI.INVALID_STRING
    Public Const UTCOFFSET_INVALID As Integer = YAPI.INVALID_INT
    Public Const TIMESET_FALSE = 0
    Public Const TIMESET_TRUE = 1
    Public Const TIMESET_INVALID = -1


    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _unixTime As Long
    Protected _dateTime As String
    Protected _utcOffset As Long
    Protected _timeSet As Long

    Public Sub New(ByVal func As String)
      MyBase.new("RealTimeClock", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _unixTime = Y_UNIXTIME_INVALID
      _dateTime = Y_DATETIME_INVALID
      _utcOffset = Y_UTCOFFSET_INVALID
      _timeSet = Y_TIMESET_INVALID
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
        ElseIf (member.name = "unixTime") Then
          _unixTime = CLng(member.ivalue)
        ElseIf (member.name = "dateTime") Then
          _dateTime = member.svalue
        ElseIf (member.name = "utcOffset") Then
          _utcOffset = member.ivalue
        ElseIf (member.name = "timeSet") Then
          If (member.ivalue > 0) Then _timeSet = 1 Else _timeSet = 0
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the clock.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the clock
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
    '''   Changes the logical name of the clock.
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
    '''   a string corresponding to the logical name of the clock
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
    '''   Returns the current value of the clock (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the clock (no more than 6 characters)
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_UNIXTIME_INVALID
        End If
      End If
      Return _unixTime
    End Function

    '''*
    ''' <summary>
    '''   Changes the current time.
    ''' <para>
    '''   Time is specifid in Unix format (number of elapsed seconds since Jan 1st, 1970).
    '''   If current UTC time is known, utcOffset will be automatically adjusted for the new specified time.
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
    '''   Returns the current time in the form "YYYY/MM/DD hh:mm:ss"
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_DATETIME_INVALID
        End If
      End If
      Return _dateTime
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_UTCOFFSET_INVALID
        End If
      End If
      Return CType(_utcOffset,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the number of seconds between current time and UTC time (time zone).
    ''' <para>
    '''   The timezone is automatically rounded to the nearest multiple of 15 minutes.
    '''   If current UTC time is known, the current time will automatically be updated according to the
    '''   selected time zone.
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
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_TIMESET_INVALID
        End If
      End If
      Return CType(_timeSet,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Continues the enumeration of clocks started using <c>yFirstRealTimeClock()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRealTimeClock</c> object, corresponding to
    '''   a clock currently online, or a <c>null</c> pointer
    '''   if there are no more clocks to enumerate.
    ''' </returns>
    '''/
    Public Function nextRealTimeClock() as YRealTimeClock
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindRealTimeClock(hwid)
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
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the clock
    ''' </param>
    ''' <returns>
    '''   a <c>YRealTimeClock</c> object allowing you to drive the clock.
    ''' </returns>
    '''/
    Public Shared Function FindRealTimeClock(ByVal func As String) As YRealTimeClock
      Dim res As YRealTimeClock
      If (_RealTimeClockCache.ContainsKey(func)) Then
        Return CType(_RealTimeClockCache(func), YRealTimeClock)
      End If
      res = New YRealTimeClock(func)
      _RealTimeClockCache.Add(func, res)
      Return res
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
    '''   the first clock currently online, or a <c>null</c> pointer
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

    REM --- (end of YRealTimeClock implementation)

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
  '''   the first clock currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstRealTimeClock() As YRealTimeClock
    Return YRealTimeClock.FirstRealTimeClock()
  End Function

  Private Sub _RealTimeClockCleanup()
  End Sub


  REM --- (end of RealTimeClock functions)

End Module

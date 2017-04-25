'*********************************************************************
'*
'* $Id: yocto_wakeupschedule.vb 27237 2017-04-21 16:36:03Z seb $
'*
'* Implements yFindWakeUpSchedule(), the high-level API for WakeUpSchedule functions
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

Module yocto_wakeupschedule

    REM --- (YWakeUpSchedule return codes)
    REM --- (end of YWakeUpSchedule return codes)
    REM --- (YWakeUpSchedule dlldef)
    REM --- (end of YWakeUpSchedule dlldef)
  REM --- (YWakeUpSchedule globals)

  Public Const Y_MINUTESA_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_MINUTESB_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_HOURS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_WEEKDAYS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_MONTHDAYS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_MONTHS_INVALID As Integer = YAPI.INVALID_UINT
  Public Const Y_NEXTOCCURENCE_INVALID As Long = YAPI.INVALID_LONG
  Public Delegate Sub YWakeUpScheduleValueCallback(ByVal func As YWakeUpSchedule, ByVal value As String)
  Public Delegate Sub YWakeUpScheduleTimedReportCallback(ByVal func As YWakeUpSchedule, ByVal measure As YMeasure)
  REM --- (end of YWakeUpSchedule globals)

  REM --- (YWakeUpSchedule class start)

  '''*
  ''' <summary>
  '''   The WakeUpSchedule function implements a wake up condition.
  ''' <para>
  '''   The wake up time is
  '''   specified as a set of months and/or days and/or hours and/or minutes when the
  '''   wake up should happen.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YWakeUpSchedule
    Inherits YFunction
    REM --- (end of YWakeUpSchedule class start)

    REM --- (YWakeUpSchedule definitions)
    Public Const MINUTESA_INVALID As Integer = YAPI.INVALID_UINT
    Public Const MINUTESB_INVALID As Integer = YAPI.INVALID_UINT
    Public Const HOURS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const WEEKDAYS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const MONTHDAYS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const MONTHS_INVALID As Integer = YAPI.INVALID_UINT
    Public Const NEXTOCCURENCE_INVALID As Long = YAPI.INVALID_LONG
    REM --- (end of YWakeUpSchedule definitions)

    REM --- (YWakeUpSchedule attributes declaration)
    Protected _minutesA As Integer
    Protected _minutesB As Integer
    Protected _hours As Integer
    Protected _weekDays As Integer
    Protected _monthDays As Integer
    Protected _months As Integer
    Protected _nextOccurence As Long
    Protected _valueCallbackWakeUpSchedule As YWakeUpScheduleValueCallback
    REM --- (end of YWakeUpSchedule attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "WakeUpSchedule"
      REM --- (YWakeUpSchedule attributes initialization)
      _minutesA = MINUTESA_INVALID
      _minutesB = MINUTESB_INVALID
      _hours = HOURS_INVALID
      _weekDays = WEEKDAYS_INVALID
      _monthDays = MONTHDAYS_INVALID
      _months = MONTHS_INVALID
      _nextOccurence = NEXTOCCURENCE_INVALID
      _valueCallbackWakeUpSchedule = Nothing
      REM --- (end of YWakeUpSchedule attributes initialization)
    End Sub

    REM --- (YWakeUpSchedule private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("minutesA") Then
        _minutesA = CInt(json_val.getLong("minutesA"))
      End If
      If json_val.has("minutesB") Then
        _minutesB = CInt(json_val.getLong("minutesB"))
      End If
      If json_val.has("hours") Then
        _hours = CInt(json_val.getLong("hours"))
      End If
      If json_val.has("weekDays") Then
        _weekDays = CInt(json_val.getLong("weekDays"))
      End If
      If json_val.has("monthDays") Then
        _monthDays = CInt(json_val.getLong("monthDays"))
      End If
      If json_val.has("months") Then
        _months = CInt(json_val.getLong("months"))
      End If
      If json_val.has("nextOccurence") Then
        _nextOccurence = json_val.getLong("nextOccurence")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YWakeUpSchedule private methods declaration)

    REM --- (YWakeUpSchedule public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the minutes in the 00-29 interval of each hour scheduled for wake up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minutes in the 00-29 interval of each hour scheduled for wake up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MINUTESA_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_minutesA() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MINUTESA_INVALID
        End If
      End If
      res = Me._minutesA
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the minutes in the 00-29 interval when a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minutes in the 00-29 interval when a wake up must take place
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
    Public Function set_minutesA(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("minutesA", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the minutes in the 30-59 intervalof each hour scheduled for wake up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minutes in the 30-59 intervalof each hour scheduled for wake up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MINUTESB_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_minutesB() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MINUTESB_INVALID
        End If
      End If
      res = Me._minutesB
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the minutes in the 30-59 interval when a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minutes in the 30-59 interval when a wake up must take place
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
    Public Function set_minutesB(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("minutesB", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the hours scheduled for wake up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the hours scheduled for wake up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_HOURS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_hours() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return HOURS_INVALID
        End If
      End If
      res = Me._hours
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the hours when a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the hours when a wake up must take place
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
    Public Function set_hours(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("hours", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the days of the week scheduled for wake up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the days of the week scheduled for wake up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_WEEKDAYS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_weekDays() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return WEEKDAYS_INVALID
        End If
      End If
      res = Me._weekDays
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the days of the week when a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the days of the week when a wake up must take place
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
    Public Function set_weekDays(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("weekDays", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the days of the month scheduled for wake up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the days of the month scheduled for wake up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MONTHDAYS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_monthDays() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MONTHDAYS_INVALID
        End If
      End If
      res = Me._monthDays
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the days of the month when a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the days of the month when a wake up must take place
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
    Public Function set_monthDays(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("monthDays", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the months scheduled for wake up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the months scheduled for wake up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MONTHS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_months() As Integer
      Dim res As Integer = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return MONTHS_INVALID
        End If
      End If
      res = Me._months
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the months when a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the months when a wake up must take place
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
    Public Function set_months(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("months", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the date/time (seconds) of the next wake up occurence.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the date/time (seconds) of the next wake up occurence
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_NEXTOCCURENCE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nextOccurence() As Long
      Dim res As Long = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return NEXTOCCURENCE_INVALID
        End If
      End If
      res = Me._nextOccurence
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves a wake up schedule for a given identifier.
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
    '''   This function does not require that the wake up schedule is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YWakeUpSchedule.isOnline()</c> to test if the wake up schedule is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a wake up schedule by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the wake up schedule
    ''' </param>
    ''' <returns>
    '''   a <c>YWakeUpSchedule</c> object allowing you to drive the wake up schedule.
    ''' </returns>
    '''/
    Public Shared Function FindWakeUpSchedule(func As String) As YWakeUpSchedule
      Dim obj As YWakeUpSchedule
      obj = CType(YFunction._FindFromCache("WakeUpSchedule", func), YWakeUpSchedule)
      If ((obj Is Nothing)) Then
        obj = New YWakeUpSchedule(func)
        YFunction._AddToCache("WakeUpSchedule", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YWakeUpScheduleValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackWakeUpSchedule = callback
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
      If (Not (Me._valueCallbackWakeUpSchedule Is Nothing)) Then
        Me._valueCallbackWakeUpSchedule(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns all the minutes of each hour that are scheduled for wake up.
    ''' <para>
    ''' </para>
    ''' </summary>
    '''/
    Public Overridable Function get_minutes() As Long
      Dim res As Long = 0
      
      res = Me.get_minutesB()
      res = ((res) << (30))
      res = res + Me.get_minutesA()
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Changes all the minutes where a wake up must take place.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitmap">
    '''   Minutes 00-59 of each hour scheduled for wake up.
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_minutes(bitmap As Long) As Integer
      Me.set_minutesA(CInt(((bitmap) And (&H3fffffff))))
      bitmap = ((bitmap) >> (30))
      Return Me.set_minutesB(CInt(((bitmap) And (&H3fffffff))))
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of wake up schedules started using <c>yFirstWakeUpSchedule()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWakeUpSchedule</c> object, corresponding to
    '''   a wake up schedule currently online, or a <c>Nothing</c> pointer
    '''   if there are no more wake up schedules to enumerate.
    ''' </returns>
    '''/
    Public Function nextWakeUpSchedule() As YWakeUpSchedule
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YWakeUpSchedule.FindWakeUpSchedule(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of wake up schedules currently accessible.
    ''' <para>
    '''   Use the method <c>YWakeUpSchedule.nextWakeUpSchedule()</c> to iterate on
    '''   next wake up schedules.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWakeUpSchedule</c> object, corresponding to
    '''   the first wake up schedule currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstWakeUpSchedule() As YWakeUpSchedule
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("WakeUpSchedule", 0, p, size, neededsize, errmsg)
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
      Return YWakeUpSchedule.FindWakeUpSchedule(serial + "." + funcId)
    End Function

    REM --- (end of YWakeUpSchedule public methods declaration)

  End Class

  REM --- (WakeUpSchedule functions)

  '''*
  ''' <summary>
  '''   Retrieves a wake up schedule for a given identifier.
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
  '''   This function does not require that the wake up schedule is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YWakeUpSchedule.isOnline()</c> to test if the wake up schedule is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a wake up schedule by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the wake up schedule
  ''' </param>
  ''' <returns>
  '''   a <c>YWakeUpSchedule</c> object allowing you to drive the wake up schedule.
  ''' </returns>
  '''/
  Public Function yFindWakeUpSchedule(ByVal func As String) As YWakeUpSchedule
    Return YWakeUpSchedule.FindWakeUpSchedule(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of wake up schedules currently accessible.
  ''' <para>
  '''   Use the method <c>YWakeUpSchedule.nextWakeUpSchedule()</c> to iterate on
  '''   next wake up schedules.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YWakeUpSchedule</c> object, corresponding to
  '''   the first wake up schedule currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstWakeUpSchedule() As YWakeUpSchedule
    Return YWakeUpSchedule.FirstWakeUpSchedule()
  End Function


  REM --- (end of WakeUpSchedule functions)

End Module

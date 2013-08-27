'*********************************************************************
'*
'* $Id: yocto_wakeupschedule.vb 12469 2013-08-22 10:11:58Z seb $
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

  REM --- (return codes)
  REM --- (end of return codes)
  
  REM --- (YWakeUpSchedule definitions)

  Public Delegate Sub UpdateCallback(ByVal func As YWakeUpSchedule, ByVal value As String)


  Public Const Y_LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_MINUTESA_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_MINUTESB_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_HOURS_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_WEEKDAYS_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_MONTHDAYS_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_MONTHS_INVALID As Integer = YAPI.INVALID_UNSIGNED
  Public Const Y_NEXTOCCURENCE_INVALID As Long = YAPI.INVALID_LONG


  REM --- (end of YWakeUpSchedule definitions)

  REM --- (YWakeUpSchedule implementation)

  Private _WakeUpScheduleCache As New Hashtable()
  Private _callback As UpdateCallback

  '''*
  ''' <summary>
  '''   The WakeUpSchedule function implements a wake-up condition.
  ''' <para>
  '''   The wake-up time is
  '''   specified as a set of months and/or days and/or hours and/or minutes where the
  '''   wake-up should happen.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YWakeUpSchedule
    Inherits YFunction
    Public Const LOGICALNAME_INVALID As String = YAPI.INVALID_STRING
    Public Const ADVERTISEDVALUE_INVALID As String = YAPI.INVALID_STRING
    Public Const MINUTESA_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const MINUTESB_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const HOURS_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const WEEKDAYS_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const MONTHDAYS_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const MONTHS_INVALID As Integer = YAPI.INVALID_UNSIGNED
    Public Const NEXTOCCURENCE_INVALID As Long = YAPI.INVALID_LONG

    Protected _logicalName As String
    Protected _advertisedValue As String
    Protected _minutesA As Long
    Protected _minutesB As Long
    Protected _hours As Long
    Protected _weekDays As Long
    Protected _monthDays As Long
    Protected _months As Long
    Protected _nextOccurence As Long

    Public Sub New(ByVal func As String)
      MyBase.new("WakeUpSchedule", func)
      _logicalName = Y_LOGICALNAME_INVALID
      _advertisedValue = Y_ADVERTISEDVALUE_INVALID
      _minutesA = Y_MINUTESA_INVALID
      _minutesB = Y_MINUTESB_INVALID
      _hours = Y_HOURS_INVALID
      _weekDays = Y_WEEKDAYS_INVALID
      _monthDays = Y_MONTHDAYS_INVALID
      _months = Y_MONTHS_INVALID
      _nextOccurence = Y_NEXTOCCURENCE_INVALID
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
        ElseIf (member.name = "minutesA") Then
          _minutesA = CLng(member.ivalue)
        ElseIf (member.name = "minutesB") Then
          _minutesB = CLng(member.ivalue)
        ElseIf (member.name = "hours") Then
          _hours = CLng(member.ivalue)
        ElseIf (member.name = "weekDays") Then
          _weekDays = CLng(member.ivalue)
        ElseIf (member.name = "monthDays") Then
          _monthDays = CLng(member.ivalue)
        ElseIf (member.name = "months") Then
          _months = CLng(member.ivalue)
        ElseIf (member.name = "nextOccurence") Then
          _nextOccurence = CLng(member.ivalue)
        End If
      Next i
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Returns the logical name of the wake-up schedule.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the logical name of the wake-up schedule
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
    '''   Changes the logical name of the wake-up schedule.
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
    '''   a string corresponding to the logical name of the wake-up schedule
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
    '''   Returns the current value of the wake-up schedule (no more than 6 characters).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a string corresponding to the current value of the wake-up schedule (no more than 6 characters)
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
    '''   Returns the minutes 00-29 of each hour scheduled for wake-up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minutes 00-29 of each hour scheduled for wake-up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MINUTESA_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_minutesA() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_MINUTESA_INVALID
        End If
      End If
      Return CType(_minutesA,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the minutes 00-29 where a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minutes 00-29 where a wake up must take place
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
    '''   Returns the minutes 30-59 of each hour scheduled for wake-up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the minutes 30-59 of each hour scheduled for wake-up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MINUTESB_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_minutesB() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_MINUTESB_INVALID
        End If
      End If
      Return CType(_minutesB,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the minutes 30-59 where a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the minutes 30-59 where a wake up must take place
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
    '''   Returns the hours  scheduled for wake-up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the hours  scheduled for wake-up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_HOURS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_hours() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_HOURS_INVALID
        End If
      End If
      Return CType(_hours,Integer)
    End Function

    '''*
    ''' <summary>
    '''   Changes the hours where a wake up must take place.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the hours where a wake up must take place
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
    '''   Returns the days of week scheduled for wake-up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the days of week scheduled for wake-up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_WEEKDAYS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_weekDays() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_WEEKDAYS_INVALID
        End If
      End If
      Return CType(_weekDays,Integer)
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
    Public Function set_weekDays(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("weekDays", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the days of week scheduled for wake-up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the days of week scheduled for wake-up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MONTHDAYS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_monthDays() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_MONTHDAYS_INVALID
        End If
      End If
      Return CType(_monthDays,Integer)
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
    Public Function set_monthDays(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("monthDays", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the days of week scheduled for wake-up.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the days of week scheduled for wake-up
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_MONTHS_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_months() As Integer
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_MONTHS_INVALID
        End If
      End If
      Return CType(_months,Integer)
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
    Public Function set_months(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("months", rest_val)
    End Function

    '''*
    ''' <summary>
    '''   Returns the  nextwake up date/time (seconds) wake up occurence
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the  nextwake up date/time (seconds) wake up occurence
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_NEXTOCCURENCE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_nextOccurence() As Long
      If (_cacheExpiration <= YAPI.GetTickCount()) Then
        If (YISERR(load(YAPI.DefaultCacheValidity))) Then
          Return Y_NEXTOCCURENCE_INVALID
        End If
      End If
      Return _nextOccurence
    End Function
    '''*
    ''' <summary>
    '''   Returns every the minutes of each hour scheduled for wake-up.
    ''' <para>
    ''' </para>
    ''' </summary>
    '''/
    public function get_minutes() as long
        dim  res as long
        res = Me.get_minutesB()
        res = res << 30
        res = res + Me.get_minutesA()
        Return res
     end function

    '''*
    ''' <summary>
    '''   Changes all the minutes where a wake up must take place.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="bitmap">
    '''   Minutes 00-59 of each hour scheduled for wake-up.,
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    public function set_minutes(bitmap as long) as integer
        Me.set_minutesA(CInt(bitmap & &H3fffffff))
        bitmap = bitmap >> 30
        Return Me.set_minutesB(CInt(bitmap & &H3fffffff))
        
     end function


    '''*
    ''' <summary>
    '''   Continues the enumeration of wake-up schedules started using <c>yFirstWakeUpSchedule()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWakeUpSchedule</c> object, corresponding to
    '''   a wake-up schedule currently online, or a <c>null</c> pointer
    '''   if there are no more wake-up schedules to enumerate.
    ''' </returns>
    '''/
    Public Function nextWakeUpSchedule() as YWakeUpSchedule
      Dim hwid As String =""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid="") Then
        Return Nothing
      End If
      Return yFindWakeUpSchedule(hwid)
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
    '''   Retrieves a wake-up schedule for a given identifier.
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
    '''   This function does not require that the wake-up schedule is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YWakeUpSchedule.isOnline()</c> to test if the wake-up schedule is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a wake-up schedule by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the wake-up schedule
    ''' </param>
    ''' <returns>
    '''   a <c>YWakeUpSchedule</c> object allowing you to drive the wake-up schedule.
    ''' </returns>
    '''/
    Public Shared Function FindWakeUpSchedule(ByVal func As String) As YWakeUpSchedule
      Dim res As YWakeUpSchedule
      If (_WakeUpScheduleCache.ContainsKey(func)) Then
        Return CType(_WakeUpScheduleCache(func), YWakeUpSchedule)
      End If
      res = New YWakeUpSchedule(func)
      _WakeUpScheduleCache.Add(func, res)
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of wake-up schedules currently accessible.
    ''' <para>
    '''   Use the method <c>YWakeUpSchedule.nextWakeUpSchedule()</c> to iterate on
    '''   next wake-up schedules.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YWakeUpSchedule</c> object, corresponding to
    '''   the first wake-up schedule currently online, or a <c>null</c> pointer
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

    REM --- (end of YWakeUpSchedule implementation)

  End Class

  REM --- (WakeUpSchedule functions)

  '''*
  ''' <summary>
  '''   Retrieves a wake-up schedule for a given identifier.
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
  '''   This function does not require that the wake-up schedule is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YWakeUpSchedule.isOnline()</c> to test if the wake-up schedule is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a wake-up schedule by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the wake-up schedule
  ''' </param>
  ''' <returns>
  '''   a <c>YWakeUpSchedule</c> object allowing you to drive the wake-up schedule.
  ''' </returns>
  '''/
  Public Function yFindWakeUpSchedule(ByVal func As String) As YWakeUpSchedule
    Return YWakeUpSchedule.FindWakeUpSchedule(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of wake-up schedules currently accessible.
  ''' <para>
  '''   Use the method <c>YWakeUpSchedule.nextWakeUpSchedule()</c> to iterate on
  '''   next wake-up schedules.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YWakeUpSchedule</c> object, corresponding to
  '''   the first wake-up schedule currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstWakeUpSchedule() As YWakeUpSchedule
    Return YWakeUpSchedule.FirstWakeUpSchedule()
  End Function

  Private Sub _WakeUpScheduleCleanup()
  End Sub


  REM --- (end of WakeUpSchedule functions)

End Module

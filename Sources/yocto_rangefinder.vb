'*********************************************************************
'*
'* $Id: yocto_rangefinder.vb 26329 2017-01-11 14:04:39Z mvuilleu $
'*
'* Implements yFindRangeFinder(), the high-level API for RangeFinder functions
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

Module yocto_rangefinder

    REM --- (YRangeFinder return codes)
    REM --- (end of YRangeFinder return codes)
    REM --- (YRangeFinder dlldef)
    REM --- (end of YRangeFinder dlldef)
  REM --- (YRangeFinder globals)

  Public Const Y_RANGEFINDERMODE_DEFAULT As Integer = 0
  Public Const Y_RANGEFINDERMODE_LONG_RANGE As Integer = 1
  Public Const Y_RANGEFINDERMODE_HIGH_ACCURACY As Integer = 2
  Public Const Y_RANGEFINDERMODE_HIGH_SPEED As Integer = 3
  Public Const Y_RANGEFINDERMODE_INVALID As Integer = -1
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YRangeFinderValueCallback(ByVal func As YRangeFinder, ByVal value As String)
  Public Delegate Sub YRangeFinderTimedReportCallback(ByVal func As YRangeFinder, ByVal measure As YMeasure)
  REM --- (end of YRangeFinder globals)

  REM --- (YRangeFinder class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce class YRangeFinder allows you to use and configure Yoctopuce range finders
  '''   sensors.
  ''' <para>
  '''   It inherits from YSensor class the core functions to read measurements,
  '''   register callback functions, access to the autonomous datalogger.
  '''   This class adds the ability to easily perform a one-point linear calibration
  '''   to compensate the effect of a glass or filter placed in front of the sensor.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YRangeFinder
    Inherits YSensor
    REM --- (end of YRangeFinder class start)

    REM --- (YRangeFinder definitions)
    Public Const RANGEFINDERMODE_DEFAULT As Integer = 0
    Public Const RANGEFINDERMODE_LONG_RANGE As Integer = 1
    Public Const RANGEFINDERMODE_HIGH_ACCURACY As Integer = 2
    Public Const RANGEFINDERMODE_HIGH_SPEED As Integer = 3
    Public Const RANGEFINDERMODE_INVALID As Integer = -1
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YRangeFinder definitions)

    REM --- (YRangeFinder attributes declaration)
    Protected _rangeFinderMode As Integer
    Protected _command As String
    Protected _valueCallbackRangeFinder As YRangeFinderValueCallback
    Protected _timedReportCallbackRangeFinder As YRangeFinderTimedReportCallback
    REM --- (end of YRangeFinder attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "RangeFinder"
      REM --- (YRangeFinder attributes initialization)
      _rangeFinderMode = RANGEFINDERMODE_INVALID
      _command = COMMAND_INVALID
      _valueCallbackRangeFinder = Nothing
      _timedReportCallbackRangeFinder = Nothing
      REM --- (end of YRangeFinder attributes initialization)
    End Sub

    REM --- (YRangeFinder private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "rangeFinderMode") Then
        _rangeFinderMode = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "command") Then
        _command = member.svalue
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YRangeFinder private methods declaration)

    REM --- (YRangeFinder public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the measuring unit for the measured temperature.
    ''' <para>
    '''   That unit is a string.
    '''   String value can be <c>"</c> or <c>mm</c>. Any other value will be ignored.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    '''   WARNING: if a specific calibration is defined for the rangeFinder function, a
    '''   unit system change will probably break it.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the measuring unit for the measured temperature
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
    Public Function set_unit(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("unit", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the rangefinder running mode.
    ''' <para>
    '''   The rangefinder running mode
    '''   allows to put priority on precision, speed or maximum range.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_RANGEFINDERMODE_DEFAULT</c>, <c>Y_RANGEFINDERMODE_LONG_RANGE</c>,
    '''   <c>Y_RANGEFINDERMODE_HIGH_ACCURACY</c> and <c>Y_RANGEFINDERMODE_HIGH_SPEED</c> corresponding to the
    '''   rangefinder running mode
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RANGEFINDERMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rangeFinderMode() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return RANGEFINDERMODE_INVALID
        End If
      End If
      Return Me._rangeFinderMode
    End Function


    '''*
    ''' <summary>
    '''   Changes the rangefinder running mode, allowing to put priority on
    '''   precision, speed or maximum range.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_RANGEFINDERMODE_DEFAULT</c>, <c>Y_RANGEFINDERMODE_LONG_RANGE</c>,
    '''   <c>Y_RANGEFINDERMODE_HIGH_ACCURACY</c> and <c>Y_RANGEFINDERMODE_HIGH_SPEED</c> corresponding to the
    '''   rangefinder running mode, allowing to put priority on
    '''   precision, speed or maximum range
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
    Public Function set_rangeFinderMode(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("rangeFinderMode", rest_val)
    End Function
    Public Function get_command() As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return COMMAND_INVALID
        End If
      End If
      Return Me._command
    End Function


    Public Function set_command(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("command", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves a range finder for a given identifier.
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
    '''   This function does not require that the range finder is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YRangeFinder.isOnline()</c> to test if the range finder is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   a range finder by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the range finder
    ''' </param>
    ''' <returns>
    '''   a <c>YRangeFinder</c> object allowing you to drive the range finder.
    ''' </returns>
    '''/
    Public Shared Function FindRangeFinder(func As String) As YRangeFinder
      Dim obj As YRangeFinder
      obj = CType(YFunction._FindFromCache("RangeFinder", func), YRangeFinder)
      If ((obj Is Nothing)) Then
        obj = New YRangeFinder(func)
        YFunction._AddToCache("RangeFinder", func, obj)
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
    Public Overloads Function registerValueCallback(callback As YRangeFinderValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackRangeFinder = callback
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
      If (Not (Me._valueCallbackRangeFinder Is Nothing)) Then
        Me._valueCallbackRangeFinder(Me, value)
      Else
        MyBase._invokeValueCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every periodic timed notification.
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
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YRangeFinderTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackRangeFinder = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackRangeFinder Is Nothing)) Then
        Me._timedReportCallbackRangeFinder(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Triggers a sensor calibration according to the current ambient temperature.
    ''' <para>
    '''   That
    '''   calibration process needs no physical interaction with the sensor. It is performed
    '''   automatically at device startup, but it is recommended to start it again when the
    '''   temperature delta since last calibration exceeds 8°C.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function triggerTempCalibration() As Integer
      Return Me.set_command("T")
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of range finders started using <c>yFirstRangeFinder()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRangeFinder</c> object, corresponding to
    '''   a range finder currently online, or a <c>Nothing</c> pointer
    '''   if there are no more range finders to enumerate.
    ''' </returns>
    '''/
    Public Function nextRangeFinder() As YRangeFinder
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YRangeFinder.FindRangeFinder(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of range finders currently accessible.
    ''' <para>
    '''   Use the method <c>YRangeFinder.nextRangeFinder()</c> to iterate on
    '''   next range finders.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YRangeFinder</c> object, corresponding to
    '''   the first range finder currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstRangeFinder() As YRangeFinder
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("RangeFinder", 0, p, size, neededsize, errmsg)
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
      Return YRangeFinder.FindRangeFinder(serial + "." + funcId)
    End Function

    REM --- (end of YRangeFinder public methods declaration)

  End Class

  REM --- (RangeFinder functions)

  '''*
  ''' <summary>
  '''   Retrieves a range finder for a given identifier.
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
  '''   This function does not require that the range finder is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YRangeFinder.isOnline()</c> to test if the range finder is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   a range finder by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the range finder
  ''' </param>
  ''' <returns>
  '''   a <c>YRangeFinder</c> object allowing you to drive the range finder.
  ''' </returns>
  '''/
  Public Function yFindRangeFinder(ByVal func As String) As YRangeFinder
    Return YRangeFinder.FindRangeFinder(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of range finders currently accessible.
  ''' <para>
  '''   Use the method <c>YRangeFinder.nextRangeFinder()</c> to iterate on
  '''   next range finders.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YRangeFinder</c> object, corresponding to
  '''   the first range finder currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstRangeFinder() As YRangeFinder
    Return YRangeFinder.FirstRangeFinder()
  End Function


  REM --- (end of RangeFinder functions)

End Module

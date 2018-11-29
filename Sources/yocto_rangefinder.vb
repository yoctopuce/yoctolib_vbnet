' ********************************************************************
'
'  $Id: yocto_rangefinder.vb 32908 2018-11-02 10:19:28Z seb $
'
'  Implements yFindRangeFinder(), the high-level API for RangeFinder functions
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

Module yocto_rangefinder

    REM --- (YRangeFinder return codes)
    REM --- (end of YRangeFinder return codes)
    REM --- (YRangeFinder dlldef)
    REM --- (end of YRangeFinder dlldef)
   REM --- (YRangeFinder yapiwrapper)
   REM --- (end of YRangeFinder yapiwrapper)
  REM --- (YRangeFinder globals)

  Public Const Y_RANGEFINDERMODE_DEFAULT As Integer = 0
  Public Const Y_RANGEFINDERMODE_LONG_RANGE As Integer = 1
  Public Const Y_RANGEFINDERMODE_HIGH_ACCURACY As Integer = 2
  Public Const Y_RANGEFINDERMODE_HIGH_SPEED As Integer = 3
  Public Const Y_RANGEFINDERMODE_INVALID As Integer = -1
  Public Const Y_HARDWARECALIBRATION_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_CURRENTTEMPERATURE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Delegate Sub YRangeFinderValueCallback(ByVal func As YRangeFinder, ByVal value As String)
  Public Delegate Sub YRangeFinderTimedReportCallback(ByVal func As YRangeFinder, ByVal measure As YMeasure)
  REM --- (end of YRangeFinder globals)

  REM --- (YRangeFinder class start)

  '''*
  ''' <summary>
  '''   The Yoctopuce class YRangeFinder allows you to use and configure Yoctopuce range finder
  '''   sensors.
  ''' <para>
  '''   It inherits from the YSensor class the core functions to read measurements,
  '''   register callback functions, access the autonomous datalogger.
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
    Public Const HARDWARECALIBRATION_INVALID As String = YAPI.INVALID_STRING
    Public Const CURRENTTEMPERATURE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    REM --- (end of YRangeFinder definitions)

    REM --- (YRangeFinder attributes declaration)
    Protected _rangeFinderMode As Integer
    Protected _hardwareCalibration As String
    Protected _currentTemperature As Double
    Protected _command As String
    Protected _valueCallbackRangeFinder As YRangeFinderValueCallback
    Protected _timedReportCallbackRangeFinder As YRangeFinderTimedReportCallback
    REM --- (end of YRangeFinder attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "RangeFinder"
      REM --- (YRangeFinder attributes initialization)
      _rangeFinderMode = RANGEFINDERMODE_INVALID
      _hardwareCalibration = HARDWARECALIBRATION_INVALID
      _currentTemperature = CURRENTTEMPERATURE_INVALID
      _command = COMMAND_INVALID
      _valueCallbackRangeFinder = Nothing
      _timedReportCallbackRangeFinder = Nothing
      REM --- (end of YRangeFinder attributes initialization)
    End Sub

    REM --- (YRangeFinder private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("rangeFinderMode") Then
        _rangeFinderMode = CInt(json_val.getLong("rangeFinderMode"))
      End If
      If json_val.has("hardwareCalibration") Then
        _hardwareCalibration = json_val.getString("hardwareCalibration")
      End If
      If json_val.has("currentTemperature") Then
        _currentTemperature = Math.Round(json_val.getDouble("currentTemperature") * 1000.0 / 65536.0) / 1000.0
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YRangeFinder private methods declaration)

    REM --- (YRangeFinder public methods declaration)

    '''*
    ''' <summary>
    '''   Changes the measuring unit for the measured range.
    ''' <para>
    '''   That unit is a string.
    '''   String value can be <c>"</c> or <c>mm</c>. Any other value is ignored.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    '''   WARNING: if a specific calibration is defined for the rangeFinder function, a
    '''   unit system change will probably break it.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a string corresponding to the measuring unit for the measured range
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
    '''   Returns the range finder running mode.
    ''' <para>
    '''   The rangefinder running mode
    '''   allows you to put priority on precision, speed or maximum range.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a value among <c>Y_RANGEFINDERMODE_DEFAULT</c>, <c>Y_RANGEFINDERMODE_LONG_RANGE</c>,
    '''   <c>Y_RANGEFINDERMODE_HIGH_ACCURACY</c> and <c>Y_RANGEFINDERMODE_HIGH_SPEED</c> corresponding to the
    '''   range finder running mode
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_RANGEFINDERMODE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_rangeFinderMode() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return RANGEFINDERMODE_INVALID
        End If
      End If
      res = Me._rangeFinderMode
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Changes the rangefinder running mode, allowing you to put priority on
    '''   precision, speed or maximum range.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a value among <c>Y_RANGEFINDERMODE_DEFAULT</c>, <c>Y_RANGEFINDERMODE_LONG_RANGE</c>,
    '''   <c>Y_RANGEFINDERMODE_HIGH_ACCURACY</c> and <c>Y_RANGEFINDERMODE_HIGH_SPEED</c> corresponding to the
    '''   rangefinder running mode, allowing you to put priority on
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
    Public Function get_hardwareCalibration() As String
      Dim res As String
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return HARDWARECALIBRATION_INVALID
        End If
      End If
      res = Me._hardwareCalibration
      Return res
    End Function


    Public Function set_hardwareCalibration(ByVal newval As String) As Integer
      Dim rest_val As String
      rest_val = newval
      Return _setAttr("hardwareCalibration", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the current sensor temperature, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the current sensor temperature, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_CURRENTTEMPERATURE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_currentTemperature() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return CURRENTTEMPERATURE_INVALID
        End If
      End If
      res = Me._currentTemperature
      Return res
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
    ''' <para>
    '''   If a call to this object's is_online() method returns FALSE although
    '''   you are certain that the matching device is plugged, make sure that you did
    '''   call registerHub() at application initialization time.
    ''' </para>
    ''' <para>
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
    '''   Returns the temperature at the time when the latest calibration was performed.
    ''' <para>
    '''   This function can be used to determine if a new calibration for ambient temperature
    '''   is required.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a temperature, as a floating point number.
    '''   On failure, throws an exception or return YAPI_INVALID_DOUBLE.
    ''' </returns>
    '''/
    Public Overridable Function get_hardwareCalibrationTemperature() As Double
      Dim hwcal As String
      hwcal = Me.get_hardwareCalibration()
      If (Not ((hwcal).Substring(0, 1) = "@")) Then
        Return YAPI.INVALID_DOUBLE
      End If
      Return YAPI._atoi((hwcal).Substring(1, (hwcal).Length))
    End Function

    '''*
    ''' <summary>
    '''   Triggers a sensor calibration according to the current ambient temperature.
    ''' <para>
    '''   That
    '''   calibration process needs no physical interaction with the sensor. It is performed
    '''   automatically at device startup, but it is recommended to start it again when the
    '''   temperature delta since the latest calibration exceeds 8°C.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function triggerTemperatureCalibration() As Integer
      Return Me.set_command("T")
    End Function

    '''*
    ''' <summary>
    '''   Triggers the photon detector hardware calibration.
    ''' <para>
    '''   This function is part of the calibration procedure to compensate for the the effect
    '''   of a cover glass. Make sure to read the chapter about hardware calibration for details
    '''   on the calibration procedure for proper results.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function triggerSpadCalibration() As Integer
      Return Me.set_command("S")
    End Function

    '''*
    ''' <summary>
    '''   Triggers the hardware offset calibration of the distance sensor.
    ''' <para>
    '''   This function is part of the calibration procedure to compensate for the the effect
    '''   of a cover glass. Make sure to read the chapter about hardware calibration for details
    '''   on the calibration procedure for proper results.
    ''' </para>
    ''' </summary>
    ''' <param name="targetDist">
    '''   true distance of the calibration target, in mm or inches, depending
    '''   on the unit selected in the device
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function triggerOffsetCalibration(targetDist As Double) As Integer
      Dim distmm As Integer = 0
      If (Me.get_unit() = """") Then
        distmm = CType(Math.Round(targetDist * 25.4), Integer)
      Else
        distmm = CType(Math.Round(targetDist), Integer)
      End If
      Return Me.set_command("O" + Convert.ToString(distmm))
    End Function

    '''*
    ''' <summary>
    '''   Triggers the hardware cross-talk calibration of the distance sensor.
    ''' <para>
    '''   This function is part of the calibration procedure to compensate for the the effect
    '''   of a cover glass. Make sure to read the chapter about hardware calibration for details
    '''   on the calibration procedure for proper results.
    ''' </para>
    ''' </summary>
    ''' <param name="targetDist">
    '''   true distance of the calibration target, in mm or inches, depending
    '''   on the unit selected in the device
    ''' </param>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function triggerXTalkCalibration(targetDist As Double) As Integer
      Dim distmm As Integer = 0
      If (Me.get_unit() = """") Then
        distmm = CType(Math.Round(targetDist * 25.4), Integer)
      Else
        distmm = CType(Math.Round(targetDist), Integer)
      End If
      Return Me.set_command("X" + Convert.ToString(distmm))
    End Function

    '''*
    ''' <summary>
    '''   Cancels the effect of previous hardware calibration procedures to compensate
    '''   for cover glass, and restores factory settings.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI_SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function cancelCoverGlassCalibrations() As Integer
      Return Me.set_hardwareCalibration("")
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of range finders started using <c>yFirstRangeFinder()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned range finders order.
    '''   If you want to find a specific a range finder, use <c>RangeFinder.findRangeFinder()</c>
    '''   and a hardwareID or a logical name.
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

  REM --- (YRangeFinder functions)

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
  ''' <para>
  '''   If a call to this object's is_online() method returns FALSE although
  '''   you are certain that the matching device is plugged, make sure that you did
  '''   call registerHub() at application initialization time.
  ''' </para>
  ''' <para>
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


  REM --- (end of YRangeFinder functions)

End Module

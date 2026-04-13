' ********************************************************************
'
'  $Id: svn_id $
'
'  Implements yFindOrientation(), the high-level API for Orientation functions
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

Module yocto_orientation

    REM --- (YOrientation return codes)
    REM --- (end of YOrientation return codes)
    REM --- (YOrientation dlldef)
    REM --- (end of YOrientation dlldef)
   REM --- (YOrientation yapiwrapper)
   REM --- (end of YOrientation yapiwrapper)
  REM --- (YOrientation globals)

  Public Const Y_COUNTERCLOCKWISE_FALSE As Integer = 0
  Public Const Y_COUNTERCLOCKWISE_TRUE As Integer = 1
  Public Const Y_COUNTERCLOCKWISE_INVALID As Integer = -1
  Public Const Y_COMMAND_INVALID As String = YAPI.INVALID_STRING
  Public Const Y_ZEROOFFSET_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Delegate Sub YOrientationValueCallback(ByVal func As YOrientation, ByVal value As String)
  Public Delegate Sub YOrientationTimedReportCallback(ByVal func As YOrientation, ByVal measure As YMeasure)
  REM --- (end of YOrientation globals)

  REM --- (YOrientation class start)

  '''*
  ''' <summary>
  '''   The <c>YOrientation</c> class allows you to read and configure Yoctopuce orientation sensors.
  ''' <para>
  '''   It inherits from <c>YSensor</c> class the core functions to read measurements,
  '''   to register callback functions, and to access the autonomous datalogger.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YOrientation
    Inherits YSensor
    REM --- (end of YOrientation class start)

    REM --- (YOrientation definitions)
    Public Const COUNTERCLOCKWISE_FALSE As Integer = 0
    Public Const COUNTERCLOCKWISE_TRUE As Integer = 1
    Public Const COUNTERCLOCKWISE_INVALID As Integer = -1
    Public Const COMMAND_INVALID As String = YAPI.INVALID_STRING
    Public Const ZEROOFFSET_INVALID As Double = YAPI.INVALID_DOUBLE
    REM --- (end of YOrientation definitions)

    REM --- (YOrientation attributes declaration)
    Protected _counterClockwise As Integer
    Protected _command As String
    Protected _zeroOffset As Double
    Protected _valueCallbackOrientation As YOrientationValueCallback
    Protected _timedReportCallbackOrientation As YOrientationTimedReportCallback
    REM --- (end of YOrientation attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Orientation"
      REM --- (YOrientation attributes initialization)
      _counterClockwise = COUNTERCLOCKWISE_INVALID
      _command = COMMAND_INVALID
      _zeroOffset = ZEROOFFSET_INVALID
      _valueCallbackOrientation = Nothing
      _timedReportCallbackOrientation = Nothing
      REM --- (end of YOrientation attributes initialization)
    End Sub

    REM --- (YOrientation private methods declaration)

    Protected Overrides Function _parseAttr(ByRef json_val As YJSONObject) As Integer
      If json_val.has("counterClockwise") Then
        If (json_val.getInt("counterClockwise") > 0) Then _counterClockwise = 1 Else _counterClockwise = 0
      End If
      If json_val.has("command") Then
        _command = json_val.getString("command")
      End If
      If json_val.has("zeroOffset") Then
        _zeroOffset = Math.Round(json_val.getDouble("zeroOffset") / 65.536) / 1000.0
      End If
      Return MyBase._parseAttr(json_val)
    End Function

    REM --- (end of YOrientation private methods declaration)

    REM --- (YOrientation public methods declaration)
    '''*
    ''' <summary>
    '''   Returns a value indicating whether the sensor is operating in a counterclockwise direction.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   either <c>YOrientation.COUNTERCLOCKWISE_FALSE</c> or <c>YOrientation.COUNTERCLOCKWISE_TRUE</c>,
    '''   according to a value indicating whether the sensor is operating in a counterclockwise direction
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YOrientation.COUNTERCLOCKWISE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_counterClockwise() As Integer
      Dim res As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return COUNTERCLOCKWISE_INVALID
        End If
      End If
      res = Me._counterClockwise
      Return res
    End Function


    '''*
    ''' <summary>
    '''   Defines the operating direction of the sensor.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the
    '''   modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   either <c>YOrientation.COUNTERCLOCKWISE_FALSE</c> or <c>YOrientation.COUNTERCLOCKWISE_TRUE</c>
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_counterClockwise(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("counterClockwise", rest_val)
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
    '''   Sets an offset between the orientation reported by the sensor and the actual orientation.
    ''' <para>
    '''   This
    '''   can typically be used  to compensate for mechanical offset. This offset can also be set
    '''   automatically using the zero() method.
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   a floating point number
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Function set_zeroOffset(ByVal newval As Double) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(Math.Round(newval * 65536.0)))
      Return _setAttr("zeroOffset", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the Offset between the orientation reported by the sensor and the actual orientation.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the Offset between the orientation reported by the sensor
    '''   and the actual orientation
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>YOrientation.ZEROOFFSET_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_zeroOffset() As Double
      Dim res As Double = 0
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI._yapiContext.GetCacheValidity()) <> YAPI.SUCCESS) Then
          Return ZEROOFFSET_INVALID
        End If
      End If
      res = Me._zeroOffset
      Return res
    End Function

    '''*
    ''' <summary>
    '''   Retrieves an orientation sensor for a given identifier.
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
    '''   This function does not require that the orientation sensor is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YOrientation.isOnline()</c> to test if the orientation sensor is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an orientation sensor by logical name, no error is notified: the first instance
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
    '''   a string that uniquely characterizes the orientation sensor, for instance
    '''   <c>MyDevice.orientation</c>.
    ''' </param>
    ''' <returns>
    '''   a <c>YOrientation</c> object allowing you to drive the orientation sensor.
    ''' </returns>
    '''/
    Public Shared Function FindOrientation(func As String) As YOrientation
      Dim obj As YOrientation
      obj = CType(YFunction._FindFromCache("Orientation", func), YOrientation)
      If ((obj Is Nothing)) Then
        obj = New YOrientation(func)
        YFunction._AddToCache("Orientation", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is then invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness,
    '''   remember to call one of these two functions periodically. The callback is called once juste after beeing
    '''   registered, passing the current advertised value  of the function, provided that it is not an empty string.
    '''   To unregister a callback, pass a Nothing pointer as argument.
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
    Public Overloads Function registerValueCallback(callback As YOrientationValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackOrientation = callback
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
      If (Not (Me._valueCallbackOrientation Is Nothing)) Then
        Me._valueCallbackOrientation(Me, value)
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
    '''   arguments: the function object of which the value has changed, and an <c>YMeasure</c> object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YOrientationTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackOrientation = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackOrientation Is Nothing)) Then
        Me._timedReportCallbackOrientation(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function

    Public Overridable Function sendCommand(command As String) As Integer
      Return Me.set_command(command)
    End Function

    '''*
    ''' <summary>
    '''   Reset the sensor's zero to current position by automatically setting a new offset.
    ''' <para>
    '''   Remember to call the <c>saveToFlash()</c> method of the module if the modification must be kept.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    '''   On failure, throws an exception or returns a negative error code.
    ''' </returns>
    '''/
    Public Overridable Function zero() As Integer
      Return Me.sendCommand("Z")
    End Function

    '''*
    ''' <summary>
    '''   Modifies the calibration of the MA600A sensor using an array of 32
    '''   values representing the offset in degrees between the true values and
    '''   those measured regularly every 11.25 degrees starting from zero.
    ''' <para>
    '''   The calibration
    '''   is applied immediately and is stored permanently in the MA600A sensor.
    '''   Before calculating the offset values, remember to clear any previous
    '''   calibration using the <c>clearCalibration</c> function and set
    '''   the zero offset  to 0. After a calibration change, the sensor will stop
    '''   measurements for about one second.
    '''   Do not confuse this function with the generic <c>calibrateFromPoints</c> function,
    '''   which works at the YSensor level and is not necessarily well suited to
    '''   a sensor returning circular values.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="offsetValues">
    '''   array of 32 floating point values in the [-11.25..+11.25] range
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function set_calibration(offsetValues As List(Of Double)) As Integer
      Dim res As String
      Dim npt As Integer = 0
      Dim idx As Integer = 0
      Dim corr As Integer = 0
      npt = offsetValues.Count
      If (npt <> 32) Then
        Me._throw(YAPI.INVALID_ARGUMENT, "Invalid calibration parameters (32 expected)")
        Return YAPI.INVALID_ARGUMENT
      End If
      res = "C"
      idx = 0
      While (idx < npt)
        corr = CType(Math.Round(offsetValues(idx) * 128 / 11.25), Integer)
        If ((corr < -128) OrElse (corr > 127)) Then
          Me._throw(YAPI.INVALID_ARGUMENT, "Calibration parameter exceeds permitted range (+/-11.25)")
          Return YAPI.INVALID_ARGUMENT
        End If
        If (corr < 0) Then
          corr = corr + 256
        End If
        res = "" + res + "" + (corr).ToString("x02")
        idx = idx + 1
      End While
      Return Me.sendCommand(res)
    End Function

    '''*
    ''' <summary>
    '''   Retrieves offset correction data points previously entered using the method
    '''   <c>set_calibration</c>.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="offsetValues">
    '''   array of 32 floating point numbers, that will be filled by the
    '''   function with the offset values for the correction points.
    ''' </param>
    ''' <para>
    ''' </para>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function get_Calibration(offsetValues As List(Of Double)) As Integer
      Return 0
    End Function

    '''*
    ''' <summary>
    '''   Cancels any calibration set with <c>set_calibration</c>.
    ''' <para>
    '''   This function
    '''   is equivalent to calling <c>set_calibration</c> with only zeros.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   <c>YAPI.SUCCESS</c> if the call succeeds.
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns a negative error code.
    ''' </para>
    '''/
    Public Overridable Function clearCalibration() As Integer
      Return Me.sendCommand("-")
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of orientation sensors started using <c>yFirstOrientation()</c>.
    ''' <para>
    '''   Caution: You can't make any assumption about the returned orientation sensors order.
    '''   If you want to find a specific an orientation sensor, use <c>Orientation.findOrientation()</c>
    '''   and a hardwareID or a logical name.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YOrientation</c> object, corresponding to
    '''   an orientation sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are no more orientation sensors to enumerate.
    ''' </returns>
    '''/
    Public Function nextOrientation() As YOrientation
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YOrientation.FindOrientation(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of orientation sensors currently accessible.
    ''' <para>
    '''   Use the method <c>YOrientation.nextOrientation()</c> to iterate on
    '''   next orientation sensors.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YOrientation</c> object, corresponding to
    '''   the first orientation sensor currently online, or a <c>Nothing</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstOrientation() As YOrientation
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Orientation", 0, p, size, neededsize, errmsg)
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
      Return YOrientation.FindOrientation(serial + "." + funcId)
    End Function

    REM --- (end of YOrientation public methods declaration)

  End Class

  REM --- (YOrientation functions)

  '''*
  ''' <summary>
  '''   Retrieves an orientation sensor for a given identifier.
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
  '''   This function does not require that the orientation sensor is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YOrientation.isOnline()</c> to test if the orientation sensor is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an orientation sensor by logical name, no error is notified: the first instance
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
  '''   a string that uniquely characterizes the orientation sensor, for instance
  '''   <c>MyDevice.orientation</c>.
  ''' </param>
  ''' <returns>
  '''   a <c>YOrientation</c> object allowing you to drive the orientation sensor.
  ''' </returns>
  '''/
  Public Function yFindOrientation(ByVal func As String) As YOrientation
    Return YOrientation.FindOrientation(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of orientation sensors currently accessible.
  ''' <para>
  '''   Use the method <c>YOrientation.nextOrientation()</c> to iterate on
  '''   next orientation sensors.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YOrientation</c> object, corresponding to
  '''   the first orientation sensor currently online, or a <c>Nothing</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstOrientation() As YOrientation
    Return YOrientation.FirstOrientation()
  End Function


  REM --- (end of YOrientation functions)

End Module

'*********************************************************************
'*
'* $Id: yocto_accelerometer.vb 24934 2016-06-30 22:32:01Z mvuilleu $
'*
'* Implements yFindAccelerometer(), the high-level API for Accelerometer functions
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

Module yocto_accelerometer

    REM --- (YAccelerometer return codes)
    REM --- (end of YAccelerometer return codes)
    REM --- (YAccelerometer dlldef)
    REM --- (end of YAccelerometer dlldef)
  REM --- (YAccelerometer globals)

  Public Const Y_BANDWIDTH_INVALID As Integer = YAPI.INVALID_INT
  Public Const Y_XVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_YVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_ZVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
  Public Const Y_GRAVITYCANCELLATION_OFF As Integer = 0
  Public Const Y_GRAVITYCANCELLATION_ON As Integer = 1
  Public Const Y_GRAVITYCANCELLATION_INVALID As Integer = -1
  Public Delegate Sub YAccelerometerValueCallback(ByVal func As YAccelerometer, ByVal value As String)
  Public Delegate Sub YAccelerometerTimedReportCallback(ByVal func As YAccelerometer, ByVal measure As YMeasure)
  REM --- (end of YAccelerometer globals)

  REM --- (YAccelerometer class start)

  '''*
  ''' <summary>
  '''   The YSensor class is the parent class for all Yoctopuce sensors.
  ''' <para>
  '''   It can be
  '''   used to read the current value and unit of any sensor, read the min/max
  '''   value, configure autonomous recording frequency and access recorded data.
  '''   It also provide a function to register a callback invoked each time the
  '''   observed value changes, or at a predefined interval. Using this class rather
  '''   than a specific subclass makes it possible to create generic applications
  '''   that work with any Yoctopuce sensor, even those that do not yet exist.
  '''   Note: The YAnButton class is the only analog input which does not inherit
  '''   from YSensor.
  ''' </para>
  ''' </summary>
  '''/
  Public Class YAccelerometer
    Inherits YSensor
    REM --- (end of YAccelerometer class start)

    REM --- (YAccelerometer definitions)
    Public Const BANDWIDTH_INVALID As Integer = YAPI.INVALID_INT
    Public Const XVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const YVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const ZVALUE_INVALID As Double = YAPI.INVALID_DOUBLE
    Public Const GRAVITYCANCELLATION_OFF As Integer = 0
    Public Const GRAVITYCANCELLATION_ON As Integer = 1
    Public Const GRAVITYCANCELLATION_INVALID As Integer = -1
    REM --- (end of YAccelerometer definitions)

    REM --- (YAccelerometer attributes declaration)
    Protected _bandwidth As Integer
    Protected _xValue As Double
    Protected _yValue As Double
    Protected _zValue As Double
    Protected _gravityCancellation As Integer
    Protected _valueCallbackAccelerometer As YAccelerometerValueCallback
    Protected _timedReportCallbackAccelerometer As YAccelerometerTimedReportCallback
    REM --- (end of YAccelerometer attributes declaration)

    Public Sub New(ByVal func As String)
      MyBase.New(func)
      _classname = "Accelerometer"
      REM --- (YAccelerometer attributes initialization)
      _bandwidth = BANDWIDTH_INVALID
      _xValue = XVALUE_INVALID
      _yValue = YVALUE_INVALID
      _zValue = ZVALUE_INVALID
      _gravityCancellation = GRAVITYCANCELLATION_INVALID
      _valueCallbackAccelerometer = Nothing
      _timedReportCallbackAccelerometer = Nothing
      REM --- (end of YAccelerometer attributes initialization)
    End Sub

    REM --- (YAccelerometer private methods declaration)

    Protected Overrides Function _parseAttr(ByRef member As TJSONRECORD) As Integer
      If (member.name = "bandwidth") Then
        _bandwidth = CInt(member.ivalue)
        Return 1
      End If
      If (member.name = "xValue") Then
        _xValue = Math.Round(member.ivalue * 1000.0 / 65536.0) / 1000.0
        Return 1
      End If
      If (member.name = "yValue") Then
        _yValue = Math.Round(member.ivalue * 1000.0 / 65536.0) / 1000.0
        Return 1
      End If
      If (member.name = "zValue") Then
        _zValue = Math.Round(member.ivalue * 1000.0 / 65536.0) / 1000.0
        Return 1
      End If
      If (member.name = "gravityCancellation") Then
        If (member.ivalue > 0) Then _gravityCancellation = 1 Else _gravityCancellation = 0
        Return 1
      End If
      Return MyBase._parseAttr(member)
    End Function

    REM --- (end of YAccelerometer private methods declaration)

    REM --- (YAccelerometer public methods declaration)
    '''*
    ''' <summary>
    '''   Returns the measure update frequency, measured in Hz (Yocto-3D-V2 only).
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   an integer corresponding to the measure update frequency, measured in Hz (Yocto-3D-V2 only)
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_BANDWIDTH_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_bandwidth() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return BANDWIDTH_INVALID
        End If
      End If
      Return Me._bandwidth
    End Function


    '''*
    ''' <summary>
    '''   Changes the measure update frequency, measured in Hz (Yocto-3D-V2 only).
    ''' <para>
    '''   When the
    '''   frequency is lower, the device performs averaging.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="newval">
    '''   an integer corresponding to the measure update frequency, measured in Hz (Yocto-3D-V2 only)
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
    Public Function set_bandwidth(ByVal newval As Integer) As Integer
      Dim rest_val As String
      rest_val = Ltrim(Str(newval))
      Return _setAttr("bandwidth", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Returns the X component of the acceleration, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the X component of the acceleration, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_XVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_xValue() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return XVALUE_INVALID
        End If
      End If
      Return Me._xValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the Y component of the acceleration, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the Y component of the acceleration, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_YVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_yValue() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return YVALUE_INVALID
        End If
      End If
      Return Me._yValue
    End Function

    '''*
    ''' <summary>
    '''   Returns the Z component of the acceleration, as a floating point number.
    ''' <para>
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a floating point number corresponding to the Z component of the acceleration, as a floating point number
    ''' </returns>
    ''' <para>
    '''   On failure, throws an exception or returns <c>Y_ZVALUE_INVALID</c>.
    ''' </para>
    '''/
    Public Function get_zValue() As Double
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return ZVALUE_INVALID
        End If
      End If
      Return Me._zValue
    End Function

    Public Function get_gravityCancellation() As Integer
      If (Me._cacheExpiration <= YAPI.GetTickCount()) Then
        If (Me.load(YAPI.DefaultCacheValidity) <> YAPI.SUCCESS) Then
          Return GRAVITYCANCELLATION_INVALID
        End If
      End If
      Return Me._gravityCancellation
    End Function


    Public Function set_gravityCancellation(ByVal newval As Integer) As Integer
      Dim rest_val As String
      If (newval > 0) Then rest_val = "1" Else rest_val = "0"
      Return _setAttr("gravityCancellation", rest_val)
    End Function
    '''*
    ''' <summary>
    '''   Retrieves an accelerometer for a given identifier.
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
    '''   This function does not require that the accelerometer is online at the time
    '''   it is invoked. The returned object is nevertheless valid.
    '''   Use the method <c>YAccelerometer.isOnline()</c> to test if the accelerometer is
    '''   indeed online at a given time. In case of ambiguity when looking for
    '''   an accelerometer by logical name, no error is notified: the first instance
    '''   found is returned. The search is performed first by hardware name,
    '''   then by logical name.
    ''' </para>
    ''' </summary>
    ''' <param name="func">
    '''   a string that uniquely characterizes the accelerometer
    ''' </param>
    ''' <returns>
    '''   a <c>YAccelerometer</c> object allowing you to drive the accelerometer.
    ''' </returns>
    '''/
    Public Shared Function FindAccelerometer(func As String) As YAccelerometer
      Dim obj As YAccelerometer
      obj = CType(YFunction._FindFromCache("Accelerometer", func), YAccelerometer)
      If ((obj Is Nothing)) Then
        obj = New YAccelerometer(func)
        YFunction._AddToCache("Accelerometer", func, obj)
      End If
      Return obj
    End Function

    '''*
    ''' <summary>
    '''   Registers the callback function that is invoked on every change of advertised value.
    ''' <para>
    '''   The callback is invoked only during the execution of <c>ySleep</c> or <c>yHandleEvents</c>.
    '''   This provides control over the time when the callback is triggered. For good responsiveness, remember to call
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and the character string describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerValueCallback(callback As YAccelerometerValueCallback) As Integer
      Dim val As String
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateValueCallbackList(Me, True)
      Else
        YFunction._UpdateValueCallbackList(Me, False)
      End If
      Me._valueCallbackAccelerometer = callback
      REM // Immediately invoke value callback with current value
      If (Not (callback Is Nothing) And Me.isOnline()) Then
        val = Me._advertisedValue
        If (Not (val = "")) Then
          Me._invokeValueCallback(val)
        End If
      End If
      Return 0
    End Function

    Public Overrides Function _invokeValueCallback(value As String) As Integer
      If (Not (Me._valueCallbackAccelerometer Is Nothing)) Then
        Me._valueCallbackAccelerometer(Me, value)
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
    '''   one of these two functions periodically. To unregister a callback, pass a null pointer as argument.
    ''' </para>
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <param name="callback">
    '''   the callback function to call, or a null pointer. The callback function should take two
    '''   arguments: the function object of which the value has changed, and an YMeasure object describing
    '''   the new advertised value.
    ''' @noreturn
    ''' </param>
    '''/
    Public Overloads Function registerTimedReportCallback(callback As YAccelerometerTimedReportCallback) As Integer
      Dim sensor As YSensor
      sensor = Me
      If (Not (callback Is Nothing)) Then
        YFunction._UpdateTimedReportCallbackList(sensor, True)
      Else
        YFunction._UpdateTimedReportCallbackList(sensor, False)
      End If
      Me._timedReportCallbackAccelerometer = callback
      Return 0
    End Function

    Public Overrides Function _invokeTimedReportCallback(value As YMeasure) As Integer
      If (Not (Me._timedReportCallbackAccelerometer Is Nothing)) Then
        Me._timedReportCallbackAccelerometer(Me, value)
      Else
        MyBase._invokeTimedReportCallback(value)
      End If
      Return 0
    End Function


    '''*
    ''' <summary>
    '''   Continues the enumeration of accelerometers started using <c>yFirstAccelerometer()</c>.
    ''' <para>
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAccelerometer</c> object, corresponding to
    '''   an accelerometer currently online, or a <c>null</c> pointer
    '''   if there are no more accelerometers to enumerate.
    ''' </returns>
    '''/
    Public Function nextAccelerometer() As YAccelerometer
      Dim hwid As String = ""
      If (YISERR(_nextFunction(hwid))) Then
        Return Nothing
      End If
      If (hwid = "") Then
        Return Nothing
      End If
      Return YAccelerometer.FindAccelerometer(hwid)
    End Function

    '''*
    ''' <summary>
    '''   Starts the enumeration of accelerometers currently accessible.
    ''' <para>
    '''   Use the method <c>YAccelerometer.nextAccelerometer()</c> to iterate on
    '''   next accelerometers.
    ''' </para>
    ''' </summary>
    ''' <returns>
    '''   a pointer to a <c>YAccelerometer</c> object, corresponding to
    '''   the first accelerometer currently online, or a <c>null</c> pointer
    '''   if there are none.
    ''' </returns>
    '''/
    Public Shared Function FirstAccelerometer() As YAccelerometer
      Dim v_fundescr(1) As YFUN_DESCR
      Dim dev As YDEV_DESCR
      Dim neededsize, err As Integer
      Dim serial, funcId, funcName, funcVal As String
      Dim errmsg As String = ""
      Dim size As Integer = Marshal.SizeOf(v_fundescr(0))
      Dim p As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(v_fundescr(0)))

      err = yapiGetFunctionsByClass("Accelerometer", 0, p, size, neededsize, errmsg)
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
      Return YAccelerometer.FindAccelerometer(serial + "." + funcId)
    End Function

    REM --- (end of YAccelerometer public methods declaration)

  End Class

  REM --- (Accelerometer functions)

  '''*
  ''' <summary>
  '''   Retrieves an accelerometer for a given identifier.
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
  '''   This function does not require that the accelerometer is online at the time
  '''   it is invoked. The returned object is nevertheless valid.
  '''   Use the method <c>YAccelerometer.isOnline()</c> to test if the accelerometer is
  '''   indeed online at a given time. In case of ambiguity when looking for
  '''   an accelerometer by logical name, no error is notified: the first instance
  '''   found is returned. The search is performed first by hardware name,
  '''   then by logical name.
  ''' </para>
  ''' </summary>
  ''' <param name="func">
  '''   a string that uniquely characterizes the accelerometer
  ''' </param>
  ''' <returns>
  '''   a <c>YAccelerometer</c> object allowing you to drive the accelerometer.
  ''' </returns>
  '''/
  Public Function yFindAccelerometer(ByVal func As String) As YAccelerometer
    Return YAccelerometer.FindAccelerometer(func)
  End Function

  '''*
  ''' <summary>
  '''   Starts the enumeration of accelerometers currently accessible.
  ''' <para>
  '''   Use the method <c>YAccelerometer.nextAccelerometer()</c> to iterate on
  '''   next accelerometers.
  ''' </para>
  ''' </summary>
  ''' <returns>
  '''   a pointer to a <c>YAccelerometer</c> object, corresponding to
  '''   the first accelerometer currently online, or a <c>null</c> pointer
  '''   if there are none.
  ''' </returns>
  '''/
  Public Function yFirstAccelerometer() As YAccelerometer
    Return YAccelerometer.FirstAccelerometer()
  End Function


  REM --- (end of Accelerometer functions)

End Module
